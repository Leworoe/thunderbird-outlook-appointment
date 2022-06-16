function base64ToText(base64String) {
  const binaryString = atob(base64String);
  const codePoints = Array(binaryString.length);
  for (let index = 0; index < binaryString.length; index++)
    codePoints[index] = binaryString.codePointAt(index);
  const uint8CodePoints = Uint8Array.from(codePoints);
  return new TextDecoder().decode(uint8CodePoints);
}

async function getBase64invite(message_id) {
  let raw = await browser.messages.getRaw(message_id);
  raw = raw.replace(/\r/g, "");
  let matches = raw.match(/(?=Content-Type: text\/calendar;)(?:[\s\S]*?\n\n)(?<base64>[\s\S]*?)\n\n/);
  if (matches === null) {
    return false;
  }
  return matches['groups'].base64;
}

async function isDarkmode() {
  let theme = await browser.theme.getCurrent();
  return theme.colors?.icons == '#fbfbfe' || theme.colors?.icons == 'rgb(249, 249, 250, 0.7)';
}

async function updateIcon() {
  if (await isDarkmode()) {
    browser.messageDisplayAction.setIcon({path: 'images/calendar_white.png'});
  } else {
    browser.messageDisplayAction.setIcon({path: 'images/calendar_black.png'});
  }
}

// Replace default location by url
function replaceLocationByUrl(ics) {
  // Remove linebreaks and initial whitespaces
  let new_ics = ics.replace(/(\r\n |\n |\r )/g, '');
  // Another option could be to use X-MICROSOFT-SKYPETEAMSMEETINGURL
  // But it is not used by every system
  const conference_url = new_ics.match(/<(https:\/\/teams\.microsoft\.com\/l\/meetup-join\/.*?)>/s);
  if (conference_url !== null) {
    return new_ics.replace(/LOCATION.*/, `LOCATION:${conference_url[1].replace(/\s+/g, '')}`);
  }

  return ics;
}

browser.messageDisplay.onMessageDisplayed.addListener(async (tab, message) => {
  if (await getBase64invite(message.id) === false) {
    browser.messageDisplayAction.disable(tab.id);
  } else {
    updateIcon();
    browser.messageDisplayAction.enable(tab.id);
  }
});

let downloading = null;
let url = null;

browser.messageDisplayAction.onClicked.addListener(async (tab) => {
  const message = await browser.messageDisplay.getDisplayedMessage(tab.id);
  let base64invite = await getBase64invite(message.id);
  if (base64invite === false) {
    browser.notifications.create(
      "no-appointment",
      {
        type: "basic",
        message: "No appointment was found.",
        title: "Outlook/Teams Appointments"
      }
    );
    setTimeout(() => {
      browser.notifications.clear("no-appointment");
    }, 5000);
    return;
  }
  let ics = base64ToText(base64invite);
  // Remove everything that is not beween "BEGIN:VCALENDAR" and "END:VCALENDAR"
  const ics_matched = ics.match(/(?:[\s\S]*?)(?<VCALENDAR>BEGIN:VCALENDAR(?:[\s\S]*?)END:VCALENDAR)/);
  if (ics_matched !== null) {
    ics = ics_matched['groups'].VCALENDAR;
  }

  ics = replaceLocationByUrl(ics);

  const blob = new Blob([ics]);
  url = URL.createObjectURL(blob);
  const escapedSubject = message.subject.replaceAll(/[/\\:*?"<>|]/g, " ")
  downloading = await browser.downloads.download(
    {
      "filename": `${escapedSubject}.ics`,
      "saveAs": true,
      "url": url
    }
  );
});

browser.downloads.onChanged.addListener((downloadDelta) => {
  if (downloadDelta.id === downloading && downloadDelta.state.current === "complete") {
    URL.revokeObjectURL(url);
  }
});
