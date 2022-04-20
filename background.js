function base64ToBlob(base64String) {
  const binaryString = atob(base64String);
  const codePoints = Array(binaryString.length);
  for (let index = 0; index < binaryString.length; index++)
    codePoints[index] = binaryString.codePointAt(index);
  const uint8CodePoints = Uint8Array.from(codePoints);
  return new Blob([uint8CodePoints]);
}

async function getRawCalendar(message_id) {
  let raw = await browser.messages.getRaw(message_id);
  raw = raw.replace(/\r/g, "");
  raw = raw.match(/Content-Type: text\/calendar;(.|\n)*\n\n/);
  if (raw === null) {
    return false;
  }
  raw = raw[0].replace(/\n\n$/, "");
  raw = raw.replace(/^(.|\n)*\n\n/, "");
  return raw;
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

function replaceLocationByUrl(cal) {
  // Replace default location by url
  let ics = atob(cal);
  // Another option could be to use X-MICROSOFT-SKYPETEAMSMEETINGURL
  // But it is not used by every system
  conference_url = ics.match(/<(https:\/\/teams\.microsoft\.com.*?)>/s);
  if (conference_url) {
    ics = ics.replace(/LOCATION.*/, `LOCATION:${conference_url[1].replace(/\s+/g, '')}`);
    cal = btoa(ics);
  }

  return cal;
}

browser.messageDisplay.onMessageDisplayed.addListener(async (tab, message) => {
  if (await getRawCalendar(message.id) === false) {
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
  let cal = await getRawCalendar(message.id);
  if (cal === false) {
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

  cal = replaceLocationByUrl(cal);

  const blob = base64ToBlob(cal);
  url = URL.createObjectURL(blob);
  const escapedSubject = message.subject.replaceAll(/[/\\:*?"<>|]/g, " ")
  downloading = await browser.downloads.download(
    {
      "filename": `${escapedSubject}.ics`,
      "saveAs": true,
      "url": url,
      "conflictAction": 'uniquify'
    }
  );
});

browser.downloads.onChanged.addListener((downloadDelta) => {
  if (downloadDelta.id === downloading && downloadDelta.state.current === "complete") {
    URL.revokeObjectURL(url);
  }
});
