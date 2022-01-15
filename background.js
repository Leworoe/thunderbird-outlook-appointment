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

browser.messageDisplay.onMessageDisplayed.addListener(async (tab, message) => {
  if (await getRawCalendar(message.id) === false) {
    browser.messageDisplayAction.disable(tab.id);
  } else {
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
  const blob = base64ToBlob(cal);
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
