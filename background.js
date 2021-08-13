var downloading = null;
var url = null;
browser.messageDisplayAction.onClicked.addListener(async (tab) => {
  browser.messageDisplay.getDisplayedMessage(tab.id).then(async (message) => {
    browser.messages.getRaw(message.id).then(async (raw) => {
      raw = raw.replace(/\r/g, "");
      raw = raw.match(/Content-Type: text\/calendar;(.|\n)*\n\n/);
      if (raw === null) {
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
      raw = raw[0].replace(/\n\n$/, "");
      raw = raw.replace(/^(.|\n)*\n\n/, "");
      var ics = atob(raw);
      var blob = new Blob([ics]);
      url = URL.createObjectURL(blob);
      downloading = await browser.downloads.download(
        {
          "filename": "calendar.ics",
          "saveAs": true,
          "url": url
        }
      );
    });
  });
});

browser.downloads.onChanged.addListener((downloadDelta) => {
  if (downloadDelta.id == downloading && downloadDelta.state.current == "complete") {
    URL.revokeObjectURL(url);
  }
});
