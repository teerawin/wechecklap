function onMessageSendHandler(event) {
  Office.context.ui.displayDialogAsync("https://teerawin.github.io/wechecklap/confirm.html",
    { height: 30, width: 20, displayInIframe: true },
    function (asyncResult) {
      event.completed({ allowEvent: true }); // หรือตาม logic จริง
    });
}
