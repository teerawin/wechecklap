
function onMessageSendHandler(event) {
  Office.context.ui.displayDialogAsync("https://teerawin.github.io/wechecklap/confirm.html",
    { height: 30, width: 20 },
    function (asyncResult) {
      const dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        const result = arg.message === "yes";
        event.completed({ allowEvent: result });
        dialog.close();
      });
    }
  );
}
