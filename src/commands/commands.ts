/* global Office */

Office.onReady(() => {
  // Office.js is ready
});

function showTaskpane(event: Office.AddinCommands.Event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

Office.actions.associate("showTaskpane", showTaskpane);
