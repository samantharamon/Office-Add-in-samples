/* global Office */

// Register keyboard shortcut handlers for showing/hiding the task pane.
Office.actions.associate("ShowTaskpane", () => {
  return Office.addin
    .showAsTaskpane()
    .then(() => {
      return;
    })
    .catch((error) => {
      return error.code;
    });
});

Office.actions.associate("HideTaskpane", () => {
  return Office.addin
    .hide()
    .then(() => {
      return;
    })
    .catch((error) => {
      return error.code;
    });
});
