Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

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
});
