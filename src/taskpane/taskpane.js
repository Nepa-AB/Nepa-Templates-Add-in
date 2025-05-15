/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Future: Any extra init
  }
});

window.insertBackground = function (imageName) {
  const imgUrl = `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/backgrounds/${encodeURIComponent(imageName)}`;
  const notify = document.getElementById("notify");
  Office.context.document.setSelectedDataAsync(
    imgUrl,
    { coercionType: Office.CoercionType.Image },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        if (notify) notify.innerText = "Background image inserted into current slide position!";
      } else {
        if (notify) notify.innerText =
          "Error: " +
          (asyncResult.error && asyncResult.error.message
            ? asyncResult.error.message
            : "Couldn't insert background. Select a placeholder or click in the slide to set position.");
      }
    }
  );
};