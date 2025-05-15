/* global Office */

// Wait until Office is initialized
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.querySelectorAll('.background-image').forEach(img => {
      img.addEventListener('click', function () {
        insertBackground(this.getAttribute('data-imagename'));
      });
    });
  }
});

// Create function in local scope
function insertBackground(imageName) {
  const notify = document.getElementById("notify");
  if (notify) {
    notify.innerHTML = '<span style="color: #444;">Inserting background image…</span>';
  }
  const imgUrl = `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/backgrounds/${encodeURIComponent(imageName)}`;
  Office.context.document.setSelectedDataAsync(
    imgUrl,
    { coercionType: Office.CoercionType.Image },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        if (notify) {
          notify.innerHTML = '<span style="color:#0a8402">Image <b>inserted!</b> You can resize/move it as needed.</span>';
        }
      } else {
        if (notify) {
          notify.innerHTML =
            '<span style="color:#d83114;">' +
            "Error: " +
            (asyncResult.error && asyncResult.error.message
              ? asyncResult.error.message
              : "Could not insert image. Please select a content placeholder or click a position on the slide, then try again.") +
            '</span>';
        }
      }
    }
  );
}

// Expose only if needed globally
// window.insertBackground = insertBackground;