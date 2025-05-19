Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");

    document.querySelectorAll('.background-image').forEach(img => {
      img.setAttribute('draggable', true);

      img.addEventListener('dragstart', (event) => {
        console.log("Native drag started for image:", img.src);
        // No need to set dataTransfer manually — browser handles image drag
      });
    });
  }
});
