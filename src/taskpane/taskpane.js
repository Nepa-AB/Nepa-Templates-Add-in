Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");

    document.querySelectorAll('.background-image').forEach(img => {
      img.setAttribute('draggable', true);

      img.addEventListener('dragstart', (event) => {
        const imageUrl = img.src;
        const imageName = img.getAttribute('data-imagename') || 'background.png';

        const downloadURL = `image/png:${imageName}:${imageUrl}`;
        event.dataTransfer.setData('DownloadURL', downloadURL);

        console.log("Dragging image:", downloadURL);
      });
    });
  }
});
