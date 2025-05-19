Office.onReady(() => {
  console.log("Nepa Templates Add-in is ready");

  const images = document.querySelectorAll(".background-image");

  images.forEach(img => {
    img.addEventListener("dragstart", (event) => {
      // This triggers PowerPoint to download the image when dropped
      const imageUrl = img.src;
      const imageName = img.alt || "background.png";

      event.dataTransfer.setData("DownloadURL", `image/png:${imageName}:${imageUrl}`);
      console.log(`Dragging ${imageName}`);
    });
  });
});
