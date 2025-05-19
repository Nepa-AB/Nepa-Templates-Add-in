Office.onReady(() => {
  console.log("Nepa Templates Add-in is ready");

  const images = document.querySelectorAll(".background-image");

  images.forEach(img => {
    img.setAttribute("draggable", "true");

    img.addEventListener("dragstart", (event) => {
      const imageUrl = img.src;
      const imageName = img.getAttribute("alt") || "background.png";

      // Format: <mime>:<filename>:<url>
      const downloadURL = `image/png:${imageName}:${imageUrl}`;
      event.dataTransfer.setData("DownloadURL", downloadURL);

      console.log(`Dragging image: ${downloadURL}`);
    });
  });
});
