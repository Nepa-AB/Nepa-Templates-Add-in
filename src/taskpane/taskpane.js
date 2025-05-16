Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");

    document.querySelectorAll('.background-image').forEach(img => {
      img.addEventListener('click', async () => {
        const imageUrl = img.src;

        try {
          await PowerPoint.run(async (context) => {
            const slide = context.presentation.slides.getSelected();
            const image = slide.shapes.addImage(imageUrl);
            image.left = 0;
            image.top = 0;
            image.width = 960; // Adjust as needed
            image.height = 540;

            await context.sync();
          });

          document.getElementById("notify").textContent = "Background inserted.";
        } catch (error) {
          console.error("Error inserting background:", error);
          document.getElementById("notify").textContent = "Error inserting background.";
        }
      });
    });
  }
});