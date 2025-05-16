Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");

    document.querySelectorAll('.background-image').forEach(img => {
      img.addEventListener('click', async () => {
        const imageUrl = img.src;

        try {
          const base64Image = await fetchImageAsBase64(imageUrl);

          await PowerPoint.run(async (context) => {
            const slide = context.presentation.slides.getActiveSlide();
            const imageShape = slide.shapes.addImage(base64Image);
            imageShape.left = 0;
            imageShape.top = 0;
            imageShape.width = 960;
            imageShape.height = 540;

            await context.sync();
            document.getElementById("notify").textContent = "Background inserted successfully.";
          });
        } catch (error) {
          console.error("Error inserting background:", error);
          document.getElementById("notify").textContent = "Error inserting background.";
        }
      });
    });

    async function fetchImageAsBase64(imageUrl) {
      const response = await fetch(imageUrl);
      const blob = await response.blob();

      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          // Strip the prefix from data URI
          const base64data = reader.result.split(',')[1];
          resolve(base64data);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    }
  }
});

