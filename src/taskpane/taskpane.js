Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");

    document.querySelectorAll('.background-image').forEach(img => {
      img.addEventListener('click', async () => {
        const imageUrl = img.src;

        try {
          const base64Image = await fetchImageAsBase64(imageUrl);

          // Build full data URI for JPEG image
          const imageBase64Uri = "data:image/jpeg;base64," + base64Image;

          // Log the full URI so you can paste it in a browser to test
          console.log("Full Data URI:", imageBase64Uri);

          // Insert image into slide using Office.js API
          Office.context.document.setSelectedDataAsync(imageBase64Uri, {
            coercionType: Office.CoercionType.Image
          }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              document.getElementById("notify").textContent = "Background inserted.";
            } else {
              console.error("Insert failed:", asyncResult.error);
              document.getElementById("notify").textContent = "Failed to insert background.";
            }
          });
        } catch (error) {
          console.error("Error inserting background:", error);
          document.getElementById("notify").textContent = "Error inserting background.";
        }
      });
    });

    // Fetch the image and convert to base64
    async function fetchImageAsBase64(imageUrl) {
      const response = await fetch(imageUrl);
      const blob = await response.blob();

      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          const dataUrl = reader.result;
          console.log("Full Data URI (for testing in browser):", dataUrl); // 👈 Full output here
          const base64data = dataUrl.split(',')[1]; // Strip prefix
          resolve(base64data);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    }
  }
});

