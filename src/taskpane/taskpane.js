Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");

    document.querySelectorAll('.background-image').forEach(img => {
      img.addEventListener('click', async () => {
        const imageUrl = img.src;

        try {
          const base64Image = await fetchImageAsBase64(imageUrl);

          // Full Data URI (important: include MIME type)
          const imageBase64Uri = "data:image/jpeg;base64," + base64Image;
          console.log("Base64 Image URI:", imageBase64Uri);

          // Insert image using Office.js
          Office.context.document.setSelectedDataAsync(imageBase64Uri, {
            coercionType: Office.CoercionType.Image
          }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              document.getElementById("notify").textContent = "Background inserted successfully.";
              console.log("Image inserted successfully.");
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

    async function fetchImageAsBase64(imageUrl) {
      const response = await fetch(imageUrl);
      const blob = await response.blob();

      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          const dataUrl = reader.result;
          console.log("Full Data URI for testing:", dataUrl); // For manual testing
          const base64data = dataUrl.split(',')[1]; // Strip 'data:image/jpeg;base64,'
          resolve(base64data);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    }
  }
});


