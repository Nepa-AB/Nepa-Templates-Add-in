Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");

    document.querySelectorAll('.background-image').forEach(img => {
      img.addEventListener('click', async () => {
        // Change extension to .png in case image src is still pointing to .jpg
        const imageUrl = img.src.replace(/\.jpg$/i, '.png');

        try {
          const imageBase64Uri = await fetchImageAsBase64(imageUrl); // Full data URI

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
          resolve(reader.result); // Use full data URI (includes MIME)
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    }
  }
});


