Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");

    document.querySelectorAll('.background-image').forEach(img => {
      img.addEventListener('click', async () => {
        const imageUrl = img.src;

        try {
          const base64Image = await fetchImageAsBase64(imageUrl);

          const imageBase64Uri = "data:image/jpeg;base64," + base64Image;

          Office.context.document.setSelectedDataAsync(imageBase64Uri, {
            coercionType: Office.CoercionType.Image
          }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              document.getElementById("notify").textContent = "Background inserted.";
            } else {
              console.error("Insertion failed: ", asyncResult.error.message);
              document.getElementById("notify").textContent = "Error inserting background.";
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
          const base64data = reader.result.split(',')[1]; // Strip data:image/jpeg;base64,
          resolve(base64data);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    }
  }
});
