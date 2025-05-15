/* global Office */

// This is required by Office - it ensures APIs are ready
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Initialization, if needed
  }
});

// Background insertion function
window.insertBackground = async function(imageName) {
  try {
    // Compose the full, encoded URL
    const imgUrl = `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/backgrounds/${encodeURIComponent(imageName)}`;
    await PowerPoint.run(async (context) => {
      // Get the active slide (much more reliable)
      const slide = context.presentation.slides.getActiveSlide();
      // Insert the background image
      const shape = slide.shapes.addImage(imgUrl);

      // OPTIONAL: Uncomment for full-slide coverage (960x540 is standard, adjust as needed)
      // shape.left = 0;
      // shape.top = 0;
      // shape.width = 960;
      // shape.height = 540;

      await context.sync();
    });
    const notify = document.getElementById("notify");
    if (notify) notify.innerText = "Background image inserted onto the slide!";
  } catch (e) {
    const notify = document.getElementById("notify");
    if (notify) notify.innerText = "Error: " + (e.message || e);
    if (window.console) console.error(e);
  }
};

// For future: add insertIcon, insertSlide, etc.