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
    // Compose full, encoded URL:
    const imgUrl = `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/backgrounds/${encodeURIComponent(imageName)}`;
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length > 0) {
        const slide = slides.items[0];
        const shape = slide.shapes.addImage(imgUrl);
        
        // Optionally cover the whole slide (try to get slideWidth/slideHeight)
        // These API props may not always be reliable in all Office versions,
        // If not, slide dimensions are often 960x540 for 16:9 aspect ratio.
        /*
        shape.left = 0;
        shape.top = 0;
        shape.width = 960;
        shape.height = 540;
        */

      }
      await context.sync();
    });
    const notify = document.getElementById("notify");
    if (notify) notify.innerText = "Background image inserted onto the slide!";
  } catch (e) {
    const notify = document.getElementById("notify");
    if (notify) notify.innerText = "Error: " + e;
  }
};

// For future: add insertIcon, insertSlide, etc.
