/* global Office */

// Called when Office is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Nothing to show/hide here unless you have a sideload message or extra UI
    // You can also wire up other startup events here
  }
});

// Function to open a new tab with the template
window.insertTemplate = function(url) {
  window.open(url, '_blank');
  const notify = document.getElementById("notify");
  if (notify) notify.innerText = "Template opened in a new tab.";
};

// Example background insertion function -- needs background image hosted online
window.insertBackground = async function(imageUrl) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length > 0) {
        const slide = slides.items[0];
        const shape = slide.shapes.addImage(imageUrl);
        // You may set shape position and size below, but PowerPoint API sizing may be limited.
        shape.left = 0;
        shape.top = 0;
        // Try to fit to the slide's width and height.
        // For default slides, often 960x540, but can be different.
        // If you want to cover the slide, use:
        // const slideWidth = context.presentation.slideWidth;
        // const slideHeight = context.presentation.slideHeight;
        // shape.width = slideWidth;
        // shape.height = slideHeight;
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

// Example: for future, you can add insertIcon, insertSlide, etc.
// window.insertIcon = function(url) { ... }
// window.insertSlide = function(url) { ... }
