// Called when Office is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("PowerPoint Add-in ready");
    // No need to bind click events — drag-and-drop handles everything now
  }
});

// Global drag handler
function handleDragStart(event) {
  const imageUrl = event.target.src;

  // Set both URI list and plain text formats
  event.dataTransfer.setData("text/uri-list", imageUrl);
  event.dataTransfer.setData("text/plain", imageUrl);

  // Optional: set drag preview
  const img = new Image();
  img.src = imageUrl;
  event.dataTransfer.setDragImage(img, 10, 10);

  console.log("Dragging image:", imageUrl);
}

