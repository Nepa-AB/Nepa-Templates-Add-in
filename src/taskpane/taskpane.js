Office.onReady(() => {
  loadImages("backgrounds");
  loadSlides(Object.keys(slidePacks)[0]); // Load the first slide pack by default

  document.getElementById("imageCategory").addEventListener("change", (e) => {
    loadImages(e.target.value);
  });

  document.getElementById("slideCategory").addEventListener("change", (e) => {
    loadSlides(e.target.value);
  });
});

const images = {
  backgrounds: [
    "background 1.png",
    "background 2.png"
  ],
  half: [
    "half page 1.jpg",
    "half page 2.jpg",
    "half page 3.jpg",
    "half page 4.jpg",
    "half page 5.jpg",
    "half page 6.jpg"
  ],
  thin: [
    "thin image 1.jpg",
    "thin image 2.jpg",
    "thin image 3.jpg",
    "thin image 4.jpg",
    "thin image 5.jpg",
    "thin image 6.jpg"
  ]
};

// Slides packs—add new ones here!
const slidePacks = {
  "Arrows, Numbers, Symbols, Banners": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Arrows, Numbers, Symbols, Banners.pptx"
  },
  "Assets": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Assets.pptx"
  },
  "Funnels, pyramids, dimensional charts": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Funnels, pyramids, dimensional charts.pptx"
  },
  "Maps": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Maps.pptx"
  },
  "Text Boxes, Layouts": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Text Boxes, Layouts.pptx"
  },
  "Timelines": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Timelines.pptx"
  }
};

function loadImages(category) {
  const container = document.getElementById("imageContainer");
  container.innerHTML = "";
  let baseUrl = "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/Images/";

  images[category].forEach((imgName) => {
    const img = document.createElement("img");
    img.src = imgName.startsWith("http") ? imgName : baseUrl + encodeURIComponent(imgName);
    img.alt = imgName;
    img.draggable = true;
    container.appendChild(img);
  });
}

function loadSlides(category) {
  const container = document.getElementById("slidePreviews");
  container.innerHTML = "";

  const pack = slidePacks[category];
  if (!pack) return;

  // Download button
  const downloadBtn = document.createElement("a");
  downloadBtn.href = pack.file;
  downloadBtn.className = "template-button";
  downloadBtn.download = ""; // uses default filename from URL
  downloadBtn.style.marginBottom = "8px";
  downloadBtn.textContent = "Download Slides Pack";
  container.appendChild(downloadBtn);

  // Message
  const msg = document.createElement("div");
  msg.style.margin = "10px 0 8px 0";
  msg.style.fontSize = "14px";
  msg.style.color = "#555";
  msg.textContent = "After downloading, open the slides pack and copy any slide or the object you want into your own presentation.";
  container.appendChild(msg);
}
