Office.onReady(() => {
  loadImages("backgrounds");
  loadSlides(Object.keys(slidePacks)[0]);

  document.getElementById("imageCategory").addEventListener("change", (e) => {
    loadImages(e.target.value);
  });

  document.getElementById("slideCategory").addEventListener("change", (e) => {
    loadSlides(e.target.value);
  });

  // Fonts section init
  const fontSelect = document.getElementById("fontCategory");
  if (fontSelect) {
    loadFont(fontSelect.value);
    fontSelect.addEventListener("change", function (e) {
      loadFont(e.target.value);
    });
  }
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

// Slides packs
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
  downloadBtn.download = "";
  downloadBtn.style.marginBottom = "8px";
  downloadBtn.textContent = "Download Slides Pack";
  container.appendChild(downloadBtn);

  // Message
  const msg = document.createElement("div");
  msg.className = "instruction-text";
  msg.style.margin = "10px 0 8px 0";
  msg.textContent = "After downloading, open the slides pack and copy any slide or the object you want into your own presentation.";
  container.appendChild(msg);
}

// Fonts section logic
const fonts = {
  "Helvetica Neue LT Pro 25 Ultra Light": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/fonts/Helvetica Neue LT Pro 25 Ultra Light.otf"
  },
  "Helvetica Neue LT STD 75 Bold": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/fonts/helvetica-neue-lt-std-75-bold.otf"
  },
  "HelveticaNeue LT 55 Roman Regular": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/fonts/HelveticaNeue LT 55 Roman Regular.ttf"
  }
};

function loadFont(selected) {
  const container = document.getElementById("fontDownloadContainer");
  container.innerHTML = "";

  const font = fonts[selected];
  if (!font) return;

  const downloadBtn = document.createElement("a");
  downloadBtn.href = font.file;
  downloadBtn.className = "template-button";
  downloadBtn.download = "";
  downloadBtn.textContent = "Download Font";
  container.appendChild(downloadBtn);

  // Comment is now static in HTML, so no code needed for it here
}
