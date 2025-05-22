Office.onReady(() => {
  loadImages("backgrounds");
  loadSlides("Arrows, Numbers, Symbols, Banners");

  document.getElementById("imageCategory").addEventListener("change", (e) => {
    loadImages(e.target.value);
  });

  document.getElementById("slideCategory").addEventListener("change", (e) => {
    loadSlides(e.target.value);
  });
});

const images = {
  backgrounds: [
    "https://upload.wikimedia.org/wikipedia/commons/4/47/PNG_transparency_demonstration_1.png",
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

function loadImages(category) {
  const container = document.getElementById("imageContainer");
  container.innerHTML = "";
  let baseUrl = "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/Images/";

  images[category].forEach((imgName) => {
    const img = document.createElement("img");
    img.src = imgName.startsWith("http") ? imgName : baseUrl + encodeURIComponent(imgName);
    img.alt = imgName;
    img.draggable = true; // enables drag and drop natively
    container.appendChild(img);
  });
}

// ------------------ SLIDES SECTION UPDATE ------------------

const slidePacks = {
  "Arrows, Numbers, Symbols, Banners": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Arrows, Numbers, Symbols, Banners.pptx",
    previews: [
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Arrows, Numbers, Symbols, Banners/slide1.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Arrows, Numbers, Symbols, Banners/slide2.jpg"
    ]
  },
  "Assets": {
    file: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Assets.pptx",
    previews: [
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide1.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide2.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide3.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide4.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide5.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide6.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide7.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide8.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide9.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide10.jpg"
    ]
  }
};

function loadSlides(category) {
  const container = document.getElementById("slidePreviews");
  container.innerHTML = "";

  const pack = slidePacks[category];
  if (!pack) return;

  // Download button
  const downloadBtn = document.createElement("a");
  downloadBtn.href = pack.file;
  downloadBtn.className = "template-button";
  downloadBtn.target = "_blank";
  downloadBtn.download = "";
  downloadBtn.style.marginBottom = "8px";
  downloadBtn.textContent = "Download Slides Pack";
  container.appendChild(downloadBtn);

  // Message
  const msg = document.createElement("div");
  msg.style.margin = "10px 0 8px 0";
  msg.style.fontSize = "14px";
  msg.style.color = "#555";
  msg.textContent = "After downloading, open the slides pack and copy any slide you want into your own presentation.";
  container.appendChild(msg);

  // Previews
  const previewsDiv = document.createElement("div");
  previewsDiv.className = "slide-previews-images";
  pack.previews.forEach((url, idx) => {
    const img = document.createElement("img");
    img.src = url;
    img.alt = `Slide ${idx + 1}`;
    img.className = "slide-preview-img";
    previewsDiv.appendChild(img);
  });
  container.appendChild(previewsDiv);
}

// (insertSlide and notification functions are no longer needed for slides)

