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

// Rest of the code (loadSlides, insertSlide, etc.) is unchanged and works as before
async function loadSlides(category) {
  const slidePreviews = {
    "Arrows, Numbers, Symbols, Banners": [
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Arrows, Numbers, Symbols, Banners/slide1.jpg",
      "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Arrows, Numbers, Symbols, Banners/slide2.jpg"
    ],
    "Assets": [
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
  };

  const slideFiles = {
    "Arrows, Numbers, Symbols, Banners": "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Arrows, Numbers, Symbols, Banners.pptx",
    "Assets": "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Assets.pptx"
  };

  const container = document.getElementById("slidePreviews");
  container.innerHTML = "";

  if (!slidePreviews[category]) return;

  slidePreviews[category].forEach((url, index) => {
    const img = document.createElement("img");
    img.src = url;
    img.alt = `Slide ${index + 1}`;
    img.style.cursor = "pointer";
    img.addEventListener("click", () => insertSlide(slideFiles[category], index + 1));
    container.appendChild(img);
  });
}

async function insertSlide(fileUrl, slideNumber) {
  try {
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      slides.load("items");
      await context.sync();
      await slides.insertFromFile(fileUrl, 0, { slideStart: slideNumber, slideEnd: slideNumber });
      await context.sync();
    });
    // No notification needed for images now!
  } catch (error) {
    console.error(error);
    // Optional: display error if needed
  }
}
