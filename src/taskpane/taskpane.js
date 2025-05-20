document.addEventListener("DOMContentLoaded", function () {
  const imageContainer = document.getElementById("imageContainer");
  const imageCategory = document.getElementById("imageCategory");
  const slideCategory = document.getElementById("slideCategory");
  const slidePreviews = document.getElementById("slidePreviews");

  const imageSources = {
    backgrounds: [
      "backgrounds/background 1.png",
      "backgrounds/background 2.png"
    ],
    half: [
      "Images/half%20page%201.jpg",
      "Images/half%20page%202.jpg",
      "Images/half%20page%203.jpg",
      "Images/half%20page%204.jpg",
      "Images/half%20page%205.jpg",
      "Images/half%20page%206.jpg"
    ],
    thin: [
      "Images/thin image 1.jpg",
      "Images/thin image 2.jpg",
      "Images/thin image 3.jpg",
      "Images/thin image 4.jpg",
      "Images/thin image 5.jpg",
      "Images/thin image 6.jpg"
    ]
  };

  const slidePreviewsData = {
    "Arrows, Numbers, Symbols, Banners": [
      "slides-previews/Arrows, Numbers, Symbols, Banners/slide1.jpg",
      "slides-previews/Arrows, Numbers, Symbols, Banners/slide2.jpg"
    ],
    Assets: [
      "slides-previews/Assets/slide1.jpg",
      "slides-previews/Assets/slide2.jpg",
      "slides-previews/Assets/slide3.jpg",
      "slides-previews/Assets/slide4.jpg",
      "slides-previews/Assets/slide5.jpg",
      "slides-previews/Assets/slide6.jpg",
      "slides-previews/Assets/slide7.jpg",
      "slides-previews/Assets/slide8.jpg",
      "slides-previews/Assets/slide9.jpg",
      "slides-previews/Assets/slide10.jpg"
    ]
  };

  function updateImages(category) {
    imageContainer.innerHTML = "";
    const basePath = "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/";
    imageSources[category].forEach((src) => {
      const img = document.createElement("img");
      img.src = basePath + src;
      img.className = "draggable-image";
      img.draggable = true;
      img.addEventListener("dragstart", (e) => {
        e.dataTransfer.setData("text/uri-list", img.src);
        console.log("Dragging image:", img.src);
      });
      imageContainer.appendChild(img);
    });
  }

  function updateSlidePreviews(category) {
    slidePreviews.innerHTML = "";
    const basePath = "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/";
    slidePreviewsData[category].forEach((src, index) => {
      const img = document.createElement("img");
      img.src = basePath + src;
      img.className = "slide-preview";
      img.title = `Click to insert slide ${index + 1}`;
      img.style.cursor = "pointer";
      img.addEventListener("click", () => {
        insertSlide(category, index + 1);
      });
      slidePreviews.appendChild(img);
    });
  }

  function insertSlide(category, slideIndex) {
    const pptxLinks = {
      "Arrows, Numbers, Symbols, Banners": "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Arrows, Numbers, Symbols, Banners.pptx",
      Assets: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Assets.pptx"
    };

    const pptxUrl = pptxLinks[category];
    Office.context.document.setSelectedDataAsync(
      `To insert slide ${slideIndex}, open: ${pptxUrl}`,
      { coercionType: Office.CoercionType.Text }
    );
  }

  imageCategory.addEventListener("change", (e) => {
    updateImages(e.target.value);
  });

  slideCategory.addEventListener("change", (e) => {
    updateSlidePreviews(e.target.value);
  });

  // Initial load
  updateImages("backgrounds");
  updateSlidePreviews("Arrows, Numbers, Symbols, Banners");
});
