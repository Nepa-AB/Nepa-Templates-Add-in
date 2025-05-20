document.addEventListener("DOMContentLoaded", () => {
  const imageCategory = document.getElementById("imageCategory");
  const imageContainer = document.getElementById("imageContainer");

  const slideCategory = document.getElementById("slideCategory");
  const slidesContainer = document.getElementById("slidesContainer");

  function renderImages(category) {
    imageContainer.innerHTML = "";
    let imageList = [];

    if (category === "backgrounds") {
      imageList = Array.from({ length: 6 }, (_, i) => `backgrounds/background ${i + 1}.png`);
    } else if (category === "halfpage") {
      imageList = Array.from({ length: 6 }, (_, i) => `Images/half page ${i + 1}.jpg`);
    } else if (category === "thin") {
      imageList = Array.from({ length: 6 }, (_, i) => `Images/thin image ${i + 1}.jpg`);
    }

    imageList.forEach((src) => {
      const img = document.createElement("img");
      img.src = `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/${src}`;
      img.alt = src;
      img.draggable = true;

      img.addEventListener("dragstart", (e) => {
        e.dataTransfer.setData("text/uri-list", img.src);
        console.log("Dragging image:", img.src);
      });

      imageContainer.appendChild(img);
    });
  }

  function renderSlidePreviews(category) {
    slidesContainer.innerHTML = "";

    let previews = [];
    if (category === "arrows") {
      previews = [1, 2].map(i => ({
        src: `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Arrows, Numbers, Symbols, Banners/slide${i}.jpg`,
        pptx: `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Arrows, Numbers, Symbols, Banners.pptx`,
        index: i
      }));
    } else if (category === "assets") {
      previews = Array.from({ length: 10 }, (_, i) => ({
        src: `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide${i + 1}.jpg`,
        pptx: `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Assets.pptx`,
        index: i + 1
      }));
    }

    previews.forEach(({ src, pptx, index }) => {
      const img = document.createElement("img");
      img.src = src;
      img.alt = `Slide ${index}`;
      img.classList.add("slide-preview");

      img.addEventListener("click", () => {
        Office.context.presentation.insertSlidesFromBase64Url(pptx, {
          formatting: Office.ImageFormatting.MatchDestination,
          position: 0
        }, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Error inserting slide:", result.error.message);
            alert("Unable to insert slide. Please try again.");
          }
        });
      });

      slidesContainer.appendChild(img);
    });
  }

  imageCategory.addEventListener("change", () => renderImages(imageCategory.value));
  slideCategory.addEventListener("change", () => renderSlidePreviews(slideCategory.value));

  renderImages("backgrounds");
  renderSlidePreviews("arrows");
});

