document.addEventListener("DOMContentLoaded", function () {
  const imageGrid = document.getElementById("image-grid");
  const categorySelect = document.getElementById("image-category");

  const imageCategories = {
    backgrounds: {
      path: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/backgrounds/",
      files: ["background 1.png", "background 2.png", "background 3.png", "background 4.png", "background 5.png", "background 6.png"]
    },
    half: {
      path: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/Images/",
      files: ["half page 1.jpg", "half page 2.jpg", "half page 3.jpg", "half page 4.jpg", "half page 5.jpg", "half page 6.jpg"]
    },
    thin: {
      path: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/Images/",
      files: ["thin image 1.jpg", "thin image 2.jpg", "thin image 3.jpg", "thin image 4.jpg", "thin image 5.jpg", "thin image 6.jpg"]
    }
  };

  function loadImages(category) {
    imageGrid.innerHTML = "";
    const { path, files } = imageCategories[category];

    files.forEach(file => {
      const img = document.createElement("img");
      img.src = path + file;
      img.alt = file;
      img.draggable = true;
      img.classList.add("draggable-image");

      img.addEventListener("dragstart", (e) => {
        e.dataTransfer.setData("text", img.src);
      });

      imageGrid.appendChild(img);
    });
  }

  categorySelect.addEventListener("change", () => {
    const selectedCategory = categorySelect.value;
    loadImages(selectedCategory);
  });

  // Load default
  loadImages("backgrounds");

  // Slides Section
  const slidePreviewMap = {
    "Arrows, Numbers, Symbols, Banners": {
      basePath: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Arrows, Numbers, Symbols, Banners/",
      count: 2,
      pptxUrl: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Arrows, Numbers, Symbols, Banners.pptx"
    },
    "Assets": {
      basePath: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/",
      count: 10,
      pptxUrl: "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Assets.pptx"
    }
  };

  const slidesDropdown = document.getElementById("slides-dropdown");
  const slidesContainer = document.getElementById("slides-previews");

  slidesDropdown.addEventListener("change", async (e) => {
    const selected = e.target.value;
    slidesContainer.innerHTML = "";

    if (slidePreviewMap[selected]) {
      const { basePath, count, pptxUrl } = slidePreviewMap[selected];
      for (let i = 1; i <= count; i++) {
        const img = document.createElement("img");
        img.src = `${basePath}slide${i}.jpg`;
        img.alt = `Slide ${i}`;
        img.classList.add("slide-preview");
        img.style.cursor = "pointer";
        img.dataset.slideIndex = i;
        img.dataset.pptxUrl = pptxUrl;
        img.dataset.category = selected;

        img.addEventListener("click", insertSlideFromPptx);
        slidesContainer.appendChild(img);
      }
    }
  });

  async function insertSlideFromPptx(event) {
    const slideIndex = parseInt(event.target.dataset.slideIndex, 10);
    const pptxUrl = event.target.dataset.pptxUrl;

    try {
      const response = await fetch(pptxUrl);
      const blob = await response.blob();
      const arrayBuffer = await blob.arrayBuffer();
      const fileBase64 = btoa(String.fromCharCode(...new Uint8Array(arrayBuffer)));

      await PowerPoint.run(async (context) => {
        const presentation = context.presentation;
        const slides = presentation.slides;
        const insertedSlide = slides.insertSlideFromBase64(fileBase64, PowerPoint.InsertSlideFormatting.useDestinationTheme);
        insertedSlide.insertAt(0);
        await context.sync();
      });
    } catch (error) {
      console.error("Failed to insert slide:", error);
    }
  }
});
