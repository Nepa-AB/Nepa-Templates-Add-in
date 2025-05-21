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
  let baseUrl = "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/backgrounds/";

  if (category === "half" || category === "thin") {
    baseUrl = "https://nepa-ab.github.io/Nepa-Templates-Add-in/src/Images/";
  }

  images[category].forEach((imgName) => {
    const wrapper = document.createElement("div");
    wrapper.className = "image-card";

    const img = document.createElement("img");
    img.src = baseUrl + encodeURIComponent(imgName);
    img.alt = imgName;

    const insertBtn = document.createElement("button");
    insertBtn.className = "insert-btn";
    insertBtn.textContent = "Insert";
    insertBtn.onclick = () => insertImageToActiveSlide(img.src);

    wrapper.appendChild(img);
    wrapper.appendChild(insertBtn);

    container.appendChild(wrapper);
  });
}

function showNotification(message) {
  const note = document.getElementById("notification");
  if (!note) return;
  note.innerText = message;
  note.style.display = "block";
  setTimeout(() => { note.style.display = "none"; }, 2500);
}

async function insertImageToActiveSlide(imgUrl) {
  try {
    // Fetch the image as a blob
    const response = await fetch(imgUrl);
    const blob = await response.blob();

    // Read blob as base64
    const base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const result = reader.result;
        const base64String = result.split(',')[1];
        resolve(base64String);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });

    // Choose correct MIME type
    let mimeType = "image/png";
    if (imgUrl.toLowerCase().endsWith(".jpg") || imgUrl.toLowerCase().endsWith(".jpeg")) {
      mimeType = "image/jpeg";
    }
    const dataUri = `data:${mimeType};base64,${base64}`;

    // Insert using Office.js (cross-platform, robust)
    Office.context.document.setSelectedDataAsync(
      dataUri,
      { coercionType: Office.CoercionType.Image },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          showNotification("Image inserted!");
        } else {
          showNotification("Insert failed: " + asyncResult.error.message);
        }
      }
    );
  } catch (error) {
    console.error(error);
    showNotification("Failed to insert image. See console for details.");
  }
}

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

      // Insert slide from file as first slide (index 0)
      await slides.insertFromFile(fileUrl, 0, { slideStart: slideNumber, slideEnd: slideNumber });
      await context.sync();
    });
    showNotification(`Inserted slide ${slideNumber}`);
  } catch (error) {
    console.error(error);
    showNotification("Failed to insert slide. See console for details.");
  }
}
