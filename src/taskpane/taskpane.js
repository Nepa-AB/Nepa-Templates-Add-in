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
    // Wikipedia PNG as a reliable public image
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
    const wrapper = document.createElement("div");
    wrapper.className = "image-card";

    const img = document.createElement("img");
    img.src = imgName.startsWith("http") ? imgName : baseUrl + encodeURIComponent(imgName);
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

// Robust: Convert Uint8Array to base64 (handles all bytes, avoids Unicode bugs)
function uint8ToBase64(u8Arr) {
  const CHUNK_SIZE = 0x8000; // 32k
  let index = 0;
  const length = u8Arr.length;
  let result = '';
  let slice;
  while (index < length) {
    slice = u8Arr.subarray(index, Math.min(index + CHUNK_SIZE, length));
    result += String.fromCharCode.apply(null, slice);
    index += CHUNK_SIZE;
  }
  return window.btoa(result);
}

// Fetch image as base64 using XHR (binary safe)
function fetchImageAsBase64(imgUrl, callback) {
  let mimeType = "image/png";
  if (imgUrl.toLowerCase().endsWith(".jpg") || imgUrl.toLowerCase().endsWith(".jpeg")) {
    mimeType = "image/jpeg";
  }
  const xhr = new XMLHttpRequest();
  xhr.open("GET", imgUrl, true);
  xhr.responseType = "arraybuffer";
  xhr.onload = function () {
    if (xhr.status === 200) {
      const uInt8Array = new Uint8Array(xhr.response);
      console.log("Fetched image byteLength:", uInt8Array.length, "MIME type:", mimeType);
      const base64 = uint8ToBase64(uInt8Array);
      const dataUri = `data:${mimeType};base64,${base64}`;
      console.log("Data URI (prefix):", dataUri.slice(0, 100) + "...");
      callback(null, dataUri, uInt8Array.length, mimeType);
    } else {
      callback(new Error("Image fetch failed: " + xhr.status), null, 0, mimeType);
    }
  };
  xhr.onerror = function () {
    callback(new Error("Image fetch network error"), null, 0, mimeType);
  };
  xhr.send();
}

function insertImageToActiveSlide(imgUrl) {
  fetchImageAsBase64(imgUrl, function(err, dataUri, byteLength, mimeType) {
    if (err) {
      showNotification("Fetch failed: " + err.message);
      console.error(err);
      return;
    }

    // For debugging: open Data URI in a new tab
    // window.open(dataUri, "_blank");

    Office.context.document.setSelectedDataAsync(
      dataUri,
      { coercionType: Office.CoercionType.Image },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          showNotification("Image inserted!");
        } else {
          showNotification("Insert failed: " + asyncResult.error.message);
          console.error("Insert failed:", asyncResult.error);
        }
      }
    );
  });
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
