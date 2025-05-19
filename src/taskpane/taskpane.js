Office.onReady(() => {
  console.log("Office ready");

  const imageContainer = document.getElementById("image-container");
  const categorySelector = document.getElementById("image-category");

  const categories = {
    "backgrounds": {
      prefix: "backgrounds/background ",
      count: 2,
      ext: "png",
      alt: "Background"
    },
    "half-page": {
      prefix: "Images/half page ",
      count: 6,
      ext: "jpg",
      alt: "Half Page"
    },
    "thin-images": {
      prefix: "Images/thin image ",
      count: 6,
      ext: "jpg",
      alt: "Thin Image"
    }
  };

  function updateImages(categoryKey) {
    const category = categories[categoryKey];
    imageContainer.innerHTML = "";

    for (let i = 1; i <= category.count; i++) {
      const fileName = `${category.prefix}${i}.${category.ext}`;
      const fullUrl = `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/${fileName}`;

      const img = document.createElement("img");
      img.src = fullUrl;
      img.alt = `${category.alt} ${i}`;
      img.className = "background-image";
      img.width = 80;
      img.height = 60;
      img.setAttribute("draggable", "true");

      img.addEventListener("dragstart", (e) => {
        console.log("Dragging image:", fullUrl);
        e.dataTransfer.setData("text/uri-list", fullUrl);
        e.dataTransfer.setData("text/plain", fullUrl);
      });

      imageContainer.appendChild(img);
    }
  }

  categorySelector.addEventListener("change", (e) => {
    updateImages(e.target.value);
  });

  // Load default category
  updateImages(categorySelector.value);
});
