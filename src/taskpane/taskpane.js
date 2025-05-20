Office.onReady(() => {
  document.getElementById('imageCategory').addEventListener('change', renderImages);
  document.getElementById('slideCategory').addEventListener('change', renderSlides);
  renderImages();
  renderSlides();
});

// Image rendering
function renderImages() {
  const container = document.getElementById('imageContainer');
  container.innerHTML = '';
  const category = document.getElementById('imageCategory').value;

  let images = [];

  if (category === 'backgrounds') {
    images = ['background 1.png', 'background 2.png'].map(name =>
      `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/backgrounds/${name}`
    );
  } else if (category === 'halfpage') {
    images = [1, 2, 3, 4, 5, 6].map(i =>
      `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/Images/half page ${i}.jpg`
    );
  } else if (category === 'thin') {
    images = [1, 2, 3, 4, 5, 6].map(i =>
      `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/Images/thin image ${i}.jpg`
    );
  }

  images.forEach(url => {
    const img = document.createElement('img');
    img.src = url;
    img.draggable = true;
    img.addEventListener('dragstart', (e) => {
      e.dataTransfer.setData('text/plain', url);
    });
    container.appendChild(img);
  });
}

// Slides rendering
function renderSlides() {
  const slidesContainer = document.getElementById('slidesContainer');
  slidesContainer.innerHTML = '';
  const selection = document.getElementById('slideCategory').value;

  let previews = [];
  let pptxUrl = '';

  if (selection === 'arrows') {
    pptxUrl = 'https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Arrows, Numbers, Symbols, Banners.pptx';
    previews = [1, 2].map(i =>
      `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Arrows, Numbers, Symbols, Banners/slide${i}.jpg`
    );
  } else if (selection === 'assets') {
    pptxUrl = 'https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides/Assets.pptx';
    previews = Array.from({ length: 10 }, (_, i) =>
      `https://nepa-ab.github.io/Nepa-Templates-Add-in/src/slides-previews/Assets/slide${i + 1}.jpg`
    );
  }

  previews.forEach((url, index) => {
    const img = document.createElement('img');
    img.src = url;
    img.title = `Insert slide ${index + 1}`;
    img.addEventListener('click', async () => {
      await insertSlideFromPptx(pptxUrl, index);
    });
    slidesContainer.appendChild(img);
  });
}

// Placeholder for inserting slide (to be implemented later)
async function insertSlideFromPptx(pptxUrl, slideIndex) {
  console.log(`Would insert slide ${slideIndex + 1} from ${pptxUrl}`);
  alert('Slide insertion will be available in a future version. For now, please copy and paste the desired slide manually.');
}
