const VIEW_WEB = 'web';
const VIEW_LIBRARY = 'library';
const STORAGE_KEY_MANUAL_LIBRARY = 'manualImageLibrary';
const STORAGE_KEY_DOMAIN_GALLERIES = 'domainImageGalleries';
const MAX_MANUAL_IMAGES = 80;
const MANUAL_GALLERY_ID = 'manual-gallery';

const elements = {
  popupRoot: document.getElementById('popupRoot'),
  logoLibraryOverlay: document.getElementById('logoLibraryOverlay'),
  logoLibraryTitleIcon: document.getElementById('logoLibraryTitleIcon'),
  logoLibraryTitleText: document.getElementById('logoLibraryTitleText'),
  logoLibraryCloseBtn: document.getElementById('logoLibraryCloseBtn'),
  logoUploadBtn: document.getElementById('logoUploadBtn'),
  selectAllWebBtn: document.getElementById('selectAllWebBtn'),
  selectAllWebBtnIcon: document.getElementById('selectAllWebBtnIcon'),
  selectAllWebBtnText: document.getElementById('selectAllWebBtnText'),
  galleryBackBtn: document.getElementById('galleryBackBtn'),
  thumbColsSliderWrap: document.getElementById('thumbColsSliderWrap'),
  thumbColsSlider: document.getElementById('thumbColsSlider'),
  logoUploadInput: document.getElementById('logoUploadInput'),
  logoLibraryDropHint: document.getElementById('logoLibraryDropHint'),
  logoLibraryGrid: document.getElementById('logoLibraryGrid'),
  userPage: document.getElementById('userPage'),
  aboutPage: document.getElementById('aboutPage'),
  pageUrlBtn: document.getElementById('pageUrlBtn'),
  pageLibraryBtn: document.getElementById('pageLibraryBtn'),
  pageUserBtn: document.getElementById('pageUserBtn'),
  pageAboutBtn: document.getElementById('pageAboutBtn'),
  saveGalleryFabBtn: document.getElementById('saveGalleryFabBtn'),
  aboutVersion: document.getElementById('aboutVersion'),
  aboutSignatureTitle: document.getElementById('aboutSignatureTitle'),
  feedbackToast: document.getElementById('feedbackToast'),
  feedbackToastIcon: document.getElementById('feedbackToastIcon'),
  feedbackToastText: document.getElementById('feedbackToastText')
};

const state = {
  activePage: 'url',
  galleryView: VIEW_WEB,
  activeTabId: null,
  activeTabUrl: '',
  webItems: [],
  selectedWebIds: new Set(),
  webLoading: false,
  manualItems: [],
  selectedManualIds: new Set(),
  savedGalleries: [],
  galleryListMode: 'list',
  activeGalleryId: '',
  thumbnailColumns: 4,
  toastTimer: null,
  fileDragDepth: 0,
  fileDragActive: false
};

void init();

async function init() {
  bindEvents();
  setThumbnailColumns(4);
  renderLocalizedCopy();
  renderVersion();
  await loadManualLibrary();
  await loadDomainGalleries();
  await resolveActiveTabContext();
  setActivePage('url');
  await loadWebGallery();
}

function bindEvents() {
  elements.pageUrlBtn.addEventListener('click', () => {
    state.galleryView = VIEW_WEB;
    setActivePage('url');
    if (state.webItems.length === 0 && !state.webLoading) {
      void loadWebGallery();
    }
  });

  elements.pageLibraryBtn.addEventListener('click', () => {
    state.galleryView = VIEW_LIBRARY;
    state.galleryListMode = 'list';
    state.activeGalleryId = '';
    setActivePage('library');
  });

  elements.pageUserBtn.addEventListener('click', () => {
    setActivePage('user');
  });

  elements.pageAboutBtn.addEventListener('click', () => {
    setActivePage('about');
  });

  if (elements.saveGalleryFabBtn instanceof HTMLButtonElement) {
    elements.saveGalleryFabBtn.addEventListener('click', () => {
      void saveSelectedWebImagesToDomainGallery();
    });
  }

  if (elements.galleryBackBtn instanceof HTMLButtonElement) {
    elements.galleryBackBtn.addEventListener('click', () => {
      state.galleryListMode = 'list';
      state.activeGalleryId = '';
      renderActiveLogoLibraryGrid();
    });
  }

  if (elements.selectAllWebBtn instanceof HTMLButtonElement) {
    elements.selectAllWebBtn.addEventListener('click', () => {
      toggleSelectAllWebImages();
    });
  }

  if (elements.thumbColsSlider instanceof HTMLInputElement) {
    const onColumnsChange = () => {
      setThumbnailColumns(Number(elements.thumbColsSlider.value));
      renderActiveLogoLibraryGrid();
    };
    elements.thumbColsSlider.addEventListener('input', onColumnsChange);
    elements.thumbColsSlider.addEventListener('change', onColumnsChange);
  }

  elements.logoUploadBtn.addEventListener('click', () => {
    elements.logoUploadInput.click();
  });

  elements.logoUploadInput.addEventListener('change', (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) {
      return;
    }
    const files = input.files ? Array.from(input.files) : [];
    input.value = '';
    if (files.length === 0) {
      return;
    }
    void addManualImages(files);
  });

  elements.logoLibraryGrid.addEventListener('click', (event) => {
    const target = event.target;
    if (!(target instanceof Element)) {
      return;
    }

    const backButton = target.closest('button[data-gallery-back]');
    if (backButton) {
      state.galleryListMode = 'list';
      state.activeGalleryId = '';
      renderActiveLogoLibraryGrid();
      return;
    }

    const openGalleryButton = target.closest('button[data-gallery-open-id]');
    if (openGalleryButton && !openGalleryButton.disabled) {
      const galleryId = String(openGalleryButton.dataset.galleryOpenId || '').trim();
      if (!galleryId) {
        return;
      }
      state.galleryListMode = 'detail';
      state.activeGalleryId = galleryId;
      renderActiveLogoLibraryGrid();
      return;
    }

    const highlightButton = target.closest('button[data-highlight-source-url]');
    if (highlightButton && !highlightButton.disabled) {
      event.preventDefault();
      event.stopPropagation();
      const sourceUrl = String(highlightButton.dataset.highlightSourceUrl || '').trim();
      if (sourceUrl) {
        void highlightSourceImageOnPage(sourceUrl);
      } else {
        showFeedbackToast('No es pot marcar aquesta imatge', 'error');
      }
      return;
    }

    const deleteButton = target.closest('button[data-manual-delete-id]');
    if (deleteButton && !deleteButton.disabled) {
      const id = String(deleteButton.dataset.manualDeleteId || '');
      if (!id) {
        return;
      }
      void removeManualImageById(id);
      return;
    }

    const selectButton = target.closest('button.logo-library-item-select');
    if (!selectButton || selectButton.disabled) {
      return;
    }
    const imageId = String(selectButton.dataset.imageId || '').trim();
    if (!imageId) {
      return;
    }
    const isManual = selectButton.dataset.imageManual === '1';
    toggleImageSelection(imageId, { manual: isManual });
  });

  elements.popupRoot.addEventListener('dragenter', (event) => {
    if (!isImageFileDragEvent(event) || state.activePage !== 'url') {
      return;
    }
    event.preventDefault();
    state.fileDragDepth += 1;
    if (!state.fileDragActive) {
      state.fileDragActive = true;
      state.galleryView = VIEW_LIBRARY;
      state.galleryListMode = 'detail';
      state.activeGalleryId = MANUAL_GALLERY_ID;
      renderActiveLogoLibraryGrid();
    }
    setDropHintVisible(true);
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = 'copy';
    }
  });

  elements.popupRoot.addEventListener('dragover', (event) => {
    if (!isImageFileDragEvent(event) || state.activePage !== 'url') {
      return;
    }
    event.preventDefault();
    if (!state.fileDragActive) {
      state.fileDragActive = true;
      state.fileDragDepth = Math.max(1, state.fileDragDepth);
      state.galleryView = VIEW_LIBRARY;
      state.galleryListMode = 'detail';
      state.activeGalleryId = MANUAL_GALLERY_ID;
      renderActiveLogoLibraryGrid();
      setDropHintVisible(true);
    }
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = 'copy';
    }
  });

  elements.popupRoot.addEventListener('dragleave', (event) => {
    if (!state.fileDragActive || !isImageFileDragEvent(event) || state.activePage !== 'url') {
      return;
    }
    event.preventDefault();
    state.fileDragDepth = Math.max(0, state.fileDragDepth - 1);
    const related = event.relatedTarget;
    const stillInside = related instanceof Node && elements.popupRoot.contains(related);
    if (!stillInside && state.fileDragDepth <= 0) {
      resetDropState();
    }
  });

  elements.popupRoot.addEventListener('drop', (event) => {
    if (!isImageFileDragEvent(event) || state.activePage !== 'url') {
      return;
    }
    event.preventDefault();
    const imageFiles = getDroppedImageFiles(event.dataTransfer);
    resetDropState();
    state.galleryView = VIEW_LIBRARY;
    state.galleryListMode = 'detail';
    state.activeGalleryId = MANUAL_GALLERY_ID;
    renderActiveLogoLibraryGrid();
    if (imageFiles.length === 0) {
      showFeedbackToast('Nomes es permeten fitxers d\'imatge', 'error');
      return;
    }
    void addManualImages(imageFiles);
  });

  elements.logoLibraryCloseBtn.addEventListener('click', () => {
    // In gallery-only mode we keep this hidden and non-functional.
  });
}

function setActivePage(page) {
  const safe = page === 'user' || page === 'about' || page === 'library' ? page : 'url';
  state.activePage = safe;

  const isGallery = safe === 'url' || safe === 'library';
  elements.popupRoot.classList.toggle('gallery-tab-active', isGallery);

  elements.logoLibraryOverlay.classList.toggle('hidden', !isGallery);
  elements.logoLibraryOverlay.setAttribute('aria-hidden', isGallery ? 'false' : 'true');
  elements.userPage.classList.toggle('hidden', safe !== 'user');
  elements.aboutPage.classList.toggle('hidden', safe !== 'about');

  setNavButtonState(elements.pageUrlBtn, safe === 'url');
  setNavButtonState(elements.pageLibraryBtn, safe === 'library');
  setNavButtonState(elements.pageUserBtn, safe === 'user');
  setNavButtonState(elements.pageAboutBtn, safe === 'about');

  if (isGallery) {
    renderActiveLogoLibraryGrid();
    return;
  }
  updateUploadButtonVisibility();
  updateSaveGalleryFabVisibility();
  updateSelectAllWebButtonState();
}

function setNavButtonState(button, active) {
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }
  button.classList.toggle('active', active);
  button.setAttribute('aria-pressed', active ? 'true' : 'false');
}

function renderVersion() {
  if (!elements.aboutVersion) {
    return;
  }
  const manifest = chrome.runtime.getManifest();
  elements.aboutVersion.textContent = `Version ${manifest.version}`;
}

function renderLocalizedCopy() {
  if (!elements.aboutSignatureTitle) {
    return;
  }
  const localized = chrome.i18n.getMessage('aboutSignaturePrefix');
  elements.aboutSignatureTitle.textContent = localized || 'una aplicació de';
}

function updateGalleryHeader() {
  const showingWeb = state.galleryView === VIEW_WEB;
  if (showingWeb) {
    elements.logoLibraryTitleText.textContent = 'Imatges web';
    elements.logoLibraryTitleIcon.innerHTML = '<i class="bi bi-globe2" aria-hidden="true"></i>';
    return;
  }

  if (state.galleryListMode === 'detail' && state.activeGalleryId) {
    if (state.activeGalleryId === MANUAL_GALLERY_ID) {
      elements.logoLibraryTitleText.textContent = 'Galeria manual';
    } else {
      const gallery = state.savedGalleries.find(
        (entry) => String(entry?.id || '').trim() === String(state.activeGalleryId || '').trim()
      );
      elements.logoLibraryTitleText.textContent = String(gallery?.domain || 'Galeria');
    }
  } else {
    elements.logoLibraryTitleText.textContent = 'Galeries';
  }
  elements.logoLibraryTitleIcon.innerHTML = '<i class="bi bi-images" aria-hidden="true"></i>';
}

function updateSaveGalleryFabVisibility() {
  if (!(elements.saveGalleryFabBtn instanceof HTMLButtonElement)) {
    return;
  }
  const visible = state.activePage === 'url' && state.galleryView === VIEW_WEB;
  elements.saveGalleryFabBtn.classList.toggle('hidden', !visible);
  elements.saveGalleryFabBtn.disabled = !visible || state.webLoading || state.webItems.length === 0;
}

function updateUploadButtonVisibility() {
  if (!(elements.logoUploadBtn instanceof HTMLButtonElement)) {
    return;
  }
  const visible =
    state.galleryView === VIEW_LIBRARY &&
    (state.activePage === 'library' || state.activePage === 'url');
  elements.logoUploadBtn.classList.toggle('hidden', !visible);
  elements.logoUploadBtn.disabled = !visible;
}

function areAllWebImagesSelected() {
  if (state.webItems.length === 0) {
    return false;
  }
  for (const item of state.webItems) {
    const id = String(item?.id || '').trim();
    if (!id || !state.selectedWebIds.has(id)) {
      return false;
    }
  }
  return true;
}

function updateSelectAllWebButtonState() {
  if (!(elements.selectAllWebBtn instanceof HTMLButtonElement)) {
    return;
  }
  const visible = state.activePage === 'url' && state.galleryView === VIEW_WEB;
  elements.selectAllWebBtn.classList.toggle('hidden', !visible);
  if (!visible) {
    return;
  }

  const canInteract = !state.webLoading && state.webItems.length > 0;
  const allSelected = canInteract && areAllWebImagesSelected();
  const text = allSelected ? 'Deseleccionar todo' : 'Seleccionar todo';

  elements.selectAllWebBtn.disabled = !canInteract;
  elements.selectAllWebBtn.classList.toggle('active', allSelected);
  elements.selectAllWebBtn.setAttribute('aria-pressed', allSelected ? 'true' : 'false');
  elements.selectAllWebBtn.setAttribute('aria-label', text);
  if (elements.selectAllWebBtnText) {
    elements.selectAllWebBtnText.textContent = text;
  }
  if (elements.selectAllWebBtnIcon) {
    elements.selectAllWebBtnIcon.innerHTML = allSelected
      ? '<i class="bi bi-check2-square" aria-hidden="true"></i>'
      : '<i class="bi bi-square" aria-hidden="true"></i>';
  }
}

function updateThumbnailSliderVisibility() {
  const wrap = elements.thumbColsSliderWrap;
  if (!(wrap instanceof HTMLElement)) {
    return;
  }
  const albumOpen =
    state.galleryView === VIEW_LIBRARY &&
    state.galleryListMode === 'detail' &&
    Boolean(String(state.activeGalleryId || '').trim());
  const webGallery = state.galleryView === VIEW_WEB;
  const sliderVisible = webGallery || albumOpen;
  wrap.classList.toggle('hidden', !sliderVisible);
  if (elements.thumbColsSlider instanceof HTMLInputElement) {
    elements.thumbColsSlider.disabled = !sliderVisible;
  }
  if (elements.galleryBackBtn instanceof HTMLButtonElement) {
    elements.galleryBackBtn.classList.toggle('hidden', !albumOpen);
    elements.galleryBackBtn.disabled = !albumOpen;
  }
  elements.popupRoot.classList.toggle('library-album-open', albumOpen);
}

function toggleSelectAllWebImages() {
  if (state.webLoading || state.webItems.length === 0) {
    return;
  }
  if (areAllWebImagesSelected()) {
    state.selectedWebIds.clear();
    renderActiveLogoLibraryGrid();
    return;
  }
  state.selectedWebIds = new Set(
    state.webItems.map((item) => String(item?.id || '').trim()).filter(Boolean)
  );
  renderActiveLogoLibraryGrid();
}

function setThumbnailColumns(columns) {
  const safe = columns === 1 || columns === 2 || columns === 3 ? columns : 4;
  state.thumbnailColumns = safe;
  elements.logoLibraryGrid.style.setProperty('--gallery-cols', String(safe));
  elements.logoLibraryGrid.dataset.cols = String(safe);
  syncThumbnailColumnSlider();
}

function syncThumbnailColumnSlider() {
  if (!(elements.thumbColsSlider instanceof HTMLInputElement)) {
    return;
  }
  elements.thumbColsSlider.value = String(state.thumbnailColumns);
}

function renderActiveLogoLibraryGrid() {
  updateGalleryHeader();
  updateUploadButtonVisibility();
  updateSaveGalleryFabVisibility();
  updateSelectAllWebButtonState();
  updateThumbnailSliderVisibility();
  if (state.galleryView === VIEW_WEB) {
    renderWebGrid();
    return;
  }
  renderLibraryGrid();
}

function renderWebGrid() {
  const grid = elements.logoLibraryGrid;
  grid.dataset.galleryMode = 'web';
  grid.replaceChildren();
  grid.classList.remove('loading-state');

  if (state.webLoading) {
    grid.classList.add('loading-state');
    const loading = document.createElement('div');
    loading.className = 'logo-library-loading';
    loading.setAttribute('role', 'status');
    loading.setAttribute('aria-live', 'polite');

    const spinner = document.createElement('span');
    spinner.className = 'logo-library-loading-spinner';
    spinner.setAttribute('aria-hidden', 'true');

    const title = document.createElement('span');
    title.className = 'logo-library-loading-title';
    title.textContent = 'Carregant imatges';

    const subtitle = document.createElement('span');
    subtitle.className = 'logo-library-loading-subtitle';
    subtitle.textContent = 'Analitzant web activa';

    loading.appendChild(spinner);
    loading.appendChild(title);
    loading.appendChild(subtitle);
    grid.appendChild(loading);
    return;
  }

  if (state.webItems.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'logo-library-empty';
    empty.textContent = 'No s\'han trobat imatges a la web activa';
    grid.appendChild(empty);
    return;
  }

  grid.appendChild(createWebSizeLegend());
  for (const item of state.webItems) {
    grid.appendChild(createImageCard(item));
  }
}

function renderLibraryGrid() {
  if (state.galleryListMode === 'detail' && state.activeGalleryId) {
    renderGalleryDetailGrid();
    return;
  }
  renderGalleryListGrid();
}

function getGalleryCollections() {
  const collections = [];

  if (state.manualItems.length > 0) {
    const latestManual = Math.max(...state.manualItems.map((item) => Number(item.createdAt) || 0), 0);
    collections.push({
      id: MANUAL_GALLERY_ID,
      title: 'Manual',
      subtitle: 'Pujades manuals',
      count: state.manualItems.length,
      updatedAt: latestManual,
      type: 'manual'
    });
  }

  for (const gallery of state.savedGalleries) {
    const domain = String(gallery?.domain || '').trim();
    if (!domain) {
      continue;
    }
    collections.push({
      id: String(gallery.id || '').trim(),
      title: domain,
      subtitle: 'Web guardada',
      count: Array.isArray(gallery.items) ? gallery.items.length : 0,
      updatedAt: Number(gallery.updatedAt) || 0,
      type: 'domain'
    });
  }

  collections.sort((a, b) => Number(b.updatedAt) - Number(a.updatedAt));
  return collections;
}

function renderGalleryListGrid() {
  const grid = elements.logoLibraryGrid;
  grid.dataset.galleryMode = 'library-list';
  grid.replaceChildren();
  grid.classList.remove('loading-state');

  const collections = getGalleryCollections();
  if (collections.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'logo-library-empty';
    empty.textContent = 'Encara no hi ha galeries guardades';
    grid.appendChild(empty);
    return;
  }

  for (const entry of collections) {
    grid.appendChild(createGalleryListCard(entry));
  }
}

function renderGalleryDetailGrid() {
  const grid = elements.logoLibraryGrid;
  grid.dataset.galleryMode = 'library-detail';
  grid.replaceChildren();
  grid.classList.remove('loading-state');

  if (state.activeGalleryId === MANUAL_GALLERY_ID) {
    if (state.manualItems.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'logo-library-empty';
      empty.textContent = 'No tens imatges manuals. Puja fitxers.';
      grid.appendChild(empty);
      return;
    }
    for (const item of state.manualItems) {
      grid.appendChild(createImageCard(item, { manual: true }));
    }
    return;
  }

  const gallery = state.savedGalleries.find(
    (entry) => String(entry?.id || '').trim() === String(state.activeGalleryId || '').trim()
  );
  if (!gallery) {
    state.galleryListMode = 'list';
    state.activeGalleryId = '';
    renderGalleryListGrid();
    return;
  }
  if (!Array.isArray(gallery.items) || gallery.items.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'logo-library-empty';
    empty.textContent = 'Aquesta galeria no te imatges';
    grid.appendChild(empty);
    return;
  }

  for (const item of gallery.items) {
    grid.appendChild(createStoredGalleryCard(item));
  }
}

function createGalleryListCard(entry) {
  const card = document.createElement('button');
  card.type = 'button';
  card.className = 'saved-gallery-entry';
  card.dataset.galleryOpenId = String(entry.id || '').trim();

  const title = document.createElement('span');
  title.className = 'saved-gallery-entry-title';
  title.textContent = String(entry.title || 'Galeria');

  const subtitle = document.createElement('span');
  subtitle.className = 'saved-gallery-entry-subtitle';
  subtitle.textContent = String(entry.subtitle || '');

  const count = document.createElement('span');
  count.className = 'saved-gallery-entry-count';
  count.textContent = `${Number(entry.count) || 0} imatges`;

  card.appendChild(title);
  card.appendChild(subtitle);
  card.appendChild(count);
  return card;
}

function createStoredGalleryCard(item) {
  const card = document.createElement('div');
  card.className = 'logo-library-item stored-gallery-item';

  const thumbWrap = createThumbWrap(item);
  const name = document.createElement('span');
  name.className = 'logo-library-name';
  name.textContent = String(item.name || 'Imatge');

  card.appendChild(thumbWrap);
  card.appendChild(name);

  if (item?.sourceUrl) {
    const highlightButton = document.createElement('button');
    highlightButton.type = 'button';
    highlightButton.className = 'logo-library-item-locate';
    highlightButton.dataset.highlightSourceUrl = String(item.sourceUrl || '').trim();
    highlightButton.setAttribute('aria-label', `Marcar en web ${item.name}`);
    highlightButton.title = 'Marcar en web';
    highlightButton.innerHTML = '<i class="bi bi-eye" aria-hidden="true"></i>';
    card.appendChild(highlightButton);
  }

  return card;
}

function isItemSelected(item, options = {}) {
  const itemId = String(item?.id || '').trim();
  if (!itemId) {
    return false;
  }
  if (options.manual) {
    return state.selectedManualIds.has(itemId);
  }
  return state.selectedWebIds.has(itemId);
}

function toggleImageSelection(itemId, options = {}) {
  const safeId = String(itemId || '').trim();
  if (!safeId) {
    return;
  }
  const selectedSet = options.manual ? state.selectedManualIds : state.selectedWebIds;
  if (selectedSet.has(safeId)) {
    selectedSet.delete(safeId);
  } else {
    selectedSet.add(safeId);
  }
  renderActiveLogoLibraryGrid();
}

function createImageCard(item, options = {}) {
  const selected = isItemSelected(item, options);
  const card = document.createElement('div');
  card.className = 'logo-library-item';
  card.classList.toggle('active', selected);

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'logo-library-item-select';
  button.setAttribute('role', 'option');
  button.setAttribute('aria-selected', selected ? 'true' : 'false');
  button.dataset.imageId = String(item?.id || '');
  button.dataset.imageManual = options.manual ? '1' : '0';
  button.title = item.name;

  const thumbWrap = createThumbWrap(item, options);
  const name = document.createElement('span');
  name.className = 'logo-library-name';
  name.textContent = item.name;

  button.appendChild(thumbWrap);
  button.appendChild(name);
  card.appendChild(button);

  if (options.manual) {
    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'logo-library-item-delete';
    deleteButton.dataset.manualDeleteId = item.id;
    deleteButton.setAttribute('aria-label', `Eliminar ${item.name}`);
    deleteButton.title = 'Eliminar imatge';
    deleteButton.innerHTML = '<i class="bi bi-trash3" aria-hidden="true"></i>';
    card.appendChild(deleteButton);
  } else if (item?.sourceUrl && item?.canHighlight) {
    const highlightButton = document.createElement('button');
    highlightButton.type = 'button';
    highlightButton.className = 'logo-library-item-locate';
    highlightButton.dataset.highlightSourceUrl = String(item.sourceUrl || '').trim();
    highlightButton.setAttribute('aria-label', `Marcar en web ${item.name}`);
    highlightButton.title = 'Marcar en web';
    highlightButton.innerHTML = '<i class="bi bi-eye" aria-hidden="true"></i>';
    card.appendChild(highlightButton);
  }

  return card;
}

function createThumbWrap(item, options = {}) {
  const wrap = document.createElement('div');
  wrap.className = 'logo-library-thumb-wrap';

  const image = document.createElement('img');
  image.className = 'logo-library-thumb';
  image.src = item.dataUrl;
  image.alt = item.name;
  image.loading = 'lazy';
  wrap.appendChild(image);

  if (!options.hideBadge) {
    const badge = createPixelBadge(item, options);
    if (badge) {
      wrap.appendChild(badge);
    }
  }
  return wrap;
}

function createWebSizeLegend() {
  const legend = document.createElement('div');
  legend.className = 'logo-library-size-legend';
  legend.setAttribute('role', 'note');
  legend.setAttribute('aria-label', 'Leyenda de tamaños');

  const entries = [
    {
      tone: 'neutral',
      text: 'Gris: tamaño no prefijado o deformable'
    },
    {
      tone: 'match',
      text: 'Verde: tamaño de imagen y web igual'
    },
    {
      tone: 'mismatch',
      text: 'Rojo: imagen mayor que tamaño web'
    }
  ];

  for (const entry of entries) {
    const item = document.createElement('span');
    item.className = `logo-library-size-legend-item ${entry.tone}`;

    const dot = document.createElement('span');
    dot.className = 'logo-library-size-legend-dot';
    dot.setAttribute('aria-hidden', 'true');

    const label = document.createElement('span');
    label.className = 'logo-library-size-legend-label';
    label.textContent = entry.text;

    item.appendChild(dot);
    item.appendChild(label);
    legend.appendChild(item);
  }

  return legend;
}

function createPixelBadge(item, options = {}) {
  const safeWidth = Number.isFinite(Number(item?.width)) ? Math.round(Number(item.width)) : 0;
  const safeHeight = Number.isFinite(Number(item?.height)) ? Math.round(Number(item.height)) : 0;
  const safeHtmlWidth = Number.isFinite(Number(item?.htmlWidth)) ? Math.round(Number(item.htmlWidth)) : 0;
  const safeHtmlHeight = Number.isFinite(Number(item?.htmlHeight)) ? Math.round(Number(item.htmlHeight)) : 0;
  const htmlFixed = Boolean(item?.htmlFixed);
  const htmlResponsive = Boolean(item?.htmlResponsive);
  const htmlUnconstrained = Boolean(item?.htmlUnconstrained);
  const isManual = Boolean(options.manual);

  const hasOriginal = safeWidth > 0 && safeHeight > 0;
  const hasHtml = safeHtmlWidth > 0 && safeHtmlHeight > 0;
  const originalLabel = hasOriginal ? `${safeWidth}x${safeHeight}` : '?x?';

  const badge = document.createElement('span');
  badge.className = 'logo-library-pixels';
  if (isManual || !hasOriginal) {
    badge.textContent = `${originalLabel}px`;
    return badge;
  }

  if (!hasHtml || htmlResponsive || !htmlFixed || htmlUnconstrained) {
    badge.textContent = `${originalLabel}px`;
    return badge;
  }

  const matchesOriginal = safeHtmlWidth === safeWidth && safeHtmlHeight === safeHeight;
  if (matchesOriginal) {
    badge.classList.add('match');
    badge.textContent = `${originalLabel}px`;
    return badge;
  }

  const imageIsBiggerThanWeb = safeWidth > safeHtmlWidth || safeHeight > safeHtmlHeight;
  if (!imageIsBiggerThanWeb) {
    badge.textContent = `${originalLabel}px`;
    return badge;
  }

  const htmlLabel = `${safeHtmlWidth}x${safeHtmlHeight}`;
  badge.classList.add('mismatch', 'two-lines');
  const line1 = document.createElement('span');
  line1.className = 'logo-library-pixels-line';
  line1.innerHTML = `<i class="bi bi-globe2" aria-hidden="true"></i><span>${htmlLabel}px</span>`;

  const line2 = document.createElement('span');
  line2.className = 'logo-library-pixels-line';
  line2.innerHTML = `<i class="bi bi-file-earmark-image" aria-hidden="true"></i><span>${originalLabel}px</span>`;

  badge.appendChild(line1);
  badge.appendChild(line2);
  return badge;
}

async function resolveActiveTabContext() {
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const activeTab = tabs[0];
    state.activeTabId = Number.isFinite(Number(activeTab?.id)) ? Number(activeTab.id) : null;
    state.activeTabUrl = String(activeTab?.url || '').trim();
  } catch (error) {
    console.error('Active tab lookup failed', error);
    state.activeTabId = null;
    state.activeTabUrl = '';
  }
}

async function loadWebGallery() {
  if (!state.activeTabId || !state.activeTabUrl) {
    state.webItems = [];
    state.selectedWebIds = new Set();
    renderActiveLogoLibraryGrid();
    return;
  }

  state.webLoading = true;
  renderActiveLogoLibraryGrid();

  try {
    const candidates = await collectWebImageCandidates(state.activeTabId, state.activeTabUrl);
    const normalized = [];
    const seenData = new Set();
    let index = 1;

    for (const candidate of candidates) {
      const asset = await normalizeRemoteImage(candidate.url);
      if (!asset || !asset.dataUrl || seenData.has(asset.dataUrl)) {
        continue;
      }
      const sourceUrl = String(candidate.url || '').trim();
      seenData.add(asset.dataUrl);
      normalized.push({
        id: `web-${Date.now()}-${index}`,
        name: extractImageFileNameFromUrl(sourceUrl, index),
        dataUrl: asset.dataUrl,
        fileSize: Math.max(0, Number(asset.fileSize) || 0),
        width: asset.width,
        height: asset.height,
        htmlWidth: Math.max(0, Number(candidate.htmlWidth) || 0),
        htmlHeight: Math.max(0, Number(candidate.htmlHeight) || 0),
        htmlFixed: Boolean(candidate.htmlFixed),
        htmlResponsive: Boolean(candidate.htmlResponsive),
        htmlUnconstrained: Boolean(candidate.htmlUnconstrained),
        canHighlight: Boolean(candidate.htmlLocatable),
        sourceUrl,
        sourcePath: extractSourcePathFromUrl(sourceUrl)
      });
      index += 1;
    }

    state.webItems = normalized;
    state.selectedWebIds = new Set(normalized.map((item) => String(item.id || '').trim()).filter(Boolean));
    if (normalized.length === 0) {
      showFeedbackToast('No s\'han trobat imatges web', 'error');
    }
  } catch (error) {
    console.error('Web gallery load failed', error);
    state.webItems = [];
    state.selectedWebIds = new Set();
    showFeedbackToast('Error carregant imatges web', 'error');
  } finally {
    state.webLoading = false;
    renderActiveLogoLibraryGrid();
  }
}

async function saveSelectedWebImagesToDomainGallery() {
  if (state.galleryView !== VIEW_WEB) {
    return;
  }

  const selected = state.webItems.filter((item) => state.selectedWebIds.has(String(item.id || '').trim()));
  if (selected.length === 0) {
    showFeedbackToast('Selecciona imatges per desar', 'error');
    return;
  }

  let domain = '';
  let pageOrigin = '';
  try {
    const parsed = new URL(String(state.activeTabUrl || ''));
    domain = String(parsed.hostname || '').trim().toLowerCase();
    pageOrigin = String(parsed.origin || '').trim();
  } catch {
    domain = '';
  }
  if (!domain) {
    showFeedbackToast('No s\'ha pogut detectar el domini', 'error');
    return;
  }

  const galleryId = `domain:${domain}`;
  const now = Date.now();
  const existing = state.savedGalleries.find((entry) => String(entry?.id || '') === galleryId);
  const mergedBySource = new Map();

  if (existing && Array.isArray(existing.items)) {
    for (const item of existing.items) {
      const key = String(item?.sourceUrl || item?.id || '').trim();
      if (!key) {
        continue;
      }
      mergedBySource.set(key, { ...item });
    }
  }

  let inserted = 0;
  for (const item of selected) {
    const sourceUrl = String(item.sourceUrl || '').trim();
    const itemId = String(item.id || '').trim();
    const mergeKey = sourceUrl || itemId;
    if (!mergeKey) {
      continue;
    }
    if (!mergedBySource.has(mergeKey)) {
      inserted += 1;
    }

    mergedBySource.set(mergeKey, {
      id: mergeKey,
      name: sanitizeDisplayName(item.name || 'Imatge', 'Imatge'),
      dataUrl: String(item.dataUrl || ''),
      sourceUrl,
      sourcePath: String(item.sourcePath || '').trim(),
      fileSize: Math.max(0, Number(item.fileSize) || 0),
      width: Math.max(0, Number(item.width) || 0),
      height: Math.max(0, Number(item.height) || 0),
      htmlWidth: Math.max(0, Number(item.htmlWidth) || 0),
      htmlHeight: Math.max(0, Number(item.htmlHeight) || 0),
      htmlFixed: Boolean(item.htmlFixed),
      htmlResponsive: Boolean(item.htmlResponsive),
      htmlUnconstrained: Boolean(item.htmlUnconstrained),
      canHighlight: Boolean(item.canHighlight),
      savedAt: now
    });
  }

  const updatedGallery = {
    id: galleryId,
    domain,
    pageOrigin,
    updatedAt: now,
    items: Array.from(mergedBySource.values()).sort((a, b) => Number(b.savedAt) - Number(a.savedAt))
  };

  const others = state.savedGalleries.filter((entry) => String(entry?.id || '') !== galleryId);
  state.savedGalleries = [updatedGallery, ...others].sort((a, b) => Number(b.updatedAt) - Number(a.updatedAt));
  await persistDomainGalleries();

  state.galleryView = VIEW_LIBRARY;
  state.galleryListMode = 'detail';
  state.activeGalleryId = galleryId;
  renderActiveLogoLibraryGrid();
  showFeedbackToast(`${inserted} imatges desades a ${domain}`);
}

async function highlightSourceImageOnPage(sourceUrl) {
  const safeUrl = String(sourceUrl || '').trim();
  if (!safeUrl || !state.activeTabId) {
    showFeedbackToast('No es pot marcar aquesta imatge', 'error');
    return;
  }

  try {
    const result = await chrome.scripting.executeScript({
      target: { tabId: state.activeTabId },
      func: (targetUrl) => {
        const styleId = 'lobotinify-highlight-style';
        const markAttr = 'data-lobotinify-highlight';
        const overlayAttr = 'data-lobotinify-highlight-overlay';
        const markClass = 'lobotinify-highlighted-image';
        const overlayClass = 'lobotinify-highlight-overlay';
        const arrowClass = 'lobotinify-highlight-arrow';
        const timerKey = '__lobotinify_highlight_timer__';
        const base = document.baseURI || location.href;

        const normalizeUrl = (value) => {
          const raw = String(value || '').trim();
          if (!raw) {
            return '';
          }
          try {
            return new URL(raw, base).href;
          } catch {
            return '';
          }
        };
        const buildUrlVariants = (value) => {
          const href = normalizeUrl(value);
          if (!href) {
            return [];
          }
          try {
            const parsed = new URL(href);
            const noHash = `${parsed.origin}${parsed.pathname}${parsed.search}`;
            const noSearchAndHash = `${parsed.origin}${parsed.pathname}`;
            return [href, noHash, noSearchAndHash];
          } catch {
            return [href];
          }
        };

        const clearMarks = () => {
          const previous = document.querySelectorAll(`[${markAttr}="1"]`);
          for (const node of previous) {
            node.classList.remove(markClass);
            node.removeAttribute(markAttr);
          }
          const overlays = document.querySelectorAll(`[${overlayAttr}="1"]`);
          for (const overlay of overlays) {
            overlay.remove();
          }
        };

        const normalizedTarget = normalizeUrl(targetUrl);
        if (!normalizedTarget) {
          clearMarks();
          return { count: 0 };
        }
        const targetVariants = new Set(buildUrlVariants(normalizedTarget));
        const matchesTarget = (value) => {
          for (const candidate of buildUrlVariants(value)) {
            if (targetVariants.has(candidate)) {
              return true;
            }
          }
          return false;
        };

        let styleNode = document.getElementById(styleId);
        if (!styleNode) {
          styleNode = document.createElement('style');
          styleNode.id = styleId;
          styleNode.textContent = `
            .${markClass} {
              box-shadow: 0 0 0 1px rgba(239, 68, 68, 0.52) !important;
              border-radius: 4px !important;
              transition: box-shadow 0.16s ease !important;
            }
            .${overlayClass} {
              position: absolute !important;
              box-sizing: border-box !important;
              border: 3px solid #ef4444 !important;
              border-radius: 5px !important;
              box-shadow: 0 0 0 2px rgba(239, 68, 68, 0.36), 0 0 16px rgba(239, 68, 68, 0.45) !important;
              pointer-events: none !important;
              z-index: 2147483647 !important;
            }
            .${arrowClass} {
              position: absolute !important;
              width: 46px !important;
              height: 46px !important;
              pointer-events: none !important;
              z-index: 2147483647 !important;
            }
            .${arrowClass} svg {
              width: 100% !important;
              height: 100% !important;
              display: block !important;
              filter: drop-shadow(0 0 6px rgba(239, 68, 68, 0.6)) !important;
            }
          `;
          const host = document.head || document.documentElement;
          host.appendChild(styleNode);
        }

        clearMarks();

        const marked = [];
        const drawOverlay = (node) => {
          if (!(node instanceof Element)) {
            return;
          }
          const rect = node.getBoundingClientRect();
          const width = Math.max(0, Math.round(rect.width));
          const height = Math.max(0, Math.round(rect.height));
          if (width < 2 || height < 2) {
            return;
          }
          const overlay = document.createElement('div');
          overlay.className = overlayClass;
          overlay.setAttribute(overlayAttr, '1');
          overlay.style.left = `${Math.round(rect.left + window.scrollX)}px`;
          overlay.style.top = `${Math.round(rect.top + window.scrollY)}px`;
          overlay.style.width = `${width}px`;
          overlay.style.height = `${height}px`;
          const arrow = document.createElement('div');
          arrow.className = arrowClass;
          arrow.setAttribute(overlayAttr, '1');
          arrow.style.left = `${Math.round(rect.left + window.scrollX - 24)}px`;
          arrow.style.top = `${Math.round(rect.top + window.scrollY - 24)}px`;

          const svgNs = 'http://www.w3.org/2000/svg';
          const arrowSvg = document.createElementNS(svgNs, 'svg');
          arrowSvg.setAttribute('viewBox', '0 0 64 64');
          arrowSvg.setAttribute('aria-hidden', 'true');

          const shaft = document.createElementNS(svgNs, 'line');
          shaft.setAttribute('x1', '10');
          shaft.setAttribute('y1', '54');
          shaft.setAttribute('x2', '46');
          shaft.setAttribute('y2', '18');
          shaft.setAttribute('fill', 'none');
          shaft.setAttribute('stroke', '#ef4444');
          shaft.setAttribute('stroke-width', '9');
          shaft.setAttribute('stroke-linecap', 'round');
          shaft.setAttribute('stroke-linejoin', 'round');

          const headA = document.createElementNS(svgNs, 'line');
          headA.setAttribute('x1', '46');
          headA.setAttribute('y1', '18');
          headA.setAttribute('x2', '41');
          headA.setAttribute('y2', '34');
          headA.setAttribute('fill', 'none');
          headA.setAttribute('stroke', '#ef4444');
          headA.setAttribute('stroke-width', '8');
          headA.setAttribute('stroke-linecap', 'round');
          headA.setAttribute('stroke-linejoin', 'round');

          const headB = document.createElementNS(svgNs, 'line');
          headB.setAttribute('x1', '46');
          headB.setAttribute('y1', '18');
          headB.setAttribute('x2', '30');
          headB.setAttribute('y2', '23');
          headB.setAttribute('fill', 'none');
          headB.setAttribute('stroke', '#ef4444');
          headB.setAttribute('stroke-width', '8');
          headB.setAttribute('stroke-linecap', 'round');
          headB.setAttribute('stroke-linejoin', 'round');

          arrowSvg.appendChild(shaft);
          arrowSvg.appendChild(headA);
          arrowSvg.appendChild(headB);
          arrow.appendChild(arrowSvg);

          const host = document.body || document.documentElement;
          host.appendChild(overlay);
          host.appendChild(arrow);
          if (typeof arrow.animate === 'function') {
            arrow.animate(
              [
                { transform: 'translate(-3px, -3px) rotate(90deg) scale(0.96)', opacity: 0.86 },
                { transform: 'translate(4px, 4px) rotate(90deg) scale(1)', opacity: 1 },
                { transform: 'translate(-3px, -3px) rotate(90deg) scale(0.96)', opacity: 0.86 }
              ],
              { duration: 880, iterations: Infinity, easing: 'ease-in-out' }
            );
          }
        };
        const markNode = (node) => {
          if (!(node instanceof Element) || marked.includes(node)) {
            return;
          }
          node.classList.add(markClass);
          node.setAttribute(markAttr, '1');
          marked.push(node);
          drawOverlay(node);
        };

        const matchSrcSet = (srcset) => {
          const raw = String(srcset || '').trim();
          if (!raw) {
            return false;
          }
          for (const chunk of raw.split(',')) {
            const [candidate] = chunk.trim().split(/\s+/u);
            if (matchesTarget(candidate)) {
              return true;
            }
          }
          return false;
        };

        const images = document.querySelectorAll('img');
        for (const node of images) {
          const currentSrc = node.currentSrc;
          const src = node.getAttribute('src');
          const srcsetMatch = matchSrcSet(node.getAttribute('srcset'));
          if (matchesTarget(currentSrc) || matchesTarget(src) || srcsetMatch) {
            markNode(node);
          }
        }

        const sources = document.querySelectorAll('source');
        for (const node of sources) {
          const src = node.getAttribute('src');
          const srcsetMatch = matchSrcSet(node.getAttribute('srcset'));
          if (matchesTarget(src) || srcsetMatch) {
            const picture = node.closest('picture');
            const rendered = picture?.querySelector('img') || picture || node;
            markNode(rendered);
          }
        }

        const inlineBackground = document.querySelectorAll('[style*="background-image"]');
        for (const node of inlineBackground) {
          const style = String(node.getAttribute('style') || '');
          const regex = /url\((['"]?)(.*?)\1\)/giu;
          let match;
          while ((match = regex.exec(style)) !== null) {
            if (matchesTarget(match[2])) {
              markNode(node);
              break;
            }
          }
        }

        if (window[timerKey]) {
          clearTimeout(window[timerKey]);
        }
        window[timerKey] = setTimeout(() => {
          clearMarks();
          window[timerKey] = null;
        }, 5000);

        if (marked.length > 0) {
          marked[0].scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
        }
        return { count: marked.length };
      },
      args: [safeUrl]
    });

    const count = Math.max(0, Number(result?.[0]?.result?.count) || 0);
    if (count > 0) {
      showFeedbackToast(`${count} imatge(s) marcada(es)`);
      return;
    }
    showFeedbackToast('No s\'ha pogut marcar a la web', 'error');
  } catch (error) {
    console.error('Highlight image failed', error);
    showFeedbackToast('No s\'ha pogut marcar a la web', 'error');
  }
}

async function collectWebImageCandidates(tabId, tabUrl) {
  const result = await chrome.scripting.executeScript({
    target: { tabId },
    func: () => {
      const found = new Map();
      const base = document.baseURI || location.href;
      const ABSOLUTE_LENGTH_RE = /^-?\d+(\.\d+)?(?:px|pt|pc|in|cm|mm|q|rem|em|ch|ex)$/iu;
      const RESPONSIVE_LENGTH_RE =
        /%|vw|vh|vmin|vmax|svw|svh|lvw|lvh|dvw|dvh|calc\(|min\(|max\(|clamp\(/iu;
      const clampDimension = (value) => {
        const numeric = Number(value);
        if (!Number.isFinite(numeric) || numeric <= 0) {
          return 0;
        }
        return Math.round(numeric);
      };
      const classifyLength = (value) => {
        const raw = String(value || '').trim().toLowerCase();
        if (!raw || raw === 'auto' || raw === 'initial' || raw === 'inherit') {
          return { fixed: false, responsive: false };
        }
        if (RESPONSIVE_LENGTH_RE.test(raw)) {
          return { fixed: false, responsive: true };
        }
        if (ABSOLUTE_LENGTH_RE.test(raw) || /^-?\d+(\.\d+)?$/u.test(raw)) {
          return { fixed: true, responsive: false };
        }
        return { fixed: false, responsive: false };
      };
      const readElementSizingHints = (node) => {
        if (!(node instanceof Element)) {
          return { fixed: false, responsive: false };
        }

        const rawWidthAttr = String(node.getAttribute('width') || '').trim();
        const rawHeightAttr = String(node.getAttribute('height') || '').trim();
        const widthAttr = Number(rawWidthAttr);
        const heightAttr = Number(rawHeightAttr);
        const attrWidthFixed = Number.isFinite(widthAttr) && widthAttr > 0;
        const attrHeightFixed = Number.isFinite(heightAttr) && heightAttr > 0;

        const styleWidth = classifyLength(node.style?.width || '');
        const styleHeight = classifyLength(node.style?.height || '');

        const widthFixed = attrWidthFixed || styleWidth.fixed;
        const heightFixed = attrHeightFixed || styleHeight.fixed;
        return {
          fixed: widthFixed && heightFixed,
          responsive: styleWidth.responsive || styleHeight.responsive
        };
      };
      const resolveSizingState = (node) => {
        let current = node instanceof Element ? node : null;
        let fixedFound = false;
        let responsiveFound = false;

        for (let depth = 0; current && depth <= 4; depth += 1) {
          const hints = readElementSizingHints(current);
          fixedFound = fixedFound || hints.fixed;
          responsiveFound = responsiveFound || hints.responsive;
          current = current.parentElement;
        }

        const fixed = fixedFound && !responsiveFound;
        return {
          fixed,
          responsive: responsiveFound,
          unconstrained: !fixedFound && !responsiveFound
        };
      };
      const createMetric = (width, height) => ({
        width,
        height,
        area: width * height,
        perimeter: width + height
      });
      const pickBetterMetric = (current, incoming) => {
        if (!current) {
          return incoming;
        }
        if (incoming.area > current.area) {
          return incoming;
        }
        if (incoming.area < current.area) {
          return current;
        }
        if (incoming.perimeter > current.perimeter) {
          return incoming;
        }
        return current;
      };
      const upsert = (
        resolved,
        dimensions = {
          width: 0,
          height: 0,
          fixed: false,
          responsive: false,
          unconstrained: false,
          locatable: false
        }
      ) => {
        const width = clampDimension(dimensions.width);
        const height = clampDimension(dimensions.height);
        const fixed = Boolean(dimensions.fixed);
        const responsive = Boolean(dimensions.responsive);
        const unconstrained = Boolean(dimensions.unconstrained);
        const locatable = Boolean(dimensions.locatable);

        let existing = found.get(resolved);
        if (!existing) {
          existing = {
            url: resolved,
            hasFixed: false,
            hasResponsive: false,
            hasUnconstrained: false,
            hasLocatable: false,
            bestFixed: null,
            bestAny: null
          };
          found.set(resolved, existing);
        }

        existing.hasFixed = existing.hasFixed || fixed;
        existing.hasResponsive = existing.hasResponsive || responsive;
        existing.hasUnconstrained = existing.hasUnconstrained || unconstrained;
        existing.hasLocatable = existing.hasLocatable || locatable;

        if (width > 0 && height > 0) {
          const metric = createMetric(width, height);
          existing.bestAny = pickBetterMetric(existing.bestAny, metric);
          if (fixed) {
            existing.bestFixed = pickBetterMetric(existing.bestFixed, metric);
          }
        }
      };

      const getRectSize = (node) => {
        if (!(node instanceof Element)) {
          return { width: 0, height: 0 };
        }
        const rect = node.getBoundingClientRect();
        return {
          width: clampDimension(rect.width),
          height: clampDimension(rect.height)
        };
      };

      const pushUrl = (value, dimensions = { width: 0, height: 0 }) => {
        const raw = String(value || '').trim();
        if (!raw) {
          return;
        }
        try {
          const resolved = new URL(raw, base).href;
          if (resolved.startsWith('http://') || resolved.startsWith('https://')) {
            upsert(resolved, dimensions);
          }
        } catch {
          // ignore malformed url
        }
      };

      const pushSrcSet = (srcset, dimensions = { width: 0, height: 0 }) => {
        const raw = String(srcset || '').trim();
        if (!raw) {
          return;
        }
        const entries = raw.split(',');
        for (const entry of entries) {
          const [candidate] = entry.trim().split(/\s+/u);
          pushUrl(candidate, dimensions);
        }
      };
      const toResultEntry = (entry) => {
        const preferred = entry.bestFixed || entry.bestAny;
        return {
          url: entry.url,
          htmlWidth: preferred ? preferred.width : 0,
          htmlHeight: preferred ? preferred.height : 0,
          htmlFixed: entry.hasFixed && !entry.hasResponsive,
          htmlResponsive: entry.hasResponsive,
          htmlUnconstrained: entry.hasUnconstrained && !entry.hasFixed && !entry.hasResponsive,
          htmlLocatable: entry.hasLocatable
        };
      };

      const metadataNodes = document.querySelectorAll(
        'meta[property="og:image"],meta[name="twitter:image"],link[rel~="icon"],link[rel="apple-touch-icon"]'
      );
      for (const node of metadataNodes) {
        pushUrl(node.getAttribute('content'));
        pushUrl(node.getAttribute('href'));
      }

      const imageNodes = document.querySelectorAll('img');
      for (const node of imageNodes) {
        const sizing = resolveSizingState(node);
        const dimensions = {
          ...getRectSize(node),
          fixed: sizing.fixed,
          responsive: sizing.responsive,
          unconstrained: sizing.unconstrained,
          locatable: true
        };
        pushUrl(node.currentSrc, dimensions);
        pushUrl(node.getAttribute('src'), dimensions);
        pushSrcSet(node.getAttribute('srcset'), dimensions);
      }

      const sourceNodes = document.querySelectorAll('source');
      for (const node of sourceNodes) {
        const parent = node.closest('picture') || node.parentElement;
        const sizing = resolveSizingState(parent || node);
        const dimensions = {
          ...getRectSize(parent),
          fixed: sizing.fixed,
          responsive: sizing.responsive,
          unconstrained: sizing.unconstrained,
          locatable: true
        };
        pushUrl(node.getAttribute('src'), dimensions);
        pushSrcSet(node.getAttribute('srcset'), dimensions);
      }

      const inlineBackground = document.querySelectorAll('[style*="background-image"]');
      for (const node of inlineBackground) {
        const sizing = resolveSizingState(node);
        const dimensions = {
          ...getRectSize(node),
          fixed: sizing.fixed,
          responsive: sizing.responsive,
          unconstrained: sizing.unconstrained,
          locatable: true
        };
        const style = String(node.getAttribute('style') || '');
        const regex = /url\((['"]?)(.*?)\1\)/giu;
        let match;
        while ((match = regex.exec(style)) !== null) {
          pushUrl(match[2], dimensions);
        }
      }

      return Array.from(found.values(), (entry) => toResultEntry(entry));
    }
  });

  const fromPageRaw = Array.isArray(result?.[0]?.result) ? result[0].result : [];
  const fromPage = fromPageRaw
    .map((entry) => ({
      url: String(entry?.url || '').trim(),
      htmlWidth: Math.max(0, Math.round(Number(entry?.htmlWidth) || 0)),
      htmlHeight: Math.max(0, Math.round(Number(entry?.htmlHeight) || 0)),
      htmlFixed: Boolean(entry?.htmlFixed),
      htmlResponsive: Boolean(entry?.htmlResponsive),
      htmlUnconstrained: Boolean(entry?.htmlUnconstrained),
      htmlLocatable: Boolean(entry?.htmlLocatable)
    }))
    .filter((entry) => Boolean(entry.url));

  const favicons = [256, 128, 64].map((size) =>
    ({
      url: chrome.runtime.getURL(`/_favicon/?pageUrl=${encodeURIComponent(tabUrl)}&size=${size}`),
      htmlWidth: 0,
      htmlHeight: 0,
      htmlFixed: false,
      htmlResponsive: false,
      htmlUnconstrained: false,
      htmlLocatable: false
    })
  );

  const merged = [...fromPage, ...favicons];
  const uniqueByUrl = new Map();

  for (const entry of merged) {
    const safeUrl = String(entry?.url || '').trim();
    if (!safeUrl) {
      continue;
    }
    const candidate = {
      url: safeUrl,
      htmlWidth: Math.max(0, Number(entry?.htmlWidth) || 0),
      htmlHeight: Math.max(0, Number(entry?.htmlHeight) || 0),
      htmlFixed: Boolean(entry?.htmlFixed),
      htmlResponsive: Boolean(entry?.htmlResponsive),
      htmlUnconstrained: Boolean(entry?.htmlUnconstrained),
      htmlLocatable: Boolean(entry?.htmlLocatable)
    };
    const existing = uniqueByUrl.get(safeUrl);
    if (!existing) {
      uniqueByUrl.set(safeUrl, candidate);
      continue;
    }
    const candidateArea = candidate.htmlWidth * candidate.htmlHeight;
    const existingArea = existing.htmlWidth * existing.htmlHeight;
    if (candidateArea > existingArea) {
      const next = { ...candidate };
      next.htmlResponsive = existing.htmlResponsive || candidate.htmlResponsive;
      next.htmlUnconstrained = existing.htmlUnconstrained || candidate.htmlUnconstrained;
      next.htmlLocatable = existing.htmlLocatable || candidate.htmlLocatable;
      next.htmlFixed = Boolean(next.htmlFixed) && !next.htmlResponsive;
      if (next.htmlResponsive || next.htmlFixed) {
        next.htmlUnconstrained = false;
      }
      uniqueByUrl.set(safeUrl, next);
      continue;
    }
    if (candidateArea === existingArea && candidate.htmlFixed && !existing.htmlFixed) {
      const next = { ...candidate };
      next.htmlResponsive = existing.htmlResponsive || candidate.htmlResponsive;
      next.htmlUnconstrained = existing.htmlUnconstrained || candidate.htmlUnconstrained;
      next.htmlLocatable = existing.htmlLocatable || candidate.htmlLocatable;
      next.htmlFixed = Boolean(next.htmlFixed) && !next.htmlResponsive;
      if (next.htmlResponsive || next.htmlFixed) {
        next.htmlUnconstrained = false;
      }
      uniqueByUrl.set(safeUrl, next);
      continue;
    }

    existing.htmlResponsive = existing.htmlResponsive || candidate.htmlResponsive;
    existing.htmlUnconstrained = existing.htmlUnconstrained || candidate.htmlUnconstrained;
    existing.htmlLocatable = existing.htmlLocatable || candidate.htmlLocatable;
    existing.htmlFixed = Boolean(existing.htmlFixed) && !existing.htmlResponsive;
    if (existing.htmlResponsive || existing.htmlFixed) {
      existing.htmlUnconstrained = false;
    }
  }

  return Array.from(uniqueByUrl.values());
}

async function normalizeRemoteImage(url) {
  try {
    const response = await fetch(url, { cache: 'no-store', credentials: 'omit' });
    if (!response.ok) {
      return null;
    }
    const blob = await response.blob();
    return normalizeBlobToAsset(blob);
  } catch {
    return null;
  }
}

async function addManualImages(files) {
  const validFiles = Array.isArray(files) ? files.filter(Boolean) : [];
  if (validFiles.length === 0) {
    return;
  }

  const created = [];
  let skipped = 0;

  for (const file of validFiles) {
    try {
      const asset = await normalizeBlobToAsset(file);
      if (!asset?.dataUrl) {
        skipped += 1;
        continue;
      }
      created.push({
        id: `manual-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
        name: sanitizeManualName(file.name || 'Imatge manual'),
        dataUrl: asset.dataUrl,
        fileSize: Math.max(0, Number(asset.fileSize) || 0),
        width: asset.width,
        height: asset.height,
        htmlWidth: 0,
        htmlHeight: 0,
        htmlFixed: false,
        htmlResponsive: false,
        htmlUnconstrained: false,
        canHighlight: false,
        sourceUrl: '',
        sourcePath: String(file.webkitRelativePath || file.name || '').trim(),
        createdAt: Date.now() + created.length
      });
    } catch {
      skipped += 1;
    }
  }

  if (created.length === 0) {
    showFeedbackToast('No s\'ha pogut pujar cap imatge', 'error');
    return;
  }

  const previousSelected = new Set(state.selectedManualIds);
  state.manualItems = [...created, ...state.manualItems].slice(0, MAX_MANUAL_IMAGES);
  const validIds = new Set(state.manualItems.map((item) => String(item.id || '').trim()).filter(Boolean));
  const nextSelected = new Set();
  for (const id of previousSelected) {
    if (validIds.has(id)) {
      nextSelected.add(id);
    }
  }
  for (const item of created) {
    const createdId = String(item.id || '').trim();
    if (createdId && validIds.has(createdId)) {
      nextSelected.add(createdId);
    }
  }
  state.selectedManualIds = nextSelected;
  await persistManualLibrary();

  state.galleryView = VIEW_LIBRARY;
  state.galleryListMode = 'detail';
  state.activeGalleryId = MANUAL_GALLERY_ID;
  renderActiveLogoLibraryGrid();

  if (skipped > 0) {
    showFeedbackToast(`${created.length} pujades, ${skipped} descartades`);
    return;
  }
  showFeedbackToast(`${created.length} imatges pujades`);
}

async function removeManualImageById(id) {
  const before = state.manualItems.length;
  state.manualItems = state.manualItems.filter((item) => item.id !== id);
  if (state.manualItems.length === before) {
    return;
  }
  state.selectedManualIds.delete(String(id || '').trim());
  await persistManualLibrary();
  renderActiveLogoLibraryGrid();
  showFeedbackToast('Imatge eliminada');
}

async function normalizeBlobToAsset(blob) {
  if (!(blob instanceof Blob) || blob.size <= 0) {
    return null;
  }
  const mimeType = String(blob.type || '').toLowerCase();
  if (mimeType && !mimeType.startsWith('image/')) {
    return null;
  }

  const dataUrl = await blobToDataUrl(blob);
  if (!dataUrl.startsWith('data:image/')) {
    return null;
  }

  let width = 0;
  let height = 0;
  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImage(objectUrl);
    width = Math.max(1, Number(image.naturalWidth) || 1);
    height = Math.max(1, Number(image.naturalHeight) || 1);
  } catch {
    // Keep original bytes even if we cannot read dimensions for this format.
  } finally {
    URL.revokeObjectURL(objectUrl);
  }

  return { dataUrl, width, height, fileSize: Math.max(0, Number(blob.size) || 0) };
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('image_load_failed'));
    image.src = src;
  });
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error('blob_to_data_url_failed'));
    reader.readAsDataURL(blob);
  });
}

async function loadManualLibrary() {
  try {
    const stored = await chrome.storage.local.get({ [STORAGE_KEY_MANUAL_LIBRARY]: [] });
    state.manualItems = normalizeManualLibrary(stored[STORAGE_KEY_MANUAL_LIBRARY]);
    state.selectedManualIds = new Set(
      state.manualItems.map((item) => String(item.id || '').trim()).filter(Boolean)
    );
  } catch (error) {
    console.error('Manual library load failed', error);
    state.manualItems = [];
    state.selectedManualIds = new Set();
  }
}

async function loadDomainGalleries() {
  try {
    const stored = await chrome.storage.local.get({ [STORAGE_KEY_DOMAIN_GALLERIES]: [] });
    state.savedGalleries = normalizeDomainGalleries(stored[STORAGE_KEY_DOMAIN_GALLERIES]);
  } catch (error) {
    console.error('Domain galleries load failed', error);
    state.savedGalleries = [];
  }
}

async function persistDomainGalleries() {
  try {
    await chrome.storage.local.set({
      [STORAGE_KEY_DOMAIN_GALLERIES]: state.savedGalleries
    });
  } catch (error) {
    console.error('Domain galleries save failed', error);
    const message = String(error?.message || '').toLowerCase();
    if (message.includes('quota')) {
      showFeedbackToast('No s\'han pogut desar les galeries: limit d\'espai', 'error');
      return;
    }
    showFeedbackToast('No s\'han pogut desar les galeries', 'error');
  }
}

function normalizeDomainGalleries(raw) {
  if (!Array.isArray(raw)) {
    return [];
  }
  const galleries = [];
  const seen = new Set();

  for (const entry of raw) {
    if (!entry || typeof entry !== 'object') {
      continue;
    }
    const id = String(entry.id || '').trim();
    const domain = String(entry.domain || '').trim().toLowerCase();
    const pageOrigin = String(entry.pageOrigin || '').trim();
    const updatedAt = Number(entry.updatedAt) || Date.now();
    if (!id || !domain || seen.has(id)) {
      continue;
    }
    seen.add(id);

    const itemsRaw = Array.isArray(entry.items) ? entry.items : [];
    const items = [];
    const seenItemIds = new Set();
    for (const item of itemsRaw) {
      if (!item || typeof item !== 'object') {
        continue;
      }
      const itemId = String(item.id || item.sourceUrl || '').trim();
      const dataUrl = String(item.dataUrl || '').trim();
      if (!itemId || !dataUrl.startsWith('data:image/') || seenItemIds.has(itemId)) {
        continue;
      }
      seenItemIds.add(itemId);
      items.push({
        id: itemId,
        name: sanitizeDisplayName(item.name || 'Imatge', 'Imatge'),
        dataUrl,
        sourceUrl: String(item.sourceUrl || '').trim(),
        sourcePath: String(item.sourcePath || '').trim(),
        fileSize: Math.max(0, Number(item.fileSize) || 0),
        width: Math.max(0, Number(item.width) || 0),
        height: Math.max(0, Number(item.height) || 0),
        htmlWidth: Math.max(0, Number(item.htmlWidth) || 0),
        htmlHeight: Math.max(0, Number(item.htmlHeight) || 0),
        htmlFixed: Boolean(item.htmlFixed),
        htmlResponsive: Boolean(item.htmlResponsive),
        htmlUnconstrained: Boolean(item.htmlUnconstrained),
        canHighlight: Boolean(item.canHighlight),
        savedAt: Number(item.savedAt) || updatedAt
      });
    }

    items.sort((a, b) => Number(b.savedAt) - Number(a.savedAt));
    galleries.push({
      id,
      domain,
      pageOrigin,
      updatedAt,
      items
    });
  }

  galleries.sort((a, b) => Number(b.updatedAt) - Number(a.updatedAt));
  return galleries;
}

async function persistManualLibrary() {
  try {
    await chrome.storage.local.set({
      [STORAGE_KEY_MANUAL_LIBRARY]: state.manualItems.slice(0, MAX_MANUAL_IMAGES)
    });
  } catch (error) {
    console.error('Manual library save failed', error);
    showFeedbackToast('No s\'han pogut desar les imatges', 'error');
  }
}

function normalizeManualLibrary(raw) {
  if (!Array.isArray(raw)) {
    return [];
  }
  const items = [];
  const seen = new Set();

  for (const entry of raw) {
    if (!entry || typeof entry !== 'object') {
      continue;
    }
    const id = String(entry.id || '').trim();
    const dataUrl = String(entry.dataUrl || '').trim();
    const name = sanitizeManualName(entry.name || 'Imatge manual');
    const fileSize = Math.max(0, Number(entry.fileSize) || 0);
    const width = Math.max(0, Number(entry.width) || 0);
    const height = Math.max(0, Number(entry.height) || 0);
    const htmlWidth = Math.max(0, Number(entry.htmlWidth) || 0);
    const htmlHeight = Math.max(0, Number(entry.htmlHeight) || 0);
    const htmlFixed = Boolean(entry.htmlFixed);
    const htmlResponsive = Boolean(entry.htmlResponsive);
    const htmlUnconstrained = Boolean(entry.htmlUnconstrained);
    const canHighlight = Boolean(entry.canHighlight);
    const sourceUrl = String(entry.sourceUrl || '').trim();
    const sourcePath = String(entry.sourcePath || '').trim();
    const createdAt = Number(entry.createdAt) || Date.now();
    if (!id || !dataUrl.startsWith('data:image/') || seen.has(id)) {
      continue;
    }
    seen.add(id);
    items.push({
      id,
      name,
      dataUrl,
      fileSize,
      width,
      height,
      htmlWidth,
      htmlHeight,
      htmlFixed,
      htmlResponsive,
      htmlUnconstrained,
      canHighlight,
      sourceUrl,
      sourcePath,
      createdAt
    });
  }

  items.sort((a, b) => Number(b.createdAt) - Number(a.createdAt));
  return items.slice(0, MAX_MANUAL_IMAGES);
}

function sanitizeManualName(name) {
  return sanitizeDisplayName(name, 'Imatge manual');
}

function sanitizeDisplayName(name, fallback = 'Imatge') {
  const safe = String(name || '').trim();
  if (!safe) {
    const safeFallback = String(fallback || 'Imatge').trim();
    return safeFallback || 'Imatge';
  }
  return safe.length > 80 ? `${safe.slice(0, 79)}...` : safe;
}

function showFeedbackToast(message, type = 'success') {
  const text = String(message || '').trim();
  if (!text) {
    return;
  }

  elements.feedbackToastText.textContent = text;
  elements.feedbackToast.classList.toggle('error', type === 'error');
  elements.feedbackToastIcon.innerHTML =
    type === 'error'
      ? '<i class="bi bi-exclamation-triangle-fill" aria-hidden="true"></i>'
      : '<i class="bi bi-check-circle-fill" aria-hidden="true"></i>';

  elements.feedbackToast.classList.add('visible');
  elements.feedbackToast.setAttribute('aria-hidden', 'false');

  if (state.toastTimer) {
    clearTimeout(state.toastTimer);
  }

  state.toastTimer = setTimeout(() => {
    elements.feedbackToast.classList.remove('visible');
    elements.feedbackToast.setAttribute('aria-hidden', 'true');
  }, 2200);
}

function extractImageFileNameFromUrl(rawUrl, fallbackIndex = 0) {
  const fallback = `imagen-${Math.max(1, Number(fallbackIndex) || 1)}`;
  try {
    const parsed = new URL(String(rawUrl || '').trim());
    const segments = String(parsed.pathname || '')
      .split('/')
      .filter(Boolean);
    let fileName = segments.length > 0 ? decodeURIComponent(segments[segments.length - 1]) : '';

    if (!fileName || fileName === '_favicon') {
      const size = String(parsed.searchParams.get('size') || '').trim();
      fileName = size ? `favicon-${size}.png` : '';
    }

    if (!fileName) {
      fileName = fallback;
    }

    return sanitizeDisplayName(fileName, fallback);
  } catch {
    return sanitizeDisplayName(fallback, fallback);
  }
}

function extractSourcePathFromUrl(rawUrl) {
  try {
    const parsed = new URL(String(rawUrl || '').trim());
    return `${parsed.pathname}${parsed.search}${parsed.hash}`;
  } catch {
    return String(rawUrl || '').trim();
  }
}

function isImageFileLike(file) {
  if (!file) {
    return false;
  }
  const mime = String(file.type || '').toLowerCase();
  if (mime.startsWith('image/')) {
    return true;
  }
  const name = String(file.name || '').toLowerCase();
  return /\.(png|jpe?g|webp|gif|bmp|svg|ico|avif)$/iu.test(name);
}

function isImageFileDragEvent(event) {
  const transfer = event?.dataTransfer;
  if (!transfer) {
    return false;
  }
  const transferTypes = Array.from(transfer.types || []);
  if (!transferTypes.includes('Files')) {
    return false;
  }
  const items = Array.from(transfer.items || []);
  if (items.length > 0) {
    return items.some((item) => item?.kind === 'file');
  }
  const files = Array.from(transfer.files || []);
  if (files.length === 0) {
    return true;
  }
  return files.some((file) => isImageFileLike(file));
}

function getDroppedImageFiles(dataTransfer) {
  const files = Array.from(dataTransfer?.files || []);
  return files.filter((file) => isImageFileLike(file));
}

function setDropHintVisible(visible) {
  const nextVisible = Boolean(visible);
  elements.logoLibraryDropHint.classList.toggle('hidden', !nextVisible);
  elements.logoLibraryDropHint.setAttribute('aria-hidden', nextVisible ? 'false' : 'true');
}

function resetDropState() {
  state.fileDragDepth = 0;
  state.fileDragActive = false;
  setDropHintVisible(false);
}
