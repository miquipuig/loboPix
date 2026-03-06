const VIEW_WEB = 'web';
const VIEW_MANUAL = 'manual';
const STORAGE_KEY_MANUAL_LIBRARY = 'manualImageLibrary';
const MAX_MANUAL_IMAGES = 80;
const MAX_WEB_IMAGES = 80;
const MAX_IMAGE_SIDE = 720;

const elements = {
  popupRoot: document.getElementById('popupRoot'),
  logoLibraryOverlay: document.getElementById('logoLibraryOverlay'),
  logoLibraryTitleIcon: document.getElementById('logoLibraryTitleIcon'),
  logoLibraryTitleText: document.getElementById('logoLibraryTitleText'),
  logoLibraryCloseBtn: document.getElementById('logoLibraryCloseBtn'),
  logoUploadBtn: document.getElementById('logoUploadBtn'),
  logoUsePageBtn: document.getElementById('logoUsePageBtn'),
  logoUsePageBtnIcon: document.getElementById('logoUsePageBtnIcon'),
  logoUsePageBtnText: document.getElementById('logoUsePageBtnText'),
  logoUploadInput: document.getElementById('logoUploadInput'),
  logoLibraryDropHint: document.getElementById('logoLibraryDropHint'),
  logoLibraryGrid: document.getElementById('logoLibraryGrid'),
  userPage: document.getElementById('userPage'),
  aboutPage: document.getElementById('aboutPage'),
  pageUrlBtn: document.getElementById('pageUrlBtn'),
  pageUserBtn: document.getElementById('pageUserBtn'),
  pageAboutBtn: document.getElementById('pageAboutBtn'),
  aboutVersion: document.getElementById('aboutVersion'),
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
  webLoading: false,
  manualItems: [],
  toastTimer: null,
  fileDragDepth: 0,
  fileDragActive: false
};

void init();

async function init() {
  bindEvents();
  renderVersion();
  await loadManualLibrary();
  await resolveActiveTabContext();
  setActivePage('url');
  await loadWebGallery();
}

function bindEvents() {
  elements.pageUrlBtn.addEventListener('click', () => {
    setActivePage('url');
    if (state.galleryView === VIEW_WEB && state.webItems.length === 0 && !state.webLoading) {
      void loadWebGallery();
    }
  });

  elements.pageUserBtn.addEventListener('click', () => {
    setActivePage('user');
  });

  elements.pageAboutBtn.addEventListener('click', () => {
    setActivePage('about');
  });

  elements.logoUsePageBtn.addEventListener('click', () => {
    if (state.galleryView === VIEW_WEB) {
      state.galleryView = VIEW_MANUAL;
      renderActiveLogoLibraryGrid();
      return;
    }
    state.galleryView = VIEW_WEB;
    renderActiveLogoLibraryGrid();
    void loadWebGallery();
  });

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
    const deleteButton = event.target.closest('button[data-manual-delete-id]');
    if (!deleteButton || deleteButton.disabled) {
      return;
    }
    const id = String(deleteButton.dataset.manualDeleteId || '');
    if (!id) {
      return;
    }
    void removeManualImageById(id);
  });

  elements.popupRoot.addEventListener('dragenter', (event) => {
    if (!isImageFileDragEvent(event) || state.activePage !== 'url') {
      return;
    }
    event.preventDefault();
    state.fileDragDepth += 1;
    if (!state.fileDragActive) {
      state.fileDragActive = true;
      state.galleryView = VIEW_MANUAL;
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
      state.galleryView = VIEW_MANUAL;
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
    state.galleryView = VIEW_MANUAL;
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
  const safe = page === 'user' || page === 'about' ? page : 'url';
  state.activePage = safe;

  const isGallery = safe === 'url';
  elements.popupRoot.classList.toggle('gallery-tab-active', isGallery);

  elements.logoLibraryOverlay.classList.toggle('hidden', !isGallery);
  elements.logoLibraryOverlay.setAttribute('aria-hidden', isGallery ? 'false' : 'true');
  elements.userPage.classList.toggle('hidden', safe !== 'user');
  elements.aboutPage.classList.toggle('hidden', safe !== 'about');

  setNavButtonState(elements.pageUrlBtn, safe === 'url');
  setNavButtonState(elements.pageUserBtn, safe === 'user');
  setNavButtonState(elements.pageAboutBtn, safe === 'about');

  if (isGallery) {
    renderActiveLogoLibraryGrid();
  }
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

function updateGalleryHeader() {
  const showingWeb = state.galleryView === VIEW_WEB;
  elements.logoLibraryTitleText.textContent = showingWeb ? 'Imatges web' : 'Galeria manual';
  elements.logoLibraryTitleIcon.innerHTML = showingWeb
    ? '<i class="bi bi-globe2" aria-hidden="true"></i>'
    : '<i class="bi bi-images" aria-hidden="true"></i>';
}

function updateToggleButtonCopy() {
  const showingWeb = state.galleryView === VIEW_WEB;
  const text = showingWeb ? 'Veure manuals' : 'Veure web';
  elements.logoUsePageBtnText.textContent = text;
  elements.logoUsePageBtn.setAttribute('aria-label', text);
  elements.logoUsePageBtnIcon.innerHTML = showingWeb
    ? '<i class="bi bi-images" aria-hidden="true"></i>'
    : '<i class="bi bi-globe2" aria-hidden="true"></i>';
}

function renderActiveLogoLibraryGrid() {
  updateGalleryHeader();
  updateToggleButtonCopy();
  if (state.galleryView === VIEW_WEB) {
    renderWebGrid();
    return;
  }
  renderManualGrid();
}

function renderWebGrid() {
  const grid = elements.logoLibraryGrid;
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

  for (const item of state.webItems) {
    grid.appendChild(createImageCard(item));
  }
}

function renderManualGrid() {
  const grid = elements.logoLibraryGrid;
  grid.replaceChildren();
  grid.classList.remove('loading-state');

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
}

function createImageCard(item, options = {}) {
  const card = document.createElement('div');
  card.className = 'logo-library-item';

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'logo-library-item-select';
  button.setAttribute('role', 'option');
  button.setAttribute('aria-selected', 'false');
  button.title = item.name;

  const thumbWrap = createThumbWrap(item);
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
  }

  return card;
}

function createThumbWrap(item) {
  const wrap = document.createElement('div');
  wrap.className = 'logo-library-thumb-wrap';

  const image = document.createElement('img');
  image.className = 'logo-library-thumb';
  image.src = item.dataUrl;
  image.alt = item.name;
  image.loading = 'lazy';
  wrap.appendChild(image);

  const badge = createPixelBadge(item.width, item.height);
  if (badge) {
    wrap.appendChild(badge);
  }
  return wrap;
}

function createPixelBadge(width, height) {
  const safeWidth = Number.isFinite(Number(width)) ? Math.round(Number(width)) : 0;
  const safeHeight = Number.isFinite(Number(height)) ? Math.round(Number(height)) : 0;
  if (safeWidth <= 0 || safeHeight <= 0) {
    return null;
  }
  const badge = document.createElement('span');
  badge.className = 'logo-library-pixels';
  badge.textContent = `${safeWidth}x${safeHeight}px`;
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

    for (const url of candidates) {
      if (normalized.length >= MAX_WEB_IMAGES) {
        break;
      }
      const asset = await normalizeRemoteImage(url);
      if (!asset || !asset.dataUrl || seenData.has(asset.dataUrl)) {
        continue;
      }
      seenData.add(asset.dataUrl);
      normalized.push({
        id: `web-${Date.now()}-${index}`,
        name: `${safeHostLabel(state.activeTabUrl)} · imatge ${index}`,
        dataUrl: asset.dataUrl,
        width: asset.width,
        height: asset.height
      });
      index += 1;
    }

    state.webItems = normalized;
    if (normalized.length === 0) {
      showFeedbackToast('No s\'han trobat imatges web', 'error');
    }
  } catch (error) {
    console.error('Web gallery load failed', error);
    state.webItems = [];
    showFeedbackToast('Error carregant imatges web', 'error');
  } finally {
    state.webLoading = false;
    renderActiveLogoLibraryGrid();
  }
}

async function collectWebImageCandidates(tabId, tabUrl) {
  const result = await chrome.scripting.executeScript({
    target: { tabId },
    func: () => {
      const found = new Set();
      const base = document.baseURI || location.href;

      const pushUrl = (value) => {
        const raw = String(value || '').trim();
        if (!raw) {
          return;
        }
        try {
          const resolved = new URL(raw, base).href;
          if (resolved.startsWith('http://') || resolved.startsWith('https://')) {
            found.add(resolved);
          }
        } catch {
          // ignore malformed url
        }
      };

      const pushSrcSet = (srcset) => {
        const raw = String(srcset || '').trim();
        if (!raw) {
          return;
        }
        const entries = raw.split(',');
        for (const entry of entries) {
          const [candidate] = entry.trim().split(/\s+/u);
          pushUrl(candidate);
        }
      };

      const selectors = [
        'meta[property="og:image"]',
        'meta[name="twitter:image"]',
        'link[rel~="icon"]',
        'link[rel="apple-touch-icon"]',
        'img[src]',
        'source[src]',
        'source[srcset]'
      ];

      for (const selector of selectors) {
        const nodes = document.querySelectorAll(selector);
        for (const node of nodes) {
          pushUrl(node.getAttribute('content'));
          pushUrl(node.getAttribute('href'));
          pushUrl(node.getAttribute('src'));
          pushSrcSet(node.getAttribute('srcset'));
        }
      }

      const inlineBackground = document.querySelectorAll('[style*="background-image"]');
      for (const node of inlineBackground) {
        const style = String(node.getAttribute('style') || '');
        const regex = /url\((['"]?)(.*?)\1\)/giu;
        let match;
        while ((match = regex.exec(style)) !== null) {
          pushUrl(match[2]);
        }
      }

      return Array.from(found);
    }
  });

  const fromPage = Array.isArray(result?.[0]?.result) ? result[0].result : [];
  const favicons = [256, 128, 64].map((size) =>
    chrome.runtime.getURL(`/_favicon/?pageUrl=${encodeURIComponent(tabUrl)}&size=${size}`)
  );

  const merged = [...fromPage, ...favicons];
  const unique = [];
  const seen = new Set();

  for (const url of merged) {
    const safe = String(url || '').trim();
    if (!safe || seen.has(safe)) {
      continue;
    }
    seen.add(safe);
    unique.push(safe);
  }

  return unique;
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
        width: asset.width,
        height: asset.height,
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

  state.manualItems = [...created, ...state.manualItems].slice(0, MAX_MANUAL_IMAGES);
  await persistManualLibrary();

  state.galleryView = VIEW_MANUAL;
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

  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImage(objectUrl);
    const width = Math.max(1, Number(image.naturalWidth) || 1);
    const height = Math.max(1, Number(image.naturalHeight) || 1);
    const scale = Math.min(1, MAX_IMAGE_SIDE / Math.max(width, height));
    const outputWidth = Math.max(1, Math.round(width * scale));
    const outputHeight = Math.max(1, Math.round(height * scale));

    const canvas = document.createElement('canvas');
    canvas.width = outputWidth;
    canvas.height = outputHeight;
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      return null;
    }
    ctx.imageSmoothingEnabled = true;
    ctx.drawImage(image, 0, 0, outputWidth, outputHeight);

    return {
      dataUrl: canvas.toDataURL('image/png'),
      width: outputWidth,
      height: outputHeight
    };
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('image_load_failed'));
    image.src = src;
  });
}

async function loadManualLibrary() {
  try {
    const stored = await chrome.storage.local.get({ [STORAGE_KEY_MANUAL_LIBRARY]: [] });
    state.manualItems = normalizeManualLibrary(stored[STORAGE_KEY_MANUAL_LIBRARY]);
  } catch (error) {
    console.error('Manual library load failed', error);
    state.manualItems = [];
  }
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
    const width = Math.max(0, Number(entry.width) || 0);
    const height = Math.max(0, Number(entry.height) || 0);
    const createdAt = Number(entry.createdAt) || Date.now();
    if (!id || !dataUrl.startsWith('data:image/') || seen.has(id)) {
      continue;
    }
    seen.add(id);
    items.push({ id, name, dataUrl, width, height, createdAt });
  }

  items.sort((a, b) => Number(b.createdAt) - Number(a.createdAt));
  return items.slice(0, MAX_MANUAL_IMAGES);
}

function sanitizeManualName(name) {
  const safe = String(name || '').trim();
  if (!safe) {
    return 'Imatge manual';
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

function safeHostLabel(rawUrl) {
  try {
    return new URL(rawUrl).hostname.replace(/^www\./iu, '') || 'web';
  } catch {
    return 'web';
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
