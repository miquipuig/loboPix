import { zipSync } from 'fflate';
import { encode as encodeJpeg } from '@jsquash/jpeg';
import { optimise as optimisePng } from '@jsquash/oxipng';
import { encode as encodeWebp } from '@jsquash/webp';
import UPNG from 'upng-js';

const VIEW_WEB = 'web';
const VIEW_LIBRARY = 'library';
const STORAGE_KEY_MANUAL_LIBRARY = 'manualImageLibrary';
const STORAGE_KEY_DOMAIN_GALLERIES = 'domainImageGalleries';
const MAX_MANUAL_IMAGES = 80;
const MANUAL_GALLERY_ID = 'manual-gallery';
const EXPORT_PRESET_ORIGINAL = 'original';
const EXPORT_PRESET_FAVICON = 'favicon';
const EXPORT_DEFAULT_FORMAT = 'image/webp';
const EXPORT_FORMAT_CANDIDATES = ['image/webp', 'image/png', 'image/jpeg'];
const EXPORT_DEFAULT_COMPRESSION = 72;
const EXPORT_MIN_QUALITY = 0.05;
const EXPORT_MAX_QUALITY = 1;
const EXPORT_CROP_MIN_RATIO = 0.05;
const EXPORT_BACKGROUND_TOLERANCE_DEFAULT = 0;
const EXPORT_BACKGROUND_TOLERANCE_MAX = 120;
const EXPORT_BACKGROUND_PROTECTION_DEFAULT = 0;
const EXPORT_BACKGROUND_PROTECTION_MAX = 100;
const EXPORT_BACKGROUND_TOLERANCE_ACTIVE_DEFAULT = 3;
const EXPORT_BACKGROUND_PROTECTION_ACTIVE_DEFAULT = 11;
const ZIP_MAX_COMPRESSION_LEVEL = 9;
const FAVICON_PRESET_SIZES = [16, 32, 48, 64, 128, 180, 192, 256, 512];
const EXPORT_RECOMMENDATION_CANDIDATES = [0, 8, 16, 24, 32, 40, 48, 56, 64, 72, 80, 88, 96, 100];
const EXPORT_RECOMMENDATION_MAX_EDGE = 256;
const EXPORT_COMPRESSION_LEVELS = ['conservative', 'optimal', 'aggressive'];
const EXPORT_DEFAULT_COMPRESSION_LEVEL = 'conservative';
const EXPORT_RECOMMENDATION_ERROR_THRESHOLD = {
  'image/png': 8.5,
  'image/webp': 6.4,
  'image/jpeg': 6.8
};
const exportCompressionRecommendationCache = new Map();
const exportFormatRecommendationCache = new Map();
const exportFormatAnalysisCache = new Map();

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
  saveGalleryFabGroup: document.getElementById('saveGalleryFabGroup'),
  saveGalleryFabBtn: document.getElementById('saveGalleryFabBtn'),
  saveGalleryAsFabBtn: document.getElementById('saveGalleryAsFabBtn'),
  saveAsModal: document.getElementById('saveAsModal'),
  saveAsNameInput: document.getElementById('saveAsNameInput'),
  saveAsCancelBtn: document.getElementById('saveAsCancelBtn'),
  saveAsConfirmBtn: document.getElementById('saveAsConfirmBtn'),
  manualUploadModal: document.getElementById('manualUploadModal'),
  manualUploadGallerySelect: document.getElementById('manualUploadGallerySelect'),
  manualUploadNameInput: document.getElementById('manualUploadNameInput'),
  manualUploadCancelBtn: document.getElementById('manualUploadCancelBtn'),
  manualUploadConfirmBtn: document.getElementById('manualUploadConfirmBtn'),
  exportWizardModal: document.getElementById('exportWizardModal'),
  exportWizardSourceName: document.getElementById('exportWizardSourceName'),
  exportWizardNameInput: document.getElementById('exportWizardNameInput'),
  exportWizardOptimalSummary: document.getElementById('exportWizardOptimalSummary'),
  exportWizardOptimalWebpValue: document.getElementById('exportWizardOptimalWebpValue'),
  exportWizardOptimalPngValue: document.getElementById('exportWizardOptimalPngValue'),
  exportWizardOptimalJpegValue: document.getElementById('exportWizardOptimalJpegValue'),
  exportWizardPreviewImg: document.getElementById('exportWizardPreviewImg'),
  exportWizardPreviewLoading: document.getElementById('exportWizardPreviewLoading'),
  exportCropToggleBtn: document.getElementById('exportCropToggleBtn'),
  exportTransparentBgBtn: document.getElementById('exportTransparentBgBtn'),
  exportCropStage: document.getElementById('exportCropStage'),
  exportCropEditorImg: document.getElementById('exportCropEditorImg'),
  exportCropBox: document.getElementById('exportCropBox'),
  exportTransparentBgPopover: document.getElementById('exportTransparentBgPopover'),
  exportTransparentBgSlider: document.getElementById('exportTransparentBgSlider'),
  exportTransparentBgSliderFill: document.getElementById('exportTransparentBgSliderFill'),
  exportTransparentBgSliderThumb: document.getElementById('exportTransparentBgSliderThumb'),
  exportTransparentBgProtectionSlider: document.getElementById('exportTransparentBgProtectionSlider'),
  exportTransparentBgProtectionSliderFill: document.getElementById('exportTransparentBgProtectionSliderFill'),
  exportTransparentBgProtectionSliderThumb: document.getElementById('exportTransparentBgProtectionSliderThumb'),
  exportWizardOriginalMeta: document.getElementById('exportWizardOriginalMeta'),
  exportWizardPreviewMeta: document.getElementById('exportWizardPreviewMeta'),
  exportTransparentBgToleranceValue: document.getElementById('exportTransparentBgToleranceValue'),
  exportTransparentBgProtectionValue: document.getElementById('exportTransparentBgProtectionValue'),
  exportFormatSelect: document.getElementById('exportFormatSelect'),
  exportPresetSelect: document.getElementById('exportPresetSelect'),
  exportCompressionLevelSelect: document.getElementById('exportCompressionLevelSelect'),
  exportCompressionRange: document.getElementById('exportCompressionRange'),
  exportCompressionValue: document.getElementById('exportCompressionValue'),
  exportCompressionRecommendedDot: document.getElementById('exportCompressionRecommendedDot'),
  exportWizardOriginalDownloadBtn: document.getElementById('exportWizardOriginalDownloadBtn'),
  exportWizardOriginalDownloadLabel: document.getElementById('exportWizardOriginalDownloadLabel'),
  exportWizardOriginalDownloadHint: document.getElementById('exportWizardOriginalDownloadHint'),
  exportWizardCloseBtn: document.getElementById('exportWizardCloseBtn'),
  exportWizardDownloadBtn: document.getElementById('exportWizardDownloadBtn'),
  exportWizardDownloadLabel: document.getElementById('exportWizardDownloadLabel'),
  exportWizardDownloadHint: document.getElementById('exportWizardDownloadHint'),
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
  pendingManualUploadFiles: [],
  pendingManualUploadTarget: null,
  exportWizardItem: null,
  exportWizardFormat: EXPORT_DEFAULT_FORMAT,
  exportWizardPreset: EXPORT_PRESET_ORIGINAL,
  exportWizardCompressionLevel: EXPORT_DEFAULT_COMPRESSION_LEVEL,
  exportWizardCompression: EXPORT_DEFAULT_COMPRESSION,
  exportWizardRecommendedCompression: null,
  exportWizardPassthroughCompression: null,
  exportWizardFormatTouched: false,
  exportWizardFormatRecommendationToken: 0,
  exportWizardFormatAnalysisToken: 0,
  exportWizardFormatAnalysesKey: '',
  exportWizardFormatAnalyses: {},
  exportWizardInitialFormatAnalysisToken: 0,
  exportWizardInitialFormatAnalysesKey: '',
  exportWizardInitialFormatAnalyses: {},
  exportWizardRecommendationToken: 0,
  exportWizardCompressionTouched: false,
  exportWizardCropEnabled: false,
  exportWizardCropRect: { x: 0, y: 0, width: 1, height: 1 },
  exportWizardCropInteraction: null,
  exportWizardTransparentBackground: false,
  exportWizardBackgroundTolerance: EXPORT_BACKGROUND_TOLERANCE_DEFAULT,
  exportWizardBackgroundProtection: EXPORT_BACKGROUND_PROTECTION_DEFAULT,
  exportWizardTransparentAdjusting: false,
  exportWizardTransparentAdjustingControl: '',
  exportWizardTransparentPointerId: null,
  exportWizardEstimatedOutputBytes: 0,
  exportWizardPreviewUrl: '',
  exportWizardCropEditorPreviewUrl: '',
  exportWizardPreviewToken: 0,
  exportWizardPreviewLoadingTimeout: null,
  exportWizardBusy: false,
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

  if (elements.saveGalleryAsFabBtn instanceof HTMLButtonElement) {
    elements.saveGalleryAsFabBtn.addEventListener('click', () => {
      openSaveAsModal();
    });
  }

  if (elements.saveAsCancelBtn instanceof HTMLButtonElement) {
    elements.saveAsCancelBtn.addEventListener('click', () => {
      closeSaveAsModal();
    });
  }

  if (elements.saveAsConfirmBtn instanceof HTMLButtonElement) {
    elements.saveAsConfirmBtn.addEventListener('click', () => {
      void confirmSaveAsModal();
    });
  }

  if (elements.saveAsNameInput instanceof HTMLInputElement) {
    elements.saveAsNameInput.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        closeSaveAsModal();
        return;
      }
      if (event.key !== 'Enter') {
        return;
      }
      event.preventDefault();
      void confirmSaveAsModal();
    });
  }

  if (elements.saveAsModal instanceof HTMLElement) {
    elements.saveAsModal.addEventListener('click', (event) => {
      if (event.target === elements.saveAsModal) {
        closeSaveAsModal();
      }
    });
  }

  if (elements.manualUploadCancelBtn instanceof HTMLButtonElement) {
    elements.manualUploadCancelBtn.addEventListener('click', () => {
      closeManualUploadModal();
    });
  }

  if (elements.manualUploadConfirmBtn instanceof HTMLButtonElement) {
    elements.manualUploadConfirmBtn.addEventListener('click', () => {
      confirmManualUploadModal();
    });
  }

  if (elements.manualUploadNameInput instanceof HTMLInputElement) {
    elements.manualUploadNameInput.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        closeManualUploadModal();
        return;
      }
      if (event.key !== 'Enter') {
        return;
      }
      event.preventDefault();
      confirmManualUploadModal();
    });
  }

  if (elements.manualUploadGallerySelect instanceof HTMLSelectElement) {
    elements.manualUploadGallerySelect.addEventListener('change', () => {
      const selectedId = String(elements.manualUploadGallerySelect.value || '').trim();
      if (selectedId && elements.manualUploadNameInput instanceof HTMLInputElement) {
        elements.manualUploadNameInput.value = '';
      }
    });
  }

  if (elements.manualUploadModal instanceof HTMLElement) {
    elements.manualUploadModal.addEventListener('click', (event) => {
      if (event.target === elements.manualUploadModal) {
        closeManualUploadModal();
      }
    });
  }

  if (elements.exportWizardCloseBtn instanceof HTMLButtonElement) {
    elements.exportWizardCloseBtn.addEventListener('click', () => {
      closeExportWizardModal();
    });
  }

  if (elements.exportWizardNameInput instanceof HTMLInputElement) {
    elements.exportWizardNameInput.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        closeExportWizardModal();
        return;
      }
      if (event.key !== 'Enter') {
        return;
      }
      event.preventDefault();
      void exportWizardDownload();
    });
  }

  if (elements.exportWizardDownloadBtn instanceof HTMLButtonElement) {
    elements.exportWizardDownloadBtn.addEventListener('click', () => {
      void exportWizardDownload();
    });
  }

  if (elements.exportWizardOriginalDownloadBtn instanceof HTMLButtonElement) {
    elements.exportWizardOriginalDownloadBtn.addEventListener('click', () => {
      void exportWizardDownloadOriginalMatchedSize();
    });
  }

  if (elements.exportCropToggleBtn instanceof HTMLButtonElement) {
    elements.exportCropToggleBtn.addEventListener('click', () => {
      const shouldEnable = state.exportWizardCropEnabled !== true;
      state.exportWizardCropEnabled = shouldEnable;
      if (shouldEnable && isExportCropRectFull(state.exportWizardCropRect)) {
        state.exportWizardCropRect = createDefaultExportCropRect(state.exportWizardItem, state.exportWizardPreset);
      }
      updateExportCropUi();
      void refreshExportWizardPreview();
      void refreshExportFormatAndCompressionRecommendations();
    });
  }

  if (elements.exportTransparentBgPopover instanceof HTMLElement) {
    elements.exportTransparentBgPopover.addEventListener('mouseenter', () => {
      logExportTransparentDebug('popover mouseenter');
    });
    elements.exportTransparentBgPopover.addEventListener('mouseleave', () => {
      logExportTransparentDebug('popover mouseleave');
    });
    elements.exportTransparentBgPopover.addEventListener('pointerdown', (event) => {
      logExportTransparentDebug('popover pointerdown', {
        targetTag: event.target instanceof HTMLElement ? event.target.tagName : null
      });
      event.stopPropagation();
    });
    elements.exportTransparentBgPopover.addEventListener('click', (event) => {
      logExportTransparentDebug('popover click', {
        targetTag: event.target instanceof HTMLElement ? event.target.tagName : null
      });
      event.stopPropagation();
    });
  }
  if (elements.exportTransparentBgBtn instanceof HTMLButtonElement) {
    elements.exportTransparentBgBtn.addEventListener('click', () => {
      const shouldEnable = state.exportWizardTransparentBackground !== true;
      if (shouldEnable) {
        setExportTransparentTolerance(EXPORT_BACKGROUND_TOLERANCE_ACTIVE_DEFAULT);
        setExportTransparentProtection(EXPORT_BACKGROUND_PROTECTION_ACTIVE_DEFAULT);
      } else {
        setExportTransparentTolerance(0);
        setExportTransparentProtection(EXPORT_BACKGROUND_PROTECTION_DEFAULT);
      }
      logExportTransparentDebug('button click', { enabled: shouldEnable });
      syncExportFormatControlState();
      updateExportTransparentBackgroundUi();
      void refreshExportWizardPreview();
      void refreshExportFormatAndCompressionRecommendations();
    });
  }

  if (elements.exportCropStage instanceof HTMLElement) {
    elements.exportCropStage.addEventListener('pointerdown', handleExportCropPointerDown);
  }

  if (elements.exportCropEditorImg instanceof HTMLImageElement) {
    elements.exportCropEditorImg.addEventListener('load', () => {
      updateExportCropUi();
    });
  }

  window.addEventListener('pointermove', handleExportCropPointerMove);
  window.addEventListener('pointerup', handleExportCropPointerUp);
  window.addEventListener('pointercancel', handleExportCropPointerUp);
  window.addEventListener('pointermove', (event) => {
    if (state.exportWizardTransparentAdjusting !== true || event.pointerId !== state.exportWizardTransparentPointerId) {
      return;
    }
    if (state.exportWizardTransparentAdjustingControl === 'protection') {
      updateExportTransparentProtectionFromClientY(event.clientY);
    } else {
      updateExportTransparentToleranceFromClientY(event.clientY);
    }
    logExportTransparentDebug('slider pointermove', {
      clientY: Math.round(event.clientY)
    });
    syncExportFormatControlState();
    updateExportTransparentBackgroundUi();
    void refreshExportWizardPreview();
  });
  window.addEventListener('pointerup', () => {
    if (state.exportWizardTransparentAdjusting !== true) {
      return;
    }
    state.exportWizardTransparentAdjusting = false;
    state.exportWizardTransparentAdjustingControl = '';
    state.exportWizardTransparentPointerId = null;
    updateExportTransparentBackgroundUi();
    logExportTransparentDebug('window pointerup');
    void refreshExportFormatAndCompressionRecommendations();
  });
  window.addEventListener('pointercancel', () => {
    if (state.exportWizardTransparentAdjusting !== true) {
      return;
    }
    state.exportWizardTransparentAdjusting = false;
    state.exportWizardTransparentAdjustingControl = '';
    state.exportWizardTransparentPointerId = null;
    updateExportTransparentBackgroundUi();
    logExportTransparentDebug('window pointercancel');
    void refreshExportFormatAndCompressionRecommendations();
  });
  window.addEventListener('resize', () => {
    if (state.exportWizardItem) {
      updateExportCropUi();
    }
  });

  if (elements.exportTransparentBgSlider instanceof HTMLElement) {
    elements.exportTransparentBgSlider.addEventListener('pointerdown', (event) => {
      state.exportWizardTransparentAdjusting = true;
      state.exportWizardTransparentAdjustingControl = 'tolerance';
      state.exportWizardTransparentPointerId = event.pointerId;
      updateExportTransparentToleranceFromClientY(event.clientY);
      elements.exportTransparentBgSlider.setPointerCapture?.(event.pointerId);
      logExportTransparentDebug('slider pointerdown', {
        clientY: Math.round(event.clientY)
      });
      syncExportFormatControlState();
      updateExportTransparentBackgroundUi();
      void refreshExportWizardPreview();
      event.preventDefault();
      event.stopPropagation();
    });
    elements.exportTransparentBgSlider.addEventListener('keydown', (event) => {
      let delta = 0;
      if (event.key === 'ArrowUp' || event.key === 'ArrowRight') {
        delta = 4;
      } else if (event.key === 'ArrowDown' || event.key === 'ArrowLeft') {
        delta = -4;
      } else if (event.key === 'Home') {
        setExportTransparentTolerance(0);
      } else if (event.key === 'End') {
        setExportTransparentTolerance(EXPORT_BACKGROUND_TOLERANCE_MAX);
      } else {
        return;
      }
      if (delta !== 0) {
        setExportTransparentTolerance(state.exportWizardBackgroundTolerance + delta);
      }
      logExportTransparentDebug('slider keydown', { key: event.key });
      syncExportFormatControlState();
      updateExportTransparentBackgroundUi();
      void refreshExportWizardPreview();
      void refreshExportFormatAndCompressionRecommendations();
      event.preventDefault();
    });
  }

  if (elements.exportTransparentBgProtectionSlider instanceof HTMLElement) {
    elements.exportTransparentBgProtectionSlider.addEventListener('pointerdown', (event) => {
      state.exportWizardTransparentAdjusting = true;
      state.exportWizardTransparentAdjustingControl = 'protection';
      state.exportWizardTransparentPointerId = event.pointerId;
      updateExportTransparentProtectionFromClientY(event.clientY);
      elements.exportTransparentBgProtectionSlider.setPointerCapture?.(event.pointerId);
      logExportTransparentDebug('protection slider pointerdown', {
        clientY: Math.round(event.clientY)
      });
      syncExportFormatControlState();
      updateExportTransparentBackgroundUi();
      void refreshExportWizardPreview();
      event.preventDefault();
      event.stopPropagation();
    });
    elements.exportTransparentBgProtectionSlider.addEventListener('keydown', (event) => {
      let delta = 0;
      if (event.key === 'ArrowUp' || event.key === 'ArrowRight') {
        delta = 4;
      } else if (event.key === 'ArrowDown' || event.key === 'ArrowLeft') {
        delta = -4;
      } else if (event.key === 'Home') {
        setExportTransparentProtection(0);
      } else if (event.key === 'End') {
        setExportTransparentProtection(EXPORT_BACKGROUND_PROTECTION_MAX);
      } else {
        return;
      }
      if (delta !== 0) {
        setExportTransparentProtection(state.exportWizardBackgroundProtection + delta);
      }
      logExportTransparentDebug('protection slider keydown', { key: event.key });
      syncExportFormatControlState();
      updateExportTransparentBackgroundUi();
      void refreshExportWizardPreview();
      void refreshExportFormatAndCompressionRecommendations();
      event.preventDefault();
    });
  }

  if (elements.exportFormatSelect instanceof HTMLSelectElement) {
    elements.exportFormatSelect.addEventListener('change', () => {
      state.exportWizardFormat = normalizeExportFormat(elements.exportFormatSelect.value);
      state.exportWizardFormatTouched = true;
      state.exportWizardCompressionTouched = false;
      updateExportWizardActionUi();
      void refreshExportWizardPreview();
      void applyRecommendedCompressionForCurrentExport({ force: true, apply: true });
    });
  }

  if (elements.exportPresetSelect instanceof HTMLSelectElement) {
    elements.exportPresetSelect.addEventListener('change', () => {
      state.exportWizardPreset = normalizeExportPreset(elements.exportPresetSelect.value);
      if (isExportCropAspectLocked(state.exportWizardPreset)) {
        state.exportWizardCropRect = createDefaultExportCropRect(state.exportWizardItem, state.exportWizardPreset);
      }
      syncExportFormatControlState();
      updateExportCropUi();
      state.exportWizardCompressionTouched = false;
      void refreshExportWizardPreview();
      void refreshExportFormatAndCompressionRecommendations();
    });
  }

  if (elements.exportCompressionLevelSelect instanceof HTMLSelectElement) {
    elements.exportCompressionLevelSelect.addEventListener('change', () => {
      state.exportWizardCompressionLevel = normalizeExportCompressionLevel(elements.exportCompressionLevelSelect.value);
      state.exportWizardCompressionTouched = false;
      updateExportCompressionLevelUi();
      updateExportOptimalSummaryUi();
      void refreshExportFormatAndCompressionRecommendations();
    });
  }

  if (elements.exportCompressionRange instanceof HTMLInputElement) {
    elements.exportCompressionRange.addEventListener('input', () => {
      state.exportWizardCompression = normalizeExportCompression(elements.exportCompressionRange.value);
      state.exportWizardCompressionTouched = true;
      updateExportCompressionUi();
      void refreshExportWizardPreview();
    });
  }

  if (elements.exportWizardModal instanceof HTMLElement) {
    elements.exportWizardModal.addEventListener('click', (event) => {
      if (event.target === elements.exportWizardModal) {
        closeExportWizardModal();
      }
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
    state.pendingManualUploadFiles = [];
    state.pendingManualUploadTarget = null;
    const activeGalleryId =
      state.galleryView === VIEW_LIBRARY && state.galleryListMode === 'detail'
        ? String(state.activeGalleryId || '').trim()
        : '';
    if (activeGalleryId) {
      state.pendingManualUploadTarget = {
        mode: 'existing',
        galleryId: activeGalleryId
      };
    }
    elements.logoUploadInput.value = '';
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
      state.pendingManualUploadFiles = [];
      state.pendingManualUploadTarget = null;
      return;
    }
    const directTarget = state.pendingManualUploadTarget;
    state.pendingManualUploadTarget = null;
    if (directTarget) {
      state.pendingManualUploadFiles = [];
      void addManualImages(files, directTarget);
      return;
    }
    state.pendingManualUploadFiles = files;
    openManualUploadModal();
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

    const exportButton = target.closest('button[data-image-export-id]');
    if (exportButton && !exportButton.disabled) {
      event.preventDefault();
      event.stopPropagation();
      const imageId = String(exportButton.dataset.imageExportId || '').trim();
      const manual = exportButton.dataset.imageExportManual === '1';
      const galleryId = String(exportButton.dataset.imageExportGalleryId || '').trim();
      if (imageId) {
        openExportWizardFromImageId(imageId, { manual, galleryId });
      }
      return;
    }

    const downloadGalleryButton = target.closest('button[data-gallery-download-id]');
    if (downloadGalleryButton && !downloadGalleryButton.disabled) {
      event.preventDefault();
      event.stopPropagation();
      const galleryId = String(downloadGalleryButton.dataset.galleryDownloadId || '').trim();
      if (!galleryId) {
        return;
      }
      void downloadGalleryAsZip(galleryId);
      return;
    }

    const deleteGalleryButton = target.closest('button[data-gallery-delete-id]');
    if (deleteGalleryButton && !deleteGalleryButton.disabled) {
      event.preventDefault();
      event.stopPropagation();
      const galleryId = String(deleteGalleryButton.dataset.galleryDeleteId || '').trim();
      if (!galleryId) {
        return;
      }
      void removeGalleryById(galleryId);
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

    const deleteGalleryItemButton = target.closest('button[data-gallery-item-delete-id]');
    if (deleteGalleryItemButton && !deleteGalleryItemButton.disabled) {
      event.preventDefault();
      event.stopPropagation();
      const imageId = String(deleteGalleryItemButton.dataset.galleryItemDeleteId || '').trim();
      const galleryId = String(
        deleteGalleryItemButton.dataset.galleryItemGalleryId || state.activeGalleryId || ''
      ).trim();
      if (!imageId || !galleryId) {
        return;
      }
      void removeStoredGalleryImageById(galleryId, imageId);
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
  if (safe !== 'url') {
    closeSaveAsModal();
    closeManualUploadModal();
  }

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
  closeExportWizardModal();
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

function openLibraryPageWhenWebGalleryUnavailable() {
  if (state.activePage !== 'url') {
    return;
  }
  state.galleryView = VIEW_LIBRARY;
  if (!String(state.activeGalleryId || '').trim()) {
    state.galleryListMode = 'list';
  }
  setActivePage('library');
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
      elements.logoLibraryTitleText.textContent = getGalleryDisplayName(gallery);
    }
  } else {
    elements.logoLibraryTitleText.textContent = 'Galeries';
  }
  elements.logoLibraryTitleIcon.innerHTML = '<i class="bi bi-images" aria-hidden="true"></i>';
}

function updateSaveGalleryFabVisibility() {
  if (
    !(elements.saveGalleryFabBtn instanceof HTMLButtonElement) ||
    !(elements.saveGalleryAsFabBtn instanceof HTMLButtonElement) ||
    !(elements.saveGalleryFabGroup instanceof HTMLElement)
  ) {
    return;
  }
  const visible = state.activePage === 'url' && state.galleryView === VIEW_WEB;
  const enabled = visible && !state.webLoading && state.webItems.length > 0;
  const domainContext = getActiveTabDomainContext();
  elements.saveGalleryFabGroup.classList.toggle('hidden', !visible);
  elements.saveGalleryFabBtn.disabled = !enabled || !domainContext.isValid;
  elements.saveGalleryAsFabBtn.disabled = !enabled;
  if (domainContext.isValid) {
    elements.saveGalleryFabBtn.removeAttribute('title');
  } else {
    elements.saveGalleryFabBtn.title = 'Domini no valid';
  }
  if (!visible) {
    closeSaveAsModal();
  }
}

function openSaveAsModal() {
  if (!(elements.saveAsModal instanceof HTMLElement) || !(elements.saveAsNameInput instanceof HTMLInputElement)) {
    return;
  }
  const canOpen = state.activePage === 'url' && state.galleryView === VIEW_WEB && !state.webLoading && state.webItems.length > 0;
  if (!canOpen) {
    return;
  }
  closeManualUploadModal();
  closeExportWizardModal();
  const defaultName = getDefaultSaveAsGalleryName();
  elements.saveAsNameInput.value = defaultName;
  elements.saveAsModal.classList.remove('hidden');
  elements.saveAsModal.setAttribute('aria-hidden', 'false');
  window.requestAnimationFrame(() => {
    elements.saveAsNameInput.focus();
    if (defaultName) {
      elements.saveAsNameInput.select();
    }
  });
}

function closeSaveAsModal() {
  if (!(elements.saveAsModal instanceof HTMLElement)) {
    return;
  }
  elements.saveAsModal.classList.add('hidden');
  elements.saveAsModal.setAttribute('aria-hidden', 'true');
}

async function confirmSaveAsModal() {
  if (!(elements.saveAsNameInput instanceof HTMLInputElement)) {
    return;
  }
  const name = String(elements.saveAsNameInput.value || '').trim();
  if (!name) {
    showFeedbackToast('Escriu un nom per guardar', 'error');
    elements.saveAsNameInput.focus();
    return;
  }
  closeSaveAsModal();
  await saveSelectedWebImagesToDomainGallery({ customName: name });
}

function openManualUploadModal() {
  if (
    !(elements.manualUploadModal instanceof HTMLElement) ||
    !(elements.manualUploadGallerySelect instanceof HTMLSelectElement) ||
    !(elements.manualUploadNameInput instanceof HTMLInputElement)
  ) {
    return;
  }
  const pendingCount = Array.isArray(state.pendingManualUploadFiles) ? state.pendingManualUploadFiles.length : 0;
  const canOpen =
    state.galleryView === VIEW_LIBRARY &&
    (state.activePage === 'library' || state.activePage === 'url') &&
    pendingCount > 0;
  if (!canOpen) {
    return;
  }

  closeSaveAsModal();
  closeExportWizardModal();
  populateManualUploadGallerySelect();
  elements.manualUploadGallerySelect.value = '';
  elements.manualUploadNameInput.value = '';
  elements.manualUploadModal.classList.remove('hidden');
  elements.manualUploadModal.setAttribute('aria-hidden', 'false');
  window.requestAnimationFrame(() => {
    elements.manualUploadGallerySelect.focus();
  });
}

function closeManualUploadModal(options = {}) {
  if (!(elements.manualUploadModal instanceof HTMLElement)) {
    return;
  }
  const clearPending = options.clearPending !== false;
  elements.manualUploadModal.classList.add('hidden');
  elements.manualUploadModal.setAttribute('aria-hidden', 'true');
  if (clearPending) {
    state.pendingManualUploadFiles = [];
    state.pendingManualUploadTarget = null;
  }
}

function confirmManualUploadModal() {
  const target = buildManualUploadTargetFromModal();
  if (!target) {
    return;
  }
  const files = Array.isArray(state.pendingManualUploadFiles) ? state.pendingManualUploadFiles : [];
  if (files.length === 0) {
    closeManualUploadModal();
    showFeedbackToast('Selecciona imatges primer', 'error');
    return;
  }
  closeManualUploadModal({ clearPending: false });
  state.pendingManualUploadFiles = [];
  void addManualImages(files, target);
}

function populateManualUploadGallerySelect() {
  if (!(elements.manualUploadGallerySelect instanceof HTMLSelectElement)) {
    return;
  }

  const select = elements.manualUploadGallerySelect;
  select.replaceChildren();
  const createOption = document.createElement('option');
  createOption.value = '';
  createOption.textContent = 'Crear nova galeria';
  select.appendChild(createOption);

  const collections = getGalleryCollections();
  for (const collection of collections) {
    const galleryId = String(collection?.id || '').trim();
    if (!galleryId) {
      continue;
    }
    const option = document.createElement('option');
    option.value = galleryId;
    option.textContent = String(collection.title || 'Galeria');
    select.appendChild(option);
  }
}

function buildManualUploadTargetFromModal() {
  if (
    !(elements.manualUploadGallerySelect instanceof HTMLSelectElement) ||
    !(elements.manualUploadNameInput instanceof HTMLInputElement)
  ) {
    return null;
  }

  const selectedId = String(elements.manualUploadGallerySelect.value || '').trim();
  if (selectedId) {
    return {
      mode: 'existing',
      galleryId: selectedId
    };
  }

  const customName = String(elements.manualUploadNameInput.value || '').trim();
  if (!customName) {
    showFeedbackToast('Escriu un nom o selecciona una galeria', 'error');
    elements.manualUploadNameInput.focus();
    return null;
  }
  return {
    mode: 'new',
    name: customName
  };
}

function normalizeExportFormat(value) {
  const safe = String(value || '').trim().toLowerCase();
  if (safe === 'image/png') {
    return 'image/png';
  }
  if (safe === 'image/jpeg' || safe === 'image/jpg') {
    return 'image/jpeg';
  }
  return 'image/webp';
}

function normalizeExportPreset(value) {
  const safe = String(value || '').trim().toLowerCase();
  return safe === EXPORT_PRESET_FAVICON ? EXPORT_PRESET_FAVICON : EXPORT_PRESET_ORIGINAL;
}

function normalizeExportCompression(value) {
  const raw = Math.round(Number(value));
  if (!Number.isFinite(raw)) {
    return EXPORT_DEFAULT_COMPRESSION;
  }
  return Math.max(0, Math.min(100, raw));
}

function extractComparableExportFormatFromDataUrl(dataUrl) {
  const safe = String(dataUrl || '').trim();
  const match = safe.match(/^data:(image\/[a-z0-9.+-]+)(?:;|,)/iu);
  if (!match) {
    return '';
  }
  const mime = String(match[1] || '').trim().toLowerCase();
  if (mime === 'image/png') {
    return 'image/png';
  }
  if (mime === 'image/jpeg' || mime === 'image/jpg') {
    return 'image/jpeg';
  }
  if (mime === 'image/webp') {
    return 'image/webp';
  }
  return '';
}

function extractMimeTypeFromDataUrl(dataUrl) {
  const safe = String(dataUrl || '').trim();
  const match = safe.match(/^data:([^;,]+)(?:;|,)/iu);
  return match ? String(match[1] || '').trim().toLowerCase() : '';
}

function isOriginalFormatPassthroughExport(
  item,
  format,
  preset,
  editOptions = {},
  compression = null,
  passthroughCompression = null
) {
  const safePreset = normalizeExportPreset(preset);
  const safeFormat = normalizeExportFormat(format);
  const originalFormat = extractComparableExportFormatFromDataUrl(item?.dataUrl);
  const safeCompression = compression === null ? 0 : normalizeExportCompression(compression);
  const safePassthroughCompression = Number.isFinite(Number(passthroughCompression))
    ? normalizeExportCompression(passthroughCompression)
    : null;
  return (
    safePreset === EXPORT_PRESET_ORIGINAL &&
    editOptions?.cropEnabled !== true &&
    editOptions?.transparentBackground !== true &&
    originalFormat === safeFormat &&
    (safeCompression <= 0 || (safePassthroughCompression !== null && safeCompression === safePassthroughCompression))
  );
}

function resolveLossyQualityFromCompression(compression) {
  const safeCompression = normalizeExportCompression(compression);
  const ratio = safeCompression / 100;
  const easedRatio = Math.pow(ratio, 1.35);
  const quality = EXPORT_MAX_QUALITY - easedRatio * (EXPORT_MAX_QUALITY - EXPORT_MIN_QUALITY);
  return Math.min(1, Math.max(0.05, Number(quality.toFixed(4))));
}

function resolveZipCompressionLevel(compression) {
  const safeCompression = normalizeExportCompression(compression);
  return Math.max(0, Math.min(ZIP_MAX_COMPRESSION_LEVEL, Math.round((safeCompression / 100) * ZIP_MAX_COMPRESSION_LEVEL)));
}

function resolvePngColorCountFromCompression(compression) {
  const safeCompression = normalizeExportCompression(compression);
  if (safeCompression <= 0) {
    return 0;
  }
  const ratio = safeCompression / 100;
  const easedRatio = Math.pow(ratio, 2.2);
  const colors = Math.round(256 - easedRatio * 254);
  return Math.max(2, Math.min(256, colors));
}

function createFullExportCropRect() {
  return { x: 0, y: 0, width: 1, height: 1 };
}

function isExportCropAspectLocked(preset = state.exportWizardPreset) {
  return normalizeExportPreset(preset) === EXPORT_PRESET_FAVICON;
}

function isExportCropRectFull(rect = {}) {
  const safeRect = clampExportCropRect(rect);
  return safeRect.x <= 0.0001 && safeRect.y <= 0.0001 && safeRect.width >= 0.9999 && safeRect.height >= 0.9999;
}

function normalizeExportCropRatio(value) {
  const raw = Number(value);
  if (!Number.isFinite(raw)) {
    return 0;
  }
  return Math.max(0, Math.min(1, raw));
}

function clampExportCropRect(rect = {}) {
  const minSize = EXPORT_CROP_MIN_RATIO;
  const width = Math.max(minSize, Math.min(1, normalizeExportCropRatio(rect.width)));
  const height = Math.max(minSize, Math.min(1, normalizeExportCropRatio(rect.height)));
  const maxX = Math.max(0, 1 - width);
  const maxY = Math.max(0, 1 - height);
  return {
    x: Math.max(0, Math.min(maxX, normalizeExportCropRatio(rect.x))),
    y: Math.max(0, Math.min(maxY, normalizeExportCropRatio(rect.y))),
    width,
    height
  };
}

function createSquareExportCropRect(item) {
  const width = Math.max(1, Number(item?.width) || 1);
  const height = Math.max(1, Number(item?.height) || 1);
  if (width === height) {
    return createFullExportCropRect();
  }
  if (width > height) {
    const ratio = height / width;
    return clampExportCropRect({
      x: (1 - ratio) / 2,
      y: 0,
      width: ratio,
      height: 1
    });
  }
  const ratio = width / height;
  return clampExportCropRect({
    x: 0,
    y: (1 - ratio) / 2,
    width: 1,
    height: ratio
  });
}

function createDefaultExportCropRect(item, preset = state.exportWizardPreset) {
  if (!isExportCropAspectLocked(preset)) {
    return createFullExportCropRect();
  }
  return createSquareExportCropRect(item);
}

function normalizeExportBackgroundTolerance(value) {
  const raw = Math.round(Number(value));
  if (!Number.isFinite(raw)) {
    return EXPORT_BACKGROUND_TOLERANCE_DEFAULT;
  }
  return Math.max(0, Math.min(EXPORT_BACKGROUND_TOLERANCE_MAX, raw));
}

function normalizeExportBackgroundProtection(value) {
  const raw = Math.round(Number(value));
  if (!Number.isFinite(raw)) {
    return EXPORT_BACKGROUND_PROTECTION_DEFAULT;
  }
  return Math.max(0, Math.min(EXPORT_BACKGROUND_PROTECTION_MAX, raw));
}

function logExportTransparentDebug(label, extra = {}) {
  const button = elements.exportTransparentBgBtn;
  const popover = elements.exportTransparentBgPopover;
  const slider = elements.exportTransparentBgSlider;
  const buttonRect = button instanceof HTMLElement ? button.getBoundingClientRect() : null;
  const popoverRect = popover instanceof HTMLElement ? popover.getBoundingClientRect() : null;
  const sliderRect = slider instanceof HTMLElement ? slider.getBoundingClientRect() : null;
  console.log(`[export-transparent] ${label}`, {
    tolerance: state.exportWizardBackgroundTolerance,
    protection: state.exportWizardBackgroundProtection,
    enabled: state.exportWizardTransparentBackground,
    adjusting: state.exportWizardTransparentAdjusting,
    adjustingControl: state.exportWizardTransparentAdjustingControl,
    buttonPressed: button instanceof HTMLButtonElement ? button.getAttribute('aria-pressed') : null,
    popoverHiddenClass: popover instanceof HTMLElement ? popover.classList.contains('hidden') : null,
    popoverAriaHidden: popover instanceof HTMLElement ? popover.getAttribute('aria-hidden') : null,
    sliderValue: state.exportWizardBackgroundTolerance,
    buttonRect: buttonRect
      ? { x: Math.round(buttonRect.x), y: Math.round(buttonRect.y), w: Math.round(buttonRect.width), h: Math.round(buttonRect.height) }
      : null,
    popoverRect: popoverRect
      ? { x: Math.round(popoverRect.x), y: Math.round(popoverRect.y), w: Math.round(popoverRect.width), h: Math.round(popoverRect.height) }
      : null,
    sliderRect: sliderRect
      ? { x: Math.round(sliderRect.x), y: Math.round(sliderRect.y), w: Math.round(sliderRect.width), h: Math.round(sliderRect.height) }
      : null,
    ...extra
  });
}

function setExportTransparentTolerance(nextTolerance) {
  const tolerance = normalizeExportBackgroundTolerance(nextTolerance);
  state.exportWizardBackgroundTolerance = tolerance;
  state.exportWizardTransparentBackground = tolerance > 0;
}

function setExportTransparentProtection(nextProtection) {
  state.exportWizardBackgroundProtection = normalizeExportBackgroundProtection(nextProtection);
}

function resolveExportTransparentSliderRatio(slider, clientY) {
  if (!(slider instanceof HTMLElement)) {
    return 0;
  }
  const rect = slider.getBoundingClientRect();
  const offset = Math.max(0, Math.min(rect.height, rect.bottom - clientY));
  return rect.height > 0 ? offset / rect.height : 0;
}

function updateExportTransparentToleranceFromClientY(clientY) {
  if (!(elements.exportTransparentBgSlider instanceof HTMLElement)) {
    return;
  }
  const ratio = resolveExportTransparentSliderRatio(elements.exportTransparentBgSlider, clientY);
  const nextTolerance = Math.round(ratio * EXPORT_BACKGROUND_TOLERANCE_MAX);
  setExportTransparentTolerance(nextTolerance);
}

function updateExportTransparentProtectionFromClientY(clientY) {
  if (!(elements.exportTransparentBgProtectionSlider instanceof HTMLElement)) {
    return;
  }
  const ratio = resolveExportTransparentSliderRatio(elements.exportTransparentBgProtectionSlider, clientY);
  const nextProtection = Math.round(ratio * EXPORT_BACKGROUND_PROTECTION_MAX);
  setExportTransparentProtection(nextProtection);
}

function getExportCropStageMetrics() {
  if (!(elements.exportCropStage instanceof HTMLElement) || !state.exportWizardItem) {
    return null;
  }
  const stageWidth = Math.max(0, Number(elements.exportCropStage.clientWidth) || 0);
  const stageHeight = Math.max(0, Number(elements.exportCropStage.clientHeight) || 0);
  if (stageWidth <= 0 || stageHeight <= 0) {
    return null;
  }
  const sourceWidth = Math.max(
    1,
    Number(state.exportWizardItem.width) || Number(elements.exportCropEditorImg?.naturalWidth) || 1
  );
  const sourceHeight = Math.max(
    1,
    Number(state.exportWizardItem.height) || Number(elements.exportCropEditorImg?.naturalHeight) || 1
  );
  const scale = Math.min(stageWidth / sourceWidth, stageHeight / sourceHeight);
  const width = Math.max(1, sourceWidth * scale);
  const height = Math.max(1, sourceHeight * scale);
  return {
    stageWidth,
    stageHeight,
    sourceWidth,
    sourceHeight,
    left: (stageWidth - width) / 2,
    top: (stageHeight - height) / 2,
    width,
    height
  };
}

function convertExportCropRectToPixels(rect, metrics) {
  const safeRect = clampExportCropRect(rect);
  return {
    x: safeRect.x * metrics.width,
    y: safeRect.y * metrics.height,
    width: safeRect.width * metrics.width,
    height: safeRect.height * metrics.height
  };
}

function convertExportCropPixelsToRect(rect, metrics) {
  return clampExportCropRect({
    x: rect.x / metrics.width,
    y: rect.y / metrics.height,
    width: rect.width / metrics.width,
    height: rect.height / metrics.height
  });
}

function getExportCropPointerPosition(event, metrics) {
  if (!(elements.exportCropStage instanceof HTMLElement)) {
    return { x: 0, y: 0 };
  }
  const stageRect = elements.exportCropStage.getBoundingClientRect();
  const rawX = event.clientX - stageRect.left - metrics.left;
  const rawY = event.clientY - stageRect.top - metrics.top;
  return {
    x: Math.max(0, Math.min(metrics.width, rawX)),
    y: Math.max(0, Math.min(metrics.height, rawY))
  };
}

function resolveExportCropPixelMinWidth(metrics, aspectRatio) {
  const minWidth = metrics.width * EXPORT_CROP_MIN_RATIO;
  const minHeight = metrics.height * EXPORT_CROP_MIN_RATIO;
  return Math.max(1, minWidth, minHeight * aspectRatio);
}

function resolveExportCropPixelMinSize(metrics) {
  return {
    width: Math.max(1, metrics.width * EXPORT_CROP_MIN_RATIO),
    height: Math.max(1, metrics.height * EXPORT_CROP_MIN_RATIO)
  };
}

function resizeExportCropRectFromCorner(interaction, point) {
  const { handle, metrics, startRect } = interaction;
  const aspectRatio = Math.max(0.0001, startRect.width / Math.max(1, startRect.height));
  const minWidth = resolveExportCropPixelMinWidth(metrics, aspectRatio);
  let anchorX = startRect.x;
  let anchorY = startRect.y;
  let availableWidth = startRect.width;
  let availableHeight = startRect.height;

  if (handle === 'nw') {
    anchorX = startRect.x + startRect.width;
    anchorY = startRect.y + startRect.height;
    availableWidth = anchorX - point.x;
    availableHeight = anchorY - point.y;
  } else if (handle === 'ne') {
    anchorX = startRect.x;
    anchorY = startRect.y + startRect.height;
    availableWidth = point.x - anchorX;
    availableHeight = anchorY - point.y;
  } else if (handle === 'sw') {
    anchorX = startRect.x + startRect.width;
    anchorY = startRect.y;
    availableWidth = anchorX - point.x;
    availableHeight = point.y - anchorY;
  } else {
    anchorX = startRect.x;
    anchorY = startRect.y;
    availableWidth = point.x - anchorX;
    availableHeight = point.y - anchorY;
  }

  const maxWidth = Math.max(1, Math.min(availableWidth, availableHeight * aspectRatio));
  const nextWidth = Math.max(Math.min(maxWidth, Math.max(metrics.width, metrics.height * aspectRatio)), minWidth);
  const nextHeight = nextWidth / aspectRatio;

  if (handle === 'nw') {
    return {
      x: anchorX - nextWidth,
      y: anchorY - nextHeight,
      width: nextWidth,
      height: nextHeight
    };
  }
  if (handle === 'ne') {
    return {
      x: anchorX,
      y: anchorY - nextHeight,
      width: nextWidth,
      height: nextHeight
    };
  }
  if (handle === 'sw') {
    return {
      x: anchorX - nextWidth,
      y: anchorY,
      width: nextWidth,
      height: nextHeight
    };
  }
  return {
    x: anchorX,
    y: anchorY,
    width: nextWidth,
    height: nextHeight
  };
}

function resizeExportCropRectFromEdge(interaction, point) {
  const { handle, metrics, startRect } = interaction;
  const minSize = resolveExportCropPixelMinSize(metrics);
  if (handle === 'n') {
    const nextTop = Math.max(0, Math.min(startRect.y + startRect.height - minSize.height, point.y));
    return {
      x: startRect.x,
      y: nextTop,
      width: startRect.width,
      height: (startRect.y + startRect.height) - nextTop
    };
  }
  if (handle === 's') {
    const nextBottom = Math.max(startRect.y + minSize.height, Math.min(metrics.height, point.y));
    return {
      x: startRect.x,
      y: startRect.y,
      width: startRect.width,
      height: nextBottom - startRect.y
    };
  }
  if (handle === 'w') {
    const nextLeft = Math.max(0, Math.min(startRect.x + startRect.width - minSize.width, point.x));
    return {
      x: nextLeft,
      y: startRect.y,
      width: (startRect.x + startRect.width) - nextLeft,
      height: startRect.height
    };
  }
  const nextRight = Math.max(startRect.x + minSize.width, Math.min(metrics.width, point.x));
  return {
    x: startRect.x,
    y: startRect.y,
    width: nextRight - startRect.x,
    height: startRect.height
  };
}

function updateExportCropUi() {
  const cropEnabled = state.exportWizardCropEnabled === true;
  const rect = clampExportCropRect(state.exportWizardCropRect);
  const aspectLocked = isExportCropAspectLocked(state.exportWizardPreset);
  state.exportWizardCropRect = rect;

  if (elements.exportCropToggleBtn instanceof HTMLButtonElement) {
    elements.exportCropToggleBtn.setAttribute('aria-pressed', cropEnabled ? 'true' : 'false');
    elements.exportCropToggleBtn.setAttribute('aria-label', cropEnabled ? 'Desactivar recorte' : 'Activar recorte');
    elements.exportCropToggleBtn.title = cropEnabled ? 'Desactivar recorte' : 'Recortar';
  }
  if (elements.exportCropEditorImg instanceof HTMLImageElement) {
    const source = String(state.exportWizardCropEditorPreviewUrl || state.exportWizardItem?.dataUrl || '').trim();
    if (source) {
      if (elements.exportCropEditorImg.src !== source) {
        elements.exportCropEditorImg.src = source;
      }
    } else {
      elements.exportCropEditorImg.removeAttribute('src');
    }
  }
  if (elements.exportWizardPreviewImg instanceof HTMLImageElement) {
    elements.exportWizardPreviewImg.classList.toggle('hidden', cropEnabled);
  }
  if (elements.exportCropStage instanceof HTMLElement) {
    elements.exportCropStage.classList.toggle('hidden', !cropEnabled);
    elements.exportCropStage.setAttribute('aria-hidden', cropEnabled ? 'false' : 'true');
  }
  if (elements.exportCropBox instanceof HTMLElement) {
    elements.exportCropBox.classList.toggle('hidden', !cropEnabled);
    elements.exportCropBox.classList.toggle('aspect-locked', aspectLocked);
  }

  const interaction = state.exportWizardCropInteraction;
  if (elements.exportCropBox instanceof HTMLElement) {
    elements.exportCropBox.classList.toggle('dragging', interaction?.mode === 'move');
    elements.exportCropBox.classList.toggle('resizing', interaction?.mode === 'resize');
    if (interaction?.mode === 'resize') {
      const cursor =
        interaction.handle === 'n' || interaction.handle === 's'
          ? 'ns-resize'
          : interaction.handle === 'e' || interaction.handle === 'w'
            ? 'ew-resize'
            : interaction.handle === 'ne' || interaction.handle === 'sw'
              ? 'nesw-resize'
              : 'nwse-resize';
      elements.exportCropBox.style.cursor = cursor;
    } else {
      elements.exportCropBox.style.cursor = interaction?.mode === 'move' ? 'grabbing' : 'grab';
    }
  }

  if (!cropEnabled) {
    return;
  }

  const metrics = getExportCropStageMetrics();
  if (!metrics) {
    return;
  }

  if (elements.exportCropEditorImg instanceof HTMLImageElement) {
    elements.exportCropEditorImg.style.left = `${metrics.left}px`;
    elements.exportCropEditorImg.style.top = `${metrics.top}px`;
    elements.exportCropEditorImg.style.width = `${metrics.width}px`;
    elements.exportCropEditorImg.style.height = `${metrics.height}px`;
  }

  if (elements.exportCropBox instanceof HTMLElement) {
    const rectPixels = convertExportCropRectToPixels(rect, metrics);
    elements.exportCropBox.style.left = `${metrics.left + rectPixels.x}px`;
    elements.exportCropBox.style.top = `${metrics.top + rectPixels.y}px`;
    elements.exportCropBox.style.width = `${rectPixels.width}px`;
    elements.exportCropBox.style.height = `${rectPixels.height}px`;
  }
}

function handleExportCropPointerDown(event) {
  if (state.exportWizardCropEnabled !== true || event.button !== 0) {
    return;
  }
  const target = event.target;
  if (!(target instanceof Element) || !(elements.exportCropBox instanceof HTMLElement)) {
    return;
  }
  if (!target.closest('#exportCropBox')) {
    return;
  }
  const metrics = getExportCropStageMetrics();
  if (!metrics) {
    return;
  }
  const handle = target.closest('[data-export-crop-handle]');
  const mode = handle ? 'resize' : 'move';
  const startPoint = getExportCropPointerPosition(event, metrics);
  state.exportWizardCropInteraction = {
    pointerId: event.pointerId,
    mode,
    handle: handle ? String(handle.getAttribute('data-export-crop-handle') || '').trim() : '',
    metrics,
    startPoint,
    startRect: convertExportCropRectToPixels(state.exportWizardCropRect, metrics)
  };
  if (event.currentTarget instanceof HTMLElement) {
    event.currentTarget.setPointerCapture?.(event.pointerId);
  }
  event.preventDefault();
  updateExportCropUi();
}

function handleExportCropPointerMove(event) {
  const interaction = state.exportWizardCropInteraction;
  if (!interaction || event.pointerId !== interaction.pointerId) {
    return;
  }
  const point = getExportCropPointerPosition(event, interaction.metrics);
  let nextRect = interaction.startRect;

  if (interaction.mode === 'move') {
    const nextX = Math.max(
      0,
      Math.min(
        interaction.metrics.width - interaction.startRect.width,
        interaction.startRect.x + (point.x - interaction.startPoint.x)
      )
    );
    const nextY = Math.max(
      0,
      Math.min(
        interaction.metrics.height - interaction.startRect.height,
        interaction.startRect.y + (point.y - interaction.startPoint.y)
      )
    );
    nextRect = {
      x: nextX,
      y: nextY,
      width: interaction.startRect.width,
      height: interaction.startRect.height
    };
  } else if (interaction.handle) {
    nextRect =
      interaction.handle.length === 2
        ? resizeExportCropRectFromCorner(interaction, point)
        : resizeExportCropRectFromEdge(interaction, point);
  }

  state.exportWizardCropRect = convertExportCropPixelsToRect(nextRect, interaction.metrics);
  updateExportCropUi();
  event.preventDefault();
}

function handleExportCropPointerUp(event) {
  const interaction = state.exportWizardCropInteraction;
  if (!interaction || event.pointerId !== interaction.pointerId) {
    return;
  }
  state.exportWizardCropInteraction = null;
  updateExportCropUi();
  void refreshExportWizardPreview();
  void refreshExportFormatAndCompressionRecommendations();
}

function updateExportTransparentBackgroundUi() {
  const enabled = normalizeExportBackgroundTolerance(state.exportWizardBackgroundTolerance) > 0;
  state.exportWizardTransparentBackground = enabled;
  const tolerance = normalizeExportBackgroundTolerance(state.exportWizardBackgroundTolerance);
  const protection = normalizeExportBackgroundProtection(state.exportWizardBackgroundProtection);
  state.exportWizardBackgroundTolerance = tolerance;
  state.exportWizardBackgroundProtection = protection;
  if (elements.exportTransparentBgBtn instanceof HTMLButtonElement) {
    elements.exportTransparentBgBtn.setAttribute('aria-pressed', enabled ? 'true' : 'false');
    elements.exportTransparentBgBtn.setAttribute(
      'aria-label',
      enabled ? 'Fondo transparente activo' : 'Fondo transparente'
    );
    elements.exportTransparentBgBtn.title = enabled
      ? `Fondo transparente ${tolerance} · Protección ${protection}`
      : 'Fondo transparente';
  }
  const transparentControl = elements.exportTransparentBgBtn?.parentElement;
  if (transparentControl instanceof HTMLElement) {
    transparentControl.classList.toggle('active', enabled);
    transparentControl.classList.toggle('adjusting', state.exportWizardTransparentAdjusting === true);
  }
  if (elements.exportTransparentBgPopover instanceof HTMLElement) {
    elements.exportTransparentBgPopover.setAttribute(
      'aria-hidden',
      enabled ? 'false' : 'true'
    );
  }
  if (elements.exportTransparentBgSlider instanceof HTMLElement) {
    const ratio = EXPORT_BACKGROUND_TOLERANCE_MAX > 0 ? tolerance / EXPORT_BACKGROUND_TOLERANCE_MAX : 0;
    const percent = Math.max(0, Math.min(100, ratio * 100));
    elements.exportTransparentBgSlider.setAttribute('aria-valuenow', String(tolerance));
    elements.exportTransparentBgSlider.setAttribute('aria-valuetext', String(tolerance));
    if (elements.exportTransparentBgSliderFill instanceof HTMLElement) {
      elements.exportTransparentBgSliderFill.style.height = `${percent}%`;
    }
    if (elements.exportTransparentBgSliderThumb instanceof HTMLElement) {
      elements.exportTransparentBgSliderThumb.style.bottom = `${percent}%`;
    }
  }
  if (elements.exportTransparentBgToleranceValue instanceof HTMLElement) {
    elements.exportTransparentBgToleranceValue.textContent = String(tolerance);
  }
  if (elements.exportTransparentBgProtectionSlider instanceof HTMLElement) {
    const ratio = EXPORT_BACKGROUND_PROTECTION_MAX > 0 ? protection / EXPORT_BACKGROUND_PROTECTION_MAX : 0;
    const percent = Math.max(0, Math.min(100, ratio * 100));
    elements.exportTransparentBgProtectionSlider.setAttribute('aria-valuenow', String(protection));
    elements.exportTransparentBgProtectionSlider.setAttribute('aria-valuetext', String(protection));
    if (elements.exportTransparentBgProtectionSliderFill instanceof HTMLElement) {
      elements.exportTransparentBgProtectionSliderFill.style.height = `${percent}%`;
    }
    if (elements.exportTransparentBgProtectionSliderThumb instanceof HTMLElement) {
      elements.exportTransparentBgProtectionSliderThumb.style.bottom = `${percent}%`;
    }
  }
  if (elements.exportTransparentBgProtectionValue instanceof HTMLElement) {
    elements.exportTransparentBgProtectionValue.textContent = String(protection);
  }
  logExportTransparentDebug('ui sync');
}

function getCurrentExportEditOptions() {
  return {
    cropEnabled: state.exportWizardCropEnabled === true,
    cropRect: clampExportCropRect(state.exportWizardCropRect),
    transparentBackground: state.exportWizardTransparentBackground === true,
    backgroundTolerance: normalizeExportBackgroundTolerance(state.exportWizardBackgroundTolerance),
    backgroundProtection: normalizeExportBackgroundProtection(state.exportWizardBackgroundProtection)
  };
}

function getCurrentExportPreviewEditOptions(preset = state.exportWizardPreset) {
  const editOptions = getCurrentExportEditOptions();
  if (normalizeExportPreset(preset) !== EXPORT_PRESET_FAVICON) {
    return editOptions;
  }
  return {
    ...editOptions,
    cropEnabled: false,
    cropRect: createFullExportCropRect()
  };
}

function normalizeExportCompressionLevel(level) {
  const safeLevel = String(level || '').trim().toLowerCase();
  return EXPORT_COMPRESSION_LEVELS.includes(safeLevel) ? safeLevel : EXPORT_DEFAULT_COMPRESSION_LEVEL;
}

function exportCompressionLevelLabel(level) {
  const safeLevel = normalizeExportCompressionLevel(level);
  if (safeLevel === 'optimal') {
    return 'Agresivo';
  }
  if (safeLevel === 'aggressive') {
    return 'Extremo';
  }
  return 'Óptimo';
}

function updateExportCompressionLevelUi() {
  if (elements.exportCompressionLevelSelect instanceof HTMLSelectElement) {
    elements.exportCompressionLevelSelect.value = normalizeExportCompressionLevel(state.exportWizardCompressionLevel);
  }
}

function updateExportCompressionUi() {
  const safeCompression = normalizeExportCompression(state.exportWizardCompression);
  if (elements.exportCompressionRange instanceof HTMLInputElement) {
    elements.exportCompressionRange.value = String(safeCompression);
  }
  if (elements.exportCompressionValue instanceof HTMLElement) {
    elements.exportCompressionValue.textContent = `${safeCompression}%`;
  }
}

function setExportWizardPreviewLoading(isLoading, options = {}) {
  if (!(elements.exportWizardPreviewLoading instanceof HTMLElement)) {
    return;
  }
  if (state.exportWizardPreviewLoadingTimeout !== null) {
    window.clearTimeout(state.exportWizardPreviewLoadingTimeout);
    state.exportWizardPreviewLoadingTimeout = null;
  }
  if (isLoading !== true) {
    elements.exportWizardPreviewLoading.classList.add('hidden');
    return;
  }
  const delayMs = Math.max(0, Math.round(Number(options.delayMs) || 0));
  if (delayMs <= 0) {
    elements.exportWizardPreviewLoading.classList.remove('hidden');
    return;
  }
  const expectedToken = Number.isFinite(Number(options.token))
    ? Math.max(0, Number(options.token))
    : state.exportWizardPreviewToken;
  state.exportWizardPreviewLoadingTimeout = window.setTimeout(() => {
    state.exportWizardPreviewLoadingTimeout = null;
    if (expectedToken !== state.exportWizardPreviewToken || !state.exportWizardItem?.dataUrl) {
      return;
    }
    elements.exportWizardPreviewLoading.classList.remove('hidden');
  }, delayMs);
}

function setExportWizardCropEditorPreviewUrl(nextUrl) {
  const safeNextUrl = String(nextUrl || '').trim();
  if (state.exportWizardCropEditorPreviewUrl && state.exportWizardCropEditorPreviewUrl !== safeNextUrl) {
    URL.revokeObjectURL(state.exportWizardCropEditorPreviewUrl);
  }
  state.exportWizardCropEditorPreviewUrl = safeNextUrl;
}

function setExportRecommendedCompression(compression) {
  const safeCompression = Number.isFinite(Number(compression))
    ? normalizeExportCompression(compression)
    : null;
  state.exportWizardRecommendedCompression = safeCompression;
  if (!(elements.exportCompressionRecommendedDot instanceof HTMLElement)) {
    return;
  }
  if (safeCompression === null) {
    elements.exportCompressionRecommendedDot.classList.add('hidden');
    elements.exportCompressionRecommendedDot.removeAttribute('style');
    return;
  }
  const dotRatio = Math.min(1, Math.max(0, safeCompression / 100));
  elements.exportCompressionRecommendedDot.classList.remove('hidden');
  elements.exportCompressionRecommendedDot.style.setProperty('--recommended-ratio', String(dotRatio));
  elements.exportCompressionRecommendedDot.title = `${exportCompressionLevelLabel(state.exportWizardCompressionLevel)} ${safeCompression}%`;
}

function setExportPassthroughCompression(compression) {
  state.exportWizardPassthroughCompression = Number.isFinite(Number(compression))
    ? normalizeExportCompression(compression)
    : null;
}

function normalizeExportRecommendationLevel(result, fallback = null) {
  const source = result && typeof result === 'object' ? result : {};
  const compression = Number.isFinite(Number(source.compression))
    ? normalizeExportCompression(source.compression)
    : Number.isFinite(Number(fallback?.compression))
      ? normalizeExportCompression(fallback.compression)
      : null;
  const size = Math.max(0, Number(source.size) || Number(fallback?.size) || 0);
  if (compression === null || size <= 0) {
    return null;
  }
  return {
    compression,
    size,
    analyzedCompression: Number.isFinite(Number(source.analyzedCompression))
      ? normalizeExportCompression(source.analyzedCompression)
      : Number.isFinite(Number(fallback?.analyzedCompression))
        ? normalizeExportCompression(fallback.analyzedCompression)
        : compression,
    forcedCompression: Number.isFinite(Number(source.forcedCompression))
      ? normalizeExportCompression(source.forcedCompression)
      : Number.isFinite(Number(fallback?.forcedCompression))
        ? normalizeExportCompression(fallback.forcedCompression)
        : null,
    passthroughCompression: Number.isFinite(Number(source.passthroughCompression))
      ? normalizeExportCompression(source.passthroughCompression)
      : Number.isFinite(Number(fallback?.passthroughCompression))
        ? normalizeExportCompression(fallback.passthroughCompression)
        : null
  };
}

function cloneExportRecommendationLevels(levels = {}) {
  const safeLevels = {};
  for (const key of EXPORT_COMPRESSION_LEVELS) {
    const normalized = normalizeExportRecommendationLevel(levels?.[key]);
    if (normalized) {
      safeLevels[key] = { ...normalized };
    }
  }
  return safeLevels;
}

function resolveExportRecommendationResultForLevel(result, level = state.exportWizardCompressionLevel) {
  const normalized = normalizeExportRecommendationResult(result);
  if (!normalized) {
    return null;
  }
  const safeLevel = normalizeExportCompressionLevel(level);
  return normalizeExportRecommendationLevel(normalized.levels?.[safeLevel], normalized) || { ...normalized };
}

function normalizeExportRecommendationResult(result) {
  const base = normalizeExportRecommendationLevel(result);
  if (!base) {
    return null;
  }
  const rawLevels = result && typeof result === 'object' && result.levels && typeof result.levels === 'object'
    ? result.levels
    : {};
  const conservative = normalizeExportRecommendationLevel(rawLevels.conservative, base) || { ...base };
  const optimal = normalizeExportRecommendationLevel(rawLevels.optimal, conservative) || { ...conservative };
  const aggressive = normalizeExportRecommendationLevel(rawLevels.aggressive, optimal) || { ...optimal };
  return {
    ...conservative,
    levels: {
      conservative,
      optimal,
      aggressive
    }
  };
}

function cloneExportRecommendationResult(result) {
  const normalized = normalizeExportRecommendationResult(result);
  return normalized
    ? {
        ...normalized,
        levels: cloneExportRecommendationLevels(normalized.levels)
      }
    : null;
}

function cloneExportFormatAnalyses(analyses) {
  const next = {};
  if (!analyses || typeof analyses !== 'object') {
    return next;
  }
  for (const format of EXPORT_FORMAT_CANDIDATES) {
    const normalized = normalizeExportRecommendationResult(analyses[format]);
    if (normalized) {
      next[format] = normalized;
    }
  }
  return next;
}

function hasCompleteExportFormatAnalyses(analyses, candidates = EXPORT_FORMAT_CANDIDATES) {
  return candidates.every((format) => normalizeExportRecommendationResult(analyses?.[normalizeExportFormat(format)]));
}

function storeExportWizardFormatAnalyses(key, analyses) {
  state.exportWizardFormatAnalysesKey = String(key || '').trim();
  state.exportWizardFormatAnalyses = cloneExportFormatAnalyses(analyses);
  updateExportOptimalSummaryUi();
}

function storeExportWizardInitialFormatAnalyses(key, analyses) {
  state.exportWizardInitialFormatAnalysesKey = String(key || '').trim();
  state.exportWizardInitialFormatAnalyses = cloneExportFormatAnalyses(analyses);
  updateExportOptimalSummaryUi();
}

function getCachedExportRecommendation(cacheKey) {
  const safeKey = String(cacheKey || '').trim();
  if (!safeKey) {
    return null;
  }
  return cloneExportRecommendationResult(exportCompressionRecommendationCache.get(safeKey));
}

function setCachedExportRecommendation(cacheKey, result) {
  const safeKey = String(cacheKey || '').trim();
  const normalized = normalizeExportRecommendationResult(result);
  if (!safeKey || !normalized) {
    return null;
  }
  exportCompressionRecommendationCache.set(safeKey, normalized);
  return { ...normalized };
}

function getExportWizardFormatAnalysesForCurrentState() {
  if (!state.exportWizardItem?.dataUrl) {
    return {};
  }
  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  const analysisKey = resolveFormatRecommendationCacheKey(state.exportWizardItem, preset, editOptions);
  if (state.exportWizardFormatAnalysesKey !== analysisKey) {
    return {};
  }
  return cloneExportFormatAnalyses(state.exportWizardFormatAnalyses);
}

function createInitialExportWizardEditOptions() {
  return {
    cropEnabled: false,
    cropRect: createFullExportCropRect(),
    transparentBackground: false,
    backgroundTolerance: EXPORT_BACKGROUND_TOLERANCE_DEFAULT,
    backgroundProtection: EXPORT_BACKGROUND_PROTECTION_DEFAULT
  };
}

function getExportWizardInitialCompressionFloors(format) {
  const normalized = normalizeExportRecommendationResult(
    state.exportWizardInitialFormatAnalyses?.[normalizeExportFormat(format)]
  );
  if (!normalized) {
    return {};
  }
  return EXPORT_COMPRESSION_LEVELS.reduce((floors, level) => {
    const result = resolveExportRecommendationResultForLevel(normalized, level);
    if (result) {
      floors[level] = result.compression;
    }
    return floors;
  }, {});
}

function updateExportOptimalSummaryUi(analyses = null) {
  const targets = [
    ['image/webp', elements.exportWizardOptimalWebpValue],
    ['image/png', elements.exportWizardOptimalPngValue],
    ['image/jpeg', elements.exportWizardOptimalJpegValue]
  ];
  const currentAnalyses =
    analyses && typeof analyses === 'object'
      ? cloneExportFormatAnalyses(analyses)
      : getExportWizardFormatAnalysesForCurrentState();
  const activeLevel = normalizeExportCompressionLevel(state.exportWizardCompressionLevel);
  const activeLevelLabel = exportCompressionLevelLabel(activeLevel);
  let hasAnyValue = false;

  for (const [format, element] of targets) {
    if (!(element instanceof HTMLElement)) {
      continue;
    }
    const result = resolveExportRecommendationResultForLevel(currentAnalyses[format], activeLevel);
    if (result) {
      hasAnyValue = true;
    }
    if (!result) {
      element.textContent = '--';
      element.title = `${exportFormatLabel(format)} sin analizar`;
      continue;
    }
    const analyzedCompression = normalizeExportCompression(result.analyzedCompression);
    element.textContent = `${analyzedCompression}%`;
    element.title = `${exportFormatLabel(format)} ${activeLevelLabel.toLowerCase()} ${analyzedCompression}%`;
  }

  if (elements.exportWizardOptimalSummary instanceof HTMLElement) {
    elements.exportWizardOptimalSummary.classList.toggle('hidden', !state.exportWizardItem && !hasAnyValue);
  }
}

function rememberExportWizardFormatAnalysis(format, result, item, preset, editOptions = {}) {
  const normalized = normalizeExportRecommendationResult(result);
  if (!normalized || !item?.dataUrl) {
    return null;
  }
  const safeFormat = normalizeExportFormat(format);
  const analysisKey = resolveFormatRecommendationCacheKey(item, preset, editOptions);
  const currentAnalyses =
    state.exportWizardFormatAnalysesKey === analysisKey ? cloneExportFormatAnalyses(state.exportWizardFormatAnalyses) : {};
  currentAnalyses[safeFormat] = normalized;
  storeExportWizardFormatAnalyses(analysisKey, currentAnalyses);
  const cachedAnalyses = cloneExportFormatAnalyses(exportFormatAnalysisCache.get(analysisKey));
  cachedAnalyses[safeFormat] = normalized;
  exportFormatAnalysisCache.set(analysisKey, cachedAnalyses);
  return { ...normalized };
}

function resolveExportFormatRecommendationCacheKey(analysisKey, level = state.exportWizardCompressionLevel) {
  return `${String(analysisKey || '').trim()}|${normalizeExportCompressionLevel(level)}`;
}

function resolveBestExportFormatFromAnalyses(
  analyses,
  candidates = EXPORT_FORMAT_CANDIDATES,
  level = state.exportWizardCompressionLevel
) {
  const safeCandidates = Array.isArray(candidates)
    ? candidates.map((candidate) => normalizeExportFormat(candidate))
    : EXPORT_FORMAT_CANDIDATES;
  const available = safeCandidates
    .map((format) => ({
      format,
      result: resolveExportRecommendationResultForLevel(analyses?.[format], level)
    }))
    .filter((entry) => entry.result);
  if (available.length === 0) {
    return '';
  }
  const best = available.reduce((winner, entry) => {
    if (!winner) {
      return entry;
    }
    return entry.result.size < winner.result.size ? entry : winner;
  }, null);
  return normalizeExportFormat(best?.format || safeCandidates[0]);
}

function pickClosestExportResultBySize(results, targetSize) {
  const safeTargetSize = Math.max(1, Number(targetSize) || 1);
  const valid = Array.isArray(results)
    ? results.filter((entry) => Number.isFinite(entry?.compression) && Number.isFinite(entry?.size) && entry.size > 0)
    : [];
  if (valid.length === 0) {
    return null;
  }
  return valid.reduce((winner, entry) => {
    if (!winner) {
      return entry;
    }
    const winnerDiff = Math.abs(winner.size - safeTargetSize);
    const entryDiff = Math.abs(entry.size - safeTargetSize);
    if (entryDiff < winnerDiff) {
      return entry;
    }
    if (entryDiff > winnerDiff) {
      return winner;
    }
    const winnerFits = winner.size <= safeTargetSize;
    const entryFits = entry.size <= safeTargetSize;
    if (entryFits !== winnerFits) {
      return entryFits ? entry : winner;
    }
    if (entry.size !== winner.size) {
      return entry.size < winner.size ? entry : winner;
    }
    return entry.compression > winner.compression ? entry : winner;
  }, null);
}

async function measureExportSizeForCompression(canvas, format, preset, compression) {
  const safeCompression = normalizeExportCompression(compression);
  const blob = await encodeCanvasToExportBlob(canvas, {
    format,
    quality: resolveExportQuality(format, preset, safeCompression),
    pngCompression: safeCompression
  });
  return {
    compression: safeCompression,
    size: Math.max(1, Number(blob?.size) || 1)
  };
}

function resolveExportPreviewTargetSize(item, preset, editOptions = {}) {
  const safePreset = normalizeExportPreset(preset);
  if (safePreset !== EXPORT_PRESET_FAVICON) {
    return 0;
  }
  return Math.min(128, resolveFaviconMaxSourceSize(item, editOptions));
}

function resolveExportValidationTargetSize(item, preset, editOptions = {}) {
  const safePreset = normalizeExportPreset(preset);
  if (safePreset !== EXPORT_PRESET_FAVICON) {
    return 0;
  }
  return resolveFaviconMaxSourceSize(item, editOptions);
}

function resolveRecommendationCacheKey(item, format, preset, editOptions = {}) {
  const safeFormat = normalizeExportFormat(format);
  const safePreset = normalizeExportPreset(preset);
  const itemId = String(item?.id || '').trim();
  const width = Math.max(0, Number(item?.width) || 0);
  const height = Math.max(0, Number(item?.height) || 0);
  const size = Math.max(0, Number(item?.fileSize) || 0) || getDataUrlByteLength(item?.dataUrl);
  const sourceSignature = `${itemId}|${width}x${height}|${size}`;
  const rect = clampExportCropRect(editOptions.cropRect);
  const editSignature = [
    editOptions.cropEnabled === true ? 'crop' : 'full',
    `${Math.round(rect.x * 100)}-${Math.round(rect.y * 100)}-${Math.round(rect.width * 100)}-${Math.round(rect.height * 100)}`,
    editOptions.transparentBackground === true
      ? `alpha-${normalizeExportBackgroundTolerance(editOptions.backgroundTolerance)}-${normalizeExportBackgroundProtection(editOptions.backgroundProtection)}`
      : 'opaque'
  ].join('|');
  return `${sourceSignature}|${safeFormat}|${safePreset}|${editSignature}`;
}

function resolveFormatRecommendationCacheKey(item, preset, editOptions = {}) {
  const safePreset = normalizeExportPreset(preset);
  const itemId = String(item?.id || '').trim();
  const width = Math.max(0, Number(item?.width) || 0);
  const height = Math.max(0, Number(item?.height) || 0);
  const size = Math.max(0, Number(item?.fileSize) || 0) || getDataUrlByteLength(item?.dataUrl);
  const rect = clampExportCropRect(editOptions.cropRect);
  const editSignature = [
    editOptions.cropEnabled === true ? 'crop' : 'full',
    `${Math.round(rect.x * 100)}-${Math.round(rect.y * 100)}-${Math.round(rect.width * 100)}-${Math.round(rect.height * 100)}`,
    editOptions.transparentBackground === true
      ? `alpha-${normalizeExportBackgroundTolerance(editOptions.backgroundTolerance)}-${normalizeExportBackgroundProtection(editOptions.backgroundProtection)}`
      : 'opaque'
  ].join('|');
  return `${itemId}|${width}x${height}|${size}|${safePreset}|${editSignature}`;
}

async function ensureExportWizardFormatAnalyses(options = {}) {
  if (!state.exportWizardItem?.dataUrl) {
    return {};
  }
  const force = options.force === true;
  const item = state.exportWizardItem;
  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  const candidates = resolveExportFormatCandidates(preset, editOptions);
  const analysisKey = resolveFormatRecommendationCacheKey(item, preset, editOptions);

  if (!force && state.exportWizardFormatAnalysesKey === analysisKey) {
    const currentAnalyses = cloneExportFormatAnalyses(state.exportWizardFormatAnalyses);
    if (hasCompleteExportFormatAnalyses(currentAnalyses, candidates)) {
      return currentAnalyses;
    }
  }

  if (!force) {
    const cachedAnalyses = cloneExportFormatAnalyses(exportFormatAnalysisCache.get(analysisKey));
    if (hasCompleteExportFormatAnalyses(cachedAnalyses, candidates)) {
      storeExportWizardFormatAnalyses(analysisKey, cachedAnalyses);
      const bestFormat = resolveBestExportFormatFromAnalyses(cachedAnalyses, candidates);
      if (bestFormat) {
        exportFormatRecommendationCache.set(resolveExportFormatRecommendationCacheKey(analysisKey), bestFormat);
      }
      return cachedAnalyses;
    }
  }

  const token = state.exportWizardFormatAnalysisToken + 1;
  state.exportWizardFormatAnalysisToken = token;

  const analyses = {};
  try {
    const results = await Promise.all(
      candidates.map(async (format) => {
        const cacheKey = resolveRecommendationCacheKey(item, format, preset, editOptions);
        const cached = !force ? getCachedExportRecommendation(cacheKey) : null;
        if (cached) {
          return [format, cached];
        }
        const recommended = await estimateRecommendedCompressionResultForFormat(format, item, {
          format,
          preset,
          ...editOptions
        });
        return [format, setCachedExportRecommendation(cacheKey, recommended)];
      })
    );

    if (token !== state.exportWizardFormatAnalysisToken || !state.exportWizardItem?.dataUrl) {
      return getExportWizardFormatAnalysesForCurrentState();
    }

    for (const [format, result] of results) {
      const normalized = normalizeExportRecommendationResult(result);
      if (normalized) {
        analyses[normalizeExportFormat(format)] = normalized;
      }
    }

    const safeAnalyses = cloneExportFormatAnalyses(analyses);
    exportFormatAnalysisCache.set(analysisKey, safeAnalyses);
    storeExportWizardFormatAnalyses(analysisKey, safeAnalyses);

    const bestFormat = resolveBestExportFormatFromAnalyses(safeAnalyses, candidates);
    if (bestFormat) {
      exportFormatRecommendationCache.set(resolveExportFormatRecommendationCacheKey(analysisKey), bestFormat);
    }

    return safeAnalyses;
  } catch (error) {
    console.error('Format analysis failed', error);
    if (token !== state.exportWizardFormatAnalysisToken) {
      return getExportWizardFormatAnalysesForCurrentState();
    }
    return {};
  }
}

async function ensureInitialExportWizardFormatAnalyses(options = {}) {
  if (!state.exportWizardItem?.dataUrl) {
    return {};
  }
  const force = options.force === true;
  const item = state.exportWizardItem;
  const preset = EXPORT_PRESET_ORIGINAL;
  const editOptions = createInitialExportWizardEditOptions();
  const candidates = resolveExportFormatCandidates(preset, editOptions);
  const analysisKey = resolveFormatRecommendationCacheKey(item, preset, editOptions);

  if (!force && state.exportWizardInitialFormatAnalysesKey === analysisKey) {
    const currentAnalyses = cloneExportFormatAnalyses(state.exportWizardInitialFormatAnalyses);
    if (hasCompleteExportFormatAnalyses(currentAnalyses, candidates)) {
      return currentAnalyses;
    }
  }

  if (!force) {
    const cachedAnalyses = cloneExportFormatAnalyses(exportFormatAnalysisCache.get(analysisKey));
    if (hasCompleteExportFormatAnalyses(cachedAnalyses, candidates)) {
      storeExportWizardInitialFormatAnalyses(analysisKey, cachedAnalyses);
      return cachedAnalyses;
    }
  }

  const token = state.exportWizardInitialFormatAnalysisToken + 1;
  state.exportWizardInitialFormatAnalysisToken = token;

  const analyses = {};
  try {
    const results = await Promise.all(
      candidates.map(async (format) => {
        const cacheKey = resolveRecommendationCacheKey(item, format, preset, editOptions);
        const cached = !force ? getCachedExportRecommendation(cacheKey) : null;
        if (cached) {
          return [format, cached];
        }
        const recommended = await estimateRecommendedCompressionResultForFormat(format, item, {
          format,
          preset,
          ...editOptions,
          ignoreInitialCompressionFloor: true
        });
        return [format, setCachedExportRecommendation(cacheKey, recommended)];
      })
    );

    if (token !== state.exportWizardInitialFormatAnalysisToken || !state.exportWizardItem?.dataUrl) {
      return cloneExportFormatAnalyses(state.exportWizardInitialFormatAnalyses);
    }

    for (const [format, result] of results) {
      const normalized = normalizeExportRecommendationResult(result);
      if (normalized) {
        analyses[normalizeExportFormat(format)] = normalized;
      }
    }

    const safeAnalyses = cloneExportFormatAnalyses(analyses);
    exportFormatAnalysisCache.set(analysisKey, safeAnalyses);
    storeExportWizardInitialFormatAnalyses(analysisKey, safeAnalyses);
    return safeAnalyses;
  } catch (error) {
    console.error('Initial format analysis failed', error);
    if (token !== state.exportWizardInitialFormatAnalysisToken) {
      return cloneExportFormatAnalyses(state.exportWizardInitialFormatAnalyses);
    }
    return {};
  }
}

function resolveExportFormatCandidates(preset, editOptions = {}) {
  const safePreset = normalizeExportPreset(preset);
  if (safePreset === EXPORT_PRESET_FAVICON) {
    return ['image/png'];
  }
  const formats = EXPORT_FORMAT_CANDIDATES
    .map((candidate) => resolveEffectiveExportFormat(candidate, safePreset, editOptions))
    .filter(Boolean);
  return Array.from(new Set(formats));
}

function resolveRecommendationErrorThreshold(format) {
  const safeFormat = normalizeExportFormat(format);
  return EXPORT_RECOMMENDATION_ERROR_THRESHOLD[safeFormat] || EXPORT_RECOMMENDATION_ERROR_THRESHOLD['image/webp'];
}

function resolveRecommendationRefinementConfig(format, strategy = null) {
  const safeFormat = normalizeExportFormat(format);
  const radiusOverride = Math.max(0, Number(strategy?.refinementRadius) || 0);
  const midpointRadiusOverride = Math.max(0, Number(strategy?.refinementMidpointRadius) || 0);
  const stepOverride = Math.max(1, Number(strategy?.refinementStep) || 0);
  if (safeFormat === 'image/png') {
    return {
      radius: radiusOverride || 8,
      midpointRadius: midpointRadiusOverride || 4,
      step: stepOverride || 2
    };
  }
  return {
    radius: radiusOverride || 12,
    midpointRadius: midpointRadiusOverride || 6,
    step: stepOverride || 2
  };
}

function buildRefinedRecommendationCandidates(baseCandidates, recommendedCompressions, format, strategy = null) {
  const safeFormat = normalizeExportFormat(format);
  const refinement = resolveRecommendationRefinementConfig(safeFormat, strategy);
  const maxCompression = Math.max(0, Math.min(100, Number(strategy?.maxCompression) || 100));
  const candidateSet = new Set();
  const safeBaseCandidates = Array.isArray(baseCandidates) ? baseCandidates : EXPORT_RECOMMENDATION_CANDIDATES;
  const addCandidate = (value) => {
    const normalized = normalizeExportCompression(value);
    if (normalized > maxCompression) {
      return;
    }
    candidateSet.add(normalized);
  };
  const addCandidateBand = (center, radius = refinement.radius) => {
    const safeCenter = normalizeExportCompression(center);
    const start = Math.max(0, safeCenter - Math.max(0, radius));
    const end = Math.min(maxCompression, safeCenter + Math.max(0, radius));
    for (let value = start; value <= end; value += refinement.step) {
      addCandidate(value);
    }
    addCandidate(start);
    addCandidate(end);
  };

  for (const candidate of safeBaseCandidates) {
    addCandidate(candidate);
  }

  const pivots = Object.values(recommendedCompressions || {})
    .filter((value) => Number.isFinite(Number(value)))
    .map((value) => normalizeExportCompression(value));
  const sortedPivots = Array.from(new Set(pivots)).sort((left, right) => left - right);
  for (const pivot of sortedPivots) {
    addCandidateBand(pivot, refinement.radius);
  }
  for (let index = 0; index < sortedPivots.length - 1; index += 1) {
    const midpoint = Math.round((sortedPivots[index] + sortedPivots[index + 1]) / 2);
    addCandidateBand(midpoint, refinement.midpointRadius);
  }

  if ((safeFormat === 'image/webp' || safeFormat === 'image/jpeg') && sortedPivots.some((value) => value <= 0)) {
    for (const candidate of [2, 4, 6, 10, 12, 14]) {
      addCandidate(candidate);
    }
  }

  return Array.from(candidateSet).sort((left, right) => left - right);
}

async function refineRecommendationResults(context, format, baseCandidates, results, strategy = null) {
  let safeResults = Array.isArray(results)
    ? results
        .filter((entry) => Number.isFinite(entry?.compression))
        .sort((left, right) => left.compression - right.compression)
    : [];
  let safeCandidates = Array.isArray(baseCandidates)
    ? baseCandidates.map((candidate) => normalizeExportCompression(candidate)).sort((left, right) => left - right)
    : [...EXPORT_RECOMMENDATION_CANDIDATES];

  for (let pass = 0; pass < 2; pass += 1) {
    const recommendedCompressions = pickRecommendedCompressionLevels(safeResults, format, strategy);
    const refinedCandidates = buildRefinedRecommendationCandidates(safeCandidates, recommendedCompressions, format, strategy);
    const measuredCompressions = new Set(safeResults.map((entry) => normalizeExportCompression(entry.compression)));
    const missingCandidates = refinedCandidates.filter((candidate) => !measuredCompressions.has(candidate));
    safeCandidates = refinedCandidates;
    if (missingCandidates.length === 0) {
      return {
        candidateCompressions: safeCandidates,
        results: safeResults,
        recommendedCompressions
      };
    }
    const extraResults = await buildExportRecommendationResults(context, missingCandidates);
    safeResults = safeResults.concat(extraResults).sort((left, right) => left.compression - right.compression);
  }

  return {
    candidateCompressions: safeCandidates,
    results: safeResults,
    recommendedCompressions: pickRecommendedCompressionLevels(safeResults, format, strategy)
  };
}

function analyzeRecommendationCanvas(canvas, options = {}) {
  const ctx = canvas?.getContext?.('2d');
  const width = Math.max(1, Number(canvas?.width) || 1);
  const height = Math.max(1, Number(canvas?.height) || 1);
  if (!ctx) {
    return null;
  }

  const imageData = ctx.getImageData(0, 0, width, height).data;
  const sampleDivisor = Math.max(1, Number(options.sampleDivisor) || 96);
  const step = Math.max(1, Math.ceil(Math.max(width, height) / sampleDivisor));
  const exactColors = new Set();
  const bucketColors = new Set();
  const maxTrackedColors = 2048;
  let pairCount = 0;
  let flatPairs = 0;
  let smoothPairs = 0;
  let hardPairs = 0;

  const readIndex = (x, y) => ((y * width) + x) * 4;
  const addColor = (index) => {
    const r = imageData[index];
    const g = imageData[index + 1];
    const b = imageData[index + 2];
    const a = imageData[index + 3];
    if (exactColors.size <= maxTrackedColors) {
      const key = (((r << 24) >>> 0) | (g << 16) | (b << 8) | a) >>> 0;
      exactColors.add(key);
    }
    if (bucketColors.size <= maxTrackedColors) {
      const key = ((r >> 3) << 14) | ((g >> 3) << 9) | ((b >> 3) << 4) | (a >> 4);
      bucketColors.add(key);
    }
  };

  const measurePair = (leftIndex, rightIndex) => {
    const diff =
      Math.abs(imageData[leftIndex] - imageData[rightIndex]) +
      Math.abs(imageData[leftIndex + 1] - imageData[rightIndex + 1]) +
      Math.abs(imageData[leftIndex + 2] - imageData[rightIndex + 2]) +
      (Math.abs(imageData[leftIndex + 3] - imageData[rightIndex + 3]) * 0.35);
    pairCount += 1;
    if (diff <= 8) {
      flatPairs += 1;
      return;
    }
    if (diff <= 42) {
      smoothPairs += 1;
      return;
    }
    hardPairs += 1;
  };

  for (let y = 0; y < height; y += step) {
    for (let x = 0; x < width; x += step) {
      const index = readIndex(x, y);
      addColor(index);
      if (x + step < width) {
        measurePair(index, readIndex(x + step, y));
      }
      if (y + step < height) {
        measurePair(index, readIndex(x, y + step));
      }
    }
  }

  const exactColorCount = exactColors.size > maxTrackedColors ? maxTrackedColors + 1 : exactColors.size;
  const bucketColorCount = bucketColors.size > maxTrackedColors ? maxTrackedColors + 1 : bucketColors.size;
  const safePairCount = Math.max(1, pairCount);
  const flatRatio = flatPairs / safePairCount;
  const smoothRatio = smoothPairs / safePairCount;
  const hardRatio = hardPairs / safePairCount;
  const edgeRatio = (smoothPairs + hardPairs) / safePairCount;
  const hardEdgeShare = hardPairs / Math.max(1, smoothPairs + hardPairs);

  return {
    exactColorCount,
    bucketColorCount,
    flatRatio,
    smoothRatio,
    hardRatio,
    edgeRatio,
    hardEdgeShare,
    paletteFriendly: exactColorCount <= 16 && smoothRatio < 0.12,
    gradientRisk: bucketColorCount >= 96 && smoothRatio >= 0.18 && hardRatio <= 0.5,
    hardEdgeLowColorRisk:
      bucketColorCount <= 256 &&
      exactColorCount <= 512 &&
      flatRatio >= 0.3 &&
      smoothRatio <= 0.32 &&
      hardRatio >= 0.14 &&
      (flatRatio + hardRatio) >= 0.68,
    strongHardEdgeLowColorRisk:
      bucketColorCount <= 160 &&
      exactColorCount <= 224 &&
      flatRatio >= 0.38 &&
      smoothRatio <= 0.24 &&
      hardRatio >= 0.18 &&
      (flatRatio + hardRatio) >= 0.74,
    webpHardEdgeRisk:
      (
        bucketColorCount <= 420 &&
        exactColorCount <= 960 &&
        flatRatio >= 0.22 &&
        smoothRatio <= 0.42 &&
        hardRatio >= 0.08 &&
        (flatRatio + hardRatio) >= 0.56
      ) ||
      (
        bucketColorCount <= 960 &&
        exactColorCount <= 1800 &&
        flatRatio >= 0.68 &&
        edgeRatio >= 0.025 &&
        edgeRatio <= 0.24 &&
        hardEdgeShare >= 0.16
      ),
    webpStrongHardEdgeRisk:
      (
        bucketColorCount <= 280 &&
        exactColorCount <= 640 &&
        flatRatio >= 0.28 &&
        smoothRatio <= 0.34 &&
        hardRatio >= 0.11 &&
        (flatRatio + hardRatio) >= 0.62
      ) ||
      (
        bucketColorCount <= 640 &&
        exactColorCount <= 1280 &&
        flatRatio >= 0.76 &&
        edgeRatio >= 0.03 &&
        edgeRatio <= 0.18 &&
        hardEdgeShare >= 0.22
      )
  };
}

function analyzePngRecommendationCanvas(canvas) {
  return analyzeRecommendationCanvas(canvas, { sampleDivisor: 96 });
}

function resolvePngRecommendationStrategy(analysis) {
  const base = {
    maxCompression: 88,
    thresholdScale: 0.86,
    sizeGainRatio: 0.995,
    preferredMinColors: 0,
    hardMinColors: 0
  };
  if (!analysis) {
    return base;
  }
  const exactColorCount = Math.max(0, Number(analysis.exactColorCount) || 0);
  const bucketColorCount = Math.max(0, Number(analysis.bucketColorCount) || 0);
  const preferredLowColorFloor = exactColorCount > 0
    ? Math.max(8, Math.min(160, Math.round(exactColorCount * 0.9)))
    : 0;
  const hardLowColorFloor = exactColorCount > 0
    ? Math.max(4, Math.min(128, Math.round(exactColorCount * 0.72)))
    : 0;
  if (analysis.paletteFriendly) {
    return {
      ...base,
      maxCompression: 64,
      thresholdScale: 0.82,
      sizeGainRatio: 0.997,
      preferredMinColors: preferredLowColorFloor,
      hardMinColors: hardLowColorFloor
    };
  }
  if (analysis.gradientRisk) {
    return {
      ...base,
      maxCompression: 60,
      thresholdScale: 0.34,
      sizeGainRatio: 0.99,
      preferredMinColors: 128,
      hardMinColors: 96
    };
  }
  if (bucketColorCount >= 160) {
    return {
      ...base,
      maxCompression: 72,
      thresholdScale: 0.52,
      sizeGainRatio: 0.993,
      preferredMinColors: 96,
      hardMinColors: 48
    };
  }
  if (exactColorCount > 0 && exactColorCount <= 32) {
    return {
      ...base,
      maxCompression: 56,
      thresholdScale: 0.78,
      sizeGainRatio: 0.997,
      preferredMinColors: Math.max(preferredLowColorFloor, exactColorCount),
      hardMinColors: Math.max(hardLowColorFloor, Math.round(exactColorCount * 0.82))
    };
  }
  return base;
}

function resolveLossyRecommendationStrategy(analysis, format) {
  const safeFormat = normalizeExportFormat(format);
  const isWebp = safeFormat === 'image/webp';
  const base = {
    maxCompression: 100,
    thresholdScale: 1,
    sizeGainRatio: 1,
    compressionBiasScale: 1,
    promotionThresholdMultiplierScale: 1,
    promotionMarginScale: 1,
    aggressiveUpgradeSizeRatioScale: 1,
    aggressiveUpgradeErrorMultiplierScale: 1,
    preferHigherCompressionOnTie: true,
    forcePositiveLossyRecommendation: true,
    refinementRadius: 12,
    refinementMidpointRadius: 6,
    refinementStep: 2,
    profileThresholdScales: {}
  };
  if (!analysis || (safeFormat !== 'image/webp' && safeFormat !== 'image/jpeg')) {
    return base;
  }

  if (isWebp && analysis.webpStrongHardEdgeRisk) {
    return {
      ...base,
      maxCompression: 16,
      thresholdScale: 0.2,
      sizeGainRatio: 0.99,
      compressionBiasScale: 0.04,
      promotionThresholdMultiplierScale: 0.12,
      promotionMarginScale: 0.14,
      aggressiveUpgradeSizeRatioScale: 0.76,
      aggressiveUpgradeErrorMultiplierScale: 0.42,
      preferHigherCompressionOnTie: false,
      forcePositiveLossyRecommendation: false,
      refinementRadius: 2,
      refinementMidpointRadius: 2,
      profileThresholdScales: {
        conservative: 0.42,
        optimal: 0.5,
        aggressive: 0.62
      }
    };
  }

  if (isWebp && analysis.webpHardEdgeRisk) {
    return {
      ...base,
      maxCompression: 22,
      thresholdScale: 0.28,
      sizeGainRatio: 0.991,
      compressionBiasScale: 0.08,
      promotionThresholdMultiplierScale: 0.18,
      promotionMarginScale: 0.2,
      aggressiveUpgradeSizeRatioScale: 0.8,
      aggressiveUpgradeErrorMultiplierScale: 0.5,
      preferHigherCompressionOnTie: false,
      forcePositiveLossyRecommendation: false,
      refinementRadius: 3,
      refinementMidpointRadius: 2,
      profileThresholdScales: {
        conservative: 0.5,
        optimal: 0.58,
        aggressive: 0.7
      }
    };
  }

  if (analysis.strongHardEdgeLowColorRisk) {
    return {
      ...base,
      maxCompression: safeFormat === 'image/jpeg' ? 46 : 36,
      thresholdScale: isWebp ? 0.34 : 0.48,
      sizeGainRatio: 0.992,
      compressionBiasScale: isWebp ? 0.14 : 0.28,
      promotionThresholdMultiplierScale: isWebp ? 0.24 : 0.42,
      promotionMarginScale: isWebp ? 0.28 : 0.45,
      aggressiveUpgradeSizeRatioScale: isWebp ? 0.84 : 0.9,
      aggressiveUpgradeErrorMultiplierScale: isWebp ? 0.58 : 0.72,
      preferHigherCompressionOnTie: false,
      forcePositiveLossyRecommendation: false,
      refinementRadius: isWebp ? 4 : 6,
      refinementMidpointRadius: 4,
      profileThresholdScales: {
        conservative: isWebp ? 0.58 : 0.72,
        optimal: isWebp ? 0.66 : 0.8,
        aggressive: isWebp ? 0.78 : 0.9
      }
    };
  }

  if (analysis.hardEdgeLowColorRisk) {
    return {
      ...base,
      maxCompression: safeFormat === 'image/jpeg' ? 60 : 48,
      thresholdScale: isWebp ? 0.46 : 0.64,
      sizeGainRatio: 0.994,
      compressionBiasScale: isWebp ? 0.26 : 0.48,
      promotionThresholdMultiplierScale: isWebp ? 0.38 : 0.58,
      promotionMarginScale: isWebp ? 0.44 : 0.62,
      aggressiveUpgradeSizeRatioScale: isWebp ? 0.88 : 0.95,
      aggressiveUpgradeErrorMultiplierScale: isWebp ? 0.66 : 0.82,
      preferHigherCompressionOnTie: false,
      forcePositiveLossyRecommendation: false,
      refinementRadius: isWebp ? 6 : 8,
      refinementMidpointRadius: 4,
      profileThresholdScales: {
        conservative: isWebp ? 0.68 : 0.8,
        optimal: isWebp ? 0.76 : 0.88,
        aggressive: isWebp ? 0.86 : 0.94
      }
    };
  }

  return base;
}

function resolveRecommendationProfileConfig(profile, format, strategy = null) {
  const safeProfile = String(profile || 'conservative').trim().toLowerCase();
  const thresholdScale = Math.max(0.2, Number(strategy?.thresholdScale) || 1);
  const profileThresholdScale = Math.max(0.2, Number(strategy?.profileThresholdScales?.[safeProfile]) || 1);
  const threshold = resolveRecommendationErrorThreshold(format) * thresholdScale * profileThresholdScale;
  const sizeGainRatio = Math.min(0.9995, Math.max(0.9, Number(strategy?.sizeGainRatio) || 0.997));
  const compressionBiasScale = Math.max(0, Number(strategy?.compressionBiasScale) || 1);
  const promotionThresholdMultiplierScale = Math.max(0.2, Number(strategy?.promotionThresholdMultiplierScale) || 1);
  const promotionMarginScale = Math.max(0.2, Number(strategy?.promotionMarginScale) || 1);
  const aggressiveUpgradeSizeRatioScale = Math.max(0.8, Number(strategy?.aggressiveUpgradeSizeRatioScale) || 1);
  const aggressiveUpgradeErrorMultiplierScale = Math.max(0.5, Number(strategy?.aggressiveUpgradeErrorMultiplierScale) || 1);
  const preferHigherCompressionOnTie =
    typeof strategy?.preferHigherCompressionOnTie === 'boolean'
      ? strategy.preferHigherCompressionOnTie
      : null;
  const forcePositiveLossyRecommendation =
    typeof strategy?.forcePositiveLossyRecommendation === 'boolean'
      ? strategy.forcePositiveLossyRecommendation
      : null;
  if (safeProfile === 'aggressive') {
    return {
      threshold: threshold * 1.9,
      sizeGainRatio: 1,
      errorPenaltyScale: 0.03,
      softErrorScale: 0.01,
      colorPenaltyScale: 0.74,
      hardPenaltyScale: 0.74,
      compressionBias: 0.1 * compressionBiasScale,
      promotionThresholdMultiplier: 2.45 * promotionThresholdMultiplierScale,
      promotionSizeGainRatio: 1,
      promotionMargin: 0.18 * promotionMarginScale,
      preferHigherCompressionOnTie: preferHigherCompressionOnTie ?? true,
      forcePositiveLossyRecommendation: forcePositiveLossyRecommendation ?? true,
      aggressiveUpgradeSizeRatio: Math.min(0.98, 0.84 * aggressiveUpgradeSizeRatioScale),
      aggressiveUpgradeErrorMultiplier: 1.45 * aggressiveUpgradeErrorMultiplierScale
    };
  }
  if (safeProfile === 'optimal') {
    return {
      threshold: threshold * 1.58,
      sizeGainRatio: 1,
      errorPenaltyScale: 0.04,
      softErrorScale: 0.013,
      colorPenaltyScale: 0.82,
      hardPenaltyScale: 0.82,
      compressionBias: 0.068 * compressionBiasScale,
      promotionThresholdMultiplier: 2.15 * promotionThresholdMultiplierScale,
      promotionSizeGainRatio: 1,
      promotionMargin: 0.13 * promotionMarginScale,
      preferHigherCompressionOnTie: preferHigherCompressionOnTie ?? true,
      forcePositiveLossyRecommendation: forcePositiveLossyRecommendation ?? true
    };
  }
  return {
    threshold,
    sizeGainRatio,
    errorPenaltyScale: 0.06,
    softErrorScale: 0.02,
    colorPenaltyScale: 1,
    hardPenaltyScale: 1,
    compressionBias: 0 * compressionBiasScale,
    promotionThresholdMultiplier: 1.55 * promotionThresholdMultiplierScale,
    promotionSizeGainRatio: 0.995,
    promotionMargin: 0.05 * promotionMarginScale,
    preferHigherCompressionOnTie: preferHigherCompressionOnTie ?? false,
    forcePositiveLossyRecommendation: forcePositiveLossyRecommendation ?? false
  };
}

function pickRecommendedCompression(results, format, strategy = null, profile = 'conservative') {
  const safeProfile = String(profile || 'conservative').trim().toLowerCase();
  const valid = Array.isArray(results)
    ? results.filter((entry) => Number.isFinite(entry?.compression) && Number.isFinite(entry?.size) && Number.isFinite(entry?.error))
    : [];
  if (valid.length === 0) {
    return EXPORT_DEFAULT_COMPRESSION;
  }
  valid.sort((left, right) => left.compression - right.compression);
  const baseline = valid.find((entry) => entry.compression === 0) || valid[0];
  const baselineSize = Math.max(1, Number(baseline.size) || 1);
  const profileConfig = resolveRecommendationProfileConfig(profile, format, strategy);
  const threshold = profileConfig.threshold;
  const sizeGainRatio = profileConfig.sizeGainRatio;
  const preferredMinColors = Math.max(0, Number(strategy?.preferredMinColors) || 0);
  const hardMinColors = Math.max(0, Number(strategy?.hardMinColors) || 0);
  const isColorFloorBlocked = (entry) => {
    const colorCount = Math.max(0, Number(entry?.colorCount) || 0);
    return hardMinColors > 0 && colorCount > 0 && colorCount < hardMinColors;
  };
  const scoreEntry = (entry, thresholdOverride = threshold) => {
    const effectiveThreshold = Math.max(1, Number(thresholdOverride) || threshold);
    const sizeScore = entry.size / baselineSize;
    const errorPenalty = Math.max(0, entry.error - effectiveThreshold) * profileConfig.errorPenaltyScale;
    const softError = (entry.error / effectiveThreshold) * profileConfig.softErrorScale;
    const colorCount = Math.max(0, Number(entry?.colorCount) || 0);
    const colorPenalty =
      preferredMinColors > 0 && colorCount > 0 && colorCount < preferredMinColors
        ? ((preferredMinColors - colorCount) / preferredMinColors) * 0.35 * profileConfig.colorPenaltyScale
        : 0;
    const hardPenalty =
      hardMinColors > 0 && colorCount > 0 && colorCount < hardMinColors
        ? 0.85 * profileConfig.hardPenaltyScale
        : 0;
    return sizeScore + errorPenalty + softError + colorPenalty + hardPenalty - ((entry.compression / 100) * profileConfig.compressionBias);
  };
  const pickAggressiveUpgrade = (currentEntry) => {
    if (safeProfile !== 'aggressive' || !currentEntry) {
      return null;
    }
    const sizeRatio = Math.min(0.98, Math.max(0.6, Number(profileConfig.aggressiveUpgradeSizeRatio) || 0.9));
    const errorCap = threshold * Math.max(1.5, Number(profileConfig.aggressiveUpgradeErrorMultiplier) || 3.2);
    const candidates = valid.filter((entry) => {
      if (entry.compression <= currentEntry.compression) {
        return false;
      }
      if (isColorFloorBlocked(entry)) {
        return false;
      }
      return entry.size <= currentEntry.size * sizeRatio && entry.error <= errorCap;
    });
    if (candidates.length === 0) {
      return null;
    }
    return candidates.reduce((winner, entry) => {
      if (!winner) {
        return entry;
      }
      if (entry.size < winner.size) {
        return entry;
      }
      if (entry.size > winner.size) {
        return winner;
      }
      return entry.compression > winner.compression ? entry : winner;
    }, null);
  };

  const qualitySafe = valid.filter((entry) => {
    if (isColorFloorBlocked(entry)) {
      return false;
    }
    return entry.error <= threshold && entry.size <= baselineSize * sizeGainRatio;
  });
  if (qualitySafe.length > 0) {
    const qualitySafeChoice = qualitySafe[qualitySafe.length - 1];
    const aggressiveUpgrade = pickAggressiveUpgrade(qualitySafeChoice);
    return (aggressiveUpgrade || qualitySafeChoice).compression;
  }

  let best = valid[0];
  let bestScore = Number.POSITIVE_INFINITY;
  for (const entry of valid) {
    const score = scoreEntry(entry);
    if (score < bestScore) {
      bestScore = score;
      best = entry;
    }
  }
  const aggressiveUpgrade = pickAggressiveUpgrade(best);
  if (aggressiveUpgrade) {
    best = aggressiveUpgrade;
    bestScore = scoreEntry(aggressiveUpgrade, threshold * Math.max(1.5, Number(profileConfig.aggressiveUpgradeErrorMultiplier) || 3.2));
  }
  if (best.compression <= 0 && (format === 'image/webp' || format === 'image/jpeg')) {
    const relaxedThreshold = threshold * profileConfig.promotionThresholdMultiplier;
    const promoted = valid
      .filter(
        (entry) =>
          entry.compression > 0 &&
          entry.size <= baselineSize * profileConfig.promotionSizeGainRatio &&
          entry.error <= relaxedThreshold
      )
      .reduce((winner, entry) => {
        const score = scoreEntry(entry, relaxedThreshold);
        if (!winner) {
          return { entry, score };
        }
        if (score < winner.score) {
          return { entry, score };
        }
        if (score > winner.score) {
          return winner;
        }
        if (profileConfig.preferHigherCompressionOnTie) {
          return entry.compression > winner.entry.compression ? { entry, score } : winner;
        }
        return entry.compression < winner.entry.compression ? { entry, score } : winner;
      }, null);
    if (promoted && promoted.score <= bestScore + profileConfig.promotionMargin) {
      return promoted.entry.compression;
    }
    if (
      promoted &&
      profileConfig.forcePositiveLossyRecommendation === true &&
      promoted.entry.size < baselineSize
    ) {
      return promoted.entry.compression;
    }
  }
  return best.compression;
}

function pickLossyIntermediateCompression(results, format, strategy = null, lowerCompression = 0, upperCompression = 0) {
  const safeFormat = normalizeExportFormat(format);
  if (safeFormat !== 'image/webp' && safeFormat !== 'image/jpeg') {
    return null;
  }
  const valid = Array.isArray(results)
    ? results.filter((entry) => Number.isFinite(entry?.compression) && Number.isFinite(entry?.size) && Number.isFinite(entry?.error))
    : [];
  if (valid.length === 0) {
    return null;
  }
  valid.sort((left, right) => left.compression - right.compression);
  const baseline = valid.find((entry) => entry.compression === 0) || valid[0];
  const baselineSize = Math.max(1, Number(baseline.size) || 1);
  const safeLower = Math.max(0, normalizeExportCompression(lowerCompression));
  const safeUpper = Math.max(0, normalizeExportCompression(upperCompression));
  if (safeUpper <= safeLower || safeUpper <= 0) {
    return null;
  }

  const profileConfig = resolveRecommendationProfileConfig('optimal', safeFormat, strategy);
  const targetCompression = safeLower > 0
    ? Math.round(((safeLower + safeUpper) / 2) / 8) * 8
    : Math.max(8, Math.round((safeUpper * 0.58) / 8) * 8);
  const pickCandidate = (errorMultiplier) => {
    const errorCap = profileConfig.threshold * Math.max(1, Number(errorMultiplier) || 1);
    const candidates = valid.filter((entry) => (
      entry.compression > safeLower &&
      entry.compression <= safeUpper &&
      entry.size < baselineSize &&
      entry.error <= errorCap
    ));
    if (candidates.length === 0) {
      return null;
    }
    return candidates.reduce((winner, entry) => {
      if (!winner) {
        return entry;
      }
      const winnerDistance = Math.abs(winner.compression - targetCompression);
      const entryDistance = Math.abs(entry.compression - targetCompression);
      if (entryDistance < winnerDistance) {
        return entry;
      }
      if (entryDistance > winnerDistance) {
        return winner;
      }
      if (entry.size < winner.size) {
        return entry;
      }
      if (entry.size > winner.size) {
        return winner;
      }
      return entry.compression > winner.compression ? entry : winner;
    }, null);
  };

  return pickCandidate(1.35)?.compression ?? pickCandidate(1.7)?.compression ?? null;
}

function pickRecommendedCompressionLevels(results, format, strategy = null) {
  const safeFormat = normalizeExportFormat(format);
  const conservative = pickRecommendedCompression(results, safeFormat, strategy, 'conservative');
  let optimal = Math.max(conservative, pickRecommendedCompression(results, safeFormat, strategy, 'optimal'));
  let aggressive = Math.max(optimal, pickRecommendedCompression(results, safeFormat, strategy, 'aggressive'));
  if ((safeFormat === 'image/webp' || safeFormat === 'image/jpeg') && optimal <= 0 && aggressive > 0) {
    optimal = pickLossyIntermediateCompression(results, safeFormat, strategy, conservative, aggressive) || optimal;
    aggressive = Math.max(optimal, aggressive);
  }
  return {
    conservative: normalizeExportCompression(conservative),
    optimal: normalizeExportCompression(optimal),
    aggressive: normalizeExportCompression(aggressive)
  };
}

async function applyRecommendedCompressionForCurrentExport(options = {}) {
  if (!state.exportWizardItem?.dataUrl) {
    setExportRecommendedCompression(null);
    setExportPassthroughCompression(null);
    return;
  }
  const force = options.force === true;
  const apply = options.apply !== false;
  const item = state.exportWizardItem;
  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  const format = resolveEffectiveExportFormat(state.exportWizardFormat, preset, editOptions);
  const cacheKey = resolveRecommendationCacheKey(item, format, preset, editOptions);
  const analysisKey = resolveFormatRecommendationCacheKey(item, preset, editOptions);
  const token = state.exportWizardRecommendationToken + 1;
  state.exportWizardRecommendationToken = token;

  const analyses =
    options.analyses && typeof options.analyses === 'object'
      ? cloneExportFormatAnalyses(options.analyses)
      : getExportWizardFormatAnalysesForCurrentState();
  const cachedResult = normalizeExportRecommendationResult(analyses[format]) || getCachedExportRecommendation(cacheKey);
  if (cachedResult) {
    const activeRecommendation = resolveExportRecommendationResultForLevel(cachedResult);
    if (token !== state.exportWizardRecommendationToken || !state.exportWizardItem?.dataUrl) {
      return;
    }
    setExportRecommendedCompression(activeRecommendation?.compression);
    setExportPassthroughCompression(activeRecommendation?.passthroughCompression);
    if (apply && (!state.exportWizardCompressionTouched || force)) {
      state.exportWizardCompression = activeRecommendation?.compression ?? EXPORT_DEFAULT_COMPRESSION;
      updateExportCompressionUi();
      void refreshExportWizardPreview();
    }
    return;
  }
  setExportRecommendedCompression(null);
  setExportPassthroughCompression(null);

  try {
    const recommended = await estimateRecommendedCompressionResultForFormat(format, item, {
      format,
      preset,
      ...editOptions
    });
    if (token !== state.exportWizardRecommendationToken || !state.exportWizardItem?.dataUrl) {
      return;
    }
    const safeRecommended = setCachedExportRecommendation(cacheKey, recommended);
    if (!safeRecommended) {
      throw new Error('invalid_recommendation_result');
    }
    if (state.exportWizardFormatAnalysesKey === analysisKey) {
      rememberExportWizardFormatAnalysis(format, safeRecommended, item, preset, editOptions);
    }
    const activeRecommendation = resolveExportRecommendationResultForLevel(safeRecommended);
    setExportRecommendedCompression(activeRecommendation?.compression);
    setExportPassthroughCompression(activeRecommendation?.passthroughCompression);
    if (apply && (!state.exportWizardCompressionTouched || force)) {
      state.exportWizardCompression = activeRecommendation?.compression ?? EXPORT_DEFAULT_COMPRESSION;
      updateExportCompressionUi();
      void refreshExportWizardPreview();
    }
  } catch (error) {
    console.error('Compression recommendation failed', error);
    if (token !== state.exportWizardRecommendationToken) {
      return;
    }
    setExportRecommendedCompression(null);
    setExportPassthroughCompression(null);
  }
}

async function autoSelectBestFormatForCurrentExport(options = {}) {
  if (!state.exportWizardItem?.dataUrl) {
    return state.exportWizardFormat;
  }

  const force = options.force === true;
  if (state.exportWizardFormatTouched && !force) {
    return state.exportWizardFormat;
  }

  const item = state.exportWizardItem;
  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  const candidates = resolveExportFormatCandidates(preset, editOptions);
  const cacheKey = resolveFormatRecommendationCacheKey(item, preset, editOptions);
  const formatRecommendationCacheKey = resolveExportFormatRecommendationCacheKey(cacheKey);
  if (candidates.length <= 1) {
    const onlyFormat = normalizeExportFormat(candidates[0] || resolveEffectiveExportFormat(state.exportWizardFormat, preset, editOptions));
    if (state.exportWizardFormat !== onlyFormat) {
      state.exportWizardFormat = onlyFormat;
      syncExportFormatControlState();
    }
    return state.exportWizardFormat;
  }

  const applyRecommendedFormat = (format) => {
    const safeFormat = normalizeExportFormat(format);
    if (state.exportWizardFormat !== safeFormat) {
      state.exportWizardFormat = safeFormat;
      syncExportFormatControlState();
    }
    return state.exportWizardFormat;
  };

  const hasProvidedAnalyses = options.analyses && typeof options.analyses === 'object';
  let analyses =
    hasProvidedAnalyses
      ? cloneExportFormatAnalyses(options.analyses)
      : getExportWizardFormatAnalysesForCurrentState();
  if (!hasCompleteExportFormatAnalyses(analyses, candidates) && !hasProvidedAnalyses) {
    analyses = await ensureExportWizardFormatAnalyses({ force: options.forceAnalysis === true });
  }
  if (hasCompleteExportFormatAnalyses(analyses, candidates)) {
    const bestFormat = resolveBestExportFormatFromAnalyses(analyses, candidates);
    if (bestFormat) {
      exportFormatRecommendationCache.set(formatRecommendationCacheKey, bestFormat);
      return applyRecommendedFormat(bestFormat);
    }
  }

  const cached = exportFormatRecommendationCache.get(formatRecommendationCacheKey);
  if (typeof cached === 'string' && cached) {
    return applyRecommendedFormat(cached);
  }

  const token = state.exportWizardFormatRecommendationToken + 1;
  state.exportWizardFormatRecommendationToken = token;

  try {
    const results = await Promise.all(
      candidates.map(async (format) => {
        const recommendation = await estimateRecommendedCompressionResultForFormat(format, item, {
          format,
          preset,
          ...editOptions
        });
        const levelRecommendation = resolveExportRecommendationResultForLevel(recommendation);
        return {
          format,
          size: Math.max(1, Number(levelRecommendation?.size) || Number(recommendation?.size) || 1)
        };
      })
    );
    if (token !== state.exportWizardFormatRecommendationToken || !state.exportWizardItem?.dataUrl) {
      return state.exportWizardFormat;
    }

    const best = results.reduce((winner, entry) => {
      if (!winner) {
        return entry;
      }
      return entry.size < winner.size ? entry : winner;
    }, null);
    const bestFormat = normalizeExportFormat(best?.format || candidates[0]);
    exportFormatRecommendationCache.set(formatRecommendationCacheKey, bestFormat);
    return applyRecommendedFormat(bestFormat);
  } catch (error) {
    console.error('Format recommendation failed', error);
    if (token !== state.exportWizardFormatRecommendationToken) {
      return state.exportWizardFormat;
    }
    return state.exportWizardFormat;
  }
}

async function refreshExportFormatAndCompressionRecommendations(options = {}) {
  if (!state.exportWizardItem?.dataUrl) {
    setExportRecommendedCompression(null);
    setExportPassthroughCompression(null);
    updateExportOptimalSummaryUi();
    return;
  }

  const forceFormat = options.forceFormat === true;
  await ensureInitialExportWizardFormatAnalyses({ force: options.forceInitialAnalysis === true });
  const analyses = await ensureExportWizardFormatAnalyses({ force: options.forceAnalysis === true });
  updateExportOptimalSummaryUi(analyses);
  const previousFormat = state.exportWizardFormat;
  await autoSelectBestFormatForCurrentExport({ force: forceFormat, analyses });
  if (previousFormat !== state.exportWizardFormat) {
    state.exportWizardCompressionTouched = false;
    await refreshExportWizardPreview();
  }
  await applyRecommendedCompressionForCurrentExport({ force: true, apply: true, analyses });
}

async function buildExportRecommendationContext(item, options = {}, forcedFormat = '') {
  const dataUrl = String(item?.dataUrl || '').trim();
  if (!dataUrl.startsWith('data:image/')) {
    return {
      fallbackResult: {
        compression: EXPORT_DEFAULT_COMPRESSION,
        size: Math.max(1, Number(item?.fileSize) || getDataUrlByteLength(item?.dataUrl) || 1)
      }
    };
  }
  const preset = normalizeExportPreset(options.preset || state.exportWizardPreset);
  const editOptions = {
    cropEnabled: options.cropEnabled === true,
    cropRect: options.cropRect,
    transparentBackground: options.transparentBackground === true,
    backgroundTolerance: options.backgroundTolerance,
    backgroundProtection: options.backgroundProtection
  };
  const format = resolveEffectiveExportFormat(forcedFormat || options.format || state.exportWizardFormat, preset, editOptions);
  const originalSize = Math.max(0, Number(item?.fileSize) || getDataUrlByteLength(item?.dataUrl));
  const canPassthroughOriginal = isOriginalFormatPassthroughExport(item, format, preset, editOptions, 0);
  const targetSize = resolveExportPreviewTargetSize(item, preset, editOptions);

  const image = await loadImage(dataUrl);
  const canvas = buildExportCanvasFromImage(image, {
    format,
    targetSize,
    maxEdge: EXPORT_RECOMMENDATION_MAX_EDGE,
    ...editOptions
  });
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    return {
      fallbackResult: {
        compression: EXPORT_DEFAULT_COMPRESSION,
        size: Math.max(1, Number(item?.fileSize) || 1)
      }
    };
  }
  const width = Math.max(1, Number(canvas.width) || 1);
  const height = Math.max(1, Number(canvas.height) || 1);
  const referencePixels = ctx.getImageData(0, 0, width, height).data;
  const validationCanvas = buildExportCanvasFromImage(image, {
    format,
    targetSize: resolveExportValidationTargetSize(item, preset, editOptions),
    ...editOptions
  });
  const validationResults = new Map();
  const measureValidationResult = async (compression) => {
    const safeCompression = normalizeExportCompression(compression);
    if (validationResults.has(safeCompression)) {
      return validationResults.get(safeCompression);
    }
    const measured = await measureExportSizeForCompression(validationCanvas, format, preset, safeCompression);
    validationResults.set(safeCompression, measured);
    return measured;
  };
  const originalComparableFormat = extractComparableExportFormatFromDataUrl(item?.dataUrl);
  const shouldMatchOriginalSize = canPassthroughOriginal && originalComparableFormat === format && originalSize > 0;
  const cropCompressionFloors =
    options.ignoreInitialCompressionFloor === true || editOptions.cropEnabled !== true
      ? {}
      : getExportWizardInitialCompressionFloors(format);

  return {
    item,
    format,
    preset,
    editOptions,
    originalSize,
    canPassthroughOriginal,
    width,
    height,
    canvas,
    referencePixels,
    measureValidationResult,
    shouldMatchOriginalSize,
    cropCompressionFloors
  };
}

async function buildExportRecommendationResults(context, candidateCompressions) {
  const results = [];
  for (const candidate of candidateCompressions) {
    const compression = normalizeExportCompression(candidate);
    const quality = resolveExportQuality(context.format, context.preset, compression);
    const colorCount = context.format === 'image/png' ? resolvePngColorCountFromCompression(compression) : 0;
    const blob = await encodeCanvasToExportBlob(context.canvas, {
      format: context.format,
      quality,
      pngCompression: compression
    });
    const error =
      compression <= 0 && context.format === 'image/png'
        ? 0
        : await computeBlobVisualMse(blob, context.referencePixels, context.width, context.height);
    results.push({
      compression,
      colorCount,
      size: Math.max(1, Number(blob?.size) || 1),
      error
    });
  }
  return results;
}

async function finalizeExportRecommendationResult(context, results, candidateCompressions, recommendedCompression, options = {}) {
  const firstResult = results[0] || {
    compression: EXPORT_DEFAULT_COMPRESSION,
    size: Math.max(1, context.originalSize || 1)
  };
  const recommendedAnalysisResult =
    results.find((entry) => entry.compression === recommendedCompression) ||
    results.find((entry) => entry.compression === 0) ||
    firstResult;
  const analyzedCompression = normalizeExportCompression(recommendedAnalysisResult.compression);
  const recommendedResult = await context.measureValidationResult(recommendedAnalysisResult.compression);

  let finalRecommendation = null;
  if (context.shouldMatchOriginalSize && recommendedResult.size >= context.originalSize) {
    const comparableCandidates = candidateCompressions.filter((candidate) => normalizeExportCompression(candidate) > 0);
    const comparableResults = [];
    for (const candidate of comparableCandidates) {
      comparableResults.push(await context.measureValidationResult(candidate));
    }
    const equivalentResult = pickClosestExportResultBySize(comparableResults, context.originalSize) || {
      compression: 0,
      size: Math.max(1, context.originalSize)
    };
    finalRecommendation = {
      compression: normalizeExportCompression(equivalentResult.compression),
      size: Math.max(1, context.originalSize),
      analyzedCompression,
      forcedCompression: normalizeExportCompression(equivalentResult.compression),
      passthroughCompression: normalizeExportCompression(equivalentResult.compression)
    };
  } else {
    finalRecommendation = {
      compression: normalizeExportCompression(recommendedResult.compression),
      size: Math.max(1, Number(recommendedResult?.size) || 1),
      analyzedCompression,
      forcedCompression: null,
      passthroughCompression: null
    };
  }

  const cropCompressionFloor = Number.isFinite(Number(options.cropCompressionFloor))
    ? normalizeExportCompression(options.cropCompressionFloor)
    : null;
  if (
    cropCompressionFloor !== null &&
    normalizeExportCompression(finalRecommendation.compression) < cropCompressionFloor
  ) {
    const flooredResult = await context.measureValidationResult(cropCompressionFloor);
    finalRecommendation = {
      compression: cropCompressionFloor,
      size: Math.max(1, Number(flooredResult?.size) || 1),
      analyzedCompression: finalRecommendation.analyzedCompression,
      forcedCompression: finalRecommendation.forcedCompression,
      passthroughCompression: null
    };
  }

  return finalRecommendation;
}

async function finalizeExportRecommendationLevels(context, results, candidateCompressions, recommendedCompressions) {
  const levels = {};
  for (const [profile, compression] of Object.entries(recommendedCompressions || {})) {
    levels[profile] = await finalizeExportRecommendationResult(context, results, candidateCompressions, compression, {
      cropCompressionFloor: context.cropCompressionFloors?.[normalizeExportCompressionLevel(profile)]
    });
  }
  const conservative = normalizeExportRecommendationLevel(levels.conservative);
  const optimal = normalizeExportRecommendationLevel(levels.optimal, conservative) || conservative;
  const aggressive = normalizeExportRecommendationLevel(levels.aggressive, optimal) || optimal;
  return {
    ...conservative,
    levels: {
      conservative,
      optimal,
      aggressive
    }
  };
}

async function estimateRecommendedPngCompressionResultForItem(item, options = {}) {
  const context = await buildExportRecommendationContext(item, options, 'image/png');
  if (context?.fallbackResult) {
    return context.fallbackResult;
  }
  const pngAnalysis = analyzePngRecommendationCanvas(context.canvas);
  const pngStrategy = resolvePngRecommendationStrategy(pngAnalysis);
  const coarseCandidates = EXPORT_RECOMMENDATION_CANDIDATES.filter(
    (candidate) => normalizeExportCompression(candidate) <= (pngStrategy?.maxCompression ?? 100)
  );
  const coarseResults = await buildExportRecommendationResults(context, coarseCandidates);
  const { candidateCompressions, results, recommendedCompressions } = await refineRecommendationResults(
    context,
    'image/png',
    coarseCandidates,
    coarseResults,
    pngStrategy
  );
  return finalizeExportRecommendationLevels(context, results, candidateCompressions, recommendedCompressions);
}

async function estimateRecommendedWebpCompressionResultForItem(item, options = {}) {
  const context = await buildExportRecommendationContext(item, options, 'image/webp');
  if (context?.fallbackResult) {
    return context.fallbackResult;
  }
  const lossyAnalysis = analyzeRecommendationCanvas(context.canvas, { sampleDivisor: 192 });
  const lossyStrategy = resolveLossyRecommendationStrategy(lossyAnalysis, 'image/webp');
  const coarseCandidates = EXPORT_RECOMMENDATION_CANDIDATES.filter(
    (candidate) => normalizeExportCompression(candidate) <= (lossyStrategy?.maxCompression ?? 100)
  );
  const coarseResults = await buildExportRecommendationResults(context, coarseCandidates);
  const { candidateCompressions, results, recommendedCompressions } = await refineRecommendationResults(
    context,
    'image/webp',
    coarseCandidates,
    coarseResults,
    lossyStrategy
  );
  return finalizeExportRecommendationLevels(context, results, candidateCompressions, recommendedCompressions);
}

async function estimateRecommendedJpegCompressionResultForItem(item, options = {}) {
  const context = await buildExportRecommendationContext(item, options, 'image/jpeg');
  if (context?.fallbackResult) {
    return context.fallbackResult;
  }
  const lossyAnalysis = analyzeRecommendationCanvas(context.canvas, { sampleDivisor: 192 });
  const lossyStrategy = resolveLossyRecommendationStrategy(lossyAnalysis, 'image/jpeg');
  const coarseCandidates = EXPORT_RECOMMENDATION_CANDIDATES.filter(
    (candidate) => normalizeExportCompression(candidate) <= (lossyStrategy?.maxCompression ?? 100)
  );
  const coarseResults = await buildExportRecommendationResults(context, coarseCandidates);
  const { candidateCompressions, results, recommendedCompressions } = await refineRecommendationResults(
    context,
    'image/jpeg',
    coarseCandidates,
    coarseResults,
    lossyStrategy
  );
  return finalizeExportRecommendationLevels(context, results, candidateCompressions, recommendedCompressions);
}

async function estimateRecommendedCompressionResultForFormat(format, item, options = {}) {
  const safeFormat = normalizeExportFormat(format);
  if (safeFormat === 'image/png') {
    return estimateRecommendedPngCompressionResultForItem(item, options);
  }
  if (safeFormat === 'image/jpeg') {
    return estimateRecommendedJpegCompressionResultForItem(item, options);
  }
  return estimateRecommendedWebpCompressionResultForItem(item, options);
}

async function estimateRecommendedCompressionResultForItem(item, options = {}) {
  const preset = normalizeExportPreset(options.preset || state.exportWizardPreset);
  const editOptions = {
    cropEnabled: options.cropEnabled === true,
    cropRect: options.cropRect,
    transparentBackground: options.transparentBackground === true,
    backgroundTolerance: options.backgroundTolerance,
    backgroundProtection: options.backgroundProtection
  };
  const format = resolveEffectiveExportFormat(options.format || state.exportWizardFormat, preset, editOptions);
  return estimateRecommendedCompressionResultForFormat(format, item, options);
}

async function estimateRecommendedCompressionForItem(item, options = {}) {
  const result = await estimateRecommendedCompressionResultForItem(item, options);
  return result.compression;
}

function computeAverageLumaBlockSsim(referenceLuma, comparedLuma, width, height) {
  const safeWidth = Math.max(1, Number(width) || 1);
  const safeHeight = Math.max(1, Number(height) || 1);
  const blockSize = safeWidth >= 128 || safeHeight >= 128 ? 8 : 6;
  const c1 = Math.pow(0.01 * 255, 2);
  const c2 = Math.pow(0.03 * 255, 2);
  let blockCount = 0;
  let ssimSum = 0;

  for (let top = 0; top < safeHeight; top += blockSize) {
    for (let left = 0; left < safeWidth; left += blockSize) {
      const right = Math.min(safeWidth, left + blockSize);
      const bottom = Math.min(safeHeight, top + blockSize);
      const sampleCount = Math.max(1, (right - left) * (bottom - top));
      let meanReference = 0;
      let meanCompared = 0;

      for (let y = top; y < bottom; y += 1) {
        let index = (y * safeWidth) + left;
        for (let x = left; x < right; x += 1, index += 1) {
          meanReference += referenceLuma[index];
          meanCompared += comparedLuma[index];
        }
      }

      meanReference /= sampleCount;
      meanCompared /= sampleCount;

      let varianceReference = 0;
      let varianceCompared = 0;
      let covariance = 0;
      for (let y = top; y < bottom; y += 1) {
        let index = (y * safeWidth) + left;
        for (let x = left; x < right; x += 1, index += 1) {
          const refDelta = referenceLuma[index] - meanReference;
          const cmpDelta = comparedLuma[index] - meanCompared;
          varianceReference += refDelta * refDelta;
          varianceCompared += cmpDelta * cmpDelta;
          covariance += refDelta * cmpDelta;
        }
      }

      const safeDivisor = Math.max(1, sampleCount - 1);
      varianceReference /= safeDivisor;
      varianceCompared /= safeDivisor;
      covariance /= safeDivisor;

      const numerator = ((2 * meanReference * meanCompared) + c1) * ((2 * covariance) + c2);
      const denominator = (((meanReference * meanReference) + (meanCompared * meanCompared) + c1) * (varianceReference + varianceCompared + c2));
      const ssim = denominator > 0 ? numerator / denominator : 1;
      ssimSum += Math.max(0, Math.min(1, ssim));
      blockCount += 1;
    }
  }

  return blockCount > 0 ? ssimSum / blockCount : 1;
}

function computePerceptualPixelError(referencePixels, comparedPixels, width, height) {
  if (!(referencePixels instanceof Uint8ClampedArray) || !(comparedPixels instanceof Uint8ClampedArray)) {
    return Number.POSITIVE_INFINITY;
  }
  const safeWidth = Math.max(1, Number(width) || 1);
  const safeHeight = Math.max(1, Number(height) || 1);
  const pixelCount = Math.max(1, safeWidth * safeHeight);
  const referenceLuma = new Float32Array(pixelCount);
  const comparedLuma = new Float32Array(pixelCount);
  let lumaSquaredError = 0;
  let chromaSquaredError = 0;
  let alphaSquaredError = 0;

  for (let pixelIndex = 0, dataIndex = 0; pixelIndex < pixelCount; pixelIndex += 1, dataIndex += 4) {
    const refR = referencePixels[dataIndex];
    const refG = referencePixels[dataIndex + 1];
    const refB = referencePixels[dataIndex + 2];
    const refA = referencePixels[dataIndex + 3];
    const cmpR = comparedPixels[dataIndex];
    const cmpG = comparedPixels[dataIndex + 1];
    const cmpB = comparedPixels[dataIndex + 2];
    const cmpA = comparedPixels[dataIndex + 3];

    const refY = (0.299 * refR) + (0.587 * refG) + (0.114 * refB);
    const cmpY = (0.299 * cmpR) + (0.587 * cmpG) + (0.114 * cmpB);
    referenceLuma[pixelIndex] = refY;
    comparedLuma[pixelIndex] = cmpY;
    const deltaY = cmpY - refY;
    lumaSquaredError += deltaY * deltaY;

    const refCb = (-0.168736 * refR) - (0.331264 * refG) + (0.5 * refB);
    const refCr = (0.5 * refR) - (0.418688 * refG) - (0.081312 * refB);
    const cmpCb = (-0.168736 * cmpR) - (0.331264 * cmpG) + (0.5 * cmpB);
    const cmpCr = (0.5 * cmpR) - (0.418688 * cmpG) - (0.081312 * cmpB);
    const deltaCb = cmpCb - refCb;
    const deltaCr = cmpCr - refCr;
    chromaSquaredError += (deltaCb * deltaCb) + (deltaCr * deltaCr);

    const deltaA = cmpA - refA;
    alphaSquaredError += deltaA * deltaA;
  }

  let gradientSquaredError = 0;
  let gradientSamples = 0;
  let strongEdgeSquaredError = 0;
  let strongEdgeSamples = 0;
  for (let y = 0; y < safeHeight - 1; y += 1) {
    for (let x = 0; x < safeWidth - 1; x += 1) {
      const index = (y * safeWidth) + x;
      const refDx = referenceLuma[index + 1] - referenceLuma[index];
      const cmpDx = comparedLuma[index + 1] - comparedLuma[index];
      const refDy = referenceLuma[index + safeWidth] - referenceLuma[index];
      const cmpDy = comparedLuma[index + safeWidth] - comparedLuma[index];
      const deltaDx = cmpDx - refDx;
      const deltaDy = cmpDy - refDy;
      const pairSquaredError = (deltaDx * deltaDx) + (deltaDy * deltaDy);
      gradientSquaredError += pairSquaredError;
      gradientSamples += 2;
      const referenceEdgeStrength = Math.abs(refDx) + Math.abs(refDy);
      if (referenceEdgeStrength >= 28) {
        const edgeWeight = 1 + Math.min(1.5, referenceEdgeStrength / 96);
        strongEdgeSquaredError += pairSquaredError * edgeWeight;
        strongEdgeSamples += 2 * edgeWeight;
      }
    }
  }

  const meanSsim = computeAverageLumaBlockSsim(referenceLuma, comparedLuma, safeWidth, safeHeight);
  const ssimPenalty = Math.max(0, (1 - meanSsim) * 100);
  const lumaRmse = Math.sqrt(lumaSquaredError / pixelCount);
  const chromaRmse = Math.sqrt(chromaSquaredError / (pixelCount * 2));
  const alphaRmse = Math.sqrt(alphaSquaredError / pixelCount);
  const gradientRmse = gradientSamples > 0 ? Math.sqrt(gradientSquaredError / gradientSamples) : 0;
  const strongEdgeRmse = strongEdgeSamples > 0 ? Math.sqrt(strongEdgeSquaredError / strongEdgeSamples) : 0;

  return (
    (ssimPenalty * 1.35) +
    (lumaRmse / 4.5) +
    (gradientRmse / 5.5) +
    (strongEdgeRmse / 3.3) +
    (chromaRmse / 9) +
    (alphaRmse / 8)
  );
}

async function computeBlobVisualMse(blob, referencePixels, width, height) {
  if (!(blob instanceof Blob) || !(referencePixels instanceof Uint8ClampedArray)) {
    return Number.POSITIVE_INFINITY;
  }
  const safeWidth = Math.max(1, Number(width) || 1);
  const safeHeight = Math.max(1, Number(height) || 1);
  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImage(objectUrl);
    const canvas = document.createElement('canvas');
    canvas.width = safeWidth;
    canvas.height = safeHeight;
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      return Number.POSITIVE_INFINITY;
    }
    ctx.clearRect(0, 0, safeWidth, safeHeight);
    ctx.drawImage(image, 0, 0, safeWidth, safeHeight);
    const comparedPixels = ctx.getImageData(0, 0, safeWidth, safeHeight).data;
    return computePerceptualPixelError(referencePixels, comparedPixels, safeWidth, safeHeight);
  } catch (error) {
    console.error('Blob visual metric failed', error);
    return Number.POSITIVE_INFINITY;
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function syncExportFormatControlState() {
  if (!(elements.exportFormatSelect instanceof HTMLSelectElement)) {
    return;
  }
  const isFaviconPreset = normalizeExportPreset(state.exportWizardPreset) === EXPORT_PRESET_FAVICON;
  const needsTransparency = state.exportWizardTransparentBackground === true;
  for (const option of Array.from(elements.exportFormatSelect.options)) {
    if (String(option.value || '').trim() === 'image/jpeg') {
      option.disabled = needsTransparency;
    }
  }
  if (isFaviconPreset) {
    state.exportWizardFormat = 'image/png';
    elements.exportFormatSelect.value = state.exportWizardFormat;
    elements.exportFormatSelect.disabled = true;
    elements.exportFormatSelect.title = 'En preset favicon el format sempre es PNG';
    updateExportWizardActionUi();
    return;
  }
  if (needsTransparency && normalizeExportFormat(state.exportWizardFormat) === 'image/jpeg') {
    state.exportWizardFormat = 'image/png';
  }
  elements.exportFormatSelect.value = state.exportWizardFormat;
  elements.exportFormatSelect.disabled = false;
  if (needsTransparency) {
    elements.exportFormatSelect.title = 'JPG no admet transparència; s\'utilitza PNG o WebP';
  } else {
    elements.exportFormatSelect.removeAttribute('title');
  }
  updateExportWizardActionUi();
}

function resolveExportFormatForPreset(format, preset) {
  const safePreset = normalizeExportPreset(preset);
  if (safePreset === EXPORT_PRESET_FAVICON) {
    return 'image/png';
  }
  return normalizeExportFormat(format);
}

function resolveEffectiveExportFormat(format, preset, editOptions = {}) {
  const safeFormat = resolveExportFormatForPreset(format, preset);
  if (editOptions.transparentBackground === true && safeFormat === 'image/jpeg') {
    return 'image/png';
  }
  return safeFormat;
}

function resolveExportQuality(format, preset, compression) {
  const safeFormat = resolveExportFormatForPreset(format, preset);
  if (safeFormat !== 'image/jpeg' && safeFormat !== 'image/webp') {
    return undefined;
  }
  return resolveLossyQualityFromCompression(compression);
}

function resolveExportSourceDimensions(item, editOptions = {}) {
  const sourceWidth = Math.max(0, Number(item?.width) || 0);
  const sourceHeight = Math.max(0, Number(item?.height) || 0);
  if (sourceWidth <= 0 || sourceHeight <= 0) {
    return { width: 0, height: 0 };
  }

  const cropEnabled = editOptions.cropEnabled === true;
  const cropRect = cropEnabled ? clampExportCropRect(editOptions.cropRect) : createFullExportCropRect();
  const cropLeft = Math.max(0, Math.min(sourceWidth - 1, Math.round(cropRect.x * sourceWidth)));
  const cropTop = Math.max(0, Math.min(sourceHeight - 1, Math.round(cropRect.y * sourceHeight)));
  const cropRight = Math.max(
    cropLeft + 1,
    Math.min(sourceWidth, Math.round((cropRect.x + cropRect.width) * sourceWidth))
  );
  const cropBottom = Math.max(
    cropTop + 1,
    Math.min(sourceHeight, Math.round((cropRect.y + cropRect.height) * sourceHeight))
  );

  return {
    width: Math.max(1, cropRight - cropLeft),
    height: Math.max(1, cropBottom - cropTop)
  };
}

function resolveFaviconMaxSourceSize(item, editOptions = {}) {
  const { width, height } = resolveExportSourceDimensions(item, editOptions);
  return Math.max(0, Math.min(width, height));
}

function resolveFaviconExportSizes(item, editOptions = {}) {
  const maxSourceSize = resolveFaviconMaxSourceSize(item, editOptions);
  if (maxSourceSize <= 0) {
    return [];
  }
  const sizes = FAVICON_PRESET_SIZES.filter((size) => Number(size) > 0 && Number(size) <= maxSourceSize);
  const roundedOriginal = Math.max(1, Math.floor(maxSourceSize));
  if (!sizes.includes(roundedOriginal)) {
    sizes.push(roundedOriginal);
  }
  return Array.from(new Set(sizes)).sort((left, right) => left - right);
}

function openExportWizardFromImageId(imageId, options = {}) {
  const item = findExportSourceItem(imageId, options);
  if (!item?.dataUrl) {
    showFeedbackToast('No s\'ha pogut obrir export wizard', 'error');
    return;
  }
  const itemName = sanitizeDisplayName(item.name || extractFileNameFromItemSource(item) || 'Imatge', 'Imatge');

  state.exportWizardItem = {
    id: String(item.id || '').trim(),
    name: itemName,
    dataUrl: String(item.dataUrl || ''),
    sourceUrl: String(item.sourceUrl || '').trim(),
    sourcePath: String(item.sourcePath || '').trim(),
    fileSize: Math.max(0, Number(item.fileSize) || 0),
    width: Math.max(0, Number(item.width) || 0),
    height: Math.max(0, Number(item.height) || 0)
  };
  state.exportWizardFormat = EXPORT_DEFAULT_FORMAT;
  state.exportWizardPreset = EXPORT_PRESET_ORIGINAL;
  state.exportWizardCompressionLevel = EXPORT_DEFAULT_COMPRESSION_LEVEL;
  state.exportWizardCompression = EXPORT_DEFAULT_COMPRESSION;
  state.exportWizardFormatTouched = false;
  state.exportWizardCompressionTouched = false;
  state.exportWizardPassthroughCompression = null;
  state.exportWizardCropEnabled = false;
  state.exportWizardCropRect = createDefaultExportCropRect(state.exportWizardItem);
  state.exportWizardCropInteraction = null;
  state.exportWizardTransparentBackground = false;
  state.exportWizardBackgroundTolerance = EXPORT_BACKGROUND_TOLERANCE_DEFAULT;
  state.exportWizardBackgroundProtection = EXPORT_BACKGROUND_PROTECTION_DEFAULT;
  state.exportWizardTransparentAdjusting = false;
  state.exportWizardTransparentAdjustingControl = '';
  state.exportWizardTransparentPointerId = null;
  state.exportWizardEstimatedOutputBytes = 0;
  state.exportWizardFormatRecommendationToken += 1;
  state.exportWizardFormatAnalysisToken += 1;
  state.exportWizardFormatAnalysesKey = '';
  state.exportWizardFormatAnalyses = {};
  state.exportWizardInitialFormatAnalysisToken += 1;
  state.exportWizardInitialFormatAnalysesKey = '';
  state.exportWizardInitialFormatAnalyses = {};
  state.exportWizardRecommendationToken += 1;
  state.exportWizardBusy = false;

  if (elements.exportFormatSelect instanceof HTMLSelectElement) {
    elements.exportFormatSelect.value = state.exportWizardFormat;
  }
  if (elements.exportPresetSelect instanceof HTMLSelectElement) {
    elements.exportPresetSelect.value = state.exportWizardPreset;
  }
  updateExportCompressionLevelUi();
  syncExportFormatControlState();
  updateExportCropUi();
  updateExportTransparentBackgroundUi();
  setExportRecommendedCompression(null);
  updateExportOptimalSummaryUi();
  updateExportCompressionUi();
  updateExportWizardActionUi();
  if (elements.exportWizardSourceName instanceof HTMLElement) {
    elements.exportWizardSourceName.textContent = itemName;
  }
  if (elements.exportWizardNameInput instanceof HTMLInputElement) {
    const suggested = sanitizeArchiveName(removeFileExtension(getExportSourceName(state.exportWizardItem))) || 'export';
    elements.exportWizardNameInput.value = suggested;
  }
  if (elements.exportWizardOriginalMeta instanceof HTMLElement) {
    const fileSize = state.exportWizardItem.fileSize || getDataUrlByteLength(state.exportWizardItem.dataUrl);
    const dims =
      state.exportWizardItem.width > 0 && state.exportWizardItem.height > 0
        ? `${state.exportWizardItem.width}x${state.exportWizardItem.height}`
        : '?x?';
    const originalFormatLabel = formatOriginalImageFormatLabel(state.exportWizardItem.dataUrl);
    elements.exportWizardOriginalMeta.textContent = `Original · ${originalFormatLabel} · ${dims} px · ${formatBytes(fileSize)}`;
  }
  if (elements.exportWizardPreviewMeta instanceof HTMLElement) {
    elements.exportWizardPreviewMeta.textContent = 'Previsualitzant...';
  }
  if (elements.exportWizardPreviewImg instanceof HTMLImageElement) {
    elements.exportWizardPreviewImg.removeAttribute('src');
  }
  setExportWizardCropEditorPreviewUrl('');
  setExportWizardPreviewLoading(false);

  closeSaveAsModal();
  closeManualUploadModal();
  if (!(elements.exportWizardModal instanceof HTMLElement)) {
    return;
  }
  elements.exportWizardModal.classList.remove('hidden');
  elements.exportWizardModal.setAttribute('aria-hidden', 'false');
  if (elements.exportWizardNameInput instanceof HTMLInputElement) {
    window.requestAnimationFrame(() => {
      elements.exportWizardNameInput.focus();
      elements.exportWizardNameInput.select();
      updateExportCropUi();
    });
  }
  void refreshExportWizardPreview();
  void refreshExportFormatAndCompressionRecommendations({ forceFormat: true });
}

function findExportSourceItem(imageId, options = {}) {
  const safeId = String(imageId || '').trim();
  if (!safeId) {
    return null;
  }

  const galleryId = String(options.galleryId || '').trim();
  if (galleryId) {
    if (galleryId === MANUAL_GALLERY_ID) {
      return state.manualItems.find((item) => String(item?.id || '').trim() === safeId) || null;
    }
    const gallery = state.savedGalleries.find((entry) => String(entry?.id || '').trim() === galleryId);
    if (!gallery || !Array.isArray(gallery.items)) {
      return null;
    }
    return gallery.items.find((item) => String(item?.id || '').trim() === safeId) || null;
  }

  if (options.manual) {
    return state.manualItems.find((item) => String(item?.id || '').trim() === safeId) || null;
  }

  const fromWeb = state.webItems.find((item) => String(item?.id || '').trim() === safeId);
  if (fromWeb) {
    return fromWeb;
  }

  for (const gallery of state.savedGalleries) {
    if (!Array.isArray(gallery.items)) {
      continue;
    }
    const found = gallery.items.find((item) => String(item?.id || '').trim() === safeId);
    if (found) {
      return found;
    }
  }
  return null;
}

function closeExportWizardModal() {
  if (elements.exportWizardModal instanceof HTMLElement) {
    elements.exportWizardModal.classList.add('hidden');
    elements.exportWizardModal.setAttribute('aria-hidden', 'true');
  }
  if (state.exportWizardPreviewUrl) {
    URL.revokeObjectURL(state.exportWizardPreviewUrl);
  }
  state.exportWizardPreviewUrl = '';
  setExportWizardCropEditorPreviewUrl('');
  state.exportWizardItem = null;
  state.exportWizardPreviewToken += 1;
  state.exportWizardFormatTouched = false;
  state.exportWizardFormatRecommendationToken += 1;
  state.exportWizardFormatAnalysisToken += 1;
  state.exportWizardFormatAnalysesKey = '';
  state.exportWizardFormatAnalyses = {};
  state.exportWizardInitialFormatAnalysisToken += 1;
  state.exportWizardInitialFormatAnalysesKey = '';
  state.exportWizardInitialFormatAnalyses = {};
  state.exportWizardRecommendationToken += 1;
  state.exportWizardCompressionLevel = EXPORT_DEFAULT_COMPRESSION_LEVEL;
  state.exportWizardCompressionTouched = false;
  state.exportWizardPassthroughCompression = null;
  state.exportWizardCropEnabled = false;
  state.exportWizardCropRect = createFullExportCropRect();
  state.exportWizardCropInteraction = null;
  state.exportWizardTransparentBackground = false;
  state.exportWizardBackgroundTolerance = EXPORT_BACKGROUND_TOLERANCE_DEFAULT;
  state.exportWizardBackgroundProtection = EXPORT_BACKGROUND_PROTECTION_DEFAULT;
  state.exportWizardTransparentAdjusting = false;
  state.exportWizardTransparentAdjustingControl = '';
  state.exportWizardTransparentPointerId = null;
  state.exportWizardEstimatedOutputBytes = 0;
  state.exportWizardBusy = false;
  setExportWizardPreviewLoading(false);
  setExportRecommendedCompression(null);
  setExportPassthroughCompression(null);
  updateExportCompressionLevelUi();
  updateExportOptimalSummaryUi();
  updateExportCropUi();
  updateExportTransparentBackgroundUi();
  updateExportWizardActionUi();
  if (elements.exportWizardPreviewImg instanceof HTMLImageElement) {
    elements.exportWizardPreviewImg.removeAttribute('src');
  }
  if (elements.exportCropEditorImg instanceof HTMLImageElement) {
    elements.exportCropEditorImg.removeAttribute('src');
    elements.exportCropEditorImg.removeAttribute('style');
  }
  if (elements.exportCropBox instanceof HTMLElement) {
    elements.exportCropBox.removeAttribute('style');
  }
}

async function refreshExportWizardPreview() {
  if (!state.exportWizardItem?.dataUrl) {
    return;
  }

  const token = state.exportWizardPreviewToken + 1;
  state.exportWizardPreviewToken = token;
  const passthroughCompression = state.exportWizardPassthroughCompression;

  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  const previewEditOptions = getCurrentExportPreviewEditOptions(preset);
  const format = resolveEffectiveExportFormat(state.exportWizardFormat, preset, editOptions);
  const exportPreviewSize = resolveExportPreviewTargetSize(state.exportWizardItem, preset, editOptions);
  const faviconMaxSize =
    preset === EXPORT_PRESET_FAVICON ? resolveFaviconMaxSourceSize(state.exportWizardItem, editOptions) : 0;
  const imagePreviewSize = preset === EXPORT_PRESET_FAVICON ? 0 : exportPreviewSize;
  const quality = resolveExportQuality(format, preset, state.exportWizardCompression);
  const needsTransparentCropStagePreview =
    editOptions.cropEnabled === true &&
    editOptions.transparentBackground === true &&
    format !== 'image/jpeg';
  const hasStableTransparentCropStagePreview =
    needsTransparentCropStagePreview &&
    String(state.exportWizardCropEditorPreviewUrl || '').trim() !== '';

  if (hasStableTransparentCropStagePreview) {
    setExportWizardPreviewLoading(false);
  } else {
    setExportWizardPreviewLoading(true, {
      delayMs: state.exportWizardPreviewUrl ? 220 : 120,
      token
    });
  }
  if (
    elements.exportWizardPreviewMeta instanceof HTMLElement &&
    !state.exportWizardPreviewUrl &&
    !hasStableTransparentCropStagePreview
  ) {
    elements.exportWizardPreviewMeta.textContent = 'Previsualitzant...';
  }

  if (!needsTransparentCropStagePreview && state.exportWizardCropEditorPreviewUrl) {
    setExportWizardCropEditorPreviewUrl('');
    updateExportCropUi();
  }

  try {
    const previewAsset = await renderImageExportAsset(state.exportWizardItem.dataUrl, {
      format,
      preset,
      targetSize: imagePreviewSize,
      quality,
      pngCompression: state.exportWizardCompression,
      passthroughCompression,
      ...previewEditOptions
    });
    if (token !== state.exportWizardPreviewToken) {
      return;
    }

    if (needsTransparentCropStagePreview) {
      const cropStageAsset = await renderImageExportAsset(state.exportWizardItem.dataUrl, {
        format,
        preset,
        targetSize: imagePreviewSize,
        quality,
        pngCompression: state.exportWizardCompression,
        passthroughCompression,
        ...editOptions,
        cropEnabled: false,
        cropRect: createFullExportCropRect()
      });
      if (token !== state.exportWizardPreviewToken) {
        return;
      }
      setExportWizardCropEditorPreviewUrl(URL.createObjectURL(cropStageAsset.blob));
      updateExportCropUi();
    }

    const metaAsset =
      preset === EXPORT_PRESET_FAVICON
        ? await renderImageExportAsset(state.exportWizardItem.dataUrl, {
            format,
            preset,
            targetSize: exportPreviewSize,
            quality,
            pngCompression: state.exportWizardCompression,
            passthroughCompression,
            ...editOptions
          })
        : previewAsset;
    if (token !== state.exportWizardPreviewToken) {
      return;
    }

    const actionAsset =
      preset === EXPORT_PRESET_FAVICON && faviconMaxSize > 0 && faviconMaxSize !== exportPreviewSize
        ? await renderImageExportAsset(state.exportWizardItem.dataUrl, {
            format,
            preset,
            targetSize: faviconMaxSize,
            quality,
            pngCompression: state.exportWizardCompression,
            passthroughCompression,
            ...editOptions
          })
        : metaAsset;
    if (token !== state.exportWizardPreviewToken) {
      return;
    }

    const nextUrl = URL.createObjectURL(previewAsset.blob);
    if (state.exportWizardPreviewUrl) {
      URL.revokeObjectURL(state.exportWizardPreviewUrl);
    }
    state.exportWizardPreviewUrl = nextUrl;
    state.exportWizardEstimatedOutputBytes = Math.max(0, Number(actionAsset.blob.size) || 0);
    updateExportWizardActionUi();

    if (elements.exportWizardPreviewImg instanceof HTMLImageElement) {
      elements.exportWizardPreviewImg.src = nextUrl;
    }
    if (elements.exportWizardPreviewMeta instanceof HTMLElement) {
      const formatLabel = exportFormatLabel(format);
      const pngColorCount = resolvePngColorCountFromCompression(state.exportWizardCompression);
      const isPassthroughPreview = isOriginalFormatPassthroughExport(
        state.exportWizardItem,
        format,
        preset,
        editOptions,
        state.exportWizardCompression,
        passthroughCompression
      );
      const compressionLabel =
        isPassthroughPreview
          ? Number.isFinite(Number(passthroughCompression)) && passthroughCompression > 0
            ? `Original sense recompressio · Equivalencia ${passthroughCompression}%`
            : 'Original sense recompressio'
          : preset === EXPORT_PRESET_FAVICON
          ? normalizeExportCompression(state.exportWizardCompression) <= 0
            ? 'PNG sense pèrdua'
            : `PNG ${normalizeExportCompression(state.exportWizardCompression)}% (${pngColorCount} colors)`
          : format === 'image/png'
            ? normalizeExportCompression(state.exportWizardCompression) <= 0
              ? 'PNG sense pèrdua'
              : `PNG quantitzat ${normalizeExportCompression(state.exportWizardCompression)}% (${pngColorCount} colors)`
            : `Compressio ${normalizeExportCompression(state.exportWizardCompression)}% · Qualitat ${Math.round((quality || 0) * 100)}%`;
      const originalWidth = Math.max(0, Number(state.exportWizardItem?.width) || 0);
      const originalHeight = Math.max(0, Number(state.exportWizardItem?.height) || 0);
      const hasOriginalDims = originalWidth > 0 && originalHeight > 0;
      const exportDimsLabel =
        preset === EXPORT_PRESET_FAVICON
          ? faviconMaxSize > 0
            ? ` · fins a ${faviconMaxSize}x${faviconMaxSize} px`
            : ''
          : hasOriginalDims && (metaAsset.width !== originalWidth || metaAsset.height !== originalHeight)
            ? ` · ${metaAsset.width}x${metaAsset.height} px`
          : '';
      elements.exportWizardPreviewMeta.textContent =
        `Export · ${formatLabel}${exportDimsLabel} · ${formatBytes(metaAsset.blob.size)} · ${compressionLabel}`;
    }
    if (token === state.exportWizardPreviewToken) {
      setExportWizardPreviewLoading(false);
    }
  } catch (error) {
    if (token !== state.exportWizardPreviewToken) {
      return;
    }
    console.error('Export preview failed', error);
    state.exportWizardEstimatedOutputBytes = 0;
    updateExportWizardActionUi();
    if (elements.exportWizardPreviewMeta instanceof HTMLElement) {
      elements.exportWizardPreviewMeta.textContent = 'No s\'ha pogut generar la previsualitzacio';
    }
    setExportWizardPreviewLoading(false);
  }
}

async function exportWizardDownload() {
  if (state.exportWizardBusy || !state.exportWizardItem?.dataUrl) {
    return;
  }
  const baseName = getExportWizardBaseName();
  if (!baseName) {
    showFeedbackToast('Escriu un nom valid per exportar', 'error');
    if (elements.exportWizardNameInput instanceof HTMLInputElement) {
      elements.exportWizardNameInput.focus();
      elements.exportWizardNameInput.select();
    }
    return;
  }

  state.exportWizardBusy = true;
  updateExportWizardActionUi();

  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  const format = resolveEffectiveExportFormat(state.exportWizardFormat, preset, editOptions);
  const quality = resolveExportQuality(format, preset, state.exportWizardCompression);
  const extension = exportExtensionFromFormat(format);
  const passthroughCompression = state.exportWizardPassthroughCompression;

  try {
    if (preset === EXPORT_PRESET_FAVICON) {
      const faviconSizes = resolveFaviconExportSizes(state.exportWizardItem, editOptions);
      if (faviconSizes.length === 0) {
        showFeedbackToast('No s\'ha pogut calcular mides favicon valides', 'error');
        return;
      }
      const entries = [];
      for (const size of faviconSizes) {
        const asset = await renderImageExportAsset(state.exportWizardItem.dataUrl, {
          format,
          preset,
          targetSize: Number(size) || 0,
          quality,
          pngCompression: state.exportWizardCompression,
          passthroughCompression,
          ...editOptions
        });
        const bytes = new Uint8Array(await asset.blob.arrayBuffer());
        entries.push({
          name: `${baseName}-${size}.${extension}`,
          bytes
        });
      }
      const zipBlob = createZipBlob(entries, {
        compressionPercent: state.exportWizardCompression
      });
      await downloadBlobAsset(zipBlob, `${baseName}-favicon-pack.zip`);
      showFeedbackToast(`Export ZIP favicon (${entries.length} mides)`);
      return;
    }

    const asset = await renderImageExportAsset(state.exportWizardItem.dataUrl, {
      format,
      preset,
      targetSize: 0,
      quality,
      pngCompression: state.exportWizardCompression,
      passthroughCompression,
      ...editOptions
    });
    await downloadBlobAsset(asset.blob, `${baseName}.${extension}`);
    showFeedbackToast('Imatge exportada');
  } catch (error) {
    console.error('Export wizard failed', error);
    showFeedbackToast('No s\'ha pogut exportar la imatge', 'error');
  } finally {
    state.exportWizardBusy = false;
    updateExportWizardActionUi();
  }
}

async function exportWizardDownloadOriginalMatchedSize() {
  if (state.exportWizardBusy || !state.exportWizardItem?.dataUrl) {
    return;
  }
  const baseName = getExportWizardBaseName();
  if (!baseName) {
    showFeedbackToast('Escriu un nom valid per exportar', 'error');
    if (elements.exportWizardNameInput instanceof HTMLInputElement) {
      elements.exportWizardNameInput.focus();
      elements.exportWizardNameInput.select();
    }
    return;
  }

  const exportConfig = resolveOriginalMatchedSizeExportConfig();
  if (!exportConfig.available) {
    showFeedbackToast(exportConfig.reason || 'No s\'ha pogut exportar en format original', 'error');
    return;
  }

  state.exportWizardBusy = true;
  updateExportWizardActionUi();

  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();

  try {
    const plan = await resolveOriginalMatchedSizeExportPlan();
    if (!plan) {
      throw new Error('original_matched_export_plan_unavailable');
    }
    const asset = await renderImageExportAsset(state.exportWizardItem.dataUrl, {
      format: plan.format,
      preset,
      targetSize: 0,
      quality: resolveExportQuality(plan.format, preset, plan.compression),
      pngCompression: plan.compression,
      ...editOptions
    });
    await downloadBlobAsset(asset.blob, `${baseName}.${exportConfig.extension}`);
  } catch (error) {
    console.error('Original matched size export failed', error);
    showFeedbackToast('No s\'ha pogut exportar en format original', 'error');
  } finally {
    state.exportWizardBusy = false;
    updateExportWizardActionUi();
  }
}

function getExportSourceName(item) {
  const sourceName = extractFileNameFromItemSource(item);
  if (sourceName) {
    return sourceName;
  }
  return sanitizeDisplayName(item?.name || 'imatge', 'imatge');
}

function getExportWizardBaseName() {
  if (!(elements.exportWizardNameInput instanceof HTMLInputElement)) {
    return '';
  }
  const raw = String(elements.exportWizardNameInput.value || '').trim();
  return sanitizeArchiveName(removeFileExtension(raw));
}

async function renderImageExportAsset(dataUrl, options = {}) {
  const source = String(dataUrl || '').trim();
  if (!source.startsWith('data:image/')) {
    throw new Error('invalid_export_source');
  }

  const format = normalizeExportFormat(options.format || EXPORT_DEFAULT_FORMAT);
  const preset = normalizeExportPreset(options.preset || state.exportWizardPreset);
  const targetSize = Math.max(0, Math.round(Number(options.targetSize) || 0));
  const passthroughEditOptions = {
    cropEnabled: options.cropEnabled === true,
    cropRect: options.cropRect,
    transparentBackground: options.transparentBackground === true,
    backgroundTolerance: options.backgroundTolerance,
    backgroundProtection: options.backgroundProtection
  };
  if (isOriginalFormatPassthroughExport(
    { dataUrl: source },
    format,
    preset,
    passthroughEditOptions,
    options.pngCompression,
    options.passthroughCompression
  ) && targetSize <= 0) {
    const bytes = decodeDataUrlToBytes(source);
    if (bytes?.length) {
      const image = await loadImage(source);
      return {
        blob: new Blob([bytes], { type: extractMimeTypeFromDataUrl(source) || format }),
        width: Math.max(1, Number(image?.naturalWidth) || 1),
        height: Math.max(1, Number(image?.naturalHeight) || 1)
      };
    }
  }
  const image = await loadImage(source);
  const canvas = buildExportCanvasFromImage(image, {
    format,
    targetSize,
    cropEnabled: options.cropEnabled === true,
    cropRect: options.cropRect,
    transparentBackground: options.transparentBackground === true,
    backgroundTolerance: options.backgroundTolerance,
    backgroundProtection: options.backgroundProtection
  });
  const blob = await encodeCanvasToExportBlob(canvas, {
    format,
    quality: options.quality,
    pngCompression: options.pngCompression
  });
  return { blob, width: canvas.width, height: canvas.height };
}

function canvasToBlobAsync(canvas, format, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (blob) {
          resolve(blob);
          return;
        }
        reject(new Error('canvas_blob_failed'));
      },
      format,
      quality
    );
  });
}

function getCanvasImageData(canvas) {
  const width = Math.max(1, Number(canvas?.width) || 1);
  const height = Math.max(1, Number(canvas?.height) || 1);
  const ctx = canvas?.getContext?.('2d');
  if (!ctx) {
    throw new Error('canvas_context_unavailable');
  }
  return ctx.getImageData(0, 0, width, height);
}

function resolveEncodeQualityPercent(quality) {
  const safeQuality = Number.isFinite(Number(quality))
    ? Math.min(1, Math.max(0, Number(quality)))
    : resolveLossyQualityFromCompression(EXPORT_DEFAULT_COMPRESSION);
  return Math.max(1, Math.min(100, Math.round(safeQuality * 100)));
}

function resolveMozJpegEncodeOptions(quality) {
  const qualityPercent = resolveEncodeQualityPercent(quality);
  return {
    quality: qualityPercent,
    progressive: true,
    optimize_coding: true,
    smoothing: 0,
    trellis_multipass: qualityPercent < 96,
    trellis_opt_zero: qualityPercent < 96,
    trellis_opt_table: qualityPercent < 96,
    trellis_loops: qualityPercent < 84 ? 2 : 1,
    auto_subsample: true
  };
}

function resolveWebpEncodeOptions(quality) {
  const qualityPercent = resolveEncodeQualityPercent(quality);
  return {
    quality: qualityPercent,
    method: 6,
    sns_strength: qualityPercent >= 90 ? 35 : qualityPercent >= 78 ? 55 : 70,
    filter_strength: qualityPercent >= 90 ? 28 : qualityPercent >= 78 ? 42 : 56,
    alpha_quality: Math.max(72, qualityPercent),
    near_lossless: 100,
    lossless: 0
  };
}

function resolveOxipngLevelFromCompression(compression) {
  const safeCompression = normalizeExportCompression(compression);
  if (safeCompression >= 80) {
    return 4;
  }
  if (safeCompression >= 44) {
    return 3;
  }
  return 2;
}

async function encodeJpegFromCanvas(canvas, quality) {
  const encoded = await encodeJpeg(getCanvasImageData(canvas), resolveMozJpegEncodeOptions(quality));
  return new Blob([encoded], { type: 'image/jpeg' });
}

async function encodeWebpFromCanvas(canvas, quality) {
  const encoded = await encodeWebp(getCanvasImageData(canvas), resolveWebpEncodeOptions(quality));
  return new Blob([encoded], { type: 'image/webp' });
}

async function optimizePngArrayBuffer(pngBuffer, compression) {
  const optimized = await optimisePng(pngBuffer, {
    level: resolveOxipngLevelFromCompression(compression),
    interlace: false,
    optimiseAlpha: false
  });
  return new Blob([optimized], { type: 'image/png' });
}

async function encodeOptimizedPngFromCanvas(canvas, compression) {
  const optimized = await optimisePng(getCanvasImageData(canvas), {
    level: resolveOxipngLevelFromCompression(compression),
    interlace: false,
    optimiseAlpha: false
  });
  return new Blob([optimized], { type: 'image/png' });
}

function buildExportCanvasFromImage(image, options = {}) {
  const format = normalizeExportFormat(options.format || EXPORT_DEFAULT_FORMAT);
  const targetSize = Math.max(0, Math.round(Number(options.targetSize) || 0));
  const maxEdge = Math.max(0, Math.round(Number(options.maxEdge) || 0));
  const sourceWidth = Math.max(1, Number(image?.naturalWidth) || 1);
  const sourceHeight = Math.max(1, Number(image?.naturalHeight) || 1);
  const cropEnabled = options.cropEnabled === true;
  const cropRect = cropEnabled ? clampExportCropRect(options.cropRect) : createFullExportCropRect();
  const cropLeft = Math.max(0, Math.min(sourceWidth - 1, Math.round(cropRect.x * sourceWidth)));
  const cropTop = Math.max(0, Math.min(sourceHeight - 1, Math.round(cropRect.y * sourceHeight)));
  const cropRight = Math.max(
    cropLeft + 1,
    Math.min(sourceWidth, Math.round((cropRect.x + cropRect.width) * sourceWidth))
  );
  const cropBottom = Math.max(
    cropTop + 1,
    Math.min(sourceHeight, Math.round((cropRect.y + cropRect.height) * sourceHeight))
  );
  const cropWidth = Math.max(1, cropRight - cropLeft);
  const cropHeight = Math.max(1, cropBottom - cropTop);
  let width = targetSize > 0 ? targetSize : cropWidth;
  let height = targetSize > 0 ? targetSize : cropHeight;

  const sourceCanvas = document.createElement('canvas');
  sourceCanvas.width = cropWidth;
  sourceCanvas.height = cropHeight;
  const sourceCtx = sourceCanvas.getContext('2d');
  if (!sourceCtx) {
    throw new Error('canvas_context_unavailable');
  }
  sourceCtx.imageSmoothingEnabled = true;
  sourceCtx.imageSmoothingQuality = 'high';
  if (format === 'image/jpeg') {
    sourceCtx.fillStyle = '#ffffff';
    sourceCtx.fillRect(0, 0, cropWidth, cropHeight);
  } else {
    sourceCtx.clearRect(0, 0, cropWidth, cropHeight);
  }
  sourceCtx.drawImage(image, cropLeft, cropTop, cropWidth, cropHeight, 0, 0, cropWidth, cropHeight);

  if (options.transparentBackground === true && format !== 'image/jpeg') {
    removeBackgroundFromCanvas(sourceCanvas, {
      tolerance: normalizeExportBackgroundTolerance(options.backgroundTolerance),
      protection: normalizeExportBackgroundProtection(options.backgroundProtection)
    });
  }

  if (maxEdge > 0) {
    const dominantEdge = Math.max(width, height);
    if (dominantEdge > maxEdge) {
      const scale = maxEdge / dominantEdge;
      width = Math.max(1, Math.round(width * scale));
      height = Math.max(1, Math.round(height * scale));
    }
  }

  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('canvas_context_unavailable');
  }
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';

  if (format === 'image/jpeg') {
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }

  if (targetSize > 0) {
    const scale = Math.min(width / cropWidth, height / cropHeight);
    const drawWidth = Math.max(1, Math.round(cropWidth * scale));
    const drawHeight = Math.max(1, Math.round(cropHeight * scale));
    const drawX = Math.round((width - drawWidth) / 2);
    const drawY = Math.round((height - drawHeight) / 2);
    ctx.drawImage(sourceCanvas, 0, 0, cropWidth, cropHeight, drawX, drawY, drawWidth, drawHeight);
  } else {
    ctx.drawImage(sourceCanvas, 0, 0, cropWidth, cropHeight, 0, 0, width, height);
  }
  return canvas;
}

function removeBackgroundFromCanvas(canvas, options = {}) {
  const width = Math.max(1, Number(canvas?.width) || 1);
  const height = Math.max(1, Number(canvas?.height) || 1);
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('canvas_context_unavailable');
  }
  const imageData = ctx.getImageData(0, 0, width, height);
  const data = imageData.data;
  const safeTolerance = normalizeExportBackgroundTolerance(options.tolerance);
  const safeProtection = normalizeExportBackgroundProtection(options.protection);
  if (safeTolerance <= 0) {
    return;
  }

  const backgroundSamples = collectCanvasCornerSamples(data, width, height);
  if (backgroundSamples.length === 0) {
    return;
  }
  const backgroundReference = computeAverageBackgroundSample(backgroundSamples);

  const visited = new Uint8Array(width * height);
  const queue = new Uint32Array(width * height);
  let head = 0;
  let tail = 0;

  const enqueue = (x, y) => {
    if (x < 0 || x >= width || y < 0 || y >= height) {
      return;
    }
    const pixelIndex = y * width + x;
    if (visited[pixelIndex] === 1) {
      return;
    }
    const dataIndex = pixelIndex * 4;
    if (!pixelMatchesBackgroundSample(data, dataIndex, backgroundSamples, backgroundReference, safeTolerance)) {
      return;
    }
    visited[pixelIndex] = 1;
    queue[tail] = pixelIndex;
    tail += 1;
  };

  for (let x = 0; x < width; x += 1) {
    enqueue(x, 0);
    enqueue(x, height - 1);
  }
  for (let y = 0; y < height; y += 1) {
    enqueue(0, y);
    enqueue(width - 1, y);
  }

  while (head < tail) {
    const pixelIndex = queue[head];
    head += 1;
    const x = pixelIndex % width;
    const y = Math.floor(pixelIndex / width);
    enqueue(x - 1, y);
    enqueue(x + 1, y);
    enqueue(x, y - 1);
    enqueue(x, y + 1);
  }

  const protectionRadius =
    safeProtection <= 0
      ? 0
      : Math.max(1, Math.round((Math.min(width, height) * 0.18) * (safeProtection / EXPORT_BACKGROUND_PROTECTION_MAX)));

  if (protectionRadius <= 0) {
    for (let pixelIndex = 0; pixelIndex < visited.length; pixelIndex += 1) {
      if (visited[pixelIndex] !== 1) {
        continue;
      }
      data[(pixelIndex * 4) + 3] = 0;
    }
    ctx.putImageData(imageData, 0, 0);
    return;
  }

  const distances = new Int32Array(width * height);
  distances.fill(-1);
  const contourQueue = new Uint32Array(width * height);
  let contourHead = 0;
  let contourTail = 0;
  const contourOffsets = [
    [-1, -1], [0, -1], [1, -1],
    [-1, 0],           [1, 0],
    [-1, 1],  [0, 1],  [1, 1]
  ];

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const pixelIndex = y * width + x;
      if (visited[pixelIndex] !== 1) {
        continue;
      }
      let touchesForeground = false;
      for (const [dx, dy] of contourOffsets) {
        const nx = x + dx;
        const ny = y + dy;
        if (nx < 0 || nx >= width || ny < 0 || ny >= height) {
          continue;
        }
        const neighborIndex = (ny * width) + nx;
        if (visited[neighborIndex] !== 1) {
          touchesForeground = true;
          break;
        }
      }
      if (!touchesForeground) {
        continue;
      }
      distances[pixelIndex] = 0;
      contourQueue[contourTail] = pixelIndex;
      contourTail += 1;
    }
  }

  if (contourTail === 0) {
    for (let pixelIndex = 0; pixelIndex < visited.length; pixelIndex += 1) {
      if (visited[pixelIndex] !== 1) {
        continue;
      }
      data[(pixelIndex * 4) + 3] = 0;
    }
    ctx.putImageData(imageData, 0, 0);
    return;
  }

  while (contourHead < contourTail) {
    const pixelIndex = contourQueue[contourHead];
    contourHead += 1;
    const distance = distances[pixelIndex];
    if (distance >= protectionRadius) {
      continue;
    }
    const x = pixelIndex % width;
    const y = Math.floor(pixelIndex / width);
    for (const [dx, dy] of contourOffsets) {
      const nx = x + dx;
      const ny = y + dy;
      if (nx < 0 || nx >= width || ny < 0 || ny >= height) {
        continue;
      }
      const neighborIndex = (ny * width) + nx;
      if (visited[neighborIndex] !== 1 || distances[neighborIndex] !== -1) {
        continue;
      }
      distances[neighborIndex] = distance + 1;
      contourQueue[contourTail] = neighborIndex;
      contourTail += 1;
    }
  }

  const transparentCandidates = new Uint8Array(width * height);
  for (let pixelIndex = 0; pixelIndex < visited.length; pixelIndex += 1) {
    if (visited[pixelIndex] !== 1) {
      continue;
    }
    if (distances[pixelIndex] >= 0 && distances[pixelIndex] < protectionRadius) {
      continue;
    }
    transparentCandidates[pixelIndex] = 1;
  }

  const exteriorTransparent = new Uint8Array(width * height);
  const exteriorQueue = new Uint32Array(width * height);
  let exteriorHead = 0;
  let exteriorTail = 0;

  const enqueueExterior = (x, y) => {
    if (x < 0 || x >= width || y < 0 || y >= height) {
      return;
    }
    const pixelIndex = (y * width) + x;
    if (transparentCandidates[pixelIndex] !== 1 || exteriorTransparent[pixelIndex] === 1) {
      return;
    }
    exteriorTransparent[pixelIndex] = 1;
    exteriorQueue[exteriorTail] = pixelIndex;
    exteriorTail += 1;
  };

  for (let x = 0; x < width; x += 1) {
    enqueueExterior(x, 0);
    enqueueExterior(x, height - 1);
  }
  for (let y = 0; y < height; y += 1) {
    enqueueExterior(0, y);
    enqueueExterior(width - 1, y);
  }

  while (exteriorHead < exteriorTail) {
    const pixelIndex = exteriorQueue[exteriorHead];
    exteriorHead += 1;
    const x = pixelIndex % width;
    const y = Math.floor(pixelIndex / width);
    enqueueExterior(x - 1, y);
    enqueueExterior(x + 1, y);
    enqueueExterior(x, y - 1);
    enqueueExterior(x, y + 1);
  }

  for (let pixelIndex = 0; pixelIndex < exteriorTransparent.length; pixelIndex += 1) {
    if (exteriorTransparent[pixelIndex] !== 1) {
      continue;
    }
    data[(pixelIndex * 4) + 3] = 0;
  }

  ctx.putImageData(imageData, 0, 0);
}

function computeAverageBackgroundSample(samples) {
  const valid = Array.isArray(samples) ? samples : [];
  if (valid.length === 0) {
    return [255, 255, 255, 255];
  }
  let r = 0;
  let g = 0;
  let b = 0;
  let a = 0;
  for (const sample of valid) {
    r += Number(sample?.[0]) || 0;
    g += Number(sample?.[1]) || 0;
    b += Number(sample?.[2]) || 0;
    a += Number(sample?.[3]) || 0;
  }
  const divisor = Math.max(1, valid.length);
  return [
    Math.round(r / divisor),
    Math.round(g / divisor),
    Math.round(b / divisor),
    Math.round(a / divisor)
  ];
}

function collectCanvasCornerSamples(data, width, height) {
  if (!(data instanceof Uint8ClampedArray)) {
    return [];
  }
  const corners = [
    [0, 0],
    [Math.max(0, width - 1), 0],
    [0, Math.max(0, height - 1)],
    [Math.max(0, width - 1), Math.max(0, height - 1)]
  ];
  const samples = [];
  for (const [x, y] of corners) {
    const index = ((y * width) + x) * 4;
    const alpha = data[index + 3];
    if (alpha <= 0) {
      continue;
    }
    samples.push([data[index], data[index + 1], data[index + 2], alpha]);
  }
  return samples;
}

function pixelMatchesBackgroundSample(data, dataIndex, samples, reference, tolerance) {
  const alpha = data[dataIndex + 3];
  if (alpha <= 0) {
    return true;
  }
  const normalizedTolerance = (Math.max(0, tolerance) / Math.max(1, EXPORT_BACKGROUND_TOLERANCE_MAX)) * 255;
  const threshold = (normalizedTolerance * normalizedTolerance * 3.2);
  const ref = Array.isArray(reference) && reference.length >= 4 ? reference : [255, 255, 255, 255];
  const drRef = data[dataIndex] - ref[0];
  const dgRef = data[dataIndex + 1] - ref[1];
  const dbRef = data[dataIndex + 2] - ref[2];
  const daRef = alpha - ref[3];
  const refDistance = (drRef * drRef) + (dgRef * dgRef) + (dbRef * dbRef) + ((daRef * daRef) * 0.12);
  if (refDistance <= threshold) {
    return true;
  }
  for (const sample of samples) {
    const dr = data[dataIndex] - sample[0];
    const dg = data[dataIndex + 1] - sample[1];
    const db = data[dataIndex + 2] - sample[2];
    const da = alpha - sample[3];
    const distance = (dr * dr) + (dg * dg) + (db * db) + ((da * da) * 0.12);
    if (distance <= threshold) {
      return true;
    }
  }
  return false;
}

async function encodeCanvasToExportBlob(canvas, options = {}) {
  const format = normalizeExportFormat(options.format || EXPORT_DEFAULT_FORMAT);
  const needsQuality = format === 'image/jpeg' || format === 'image/webp';
  const requestedQuality = Number(options.quality);
  const quality = needsQuality
    ? Number.isFinite(requestedQuality)
      ? Math.min(1, Math.max(0, requestedQuality))
      : resolveLossyQualityFromCompression(EXPORT_DEFAULT_COMPRESSION)
    : undefined;
  const pngCompression = normalizeExportCompression(options.pngCompression);
  if (format === 'image/jpeg') {
    try {
      return await encodeJpegFromCanvas(canvas, quality);
    } catch (error) {
      console.error('MozJPEG encode failed, fallback to canvas encoder', error);
      return canvasToBlobAsync(canvas, format, quality);
    }
  }
  if (format === 'image/webp') {
    try {
      return await encodeWebpFromCanvas(canvas, quality);
    } catch (error) {
      console.error('WebP encode failed, fallback to canvas encoder', error);
      return canvasToBlobAsync(canvas, format, quality);
    }
  }
  if (format === 'image/png') {
    if (pngCompression > 0) {
      return encodeQuantizedPngFromCanvas(canvas, pngCompression);
    }
    try {
      return await encodeOptimizedPngFromCanvas(canvas, pngCompression);
    } catch (error) {
      console.error('OXIPNG encode failed, fallback to canvas encoder', error);
      return canvasToBlobAsync(canvas, 'image/png');
    }
  }
  return canvasToBlobAsync(canvas, format, quality);
}

async function encodeQuantizedPngFromCanvas(canvas, compression) {
  const width = Math.max(1, Number(canvas.width) || 1);
  const height = Math.max(1, Number(canvas.height) || 1);
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('canvas_context_unavailable');
  }
  const imageData = ctx.getImageData(0, 0, width, height);
  const rgba = new Uint8Array(imageData.data);
  const colorCount = resolvePngColorCountFromCompression(compression);
  if (colorCount <= 0) {
    try {
      return await encodeOptimizedPngFromCanvas(canvas, compression);
    } catch (error) {
      console.error('OXIPNG encode failed, fallback to canvas PNG', error);
      return canvasToBlobAsync(canvas, 'image/png');
    }
  }
  try {
    const encoded = UPNG.encode([rgba.buffer], width, height, colorCount);
    try {
      return await optimizePngArrayBuffer(encoded, compression);
    } catch (optimizationError) {
      console.error('OXIPNG optimise failed after quantization, fallback to raw quantized PNG', optimizationError);
      return new Blob([encoded], { type: 'image/png' });
    }
  } catch (error) {
    console.error('UPNG quantization failed, fallback to canvas PNG', error);
    try {
      return await encodeOptimizedPngFromCanvas(canvas, compression);
    } catch (optimizationError) {
      console.error('OXIPNG encode failed, fallback to canvas PNG', optimizationError);
      return canvasToBlobAsync(canvas, 'image/png');
    }
  }
}

function exportExtensionFromFormat(format) {
  const safe = normalizeExportFormat(format);
  if (safe === 'image/png') {
    return 'png';
  }
  if (safe === 'image/jpeg') {
    return 'jpg';
  }
  return 'webp';
}

function exportFormatLabel(format) {
  const safe = normalizeExportFormat(format);
  if (safe === 'image/png') {
    return 'PNG';
  }
  if (safe === 'image/jpeg') {
    return 'JPG';
  }
  return 'WebP';
}

function resolveOriginalMatchedSizeExportConfig() {
  const item = state.exportWizardItem;
  const originalFormat = extractComparableExportFormatFromDataUrl(item?.dataUrl);
  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  if (!item?.dataUrl || !originalFormat) {
    return {
      available: false,
      reason: 'Formato original no compatible'
    };
  }
  if (preset === EXPORT_PRESET_FAVICON) {
    return {
      available: false,
      reason: 'No disponible con favicon pack'
    };
  }
  if (editOptions.transparentBackground === true && originalFormat === 'image/jpeg') {
    return {
      available: false,
      reason: 'El JPG original no admite transparencia'
    };
  }
  return {
    available: true,
    format: originalFormat,
    formatLabel: exportFormatLabel(originalFormat),
    extension: exportExtensionFromFormat(originalFormat)
  };
}

function buildExportCompressionSweepCandidates(step = 4) {
  const safeStep = Math.max(1, Math.round(Number(step) || 1));
  const candidates = [];
  for (let compression = 0; compression <= 100; compression += safeStep) {
    candidates.push(normalizeExportCompression(compression));
  }
  if (!candidates.includes(100)) {
    candidates.push(100);
  }
  return Array.from(new Set(candidates)).sort((left, right) => left - right);
}

function pickClosestExportResultBySizePreferLowerCompression(results, targetSize) {
  const safeTargetSize = Math.max(1, Number(targetSize) || 1);
  const valid = Array.isArray(results)
    ? results.filter((entry) => Number.isFinite(entry?.compression) && Number.isFinite(entry?.size) && entry.size > 0)
    : [];
  if (valid.length === 0) {
    return null;
  }
  return valid.reduce((winner, entry) => {
    if (!winner) {
      return entry;
    }
    const winnerDiff = Math.abs(winner.size - safeTargetSize);
    const entryDiff = Math.abs(entry.size - safeTargetSize);
    if (entryDiff < winnerDiff) {
      return entry;
    }
    if (entryDiff > winnerDiff) {
      return winner;
    }
    const winnerFits = winner.size <= safeTargetSize;
    const entryFits = entry.size <= safeTargetSize;
    if (entryFits !== winnerFits) {
      return entryFits ? entry : winner;
    }
    if (entryFits && winnerFits && entry.size !== winner.size) {
      return entry.size > winner.size ? entry : winner;
    }
    if (!entryFits && !winnerFits && entry.size !== winner.size) {
      return entry.size < winner.size ? entry : winner;
    }
    return entry.compression < winner.compression ? entry : winner;
  }, null);
}

async function resolveOriginalMatchedSizeExportPlan() {
  const exportConfig = resolveOriginalMatchedSizeExportConfig();
  if (!exportConfig.available || !state.exportWizardItem?.dataUrl) {
    return null;
  }
  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  const context = await buildExportRecommendationContext(state.exportWizardItem, {
    format: exportConfig.format,
    preset,
    ...editOptions,
    ignoreInitialCompressionFloor: true
  }, exportConfig.format);
  if (context?.fallbackResult) {
    return {
      format: exportConfig.format,
      compression: EXPORT_DEFAULT_COMPRESSION,
      estimatedSize: Math.max(1, Number(context.fallbackResult.size) || 1)
    };
  }

  const coarseCandidates = buildExportCompressionSweepCandidates(4);
  const measuredResults = [];
  for (const candidate of coarseCandidates) {
    measuredResults.push(await context.measureValidationResult(candidate));
  }
  const coarseClosest =
    pickClosestExportResultBySizePreferLowerCompression(measuredResults, context.originalSize) || measuredResults[0];
  const refinedCandidates = new Set(coarseCandidates);
  for (
    let candidate = Math.max(0, normalizeExportCompression(coarseClosest?.compression) - 6);
    candidate <= Math.min(100, normalizeExportCompression(coarseClosest?.compression) + 6);
    candidate += 1
  ) {
    refinedCandidates.add(candidate);
  }
  const refinedResults = measuredResults.slice();
  for (const candidate of Array.from(refinedCandidates).sort((left, right) => left - right)) {
    if (measuredResults.some((entry) => normalizeExportCompression(entry.compression) === candidate)) {
      continue;
    }
    refinedResults.push(await context.measureValidationResult(candidate));
  }
  const matched =
    pickClosestExportResultBySizePreferLowerCompression(refinedResults, context.originalSize) ||
    coarseClosest ||
    {
      compression: EXPORT_DEFAULT_COMPRESSION,
      size: Math.max(1, context.originalSize || 1)
    };
  return {
    format: exportConfig.format,
    compression: normalizeExportCompression(matched.compression),
    estimatedSize: Math.max(1, Number(matched.size) || 1)
  };
}

function updateExportWizardActionUi() {
  const preset = normalizeExportPreset(state.exportWizardPreset);
  const editOptions = getCurrentExportEditOptions();
  const format = resolveEffectiveExportFormat(state.exportWizardFormat, preset, editOptions);
  const formatLabel = exportFormatLabel(format);
  const originalMatchedExportConfig = resolveOriginalMatchedSizeExportConfig();
  const isBusy = state.exportWizardBusy === true;
  const originalBytes =
    Math.max(0, Number(state.exportWizardItem?.fileSize) || 0) || getDataUrlByteLength(state.exportWizardItem?.dataUrl);
  const estimatedBytes = Math.max(0, Number(state.exportWizardEstimatedOutputBytes) || 0);
  let gainLabel = '';
  if (originalBytes > 0 && estimatedBytes > 0) {
    const deltaRatio = (estimatedBytes - originalBytes) / originalBytes;
    const deltaPercent = Math.round(Math.abs(deltaRatio) * 100);
    if (deltaPercent > 0) {
      gainLabel = deltaRatio < 0 ? `${deltaPercent}% menos` : `${deltaPercent}% más`;
    } else {
      gainLabel = 'Mismo tamaño';
    }
  }
  const title = isBusy
    ? 'Exportando...'
    : preset === EXPORT_PRESET_FAVICON
      ? 'Exportar favicon pack'
      : 'Exportar imagen';
  const hint = isBusy
    ? 'Preparando la descarga'
    : preset === EXPORT_PRESET_FAVICON
      ? gainLabel
        ? `${formatLabel} · ${gainLabel}`
        : `${formatLabel} · ZIP multimida`
      : gainLabel
        ? `${formatLabel} · ${gainLabel}`
        : `${formatLabel} · Descarga directa`;
  const originalTitle = isBusy
    ? 'Exportando...'
    : originalMatchedExportConfig.available
      ? 'Exportar sin pérdidas'
      : 'Sin pérdidas no disponible';

  if (elements.exportWizardDownloadBtn instanceof HTMLButtonElement) {
    elements.exportWizardDownloadBtn.disabled = isBusy;
    elements.exportWizardDownloadBtn.setAttribute('aria-busy', isBusy ? 'true' : 'false');
    elements.exportWizardDownloadBtn.title = title;
  }
  if (elements.exportWizardDownloadLabel instanceof HTMLElement) {
    elements.exportWizardDownloadLabel.textContent = title;
  }
  if (elements.exportWizardDownloadHint instanceof HTMLElement) {
    elements.exportWizardDownloadHint.textContent = hint;
  }
  if (elements.exportWizardOriginalDownloadBtn instanceof HTMLButtonElement) {
    elements.exportWizardOriginalDownloadBtn.disabled = isBusy || !originalMatchedExportConfig.available;
    elements.exportWizardOriginalDownloadBtn.setAttribute('aria-busy', isBusy ? 'true' : 'false');
    elements.exportWizardOriginalDownloadBtn.title = originalTitle;
  }
  if (elements.exportWizardOriginalDownloadLabel instanceof HTMLElement) {
    elements.exportWizardOriginalDownloadLabel.textContent = 'Exportar sin pérdidas';
  }
}

async function downloadBlobAsset(blob, filename) {
  if (!(blob instanceof Blob) || blob.size <= 0) {
    throw new Error('invalid_blob_download');
  }
  const safeFilename = String(filename || '').trim() || `export-${Date.now()}`;
  const objectUrl = URL.createObjectURL(blob);

  try {
    if (chrome?.downloads?.download) {
      await chrome.downloads.download({
        url: objectUrl,
        filename: safeFilename,
        saveAs: true
      });
      return;
    }

    const anchor = document.createElement('a');
    anchor.href = objectUrl;
    anchor.download = safeFilename;
    anchor.click();
  } finally {
    window.setTimeout(() => {
      URL.revokeObjectURL(objectUrl);
    }, 60_000);
  }
}

function removeFileExtension(name) {
  const safe = String(name || '').trim();
  if (!safe) {
    return '';
  }
  return safe.replace(/\.[a-z0-9]{2,8}$/iu, '');
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
  if (!visible) {
    closeManualUploadModal();
  }
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
    const manualSizeBytes = state.manualItems.reduce(
      (sum, item) => sum + getStoredItemSizeBytes(item),
      0
    );
    collections.push({
      id: MANUAL_GALLERY_ID,
      title: 'Manual',
      subtitle: 'Pujades manuals',
      count: state.manualItems.length,
      sizeBytes: manualSizeBytes,
      updatedAt: latestManual,
      type: 'manual'
    });
  }

  for (const gallery of state.savedGalleries) {
    const domain = String(gallery?.domain || '').trim();
    if (!domain) {
      continue;
    }
    const items = Array.isArray(gallery.items) ? gallery.items : [];
    const gallerySizeBytes = items.reduce((sum, item) => sum + getStoredItemSizeBytes(item), 0);
    collections.push({
      id: String(gallery.id || '').trim(),
      title: getGalleryDisplayName(gallery),
      subtitle: 'Web guardada',
      count: items.length,
      sizeBytes: gallerySizeBytes,
      faviconUrl: resolveGalleryFaviconUrl(gallery),
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
    grid.appendChild(createStoredGalleryCard(item, gallery.id));
  }
}

function createGalleryListCard(entry) {
  const cardWrap = document.createElement('div');
  cardWrap.className = 'saved-gallery-entry-wrap';

  const card = document.createElement('button');
  card.type = 'button';
  card.className = 'saved-gallery-entry';
  card.dataset.galleryOpenId = String(entry.id || '').trim();

  const titleRow = document.createElement('span');
  titleRow.className = 'saved-gallery-entry-title-row';

  const title = document.createElement('span');
  title.className = 'saved-gallery-entry-title';
  title.textContent = String(entry.title || 'Galeria');
  titleRow.appendChild(title);

  if (entry.type === 'domain') {
    const faviconUrl = String(entry.faviconUrl || '').trim();
    if (faviconUrl) {
      const favicon = document.createElement('img');
      favicon.className = 'saved-gallery-entry-favicon';
      favicon.src = faviconUrl;
      favicon.alt = '';
      favicon.loading = 'lazy';
      favicon.decoding = 'async';
      favicon.setAttribute('aria-hidden', 'true');
      favicon.addEventListener('error', () => {
        favicon.remove();
      });
      titleRow.prepend(favicon);
    }
  }

  const subtitle = document.createElement('span');
  subtitle.className = 'saved-gallery-entry-subtitle';
  subtitle.textContent = String(entry.subtitle || '');

  const count = document.createElement('span');
  count.className = 'saved-gallery-entry-count';
  count.textContent = `${Number(entry.count) || 0} imatges`;

  const size = document.createElement('span');
  size.className = 'saved-gallery-entry-size';
  size.textContent = formatGallerySizeMb(entry.sizeBytes);

  card.appendChild(titleRow);
  card.appendChild(subtitle);
  card.appendChild(count);
  card.appendChild(size);

  const actions = document.createElement('div');
  actions.className = 'saved-gallery-entry-actions';

  const downloadButton = document.createElement('button');
  downloadButton.type = 'button';
  downloadButton.className = 'saved-gallery-entry-download';
  downloadButton.dataset.galleryDownloadId = String(entry.id || '').trim();
  downloadButton.setAttribute('aria-label', `Descarregar galeria ${entry.title}`);
  downloadButton.title = 'Descarregar galeria (ZIP)';
  downloadButton.innerHTML = '<i class="bi bi-download" aria-hidden="true"></i>';
  downloadButton.disabled = Number(entry.count) <= 0;

  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.className = 'saved-gallery-entry-delete';
  deleteButton.dataset.galleryDeleteId = String(entry.id || '').trim();
  deleteButton.setAttribute('aria-label', `Eliminar galeria ${entry.title}`);
  deleteButton.title = 'Eliminar galeria';
  deleteButton.innerHTML = '<i class="bi bi-trash3" aria-hidden="true"></i>';

  actions.appendChild(downloadButton);
  actions.appendChild(deleteButton);
  cardWrap.appendChild(card);
  cardWrap.appendChild(actions);
  return cardWrap;
}

async function downloadGalleryAsZip(galleryId) {
  const safeGalleryId = String(galleryId || '').trim();
  if (!safeGalleryId) {
    return;
  }

  const gallerySnapshot = getGallerySnapshotById(safeGalleryId);
  if (!gallerySnapshot || gallerySnapshot.items.length === 0) {
    showFeedbackToast('Aquesta galeria no te imatges', 'error');
    return;
  }

  try {
    const usedNames = new Set();
    const zipEntries = [];

    for (let index = 0; index < gallerySnapshot.items.length; index += 1) {
      const item = gallerySnapshot.items[index];
      const dataUrl = String(item?.dataUrl || '').trim();
      const bytes = decodeDataUrlToBytes(dataUrl);
      if (!bytes || bytes.length === 0) {
        continue;
      }
      const preferredName = buildZipEntryName(item, index);
      const uniqueName = ensureUniqueArchiveName(preferredName, usedNames);
      zipEntries.push({ name: uniqueName, bytes });
    }

    if (zipEntries.length === 0) {
      showFeedbackToast('No hi ha imatges valides per descarregar', 'error');
      return;
    }

    const archiveName = buildGalleryArchiveFileName(gallerySnapshot);
    const blob = createZipBlob(zipEntries);
    const objectUrl = URL.createObjectURL(blob);

    try {
      if (chrome?.downloads?.download) {
        await chrome.downloads.download({
          url: objectUrl,
          filename: archiveName,
          saveAs: true
        });
      } else {
        const anchor = document.createElement('a');
        anchor.href = objectUrl;
        anchor.download = archiveName;
        anchor.click();
      }
      showFeedbackToast(`ZIP descarregat (${zipEntries.length} imatges)`);
    } finally {
      window.setTimeout(() => {
        URL.revokeObjectURL(objectUrl);
      }, 60_000);
    }
  } catch (error) {
    console.error('Gallery zip download failed', error);
    showFeedbackToast('No s\'ha pogut descarregar la galeria', 'error');
  }
}

function getGallerySnapshotById(galleryId) {
  const safeGalleryId = String(galleryId || '').trim();
  if (!safeGalleryId) {
    return null;
  }

  if (safeGalleryId === MANUAL_GALLERY_ID) {
    return {
      id: MANUAL_GALLERY_ID,
      name: 'manual',
      items: Array.isArray(state.manualItems) ? state.manualItems : []
    };
  }

  const gallery = state.savedGalleries.find((entry) => String(entry?.id || '').trim() === safeGalleryId);
  if (!gallery) {
    return null;
  }
  return {
    id: safeGalleryId,
    name: getGalleryDisplayName(gallery),
    items: Array.isArray(gallery.items) ? gallery.items : []
  };
}

function decodeDataUrlToBytes(dataUrl) {
  const safeDataUrl = String(dataUrl || '').trim();
  if (!safeDataUrl.startsWith('data:image/')) {
    return null;
  }
  const commaIndex = safeDataUrl.indexOf(',');
  if (commaIndex < 0) {
    return null;
  }

  const metadata = safeDataUrl.slice(5, commaIndex);
  const payload = safeDataUrl.slice(commaIndex + 1);
  const isBase64 = metadata.toLowerCase().includes(';base64');

  if (isBase64) {
    const base64 = payload.replace(/\s+/g, '');
    if (!base64) {
      return null;
    }
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  }

  try {
    const decoded = decodeURIComponent(payload);
    return new TextEncoder().encode(decoded);
  } catch {
    return null;
  }
}

function getDataUrlByteLength(dataUrl) {
  const bytes = decodeDataUrlToBytes(dataUrl);
  return bytes ? bytes.length : 0;
}

function buildGalleryArchiveFileName(gallerySnapshot) {
  const fallback = `galeria-${Date.now()}.zip`;
  const name = sanitizeArchiveName(gallerySnapshot?.name || '');
  if (!name) {
    return fallback;
  }
  return `${name}.zip`;
}

function buildZipEntryName(item, index = 0) {
  const sourceName = extractFileNameFromItemSource(item);
  const fromItemName = String(item?.name || '').trim();
  let name = sanitizeArchiveName(sourceName || fromItemName || `imagen-${index + 1}`);
  if (!name) {
    name = `imagen-${index + 1}`;
  }

  if (!/\.[a-z0-9]{2,8}$/iu.test(name)) {
    const ext = extractImageExtensionFromDataUrl(item?.dataUrl);
    if (ext) {
      name = `${name}.${ext}`;
    }
  }
  return name;
}

function extractFileNameFromItemSource(item) {
  const candidates = [String(item?.sourcePath || '').trim(), String(item?.sourceUrl || '').trim()];
  for (const candidate of candidates) {
    if (!candidate) {
      continue;
    }
    try {
      const parsed = candidate.startsWith('http://') || candidate.startsWith('https://')
        ? new URL(candidate)
        : new URL(candidate, 'https://placeholder.local');
      const segments = String(parsed.pathname || '')
        .split('/')
        .filter(Boolean);
      if (segments.length === 0) {
        continue;
      }
      const rawName = decodeURIComponent(segments[segments.length - 1]);
      if (rawName) {
        return rawName;
      }
    } catch {
      // ignore malformed source
    }
  }
  return '';
}

function extractImageExtensionFromDataUrl(dataUrl) {
  const safe = String(dataUrl || '').trim();
  const match = safe.match(/^data:image\/([a-z0-9.+-]+)(?:;|,)/iu);
  if (!match) {
    return '';
  }
  const mimeSubtype = String(match[1] || '').toLowerCase();
  if (!mimeSubtype) {
    return '';
  }
  if (mimeSubtype === 'jpeg') {
    return 'jpg';
  }
  if (mimeSubtype === 'svg+xml') {
    return 'svg';
  }
  if (mimeSubtype === 'x-icon' || mimeSubtype === 'vnd.microsoft.icon') {
    return 'ico';
  }
  const normalized = mimeSubtype.split('+')[0];
  return normalized || '';
}

function formatOriginalImageFormatLabel(dataUrl) {
  const ext = extractImageExtensionFromDataUrl(dataUrl);
  if (!ext) {
    return 'IMG';
  }
  if (ext === 'jpg') {
    return 'JPG';
  }
  if (ext === 'ico') {
    return 'ICO';
  }
  return ext.toUpperCase();
}

function sanitizeArchiveName(value) {
  const safe = String(value || '')
    .trim()
    .replace(/[\\/:*?"<>|]/gu, '-')
    .replace(/\s+/gu, ' ');
  if (!safe) {
    return '';
  }
  return safe.length > 140 ? safe.slice(0, 140) : safe;
}

function ensureUniqueArchiveName(name, usedNames) {
  const safeName = String(name || '').trim() || 'imatge';
  if (!usedNames.has(safeName)) {
    usedNames.add(safeName);
    return safeName;
  }

  const dotIndex = safeName.lastIndexOf('.');
  const hasExt = dotIndex > 0 && dotIndex < safeName.length - 1;
  const base = hasExt ? safeName.slice(0, dotIndex) : safeName;
  const ext = hasExt ? safeName.slice(dotIndex) : '';

  let suffix = 2;
  while (suffix < 10_000) {
    const candidate = `${base}-${suffix}${ext}`;
    if (!usedNames.has(candidate)) {
      usedNames.add(candidate);
      return candidate;
    }
    suffix += 1;
  }

  const fallback = `${base}-${Date.now()}${ext}`;
  usedNames.add(fallback);
  return fallback;
}

function createZipBlob(entries, options = {}) {
  const safeEntries = Array.isArray(entries) ? entries.filter(Boolean) : [];
  const compressionPercent =
    options && Object.prototype.hasOwnProperty.call(options, 'compressionPercent')
      ? normalizeExportCompression(options.compressionPercent)
      : 100;
  const compressionLevel = resolveZipCompressionLevel(compressionPercent);
  const zipInput = {};

  for (const entry of safeEntries) {
    const fileName = String(entry?.name || '').trim();
    const data = entry?.bytes instanceof Uint8Array ? entry.bytes : null;
    if (!fileName || !data) {
      continue;
    }
    zipInput[fileName] = [data, { level: compressionLevel }];
  }

  const zipped = zipSync(zipInput, { level: compressionLevel });
  return new Blob([zipped], { type: 'application/zip' });
}

function getStoredItemSizeBytes(item) {
  const rawSize = Math.max(0, Number(item?.fileSize) || 0);
  if (rawSize > 0) {
    return rawSize;
  }

  const dataUrl = String(item?.dataUrl || '').trim();
  if (!dataUrl.startsWith('data:image/')) {
    return 0;
  }
  const commaIndex = dataUrl.indexOf(',');
  if (commaIndex < 0) {
    return 0;
  }
  const base64 = dataUrl.slice(commaIndex + 1).replace(/\s+/g, '');
  if (!base64) {
    return 0;
  }

  const padding = base64.endsWith('==') ? 2 : base64.endsWith('=') ? 1 : 0;
  return Math.max(0, Math.floor((base64.length * 3) / 4) - padding);
}

function formatGallerySizeMb(bytes) {
  const safeBytes = Math.max(0, Number(bytes) || 0);
  if (safeBytes <= 0) {
    return '0 MB';
  }
  const megaBytes = safeBytes / (1024 * 1024);
  if (megaBytes < 0.01) {
    return '0.01 MB';
  }
  return `${megaBytes.toFixed(2)} MB`;
}

function formatBytes(bytes) {
  const safe = Math.max(0, Number(bytes) || 0);
  if (safe < 1024) {
    return `${safe} B`;
  }
  if (safe < 1024 * 1024) {
    return `${(safe / 1024).toFixed(1)} KB`;
  }
  return `${(safe / (1024 * 1024)).toFixed(2)} MB`;
}

function formatBytesPrecise(bytes) {
  const safe = Math.max(0, Math.round(Number(bytes) || 0));
  if (safe < 1024) {
    return `${safe} B`;
  }
  if (safe < 1024 * 1024) {
    return `${(safe / 1024).toFixed(2)} KB (${safe} B)`;
  }
  return `${(safe / (1024 * 1024)).toFixed(3)} MB (${safe} B)`;
}

function getGalleryDisplayName(gallery) {
  const rawName = String(gallery?.name || '').trim();
  if (rawName) {
    return rawName;
  }
  const domain = String(gallery?.domain || '').trim();
  if (domain) {
    return domain;
  }
  return 'Galeria';
}

function resolveGalleryFaviconUrl(gallery) {
  const stored = String(gallery?.faviconUrl || '').trim();
  if (stored) {
    return stored;
  }
  return buildGalleryFaviconUrl(gallery?.pageOrigin, gallery?.domain);
}

function buildGalleryFaviconUrl(pageOrigin, domain) {
  let pageUrl = '';
  const safePageOrigin = String(pageOrigin || '').trim();
  if (safePageOrigin) {
    try {
      pageUrl = new URL(safePageOrigin).href;
    } catch {
      pageUrl = '';
    }
  }

  if (!pageUrl) {
    const safeDomain = String(domain || '').trim().toLowerCase();
    if (!safeDomain) {
      return '';
    }
    pageUrl = `https://${safeDomain}`;
  }

  return chrome.runtime.getURL(`/_favicon/?pageUrl=${encodeURIComponent(pageUrl)}&size=32`);
}

function createStoredGalleryCard(item, galleryId) {
  const card = document.createElement('div');
  card.className = 'logo-library-item stored-gallery-item';

  const thumbWrap = createThumbWrap(item);
  const name = document.createElement('span');
  name.className = 'logo-library-name';
  name.textContent = String(item.name || 'Imatge');

  card.appendChild(thumbWrap);
  card.appendChild(name);

  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.className = 'logo-library-item-delete-small';
  deleteButton.dataset.galleryItemDeleteId = String(item?.id || '').trim();
  deleteButton.dataset.galleryItemGalleryId = String(galleryId || '').trim();
  deleteButton.setAttribute('aria-label', `Eliminar ${item.name}`);
  deleteButton.title = 'Eliminar imatge';
  deleteButton.innerHTML = '<i class="bi bi-trash3" aria-hidden="true"></i>';
  card.appendChild(deleteButton);

  const exportButton = createImageExportButton(item, { galleryId });
  if (exportButton) {
    card.appendChild(exportButton);
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
  const exportButton = createImageExportButton(item, {
    manual: Boolean(options.manual),
    galleryId: options.manual ? MANUAL_GALLERY_ID : ''
  });
  if (exportButton) {
    thumbWrap.appendChild(exportButton);
  }
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

function createImageExportButton(item, options = {}) {
  const imageId = String(item?.id || '').trim();
  if (!imageId) {
    return null;
  }

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'logo-library-item-export';
  button.dataset.imageExportId = imageId;
  if (options.manual) {
    button.dataset.imageExportManual = '1';
  }
  if (options.galleryId) {
    button.dataset.imageExportGalleryId = String(options.galleryId || '').trim();
  }
  const imageName = String(item?.name || 'imatge');
  button.setAttribute('aria-label', `Export wizard ${imageName}`);
  button.title = 'Export wizard';
  button.innerHTML = '<i class="bi bi-stars" aria-hidden="true"></i>';
  return button;
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
    openLibraryPageWhenWebGalleryUnavailable();
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
      openLibraryPageWhenWebGalleryUnavailable();
    }
  } catch (error) {
    console.error('Web gallery load failed', error);
    state.webItems = [];
    state.selectedWebIds = new Set();
    showFeedbackToast('Error carregant imatges web', 'error');
    openLibraryPageWhenWebGalleryUnavailable();
  } finally {
    state.webLoading = false;
    renderActiveLogoLibraryGrid();
  }
}

async function saveSelectedWebImagesToDomainGallery(options = {}) {
  if (state.galleryView !== VIEW_WEB) {
    return;
  }

  const selected = state.webItems.filter((item) => state.selectedWebIds.has(String(item.id || '').trim()));
  if (selected.length === 0) {
    showFeedbackToast('Selecciona imatges per desar', 'error');
    return;
  }

  const domainContext = getActiveTabDomainContext();
  if (!domainContext.isValid || !domainContext.domain) {
    showFeedbackToast('No s\'ha pogut detectar el domini', 'error');
    return;
  }
  const domain = domainContext.domain;
  const pageOrigin = domainContext.pageOrigin;

  const customNameRaw = String(options.customName || '').trim();
  const hasCustomName = customNameRaw.length > 0;
  const galleryName = sanitizeGalleryName(hasCustomName ? customNameRaw : domain, domain);
  const galleryId = hasCustomName
    ? buildCustomGalleryId(galleryName)
    : `domain:${domain}`;
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
    name: galleryName,
    domain,
    pageOrigin,
    faviconUrl: buildGalleryFaviconUrl(pageOrigin, domain),
    updatedAt: now,
    items: Array.from(mergedBySource.values()).sort((a, b) => Number(b.savedAt) - Number(a.savedAt))
  };

  const others = state.savedGalleries.filter((entry) => String(entry?.id || '') !== galleryId);
  state.savedGalleries = [updatedGallery, ...others].sort((a, b) => Number(b.updatedAt) - Number(a.updatedAt));
  await persistDomainGalleries();

  state.galleryView = VIEW_LIBRARY;
  state.galleryListMode = 'detail';
  state.activeGalleryId = galleryId;
  setActivePage('library');
  showFeedbackToast(`${inserted} imatges desades a ${galleryName}`);
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

async function addManualImages(files, uploadTarget = null) {
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
        createdAt: Date.now() + created.length,
        savedAt: Date.now() + created.length
      });
    } catch {
      skipped += 1;
    }
  }

  if (created.length === 0) {
    showFeedbackToast('No s\'ha pogut pujar cap imatge', 'error');
    return;
  }

  const target = resolveManualUploadTarget(uploadTarget);
  let stored = false;

  if (target.mode === 'existing') {
    stored = await addManualImagesToExistingGallery(created, target.galleryId);
  } else if (target.mode === 'new') {
    stored = await addManualImagesToNamedGallery(created, target.name);
  } else {
    stored = await addManualImagesToManualGallery(created);
  }

  if (!stored) {
    return;
  }

  if (skipped > 0) {
    showFeedbackToast(`${created.length} pujades, ${skipped} descartades`);
    return;
  }
  showFeedbackToast(`${created.length} imatges pujades`);
}

function resolveManualUploadTarget(uploadTarget) {
  if (!uploadTarget || typeof uploadTarget !== 'object') {
    return { mode: 'manual' };
  }

  const mode = String(uploadTarget.mode || '').trim();
  if (mode === 'existing') {
    const galleryId = String(uploadTarget.galleryId || '').trim();
    if (galleryId) {
      return { mode: 'existing', galleryId };
    }
    return { mode: 'manual' };
  }

  if (mode === 'new') {
    const rawName = String(uploadTarget.name || '').trim();
    if (rawName) {
      const name = sanitizeGalleryName(rawName, 'Galeria');
      return { mode: 'new', name };
    }
    return { mode: 'manual' };
  }

  return { mode: 'manual' };
}

async function addManualImagesToManualGallery(created) {
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
  return true;
}

async function addManualImagesToExistingGallery(created, galleryId) {
  const safeGalleryId = String(galleryId || '').trim();
  if (!safeGalleryId) {
    return addManualImagesToManualGallery(created);
  }
  if (safeGalleryId === MANUAL_GALLERY_ID) {
    return addManualImagesToManualGallery(created);
  }

  const existing = state.savedGalleries.find((entry) => String(entry?.id || '').trim() === safeGalleryId);
  if (!existing) {
    showFeedbackToast('No s\'ha trobat la galeria seleccionada', 'error');
    return false;
  }

  const now = Date.now();
  const appended = created.map((item, index) => mapManualAssetToGalleryItem(item, now + index));
  const updatedGallery = {
    ...existing,
    updatedAt: now,
    items: [...appended, ...(Array.isArray(existing.items) ? existing.items : [])]
  };

  const others = state.savedGalleries.filter((entry) => String(entry?.id || '').trim() !== safeGalleryId);
  state.savedGalleries = [updatedGallery, ...others].sort((a, b) => Number(b.updatedAt) - Number(a.updatedAt));
  await persistDomainGalleries();

  state.galleryView = VIEW_LIBRARY;
  state.galleryListMode = 'detail';
  state.activeGalleryId = safeGalleryId;
  setActivePage('library');
  return true;
}

async function addManualImagesToNamedGallery(created, rawName) {
  const { domain, pageOrigin } = resolveGalleryDomainContextFromActiveTab();
  const galleryName = sanitizeGalleryName(rawName || '', domain);
  if (!galleryName) {
    showFeedbackToast('Nom de galeria invalid', 'error');
    return false;
  }

  const galleryId = buildCustomGalleryId(galleryName);
  const existing = state.savedGalleries.find((entry) => String(entry?.id || '').trim() === galleryId);
  const now = Date.now();
  const appended = created.map((item, index) => mapManualAssetToGalleryItem(item, now + index));

  const updatedGallery = {
    id: galleryId,
    name: galleryName,
    domain: String(existing?.domain || domain || 'manual.local').trim().toLowerCase(),
    pageOrigin: String(existing?.pageOrigin || pageOrigin || '').trim(),
    faviconUrl: String(
      existing?.faviconUrl ||
      buildGalleryFaviconUrl(existing?.pageOrigin || pageOrigin, existing?.domain || domain || 'manual.local')
    ).trim(),
    updatedAt: now,
    items: [...appended, ...(Array.isArray(existing?.items) ? existing.items : [])]
  };

  const others = state.savedGalleries.filter((entry) => String(entry?.id || '').trim() !== galleryId);
  state.savedGalleries = [updatedGallery, ...others].sort((a, b) => Number(b.updatedAt) - Number(a.updatedAt));
  await persistDomainGalleries();

  state.galleryView = VIEW_LIBRARY;
  state.galleryListMode = 'detail';
  state.activeGalleryId = galleryId;
  setActivePage('library');
  return true;
}

function mapManualAssetToGalleryItem(item, savedAt) {
  const manualId = String(item?.id || '').trim() || `manual-item-${savedAt}`;
  return {
    id: manualId,
    name: sanitizeDisplayName(item?.name || 'Imatge', 'Imatge'),
    dataUrl: String(item?.dataUrl || ''),
    sourceUrl: '',
    sourcePath: String(item?.sourcePath || '').trim(),
    fileSize: Math.max(0, Number(item?.fileSize) || 0),
    width: Math.max(0, Number(item?.width) || 0),
    height: Math.max(0, Number(item?.height) || 0),
    htmlWidth: 0,
    htmlHeight: 0,
    htmlFixed: false,
    htmlResponsive: false,
    htmlUnconstrained: false,
    canHighlight: false,
    savedAt: Number(savedAt) || Date.now()
  };
}

function getActiveTabDomainContext() {
  const rawUrl = String(state.activeTabUrl || '').trim();
  if (!rawUrl) {
    return { isValid: false, domain: '', pageOrigin: '', href: '' };
  }

  try {
    const parsed = new URL(rawUrl);
    const protocol = String(parsed.protocol || '').toLowerCase();
    const domain = String(parsed.hostname || '').trim().toLowerCase();
    const isValidProtocol = protocol === 'http:' || protocol === 'https:';
    if (!isValidProtocol || !domain) {
      return { isValid: false, domain: '', pageOrigin: '', href: '' };
    }
    return {
      isValid: true,
      domain,
      pageOrigin: String(parsed.origin || '').trim(),
      href: String(parsed.href || rawUrl).trim()
    };
  } catch {
    return { isValid: false, domain: '', pageOrigin: '', href: '' };
  }
}

function getDefaultSaveAsGalleryName() {
  const context = getActiveTabDomainContext();
  if (!context.isValid) {
    return '';
  }
  const href = String(context.href || '').trim();
  if (!href) {
    return '';
  }
  return href.length > 512 ? href.slice(0, 512) : href;
}

function resolveGalleryDomainContextFromActiveTab() {
  const context = getActiveTabDomainContext();
  if (context.isValid && context.domain) {
    return {
      domain: context.domain,
      pageOrigin: context.pageOrigin
    };
  }

  let domain = 'manual.local';
  let pageOrigin = '';
  try {
    const parsed = new URL(String(state.activeTabUrl || '').trim());
    domain = String(parsed.hostname || domain).trim().toLowerCase() || domain;
    pageOrigin = String(parsed.origin || '').trim();
  } catch {
    // keep fallback values
  }
  return { domain, pageOrigin };
}

function buildCustomGalleryId(name) {
  const safeName = sanitizeGalleryName(name || '', 'galeria').toLowerCase();
  return `custom:${encodeURIComponent(safeName)}`;
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

async function removeGalleryById(galleryId) {
  const safeGalleryId = String(galleryId || '').trim();
  if (!safeGalleryId) {
    return;
  }

  if (safeGalleryId === MANUAL_GALLERY_ID) {
    if (state.manualItems.length === 0) {
      return;
    }
    state.manualItems = [];
    state.selectedManualIds = new Set();
    await persistManualLibrary();
    if (String(state.activeGalleryId || '').trim() === MANUAL_GALLERY_ID) {
      state.galleryListMode = 'list';
      state.activeGalleryId = '';
    }
    renderActiveLogoLibraryGrid();
    showFeedbackToast('Galeria eliminada');
    return;
  }

  const before = state.savedGalleries.length;
  state.savedGalleries = state.savedGalleries.filter(
    (entry) => String(entry?.id || '').trim() !== safeGalleryId
  );
  if (state.savedGalleries.length === before) {
    return;
  }

  if (String(state.activeGalleryId || '').trim() === safeGalleryId) {
    state.galleryListMode = 'list';
    state.activeGalleryId = '';
  }

  await persistDomainGalleries();
  renderActiveLogoLibraryGrid();
  showFeedbackToast('Galeria eliminada');
}

async function removeStoredGalleryImageById(galleryId, imageId) {
  const safeGalleryId = String(galleryId || '').trim();
  const safeImageId = String(imageId || '').trim();
  if (!safeGalleryId || !safeImageId) {
    return;
  }

  if (safeGalleryId === MANUAL_GALLERY_ID) {
    await removeManualImageById(safeImageId);
    return;
  }

  const nextGalleries = [];
  let removed = false;
  const now = Date.now();

  for (const gallery of state.savedGalleries) {
    const currentId = String(gallery?.id || '').trim();
    if (currentId !== safeGalleryId) {
      nextGalleries.push(gallery);
      continue;
    }

    const items = Array.isArray(gallery.items) ? gallery.items : [];
    const nextItems = items.filter((item) => String(item?.id || '').trim() !== safeImageId);
    if (nextItems.length === items.length) {
      nextGalleries.push(gallery);
      continue;
    }

    removed = true;
    nextGalleries.push({
      ...gallery,
      updatedAt: now,
      items: nextItems
    });
  }

  if (!removed) {
    return;
  }

  state.savedGalleries = nextGalleries.sort((a, b) => Number(b.updatedAt) - Number(a.updatedAt));
  await persistDomainGalleries();
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
    const name = sanitizeGalleryName(entry.name || '', domain || 'Galeria');
    const pageOrigin = String(entry.pageOrigin || '').trim();
    const faviconUrl = String(entry.faviconUrl || '').trim();
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
      name,
      domain,
      pageOrigin,
      faviconUrl: faviconUrl || buildGalleryFaviconUrl(pageOrigin, domain),
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

function sanitizeGalleryName(name, fallback = 'Galeria') {
  const safe = String(name || '').trim();
  const safeFallback = String(fallback || 'Galeria').trim() || 'Galeria';
  if (!safe) {
    return safeFallback;
  }
  return safe.length > 512 ? safe.slice(0, 512) : safe;
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
