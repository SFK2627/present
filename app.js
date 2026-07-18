/* Presentation Hub static app
   - Open index.html directly or upload the folder to any static host.
   - PDF rendering uses PDF.js from CDN with high-DPI fullscreen rendering.
   - PPTX visual mode uses PPTXjs when online. For pixel-perfect PowerPoint output, export PPTX as PDF and upload the PDF.
   - Phone remote uses Firebase when firebase-config.js is configured.
*/

(function () {
  'use strict';

  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }

  const $ = (id) => document.getElementById(id);
  const qs = new URLSearchParams(window.location.search);
  const isRemoteMode = qs.get('remote') === '1';

  const DB_NAME = 'presentation-hub-static';
  const DB_VERSION = 1;
  const STORE = 'presentations';
  const SESSION_COLLECTION = 'presentationHubSessions';

  const state = {
    files: [],
    folders: [],
    activeFolderId: localStorage.getItem('presentationHubActiveFolder') || 'all',
    activeFile: null,
    activePdf: null,
    activePptxSlides: [],
    pptxVisualReady: false,
    pptxRenderedSlides: [],
    pptxBlobUrl: '',
    currentPage: 1,
    totalPages: 1,
    zoom: 1,
    viewportCenterX: 0.5,
    viewportCenterY: 0.5,
    timingMode: 'global',
    perPageTiming: {},
    autoTimer: null,
    autoPlaying: false,
    autoPaused: false,
    autoStartedAt: null,
    autoElapsedBeforePause: 0,
    autoCurrentDuration: 10,
    autoAdvancing: false,
    countdownAlert: 'off',
    countdownVoiceGender: 'soft-female',
    countdownVoiceStart: 5,
    countdownPreviewTimers: [],
    lastCountdownAlertSecond: null,
    audioContext: null,
    audioUnlocked: false,
    transitionEffect: 'fade',
    transitionDuration: 420,
    slideChangePending: false,
    pdfPageCache: new Map(),
    pdfBitmapCache: new Map(),
    activePdfCanvas: null,
    pdfRenderQueueKey: '',
    timerTick: null,
    timerStartedAt: null,
    timerElapsedBeforePause: 0,
    timer: {
      visible: false,
      mode: 'up',
      countdownSeconds: 600,
      position: 'bottom-right',
      opacity: 75,
      size: 28,
    },
    toolbarHideTimer: null,
    renderToken: 0,
    activePdfRenderTask: null,
    lastRemoteThumbKey: '',
    remotePreviewBusy: false,
    remoteThumbTimer: null,
    viewportQualityTimer: null,
    firebaseReady: false,
    firebaseDb: null,
    firebaseAuthReady: false,
    sessionId: null,
    sessionRef: null,
    unsubscribeSession: null,
    lastCommandId: null,
    publishLock: false,
    inkStrokes: [],
    remoteSlideThumbs: {},
    remoteSlideThumbsCount: 0,
    remoteSlideThumbsBusy: false,
  };



  const COUNTDOWN_VOICE_STYLES = {
    'soft-female': {
      label: 'Soft Female',
      hints: ['female', 'woman', 'samantha', 'zira', 'victoria', 'karen', 'moira', 'tessa', 'susan', 'allison', 'ava', 'aria', 'jenny'],
      pitch: 1.18,
      rate: 0.95,
    },
    'bright-female': {
      label: 'Bright Female',
      hints: ['female', 'woman', 'jenny', 'aria', 'ava', 'samantha', 'zira', 'victoria', 'allison', 'susan'],
      pitch: 1.32,
      rate: 1.08,
    },
    'calm-male': {
      label: 'Calm Male',
      hints: ['male', 'man', 'david', 'mark', 'alex', 'daniel', 'tom', 'fred', 'guy', 'ryan', 'george'],
      pitch: 0.82,
      rate: 0.94,
    },
    'deep-male': {
      label: 'Deep Male',
      hints: ['male', 'man', 'david', 'mark', 'daniel', 'george', 'guy', 'ryan', 'alex'],
      pitch: 0.68,
      rate: 0.9,
    },
    teacher: {
      label: 'Teacher Voice',
      hints: ['female', 'woman', 'samantha', 'zira', 'jenny', 'aria', 'victoria', 'karen', 'moira', 'ava', 'alex'],
      pitch: 1.05,
      rate: 0.92,
    },
    announcer: {
      label: 'Announcer Voice',
      hints: ['male', 'man', 'david', 'mark', 'daniel', 'george', 'ryan', 'alex', 'guy'],
      pitch: 0.9,
      rate: 1.02,
    },
  };

  function normalizeCountdownVoiceStyle(value) {
    const key = String(value || '').trim().toLowerCase();
    if (COUNTDOWN_VOICE_STYLES[key]) return key;
    if (key === 'boy' || key === 'male') return 'calm-male';
    if (key === 'girl' || key === 'female') return 'soft-female';
    return 'soft-female';
  }

  function getCountdownVoiceProfile() {
    const key = normalizeCountdownVoiceStyle(state.countdownVoiceGender);
    return COUNTDOWN_VOICE_STYLES[key] || COUNTDOWN_VOICE_STYLES['soft-female'];
  }
  const els = {};

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    if (isRemoteMode) {
      document.documentElement.classList.add('presentation-hub-remote-mode');
      document.body.classList.add('presentation-hub-remote-mode');
      document.body.classList.remove('remote-fullscreen-open');
    } else {
      document.documentElement.classList.remove('presentation-hub-remote-mode');
      document.body.classList.remove('presentation-hub-remote-mode', 'remote-fullscreen-open');
    }
    cacheElements();
    applySavedTheme();
    setupBaseEvents();

    // Do not block the dashboard while Firebase signs in. The viewer can open fast,
    // then the remote status updates as soon as Firebase is ready.
    const firebaseInitPromise = initFirebaseIfConfigured();

    if (isRemoteMode) {
      await firebaseInitPromise;
      renderRemoteApp();
      return;
    }

    await loadLibrary();
    loadFolders();
    renderFolderList();
    renderLibrary();
    registerServiceWorker();
    window.addEventListener('resize', () => {
      if (!state.activeFile) return;
      if (state.activeFile.type === 'pptx' && state.pptxVisualReady) showOnlyCurrentPptxSlide();
      if (state.activeFile.type === 'pdf') renderCurrentPage();
    });
  }

  function cacheElements() {
    [
      'app', 'remoteApp', 'homeView', 'viewerView', 'fileInput', 'uploadZone', 'searchInput', 'sortSelect',
      'cardsGrid', 'emptyState', 'libraryCount', 'clearLibraryBtn', 'themeToggle', 'firebaseStatus',
      'createFolderBtn', 'folderList', 'canvaTitleInput', 'canvaLinkInput', 'addCanvaBtn',
      'thumbnailSidebar', 'viewerStage', 'viewerToolbar', 'controlPanel', 'settingsBtn', 'backHomeBtn',
      'prevBtn', 'nextBtn', 'jumpInput', 'pageTotalLabel', 'zoomOutBtn', 'zoomInBtn', 'resetZoomBtn',
      'zoomLabel', 'fullscreenBtn', 'qrBtn', 'viewerCanvasWrap', 'pdfCanvas', 'pptxSlide', 'canvaFrame', 'inkCanvas',
      'timingModeSelect', 'globalTimingSelect', 'customTimingWrap', 'customTimingInput', 'perSlideTimingWrap', 'perSlideTimingInput',
      'countdownAlertSelect', 'countdownVoiceSelect', 'countdownVoiceStartSelect', 'soundTestBtn', 'slideTransitionSelect',
      'autoStartBtn', 'autoPauseBtn', 'autoResumeBtn', 'autoStopBtn', 'timerOverlay', 'timerModeSelect',
      'countdownMinutesInput', 'timerPositionSelect', 'timerOpacityInput', 'timerSizeInput', 'timerSizeLabel', 'timerShowBtn', 'timerHideBtn',
      'timerResetBtn', 'qrModal', 'closeQrBtn', 'hostQr', 'viewerQr', 'hostRemoteLink', 'viewerRemoteLink',
      'qrHelp', 'setupModal', 'closeSetupBtn'
    ].forEach((id) => { els[id] = $(id); });
  }

  function setupBaseEvents() {
    if (els.fileInput) els.fileInput.addEventListener('change', (event) => handleFiles(event.target.files));
    if (els.uploadZone) {
      ['dragenter', 'dragover'].forEach((type) => els.uploadZone.addEventListener(type, (event) => {
        event.preventDefault();
        els.uploadZone.classList.add('drag-over');
      }));
      ['dragleave', 'drop'].forEach((type) => els.uploadZone.addEventListener(type, (event) => {
        event.preventDefault();
        els.uploadZone.classList.remove('drag-over');
      }));
      els.uploadZone.addEventListener('drop', (event) => handleFiles(event.dataTransfer.files));
    }

    if (els.searchInput) els.searchInput.addEventListener('input', renderLibrary);
    if (els.sortSelect) els.sortSelect.addEventListener('change', renderLibrary);
    if (els.clearLibraryBtn) els.clearLibraryBtn.addEventListener('click', clearLibrary);
    if (els.createFolderBtn) els.createFolderBtn.addEventListener('click', createFolder);
    if (els.addCanvaBtn) els.addCanvaBtn.addEventListener('click', addCanvaLink);
    if (els.canvaLinkInput) els.canvaLinkInput.addEventListener('keydown', (event) => { if (event.key === 'Enter') addCanvaLink(); });
    if (els.themeToggle) els.themeToggle.addEventListener('click', toggleTheme);
    if (els.firebaseStatus) els.firebaseStatus.addEventListener('click', () => showModal(els.setupModal));
    if (els.closeSetupBtn) els.closeSetupBtn.addEventListener('click', () => hideModal(els.setupModal));

    if (els.backHomeBtn) els.backHomeBtn.addEventListener('click', closeViewer);
    if (els.prevBtn) els.prevBtn.addEventListener('click', previousPage);
    if (els.nextBtn) els.nextBtn.addEventListener('click', nextPage);
    if (els.jumpInput) els.jumpInput.addEventListener('change', () => jumpToPage(Number(els.jumpInput.value)));
    if (els.zoomInBtn) els.zoomInBtn.addEventListener('click', () => setZoom(state.zoom + 0.1));
    if (els.zoomOutBtn) els.zoomOutBtn.addEventListener('click', () => setZoom(state.zoom - 0.1));
    if (els.resetZoomBtn) els.resetZoomBtn.addEventListener('click', () => setZoom(1));
    if (els.fullscreenBtn) els.fullscreenBtn.addEventListener('click', toggleFullscreen);
    if (els.qrBtn) els.qrBtn.addEventListener('click', openQrModal);
    if (els.closeQrBtn) els.closeQrBtn.addEventListener('click', () => hideModal(els.qrModal));
    if (els.settingsBtn) els.settingsBtn.addEventListener('click', () => els.controlPanel.classList.toggle('hidden'));

    // Browsers only allow presentation sounds after a user gesture. Prime audio early so
    // the last-5-second alert still works later, including when autoplay is started by phone.
    document.addEventListener('pointerdown', primePresentationAudio, { passive: true });
    document.addEventListener('keydown', primePresentationAudio, true);

    if (els.timingModeSelect) els.timingModeSelect.addEventListener('change', () => {
      state.timingMode = els.timingModeSelect.value === 'per-slide' ? 'per-slide' : 'global';
      updateTimingModeUI();
      resetAutoClockForCurrentSlide();
      renderThumbnailSidebar();
      updateTimerText();
      publishSessionState();
    });
    if (els.globalTimingSelect) els.globalTimingSelect.addEventListener('change', () => {
      els.customTimingWrap.classList.toggle('hidden', els.globalTimingSelect.value !== 'custom');
      resetAutoClockForCurrentSlide();
      updateTimerText();
      publishSessionState();
    });
    if (els.customTimingInput) els.customTimingInput.addEventListener('input', () => {
      resetAutoClockForCurrentSlide();
      updateTimerText();
      publishSessionState();
    });
    if (els.perSlideTimingInput) els.perSlideTimingInput.addEventListener('change', () => {
      const val = Number(els.perSlideTimingInput.value);
      if (val > 0) state.perPageTiming[state.currentPage] = val;
      else delete state.perPageTiming[state.currentPage];
      state.timingMode = 'per-slide';
      if (els.timingModeSelect) els.timingModeSelect.value = 'per-slide';
      updateTimingModeUI();
      resetAutoClockForCurrentSlide();
      renderThumbnailSidebar();
      updateTimerText();
      publishSessionState();
    });
    if (els.countdownAlertSelect) els.countdownAlertSelect.addEventListener('change', () => {
      state.countdownAlert = els.countdownAlertSelect.value || 'off';
      state.lastCountdownAlertSecond = null;
      if (state.countdownAlert !== 'off') primePresentationAudio();
      publishSessionState();
    });
    if (els.countdownVoiceSelect) els.countdownVoiceSelect.addEventListener('change', () => {
      state.countdownVoiceGender = normalizeCountdownVoiceStyle(els.countdownVoiceSelect.value);
      primePresentationAudio();
      publishSessionState();
    });
    if (els.countdownVoiceStartSelect) els.countdownVoiceStartSelect.addEventListener('change', () => {
      state.countdownVoiceStart = Number(els.countdownVoiceStartSelect.value) === 3 ? 3 : 5;
      state.lastCountdownAlertSecond = null;
      primePresentationAudio();
      publishSessionState();
    });
    if (els.soundTestBtn) els.soundTestBtn.addEventListener('click', () => {
      const mode = state.countdownAlert === 'off' ? 'both' : state.countdownAlert;
      primePresentationAudio();
      runCountdownAlertPreview(mode);
    });
    if (els.slideTransitionSelect) els.slideTransitionSelect.addEventListener('change', () => {
      state.transitionEffect = els.slideTransitionSelect.value || 'fade';
      publishSessionState();
    });

    if (els.autoStartBtn) els.autoStartBtn.addEventListener('click', startAutoPlay);
    if (els.autoPauseBtn) els.autoPauseBtn.addEventListener('click', pauseAutoPlay);
    if (els.autoResumeBtn) els.autoResumeBtn.addEventListener('click', resumeAutoPlay);
    if (els.autoStopBtn) els.autoStopBtn.addEventListener('click', stopAutoPlay);

    if (els.timerModeSelect) els.timerModeSelect.addEventListener('change', () => {
      state.timer.mode = els.timerModeSelect.value;
      resetTimer();
      publishSessionState();
    });
    if (els.countdownMinutesInput) els.countdownMinutesInput.addEventListener('change', () => {
      state.timer.countdownSeconds = Math.max(1, Number(els.countdownMinutesInput.value) || 10) * 60;
      resetTimer();
      publishSessionState();
    });
    if (els.timerPositionSelect) els.timerPositionSelect.addEventListener('change', () => {
      state.timer.position = els.timerPositionSelect.value;
      applyTimerSettings();
      publishSessionState();
    });
    if (els.timerOpacityInput) els.timerOpacityInput.addEventListener('input', () => {
      state.timer.opacity = Number(els.timerOpacityInput.value);
      applyTimerSettings();
      publishSessionState();
    });
    if (els.timerSizeInput) els.timerSizeInput.addEventListener('input', () => {
      state.timer.size = Math.max(16, Math.min(72, Number(els.timerSizeInput.value) || 28));
      applyTimerSettings();
      publishSessionState();
    });
    if (els.timerShowBtn) els.timerShowBtn.addEventListener('click', showTimer);
    if (els.timerHideBtn) els.timerHideBtn.addEventListener('click', hideTimer);
    if (els.timerResetBtn) els.timerResetBtn.addEventListener('click', resetTimer);

    document.addEventListener('keydown', handleKeyboard, true);
    window.addEventListener('keydown', handleKeyboard, true);
    document.addEventListener('fullscreenchange', syncFullscreenState);
    if (els.viewerStage) {
      els.viewerStage.addEventListener('mousemove', revealToolbarTemporarily);
      els.viewerStage.addEventListener('touchstart', revealToolbarTemporarily, { passive: true });
    }
    if (els.viewerView) {
      els.viewerView.addEventListener('mousemove', revealToolbarTemporarily);
      els.viewerView.addEventListener('touchstart', revealToolbarTemporarily, { passive: true });
    }
  }

  function applySavedTheme() {
    const saved = localStorage.getItem('presentationHubTheme');
    const dark = saved ? saved === 'dark' : window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    document.body.classList.toggle('dark', dark);
    if (els.themeToggle) els.themeToggle.textContent = dark ? '☀️' : '🌙';
  }

  function toggleTheme() {
    const dark = !document.body.classList.contains('dark');
    document.body.classList.toggle('dark', dark);
    localStorage.setItem('presentationHubTheme', dark ? 'dark' : 'light');
    els.themeToggle.textContent = dark ? '☀️' : '🌙';
  }

  function hasFirebaseConfig() {
    const cfg = window.PRESENTATION_HUB_FIREBASE_CONFIG || {};
    return Boolean(cfg.apiKey && cfg.projectId && cfg.appId);
  }

  async function initFirebaseIfConfigured() {
    if (!hasFirebaseConfig() || !window.firebase) {
      state.firebaseReady = false;
      updateFirebaseStatus();
      return false;
    }
    try {
      if (!firebase.apps.length) firebase.initializeApp(window.PRESENTATION_HUB_FIREBASE_CONFIG);
      state.firebaseDb = firebase.firestore();
      await firebase.auth().signInAnonymously();
      state.firebaseAuthReady = true;
      state.firebaseReady = true;
      updateFirebaseStatus();
      return true;
    } catch (error) {
      console.warn('Firebase setup failed:', error);
      state.firebaseReady = false;
      updateFirebaseStatus('Remote offline');
      return false;
    }
  }

  function updateFirebaseStatus(customText) {
    if (!els.firebaseStatus) return;
    if (state.firebaseReady) {
      els.firebaseStatus.textContent = 'Remote ready';
      els.firebaseStatus.classList.add('ready');
      els.firebaseStatus.classList.remove('muted');
    } else {
      els.firebaseStatus.textContent = customText || 'Remote setup';
      els.firebaseStatus.classList.remove('ready');
      els.firebaseStatus.classList.add('muted');
    }
  }

  function showModal(modal) {
    if (!modal) return;
    const root = document.fullscreenElement || document.body;
    if (modal.parentElement !== root) root.appendChild(modal);
    modal.classList.remove('hidden');
    revealToolbarTemporarily();
  }

  function hideModal(modal) {
    if (!modal) return;
    modal.classList.add('hidden');
    if (els.app && modal.parentElement !== els.app) els.app.appendChild(modal);
  }

  function loadScriptOnce(src) {
    return new Promise((resolve, reject) => {
      const existing = document.querySelector(`script[data-src="${src}"]`);
      if (existing) {
        if (existing.dataset.loaded === '1') resolve();
        else existing.addEventListener('load', resolve, { once: true });
        return;
      }
      const script = document.createElement('script');
      script.src = src;
      script.dataset.src = src;
      script.async = false;
      script.onload = () => { script.dataset.loaded = '1'; resolve(); };
      script.onerror = () => reject(new Error(`Failed to load ${src}`));
      document.head.appendChild(script);
    });
  }

  function loadStyleOnce(href) {
    if (document.querySelector(`link[data-href="${href}"]`)) return;
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = href;
    link.dataset.href = href;
    document.head.appendChild(link);
  }

  async function ensurePptxRendererAssets() {
    loadStyleOnce('https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@master/css/pptxjs.css');
    loadStyleOnce('https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@master/css/nv.d3.min.css');

    if (!window.jQuery) {
      await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.7.1/jquery.min.js');
    }
    if (!window.d3) {
      await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/d3/3.5.17/d3.min.js');
    }
    if (!window.nv) {
      await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/nvd3/1.8.6/nv.d3.min.js');
    }
    await loadScriptOnce('https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@master/filereader.js');
    await loadScriptOnce('https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@master/js/dingbat.js');
    await loadScriptOnce('https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@master/js/pptxjs.js');
    await loadScriptOnce('https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@master/js/divs2slides.js');
    if (!window.html2canvas) {
      await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
    }
  }

  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: 'id' });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function dbPut(record) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(record);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function dbGetAll() {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const request = tx.objectStore(STORE).getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  }

  async function dbDelete(id) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(id);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function dbClear() {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function loadLibrary() {
    try {
      state.files = await dbGetAll();
    } catch (error) {
      console.warn(error);
      state.files = [];
    }
  }


  function loadFolders() {
    try {
      const raw = localStorage.getItem('presentationHubFolders');
      state.folders = raw ? JSON.parse(raw).filter((folder) => folder && folder.id && folder.name) : [];
    } catch (error) {
      state.folders = [];
    }
    if (state.activeFolderId !== 'all' && !state.folders.some((folder) => folder.id === state.activeFolderId)) {
      state.activeFolderId = 'all';
      localStorage.setItem('presentationHubActiveFolder', 'all');
    }
  }

  function saveFolders() {
    localStorage.setItem('presentationHubFolders', JSON.stringify(state.folders));
    localStorage.setItem('presentationHubActiveFolder', state.activeFolderId || 'all');
  }

  function getFolderName(folderId) {
    if (!folderId) return 'Unfiled';
    const folder = state.folders.find((item) => item.id === folderId);
    return folder ? folder.name : 'Unfiled';
  }

  function createFolder() {
    const name = prompt('Folder name:');
    if (!name || !name.trim()) return;
    const folder = {
      id: crypto.randomUUID ? crypto.randomUUID() : `folder-${Date.now()}`,
      name: name.trim(),
      createdAt: new Date().toISOString(),
    };
    state.folders.push(folder);
    state.activeFolderId = folder.id;
    saveFolders();
    renderFolderList();
    renderLibrary();
  }

  function renameFolder(folderId) {
    const folder = state.folders.find((item) => item.id === folderId);
    if (!folder) return;
    const name = prompt('Rename folder:', folder.name);
    if (!name || !name.trim()) return;
    folder.name = name.trim();
    saveFolders();
    renderFolderList();
    renderLibrary();
  }

  async function deleteFolder(folderId) {
    const folder = state.folders.find((item) => item.id === folderId);
    if (!folder) return;
    if (!confirm(`Delete folder "${folder.name}"? Files will stay in Unfiled.`)) return;
    state.folders = state.folders.filter((item) => item.id !== folderId);
    state.files.forEach((file) => {
      if (file.folderId === folderId) file.folderId = '';
    });
    await Promise.all(state.files.map((file) => dbPut(file)));
    if (state.activeFolderId === folderId) state.activeFolderId = 'all';
    saveFolders();
    renderFolderList();
    renderLibrary();
  }

  function setActiveFolder(folderId) {
    state.activeFolderId = folderId || 'all';
    saveFolders();
    renderFolderList();
    renderLibrary();
  }

  function getFolderCounts() {
    const counts = { all: state.files.length, unfiled: state.files.filter((file) => !file.folderId).length };
    state.folders.forEach((folder) => {
      counts[folder.id] = state.files.filter((file) => file.folderId === folder.id).length;
    });
    return counts;
  }

  function renderFolderList() {
    if (!els.folderList) return;
    const counts = getFolderCounts();
    const chips = [
      `<button class="folder-chip ${state.activeFolderId === 'all' ? 'active' : ''}" data-folder="all"><span>All</span><b>${counts.all || 0}</b></button>`,
      `<button class="folder-chip ${state.activeFolderId === 'unfiled' ? 'active' : ''}" data-folder="unfiled" data-folder-drop="unfiled" title="Drop files here to remove from folders"><span>Unfiled</span><b>${counts.unfiled || 0}</b></button>`,
      ...state.folders.map((folder) => `
        <div class="folder-chip-wrap ${state.activeFolderId === folder.id ? 'active' : ''}" data-folder-drop="${folder.id}" title="Drop presentations here">
          <button class="folder-chip folder-main ${state.activeFolderId === folder.id ? 'active' : ''}" data-folder="${folder.id}">
            <span>${escapeHtml(folder.name)}</span><b>${counts[folder.id] || 0}</b>
          </button>
          <button class="folder-mini-btn" data-rename-folder="${folder.id}" title="Rename folder">✎</button>
          <button class="folder-mini-btn danger" data-delete-folder="${folder.id}" title="Delete folder">×</button>
        </div>`),
    ];
    els.folderList.innerHTML = chips.join('');
    els.folderList.querySelectorAll('[data-folder]').forEach((button) => {
      button.addEventListener('click', () => setActiveFolder(button.dataset.folder));
    });
    els.folderList.querySelectorAll('[data-rename-folder]').forEach((button) => {
      button.addEventListener('click', () => renameFolder(button.dataset.renameFolder));
    });
    els.folderList.querySelectorAll('[data-delete-folder]').forEach((button) => {
      button.addEventListener('click', () => deleteFolder(button.dataset.deleteFolder));
    });
    setupFolderDropTargets();
  }

  function resolveDropFolderId(dropValue) {
    if (dropValue === 'unfiled') return '';
    return state.folders.some((folder) => folder.id === dropValue) ? dropValue : null;
  }

  function setupFolderDropTargets() {
    if (!els.folderList) return;
    els.folderList.querySelectorAll('[data-folder-drop]').forEach((target) => {
      const clearDropState = () => target.classList.remove('drop-ready');
      target.addEventListener('dragenter', (event) => {
        if (!isSupportedFolderDrop(event)) return;
        event.preventDefault();
        target.classList.add('drop-ready');
      });
      target.addEventListener('dragover', (event) => {
        if (!isSupportedFolderDrop(event)) return;
        event.preventDefault();
        event.dataTransfer.dropEffect = 'move';
        target.classList.add('drop-ready');
      });
      target.addEventListener('dragleave', (event) => {
        if (!target.contains(event.relatedTarget)) clearDropState();
      });
      target.addEventListener('drop', async (event) => {
        if (!isSupportedFolderDrop(event)) return;
        event.preventDefault();
        event.stopPropagation();
        clearDropState();
        const folderId = resolveDropFolderId(target.dataset.folderDrop);
        if (folderId === null) return;

        const droppedFiles = event.dataTransfer.files;
        if (droppedFiles && droppedFiles.length) {
          await handleFiles(droppedFiles, folderId);
          return;
        }

        const presentationId = event.dataTransfer.getData('application/x-presentation-id') || event.dataTransfer.getData('text/plain');
        if (presentationId) await movePresentationToFolder(presentationId, folderId);
      });
    });
  }

  function isSupportedFolderDrop(event) {
    const types = Array.from(event.dataTransfer ? event.dataTransfer.types || [] : []);
    return types.includes('Files') || types.includes('application/x-presentation-id') || types.includes('text/plain');
  }

  function getUploadFolderId() {
    return state.activeFolderId && !['all', 'unfiled'].includes(state.activeFolderId) ? state.activeFolderId : '';
  }

  function folderSelectOptions(selectedId) {
    const base = [`<option value="">Unfiled</option>`];
    state.folders.forEach((folder) => {
      base.push(`<option value="${folder.id}" ${selectedId === folder.id ? 'selected' : ''}>${escapeHtml(folder.name)}</option>`);
    });
    return base.join('');
  }

  async function movePresentationToFolder(id, folderId) {
    const file = state.files.find((item) => item.id === id);
    if (!file) return;
    file.folderId = folderId || '';
    await dbPut(file);
    renderFolderList();
    renderLibrary();
  }

  function decodeLooseHtml(value) {
    const textarea = document.createElement('textarea');
    textarea.innerHTML = value;
    return textarea.value;
  }

  function normalizeCanvaUrl(raw) {
    let value = decodeLooseHtml(String(raw || '').trim());
    if (!value) return null;

    // Accept full Canva iframe/embed snippets, normal share links, links copied with
    // extra text around them, and links where the browser/app escaped ampersands.
    const srcMatch = value.match(/src\s*=\s*["']([^"']+)["']/i);
    if (srcMatch) value = srcMatch[1];
    else {
      const urlMatch = value.match(/https?:\/\/[^\s<>"']+|(?:www\.)?canva\.(?:com|site|cn|me)\/[^\s<>"']+/i);
      if (urlMatch) value = urlMatch[0];
    }

    value = decodeLooseHtml(value)
      .replace(/&amp;/gi, '&')
      .trim()
      .replace(/[),.;]+$/g, '');

    if (!/^https?:\/\//i.test(value)) value = `https://${value}`;

    let url;
    try { url = new URL(value); } catch (error) { return null; }

    const host = url.hostname.toLowerCase().replace(/^www\./, '');
    const allowed = host === 'canva.com' || host.endsWith('.canva.com') ||
      host === 'canva.site' || host.endsWith('.canva.site') ||
      host === 'canva.cn' || host.endsWith('.canva.cn') ||
      host === 'canva.me' || host.endsWith('.canva.me');
    if (!allowed) return null;

    // Canva's public design/view links usually embed when ?embed is present.
    // Keep public canva.site pages as-is because they are already designed for viewing.
    if ((host === 'canva.com' || host.endsWith('.canva.com') || host === 'canva.cn' || host.endsWith('.canva.cn')) && !url.searchParams.has('embed')) {
      url.searchParams.set('embed', '');
    }
    return url.toString();
  }

  async function addCanvaLink() {
    const raw = els.canvaLinkInput ? els.canvaLinkInput.value : '';
    const embedUrl = normalizeCanvaUrl(raw);
    if (!embedUrl) {
      alert('Please paste a valid Canva public view or embed link.');
      return;
    }
    const id = crypto.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random());
    const now = new Date().toISOString();
    let title = els.canvaTitleInput && els.canvaTitleInput.value.trim();
    if (!title) title = 'Canva presentation';
    const record = {
      id,
      name: title,
      type: 'canva',
      url: raw.trim(),
      embedUrl,
      uploadedAt: now,
      lastViewed: now,
      pageCount: 1,
      folderId: getUploadFolderId(),
      thumbnail: createCanvaThumbDataUrl(title),
      note: 'Canva links are embedded in the viewer. Use a public view or embed link; private Canva links may require login.',
    };
    state.files.unshift(record);
    await dbPut(record);
    if (els.canvaTitleInput) els.canvaTitleInput.value = '';
    if (els.canvaLinkInput) els.canvaLinkInput.value = '';
    renderFolderList();
    renderLibrary();
  }

  function createCanvaThumbDataUrl(name) {
    const canvas = document.createElement('canvas');
    canvas.width = 640;
    canvas.height = 400;
    const ctx = canvas.getContext('2d');
    const gradient = ctx.createLinearGradient(0, 0, 640, 400);
    gradient.addColorStop(0, '#00c4cc');
    gradient.addColorStop(0.55, '#7d2ae8');
    gradient.addColorStop(1, '#ff66c4');
    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, 640, 400);
    ctx.fillStyle = 'rgba(255,255,255,.18)';
    roundRect(ctx, 54, 54, 532, 292, 30);
    ctx.fill();
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 48px Inter, Arial';
    ctx.fillText('CANVA', 86, 150);
    ctx.font = '26px Inter, Arial';
    ctx.fillText('Embedded presentation link', 86, 198);
    ctx.font = '22px Inter, Arial';
    wrapCanvasText(ctx, name, 86, 255, 460, 30, 2);
    return canvas.toDataURL('image/jpeg', 0.86);
  }

  async function handleFiles(fileList, folderIdOverride) {
    const files = Array.from(fileList || []).filter((file) => /\.(pdf|pptx)$/i.test(file.name));
    if (!files.length) return;

    for (const file of files) {
      const record = await createPresentationRecord(file, typeof folderIdOverride === 'string' ? folderIdOverride : getUploadFolderId());
      state.files.unshift(record);
      await dbPut(record);
      renderLibrary();
    }
    if (els.fileInput) els.fileInput.value = '';
  }

  async function createPresentationRecord(file, folderId = '') {
    const id = crypto.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random());
    const now = new Date().toISOString();
    const type = file.name.toLowerCase().endsWith('.pdf') ? 'pdf' : 'pptx';
    let pageCount = 1;
    let thumbnail = '';
    let note = '';

    try {
      const buffer = await file.arrayBuffer();
      if (type === 'pdf') {
        const pdf = await pdfjsLib.getDocument({ data: buffer.slice(0) }).promise;
        pageCount = pdf.numPages;
        thumbnail = await renderPdfPageToDataUrl(pdf, 1, 0.22);
        if (pdf.destroy) pdf.destroy();
      } else {
        const result = await inspectPptx(buffer.slice(0));
        pageCount = result.count;
        thumbnail = createPptxThumbDataUrl(file.name, pageCount);
        note = 'PowerPoint opens with the visual renderer when possible. For exact layout, export PPTX as PDF and upload the PDF.';
      }
    } catch (error) {
      console.warn('Could not inspect file:', error);
      thumbnail = type === 'pptx' ? createPptxThumbDataUrl(file.name, pageCount) : createGenericThumbDataUrl(file.name, type);
    }

    return {
      id,
      name: file.name,
      type,
      uploadedAt: now,
      lastViewed: now,
      pageCount,
      thumbnail,
      folderId,
      blob: file,
      note,
    };
  }

  async function inspectPptx(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const slideFiles = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort((a, b) => slideNumber(a) - slideNumber(b));
    return { count: Math.max(1, slideFiles.length), slideFiles };
  }

  function slideNumber(path) {
    const match = path.match(/slide(\d+)\.xml/i);
    return match ? Number(match[1]) : 0;
  }

  async function parsePptxSlides(blob) {
    const buffer = await blob.arrayBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const slideFiles = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort((a, b) => slideNumber(a) - slideNumber(b));

    const slides = [];
    for (const fileName of slideFiles) {
      const xml = await zip.file(fileName).async('string');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const texts = Array.from(doc.getElementsByTagName('a:t')).map((node) => node.textContent.trim()).filter(Boolean);
      slides.push({
        title: texts[0] || `Slide ${slides.length + 1}`,
        lines: texts.slice(1, 10),
      });
    }
    return slides.length ? slides : [{ title: 'PowerPoint file', lines: ['No extractable slide text found. Convert this PPTX to PDF for exact viewing.'] }];
  }

  function createGenericThumbDataUrl(name, type) {
    const canvas = document.createElement('canvas');
    canvas.width = 640;
    canvas.height = 400;
    const ctx = canvas.getContext('2d');
    const gradient = ctx.createLinearGradient(0, 0, 640, 400);
    gradient.addColorStop(0, '#6557ff');
    gradient.addColorStop(1, '#00b7c7');
    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, 640, 400);
    ctx.fillStyle = 'rgba(255,255,255,.16)';
    roundRect(ctx, 48, 58, 544, 284, 28);
    ctx.fill();
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 54px Inter, Arial';
    ctx.fillText(type.toUpperCase(), 70, 170);
    ctx.font = '28px Inter, Arial';
    wrapCanvasText(ctx, name, 70, 230, 500, 34, 3);
    return canvas.toDataURL('image/jpeg', 0.82);
  }

  function createPptxThumbDataUrl(name, count) {
    const canvas = document.createElement('canvas');
    canvas.width = 640;
    canvas.height = 400;
    const ctx = canvas.getContext('2d');
    const gradient = ctx.createLinearGradient(0, 0, 640, 400);
    gradient.addColorStop(0, '#fb923c');
    gradient.addColorStop(1, '#6557ff');
    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, 640, 400);
    ctx.fillStyle = 'rgba(255,255,255,.92)';
    roundRect(ctx, 82, 64, 476, 272, 28);
    ctx.fill();
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 42px Inter, Arial';
    ctx.fillText('PPTX', 116, 150);
    ctx.font = '24px Inter, Arial';
    ctx.fillText(`${count} slide${count === 1 ? '' : 's'}`, 116, 194);
    ctx.fillStyle = '#475569';
    ctx.font = '22px Inter, Arial';
    wrapCanvasText(ctx, name, 116, 245, 400, 28, 2);
    return canvas.toDataURL('image/jpeg', 0.84);
  }

  function roundRect(ctx, x, y, width, height, radius) {
    ctx.beginPath();
    ctx.moveTo(x + radius, y);
    ctx.arcTo(x + width, y, x + width, y + height, radius);
    ctx.arcTo(x + width, y + height, x, y + height, radius);
    ctx.arcTo(x, y + height, x, y, radius);
    ctx.arcTo(x, y, x + width, y, radius);
    ctx.closePath();
  }

  function wrapCanvasText(ctx, text, x, y, maxWidth, lineHeight, maxLines) {
    const words = String(text || '').split(/\s+/);
    let line = '';
    let lines = 0;
    for (let i = 0; i < words.length; i++) {
      const test = line + words[i] + ' ';
      if (ctx.measureText(test).width > maxWidth && i > 0) {
        ctx.fillText(line, x, y + lines * lineHeight);
        line = words[i] + ' ';
        lines++;
        if (lines >= maxLines) return;
      } else {
        line = test;
      }
    }
    if (line && lines < maxLines) ctx.fillText(line, x, y + lines * lineHeight);
  }

  async function renderPdfPageToDataUrl(pdf, pageNumber, targetWidth = 900, quality = 0.86) {
    const page = await pdf.getPage(pageNumber);
    const base = page.getViewport({ scale: 1 });
    const scale = Math.min(1.25, Math.max(0.25, targetWidth / base.width));
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d', { alpha: false });
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    return canvas.toDataURL('image/jpeg', quality);
  }

  function renderLibrary() {
    if (!els.cardsGrid) return;
    const query = (els.searchInput.value || '').toLowerCase().trim();
    const sort = els.sortSelect.value;
    let files = [...state.files];

    if (state.activeFolderId === 'unfiled') files = files.filter((file) => !file.folderId);
    else if (state.activeFolderId && state.activeFolderId !== 'all') files = files.filter((file) => file.folderId === state.activeFolderId);

    if (query) {
      files = files.filter((file) => {
        const haystack = [file.name, file.type, file.url, getFolderName(file.folderId)].join(' ').toLowerCase();
        return haystack.includes(query);
      });
    }
    files.sort((a, b) => {
      if (sort === 'name') return a.name.localeCompare(b.name);
      if (sort === 'date') return new Date(b.uploadedAt) - new Date(a.uploadedAt);
      return new Date(b.lastViewed || b.uploadedAt) - new Date(a.lastViewed || a.uploadedAt);
    });

    els.cardsGrid.innerHTML = files.map(cardTemplate).join('');
    els.emptyState.classList.toggle('hidden', files.length > 0);
    const folderText = state.activeFolderId === 'all' ? 'all folders' : state.activeFolderId === 'unfiled' ? 'Unfiled' : getFolderName(state.activeFolderId);
    els.libraryCount.textContent = `${files.length} shown • ${state.files.length} total • ${folderText}`;

    els.cardsGrid.querySelectorAll('[data-open]').forEach((button) => {
      button.addEventListener('click', () => openPresentation(button.dataset.open));
    });
    els.cardsGrid.querySelectorAll('[data-delete]').forEach((button) => {
      button.addEventListener('click', () => deletePresentation(button.dataset.delete));
    });
    els.cardsGrid.querySelectorAll('[data-move-folder]').forEach((select) => {
      select.addEventListener('change', () => movePresentationToFolder(select.dataset.moveFolder, select.value));
    });
    setupCardDragEvents();
  }

  function setupCardDragEvents() {
    if (!els.cardsGrid) return;
    els.cardsGrid.querySelectorAll('[data-card-id]').forEach((card) => {
      card.addEventListener('dragstart', (event) => {
        const tag = event.target && event.target.tagName ? event.target.tagName.toUpperCase() : '';
        if (['BUTTON', 'SELECT', 'OPTION', 'INPUT', 'TEXTAREA', 'A'].includes(tag)) {
          event.preventDefault();
          return;
        }
        const id = card.dataset.cardId;
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('application/x-presentation-id', id);
        event.dataTransfer.setData('text/plain', id);
        card.classList.add('dragging');
        document.body.classList.add('is-dragging-card');
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        document.body.classList.remove('is-dragging-card');
        document.querySelectorAll('.drop-ready').forEach((item) => item.classList.remove('drop-ready'));
      });
    });
  }

  function cardTemplate(file) {
    const date = new Date(file.uploadedAt).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
    const typeLabel = file.type === 'pdf' ? 'PDF' : file.type === 'pptx' ? 'PPTX' : 'CANVA';
    const countLabel = file.type === 'pdf' ? `${file.pageCount} pages` : file.type === 'pptx' ? `${file.pageCount} slides` : 'Canva link';
    return `
      <article class="presentation-card glass" data-card-id="${file.id}" draggable="true" title="Drag this card to a folder">
        <span class="file-badge ${file.type === 'canva' ? 'canva-badge' : ''}">${typeLabel}</span>
        <div class="card-thumb"><img src="${file.thumbnail || createGenericThumbDataUrl(file.name, file.type)}" alt="${escapeHtml(file.name)} thumbnail"></div>
        <div class="card-body">
          <h4 title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</h4>
          <div class="card-meta">
            <span>${countLabel}</span>
            <span>${date}</span>
          </div>
          <div class="card-folder-row">
            <span>${escapeHtml(getFolderName(file.folderId))}</span>
            <select data-move-folder="${file.id}" title="Move to folder">
              ${folderSelectOptions(file.folderId)}
            </select>
          </div>
          ${file.note ? `<p class="card-note">${escapeHtml(file.note)}</p>` : ''}
          <div class="card-actions">
            <button data-open="${file.id}">Open</button>
            <button class="delete-btn" data-delete="${file.id}">Delete</button>
          </div>
        </div>
      </article>
    `;
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
  }

  async function deletePresentation(id) {
    state.files = state.files.filter((file) => file.id !== id);
    await dbDelete(id);
    renderFolderList();
    renderLibrary();
  }

  async function clearLibrary() {
    if (!confirm('Clear all locally saved presentations?')) return;
    state.files = [];
    await dbClear();
    renderFolderList();
    renderLibrary();
  }

  async function openPresentation(id) {
    const file = state.files.find((item) => item.id === id);
    if (!file) return;
    stopAutoPlay();
    stopTimerInterval();
    state.activeFile = file;
    state.currentPage = 1;
    state.totalPages = file.pageCount || 1;
    state.zoom = 1;
    state.viewportCenterX = 0.5;
    state.viewportCenterY = 0.5;
    state.timingMode = 'global';
    state.perPageTiming = {};
    state.autoStartedAt = null;
    state.autoElapsedBeforePause = 0;
    state.autoCurrentDuration = getCurrentTimingSeconds();
    state.activePdf = null;
    state.pdfPageCache = new Map();
    state.pdfBitmapCache = new Map();
    resetPdfLayers();
    state.activePptxSlides = [];
    state.pptxVisualReady = false;
    state.pptxRenderedSlides = [];
    state.lastCountdownAlertSecond = null;
    state.inkStrokes = [];
    state.remoteSlideThumbs = {};
    state.remoteSlideThumbsCount = 0;
    if (state.pptxBlobUrl) {
      URL.revokeObjectURL(state.pptxBlobUrl);
      state.pptxBlobUrl = '';
    }

    file.lastViewed = new Date().toISOString();
    await dbPut(file);

    els.homeView.classList.add('hidden');
    els.viewerView.classList.remove('hidden');
    els.pdfCanvas.classList.add('hidden');
    els.pptxSlide.classList.add('hidden');
    if (els.canvaFrame) {
      els.canvaFrame.classList.add('hidden');
      els.canvaFrame.removeAttribute('src');
    }
    els.pageTotalLabel.textContent = `/ ${state.totalPages}`;
    els.jumpInput.max = state.totalPages;
    els.jumpInput.value = '1';
    updateZoomLabel();
    updateTimingModeUI();

    try {
      if (file.type === 'pdf') {
        const buffer = await file.blob.arrayBuffer();
        state.activePdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        state.totalPages = state.activePdf.numPages;
      } else if (file.type === 'pptx') {
        await preparePptxVisualRenderer(file.blob);
        if (!state.pptxVisualReady) {
          state.activePptxSlides = await parsePptxSlides(file.blob);
          state.totalPages = state.activePptxSlides.length;
        }
      } else if (file.type === 'canva') {
        state.totalPages = 1;
      }
      file.pageCount = state.totalPages;
      els.pageTotalLabel.textContent = `/ ${state.totalPages}`;
      els.jumpInput.max = state.totalPages;
      renderThumbnailSidebar();
      await renderCurrentPage();
      setupRemoteSessionIfPossible();
      revealToolbarTemporarily();
    } catch (error) {
      console.error(error);
      alert('This file could not be opened. Try converting it to PDF first.');
      closeViewer();
    }
  }

  function closeViewer() {
    if (document.fullscreenElement === els.viewerView) document.exitFullscreen().catch(() => {});
    els.viewerView.classList.remove('presentation-fullscreen');
    stopAutoPlay();
    stopTimerInterval();
    if (state.unsubscribeSession) state.unsubscribeSession();
    state.unsubscribeSession = null;
    state.sessionRef = null;
    state.sessionId = null;
    if (state.activePdf && state.activePdf.destroy) state.activePdf.destroy();
    state.activePdf = null;
    state.pdfPageCache = new Map();
    state.pdfBitmapCache = new Map();
    resetPdfLayers();
    state.activePptxSlides = [];
    state.pptxVisualReady = false;
    state.pptxRenderedSlides = [];
    state.lastCountdownAlertSecond = null;
    state.inkStrokes = [];
    state.remoteSlideThumbs = {};
    state.remoteSlideThumbsCount = 0;
    if (state.pptxBlobUrl) {
      URL.revokeObjectURL(state.pptxBlobUrl);
      state.pptxBlobUrl = '';
    }
    if (els.canvaFrame) {
      els.canvaFrame.classList.add('hidden');
      els.canvaFrame.removeAttribute('src');
    }
    state.activeFile = null;
    els.viewerView.classList.add('hidden');
    els.homeView.classList.remove('hidden');
    renderLibrary();
  }

  async function preparePptxVisualRenderer(blob) {
    state.pptxVisualReady = false;
    state.pptxRenderedSlides = [];
    els.pptxSlide.classList.remove('hidden');
    els.pptxSlide.classList.add('visual-pptx');
    els.pptxSlide.style.transform = '';
    els.pptxSlide.innerHTML = '<div class="viewer-message">Rendering PowerPoint visually...</div>';

    try {
      await ensurePptxRendererAssets();
    } catch (error) {
      console.warn('PPTXjs assets could not be loaded. Visual PPTX render is unavailable.', error);
      return;
    }

    if (!window.jQuery || !window.jQuery.fn || !window.jQuery.fn.pptxToHtml) {
      console.warn('PPTXjs visual renderer is not available.');
      return;
    }

    try {
      state.pptxBlobUrl = URL.createObjectURL(blob);
      await new Promise((resolve, reject) => {
        const host = window.jQuery(els.pptxSlide);
        host.empty();

        try {
          host.pptxToHtml({
            pptxFileUrl: state.pptxBlobUrl,
            slidesScale: '100%',
            slideMode: false,
            keyBoardShortCut: false,
            mediaProcess: true,
            themeProcess: true,
            incSlide: { height: 2, width: 2 },
            jsZipV2: 'https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@master/js/jszip.min.js',
          });
        } catch (error) {
          reject(error);
          return;
        }

        let tries = 0;
        const timer = setInterval(() => {
          const slides = collectPptxRenderedSlides();
          if (slides.length) {
            clearInterval(timer);
            state.pptxRenderedSlides = slides;
            state.pptxVisualReady = true;
            state.totalPages = slides.length;
            slides.forEach((slide, index) => {
              slide.classList.add('pptx-rendered-slide');
              slide.dataset.pageNumber = String(index + 1);
            });
            resolve();
            return;
          }

          tries += 1;
          if (tries > 140) {
            clearInterval(timer);
            reject(new Error('PPTX visual renderer timed out.'));
          }
        }, 150);
      });
    } catch (error) {
      console.warn('PPTX visual rendering failed:', error);
      state.pptxVisualReady = false;
      state.pptxRenderedSlides = [];
      els.pptxSlide.classList.remove('visual-pptx');
    }
  }

  function collectPptxRenderedSlides() {
    const host = els.pptxSlide;
    if (!host) return [];

    const selector = [
      '.slide',
      '.pptx-slide',
      '.pptxjs-slide',
      '.slide-wrapper',
      '.slideContainer',
      '.pptx-page',
      '.presentation-slide',
      '.reveal section',
    ].join(',');

    let candidates = Array.from(host.querySelectorAll(selector))
      .filter((el) => el !== host && !el.closest('.thumb-item'));

    if (!candidates.length) {
      candidates = Array.from(host.children).filter((el) => {
        const rect = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        return rect.width > 250 || rect.height > 150 || /px|%/.test(style.width + style.height);
      });
    }

    const unique = [];
    for (const el of candidates) {
      if (unique.some((item) => item.contains(el))) continue;
      unique.push(el);
    }
    return unique;
  }

  function isPresentationFullscreen() {
    return document.fullscreenElement === els.viewerView || els.viewerView.classList.contains('presentation-fullscreen');
  }

  function getAvailableStageSize() {
    const fullscreen = isPresentationFullscreen();
    if (fullscreen) {
      return { width: Math.max(320, window.innerWidth), height: Math.max(240, window.innerHeight) };
    }

    // Normal desktop view has a floating toolbar. Use the actual canvas well,
    // not the full stage, so the PDF fits below the toolbar instead of looking
    // cropped or oversized.
    const wrapRect = els.viewerCanvasWrap ? els.viewerCanvasWrap.getBoundingClientRect() : null;
    if (wrapRect && wrapRect.width > 120 && wrapRect.height > 120) {
      return {
        width: Math.max(320, wrapRect.width - 40),
        height: Math.max(240, wrapRect.height - 40),
      };
    }

    const rect = els.viewerStage ? els.viewerStage.getBoundingClientRect() : { width: window.innerWidth, height: window.innerHeight };
    return {
      width: Math.max(320, rect.width - 96),
      height: Math.max(240, rect.height - 150),
    };
  }

  function showOnlyCurrentPptxSlide() {
    if (!state.pptxVisualReady || !state.pptxRenderedSlides.length) return;
    const current = state.pptxRenderedSlides[state.currentPage - 1] || state.pptxRenderedSlides[0];
    const available = getAvailableStageSize();
    const maxWidth = available.width;
    const maxHeight = available.height;

    state.pptxRenderedSlides.forEach((slide) => {
      slide.style.display = 'none';
      slide.style.visibility = 'visible';
      slide.style.transformOrigin = 'center center';
      slide.style.margin = '0';
    });

    current.style.display = 'block';
    current.style.transform = 'none';
    current.style.margin = '0';

    const rect = current.getBoundingClientRect();
    const naturalWidth = current.offsetWidth || rect.width || 960;
    const naturalHeight = current.offsetHeight || rect.height || 540;
    const fitScale = Math.min(maxWidth / naturalWidth, maxHeight / naturalHeight);
    const finalScale = Math.max(0.15, fitScale * state.zoom);

    current.style.transform = `scale(${finalScale})`;
    els.pptxSlide.style.width = `${naturalWidth * finalScale}px`;
    els.pptxSlide.style.height = `${naturalHeight * finalScale}px`;
    els.pptxSlide.style.transform = '';
  }

  async function renderCurrentPage() {
    if (!state.activeFile) return;
    const token = ++state.renderToken;
    state.currentPage = Math.min(Math.max(1, state.currentPage), state.totalPages);
    if (els.jumpInput) els.jumpInput.value = state.currentPage;
    if (els.perSlideTimingInput) els.perSlideTimingInput.value = state.perPageTiming[state.currentPage] || '';
    updateTimingModeUI();
    updateActiveThumb();
    updateZoomLabel();
    publishSessionState();

    if (state.activeFile.type === 'pdf' && state.activePdfRenderTask) {
      try { state.activePdfRenderTask.cancel(); } catch (error) {}
      state.activePdfRenderTask = null;
    }

    try {
      if (state.activeFile.type === 'canva') {
        els.pdfCanvas.classList.add('hidden');
        els.pptxSlide.classList.add('hidden');
        if (els.canvaFrame) {
          els.canvaFrame.classList.remove('hidden');
          if (els.canvaFrame.getAttribute('src') !== state.activeFile.embedUrl) els.canvaFrame.setAttribute('src', state.activeFile.embedUrl);
        }
        applyViewportScroll();
      } else if (state.activeFile.type === 'pdf') {
        if (els.canvaFrame) els.canvaFrame.classList.add('hidden');
        els.pptxSlide.classList.add('hidden');
        await renderPdfPageSmooth(token);
        if (token !== state.renderToken) return;
        warmAdjacentPdfPages();
        warmAdjacentPdfBitmaps();
        applyViewportScroll();
      } else {
        if (els.canvaFrame) els.canvaFrame.classList.add('hidden');
        els.pdfCanvas.classList.add('hidden');
        els.pptxSlide.classList.remove('hidden');
        if (state.pptxVisualReady) {
          showOnlyCurrentPptxSlide();
          applyViewportScroll();
        } else {
          els.pptxSlide.classList.remove('visual-pptx');
          els.pptxSlide.style.transform = '';
          els.pptxSlide.innerHTML = `
            <div class="viewer-message viewer-message-warning">
              <strong>PowerPoint visual render is not available.</strong>
              <span>For the exact colors, pictures, fonts, and layout, export this PPTX as PDF and upload the PDF here. Static browser PPTX rendering cannot guarantee every PowerPoint element.</span>
            </div>
          `;
        }
      }
      renderInkOverlay();
      if (state.slideChangePending && token === state.renderToken) {
        playSlideTransition();
        state.slideChangePending = false;
      }
    } catch (error) {
      const cancelled = error && (error.name === 'RenderingCancelledException' || /cancel/i.test(String(error.message || error)));
      if (!cancelled) console.warn('Slide render failed:', error);
    }
  }

  function renderThumbnailSidebar() {
    els.thumbnailSidebar.innerHTML = '';
    for (let i = 1; i <= state.totalPages; i++) {
      const item = document.createElement('button');
      item.className = `thumb-item${i === state.currentPage ? ' active' : ''}`;
      item.type = 'button';
      item.dataset.page = String(i);
      item.innerHTML = `
        <div class="thumb-canvas-wrap" data-thumb-wrap="${i}"><span>Loading...</span></div>
        <div class="thumb-label"><span>${state.activeFile.type === 'pdf' ? 'Page' : state.activeFile.type === 'canva' ? 'Canva' : 'Slide'} ${i}</span><span>${state.perPageTiming[i] ? state.perPageTiming[i] + 's' : ''}</span></div>
      `;
      item.addEventListener('click', () => jumpToPage(i));
      els.thumbnailSidebar.appendChild(item);
    }

    const observer = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          renderSidebarThumb(Number(entry.target.dataset.page));
          observer.unobserve(entry.target);
        }
      });
    }, { root: els.thumbnailSidebar, rootMargin: '220px' });

    els.thumbnailSidebar.querySelectorAll('.thumb-item').forEach((item) => observer.observe(item));
  }

  async function renderSidebarThumb(pageNumber) {
    const wrap = els.thumbnailSidebar.querySelector(`[data-thumb-wrap="${pageNumber}"]`);
    if (!wrap || wrap.dataset.rendered) return;
    wrap.dataset.rendered = '1';
    wrap.innerHTML = '';

    if (state.activeFile.type === 'canva') {
      const img = document.createElement('img');
      img.alt = 'Canva presentation link';
      img.src = state.activeFile.thumbnail || createCanvaThumbDataUrl(state.activeFile.name);
      wrap.appendChild(img);
    } else if (state.activeFile.type === 'pdf') {
      const page = await state.activePdf.getPage(pageNumber);
      const viewportBase = page.getViewport({ scale: 1 });
      const scale = 170 / viewportBase.width;
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      wrap.appendChild(canvas);
    } else {
      if (state.pptxVisualReady && state.pptxRenderedSlides[pageNumber - 1]) {
        const clone = state.pptxRenderedSlides[pageNumber - 1].cloneNode(true);
        clone.style.display = 'block';
        clone.style.transformOrigin = 'top left';
        clone.style.transform = 'scale(0.16)';
        clone.style.margin = '0';
        clone.style.pointerEvents = 'none';
        const thumbShell = document.createElement('div');
        thumbShell.className = 'pptx-thumb-shell';
        thumbShell.appendChild(clone);
        wrap.appendChild(thumbShell);
      } else {
        const slide = state.activePptxSlides[pageNumber - 1] || { title: `Slide ${pageNumber}` };
        const img = document.createElement('img');
        img.alt = `Slide ${pageNumber}`;
        img.src = createPptxThumbDataUrl(slide.title, pageNumber);
        wrap.appendChild(img);
      }
    }
  }

  function updateActiveThumb() {
    if (!els.thumbnailSidebar) return;
    els.thumbnailSidebar.querySelectorAll('.thumb-item').forEach((item) => {
      item.classList.toggle('active', Number(item.dataset.page) === state.currentPage);
    });
  }

  function previousPage() {
    if (state.currentPage <= 1) return;
    state.currentPage--;
    handleSlideChanged();
    renderCurrentPage();
  }

  function nextPage() {
    if (state.currentPage >= state.totalPages) return;
    state.currentPage++;
    handleSlideChanged();
    renderCurrentPage();
  }

  function jumpToPage(page) {
    const target = Math.min(Math.max(1, Number(page) || 1), state.totalPages);
    if (target === state.currentPage) return;
    state.currentPage = target;
    handleSlideChanged();
    renderCurrentPage();
  }

  function handleSlideChanged() {
    state.viewportCenterX = 0.5;
    state.viewportCenterY = 0.5;
    state.slideChangePending = true;
    state.lastCountdownAlertSecond = null;
    resetAutoClockForCurrentSlide();
    if (isAutoClockActive()) resetTimer({ publish: false, start: state.autoPlaying });
    // Pro v10: publish the new slide number immediately before the heavier render/preview path.
    publishSessionState(true);
  }

  function setZoom(value, options = {}) {
    state.zoom = Math.min(4, Math.max(0.25, Number(value)));
    if (Number.isFinite(Number(options.centerX))) state.viewportCenterX = clamp01(Number(options.centerX));
    if (Number.isFinite(Number(options.centerY))) state.viewportCenterY = clamp01(Number(options.centerY));
    renderCurrentPage();
  }

  function setViewportTransform(value = {}) {
    const nextZoom = Math.min(4, Math.max(0.25, Number(value.zoom ?? state.zoom) || 1));
    state.zoom = nextZoom;
    if (Number.isFinite(Number(value.centerX))) state.viewportCenterX = clamp01(Number(value.centerX));
    if (Number.isFinite(Number(value.centerY))) state.viewportCenterY = clamp01(Number(value.centerY));
    updateZoomLabel();

    // Remote pinch/drag should feel instant. For PDFs, resize/reposition the
    // already-rendered page first, then do a quiet high-quality re-render after
    // the gesture settles. This avoids the desktop flash/jitter caused by
    // running PDF.js on every tiny finger movement.
    if (state.activeFile && state.activeFile.type === 'pdf' && applyLivePdfViewportTransform()) {
      scheduleViewportQualityRender();
      return;
    }

    renderCurrentPage();
  }

  function scheduleViewportQualityRender() {
    clearTimeout(state.viewportQualityTimer);
    state.viewportQualityTimer = setTimeout(() => {
      if (state.activeFile && state.activeFile.type === 'pdf') renderCurrentPage();
    }, 380);
  }

  function applyLivePdfViewportTransform() {
    const layer = state.activePdfCanvas;
    if (!layer || !layer.isConnected) return false;
    const renderedZoom = Math.max(0.001, Number(layer.dataset.renderZoom) || 1);
    const baseWidth = Number(layer.dataset.baseWidth) || ((layer.offsetWidth || 1) / renderedZoom);
    const baseHeight = Number(layer.dataset.baseHeight) || ((layer.offsetHeight || 1) / renderedZoom);
    layer.style.width = `${Math.max(1, Math.round(baseWidth * state.zoom))}px`;
    layer.style.height = `${Math.max(1, Math.round(baseHeight * state.zoom))}px`;
    layer.style.maxWidth = 'none';
    layer.style.maxHeight = 'none';
    applyViewportScroll();
    return true;
  }

  function clamp01(value) {
    return Math.max(0, Math.min(1, Number(value) || 0));
  }

  function applyViewportScroll() {
    requestAnimationFrame(() => {
      const wrap = els.viewerCanvasWrap;
      if (!wrap) return;
      const maxLeft = Math.max(0, wrap.scrollWidth - wrap.clientWidth);
      const maxTop = Math.max(0, wrap.scrollHeight - wrap.clientHeight);
      wrap.scrollLeft = Math.max(0, Math.min(maxLeft, (wrap.scrollWidth * state.viewportCenterX) - (wrap.clientWidth / 2)));
      wrap.scrollTop = Math.max(0, Math.min(maxTop, (wrap.scrollHeight * state.viewportCenterY) - (wrap.clientHeight / 2)));
      renderInkOverlay();
    });
  }

  function updateZoomLabel() {
    els.zoomLabel.textContent = `${Math.round(state.zoom * 100)}%`;
  }

  function handleKeyboard(event) {
    if (!els.viewerView || els.viewerView.classList.contains('hidden')) return;

    const key = event.key;
    const isPresentationKey = key === 'ArrowRight' || key === 'PageDown' || key === ' ' || key === 'ArrowLeft' || key === 'PageUp' || key.toLowerCase() === 'f' || key === 'Escape';
    if (!isPresentationKey) return;

    if (key === 'Escape' && document.fullscreenElement) {
      event.preventDefault();
      event.stopPropagation();
      document.exitFullscreen().catch(() => {});
      return;
    }

    if (shouldIgnoreKeyboardForTyping()) return;

    if (key === 'ArrowRight' || key === 'PageDown' || key === ' ') {
      event.preventDefault();
      event.stopPropagation();
      nextPage();
      return;
    }
    if (key === 'ArrowLeft' || key === 'PageUp') {
      event.preventDefault();
      event.stopPropagation();
      previousPage();
      return;
    }
    if (key.toLowerCase() === 'f') {
      event.preventDefault();
      event.stopPropagation();
      toggleFullscreen();
    }
  }

  function shouldIgnoreKeyboardForTyping() {
    const el = document.activeElement;
    if (!el) return false;
    const tag = el.tagName || '';
    const typingTarget = el.isContentEditable || ['INPUT', 'TEXTAREA', 'SELECT'].includes(tag);
    if (!typingTarget) return false;

    // In fullscreen presentation mode, the last-focused setting input can remain active
    // even after the controls are hidden. Do not let that hidden input steal arrow keys.
    if (isPresentationFullscreen()) {
      const inVisibleControlPanel = els.controlPanel && !els.controlPanel.classList.contains('hidden') && els.controlPanel.contains(el);
      const inVisibleQrModal = els.qrModal && !els.qrModal.classList.contains('hidden') && els.qrModal.contains(el);
      const inVisibleSetupModal = els.setupModal && !els.setupModal.classList.contains('hidden') && els.setupModal.contains(el);
      if (!inVisibleControlPanel && !inVisibleQrModal && !inVisibleSetupModal) return false;
    }

    return isElementVisible(el);
  }

  function isElementVisible(el) {
    if (!el || !document.body.contains(el)) return false;
    const style = window.getComputedStyle(el);
    const rect = el.getBoundingClientRect();
    return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
  }

  async function toggleFullscreen() {
    if (document.activeElement && typeof document.activeElement.blur === 'function') {
      document.activeElement.blur();
    }
    if (!document.fullscreenElement) {
      await els.viewerView.requestFullscreen().catch(() => {});
    } else {
      await document.exitFullscreen().catch(() => {});
    }
  }

  function syncFullscreenState() {
    if (!els.viewerView) return;
    const fullscreen = document.fullscreenElement === els.viewerView;
    els.viewerView.classList.toggle('presentation-fullscreen', fullscreen);
    if (els.fullscreenBtn) els.fullscreenBtn.textContent = fullscreen ? 'Exit Fullscreen' : 'Fullscreen';
    if (els.controlPanel && fullscreen) els.controlPanel.classList.add('hidden');
    if (fullscreen && document.activeElement && typeof document.activeElement.blur === 'function') {
      document.activeElement.blur();
    }
    revealToolbarTemporarily();
    if (state.activeFile) renderCurrentPage();
  }

  function revealToolbarTemporarily() {
    if (!els.viewerToolbar) return;
    els.viewerToolbar.classList.remove('toolbar-hidden');
    if (els.viewerStage) els.viewerStage.classList.remove('cursor-hidden');
    clearTimeout(state.toolbarHideTimer);
    state.toolbarHideTimer = setTimeout(() => {
      if (!els.controlPanel.classList.contains('hidden')) return;
      els.viewerToolbar.classList.add('toolbar-hidden');
      if (els.viewerStage && isPresentationFullscreen()) els.viewerStage.classList.add('cursor-hidden');
    }, 2400);
  }

  function updateTimingModeUI() {
    if (els.timingModeSelect) els.timingModeSelect.value = state.timingMode;
    const perSlideMode = state.timingMode === 'per-slide';
    if (els.perSlideTimingWrap) els.perSlideTimingWrap.classList.toggle('timing-disabled', !perSlideMode);
    if (els.perSlideTimingInput) {
      els.perSlideTimingInput.disabled = !perSlideMode;
      els.perSlideTimingInput.placeholder = perSlideMode ? 'Use global fallback' : 'Disabled in global mode';
    }
  }

  function getGlobalTimingSeconds() {
    const selected = els.globalTimingSelect.value;
    if (selected === 'custom') return Math.max(1, Number(els.customTimingInput.value) || 10);
    return Number(selected) || 10;
  }

  function getCurrentTimingSeconds() {
    if (state.timingMode !== 'per-slide') return getGlobalTimingSeconds();
    return state.perPageTiming[state.currentPage] || getGlobalTimingSeconds();
  }

  function setGlobalTimingSeconds(seconds) {
    const safe = Math.max(1, Number(seconds) || 10);
    const preset = ['5', '10', '15', '20', '30'].includes(String(safe)) ? String(safe) : 'custom';
    els.globalTimingSelect.value = preset;
    els.customTimingWrap.classList.toggle('hidden', preset !== 'custom');
    els.customTimingInput.value = safe;
    resetAutoClockForCurrentSlide();
    updateTimerText();
  }

  function setCurrentSlideTimingSeconds(seconds) {
    const safe = Math.max(1, Number(seconds) || getGlobalTimingSeconds());
    state.perPageTiming[state.currentPage] = safe;
    state.timingMode = 'per-slide';
    updateTimingModeUI();
    renderThumbnailSidebar();
    resetAutoClockForCurrentSlide();
    updateTimerText();
  }

  function isAutoClockActive() {
    return state.autoPlaying || state.autoPaused || Boolean(state.autoTimer);
  }

  function getAutoElapsedSeconds(raw = false) {
    const running = state.autoPlaying && state.autoStartedAt ? (Date.now() - state.autoStartedAt) / 1000 : 0;
    const total = state.autoElapsedBeforePause + running;
    return raw ? total : Math.floor(total);
  }

  function resetAutoClockForCurrentSlide() {
    state.autoCurrentDuration = getCurrentTimingSeconds();
    state.autoElapsedBeforePause = 0;
    state.autoStartedAt = state.autoPlaying ? Date.now() : null;
    state.lastCountdownAlertSecond = null;
  }

  function startAutoPlay() {
    stopAutoPlay(false);
    state.autoPlaying = true;
    state.autoPaused = false;
    resetAutoClockForCurrentSlide();
    startAutoLoop();
    if (state.timer.visible) resetTimer({ publish: false, start: true });
    publishSessionState();
  }

  function startAutoLoop() {
    clearInterval(state.autoTimer);
    state.autoTimer = setInterval(() => {
      if (!state.autoPlaying || state.autoPaused) return;
      updateTimerText();
      checkCountdownAlert();
      if (getAutoElapsedSeconds(true) < state.autoCurrentDuration || state.autoAdvancing) return;
      if (state.currentPage >= state.totalPages) {
        stopAutoPlay();
        return;
      }
      state.autoAdvancing = true;
      state.currentPage += 1;
      handleSlideChanged();
      renderCurrentPage().finally(() => { state.autoAdvancing = false; });
    }, 120);
  }

  function pauseAutoPlay() {
    if (!state.autoPlaying && !state.autoTimer) return;
    state.autoElapsedBeforePause = getAutoElapsedSeconds(true);
    state.autoStartedAt = null;
    state.autoPlaying = false;
    state.autoPaused = true;
    clearInterval(state.autoTimer);
    state.autoTimer = null;
    updateTimerText();
    publishSessionState();
  }

  function resumeAutoPlay() {
    if (!state.activeFile) return;
    state.autoPlaying = true;
    state.autoPaused = false;
    state.autoStartedAt = Date.now();
    startAutoLoop();
    if (state.timer.visible) startTimerInterval();
    publishSessionState();
  }

  function stopAutoPlay(publish = true) {
    clearInterval(state.autoTimer);
    state.autoTimer = null;
    state.autoPlaying = false;
    state.autoPaused = false;
    state.autoStartedAt = null;
    state.autoElapsedBeforePause = 0;
    state.autoAdvancing = false;
    state.autoCurrentDuration = getCurrentTimingSeconds();
    resetTimer({ publish: false, start: false });
    if (publish) publishSessionState();
  }

  function showTimer() {
    state.timer.visible = true;
    els.timerOverlay.classList.remove('hidden');
    if (isAutoClockActive()) {
      updateTimerText();
    } else if (!state.timerStartedAt) {
      state.timerStartedAt = Date.now();
    }
    if (!state.autoPaused) startTimerInterval();
    applyTimerSettings();
    publishSessionState();
  }

  function hideTimer() {
    state.timer.visible = false;
    els.timerOverlay.classList.add('hidden');
    stopTimerInterval();
    publishSessionState();
  }

  function resetTimer(options = {}) {
    const publish = options.publish !== false;
    const shouldStart = options.start !== false && state.timer.visible && !state.autoPaused;
    state.timerElapsedBeforePause = 0;
    state.timerStartedAt = shouldStart && !isAutoClockActive() ? Date.now() : null;
    if (shouldStart) startTimerInterval();
    else stopTimerInterval(true);
    updateTimerText();
    if (publish) publishSessionState();
  }

  function startTimerInterval() {
    stopTimerInterval(false);
    state.timerTick = setInterval(updateTimerText, 250);
    updateTimerText();
  }

  function stopTimerInterval(clearStartedAt = true) {
    clearInterval(state.timerTick);
    state.timerTick = null;
    if (clearStartedAt) state.timerStartedAt = null;
  }

  function elapsedSeconds() {
    if (isAutoClockActive()) return getAutoElapsedSeconds(false);
    if (!state.timerStartedAt) return 0;
    return Math.floor((Date.now() - state.timerStartedAt) / 1000) + state.timerElapsedBeforePause;
  }

  function timerCountdownTargetSeconds() {
    if (isAutoClockActive()) return state.autoCurrentDuration || getCurrentTimingSeconds();
    return state.timer.countdownSeconds;
  }

  function updateTimerText() {
    if (!els.timerOverlay) return;
    let seconds = elapsedSeconds();
    if (state.timer.mode === 'down') seconds = Math.max(0, timerCountdownTargetSeconds() - seconds);
    els.timerOverlay.textContent = formatTime(seconds);
  }

  function formatTime(totalSeconds) {
    const safeSeconds = Math.max(0, Math.floor(totalSeconds));
    const hours = Math.floor(safeSeconds / 3600);
    const minutes = Math.floor((safeSeconds % 3600) / 60);
    const seconds = safeSeconds % 60;
    return [hours, minutes, seconds].map((n) => String(n).padStart(2, '0')).join(':');
  }

  function applyTimerSettings() {
    els.timerOverlay.classList.remove('bottom-right', 'bottom-left', 'top-right', 'top-left');
    els.timerOverlay.classList.add(state.timer.position);
    els.timerOverlay.style.opacity = String(state.timer.opacity / 100);
    const timerSize = Math.max(16, Math.min(72, Number(state.timer.size) || 28));
    els.timerOverlay.style.setProperty('--timer-size', `${timerSize}px`);
    els.timerOverlay.style.setProperty('--timer-pad-y', `${Math.max(8, Math.round(timerSize * 0.42))}px`);
    els.timerOverlay.style.setProperty('--timer-pad-x', `${Math.max(12, Math.round(timerSize * 0.58))}px`);
    els.timerOverlay.style.setProperty('--timer-min-width', `${Math.max(126, Math.round(timerSize * 5.4))}px`);
    els.timerModeSelect.value = state.timer.mode;
    els.timerPositionSelect.value = state.timer.position;
    els.timerOpacityInput.value = state.timer.opacity;
    if (els.timerSizeInput) els.timerSizeInput.value = timerSize;
    if (els.timerSizeLabel) els.timerSizeLabel.textContent = `${timerSize}px`;
    if (els.countdownAlertSelect) els.countdownAlertSelect.value = state.countdownAlert || 'off';
    if (els.countdownVoiceSelect) els.countdownVoiceSelect.value = normalizeCountdownVoiceStyle(state.countdownVoiceGender);
    if (els.countdownVoiceStartSelect) els.countdownVoiceStartSelect.value = String(getCountdownVoiceStart());
    if (els.slideTransitionSelect) els.slideTransitionSelect.value = state.transitionEffect || 'fade';
  }

  async function getCachedPdfPage(pageNumber) {
    if (!state.activePdf) return null;
    const safePage = Math.min(Math.max(1, Number(pageNumber) || 1), state.totalPages || 1);
    if (!state.pdfPageCache) state.pdfPageCache = new Map();
    if (!state.pdfPageCache.has(safePage)) {
      state.pdfPageCache.set(safePage, state.activePdf.getPage(safePage));
    }
    return state.pdfPageCache.get(safePage);
  }

  function warmAdjacentPdfPages() {
    if (!state.activePdf || !state.pdfPageCache) return;
    const keep = new Set();
    for (let page = state.currentPage - 3; page <= state.currentPage + 3; page++) {
      if (page >= 1 && page <= state.totalPages) {
        keep.add(page);
        if (!state.pdfPageCache.has(page)) state.pdfPageCache.set(page, state.activePdf.getPage(page));
      }
    }
    for (const key of Array.from(state.pdfPageCache.keys())) {
      if (!keep.has(key)) state.pdfPageCache.delete(key);
    }
  }

  function getTransitionClasses() {
    return ['slide-transition', 'transition-fade', 'transition-slide-left', 'transition-slide-right', 'transition-slide-up', 'transition-zoom-in', 'transition-zoom-out', 'transition-soft-blur'];
  }

  function playSlideTransition() {
    const effect = state.transitionEffect || 'fade';
    if (effect === 'none') return;
    if (state.activeFile && state.activeFile.type === 'pdf') return; // PDF uses the double-buffer renderer, not wrapper animation.
    const target = getTransitionTarget();
    if (!target) return;
    const classes = getTransitionClasses();
    target.classList.remove(...classes);
    void target.offsetWidth;
    target.classList.add('slide-transition', `transition-${effect}`);
    window.setTimeout(() => target.classList.remove(...classes), state.transitionDuration + 180);
  }

  function getTransitionTarget() {
    return els.pptxSlide || els.viewerCanvasWrap;
  }

  async function renderPdfPageSmooth(token) {
    const wrap = els.viewerCanvasWrap;
    if (!wrap || !state.activePdf) return;
    wrap.classList.add('pdf-layer-mode');
    if (els.pdfCanvas) els.pdfCanvas.classList.add('hidden');

    const pageNumber = state.currentPage;
    const renderInfo = await getPdfRenderInfo(pageNumber);
    if (token !== state.renderToken || !renderInfo) return;

    const bitmapKey = makePdfBitmapKey(pageNumber, renderInfo);
    let incoming = takeCachedPdfBitmap(bitmapKey);
    if (!incoming) incoming = await renderPdfPageToCanvas(pageNumber, renderInfo, token);
    if (token !== state.renderToken || !incoming) return;

    preparePdfLayer(incoming, renderInfo);
    const oldLayer = state.activePdfCanvas && state.activePdfCanvas.isConnected ? state.activePdfCanvas : null;
    const effect = state.transitionEffect || 'fade';
    const shouldAnimate = effect !== 'none' && (state.slideChangePending || !oldLayer);

    incoming.classList.add('pdf-page-layer', 'pdf-layer-current');
    incoming.dataset.pageNumber = String(pageNumber);
    incoming.dataset.bitmapKey = bitmapKey;

    if (oldLayer && oldLayer.dataset.pageNumber === String(pageNumber) && oldLayer.dataset.bitmapKey === bitmapKey) {
      // Same rendered page/zoom after a resize race; keep a single clean layer.
      oldLayer.remove();
    }

    if (shouldAnimate) {
      incoming.classList.add('pdf-layer-incoming', `transition-${effect}`);
      incoming.style.setProperty('--ph-transition-ms', `${state.transitionDuration}ms`);
    }

    wrap.appendChild(incoming);
    state.activePdfCanvas = incoming;

    // Keep the old PDF visible only until the incoming bitmap is ready, then fade it out cleanly.
    if (oldLayer && oldLayer !== incoming) {
      oldLayer.classList.remove('pdf-layer-current');
      oldLayer.classList.add('pdf-layer-outgoing');
      oldLayer.style.setProperty('--ph-transition-ms', `${Math.max(180, Math.round(state.transitionDuration * 0.55))}ms`);
      window.setTimeout(() => oldLayer.remove(), Math.max(260, state.transitionDuration + 120));
    }

    if (shouldAnimate) {
      requestAnimationFrame(() => {
        incoming.classList.add('pdf-layer-show');
        window.setTimeout(() => {
          incoming.classList.remove('pdf-layer-incoming', 'pdf-layer-show', `transition-${effect}`);
        }, state.transitionDuration + 120);
      });
    }

    // Remove stale PDF layers left by fast keyboard presses.
    window.setTimeout(() => cleanupPdfLayers(), Math.max(500, state.transitionDuration + 220));
  }

  async function getPdfRenderInfo(pageNumber) {
    const page = await getCachedPdfPage(pageNumber);
    if (!page) return null;
    const viewportBase = page.getViewport({ scale: 1 });
    const available = getAvailableStageSize();
    const fitScale = Math.min(available.width / viewportBase.width, available.height / viewportBase.height);
    const cssScale = Math.max(0.15, fitScale * state.zoom);
    const dpr = Math.min(window.devicePixelRatio || 1, 2.4);
    const renderViewport = page.getViewport({ scale: cssScale * dpr });
    const cssViewport = page.getViewport({ scale: cssScale });
    return { page, viewportBase, available, fitScale, cssScale, dpr, renderViewport, cssViewport };
  }

  function makePdfBitmapKey(pageNumber, info) {
    const w = Math.round(info.cssViewport.width);
    const h = Math.round(info.cssViewport.height);
    const z = Math.round(state.zoom * 1000);
    const fs = isPresentationFullscreen() ? 'fs' : 'normal';
    return `${pageNumber}:${w}x${h}:z${z}:${fs}`;
  }

  function takeCachedPdfBitmap(key) {
    if (!state.pdfBitmapCache || !state.pdfBitmapCache.has(key)) return null;
    const cached = state.pdfBitmapCache.get(key);
    state.pdfBitmapCache.delete(key);
    return cached;
  }

  async function renderPdfPageToCanvas(pageNumber, info, token) {
    const canvas = document.createElement('canvas');
    canvas.className = 'pdf-canvas pdf-page-layer';
    preparePdfLayer(canvas, info);
    canvas.width = Math.floor(info.renderViewport.width);
    canvas.height = Math.floor(info.renderViewport.height);
    const ctx = canvas.getContext('2d', { alpha: false });
    ctx.save();
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.restore();

    const task = info.page.render({ canvasContext: ctx, viewport: info.renderViewport });
    state.activePdfRenderTask = task;
    try {
      await task.promise;
    } finally {
      if (state.activePdfRenderTask === task) state.activePdfRenderTask = null;
    }
    if (token !== state.renderToken) return null;
    return canvas;
  }

  function preparePdfLayer(canvas, info) {
    canvas.classList.remove('hidden', ...getTransitionClasses(), 'pdf-layer-incoming', 'pdf-layer-show', 'pdf-layer-outgoing');
    const renderZoom = Math.max(0.001, Number(state.zoom) || 1);
    const width = Math.floor(info.cssViewport.width);
    const height = Math.floor(info.cssViewport.height);
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;
    canvas.style.maxWidth = 'none';
    canvas.style.maxHeight = 'none';
    canvas.dataset.renderZoom = String(renderZoom);
    canvas.dataset.baseWidth = String(width / renderZoom);
    canvas.dataset.baseHeight = String(height / renderZoom);
  }

  async function warmAdjacentPdfBitmaps() {
    if (!state.activePdf || !state.pdfBitmapCache) return;
    const center = state.currentPage;
    const pages = [center + 1, center - 1].filter((p) => p >= 1 && p <= state.totalPages);
    for (const pageNumber of pages) {
      try {
        const info = await getPdfRenderInfo(pageNumber);
        if (!info) continue;
        const key = makePdfBitmapKey(pageNumber, info);
        if (state.pdfBitmapCache.has(key)) continue;
        const token = state.renderToken;
        const canvas = await renderPdfPageToCanvas(pageNumber, info, token);
        if (canvas && token === state.renderToken) {
          canvas.dataset.bitmapKey = key;
          canvas.dataset.pageNumber = String(pageNumber);
          state.pdfBitmapCache.set(key, canvas);
        }
      } catch (error) {
        // Pre-rendering is optional. The main renderer will still work.
      }
    }
    trimPdfBitmapCache();
  }

  function trimPdfBitmapCache() {
    if (!state.pdfBitmapCache) return;
    const keepKeys = [];
    for (const key of state.pdfBitmapCache.keys()) {
      const page = Number(String(key).split(':')[0]);
      if (Math.abs(page - state.currentPage) <= 2) keepKeys.push(key);
    }
    for (const key of Array.from(state.pdfBitmapCache.keys())) {
      if (!keepKeys.includes(key) || keepKeys.length > 4) state.pdfBitmapCache.delete(key);
    }
  }


  function resetPdfLayers() {
    state.activePdfCanvas = null;
    state.pdfRenderQueueKey = '';
    if (!els.viewerCanvasWrap) return;
    els.viewerCanvasWrap.classList.remove('pdf-layer-mode');
    els.viewerCanvasWrap.querySelectorAll('.pdf-page-layer').forEach((layer) => layer.remove());
    if (els.pdfCanvas && !els.pdfCanvas.isConnected) els.viewerCanvasWrap.prepend(els.pdfCanvas);
    if (els.pdfCanvas) {
      els.pdfCanvas.className = 'pdf-canvas hidden';
      els.pdfCanvas.removeAttribute('style');
    }
  }

  function cleanupPdfLayers() {
    if (!els.viewerCanvasWrap) return;
    const layers = Array.from(els.viewerCanvasWrap.querySelectorAll('.pdf-page-layer'));
    layers.forEach((layer) => {
      if (layer !== state.activePdfCanvas && !layer.classList.contains('pdf-layer-outgoing')) layer.remove();
    });
  }

  function checkCountdownAlert() {
    if (!state.autoPlaying || state.autoPaused || state.countdownAlert === 'off') return;
    const remaining = Math.ceil(Math.max(0, state.autoCurrentDuration - getAutoElapsedSeconds(true)));
    if (remaining < 1 || remaining > 5) return;
    if (remaining === state.lastCountdownAlertSecond) return;
    state.lastCountdownAlertSecond = remaining;
    triggerCountdownAlert(remaining);
  }

  function getCountdownVoiceStart() {
    return Number(state.countdownVoiceStart) === 3 ? 3 : 5;
  }

  function triggerCountdownAlert(second, overrideMode, options = {}) {
    const mode = overrideMode || state.countdownAlert;
    if (mode === 'sound' || mode === 'both') playCountdownBeep(second);
    const voiceEnabled = mode === 'voice' || mode === 'both';
    const voiceShouldPlay = options.forceVoice || second <= getCountdownVoiceStart();
    if (voiceEnabled && voiceShouldPlay) speakCountdown(second);
  }

  function runCountdownAlertPreview(overrideMode) {
    const mode = overrideMode || (state.countdownAlert === 'off' ? 'both' : state.countdownAlert);
    if (state.countdownPreviewTimers) state.countdownPreviewTimers.forEach((timer) => clearTimeout(timer));
    state.countdownPreviewTimers = [];
    if (window.speechSynthesis) window.speechSynthesis.cancel();
    primePresentationAudio();
    [5, 4, 3, 2, 1].forEach((second, index) => {
      const timer = setTimeout(() => triggerCountdownAlert(second, mode), index * 520);
      state.countdownPreviewTimers.push(timer);
    });
  }

  function primePresentationAudio() {
    try {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (AudioContext) {
        state.audioContext = state.audioContext || new AudioContext();
        if (state.audioContext.state === 'suspended') {
          state.audioContext.resume().catch(() => {});
        }
        state.audioUnlocked = true;
      }
      if (window.speechSynthesis && window.speechSynthesis.getVoices) {
        window.speechSynthesis.getVoices();
      }
    } catch (error) {}
  }

  async function playCountdownBeep(second) {
    try {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (!AudioContext) return;
      state.audioContext = state.audioContext || new AudioContext();
      const ctx = state.audioContext;
      if (ctx.state === 'suspended') {
        await ctx.resume().catch(() => {});
      }
      if (ctx.state === 'suspended') return;
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.value = second === 1 ? 1100 : 760;
      gain.gain.setValueAtTime(0.0001, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.22, ctx.currentTime + 0.012);
      gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + 0.22);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(ctx.currentTime);
      osc.stop(ctx.currentTime + 0.24);
    } catch (error) {}
  }

  function getPreferredCountdownVoice() {
    try {
      if (!window.speechSynthesis || !window.speechSynthesis.getVoices) return null;
      const voices = window.speechSynthesis.getVoices().filter((voice) => /^en/i.test(voice.lang || ''));
      if (!voices.length) return null;
      const profile = getCountdownVoiceProfile();
      const hints = profile.hints || [];
      const byName = voices.find((voice) => hints.some((hint) => `${voice.name} ${voice.voiceURI}`.toLowerCase().includes(hint)));
      return byName || voices.find((voice) => /en-US/i.test(voice.lang || '')) || voices[0];
    } catch (error) {
      return null;
    }
  }

  function speakCountdown(second) {
    try {
      if (!window.speechSynthesis) return;
      const words = { 5: 'five', 4: 'four', 3: 'three', 2: 'two', 1: 'one' };
      window.speechSynthesis.cancel();
      const utterance = new SpeechSynthesisUtterance(words[second] || String(second));
      utterance.lang = 'en-US';
      utterance.voice = getPreferredCountdownVoice();
      const voiceProfile = getCountdownVoiceProfile();
      utterance.rate = voiceProfile.rate || 1;
      utterance.pitch = voiceProfile.pitch || 1;
      utterance.volume = 1;
      window.speechSynthesis.speak(utterance);
    } catch (error) {}
  }

  async function setupRemoteSessionIfPossible() {
    if (!state.firebaseReady || !state.activeFile) return;
    if (state.unsubscribeSession) state.unsubscribeSession();
    state.sessionId = state.sessionId || makeSessionId();
    state.sessionRef = state.firebaseDb.collection(SESSION_COLLECTION).doc(state.sessionId);
    state.lastCommandId = null;
    // v10.6 cleanup: old builds stored every slide thumbnail inside the live
    // session document. That made every phone command feel delayed. Remove it
    // once, then keep thumbnails in a lightweight subcollection instead.
    try {
      if (window.firebase && window.firebase.firestore && window.firebase.firestore.FieldValue) {
        await state.sessionRef.set({ slideThumbs: window.firebase.firestore.FieldValue.delete() }, { merge: true });
      }
    } catch (error) {}
    await publishSessionState(true);
    state.unsubscribeSession = state.sessionRef.onSnapshot((snap) => {
      if (!snap.exists) return;
      const data = snap.data();
      if (!data || !data.command || data.command.id === state.lastCommandId) return;
      state.lastCommandId = data.command.id;
      applyRemoteCommand(data.command);
    });
  }

  function makeSessionId() {
    return Math.random().toString(36).slice(2, 6).toUpperCase() + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
  }

  async function publishSessionState(force = false) {
    if (!state.sessionRef || !state.activeFile || (state.publishLock && !force)) return;
    state.publishLock = true;
    setTimeout(() => { state.publishLock = false; }, force ? 0 : 120);

    // Keep the hot remote-control path light: write slide/control state first,
    // then upload the heavier thumbnail preview in a short deferred task.
    const payload = {
      fileName: state.activeFile.name,
      type: state.activeFile.type,
      currentPage: state.currentPage,
      totalPages: state.totalPages,
      zoom: state.zoom,
      viewportCenterX: state.viewportCenterX,
      viewportCenterY: state.viewportCenterY,
      autoPlaying: state.autoPlaying,
      autoPaused: state.autoPaused,
      autoElapsed: getAutoElapsedSeconds(false),
      autoDuration: state.autoCurrentDuration || getCurrentTimingSeconds(),
      timingMode: state.timingMode,
      globalTiming: getGlobalTimingSeconds(),
      currentTiming: getCurrentTimingSeconds(),
      perPageTiming: state.perPageTiming,
      timerVisible: state.timer.visible,
      timerOpacity: state.timer.opacity,
      timerSize: state.timer.size,
      timerMode: state.timer.mode,
      timerPosition: state.timer.position,
      countdownAlert: state.countdownAlert,
      countdownVoiceGender: state.countdownVoiceGender,
      countdownVoiceStart: getCountdownVoiceStart(),
      transitionEffect: state.transitionEffect,
      inkStrokes: state.inkStrokes.slice(-220),
      slideThumbsCount: state.remoteSlideThumbsCount || 0,
      slideThumbsReadyAt: state.remoteSlideThumbsReadyAt || 0,
      updatedAt: Date.now(),
    };

    try {
      await state.sessionRef.set(payload, { merge: true });
      scheduleRemotePreviewPublish(force ? 60 : 220);
    } catch (error) {
      console.warn('Could not publish session:', error);
    }
  }

  function scheduleRemotePreviewPublish(delay = 260) {
    if (!state.sessionRef || !state.activeFile) return;
    clearTimeout(state.remoteThumbTimer);
    state.remoteThumbTimer = setTimeout(async () => {
      try {
        const thumb = await getCurrentThumbForRemote();
        if (thumb && state.sessionRef) {
          await state.sessionRef.set({ thumb, thumbUpdatedAt: Date.now() }, { merge: true });
        }
      } catch (error) {}
    }, delay);
  }

  async function getCurrentThumbForRemote() {
    const key = `${state.activeFile ? state.activeFile.id : ''}:${state.currentPage}:base`;
    if (!state.remotePreviewBusy && state.lastRemoteThumbKey === key && !isPresentationFullscreen()) return '';
    state.remotePreviewBusy = true;
    try {
      state.lastRemoteThumbKey = key;
      if (state.activeFile.type === 'pdf' && state.activePdf) {
        return await renderPdfPageToDataUrl(state.activePdf, state.currentPage, 680, 0.72);
      }
      if (state.activeFile.type === 'canva') {
        return state.activeFile.thumbnail || createCanvaThumbDataUrl(state.activeFile.name);
      }
      if (state.activeFile.type === 'pptx') {
        if (state.pptxVisualReady && window.html2canvas && state.pptxRenderedSlides[state.currentPage - 1]) {
          const canvas = await window.html2canvas(state.pptxRenderedSlides[state.currentPage - 1], {
            backgroundColor: '#ffffff',
            scale: 0.9,
            logging: false,
            useCORS: true,
          });
          return canvas.toDataURL('image/jpeg', 0.86);
        }
        return createPptxThumbDataUrl(`Slide ${state.currentPage}`, state.currentPage);
      }
    } catch (error) {
      return '';
    } finally {
      state.remotePreviewBusy = false;
    }
    return '';
  }


  function removeRemoteFullscreenPrompt() {
    const old = document.getElementById('fullscreenRequestPrompt');
    if (old) old.remove();
  }

  function showRemoteFullscreenRequestPrompt() {
    removeRemoteFullscreenPrompt();
    if (!els.viewerView) return;
    const prompt = document.createElement('div');
    prompt.id = 'fullscreenRequestPrompt';
    prompt.className = 'fullscreen-request-prompt';
    prompt.innerHTML = `
      <div class="fullscreen-request-card">
        <strong>Phone requested fullscreen</strong>
        <span>Browsers require one tap on this desktop screen before entering fullscreen.</span>
        <div class="fullscreen-request-actions">
          <button id="fullscreenRequestEnter" type="button">Enter Fullscreen</button>
          <button id="fullscreenRequestDismiss" type="button">Dismiss</button>
        </div>
      </div>
    `;
    const root = document.fullscreenElement || els.viewerView || document.body;
    root.appendChild(prompt);
    revealToolbarTemporarily();
    const enter = prompt.querySelector('#fullscreenRequestEnter');
    const dismiss = prompt.querySelector('#fullscreenRequestDismiss');
    if (enter) enter.addEventListener('click', async () => {
      try { await toggleFullscreen(); } finally { removeRemoteFullscreenPrompt(); }
    });
    if (dismiss) dismiss.addEventListener('click', removeRemoteFullscreenPrompt);
    window.setTimeout(() => {
      if (document.getElementById('fullscreenRequestPrompt')) revealToolbarTemporarily();
    }, 120);
  }

  function handleRemoteFullscreenToggle() {
    if (isPresentationFullscreen() || document.fullscreenElement) {
      document.exitFullscreen().catch(() => {});
      return;
    }
    // A remote phone/Firebase command is not considered a user gesture by desktop browsers,
    // so entering fullscreen silently is blocked. Show a desktop-side one-tap prompt instead.
    showRemoteFullscreenRequestPrompt();
  }

  async function applyRemoteCommand(command) {
    switch (command.action) {
      case 'next': nextPage(); break;
      case 'prev': previousPage(); break;
      case 'first': jumpToPage(1); break;
      case 'last': jumpToPage(state.totalPages); break;
      case 'jumpTo': jumpToPage(command.value); break;
      case 'requestSlideThumbs': generateRemoteSlideThumbs(); break;
      case 'addInkStroke': addInkStroke(command.value); break;
      case 'clearInk': clearInkStrokes(); break;
      case 'undoInk': undoInkStroke(); break;
      case 'eraseInkAt': eraseInkAt(command.value); break;
      case 'zoomIn': setZoom(state.zoom + 0.1); break;
      case 'zoomOut': setZoom(state.zoom - 0.1); break;
      case 'resetZoom': setZoom(1, { centerX: 0.5, centerY: 0.5 }); break;
      case 'toggleFullscreen': handleRemoteFullscreenToggle(); break;
      case 'setZoom': setZoom(Number(command.value) || 1); break;
      case 'setViewport': setViewportTransform(command.value || {}); break;
      case 'autoStart': startAutoPlay(); break;
      case 'autoPause': pauseAutoPlay(); break;
      case 'autoStop': stopAutoPlay(); break;
      case 'setGlobalTiming':
        state.timingMode = 'global';
        updateTimingModeUI();
        setGlobalTimingSeconds(command.value);
        publishSessionState();
        break;
      case 'setCurrentSlideTiming':
        setCurrentSlideTimingSeconds(command.value);
        publishSessionState();
        break;
      case 'setTimingMode':
        state.timingMode = command.value === 'per-slide' ? 'per-slide' : 'global';
        updateTimingModeUI();
        resetAutoClockForCurrentSlide();
        updateTimerText();
        publishSessionState();
        break;
      case 'setTiming':
        state.timingMode = 'global';
        updateTimingModeUI();
        setGlobalTimingSeconds(command.value);
        publishSessionState();
        break;
      case 'timerShow': showTimer(); break;
      case 'timerHide': hideTimer(); break;
      case 'timerReset': resetTimer(); break;
      case 'setTimerPosition':
        state.timer.position = ['bottom-right', 'bottom-left', 'top-right', 'top-left'].includes(command.value) ? command.value : 'bottom-right';
        applyTimerSettings();
        publishSessionState();
        break;
      case 'setTimerMode':
        state.timer.mode = command.value === 'down' ? 'down' : 'up';
        resetTimer();
        publishSessionState();
        break;
      case 'setTimerOpacity':
        state.timer.opacity = Number(command.value);
        applyTimerSettings();
        publishSessionState();
        break;
      case 'setTimerSize':
        state.timer.size = Math.max(16, Math.min(72, Number(command.value) || 28));
        applyTimerSettings();
        publishSessionState();
        break;
      case 'setCountdownAlert':
        state.countdownAlert = ['off', 'sound', 'voice', 'both'].includes(command.value) ? command.value : 'off';
        state.lastCountdownAlertSecond = null;
        if (state.countdownAlert !== 'off') primePresentationAudio();
        applyTimerSettings();
        publishSessionState();
        break;
      case 'setCountdownVoiceGender':
        state.countdownVoiceGender = normalizeCountdownVoiceStyle(command.value);
        primePresentationAudio();
        applyTimerSettings();
        publishSessionState();
        break;
      case 'setCountdownVoiceStart':
        state.countdownVoiceStart = Number(command.value) === 3 ? 3 : 5;
        state.lastCountdownAlertSecond = null;
        primePresentationAudio();
        applyTimerSettings();
        publishSessionState();
        break;
      case 'testCountdownAlert': {
        const mode = ['sound', 'voice', 'both'].includes(command.value) ? command.value : (state.countdownAlert === 'off' ? 'both' : state.countdownAlert);
        primePresentationAudio();
        runCountdownAlertPreview(mode);
        break;
      }
      case 'setTransitionEffect':
        state.transitionEffect = command.value || 'fade';
        applyTimerSettings();
        publishSessionState();
        break;
      default: break;
    }
  }


  function normalizeStroke(raw = {}) {
    const page = Math.min(Math.max(1, Number(raw.page) || state.currentPage || 1), state.totalPages || 1);
    const tool = raw.tool === 'highlighter' ? 'highlighter' : 'pen';
    const points = Array.isArray(raw.points)
      ? raw.points.map((pt) => ({ x: clamp01(pt.x), y: clamp01(pt.y) })).filter((pt) => Number.isFinite(pt.x) && Number.isFinite(pt.y))
      : [];
    if (points.length < 2) return null;
    return {
      id: raw.id || `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      page,
      tool,
      color: typeof raw.color === 'string' && raw.color ? raw.color : (tool === 'highlighter' ? '#facc15' : '#ef4444'),
      size: Math.max(2, Math.min(28, Number(raw.size) || (tool === 'highlighter' ? 18 : 6))),
      points,
      createdAt: Number(raw.createdAt) || Date.now(),
    };
  }

  function addInkStroke(raw) {
    const stroke = normalizeStroke(raw);
    if (!stroke) return;
    state.inkStrokes = state.inkStrokes.filter((item) => item.id !== stroke.id).concat(stroke).slice(-350);
    renderInkOverlay();
    publishSessionState(true);
  }

  function clearInkStrokes(pageOnly = true) {
    if (pageOnly) state.inkStrokes = state.inkStrokes.filter((stroke) => stroke.page !== state.currentPage);
    else state.inkStrokes = [];
    renderInkOverlay();
    publishSessionState(true);
  }

  function undoInkStroke() {
    const idx = [...state.inkStrokes].map((stroke, index) => ({ stroke, index })).reverse().find((item) => item.stroke.page === state.currentPage)?.index;
    if (idx === undefined) return;
    state.inkStrokes.splice(idx, 1);
    renderInkOverlay();
    publishSessionState(true);
  }

  function eraseInkAt(value = {}) {
    const page = Math.min(Math.max(1, Number(value.page) || state.currentPage), state.totalPages || 1);
    const x = clamp01(value.x);
    const y = clamp01(value.y);
    const radius = Math.max(0.01, Math.min(0.09, Number(value.radius) || 0.035));
    const before = state.inkStrokes.length;
    state.inkStrokes = state.inkStrokes.filter((stroke) => {
      if (stroke.page !== page) return true;
      return !stroke.points.some((pt) => Math.hypot(pt.x - x, pt.y - y) <= radius);
    });
    if (state.inkStrokes.length !== before) {
      renderInkOverlay();
      publishSessionState(true);
    }
  }

  function getActiveSlideSurface() {
    if (state.activeFile && state.activeFile.type === 'pdf') return state.activePdfCanvas || els.pdfCanvas;
    if (state.activeFile && state.activeFile.type === 'pptx') return state.pptxRenderedSlides[state.currentPage - 1] || els.pptxSlide;
    if (state.activeFile && state.activeFile.type === 'canva') return els.canvaFrame;
    return null;
  }

  function renderInkOverlay() {
    const canvas = els.inkCanvas;
    const wrap = els.viewerCanvasWrap;
    const surface = getActiveSlideSurface();
    if (!canvas || !wrap || !surface || !surface.isConnected || !state.inkStrokes.length) {
      if (canvas) canvas.classList.add('hidden');
      return;
    }
    const wrapRect = wrap.getBoundingClientRect();
    const surfRect = surface.getBoundingClientRect();
    const dpr = Math.max(1, Math.min(2.5, window.devicePixelRatio || 1));
    const width = Math.max(1, wrap.scrollWidth || wrap.clientWidth);
    const height = Math.max(1, wrap.scrollHeight || wrap.clientHeight);
    canvas.classList.remove('hidden');
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;
    canvas.width = Math.round(width * dpr);
    canvas.height = Math.round(height * dpr);
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, width, height);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    const offsetX = surfRect.left - wrapRect.left + wrap.scrollLeft;
    const offsetY = surfRect.top - wrapRect.top + wrap.scrollTop;
    const surfW = Math.max(1, surfRect.width);
    const surfH = Math.max(1, surfRect.height);
    state.inkStrokes.filter((stroke) => stroke.page === state.currentPage).forEach((stroke) => {
      if (!stroke.points || stroke.points.length < 2) return;
      ctx.save();
      ctx.globalAlpha = stroke.tool === 'highlighter' ? 0.35 : 0.95;
      ctx.globalCompositeOperation = stroke.tool === 'highlighter' ? 'multiply' : 'source-over';
      ctx.strokeStyle = stroke.color || (stroke.tool === 'highlighter' ? '#facc15' : '#ef4444');
      ctx.lineWidth = Math.max(2, Number(stroke.size) || 6);
      ctx.beginPath();
      stroke.points.forEach((pt, index) => {
        const x = offsetX + pt.x * surfW;
        const y = offsetY + pt.y * surfH;
        if (index === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.stroke();
      ctx.restore();
    });
  }

  async function generateRemoteSlideThumbs() {
    if (!state.sessionRef || !state.activeFile || state.remoteSlideThumbsBusy) return;
    state.remoteSlideThumbsBusy = true;
    // Keep this safe for Firestore's 1 MiB document limit and for phone speed.
    // For classroom decks this shows real slide previews; for huge PDFs it still
    // allows number-jump placeholders after the cap.
    const maxPages = Math.min(state.totalPages || 1, 60);
    const lite = {};
    try {
      let thumbsRef = null;
      try { thumbsRef = state.sessionRef.collection('slideThumbs'); } catch (error) { thumbsRef = null; }
      for (let page = 1; page <= maxPages; page++) {
        let src = '';
        if (state.activeFile.type === 'pdf' && state.activePdf) {
          src = await renderPdfPageToDataUrl(state.activePdf, page, 120, 0.34);
        } else if (state.activeFile.type === 'pptx') {
          src = createPptxThumbDataUrl(`Slide ${page}`, page);
        } else if (state.activeFile.type === 'canva') {
          src = state.activeFile.thumbnail || createCanvaThumbDataUrl(state.activeFile.name);
        }
        if (!src) continue;
        lite[String(page)] = src;
        // Try subcollection for projects with broader rules, but never depend on it.
        if (thumbsRef) {
          try { await thumbsRef.doc(String(page)).set({ page, src, updatedAt: Date.now() }, { merge: true }); } catch (error) { thumbsRef = null; }
        }
        if (page === 1 || page % 4 === 0 || page === maxPages) {
          state.remoteSlideThumbsCount = page;
          state.remoteSlideThumbsReadyAt = Date.now();
          await state.sessionRef.set({
            slideThumbsLite: lite,
            slideThumbsCount: page,
            slideThumbsTotal: maxPages,
            slideThumbsReadyAt: state.remoteSlideThumbsReadyAt
          }, { merge: true });
        }
      }
      if (maxPages < (state.totalPages || 1)) {
        await state.sessionRef.set({ slideThumbsTruncated: true }, { merge: true });
      }
    } catch (error) {
      console.warn('Could not generate all slide thumbnails:', error);
    } finally {
      state.remoteSlideThumbsBusy = false;
    }
  }

  async function openQrModal() {
    if (!state.activeFile) return;
    if (!state.firebaseReady) {
      els.qrHelp.textContent = 'Phone remote across devices needs Firebase config. You can still present locally. Click Remote setup on the home screen to see where to add your Firebase keys.';
    } else {
      await setupRemoteSessionIfPossible();
      els.qrHelp.textContent = `Session ${state.sessionId} is live. Scan the CONTROL QR for buttons. Viewer QR is preview-only / view-only.`;
    }

    const session = state.sessionId || 'NO-FIREBASE';
    const baseUrl = window.location.href.split('?')[0].split('#')[0];
    const hostUrl = `${baseUrl}?remote=1&session=${encodeURIComponent(session)}&role=host&screen=controls`;
    const viewerUrl = `${baseUrl}?remote=1&session=${encodeURIComponent(session)}&role=viewer&screen=preview`;

    els.hostRemoteLink.value = hostUrl;
    els.viewerRemoteLink.value = viewerUrl;
    els.hostQr.innerHTML = '';
    els.viewerQr.innerHTML = '';
    new QRCode(els.hostQr, { text: hostUrl, width: 210, height: 210 });
    new QRCode(els.viewerQr, { text: viewerUrl, width: 210, height: 210 });
    showModal(els.qrModal);
  }

  function renderRemoteApp() {
    document.documentElement.classList.add('presentation-hub-remote-mode');
    document.body.classList.add('presentation-hub-remote-mode');
    document.body.classList.remove('remote-fullscreen-open');
    if (document.fullscreenElement) { try { document.exitFullscreen(); } catch (error) {} }
    els.app.classList.add('hidden');
    els.remoteApp.classList.remove('hidden');
    const sessionId = qs.get('session') || '';
    const requestedRole = (qs.get('role') || 'host').toLowerCase();
    const role = requestedRole === 'viewer' ? 'viewer' : 'host';
    const isHost = role === 'host';
    let remoteViewport = { zoom: 1, centerX: 0.5, centerY: 0.5 };
    let latestThumb = '';
    let latestRemoteData = null;
    let remoteSlideThumbs = {};
    let remoteSlideThumbsLoading = false;
    let remoteLastThumbsReadyAt = 0;

    els.remoteApp.innerHTML = `
      <main id="remoteControlDashboard" class="remote-card remote-premium-card ${isHost ? '' : 'remote-viewer-only'}">
        <div class="remote-head remote-premium-head">
          <div class="remote-title-block">
            <span class="remote-eyebrow">${isHost ? 'HOST REMOTE' : 'VIEW ONLY'}</span>
            <h1>Presentation Hub Pro</h1>
            <p class="remote-sub">${escapeHtml(sessionId || 'No session')} • ${isHost ? 'Host control' : 'Preview only'} • low-latency mode</p>
          </div>
          <span id="remoteStatusPill" class="remote-status-pill">Connecting</span>
        </div>

        <section class="remote-now-card">
          <div class="remote-slide-meta">
            <div>
              <span class="remote-mini-label">Now showing</span>
              <h2 id="remoteSlideLabel">Slide -- / --</h2>
            </div>
            <p id="remoteAutoLabel" class="remote-sub remote-auto-label">Auto Play idle</p>
          </div>
          <div class="remote-preview-shell">
            <div class="remote-preview" id="remotePreview"><span>Waiting for presentation...</span></div>
            <div class="remote-preview-actions">
              <button id="remotePreviewFullBtn" class="remote-preview-full-btn" data-host-only="false">Portrait Preview</button>
              <button id="remoteAllSlidesBtn" class="remote-preview-full-btn remote-all-slides-btn" data-host-only="true">All Slides</button>
            </div>
          </div>
          <p id="remoteFileLabel" class="remote-sub remote-file-label">Connect to an active desktop session.</p>
        </section>

        <div class="remote-control-banner ${isHost ? '' : 'viewer'}">${isHost ? 'Host control active. Tap commands first; preview updates after for lower delay.' : 'Viewer mode: preview is visible, but controls are disabled.'}</div>
        <p class="remote-hint">Portrait preview is supported. Pinch and drag only after opening the preview screen.</p>

        <section class="remote-nav-pad" aria-label="Presentation navigation">
          <button class="remote-small-action" data-command="first" data-host-only="true">First</button>
          <button class="remote-main-action remote-prev-action" data-command="prev" data-host-only="true">← Prev</button>
          <button class="remote-main-action remote-next-action" data-command="next" data-host-only="true">Next →</button>
          <button class="remote-small-action" data-command="last" data-host-only="true">Last</button>
        </section>

        <section class="remote-section remote-premium-section">
          <div class="remote-section-title"><span>Presentation</span><small>Desktop screen</small></div>
          <div class="remote-control-row remote-segment-row">
            <button data-command="toggleFullscreen" data-host-only="true">Fullscreen</button>
            <button data-command="resetZoom" data-host-only="true">Fit Slide</button>
            <button data-command="timerReset" data-host-only="true">Reset Time</button>
          </div>
          <div class="remote-wide remote-timing-box remote-form-grid two">
            <select id="remoteTransition" data-host-only="true">
              <option value="fade">Fade transition</option>
              <option value="slide-left">Slide left</option>
              <option value="slide-right">Slide right</option>
              <option value="slide-up">Slide up</option>
              <option value="zoom-in">Zoom in</option>
              <option value="zoom-out">Zoom out</option>
              <option value="soft-blur">Soft blur</option>
              <option value="none">No transition</option>
            </select>
            <select id="remoteTimerPosition" data-host-only="true">
              <option value="bottom-right">Timer bottom right</option>
              <option value="bottom-left">Timer bottom left</option>
              <option value="top-right">Timer top right</option>
              <option value="top-left">Timer top left</option>
            </select>
          </div>
        </section>

        <section class="remote-section remote-premium-section">
          <div class="remote-section-title"><span>Zoom</span><small>Desktop viewer</small></div>
          <div class="remote-control-row remote-segment-row">
            <button data-command="zoomOut" data-host-only="true">− Out</button>
            <button data-command="resetZoom" data-host-only="true">Reset</button>
            <button data-command="zoomIn" data-host-only="true">+ In</button>
          </div>
        </section>

        <section class="remote-section remote-premium-section">
          <div class="remote-section-title"><span>Auto Play</span><small>Global or per-slide</small></div>
          <div class="remote-control-row remote-segment-row">
            <button data-command="autoStart" data-host-only="true">Start</button>
            <button data-command="autoPause" data-host-only="true">Pause</button>
            <button data-command="autoStop" data-host-only="true">Stop</button>
          </div>
          <div class="remote-wide remote-timing-box remote-form-grid">
            <select id="remoteTimingMode" data-host-only="true">
              <option value="global">Global - all slides</option>
              <option value="per-slide">Per-slide - current slide</option>
            </select>
            <input id="remoteTiming" type="number" min="1" value="10" placeholder="Seconds">
            <button id="remoteSetTiming" data-host-only="true">Apply Timing</button>
          </div>
          <div class="remote-wide remote-timing-box remote-form-grid two">
            <select id="remoteCountdownAlert" data-host-only="true">
              <option value="off">Alert off</option>
              <option value="sound">Sound only</option>
              <option value="voice">Voice count</option>
              <option value="both">Sound + voice</option>
            </select>
            <select id="remoteCountdownVoice" data-host-only="true">
              <option value="soft-female">Soft Female</option>
              <option value="bright-female">Bright Female</option>
              <option value="calm-male">Calm Male</option>
              <option value="deep-male">Deep Male</option>
              <option value="teacher">Teacher Voice</option>
              <option value="announcer">Announcer Voice</option>
            </select>
            <select id="remoteCountdownVoiceStart" data-host-only="true">
              <option value="5">Voice from 5 sec</option>
              <option value="3">Voice from 3 sec</option>
            </select>
            <button id="remoteTestAlert" data-host-only="true">Test Sound/Voice</button>
          </div>
        </section>

        <section class="remote-section remote-premium-section">
          <div class="remote-section-title"><span>Timer</span><small>Overlay controls</small></div>
          <div class="remote-control-row remote-segment-row">
            <button data-command="timerShow" data-host-only="true">Show</button>
            <button data-command="timerHide" data-host-only="true">Hide</button>
            <button data-command="timerReset" data-host-only="true">Reset</button>
          </div>
          <div class="remote-wide remote-timing-box remote-form-grid two">
            <select id="remoteTimerMode" data-host-only="true">
              <option value="up">Timer count up</option>
              <option value="down">Timer countdown</option>
            </select>
            <button data-command="timerReset" data-host-only="true">Sync Timer Reset</button>
          </div>
          <div class="remote-slider-stack remote-premium-sliders">
            <label>Opacity <span id="remoteOpacityLabel">75%</span>
              <input id="remoteOpacity" type="range" min="0" max="100" step="1" value="75" data-host-only="true">
            </label>
            <label>Timer size <span id="remoteTimerSizeLabel">28px</span>
              <input id="remoteTimerSize" type="range" min="16" max="72" step="1" value="28" data-host-only="true">
            </label>
          </div>
        </section>
      </main>
      <div id="remoteSlidesPanel" class="remote-slides-panel hidden" data-host-only="true">
        <div class="remote-slides-head">
          <div>
            <strong>All Slides</strong>
            <span id="remoteSlidesHelp">Choose a slide to show on desktop.</span>
          </div>
          <button id="remoteCloseSlides" type="button">Close</button>
        </div>
        <div id="remoteSlidesGrid" class="remote-slides-grid"></div>
      </div>

      <div id="remoteFullPreview" class="remote-full-preview remote-portrait-preview hidden">
        <div class="remote-full-toolbar">
          <span id="remoteFullLabel">Slide Preview</span>
          <button id="remoteBackToControls" class="remote-back-controls">Controls</button>
          <button id="remoteClosePreview">Close</button>
        </div>
        <div class="remote-draw-toolbar" data-host-only="true">
          <button class="active" id="remoteToolMove" type="button">Move</button>
          <button id="remoteToolPen" type="button">Pen</button>
          <button id="remoteToolHighlighter" type="button">Highlight</button>
          <button id="remoteToolErase" type="button">Erase</button>
          <button id="remoteUndoInk" type="button">Undo</button>
          <button id="remoteClearInk" type="button">Clear</button>
        </div>
        <div id="remoteFullStage" class="remote-full-stage remote-monitor-stage">
          <div class="remote-monitor-frame" aria-label="Desktop monitor preview">
            <div class="remote-monitor-camera"></div>
            <div class="remote-monitor-screen">
              <img id="remoteFullImg" alt="Desktop monitor preview of current slide" />
              <canvas id="remoteInkPreview" class="remote-ink-preview" aria-hidden="true"></canvas>
            </div>
            <div class="remote-monitor-neck"></div>
            <div class="remote-monitor-base"></div>
          </div>
        </div>
        <div class="remote-full-help">Portrait preview shows the desktop monitor look. Pinch inside the screen, then drag to choose the exact desktop area.</div>
      </div>
    `;

    if (!isHost) {
      els.remoteApp.querySelectorAll('[data-host-only="true"]').forEach((node) => {
        node.setAttribute('aria-disabled', 'true');
        if ('disabled' in node) node.disabled = true;
      });
    }

    if (!hasFirebaseConfig()) {
      $('remoteFileLabel').textContent = 'Remote needs firebase-config.js because phones need realtime sync across devices.';
      $('remoteStatusPill').textContent = 'Setup needed';
      return;
    }

    initFirebaseIfConfigured().then(() => {
      if (!state.firebaseReady || !sessionId) {
        $('remoteFileLabel').textContent = 'Could not connect to Firebase or missing session.';
        $('remoteStatusPill').textContent = 'Offline';
        return;
      }
      const ref = state.firebaseDb.collection(SESSION_COLLECTION).doc(sessionId);
      ref.onSnapshot((snap) => {
        if (!snap.exists) {
          $('remoteFileLabel').textContent = 'Session not found. Open QR from the desktop viewer first.';
          $('remoteStatusPill').textContent = 'No session';
          return;
        }
        const data = snap.data();
        latestRemoteData = data;
        window.__phRemoteCurrentPage = Number(data.currentPage) || 1;
        window.__phRemoteInkStrokes = Array.isArray(data.inkStrokes) ? data.inkStrokes : [];
        $('remoteStatusPill').textContent = isHost ? 'Host' : 'Viewer';
        $('remoteSlideLabel').textContent = `${data.type === 'pdf' ? 'Page' : data.type === 'canva' ? 'Canva' : 'Slide'} ${data.currentPage || '--'} / ${data.totalPages || '--'}`;
        $('remoteFullLabel').textContent = `${data.type === 'pdf' ? 'Page' : data.type === 'canva' ? 'Canva' : 'Slide'} ${data.currentPage || '--'} / ${data.totalPages || '--'}`;
        $('remoteFileLabel').textContent = data.fileName || 'Active presentation';
        if (data.thumb) latestThumb = data.thumb;
        const preview = $('remotePreview');
        preview.innerHTML = latestThumb ? `<img src="${latestThumb}" alt="Current slide preview">` : '<span>No preview yet</span>';
        const fullImg = $('remoteFullImg');
        if (fullImg && latestThumb && fullImg.src !== latestThumb) fullImg.src = latestThumb;
        if (!window.__presentationHubRemoteGestureActive) {
          remoteViewport = {
            zoom: Number(data.zoom) || 1,
            centerX: Number.isFinite(Number(data.viewportCenterX)) ? Number(data.viewportCenterX) : 0.5,
            centerY: Number.isFinite(Number(data.viewportCenterY)) ? Number(data.viewportCenterY) : 0.5,
          };
        }
        $('remoteTimingMode').value = data.timingMode || 'global';
        $('remoteTiming').value = (data.timingMode === 'per-slide' ? data.currentTiming : data.globalTiming) || 10;
        if ($('remoteCountdownAlert')) $('remoteCountdownAlert').value = data.countdownAlert || 'off';
        if ($('remoteCountdownVoice')) $('remoteCountdownVoice').value = normalizeCountdownVoiceStyle(data.countdownVoiceGender);
        if ($('remoteCountdownVoiceStart')) $('remoteCountdownVoiceStart').value = String(Number(data.countdownVoiceStart) === 3 ? 3 : 5);
        if ($('remoteTransition')) $('remoteTransition').value = data.transitionEffect || 'fade';
        if ($('remoteTimerPosition')) $('remoteTimerPosition').value = data.timerPosition || 'bottom-right';
        if ($('remoteTimerMode')) $('remoteTimerMode').value = data.timerMode || 'up';
        $('remoteOpacity').value = data.timerOpacity ?? 75;
        if ($('remoteOpacityLabel')) $('remoteOpacityLabel').textContent = `${data.timerOpacity ?? 75}%`;
        $('remoteTimerSize').value = data.timerSize ?? 28;
        if ($('remoteTimerSizeLabel')) $('remoteTimerSizeLabel').textContent = `${data.timerSize ?? 28}px`;
        $('remoteAutoLabel').textContent = data.autoPlaying
          ? `Auto Play: ${data.autoElapsed || 0}s / ${data.autoDuration || data.currentTiming || data.globalTiming || 10}s`
          : (data.autoPaused ? `Auto Play paused at ${data.autoElapsed || 0}s` : 'Auto Play idle');
        applyRemoteFullPreviewTransform(remoteViewport);
        if (data.slideThumbsLite && typeof data.slideThumbsLite === 'object') {
          // Reliable fallback for existing simple Firestore rules: thumbnails are
          // stored as tiny data URLs on the session document only after the host
          // taps All Slides. This avoids blank grids when subcollections are not
          // allowed by the user's current Firebase rules.
          remoteSlideThumbs = { ...remoteSlideThumbs, ...data.slideThumbsLite };
        }
        if (data.slideThumbsReadyAt && data.slideThumbsReadyAt !== remoteLastThumbsReadyAt) {
          remoteLastThumbsReadyAt = data.slideThumbsReadyAt;
          if ($('remoteSlidesPanel') && !$('remoteSlidesPanel').classList.contains('hidden')) {
            loadRemoteSlideThumbs(ref, true).then((thumbs) => {
              remoteSlideThumbs = { ...remoteSlideThumbs, ...thumbs };
              renderRemoteSlidesGrid(latestRemoteData || {}, ref, isHost, remoteSlideThumbs);
            });
          }
        }
        renderRemoteSlidesGrid(latestRemoteData || {}, ref, isHost, remoteSlideThumbs);
        renderRemoteInkPreview(window.__phRemoteInkStrokes, window.__phRemoteCurrentPage, remoteViewport);
      });

      els.remoteApp.querySelectorAll('[data-command]').forEach((button) => {
        button.addEventListener('click', () => {
          if (!isHost) return;
          sendRemoteCommand(ref, button.dataset.command);
        });
      });

      const allSlidesBtn = $('remoteAllSlidesBtn');
      const slidesPanel = $('remoteSlidesPanel');
      const closeSlidesBtn = $('remoteCloseSlides');
      if (allSlidesBtn && slidesPanel) {
        allSlidesBtn.addEventListener('click', () => {
          if (!isHost) return;
          slidesPanel.classList.remove('hidden');
          renderRemoteSlidesGrid(latestRemoteData || {}, ref, isHost, remoteSlideThumbs);
          loadRemoteSlideThumbs(ref).then((thumbs) => {
            remoteSlideThumbs = { ...remoteSlideThumbs, ...thumbs };
            renderRemoteSlidesGrid(latestRemoteData || {}, ref, isHost, remoteSlideThumbs);
          });
          sendRemoteCommand(ref, 'requestSlideThumbs');
          window.setTimeout(() => loadRemoteSlideThumbs(ref, true).then((thumbs) => {
            remoteSlideThumbs = { ...remoteSlideThumbs, ...thumbs };
            renderRemoteSlidesGrid(latestRemoteData || {}, ref, isHost, remoteSlideThumbs);
          }), 1400);
        });
      }
      if (closeSlidesBtn && slidesPanel) {
        closeSlidesBtn.addEventListener('click', () => {
          slidesPanel.classList.add('hidden');
        });
      }

      // Keep button handlers direct and single-fire. Heavy All Slides data is loaded separately
      // so normal host commands stay fast.

      attachRemotePreviewControls(ref, isHost, () => remoteViewport, (next) => {
        remoteViewport = next;
        applyRemoteFullPreviewTransform(remoteViewport);
        renderRemoteSlidesGrid(latestRemoteData || {}, ref, isHost, remoteSlideThumbs);
        renderRemoteInkPreview(window.__phRemoteInkStrokes, window.__phRemoteCurrentPage, remoteViewport);
      });
      $('remoteTimingMode').addEventListener('change', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setTimingMode', $('remoteTimingMode').value);
      });
      $('remoteSetTiming').addEventListener('click', () => {
        if (!isHost) return;
        const seconds = Number($('remoteTiming').value) || 10;
        const mode = $('remoteTimingMode').value;
        sendRemoteCommand(ref, mode === 'per-slide' ? 'setCurrentSlideTiming' : 'setGlobalTiming', seconds);
      });
      if ($('remoteCountdownAlert')) $('remoteCountdownAlert').addEventListener('change', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setCountdownAlert', $('remoteCountdownAlert').value);
      });
      if ($('remoteCountdownVoice')) $('remoteCountdownVoice').addEventListener('change', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setCountdownVoiceGender', $('remoteCountdownVoice').value);
      });
      if ($('remoteCountdownVoiceStart')) $('remoteCountdownVoiceStart').addEventListener('change', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setCountdownVoiceStart', Number($('remoteCountdownVoiceStart').value) === 3 ? 3 : 5);
      });
      if ($('remoteTestAlert')) $('remoteTestAlert').addEventListener('click', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'testCountdownAlert', $('remoteCountdownAlert') ? $('remoteCountdownAlert').value : 'both');
      });
      if ($('remoteTransition')) $('remoteTransition').addEventListener('change', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setTransitionEffect', $('remoteTransition').value);
      });
      if ($('remoteTimerPosition')) $('remoteTimerPosition').addEventListener('change', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setTimerPosition', $('remoteTimerPosition').value);
      });
      if ($('remoteTimerMode')) $('remoteTimerMode').addEventListener('change', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setTimerMode', $('remoteTimerMode').value);
      });
      $('remoteOpacity').addEventListener('input', () => {
        if (!isHost) return;
        const value = Number($('remoteOpacity').value);
        if ($('remoteOpacityLabel')) $('remoteOpacityLabel').textContent = `${value}%`;
        sendRemoteCommand(ref, 'setTimerOpacity', value);
      });
      $('remoteTimerSize').addEventListener('input', () => {
        if (!isHost) return;
        const value = Number($('remoteTimerSize').value);
        if ($('remoteTimerSizeLabel')) $('remoteTimerSizeLabel').textContent = `${value}px`;
        sendRemoteCommand(ref, 'setTimerSize', value);
      });
    });
  }


  async function loadRemoteSlideThumbs(ref, force = false) {
    if (!ref || remoteSlideThumbsLoading) return remoteSlideThumbs || {};
    if (!force && remoteSlideThumbs && Object.keys(remoteSlideThumbs).length) return remoteSlideThumbs;
    remoteSlideThumbsLoading = true;
    try {
      const snap = await ref.collection('slideThumbs').orderBy('page').limit(180).get();
      const next = {};
      snap.forEach((doc) => {
        const item = doc.data() || {};
        if (item.page && item.src) next[item.page] = item.src;
      });
      return next;
    } catch (error) {
      console.warn('Could not load slide thumbnails:', error);
      return remoteSlideThumbs || {};
    } finally {
      remoteSlideThumbsLoading = false;
    }
  }

  function renderRemoteSlidesGrid(data = {}, ref, isHost, localThumbs = {}) {
    const grid = $('remoteSlidesGrid');
    const panel = $('remoteSlidesPanel');
    if (!grid || !panel || panel.classList.contains('hidden')) return;
    const total = Math.max(1, Number(data.totalPages) || 1);
    const current = Math.max(1, Number(data.currentPage) || 1);
    const thumbs = localThumbs || {};
    const availableCount = Object.keys(thumbs).length || Number(data.slideThumbsCount) || 0;
    const help = $('remoteSlidesHelp');
    if (help) help.textContent = availableCount
      ? `Tap any slide. Thumbnails loaded: ${Math.min(availableCount, total)} / ${total}.`
      : 'Loading real slide previews from desktop... numbers still work while waiting.';
    const limit = Math.min(total, 180);
    const parts = [];
    for (let i = 1; i <= limit; i++) {
      const src = thumbs[String(i)] || thumbs[i] || '';
      parts.push(`
        <button class="remote-slide-choice${i === current ? ' active' : ''}" data-jump-slide="${i}" type="button">
          <span class="remote-slide-num">${i}</span>
          ${src ? `<img src="${src}" alt="Slide ${i} thumbnail">` : `<span class="remote-slide-placeholder">${i}</span>`}
        </button>
      `);
    }
    if (total > limit) {
      parts.push(`<div class="remote-slide-note">Showing first ${limit} of ${total}. Use First/Last/Next for far pages.</div>`);
    }
    grid.innerHTML = parts.join('');
    grid.querySelectorAll('[data-jump-slide]').forEach((button) => {
      button.addEventListener('click', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'jumpTo', Number(button.dataset.jumpSlide));
      });
    });
  }

  function renderRemoteInkPreview(strokes = [], page = 1, viewport = {}) {
    const canvas = $('remoteInkPreview');
    const img = $('remoteFullImg');
    const screen = img ? (img.closest('.remote-monitor-screen') || $('remoteFullStage')) : null;
    if (!canvas || !img || !screen) return;
    const screenW = Math.max(1, screen.clientWidth || screen.getBoundingClientRect().width || 1);
    const screenH = Math.max(1, screen.clientHeight || screen.getBoundingClientRect().height || 1);
    const dpr = Math.max(1, Math.min(2, window.devicePixelRatio || 1));
    canvas.style.width = `${screenW}px`;
    canvas.style.height = `${screenH}px`;
    canvas.width = Math.round(screenW * dpr);
    canvas.height = Math.round(screenH * dpr);
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, screenW, screenH);
    const left = Number(img.dataset.remoteLeft) || 0;
    const top = Number(img.dataset.remoteTop) || 0;
    const contentW = Number(img.dataset.remoteContentW) || screenW;
    const contentH = Number(img.dataset.remoteContentH) || screenH;
    const all = Array.isArray(strokes) ? strokes.slice() : [];
    if (window.__phRemoteTempStroke) all.push(window.__phRemoteTempStroke);
    all.filter((stroke) => Number(stroke.page) === Number(page)).forEach((stroke) => {
      if (!stroke.points || stroke.points.length < 2) return;
      ctx.save();
      ctx.globalAlpha = stroke.tool === 'highlighter' ? 0.35 : 0.95;
      ctx.globalCompositeOperation = stroke.tool === 'highlighter' ? 'multiply' : 'source-over';
      ctx.strokeStyle = stroke.color || (stroke.tool === 'highlighter' ? '#facc15' : '#ef4444');
      ctx.lineWidth = Math.max(2, Math.min(22, Number(stroke.size) || 6));
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.beginPath();
      stroke.points.forEach((pt, index) => {
        const x = left + clamp01(pt.x) * contentW;
        const y = top + clamp01(pt.y) * contentH;
        if (index === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.stroke();
      ctx.restore();
    });
  }

  function remoteScreenPointToSlide(clientX, clientY) {
    const img = $('remoteFullImg');
    const screen = img ? (img.closest('.remote-monitor-screen') || $('remoteFullStage')) : null;
    if (!img || !screen) return { x: 0.5, y: 0.5 };
    const rect = screen.getBoundingClientRect();
    const left = Number(img.dataset.remoteLeft) || 0;
    const top = Number(img.dataset.remoteTop) || 0;
    const contentW = Math.max(1, Number(img.dataset.remoteContentW) || rect.width || 1);
    const contentH = Math.max(1, Number(img.dataset.remoteContentH) || rect.height || 1);
    return {
      x: clamp01(((clientX - rect.left) - left) / contentW),
      y: clamp01(((clientY - rect.top) - top) / contentH),
    };
  }

  function applyRemoteFullPreviewTransform(viewport) {
    const img = $('remoteFullImg');
    if (!img) return;
    const screen = img.closest('.remote-monitor-screen') || $('remoteFullStage');
    if (!screen) return;

    const zoom = Math.min(4, Math.max(1, Number(viewport.zoom) || 1));
    let centerX = Math.max(0, Math.min(1, Number(viewport.centerX) || 0.5));
    let centerY = Math.max(0, Math.min(1, Number(viewport.centerY) || 0.5));
    const screenRect = screen.getBoundingClientRect();
    const screenW = Math.max(1, screen.clientWidth || screenRect.width || 1);
    const screenH = Math.max(1, screen.clientHeight || screenRect.height || 1);
    const naturalW = Math.max(1, img.naturalWidth || 16);
    const naturalH = Math.max(1, img.naturalHeight || 9);
    const aspect = naturalW / naturalH;

    let fitW = screenW;
    let fitH = fitW / aspect;
    if (fitH > screenH) {
      fitH = screenH;
      fitW = fitH * aspect;
    }

    const contentW = fitW * zoom;
    const contentH = fitH * zoom;
    const minCenterX = contentW > screenW ? screenW / (2 * contentW) : 0.5;
    const maxCenterX = contentW > screenW ? 1 - minCenterX : 0.5;
    const minCenterY = contentH > screenH ? screenH / (2 * contentH) : 0.5;
    const maxCenterY = contentH > screenH ? 1 - minCenterY : 0.5;
    centerX = Math.max(minCenterX, Math.min(maxCenterX, centerX));
    centerY = Math.max(minCenterY, Math.min(maxCenterY, centerY));

    const left = (screenW / 2) - (centerX * contentW);
    const top = (screenH / 2) - (centerY * contentH);

    // Pixel layout mirrors the desktop scroll viewport: fit image first, then
    // enlarge and position it by normalized center. This makes the phone monitor
    // crop match the desktop crop instead of using a separate percent transform.
    img.style.setProperty('--remote-full-transform', 'none');
    img.style.setProperty('transform', 'none', 'important');
    img.style.position = 'absolute';
    img.style.left = `${left}px`;
    img.style.top = `${top}px`;
    img.style.width = `${contentW}px`;
    img.style.height = `${contentH}px`;
    img.style.maxWidth = 'none';
    img.style.maxHeight = 'none';
    img.dataset.remoteLeft = String(left);
    img.dataset.remoteTop = String(top);
    img.dataset.remoteContentW = String(contentW);
    img.dataset.remoteContentH = String(contentH);
    img.dataset.actualCenterX = String(centerX);
    img.dataset.actualCenterY = String(centerY);
    renderRemoteInkPreview(window.__phRemoteInkStrokes || [], window.__phRemoteCurrentPage || 1, { zoom, centerX, centerY });
  }

  function clampViewportToRemoteScreen(viewport) {
    const img = $('remoteFullImg');
    const screen = img ? (img.closest('.remote-monitor-screen') || $('remoteFullStage')) : null;
    if (!img || !screen) {
      return {
        zoom: Math.min(4, Math.max(1, Number(viewport.zoom) || 1)),
        centerX: Math.max(0, Math.min(1, Number(viewport.centerX) || 0.5)),
        centerY: Math.max(0, Math.min(1, Number(viewport.centerY) || 0.5)),
      };
    }
    const zoom = Math.min(4, Math.max(1, Number(viewport.zoom) || 1));
    const screenW = Math.max(1, screen.clientWidth || screen.getBoundingClientRect().width || 1);
    const screenH = Math.max(1, screen.clientHeight || screen.getBoundingClientRect().height || 1);
    const naturalW = Math.max(1, img.naturalWidth || 16);
    const naturalH = Math.max(1, img.naturalHeight || 9);
    const aspect = naturalW / naturalH;
    let fitW = screenW;
    let fitH = fitW / aspect;
    if (fitH > screenH) {
      fitH = screenH;
      fitW = fitH * aspect;
    }
    const contentW = fitW * zoom;
    const contentH = fitH * zoom;
    const minX = contentW > screenW ? screenW / (2 * contentW) : 0.5;
    const maxX = contentW > screenW ? 1 - minX : 0.5;
    const minY = contentH > screenH ? screenH / (2 * contentH) : 0.5;
    const maxY = contentH > screenH ? 1 - minY : 0.5;
    return {
      zoom,
      centerX: Math.max(minX, Math.min(maxX, Number(viewport.centerX) || 0.5)),
      centerY: Math.max(minY, Math.min(maxY, Number(viewport.centerY) || 0.5)),
    };
  }


  function attachRemotePreviewControls(ref, isHost, getViewport, setLocalViewport) {
    const openBtn = $('remotePreviewFullBtn');
    const closeBtn = $('remoteClosePreview');
    const backBtn = $('remoteBackToControls');
    const full = $('remoteFullPreview');
    const stage = $('remoteFullStage');
    if (!openBtn || !full || !stage) return;

    const safeViewport = (value) => clampViewportToRemoteScreen(value || {});

    let localViewport = safeViewport(getViewport());
    let rafId = 0;
    let gestureSettleTimer = null;

    const sendViewport = throttle((viewport) => {
      if (!isHost) return;
      sendRemoteCommand(ref, 'setViewport', safeViewport(viewport));
    }, 110);

    function renderLocal(viewport) {
      localViewport = safeViewport(viewport);
      setLocalViewport(localViewport);
      if (rafId) return;
      rafId = requestAnimationFrame(() => {
        rafId = 0;
        applyRemoteFullPreviewTransform(localViewport);
      });
    }

    function beginGesture() {
      window.__presentationHubRemoteGestureActive = true;
      clearTimeout(gestureSettleTimer);
    }

    function endGesture() {
      clearTimeout(gestureSettleTimer);
      gestureSettleTimer = setTimeout(() => {
        window.__presentationHubRemoteGestureActive = false;
        sendViewport(localViewport);
      }, 220);
    }

    function openFullPreview() {
      localViewport = safeViewport(getViewport());
      renderLocal(localViewport);
      full.classList.remove('hidden');
      document.body.classList.add('remote-fullscreen-open');
      if (full.requestFullscreen) full.requestFullscreen().catch(() => {});
    }

    function closeFullPreview() {
      window.__presentationHubRemoteGestureActive = false;
      full.classList.add('hidden');
      document.body.classList.remove('remote-fullscreen-open');
      document.documentElement.classList.add('presentation-hub-remote-mode');
      document.body.classList.add('presentation-hub-remote-mode');
      if (document.fullscreenElement === full) document.exitFullscreen().catch(() => {});
      requestAnimationFrame(() => {
        document.body.classList.remove('remote-fullscreen-open');
      });
    }

    openBtn.addEventListener('click', openFullPreview);
    if (closeBtn) closeBtn.addEventListener('click', closeFullPreview);
    if (backBtn) backBtn.addEventListener('click', closeFullPreview);
    document.addEventListener('fullscreenchange', () => {
      if (!document.fullscreenElement && !full.classList.contains('hidden')) closeFullPreview();
    });

    if (!isHost) return;

    let startDistance = 0;
    let startZoom = 1;
    let startCenter = { centerX: 0.5, centerY: 0.5 };
    let startMid = { x: 0.5, y: 0.5 };
    let startPoint = null;

    const distance = (touches) => {
      const dx = touches[0].clientX - touches[1].clientX;
      const dy = touches[0].clientY - touches[1].clientY;
      return Math.hypot(dx, dy);
    };

    const midpoint = (touches, rect) => ({
      x: ((touches[0].clientX + touches[1].clientX) / 2 - rect.left) / Math.max(1, rect.width),
      y: ((touches[0].clientY + touches[1].clientY) / 2 - rect.top) / Math.max(1, rect.height),
    });

    const gestureRect = () => {
      const screen = stage.querySelector('.remote-monitor-screen');
      return (screen || stage).getBoundingClientRect();
    };

    const screenGeometry = (zoomValue) => {
      const img = $('remoteFullImg');
      const screen = stage.querySelector('.remote-monitor-screen') || stage;
      const rect = screen.getBoundingClientRect();
      const naturalW = Math.max(1, img && img.naturalWidth ? img.naturalWidth : 16);
      const naturalH = Math.max(1, img && img.naturalHeight ? img.naturalHeight : 9);
      const aspect = naturalW / naturalH;
      let fitW = Math.max(1, rect.width);
      let fitH = fitW / aspect;
      if (fitH > rect.height) {
        fitH = Math.max(1, rect.height);
        fitW = fitH * aspect;
      }
      return { rect, fitW, fitH, contentW: fitW * zoomValue, contentH: fitH * zoomValue };
    };

    let activeTool = 'move';
    let currentStroke = null;
    const toolButtons = {
      move: $('remoteToolMove'),
      pen: $('remoteToolPen'),
      highlighter: $('remoteToolHighlighter'),
      erase: $('remoteToolErase'),
    };

    function setActiveTool(tool) {
      activeTool = tool;
      Object.entries(toolButtons).forEach(([name, button]) => {
        if (button) button.classList.toggle('active', name === tool);
      });
      const screen = stage.querySelector('.remote-monitor-screen');
      if (screen) screen.dataset.tool = tool;
    }

    Object.entries(toolButtons).forEach(([tool, button]) => {
      if (button) button.addEventListener('click', () => setActiveTool(tool));
    });
    const undoBtn = $('remoteUndoInk');
    const clearBtn = $('remoteClearInk');
    if (undoBtn) undoBtn.addEventListener('click', () => sendRemoteCommand(ref, 'undoInk'));
    if (clearBtn) clearBtn.addEventListener('click', () => {
      if (confirm('Clear drawings on this slide?')) sendRemoteCommand(ref, 'clearInk');
    });

    const sendErase = throttle((point) => sendRemoteCommand(ref, 'eraseInkAt', {
      page: window.__phRemoteCurrentPage || 1,
      x: point.x,
      y: point.y,
      radius: 0.035,
    }), 120);

    stage.addEventListener('touchstart', (event) => {
      if (event.touches.length) event.preventDefault();
      if (activeTool !== 'move' && event.touches.length === 1) {
        const point = remoteScreenPointToSlide(event.touches[0].clientX, event.touches[0].clientY);
        if (activeTool === 'erase') {
          sendErase(point);
          return;
        }
        currentStroke = {
          id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
          page: window.__phRemoteCurrentPage || 1,
          tool: activeTool === 'highlighter' ? 'highlighter' : 'pen',
          color: activeTool === 'highlighter' ? '#facc15' : '#ef4444',
          size: activeTool === 'highlighter' ? 18 : 6,
          points: [point],
          createdAt: Date.now(),
        };
        window.__phRemoteTempStroke = currentStroke;
        renderRemoteInkPreview(window.__phRemoteInkStrokes || [], window.__phRemoteCurrentPage || 1, localViewport);
        return;
      }
      beginGesture();
      const current = safeViewport(localViewport || getViewport());
      startZoom = current.zoom;
      startCenter = { centerX: current.centerX, centerY: current.centerY };
      const rect = gestureRect();
      if (event.touches.length === 2) {
        startDistance = Math.max(1, distance(event.touches));
        startMid = midpoint(event.touches, rect);
        startPoint = null;
      } else if (event.touches.length === 1) {
        startPoint = { x: event.touches[0].clientX, y: event.touches[0].clientY };
        startDistance = 0;
      }
    }, { passive: false });

    stage.addEventListener('touchmove', (event) => {
      if (event.touches.length !== 1 && event.touches.length !== 2) return;
      event.preventDefault();
      if (activeTool !== 'move' && event.touches.length === 1) {
        const point = remoteScreenPointToSlide(event.touches[0].clientX, event.touches[0].clientY);
        if (activeTool === 'erase') {
          sendErase(point);
          return;
        }
        if (currentStroke) {
          const last = currentStroke.points[currentStroke.points.length - 1];
          if (!last || Math.hypot(last.x - point.x, last.y - point.y) > 0.0025) currentStroke.points.push(point);
          window.__phRemoteTempStroke = currentStroke;
          renderRemoteInkPreview(window.__phRemoteInkStrokes || [], window.__phRemoteCurrentPage || 1, localViewport);
        }
        return;
      }
      beginGesture();
      const rect = gestureRect();
      let next = safeViewport(localViewport || getViewport());

      if (event.touches.length === 2 && startDistance) {
        const ratio = distance(event.touches) / startDistance;
        const mid = midpoint(event.touches, rect);
        const nextZoom = Math.min(4, Math.max(1, startZoom * ratio));
        const geometry = screenGeometry(nextZoom);
        const dragX = ((mid.x - startMid.x) * geometry.rect.width) / Math.max(1, geometry.contentW);
        const dragY = ((mid.y - startMid.y) * geometry.rect.height) / Math.max(1, geometry.contentH);
        next = safeViewport({
          zoom: nextZoom,
          centerX: startCenter.centerX - dragX,
          centerY: startCenter.centerY - dragY,
        });
      } else if (event.touches.length === 1 && startPoint) {
        const dx = event.touches[0].clientX - startPoint.x;
        const dy = event.touches[0].clientY - startPoint.y;
        const geometry = screenGeometry(startZoom);
        next = safeViewport({
          zoom: startZoom,
          centerX: startCenter.centerX - dx / Math.max(1, geometry.contentW),
          centerY: startCenter.centerY - dy / Math.max(1, geometry.contentH),
        });
      }

      renderLocal(next);
      sendViewport(next);
    }, { passive: false });

    stage.addEventListener('touchend', () => {
      if (currentStroke) {
        const finished = currentStroke;
        currentStroke = null;
        window.__phRemoteTempStroke = null;
        if (finished.points && finished.points.length > 1) sendRemoteCommand(ref, 'addInkStroke', finished);
        renderRemoteInkPreview(window.__phRemoteInkStrokes || [], window.__phRemoteCurrentPage || 1, localViewport);
        return;
      }
      endGesture();
    }, { passive: true });
    stage.addEventListener('touchcancel', () => {
      currentStroke = null;
      window.__phRemoteTempStroke = null;
      renderRemoteInkPreview(window.__phRemoteInkStrokes || [], window.__phRemoteCurrentPage || 1, localViewport);
      endGesture();
    }, { passive: true });

    stage.addEventListener('wheel', (event) => {
      event.preventDefault();
      beginGesture();
      const current = safeViewport(localViewport || getViewport());
      const rect = gestureRect();
      const next = safeViewport({
        zoom: current.zoom + (event.deltaY < 0 ? 0.12 : -0.12),
        centerX: (event.clientX - rect.left) / Math.max(1, rect.width),
        centerY: (event.clientY - rect.top) / Math.max(1, rect.height),
      });
      renderLocal(next);
      sendViewport(next);
      endGesture();
    }, { passive: false });
  }

  function throttle(fn, delay) {
    let last = 0;
    let trailing = null;
    return (...args) => {
      const now = Date.now();
      if (now - last >= delay) {
        last = now;
        fn(...args);
        return;
      }
      clearTimeout(trailing);
      trailing = setTimeout(() => {
        last = Date.now();
        fn(...args);
      }, delay - (now - last));
    };
  }

  function sendRemoteCommand(ref, action, value) {
    return ref.set({
      command: {
        id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
        action,
        value: value ?? null,
        issuedAt: Date.now(),
        v: 10,
      }
    }, { merge: true });
  }

  function registerServiceWorker() {
    if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }
})();
