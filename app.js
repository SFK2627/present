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
    lastCountdownAlertSecond: null,
    audioContext: null,
    transitionEffect: 'fade',
    slideChangePending: false,
    pdfPageCache: new Map(),
    timerTick: null,
    timerStartedAt: null,
    timerElapsedBeforePause: 0,
    timer: {
      visible: false,
      mode: 'up',
      countdownSeconds: 600,
      position: 'bottom-right',
      opacity: 75,
    },
    toolbarHideTimer: null,
    renderToken: 0,
    lastRemoteThumbKey: '',
    remotePreviewBusy: false,
    firebaseReady: false,
    firebaseDb: null,
    firebaseAuthReady: false,
    sessionId: null,
    sessionRef: null,
    unsubscribeSession: null,
    lastCommandId: null,
    publishLock: false,
  };

  const els = {};

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
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
      'thumbnailSidebar', 'viewerStage', 'viewerToolbar', 'controlPanel', 'settingsBtn', 'backHomeBtn',
      'prevBtn', 'nextBtn', 'jumpInput', 'pageTotalLabel', 'zoomOutBtn', 'zoomInBtn', 'resetZoomBtn',
      'zoomLabel', 'fullscreenBtn', 'qrBtn', 'viewerCanvasWrap', 'pdfCanvas', 'pptxSlide',
      'timingModeSelect', 'globalTimingSelect', 'customTimingWrap', 'customTimingInput', 'perSlideTimingWrap', 'perSlideTimingInput',
      'countdownAlertSelect', 'slideTransitionSelect',
      'autoStartBtn', 'autoPauseBtn', 'autoResumeBtn', 'autoStopBtn', 'timerOverlay', 'timerModeSelect',
      'countdownMinutesInput', 'timerPositionSelect', 'timerOpacityInput', 'timerShowBtn', 'timerHideBtn',
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
      publishSessionState();
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
    if (els.timerShowBtn) els.timerShowBtn.addEventListener('click', showTimer);
    if (els.timerHideBtn) els.timerHideBtn.addEventListener('click', hideTimer);
    if (els.timerResetBtn) els.timerResetBtn.addEventListener('click', resetTimer);

    document.addEventListener('keydown', handleKeyboard);
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

  async function handleFiles(fileList) {
    const files = Array.from(fileList || []).filter((file) => /\.(pdf|pptx)$/i.test(file.name));
    if (!files.length) return;

    for (const file of files) {
      const record = await createPresentationRecord(file);
      state.files.unshift(record);
      await dbPut(record);
      renderLibrary();
    }
    if (els.fileInput) els.fileInput.value = '';
  }

  async function createPresentationRecord(file) {
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

  async function renderPdfPageToDataUrl(pdf, pageNumber, targetWidth = 900) {
    const page = await pdf.getPage(pageNumber);
    const base = page.getViewport({ scale: 1 });
    const scale = Math.min(1.25, Math.max(0.25, targetWidth / base.width));
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d', { alpha: false });
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    return canvas.toDataURL('image/jpeg', 0.86);
  }

  function renderLibrary() {
    if (!els.cardsGrid) return;
    const query = (els.searchInput.value || '').toLowerCase().trim();
    const sort = els.sortSelect.value;
    let files = [...state.files];

    if (query) files = files.filter((file) => file.name.toLowerCase().includes(query));
    files.sort((a, b) => {
      if (sort === 'name') return a.name.localeCompare(b.name);
      if (sort === 'date') return new Date(b.uploadedAt) - new Date(a.uploadedAt);
      return new Date(b.lastViewed || b.uploadedAt) - new Date(a.lastViewed || a.uploadedAt);
    });

    els.cardsGrid.innerHTML = files.map(cardTemplate).join('');
    els.emptyState.classList.toggle('hidden', files.length > 0);
    els.libraryCount.textContent = `${state.files.length} file${state.files.length === 1 ? '' : 's'} saved locally`;

    els.cardsGrid.querySelectorAll('[data-open]').forEach((button) => {
      button.addEventListener('click', () => openPresentation(button.dataset.open));
    });
    els.cardsGrid.querySelectorAll('[data-delete]').forEach((button) => {
      button.addEventListener('click', () => deletePresentation(button.dataset.delete));
    });
  }

  function cardTemplate(file) {
    const date = new Date(file.uploadedAt).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
    const typeLabel = file.type === 'pdf' ? 'PDF' : 'PPTX';
    return `
      <article class="presentation-card glass">
        <span class="file-badge">${typeLabel}</span>
        <div class="card-thumb"><img src="${file.thumbnail || createGenericThumbDataUrl(file.name, file.type)}" alt="${escapeHtml(file.name)} thumbnail"></div>
        <div class="card-body">
          <h4 title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</h4>
          <div class="card-meta">
            <span>${file.pageCount} ${file.type === 'pdf' ? 'pages' : 'slides'}</span>
            <span>${date}</span>
          </div>
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
    renderLibrary();
  }

  async function clearLibrary() {
    if (!confirm('Clear all locally saved presentations?')) return;
    state.files = [];
    await dbClear();
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
    state.activePptxSlides = [];
    state.pptxVisualReady = false;
    state.pptxRenderedSlides = [];
    state.pdfPageCache = new Map();
    state.lastCountdownAlertSecond = null;
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
      } else {
        await preparePptxVisualRenderer(file.blob);
        if (!state.pptxVisualReady) {
          state.activePptxSlides = await parsePptxSlides(file.blob);
          state.totalPages = state.activePptxSlides.length;
        }
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
    state.activePptxSlides = [];
    state.pptxVisualReady = false;
    state.pptxRenderedSlides = [];
    state.pdfPageCache = new Map();
    state.lastCountdownAlertSecond = null;
    if (state.pptxBlobUrl) {
      URL.revokeObjectURL(state.pptxBlobUrl);
      state.pptxBlobUrl = '';
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
    const rect = fullscreen
      ? { width: window.innerWidth, height: window.innerHeight }
      : (els.viewerStage ? els.viewerStage.getBoundingClientRect() : { width: window.innerWidth, height: window.innerHeight });
    const horizontalPadding = fullscreen ? 0 : 64;
    const verticalPadding = fullscreen ? 0 : 46;
    return {
      width: Math.max(320, rect.width - horizontalPadding),
      height: Math.max(240, rect.height - verticalPadding),
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
    els.jumpInput.value = state.currentPage;
    els.perSlideTimingInput.value = state.perPageTiming[state.currentPage] || '';
    updateTimingModeUI();
    updateActiveThumb();

    if (state.activeFile.type === 'pdf') {
      els.pptxSlide.classList.add('hidden');
      els.pdfCanvas.classList.remove('hidden');
      const page = await getCachedPdfPage(state.currentPage);
      if (token !== state.renderToken) return;
      const viewportBase = page.getViewport({ scale: 1 });
      const available = getAvailableStageSize();
      const fitScale = Math.min(available.width / viewportBase.width, available.height / viewportBase.height);
      const cssScale = Math.max(0.15, fitScale * state.zoom);
      const dpr = Math.min(window.devicePixelRatio || 1, 2.4);
      const renderViewport = page.getViewport({ scale: cssScale * dpr });
      const cssViewport = page.getViewport({ scale: cssScale });
      const canvas = els.pdfCanvas;
      const ctx = canvas.getContext('2d', { alpha: false });
      canvas.width = Math.floor(renderViewport.width);
      canvas.height = Math.floor(renderViewport.height);
      canvas.style.width = `${Math.floor(cssViewport.width)}px`;
      canvas.style.height = `${Math.floor(cssViewport.height)}px`;
      await page.render({ canvasContext: ctx, viewport: renderViewport }).promise;
      warmAdjacentPdfPages();
      applyViewportScroll();
    } else {
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
    if (state.slideChangePending) {
      playSlideTransition();
      state.slideChangePending = false;
    }
    updateZoomLabel();
    publishSessionState();
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
        <div class="thumb-label"><span>${state.activeFile.type === 'pdf' ? 'Page' : 'Slide'} ${i}</span><span>${state.perPageTiming[i] ? state.perPageTiming[i] + 's' : ''}</span></div>
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

    if (state.activeFile.type === 'pdf') {
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
  }

  function setZoom(value, options = {}) {
    state.zoom = Math.min(4, Math.max(0.25, Number(value)));
    if (Number.isFinite(Number(options.centerX))) state.viewportCenterX = clamp01(Number(options.centerX));
    if (Number.isFinite(Number(options.centerY))) state.viewportCenterY = clamp01(Number(options.centerY));
    renderCurrentPage();
  }

  function setViewportTransform(value = {}) {
    setZoom(value.zoom ?? state.zoom, { centerX: value.centerX ?? state.viewportCenterX, centerY: value.centerY ?? state.viewportCenterY });
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
    });
  }

  function updateZoomLabel() {
    els.zoomLabel.textContent = `${Math.round(state.zoom * 100)}%`;
  }

  function handleKeyboard(event) {
    if (els.viewerView.classList.contains('hidden')) return;
    const focusedTag = document.activeElement ? document.activeElement.tagName : '';
    const isTyping = ['INPUT', 'TEXTAREA', 'SELECT'].includes(focusedTag);

    if (event.key === 'Escape' && document.fullscreenElement) {
      event.preventDefault();
      document.exitFullscreen().catch(() => {});
      return;
    }
    if (isTyping) return;
    if (event.key === 'ArrowRight') {
      event.preventDefault();
      nextPage();
    }
    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      previousPage();
    }
    if (event.key.toLowerCase() === 'f') {
      event.preventDefault();
      toggleFullscreen();
    }
  }

  async function toggleFullscreen() {
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
    els.timerModeSelect.value = state.timer.mode;
    els.timerPositionSelect.value = state.timer.position;
    els.timerOpacityInput.value = state.timer.opacity;
    if (els.countdownAlertSelect) els.countdownAlertSelect.value = state.countdownAlert || 'off';
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

  function playSlideTransition() {
    const effect = state.transitionEffect || 'fade';
    if (effect === 'none') return;
    const target = getTransitionTarget();
    if (!target) return;
    const classes = ['slide-transition', 'transition-fade', 'transition-slide-left', 'transition-slide-right', 'transition-slide-up', 'transition-zoom-in', 'transition-zoom-out', 'transition-soft-blur'];
    target.classList.remove(...classes);
    // Force restart animation.
    void target.offsetWidth;
    target.classList.add('slide-transition', `transition-${effect}`);
    window.setTimeout(() => target.classList.remove(...classes), 620);
  }

  function getTransitionTarget() {
    // Animate the viewport wrapper, not the canvas/slide itself, so PPTX/PDF scale transforms stay intact.
    return els.viewerCanvasWrap || els.pdfCanvas || els.pptxSlide;
  }

  function checkCountdownAlert() {
    if (!state.autoPlaying || state.autoPaused || state.countdownAlert === 'off') return;
    const remaining = Math.ceil(Math.max(0, state.autoCurrentDuration - getAutoElapsedSeconds(true)));
    if (remaining < 1 || remaining > 5) return;
    if (remaining === state.lastCountdownAlertSecond) return;
    state.lastCountdownAlertSecond = remaining;
    triggerCountdownAlert(remaining);
  }

  function triggerCountdownAlert(second) {
    const mode = state.countdownAlert;
    if (mode === 'sound' || mode === 'both') playCountdownBeep(second);
    if (mode === 'voice' || mode === 'both') speakCountdown(second);
  }

  function playCountdownBeep(second) {
    try {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (!AudioContext) return;
      state.audioContext = state.audioContext || new AudioContext();
      if (state.audioContext.state === 'suspended') state.audioContext.resume().catch(() => {});
      const ctx = state.audioContext;
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.value = second === 1 ? 1040 : 760;
      gain.gain.setValueAtTime(0.0001, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.16, ctx.currentTime + 0.015);
      gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + 0.18);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(ctx.currentTime);
      osc.stop(ctx.currentTime + 0.2);
    } catch (error) {}
  }

  function speakCountdown(second) {
    try {
      if (!window.speechSynthesis) return;
      window.speechSynthesis.cancel();
      const utterance = new SpeechSynthesisUtterance(String(second));
      utterance.rate = 1.05;
      utterance.pitch = 1;
      utterance.volume = 0.9;
      window.speechSynthesis.speak(utterance);
    } catch (error) {}
  }

  async function setupRemoteSessionIfPossible() {
    if (!state.firebaseReady || !state.activeFile) return;
    if (state.unsubscribeSession) state.unsubscribeSession();
    state.sessionId = state.sessionId || makeSessionId();
    state.sessionRef = state.firebaseDb.collection(SESSION_COLLECTION).doc(state.sessionId);
    state.lastCommandId = null;
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
    if (!state.sessionRef || !state.activeFile || state.publishLock) return;
    state.publishLock = true;
    setTimeout(() => { state.publishLock = false; }, force ? 0 : 200);

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
      timerMode: state.timer.mode,
      timerPosition: state.timer.position,
      countdownAlert: state.countdownAlert,
      transitionEffect: state.transitionEffect,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    };

    try {
      const thumb = await getCurrentThumbForRemote();
      if (thumb) payload.thumb = thumb;
      await state.sessionRef.set(payload, { merge: true });
    } catch (error) {
      console.warn('Could not publish session:', error);
    }
  }

  async function getCurrentThumbForRemote() {
    const key = `${state.activeFile ? state.activeFile.id : ''}:${state.currentPage}:${Math.round(state.zoom * 100)}:${Math.round(state.viewportCenterX * 100)}:${Math.round(state.viewportCenterY * 100)}:${state.timer.visible ? 1 : 0}`;
    if (!state.remotePreviewBusy && state.lastRemoteThumbKey === key && !isPresentationFullscreen()) return '';
    state.remotePreviewBusy = true;
    try {
      state.lastRemoteThumbKey = key;
      if (state.activeFile.type === 'pdf' && state.activePdf) {
        return await renderPdfPageToDataUrl(state.activePdf, state.currentPage, 900);
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

  async function applyRemoteCommand(command) {
    switch (command.action) {
      case 'next': nextPage(); break;
      case 'prev': previousPage(); break;
      case 'first': jumpToPage(1); break;
      case 'last': jumpToPage(state.totalPages); break;
      case 'zoomIn': setZoom(state.zoom + 0.1); break;
      case 'zoomOut': setZoom(state.zoom - 0.1); break;
      case 'resetZoom': setZoom(1, { centerX: 0.5, centerY: 0.5 }); break;
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
      case 'setTimerOpacity':
        state.timer.opacity = Number(command.value);
        applyTimerSettings();
        publishSessionState();
        break;
      case 'setCountdownAlert':
        state.countdownAlert = ['off', 'sound', 'voice', 'both'].includes(command.value) ? command.value : 'off';
        state.lastCountdownAlertSecond = null;
        applyTimerSettings();
        publishSessionState();
        break;
      case 'setTransitionEffect':
        state.transitionEffect = command.value || 'fade';
        applyTimerSettings();
        publishSessionState();
        break;
      default: break;
    }
  }

  async function openQrModal() {
    if (!state.activeFile) return;
    if (!state.firebaseReady) {
      els.qrHelp.textContent = 'Phone remote across devices needs Firebase config. You can still present locally. Click Remote setup on the home screen to see where to add your Firebase keys.';
    } else {
      await setupRemoteSessionIfPossible();
      els.qrHelp.textContent = `Session ${state.sessionId} is live. Multiple phones can connect. Host can control; Viewer is view-only.`;
    }

    const session = state.sessionId || 'NO-FIREBASE';
    const baseUrl = window.location.href.split('?')[0].split('#')[0];
    const hostUrl = `${baseUrl}?remote=1&session=${encodeURIComponent(session)}&role=host`;
    const viewerUrl = `${baseUrl}?remote=1&session=${encodeURIComponent(session)}&role=viewer`;

    els.hostRemoteLink.value = hostUrl;
    els.viewerRemoteLink.value = viewerUrl;
    els.hostQr.innerHTML = '';
    els.viewerQr.innerHTML = '';
    new QRCode(els.hostQr, { text: hostUrl, width: 210, height: 210 });
    new QRCode(els.viewerQr, { text: viewerUrl, width: 210, height: 210 });
    showModal(els.qrModal);
  }

  function renderRemoteApp() {
    els.app.classList.add('hidden');
    els.remoteApp.classList.remove('hidden');
    const sessionId = qs.get('session') || '';
    const role = qs.get('role') || 'viewer';
    const isHost = role === 'host';
    let remoteViewport = { zoom: 1, centerX: 0.5, centerY: 0.5 };
    let latestThumb = '';

    els.remoteApp.innerHTML = `
      <main class="remote-card ${isHost ? '' : 'remote-viewer-only'}">
        <div class="remote-head">
          <div>
            <h1>Presentation Remote</h1>
            <p class="remote-sub">${escapeHtml(sessionId || 'No session')} • ${isHost ? 'Host control' : 'Viewer only'}</p>
          </div>
          <span id="remoteStatusPill" class="remote-status-pill">Connecting</span>
        </div>
        <div class="remote-preview-shell">
          <div class="remote-preview" id="remotePreview"><span>Waiting for presentation...</span></div>
          <button id="remotePreviewFullBtn" class="remote-preview-full-btn" data-host-only="false">Fullscreen Preview</button>
        </div>
        <p class="remote-hint">Open fullscreen, rotate your phone landscape, then pinch and drag the preview. The desktop follows the same zoomed area.</p>
        <h2 id="remoteSlideLabel">Slide -- / --</h2>
        <p id="remoteFileLabel" class="remote-sub">Connect to an active desktop session.</p>
        <p id="remoteAutoLabel" class="remote-sub remote-auto-label">Auto Play idle</p>
        <section class="remote-grid">
          <button data-command="first" data-host-only="true">First</button>
          <button data-command="last" data-host-only="true">Last</button>
          <button data-command="prev" data-host-only="true">Previous</button>
          <button data-command="next" data-host-only="true">Next</button>
        </section>
        <section class="remote-section">
          <h3>Zoom</h3>
          <div class="remote-control-row">
            <button data-command="zoomOut" data-host-only="true">Zoom −</button>
            <button data-command="resetZoom" data-host-only="true">Reset</button>
            <button data-command="zoomIn" data-host-only="true">Zoom +</button>
          </div>
        </section>
        <section class="remote-section">
          <h3>Auto Play</h3>
          <div class="remote-control-row">
            <button data-command="autoStart" data-host-only="true">Start</button>
            <button data-command="autoPause" data-host-only="true">Pause</button>
            <button data-command="autoStop" data-host-only="true">Stop</button>
          </div>
          <div class="remote-wide remote-timing-box">
            <select id="remoteTimingMode" data-host-only="true">
              <option value="global">Global - all slides</option>
              <option value="per-slide">Per-slide - current slide</option>
            </select>
            <input id="remoteTiming" type="number" min="1" value="10" placeholder="Seconds">
            <button id="remoteSetTiming" data-host-only="true">Apply Timing</button>
          </div>
        </section>
        <section class="remote-section">
          <h3>Timer</h3>
          <div class="remote-control-row">
            <button data-command="timerShow" data-host-only="true">Show</button>
            <button data-command="timerHide" data-host-only="true">Hide</button>
            <button data-command="timerReset" data-host-only="true">Reset</button>
          </div>
          <div class="remote-wide">
            <input id="remoteOpacity" type="range" min="0" max="100" step="25" value="75">
            <button id="remoteSetOpacity" data-host-only="true">Change Opacity</button>
          </div>
        </section>
      </main>
      <div id="remoteFullPreview" class="remote-full-preview hidden">
        <div class="remote-full-toolbar">
          <span id="remoteFullLabel">Slide Preview</span>
          <button id="remoteClosePreview">Close</button>
        </div>
        <div id="remoteFullStage" class="remote-full-stage">
          <img id="remoteFullImg" alt="Fullscreen current slide preview" />
        </div>
        <div class="remote-full-help">Pinch to zoom. Drag to choose the exact area shown on the desktop.</div>
      </div>
    `;

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
        $('remoteStatusPill').textContent = isHost ? 'Host' : 'Viewer';
        $('remoteSlideLabel').textContent = `${data.type === 'pdf' ? 'Page' : 'Slide'} ${data.currentPage || '--'} / ${data.totalPages || '--'}`;
        $('remoteFullLabel').textContent = `${data.type === 'pdf' ? 'Page' : 'Slide'} ${data.currentPage || '--'} / ${data.totalPages || '--'}`;
        $('remoteFileLabel').textContent = data.fileName || 'Active presentation';
        if (data.thumb) latestThumb = data.thumb;
        const preview = $('remotePreview');
        preview.innerHTML = latestThumb ? `<img src="${latestThumb}" alt="Current slide preview">` : '<span>No preview yet</span>';
        const fullImg = $('remoteFullImg');
        if (fullImg && latestThumb && fullImg.src !== latestThumb) fullImg.src = latestThumb;
        remoteViewport = {
          zoom: Number(data.zoom) || 1,
          centerX: Number.isFinite(Number(data.viewportCenterX)) ? Number(data.viewportCenterX) : 0.5,
          centerY: Number.isFinite(Number(data.viewportCenterY)) ? Number(data.viewportCenterY) : 0.5,
        };
        $('remoteTimingMode').value = data.timingMode || 'global';
        $('remoteTiming').value = (data.timingMode === 'per-slide' ? data.currentTiming : data.globalTiming) || 10;
        $('remoteOpacity').value = data.timerOpacity ?? 75;
        $('remoteAutoLabel').textContent = data.autoPlaying
          ? `Auto Play: ${data.autoElapsed || 0}s / ${data.autoDuration || data.currentTiming || data.globalTiming || 10}s`
          : (data.autoPaused ? `Auto Play paused at ${data.autoElapsed || 0}s` : 'Auto Play idle');
        applyRemoteFullPreviewTransform(remoteViewport);
      });

      els.remoteApp.querySelectorAll('[data-command]').forEach((button) => {
        button.addEventListener('click', () => {
          if (!isHost) return;
          sendRemoteCommand(ref, button.dataset.command);
        });
      });
      attachRemotePreviewControls(ref, isHost, () => remoteViewport, (next) => {
        remoteViewport = next;
        applyRemoteFullPreviewTransform(remoteViewport);
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
      $('remoteSetOpacity').addEventListener('click', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setTimerOpacity', Number($('remoteOpacity').value));
      });
    });
  }

  function applyRemoteFullPreviewTransform(viewport) {
    const img = $('remoteFullImg');
    if (!img) return;
    const zoom = Math.min(4, Math.max(1, Number(viewport.zoom) || 1));
    const centerX = Math.max(0, Math.min(1, Number(viewport.centerX) || 0.5));
    const centerY = Math.max(0, Math.min(1, Number(viewport.centerY) || 0.5));
    const offsetStrength = Math.max(0, zoom - 1);
    const translateX = (0.5 - centerX) * 100 * offsetStrength;
    const translateY = (0.5 - centerY) * 100 * offsetStrength;
    img.style.transformOrigin = 'center center';
    img.style.transform = `translate3d(${translateX}%, ${translateY}%, 0) scale(${zoom})`;
  }

  function attachRemotePreviewControls(ref, isHost, getViewport, setLocalViewport) {
    const openBtn = $('remotePreviewFullBtn');
    const closeBtn = $('remoteClosePreview');
    const full = $('remoteFullPreview');
    const stage = $('remoteFullStage');
    if (!openBtn || !full || !stage) return;

    const safeViewport = (value) => ({
      zoom: Math.min(4, Math.max(1, Number(value.zoom) || 1)),
      centerX: Math.max(0, Math.min(1, Number(value.centerX) || 0.5)),
      centerY: Math.max(0, Math.min(1, Number(value.centerY) || 0.5)),
    });

    const sendViewport = throttle((viewport) => {
      if (!isHost) return;
      sendRemoteCommand(ref, 'setViewport', safeViewport(viewport));
    }, 90);

    function openFullPreview() {
      full.classList.remove('hidden');
      document.body.classList.add('remote-fullscreen-open');
      applyRemoteFullPreviewTransform(getViewport());
      if (full.requestFullscreen) full.requestFullscreen().catch(() => {});
      if (screen.orientation && screen.orientation.lock) screen.orientation.lock('landscape').catch(() => {});
    }

    function closeFullPreview() {
      full.classList.add('hidden');
      document.body.classList.remove('remote-fullscreen-open');
      if (document.fullscreenElement === full) document.exitFullscreen().catch(() => {});
      if (screen.orientation && screen.orientation.unlock) {
        try { screen.orientation.unlock(); } catch (error) {}
      }
    }

    openBtn.addEventListener('click', openFullPreview);
    if (closeBtn) closeBtn.addEventListener('click', closeFullPreview);
    document.addEventListener('fullscreenchange', () => {
      if (!document.fullscreenElement && !full.classList.contains('hidden')) closeFullPreview();
    });

    if (!isHost) return;

    let startDistance = 0;
    let startZoom = 1;
    let startCenter = { centerX: 0.5, centerY: 0.5 };
    let startPoint = null;

    const distance = (touches) => {
      const dx = touches[0].clientX - touches[1].clientX;
      const dy = touches[0].clientY - touches[1].clientY;
      return Math.hypot(dx, dy);
    };

    const midpoint = (touches, rect) => ({
      x: ((touches[0].clientX + touches[1].clientX) / 2 - rect.left) / rect.width,
      y: ((touches[0].clientY + touches[1].clientY) / 2 - rect.top) / rect.height,
    });

    stage.addEventListener('touchstart', (event) => {
      if (event.touches.length) event.preventDefault();
      const current = safeViewport(getViewport());
      startZoom = current.zoom;
      startCenter = { centerX: current.centerX, centerY: current.centerY };
      if (event.touches.length === 2) {
        startDistance = distance(event.touches);
        startPoint = null;
      } else if (event.touches.length === 1) {
        startPoint = { x: event.touches[0].clientX, y: event.touches[0].clientY };
        startDistance = 0;
      }
    }, { passive: false });

    stage.addEventListener('touchmove', (event) => {
      if (event.touches.length !== 1 && event.touches.length !== 2) return;
      event.preventDefault();
      const rect = stage.getBoundingClientRect();
      let next = safeViewport(getViewport());

      if (event.touches.length === 2 && startDistance) {
        const ratio = distance(event.touches) / startDistance;
        const mid = midpoint(event.touches, rect);
        next = safeViewport({
          zoom: startZoom * ratio,
          centerX: mid.x,
          centerY: mid.y,
        });
      } else if (event.touches.length === 1 && startPoint) {
        const dx = event.touches[0].clientX - startPoint.x;
        const dy = event.touches[0].clientY - startPoint.y;
        const divisor = Math.max(1, startZoom);
        next = safeViewport({
          zoom: startZoom,
          centerX: startCenter.centerX - dx / (rect.width * divisor),
          centerY: startCenter.centerY - dy / (rect.height * divisor),
        });
      }

      setLocalViewport(next);
      sendViewport(next);
    }, { passive: false });

    stage.addEventListener('wheel', (event) => {
      event.preventDefault();
      const current = safeViewport(getViewport());
      const rect = stage.getBoundingClientRect();
      const next = safeViewport({
        zoom: current.zoom + (event.deltaY < 0 ? 0.12 : -0.12),
        centerX: (event.clientX - rect.left) / rect.width,
        centerY: (event.clientY - rect.top) / rect.height,
      });
      setLocalViewport(next);
      sendViewport(next);
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
        issuedAt: firebase.firestore.FieldValue.serverTimestamp(),
      }
    }, { merge: true });
  }

  function registerServiceWorker() {
    if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }
})();
