/* Presentation Hub static app
   - Open index.html directly or upload the folder to any static host.
   - PDF rendering uses PDF.js from CDN.
   - PPTX static mode extracts slide text and slide count from the PPTX zip. For pixel-perfect PPTX, convert PPTX to PDF first.
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
    perPageTiming: {},
    autoTimer: null,
    autoPaused: false,
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
    await initFirebaseIfConfigured();

    if (isRemoteMode) {
      renderRemoteApp();
      return;
    }

    await loadLibrary();
    renderLibrary();
    registerServiceWorker();
    window.addEventListener('resize', () => {
      if (state.activeFile && state.activeFile.type === 'pptx' && state.pptxVisualReady) showOnlyCurrentPptxSlide();
    });
  }

  function cacheElements() {
    [
      'app', 'remoteApp', 'homeView', 'viewerView', 'fileInput', 'uploadZone', 'searchInput', 'sortSelect',
      'cardsGrid', 'emptyState', 'libraryCount', 'clearLibraryBtn', 'themeToggle', 'firebaseStatus',
      'thumbnailSidebar', 'viewerStage', 'viewerToolbar', 'controlPanel', 'settingsBtn', 'backHomeBtn',
      'prevBtn', 'nextBtn', 'jumpInput', 'pageTotalLabel', 'zoomOutBtn', 'zoomInBtn', 'resetZoomBtn',
      'zoomLabel', 'fullscreenBtn', 'qrBtn', 'viewerCanvasWrap', 'pdfCanvas', 'pptxSlide',
      'globalTimingSelect', 'customTimingWrap', 'customTimingInput', 'perSlideTimingInput',
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
    if (els.firebaseStatus) els.firebaseStatus.addEventListener('click', () => els.setupModal.classList.remove('hidden'));
    if (els.closeSetupBtn) els.closeSetupBtn.addEventListener('click', () => els.setupModal.classList.add('hidden'));

    if (els.backHomeBtn) els.backHomeBtn.addEventListener('click', closeViewer);
    if (els.prevBtn) els.prevBtn.addEventListener('click', previousPage);
    if (els.nextBtn) els.nextBtn.addEventListener('click', nextPage);
    if (els.jumpInput) els.jumpInput.addEventListener('change', () => jumpToPage(Number(els.jumpInput.value)));
    if (els.zoomInBtn) els.zoomInBtn.addEventListener('click', () => setZoom(state.zoom + 0.1));
    if (els.zoomOutBtn) els.zoomOutBtn.addEventListener('click', () => setZoom(state.zoom - 0.1));
    if (els.resetZoomBtn) els.resetZoomBtn.addEventListener('click', () => setZoom(1));
    if (els.fullscreenBtn) els.fullscreenBtn.addEventListener('click', toggleFullscreen);
    if (els.qrBtn) els.qrBtn.addEventListener('click', openQrModal);
    if (els.closeQrBtn) els.closeQrBtn.addEventListener('click', () => els.qrModal.classList.add('hidden'));
    if (els.settingsBtn) els.settingsBtn.addEventListener('click', () => els.controlPanel.classList.toggle('hidden'));

    if (els.globalTimingSelect) els.globalTimingSelect.addEventListener('change', () => {
      els.customTimingWrap.classList.toggle('hidden', els.globalTimingSelect.value !== 'custom');
      publishSessionState();
    });
    if (els.customTimingInput) els.customTimingInput.addEventListener('input', publishSessionState);
    if (els.perSlideTimingInput) els.perSlideTimingInput.addEventListener('change', () => {
      const val = Number(els.perSlideTimingInput.value);
      if (val > 0) state.perPageTiming[state.currentPage] = val;
      else delete state.perPageTiming[state.currentPage];
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
    if (els.viewerStage) {
      els.viewerStage.addEventListener('mousemove', revealToolbarTemporarily);
      els.viewerStage.addEventListener('touchstart', revealToolbarTemporarily, { passive: true });
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
        note = 'Static PPTX preview extracts slide text. Convert to PDF for exact PowerPoint layout.';
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

  async function renderPdfPageToDataUrl(pdf, pageNumber, scale) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: ctx, viewport }).promise;
    return canvas.toDataURL('image/jpeg', 0.8);
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
    state.perPageTiming = {};
    state.activePdf = null;
    state.activePptxSlides = [];
    state.pptxVisualReady = false;
    state.pptxRenderedSlides = [];
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
    stopAutoPlay();
    stopTimerInterval();
    if (state.unsubscribeSession) state.unsubscribeSession();
    state.unsubscribeSession = null;
    state.sessionRef = null;
    state.sessionId = null;
    if (state.activePdf && state.activePdf.destroy) state.activePdf.destroy();
    state.activePdf = null;
    state.activePptxSlides = [];
    state.pptxVisualReady = false;
    state.pptxRenderedSlides = [];
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

    if (!window.jQuery || !window.jQuery.fn || !window.jQuery.fn.pptxToHtml) {
      console.warn('PPTXjs visual renderer is not available. Falling back to text preview.');
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

  function showOnlyCurrentPptxSlide() {
    if (!state.pptxVisualReady || !state.pptxRenderedSlides.length) return;
    const current = state.pptxRenderedSlides[state.currentPage - 1] || state.pptxRenderedSlides[0];
    const maxWidth = Math.max(320, window.innerWidth - 300);
    const maxHeight = Math.max(240, window.innerHeight - 110);

    state.pptxRenderedSlides.forEach((slide) => {
      slide.style.display = 'none';
      slide.style.visibility = 'visible';
      slide.style.transformOrigin = 'center center';
    });

    current.style.display = 'block';
    current.style.transform = 'none';
    current.style.margin = '0 auto';

    const rect = current.getBoundingClientRect();
    const naturalWidth = current.offsetWidth || rect.width || 960;
    const naturalHeight = current.offsetHeight || rect.height || 540;
    const fitScale = Math.min(maxWidth / naturalWidth, maxHeight / naturalHeight, 1.3);
    const finalScale = Math.max(0.15, fitScale * state.zoom);

    current.style.transform = `scale(${finalScale})`;
    els.pptxSlide.style.width = `${naturalWidth * finalScale}px`;
    els.pptxSlide.style.height = `${naturalHeight * finalScale}px`;
    els.pptxSlide.style.transform = '';
  }

  async function renderCurrentPage() {
    if (!state.activeFile) return;
    state.currentPage = Math.min(Math.max(1, state.currentPage), state.totalPages);
    els.jumpInput.value = state.currentPage;
    els.perSlideTimingInput.value = state.perPageTiming[state.currentPage] || '';
    updateActiveThumb();

    if (state.activeFile.type === 'pdf') {
      els.pptxSlide.classList.add('hidden');
      els.pdfCanvas.classList.remove('hidden');
      const page = await state.activePdf.getPage(state.currentPage);
      const baseWidth = Math.min(1180, Math.max(700, window.innerWidth - 320));
      const viewportBase = page.getViewport({ scale: 1 });
      const scale = (baseWidth / viewportBase.width) * state.zoom;
      const viewport = page.getViewport({ scale });
      const canvas = els.pdfCanvas;
      const ctx = canvas.getContext('2d');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      await page.render({ canvasContext: ctx, viewport }).promise;
    } else {
      els.pdfCanvas.classList.add('hidden');
      els.pptxSlide.classList.remove('hidden');
      if (state.pptxVisualReady) {
        showOnlyCurrentPptxSlide();
      } else {
        els.pptxSlide.classList.remove('visual-pptx');
        const slide = state.activePptxSlides[state.currentPage - 1] || { title: `Slide ${state.currentPage}`, lines: [] };
        els.pptxSlide.style.transform = `scale(${state.zoom})`;
        els.pptxSlide.innerHTML = `
          <h2>${escapeHtml(slide.title)}</h2>
          ${slide.lines.slice(0, 8).map((line) => `<p>${escapeHtml(line)}</p>`).join('')}
          <div class="pptx-note">Visual PPTX renderer did not load. For perfect layout, export this PowerPoint as PDF and upload the PDF.</div>
        `;
      }
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
    renderCurrentPage();
  }

  function nextPage() {
    if (state.currentPage >= state.totalPages) return;
    state.currentPage++;
    renderCurrentPage();
  }

  function jumpToPage(page) {
    const target = Math.min(Math.max(1, Number(page) || 1), state.totalPages);
    if (target === state.currentPage) return;
    state.currentPage = target;
    renderCurrentPage();
  }

  function setZoom(value) {
    state.zoom = Math.min(2.5, Math.max(0.45, Number(value)));
    renderCurrentPage();
  }

  function updateZoomLabel() {
    els.zoomLabel.textContent = `${Math.round(state.zoom * 100)}%`;
  }

  function handleKeyboard(event) {
    if (els.viewerView.classList.contains('hidden')) return;
    if (event.key === 'ArrowRight') nextPage();
    if (event.key === 'ArrowLeft') previousPage();
    if (event.key.toLowerCase() === 'f') toggleFullscreen();
  }

  async function toggleFullscreen() {
    if (!document.fullscreenElement) {
      await els.viewerStage.requestFullscreen().catch(() => {});
    } else {
      await document.exitFullscreen().catch(() => {});
    }
  }

  function revealToolbarTemporarily() {
    if (!els.viewerToolbar) return;
    els.viewerToolbar.classList.remove('toolbar-hidden');
    clearTimeout(state.toolbarHideTimer);
    state.toolbarHideTimer = setTimeout(() => {
      if (!els.controlPanel.classList.contains('hidden')) return;
      els.viewerToolbar.classList.add('toolbar-hidden');
    }, 3200);
  }

  function getGlobalTimingSeconds() {
    const selected = els.globalTimingSelect.value;
    if (selected === 'custom') return Math.max(1, Number(els.customTimingInput.value) || 10);
    return Number(selected) || 10;
  }

  function getCurrentTimingSeconds() {
    return state.perPageTiming[state.currentPage] || getGlobalTimingSeconds();
  }

  function startAutoPlay() {
    stopAutoPlay(false);
    state.autoPaused = false;
    scheduleNextAutoSlide();
    publishSessionState();
  }

  function scheduleNextAutoSlide() {
    clearTimeout(state.autoTimer);
    state.autoTimer = setTimeout(() => {
      if (state.autoPaused) return;
      if (state.currentPage >= state.totalPages) {
        stopAutoPlay();
        return;
      }
      state.currentPage++;
      renderCurrentPage().then(scheduleNextAutoSlide);
    }, getCurrentTimingSeconds() * 1000);
  }

  function pauseAutoPlay() {
    state.autoPaused = true;
    clearTimeout(state.autoTimer);
    publishSessionState();
  }

  function resumeAutoPlay() {
    if (!state.activeFile) return;
    state.autoPaused = false;
    scheduleNextAutoSlide();
    publishSessionState();
  }

  function stopAutoPlay(publish = true) {
    clearTimeout(state.autoTimer);
    state.autoTimer = null;
    state.autoPaused = false;
    if (publish) publishSessionState();
  }

  function showTimer() {
    state.timer.visible = true;
    els.timerOverlay.classList.remove('hidden');
    if (!state.timerStartedAt) state.timerStartedAt = Date.now();
    startTimerInterval();
    applyTimerSettings();
    publishSessionState();
  }

  function hideTimer() {
    state.timer.visible = false;
    els.timerOverlay.classList.add('hidden');
    stopTimerInterval();
    publishSessionState();
  }

  function resetTimer() {
    state.timerElapsedBeforePause = 0;
    state.timerStartedAt = Date.now();
    updateTimerText();
    if (state.timer.visible) startTimerInterval();
    publishSessionState();
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
    if (!state.timerStartedAt) return 0;
    return Math.floor((Date.now() - state.timerStartedAt) / 1000) + state.timerElapsedBeforePause;
  }

  function updateTimerText() {
    let seconds = elapsedSeconds();
    if (state.timer.mode === 'down') seconds = Math.max(0, state.timer.countdownSeconds - seconds);
    els.timerOverlay.textContent = formatTime(seconds);
  }

  function formatTime(totalSeconds) {
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return [hours, minutes, seconds].map((n) => String(n).padStart(2, '0')).join(':');
  }

  function applyTimerSettings() {
    els.timerOverlay.classList.remove('bottom-right', 'bottom-left', 'top-right', 'top-left');
    els.timerOverlay.classList.add(state.timer.position);
    els.timerOverlay.style.opacity = String(state.timer.opacity / 100);
    els.timerModeSelect.value = state.timer.mode;
    els.timerPositionSelect.value = state.timer.position;
    els.timerOpacityInput.value = state.timer.opacity;
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
      autoPlaying: Boolean(state.autoTimer && !state.autoPaused),
      autoPaused: state.autoPaused,
      globalTiming: getGlobalTimingSeconds(),
      perPageTiming: state.perPageTiming,
      timerVisible: state.timer.visible,
      timerOpacity: state.timer.opacity,
      timerMode: state.timer.mode,
      timerPosition: state.timer.position,
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
    try {
      if (state.activeFile.type === 'pdf' && state.activePdf) {
        return await renderPdfPageToDataUrl(state.activePdf, state.currentPage, 0.16);
      }
      if (state.activeFile.type === 'pptx') {
        if (state.pptxVisualReady && window.html2canvas && state.pptxRenderedSlides[state.currentPage - 1]) {
          const canvas = await window.html2canvas(state.pptxRenderedSlides[state.currentPage - 1], {
            backgroundColor: null,
            scale: 0.35,
            logging: false,
          });
          return canvas.toDataURL('image/jpeg', 0.72);
        }
        const slide = state.activePptxSlides[state.currentPage - 1];
        return createPptxThumbDataUrl(slide ? slide.title : `Slide ${state.currentPage}`, state.currentPage);
      }
    } catch (error) {
      return '';
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
      case 'resetZoom': setZoom(1); break;
      case 'autoStart': startAutoPlay(); break;
      case 'autoPause': pauseAutoPlay(); break;
      case 'autoStop': stopAutoPlay(); break;
      case 'setTiming':
        els.globalTimingSelect.value = 'custom';
        els.customTimingWrap.classList.remove('hidden');
        els.customTimingInput.value = Math.max(1, Number(command.value) || 10);
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
    els.qrModal.classList.remove('hidden');
  }

  function renderRemoteApp() {
    els.app.classList.add('hidden');
    els.remoteApp.classList.remove('hidden');
    const sessionId = qs.get('session') || '';
    const role = qs.get('role') || 'viewer';
    const isHost = role === 'host';

    els.remoteApp.innerHTML = `
      <main class="remote-card ${isHost ? '' : 'remote-viewer-only'}">
        <h1>Presentation Remote</h1>
        <p class="remote-sub">${escapeHtml(sessionId || 'No session')} • ${isHost ? 'Host control' : 'Viewer only'}</p>
        <div class="remote-preview" id="remotePreview"><span>Waiting for presentation...</span></div>
        <h2 id="remoteSlideLabel">Slide -- / --</h2>
        <p id="remoteFileLabel" class="remote-sub">Connect to an active desktop session.</p>
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
          <div class="remote-wide">
            <input id="remoteTiming" type="number" min="1" value="10" placeholder="Seconds">
            <button id="remoteSetTiming" data-host-only="true">Change Timing</button>
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
    `;

    if (!hasFirebaseConfig()) {
      $('remoteFileLabel').textContent = 'Remote needs firebase-config.js because phones need realtime sync across devices.';
      return;
    }

    initFirebaseIfConfigured().then(() => {
      if (!state.firebaseReady || !sessionId) {
        $('remoteFileLabel').textContent = 'Could not connect to Firebase or missing session.';
        return;
      }
      const ref = state.firebaseDb.collection(SESSION_COLLECTION).doc(sessionId);
      ref.onSnapshot((snap) => {
        if (!snap.exists) {
          $('remoteFileLabel').textContent = 'Session not found. Open QR from the desktop viewer first.';
          return;
        }
        const data = snap.data();
        $('remoteSlideLabel').textContent = `${data.type === 'pdf' ? 'Page' : 'Slide'} ${data.currentPage || '--'} / ${data.totalPages || '--'}`;
        $('remoteFileLabel').textContent = data.fileName || 'Active presentation';
        const preview = $('remotePreview');
        preview.innerHTML = data.thumb ? `<img src="${data.thumb}" alt="Current slide preview">` : '<span>No preview yet</span>';
        $('remoteTiming').value = data.globalTiming || 10;
        $('remoteOpacity').value = data.timerOpacity ?? 75;
      });

      els.remoteApp.querySelectorAll('[data-command]').forEach((button) => {
        button.addEventListener('click', () => {
          if (!isHost) return;
          sendRemoteCommand(ref, button.dataset.command);
        });
      });
      $('remoteSetTiming').addEventListener('click', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setTiming', Number($('remoteTiming').value) || 10);
      });
      $('remoteSetOpacity').addEventListener('click', () => {
        if (!isHost) return;
        sendRemoteCommand(ref, 'setTimerOpacity', Number($('remoteOpacity').value));
      });
    });
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
