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
  const MEDIA_CAST_ICE_SERVERS = {
    iceServers: [
      { urls: 'stun:stun.l.google.com:19302' },
      { urls: 'stun:stun1.l.google.com:19302' }
    ]
  };
  const MEDIA_CAST_CHUNK_SIZE = 64 * 1024;
  const MEDIA_CAST_MAX_FILE_BYTES = 350 * 1024 * 1024;

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
    transitionDuration: 920,
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
    remoteCommandPollTimer: null,
    processedRemoteCommands: new Set(),
    publishLock: false,
    inkStrokes: [],
    remoteSlideThumbs: {},
    remoteSlideThumbsCount: 0,
    remoteSlideThumbsBusy: false,
    magicEffectTimer: null,
    activeMagicToggle: null,
    magicEffectVolume: 200,
    magicEffectSound: true,
    magicEffectIntensity: 'grand',
    classroom: { sections: [], activeSectionId: '', groupCount: 6, removePicked: true, balanced: true, groupRollMode: 'balanced', rolledGroups: [], pickedIds: [], assignments: {}, lastStudent: null, lastGroup: null, revealMode: '', busy: false },
    classroomSyncTimer: null,
    firebaseUser: null,
    mediaCastReceiver: null,
    mediaCastReceiverId: '',
    mediaCastBlobUrl: '',
    mediaMode: false,
    mediaLinkVolume: 0.9,
  };




  const MAGIC_EFFECTS = [
    { id: 'drumroll', label: 'Drumroll', emoji: '🥁', shortcut: 'D', hint: 'Suspense before reveal' },
    { id: 'confetti', label: 'Confetti', emoji: '🎉', shortcut: 'C', hint: 'Celebration burst' },
    { id: 'micdrop', label: 'Mic Drop', emoji: '🎤', shortcut: 'M', hint: 'Strong ending moment' },
    { id: 'curtain', label: 'Curtain Reveal', emoji: '🎭', shortcut: 'U', hint: 'Reveal with curtains' },
    { id: 'bubbles', label: 'Bubbles', emoji: '🫧', shortcut: 'O', hint: 'Light floating effect' },
    { id: 'blur', label: 'Blur Screen', emoji: '🌫️', shortcut: 'B', hint: 'Hide answers until toggled again' },
    { id: 'quiet', label: 'Be Quiet', emoji: '🤫', shortcut: 'Q', hint: 'Calm class reminder' },
    { id: 'applause', label: 'Applause', emoji: '👏', shortcut: 'P', hint: 'Clap for students' },
    { id: 'spotlight', label: 'Spotlight', emoji: '🔦', shortcut: 'S', hint: 'Focus attention' },
    { id: 'correct', label: 'Correct Stamp', emoji: '✅', shortcut: 'A', hint: 'Correct answer' },
    { id: 'wrong', label: 'Wrong Buzzer', emoji: '❌', shortcut: 'X', hint: 'Fun wrong answer' },
    { id: 'timesup', label: "Time's Up", emoji: '⏰', shortcut: 'T', hint: 'End of time' },
    { id: 'sparkle', label: 'Sparkle', emoji: '✨', shortcut: 'K', hint: 'Highlight moment' },
    { id: 'stars', label: 'Star Rain', emoji: '🌟', shortcut: 'R', hint: 'Soft celebration' },
    { id: 'hype', label: 'Hype Burst', emoji: '🔥', shortcut: 'H', hint: 'Energy boost' },
    { id: 'freeze', label: 'Freeze', emoji: '🧊', shortcut: 'Z', hint: 'Pause for suspense' },
  ];

  const MAGIC_EFFECT_MAP = MAGIC_EFFECTS.reduce((acc, effect) => {
    acc[effect.id] = effect;
    acc[effect.shortcut.toLowerCase()] = effect;
    return acc;
  }, {});

  function getMagicEffectByShortcut(key) {
    const normalized = String(key || '').trim().toLowerCase();
    if (!normalized) return null;
    return MAGIC_EFFECTS.find((effect) => String(effect.shortcut || '').toLowerCase() === normalized) || null;
  }

  function magicParticleMarkup(kind, count = 72, symbols = ['✨']) {
    return Array.from({ length: count }, (_, index) => {
      const x = Math.round((index * 37 + Math.random() * 18) % 100);
      const size = Math.round(9 + Math.random() * 18);
      const delay = (Math.random() * 0.9).toFixed(2);
      const drift = Math.round((Math.random() - 0.5) * 260);
      const symbol = symbols[index % symbols.length];
      const style = `--x:${x}%;--s:${size}px;--d:${delay}s;--drift:${drift}px`;
      if (kind === 'symbol') return `<i class="magic-symbol" style="${style}">${symbol}</i>`;
      if (kind === 'bubble') return `<i class="magic-bubble" style="${style}"></i>`;
      return `<i class="magic-confetti" style="${style}"></i>`;
    }).join('');
  }

  function magicCenter(icon, title = '', subtitle = '', extraClass = '') {
    // Magic effects are visual-only on the presentation screen: no card titles,
    // no labels, and no subtitle text. The phone remote can still show names.
    return `<div class="magic-center magic-visual-only ${extraClass}"><div class="magic-emoji" aria-hidden="true">${icon}</div></div>`;
  }

  function magicEffectMarkup(effectId) {
    switch (effectId) {
      case 'drumroll':
        return `<div class="magic-stage-flash amber"></div><div class="magic-center magic-plain magic-visual-only"><div class="magic-drum-hero magic-drum-big-center"><span class="magic-stick left"></span><span class="magic-stick right"></span><span class="magic-big-drum">🥁</span></div></div>`;
      case 'confetti':
        return `<canvas class="magic-fx-canvas"></canvas><div class="magic-stage-flash"></div><div class="magic-glow-ring rainbow"></div><div class="magic-particles confetti-full">${magicParticleMarkup('confetti', 120)}</div>${magicCenter('🎉', '', '', 'magic-pop')}`;
      case 'micdrop':
        return `<style>
          #magicEffectLayer .ph-micdrop-safe{position:fixed!important;inset:0!important;z-index:2147483640!important;pointer-events:none!important;display:block!important;overflow:hidden!important;isolation:isolate!important}
          #magicEffectLayer .ph-micdrop-safe *{box-sizing:border-box!important}
          #magicEffectLayer .ph-micdrop-flash{position:absolute!important;inset:0!important;z-index:1!important;background:radial-gradient(circle at center,rgba(255,239,184,.42),rgba(250,204,21,.18) 30%,transparent 62%)!important;animation:phMicdropFlash 2.45s ease-out both!important}
          #magicEffectLayer .ph-micdrop-glow{position:absolute!important;left:50%!important;top:73%!important;width:min(56vw,720px)!important;height:min(24vw,310px)!important;border-radius:50%!important;transform:translate(-50%,-50%)!important;background:radial-gradient(ellipse,rgba(255,209,102,.75),rgba(249,115,22,.28) 46%,transparent 75%)!important;filter:blur(22px)!important;opacity:.95!important;z-index:2!important;animation:phMicdropGlow 2.45s ease-out both!important}
          #magicEffectLayer .ph-micdrop-mic{position:absolute!important;left:50%!important;top:12vh!important;width:clamp(92px,11vw,170px)!important;height:auto!important;z-index:12!important;opacity:1!important;transform:translate(-50%,0) rotate(-14deg) scale(1)!important;transform-origin:50% 76%!important;filter:drop-shadow(0 32px 24px rgba(0,0,0,.5))!important;animation:phMicdropFall 2.45s cubic-bezier(.18,.78,.18,1.06) both!important}
          #magicEffectLayer .ph-micdrop-shadow{position:absolute!important;left:50%!important;top:82%!important;width:clamp(190px,29vw,470px)!important;height:clamp(28px,4vw,62px)!important;border-radius:50%!important;z-index:3!important;transform:translate(-50%,-50%) scale(.25)!important;background:radial-gradient(ellipse,rgba(0,0,0,.58),rgba(0,0,0,.18) 58%,transparent 78%)!important;filter:blur(4px)!important;opacity:0!important;animation:phMicdropShadow 2.45s ease-out both!important}
          #magicEffectLayer .ph-micdrop-ring{position:absolute!important;left:50%!important;top:81%!important;width:clamp(160px,23vw,390px)!important;height:clamp(54px,6vw,120px)!important;border:clamp(4px,.55vw,9px) solid rgba(255,255,255,.82)!important;border-radius:50%!important;z-index:4!important;transform:translate(-50%,-50%) scale(.25)!important;opacity:0!important;animation:phMicdropRing 2.45s ease-out both!important}
          #magicEffectLayer .ph-micdrop-burst{position:absolute!important;left:50%!important;top:74%!important;width:clamp(250px,34vw,560px)!important;height:clamp(150px,18vw,290px)!important;transform:translate(-50%,-50%)!important;z-index:5!important}
          #magicEffectLayer .ph-micdrop-burst i{position:absolute!important;left:50%!important;bottom:25%!important;width:clamp(5px,.6vw,10px)!important;height:0!important;border-radius:999px!important;background:linear-gradient(180deg,rgba(255,255,255,1),rgba(251,191,36,.9),transparent)!important;transform-origin:50% 100%!important;opacity:0!important;animation:phMicdropBurst 2.45s ease-out both!important}
          #magicEffectLayer .ph-micdrop-burst i:nth-child(1){transform:rotate(-68deg)!important;animation-delay:1.28s!important}#magicEffectLayer .ph-micdrop-burst i:nth-child(2){transform:rotate(-42deg)!important;animation-delay:1.30s!important}#magicEffectLayer .ph-micdrop-burst i:nth-child(3){transform:rotate(-15deg)!important;animation-delay:1.32s!important}#magicEffectLayer .ph-micdrop-burst i:nth-child(4){transform:rotate(15deg)!important;animation-delay:1.32s!important}#magicEffectLayer .ph-micdrop-burst i:nth-child(5){transform:rotate(42deg)!important;animation-delay:1.30s!important}#magicEffectLayer .ph-micdrop-burst i:nth-child(6){transform:rotate(68deg)!important;animation-delay:1.28s!important}
          @keyframes phMicdropFall{0%{opacity:1;transform:translate(-50%,0) rotate(-14deg) scale(1)}42%{transform:translate(-50%,43vh) rotate(16deg) scale(1.06)}56%{transform:translate(-50%,57vh) rotate(25deg) scale(1.1)}70%{transform:translate(-50%,47vh) rotate(-9deg) scale(1.04)}84%{transform:translate(-50%,55vh) rotate(11deg) scale(1.02)}100%{opacity:1;transform:translate(-50%,52vh) rotate(-7deg) scale(1)}}
          @keyframes phMicdropShadow{0%,34%{opacity:0;transform:translate(-50%,-50%) scale(.22)}56%{opacity:.66;transform:translate(-50%,-50%) scale(1.22)}100%{opacity:.48;transform:translate(-50%,-50%) scale(1)}}
          @keyframes phMicdropRing{0%,46%{opacity:0;transform:translate(-50%,-50%) scale(.22)}58%{opacity:.98;transform:translate(-50%,-50%) scale(.86)}100%{opacity:0;transform:translate(-50%,-50%) scale(2.05)}}
          @keyframes phMicdropBurst{0%,48%{opacity:0;height:0}60%{opacity:1;height:clamp(52px,8vw,120px)}100%{opacity:0;height:0}}
          @keyframes phMicdropGlow{0%{opacity:0;transform:translate(-50%,-50%) scale(.7)}34%,80%{opacity:.95;transform:translate(-50%,-50%) scale(1)}100%{opacity:0;transform:translate(-50%,-50%) scale(1.18)}}
          @keyframes phMicdropFlash{0%{opacity:0}18%{opacity:1}100%{opacity:0}}
          @media(max-width:700px){#magicEffectLayer .ph-micdrop-mic{top:13vh!important;width:clamp(92px,24vw,150px)!important}#magicEffectLayer .ph-micdrop-shadow{top:83%!important}#magicEffectLayer .ph-micdrop-ring{top:82%!important}#magicEffectLayer .ph-micdrop-burst{top:75%!important}}
        </style><div class="ph-micdrop-safe" aria-hidden="true"><div class="ph-micdrop-flash"></div><div class="ph-micdrop-glow"></div><svg class="ph-micdrop-mic" viewBox="0 0 120 310" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><defs><linearGradient id="phMicHead" x1="10" y1="0" x2="108" y2="104" gradientUnits="userSpaceOnUse"><stop offset="0" stop-color="#ffffff"/><stop offset="0.24" stop-color="#cbd5e1"/><stop offset="0.56" stop-color="#475569"/><stop offset="1" stop-color="#020617"/></linearGradient><linearGradient id="phMicGold" x1="0" y1="0" x2="120" y2="0" gradientUnits="userSpaceOnUse"><stop offset="0" stop-color="#f59e0b"/><stop offset="0.5" stop-color="#fff7cc"/><stop offset="1" stop-color="#f59e0b"/></linearGradient><linearGradient id="phMicHandle" x1="35" y1="130" x2="96" y2="294" gradientUnits="userSpaceOnUse"><stop offset="0" stop-color="#f8fafc"/><stop offset="0.36" stop-color="#94a3b8"/><stop offset="1" stop-color="#111827"/></linearGradient></defs><rect x="19" y="8" width="82" height="128" rx="41" fill="url(#phMicHead)" stroke="rgba(255,255,255,.9)" stroke-width="8"/><path d="M31 36 H89 M28 59 H92 M28 82 H92 M33 105 H87" stroke="rgba(255,255,255,.62)" stroke-width="8" stroke-linecap="round"/><rect x="15" y="128" width="90" height="34" rx="17" fill="url(#phMicGold)" stroke="rgba(255,255,255,.82)" stroke-width="6"/><rect x="38" y="153" width="44" height="132" rx="21" fill="url(#phMicHandle)" stroke="rgba(255,255,255,.75)" stroke-width="7"/><rect x="52" y="278" width="16" height="30" rx="8" fill="#020617" opacity=".82"/></svg><div class="ph-micdrop-shadow"></div><div class="ph-micdrop-ring"></div><div class="ph-micdrop-burst"><i></i><i></i><i></i><i></i><i></i><i></i></div></div>`;
      case 'curtain':
        return `<div class="magic-curtain-stage"><div class="magic-curtain left"></div><div class="magic-curtain right"></div><div class="magic-curtain valance"></div><span class="magic-curtain-tie left"></span><span class="magic-curtain-tie right"></span></div>`;
      case 'bubbles':
        return `<canvas class="magic-fx-canvas magic-bubbles-canvas"></canvas><div class="magic-glow-ring"></div><div class="magic-particles bubbles-full">${magicParticleMarkup('bubble', 86)}</div>${magicCenter('🫧', '', '', 'magic-pop')}`;
      case 'blur':
        return `<div class="magic-screen-blur"><span class="magic-blur-cloud one"></span><span class="magic-blur-cloud two"></span><span class="magic-blur-cloud three"></span></div>`;
      case 'quiet':
        return `<div class="magic-quiet-backdrop"></div><div class="magic-quiet-halo"></div><div class="magic-quiet-halo halo-two"></div><div class="magic-center magic-plain magic-quiet-hero magic-visual-only"><span class="magic-quiet-aura"></span><span class="magic-shush-wave"></span><span class="magic-shush-wave two"></span><span class="magic-shush-wave three"></span><div class="magic-emoji">🤫</div></div>`;
      case 'applause':
        return `<div class="magic-stage-flash"></div><div class="magic-glow-ring green"></div><div class="magic-particles applause-full">${magicParticleMarkup('symbol', 84, ['👏','🙌','👏','✨'])}</div><div class="magic-center magic-plain magic-applause-hero magic-visual-only"><div class="magic-emoji">👏</div></div>`;
      case 'spotlight':
        return `<div class="magic-spotlight-dimmer"></div><div class="magic-moving-spotlight"></div>`;
      case 'correct':
        return `<div class="magic-stage-flash"></div><div class="magic-glow-ring green"></div><div class="magic-center magic-plain magic-big-mark correct magic-visual-only"><div class="magic-mark">✅</div></div>`;
      case 'wrong':
        return `<div class="magic-stage-flash red"></div><div class="magic-center magic-plain magic-big-mark wrong magic-visual-only"><div class="magic-mark">❌</div></div>`;
      case 'timesup':
        return `<div class="magic-time-rings"></div><div class="magic-center magic-plain magic-clock-wrap magic-visual-only"><div class="magic-clock">⏰</div></div>`;
      case 'sparkle':
        return `<div class="magic-stage-flash"></div><div class="magic-glow-ring rainbow"></div><div class="magic-particles">${magicParticleMarkup('symbol', 90, ['✨','✦','✧','💫'])}</div><div class="magic-center magic-plain magic-sparkle-hero magic-visual-only"><div class="magic-emoji">✨</div></div>`;
      case 'stars':
        return `<div class="magic-glow-ring rainbow"></div><div class="magic-particles">${magicParticleMarkup('symbol', 96, ['⭐','🌟','✦','✨'])}</div><div class="magic-center magic-plain magic-stars-hero magic-visual-only"><div class="magic-emoji">🌟</div></div>`;
      case 'hype':
        return `<div class="magic-stage-flash amber"></div><div class="magic-glow-ring fire"></div><div class="magic-particles">${magicParticleMarkup('symbol', 88, ['🔥','⚡','✨'])}</div><div class="magic-center magic-plain magic-hype-hero magic-visual-only"><div class="magic-emoji">🔥</div></div>`;
      case 'freeze':
        return `<div class="magic-freeze"></div><div class="magic-center magic-plain magic-freeze-hero magic-visual-only"><div class="magic-emoji">🧊</div></div>`;
      default:
        return magicCenter('✨', '', '', 'magic-pop');
    }
  }

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
    loadClassroomLocal();
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
    const existingBalancedInput = $('balancedGroupsInput');
    const existingBalancedLabel = existingBalancedInput && existingBalancedInput.closest('label');
    if (existingBalancedLabel && !$('groupRollModeSelect')) {
      existingBalancedLabel.style.display = 'none';
      const modeLabel = document.createElement('label');
      modeLabel.className = 'classroom-group-mode-label';
      modeLabel.innerHTML = `Group rolls<select id="groupRollModeSelect"><option value="balanced">Balanced</option><option value="no-repeat">No repeat</option><option value="free">Free repeat</option></select>`;
      existingBalancedLabel.insertAdjacentElement('afterend', modeLabel);
    }
    [
      'app', 'remoteApp', 'homeView', 'viewerView', 'fileInput', 'uploadZone', 'searchInput', 'sortSelect',
      'cardsGrid', 'emptyState', 'libraryCount', 'clearLibraryBtn', 'themeToggle', 'firebaseStatus',
      'mediaViewerBtn', 'createFolderBtn', 'folderList',
      'thumbnailSidebar', 'viewerStage', 'viewerToolbar', 'controlPanel', 'settingsBtn', 'backHomeBtn',
      'prevBtn', 'nextBtn', 'jumpInput', 'pageTotalLabel', 'zoomOutBtn', 'zoomInBtn', 'resetZoomBtn',
      'zoomLabel', 'fullscreenBtn', 'qrBtn', 'viewerCanvasWrap', 'pdfCanvas', 'pptxSlide', 'inkCanvas',
      'timingModeSelect', 'globalTimingSelect', 'customTimingWrap', 'customTimingInput', 'perSlideTimingWrap', 'perSlideTimingInput',
      'countdownAlertSelect', 'countdownVoiceSelect', 'countdownVoiceStartSelect', 'soundTestBtn', 'slideTransitionSelect',
      'autoStartBtn', 'autoPauseBtn', 'autoResumeBtn', 'autoStopBtn', 'timerOverlay', 'timerModeSelect',
      'countdownMinutesInput', 'timerPositionSelect', 'timerOpacityInput', 'timerSizeInput', 'timerSizeLabel', 'timerShowBtn', 'timerHideBtn',
      'timerResetBtn', 'qrModal', 'closeQrBtn', 'hostQr', 'viewerQr', 'hostRemoteLink', 'viewerRemoteLink',
      'qrHelp', 'setupModal', 'closeSetupBtn', 'classroomBtn', 'classroomModal', 'closeClassroomBtn', 'classroomAuthBtn', 'classroomAuthStatus', 'sectionSelect', 'newSectionName', 'addSectionBtn', 'studentNamesInput', 'classListFileInput', 'importClassListBtn', 'classImportStatus', 'saveNamesBtn', 'groupCountInput', 'removePickedInput', 'balancedGroupsInput', 'groupRollModeSelect', 'pickNameBtn', 'rollGroupBtn', 'resetPicksBtn', 'presentClassroomBtn', 'classroomResult', 'classroomRoster'
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
    if (els.themeToggle) els.themeToggle.addEventListener('click', toggleTheme);
    if (els.firebaseStatus) els.firebaseStatus.addEventListener('click', () => showModal(els.setupModal));
    if (els.mediaViewerBtn) els.mediaViewerBtn.addEventListener('click', openMediaViewer);
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

    if (els.classroomBtn) els.classroomBtn.addEventListener('click', openClassroomTools);
    if (els.closeClassroomBtn) els.closeClassroomBtn.addEventListener('click', () => hideModal(els.classroomModal));
    if (els.presentClassroomBtn) els.presentClassroomBtn.addEventListener('click', openClassroomPresentationMode);
    if (els.classroomAuthBtn) els.classroomAuthBtn.addEventListener('click', signInClassroomAccount);
    if (els.addSectionBtn) els.addSectionBtn.addEventListener('click', addClassroomSection);
    if (els.sectionSelect) els.sectionSelect.addEventListener('change', () => { state.classroom.activeSectionId = els.sectionSelect.value; state.classroom.pickedIds = []; state.classroom.rolledGroups = []; renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true); });
    if (els.saveNamesBtn) els.saveNamesBtn.addEventListener('click', saveClassroomNames);
    if (els.importClassListBtn) els.importClassListBtn.addEventListener('click', () => els.classListFileInput && els.classListFileInput.click());
    if (els.classListFileInput) els.classListFileInput.addEventListener('change', importClassListFile);
    if (els.groupCountInput) els.groupCountInput.addEventListener('change', () => {
      state.classroom.groupCount = Math.max(2, Math.min(20, Number(els.groupCountInput.value) || 6));
      state.classroom.rolledGroups = [];
      scheduleClassroomSave(); publishSessionState(true);
    });
    if (els.removePickedInput) els.removePickedInput.addEventListener('change', () => { state.classroom.removePicked = els.removePickedInput.checked; scheduleClassroomSave(); publishSessionState(true); });
    if (els.groupRollModeSelect) els.groupRollModeSelect.addEventListener('change', () => {
      const mode = ['balanced', 'no-repeat', 'free'].includes(els.groupRollModeSelect.value) ? els.groupRollModeSelect.value : 'balanced';
      state.classroom.groupRollMode = mode;
      state.classroom.balanced = mode === 'balanced';
      state.classroom.rolledGroups = [];
      scheduleClassroomSave(); publishSessionState(true);
    });
    if (els.balancedGroupsInput) els.balancedGroupsInput.addEventListener('change', () => {
      state.classroom.groupRollMode = els.balancedGroupsInput.checked ? 'balanced' : 'free';
      state.classroom.balanced = els.balancedGroupsInput.checked;
      state.classroom.rolledGroups = [];
      scheduleClassroomSave(); publishSessionState(true);
    });
    if (els.pickNameBtn) els.pickNameBtn.addEventListener('click', pickRandomStudent);
    if (els.rollGroupBtn) els.rollGroupBtn.addEventListener('click', rollGroupDice);
    if (els.resetPicksBtn) els.resetPicksBtn.addEventListener('click', resetClassroomRound);

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
      let user = firebase.auth().currentUser;
      if (!user) {
        const cred = await firebase.auth().signInAnonymously();
        user = cred.user;
      }
      state.firebaseUser = user;
      state.firebaseAuthReady = true;
      state.firebaseReady = true;
      firebase.auth().onAuthStateChanged(async (nextUser) => {
        state.firebaseUser = nextUser || null;
        updateFirebaseStatus();
        updateClassroomAuthUI();
        if (nextUser && !nextUser.isAnonymous) await loadClassroomCloud();
      });
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
      els.firebaseStatus.textContent = state.firebaseUser && !state.firebaseUser.isAnonymous ? 'Account synced' : 'Remote ready';
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
    if (!window.jspdf || !window.jspdf.jsPDF) {
      await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
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
      const allFiles = await dbGetAll();
      const removedCanva = allFiles.filter((file) => file && file.type === 'canva');
      state.files = allFiles.filter((file) => file && file.type !== 'canva');
      if (removedCanva.length) removedCanva.forEach((file) => dbDelete(file.id).catch(() => undefined));
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
    textarea.innerHTML = String(value || '');
    return textarea.value;
  }

  function decodeMaybeEncodedUrl(value) {
    let output = String(value || '').trim();
    for (let i = 0; i < 3; i++) {
      const before = output;
      output = decodeLooseHtml(output).replace(/&amp;/gi, '&').trim();
      try {
        output = decodeURIComponent(output).trim();
      } catch (error) {}
      if (output === before) break;
    }
    return output;
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
        note = 'PowerPoint will auto-convert into a PDF-style viewer when opened. Conversion uses the browser renderer and may take a moment.';
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
    const typeLabel = file.type === 'pdf' ? 'PDF' : 'PPTX';
    const countLabel = file.type === 'pdf' ? `${file.pageCount} pages` : `${file.pageCount} slides`;
    return `
      <article class="presentation-card glass" data-card-id="${file.id}" draggable="true" title="Drag this card to a folder">
        <span class="file-badge">${typeLabel}</span>
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
    closeMediaCastOverlay(false);
    state.mediaMode = false;
    els.viewerView.classList.remove('media-viewer-mode');
    const mediaWelcome = document.getElementById('mediaViewerWelcome');
    if (mediaWelcome) mediaWelcome.remove();
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
    els.pageTotalLabel.textContent = `/ ${state.totalPages}`;
    els.jumpInput.max = state.totalPages;
    els.jumpInput.value = '1';
    updateZoomLabel();
    updateTimingModeUI();
    updateViewerNavigationUI();

    try {
      if (file.type === 'pdf') {
        const buffer = await file.blob.arrayBuffer();
        state.activePdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        state.totalPages = state.activePdf.numPages;
      } else if (file.type === 'pptx') {
        await preparePptxVisualRenderer(file.blob);
        if (state.pptxVisualReady) {
          const convertedBlob = await convertRenderedPptxToPdfBlob();
          if (convertedBlob) {
            const buffer = await convertedBlob.arrayBuffer();
            state.activePdf = await pdfjsLib.getDocument({ data: buffer }).promise;
            state.totalPages = state.activePdf.numPages;
            state.activeFile.originalType = 'pptx';
            state.activeFile.type = 'pdf';
            state.activeFile.convertedFromPptx = true;
            state.pptxVisualReady = false;
            els.pptxSlide.classList.add('hidden');
          }
        }
        if (!state.activePdf && !state.pptxVisualReady) {
          state.activePptxSlides = await parsePptxSlides(file.blob);
          state.totalPages = state.activePptxSlides.length;
        }
      }
      file.pageCount = state.totalPages;
      els.pageTotalLabel.textContent = `/ ${state.totalPages}`;
      els.jumpInput.max = state.totalPages;
      updateViewerNavigationUI();
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
    els.viewerView.classList.remove('presentation-fullscreen', 'media-viewer-mode');
    stopAutoPlay();
    stopTimerInterval();
    closeMediaCastOverlay(false);
    state.mediaMode = false;
    const mediaWelcome = document.getElementById('mediaViewerWelcome');
    if (mediaWelcome) mediaWelcome.remove();
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
    state.activeFile = null;
    els.viewerView.classList.add('hidden');
    els.homeView.classList.remove('hidden');
    renderLibrary();
  }


  function ensureMediaViewerWelcome() {
    if (!els.viewerCanvasWrap) return null;
    let welcome = document.getElementById('mediaViewerWelcome');
    if (!welcome) {
      welcome = document.createElement('div');
      welcome.id = 'mediaViewerWelcome';
      welcome.className = 'media-viewer-welcome glass';
      els.viewerCanvasWrap.appendChild(welcome);
    }
    welcome.innerHTML = `
      <div class="media-viewer-welcome-icon">🎬</div>
      <h2>Media Viewing Mode</h2>
      <p>Use the phone host remote to cast pictures, MP3, MP4, or paste online links like YouTube, TikTok, and direct media URLs.</p>
      <div class="media-viewer-welcome-actions">
        <button id="mediaViewerQrNow" type="button">Open Remote QR</button>
        <button id="mediaViewerFullscreenNow" type="button">Fullscreen Screen</button>
      </div>
      <small>No PDF/PPT is needed. Media is shown only during this live session.</small>
    `;
    const qr = welcome.querySelector('#mediaViewerQrNow');
    if (qr) qr.addEventListener('click', openQrModal);
    const full = welcome.querySelector('#mediaViewerFullscreenNow');
    if (full) full.addEventListener('click', toggleFullscreen);
    return welcome;
  }

  async function openMediaViewer() {
    stopAutoPlay();
    stopTimerInterval();
    state.mediaMode = true;
    state.activeFile = null;
    state.activePdf = null;
    state.activePptxSlides = [];
    state.pptxVisualReady = false;
    state.currentPage = 0;
    state.totalPages = 0;
    state.zoom = 1;
    state.viewportCenterX = 0.5;
    state.viewportCenterY = 0.5;
    state.inkStrokes = [];
    state.remoteSlideThumbs = {};
    state.remoteSlideThumbsCount = 0;
    if (state.pptxBlobUrl) {
      try { URL.revokeObjectURL(state.pptxBlobUrl); } catch (error) {}
      state.pptxBlobUrl = '';
    }
    resetPdfLayers();
    els.homeView.classList.add('hidden');
    els.viewerView.classList.remove('hidden');
    els.viewerView.classList.add('media-viewer-mode');
    if (els.thumbnailSidebar) els.thumbnailSidebar.innerHTML = '';
    if (els.pdfCanvas) els.pdfCanvas.classList.add('hidden');
    if (els.pptxSlide) els.pptxSlide.classList.add('hidden');
    if (els.inkCanvas) els.inkCanvas.classList.add('hidden');
    if (els.jumpInput) { els.jumpInput.value = ''; els.jumpInput.disabled = true; els.jumpInput.placeholder = 'Media'; }
    if (els.pageTotalLabel) els.pageTotalLabel.textContent = '/ Media';
    if (els.prevBtn) els.prevBtn.disabled = true;
    if (els.nextBtn) els.nextBtn.disabled = true;
    updateZoomLabel();
    ensureMediaViewerWelcome();
    revealToolbarTemporarily();
    if (!state.firebaseReady && hasFirebaseConfig()) await initFirebaseIfConfigured();
    await setupRemoteSessionIfPossible();
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

  async function convertRenderedPptxToPdfBlob() {
    if (!state.pptxVisualReady || !state.pptxRenderedSlides.length || !window.html2canvas || !window.jspdf || !window.jspdf.jsPDF) return null;
    const sourceSlides = state.pptxRenderedSlides.slice();
    els.pptxSlide.innerHTML = '<div class="viewer-message">Converting PowerPoint to PDF viewer...</div>';
    await wait(120);
    try {
      const { jsPDF } = window.jspdf;
      let pdf = null;
      for (let index = 0; index < sourceSlides.length; index++) {
        const slide = sourceSlides[index];
        slide.classList.remove('hidden');
        slide.style.display = '';
        slide.style.visibility = 'visible';
        slide.style.opacity = '1';
        els.pptxSlide.innerHTML = '';
        els.pptxSlide.appendChild(slide);
        await wait(80);
        const canvas = await html2canvas(slide, {
          backgroundColor: '#ffffff',
          scale: Math.min(2, window.devicePixelRatio || 1.5),
          useCORS: true,
          allowTaint: true,
          logging: false,
        });
        const width = Math.max(1, canvas.width);
        const height = Math.max(1, canvas.height);
        const orientation = width >= height ? 'landscape' : 'portrait';
        const image = canvas.toDataURL('image/jpeg', 0.95);
        if (!pdf) pdf = new jsPDF({ orientation, unit: 'px', format: [width, height], compress: true });
        else pdf.addPage([width, height], orientation);
        pdf.addImage(image, 'JPEG', 0, 0, width, height);
      }
      return pdf ? pdf.output('blob') : null;
    } catch (error) {
      console.warn('PPTX to PDF-style conversion failed:', error);
      return null;
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

    // Normal viewer must open in a true fit-to-page layout: the whole PDF/PPT
    // page should be visible below the floating toolbar, like a presentation
    // preview, instead of being rendered too tall and getting cut off.
    const wrapRect = els.viewerCanvasWrap ? els.viewerCanvasWrap.getBoundingClientRect() : null;
    if (wrapRect && wrapRect.width > 120 && wrapRect.height > 120) {
      const style = window.getComputedStyle ? window.getComputedStyle(els.viewerCanvasWrap) : null;
      const padX = style ? (parseFloat(style.paddingLeft) || 0) + (parseFloat(style.paddingRight) || 0) : 0;
      const padY = style ? (parseFloat(style.paddingTop) || 0) + (parseFloat(style.paddingBottom) || 0) : 0;
      const stageRect = els.viewerStage ? els.viewerStage.getBoundingClientRect() : { height: window.innerHeight };
      const toolbarRect = els.viewerToolbar ? els.viewerToolbar.getBoundingClientRect() : null;
      const toolbarReserve = toolbarRect ? Math.max(96, toolbarRect.height + 44) : 110;
      const safeHeightBelowToolbar = Math.max(240, stageRect.height - toolbarReserve);
      return {
        width: Math.max(320, wrapRect.width - padX),
        height: Math.max(240, Math.min(wrapRect.height - padY, safeHeightBelowToolbar)),
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
    updateViewerNavigationUI();
    publishSessionState();

    if (state.activeFile.type === 'pdf' && state.activePdfRenderTask) {
      try { state.activePdfRenderTask.cancel(); } catch (error) {}
      state.activePdfRenderTask = null;
    }

    try {
      if (state.activeFile.type === 'pdf') {
            els.pptxSlide.classList.add('hidden');
        await renderPdfPageSmooth(token);
        if (token !== state.renderToken) return;
        warmAdjacentPdfPages();
        warmAdjacentPdfBitmaps();
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
              <strong>PowerPoint conversion is not available.</strong>
              <span>The app tried to convert this PPTX into a PDF-style viewer. For perfect fidelity, export the file as PDF from PowerPoint and upload the PDF.</span>
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

  function updateViewerNavigationUI() {
    if (!state.activeFile) return;
    if (els.prevBtn) els.prevBtn.disabled = state.currentPage <= 1;
    if (els.nextBtn) els.nextBtn.disabled = state.currentPage >= state.totalPages;
    if (els.jumpInput) {
      els.jumpInput.disabled = false;
      els.jumpInput.placeholder = '';
      els.jumpInput.max = state.totalPages || 1;
      els.jumpInput.value = String(state.currentPage || 1);
    }
    if (els.pageTotalLabel) els.pageTotalLabel.textContent = `/ ${state.totalPages || 1}`;
  }

  function createMagicLayer() {
    let layer = document.getElementById('magicEffectLayer');
    // Put magic effects on the fullscreen element when available; otherwise put
    // them directly on <body>. Appending to the slide/stage can be clipped by
    // transforms, overflow rules, or stacking contexts, which is why remote
    // effects could be received but never visually appear on the desktop.
    const host = document.fullscreenElement || document.body;
    if (!layer) {
      layer = document.createElement('div');
      layer.id = 'magicEffectLayer';
      layer.className = 'magic-effect-layer hidden';
    }
    if (layer.parentElement !== host) host.appendChild(layer);
    layer.className = 'magic-effect-layer hidden';
    layer.style.position = 'fixed';
    layer.style.inset = '0';
    layer.style.zIndex = '2147483600';
    layer.style.pointerEvents = 'none';
    layer.innerHTML = '';
    return layer;
  }

  function clearMagicEffectLayer(layer) {
    if (!layer) return;
    layer.classList.remove('show', 'magic-toggle-active', 'magic-curtain-close', 'magic-curtain-open', 'magic-blur-on', 'magic-blur-off');
    layer.classList.add('hidden');
    layer.style.display = 'none';
    layer.style.opacity = '';
    layer.style.visibility = '';
    layer.innerHTML = '';
  }

  function showMagicLayer(layer) {
    layer.style.display = 'grid';
    layer.style.opacity = '1';
    layer.style.visibility = 'visible';
    void layer.offsetWidth;
    layer.classList.add('show');
  }

  function togglePersistentMagicEffect(effect) {
    if (!['blur', 'curtain'].includes(effect.id)) return false;
    const existing = document.getElementById('magicEffectLayer');
    const intensity = state.magicEffectIntensity || 'grand';

    if (state.activeMagicToggle === effect.id && existing) {
      clearTimeout(state.magicEffectTimer);
      state.activeMagicToggle = null;
      existing.classList.remove('magic-toggle-active', 'magic-blur-on', 'magic-curtain-close');
      existing.classList.add(effect.id === 'curtain' ? 'magic-curtain-open' : 'magic-blur-off');
      showMagicLayer(existing);
      playMagicEffectSound(effect.id);
      state.magicEffectTimer = setTimeout(() => clearMagicEffectLayer(existing), effect.id === 'curtain' ? 1850 : 560);
      return true;
    }

    if (state.activeMagicToggle && existing) clearMagicEffectLayer(existing);
    const layer = createMagicLayer();
    layer.className = `magic-effect-layer magic-${effect.id} magic-toggle-active magic-intensity-${intensity} ${effect.id === 'curtain' ? 'magic-curtain-close' : 'magic-blur-on'}`;
    try {
      layer.innerHTML = magicEffectMarkup(effect.id);
    } catch (error) {
      console.warn('Magic effect markup failed; using safe fallback.', error);
      layer.innerHTML = magicCenter(effect.emoji || '✨', '', '', 'magic-pop');
    }
    showMagicLayer(layer);
    state.activeMagicToggle = effect.id;
    clearTimeout(state.magicEffectTimer);
    playMagicEffectSound(effect.id);
    return true;
  }

  function triggerMagicEffect(effectId) {
    const effectKey = String(effectId || 'confetti').trim().toLowerCase();
    const effect = MAGIC_EFFECT_MAP[effectKey] || MAGIC_EFFECT_MAP.confetti;
    if (togglePersistentMagicEffect(effect)) return;
    const oldLayer = document.getElementById('magicEffectLayer');
    if (state.activeMagicToggle && oldLayer) {
      clearMagicEffectLayer(oldLayer);
      state.activeMagicToggle = null;
    }
    const layer = createMagicLayer();
    layer.className = `magic-effect-layer magic-${effect.id} magic-intensity-${state.magicEffectIntensity || 'grand'}`;
    try {
      layer.innerHTML = magicEffectMarkup(effect.id);
    } catch (error) {
      console.warn('Magic effect markup failed; using safe fallback.', error);
      layer.innerHTML = magicCenter(effect.emoji || '✨', effect.label || 'Magic!', '', 'magic-pop');
    }
    // Inline visibility guards make the effect show even if an old CSS cache or
    // a stacking-context bug is still present on GitHub Pages.
    showMagicLayer(layer);
    if (effect.id === 'confetti' || effect.id === 'bubbles') startCanvasMagicEffect(layer, effect.id);
    playMagicEffectSound(effect.id);
    clearTimeout(state.magicEffectTimer);
    const durations = { drumroll: 3300, confetti: 3200, micdrop: 3200, curtain: 3400, bubbles: 8200, blur: 999999, quiet: 2200, applause: 2400, spotlight: 5400, correct: 2200, wrong: 2100, timesup: 2500, sparkle: 2300, stars: 2400, hype: 2300, freeze: 2300 };
    state.magicEffectTimer = setTimeout(() => {
      clearMagicEffectLayer(layer);
    }, durations[effect.id] || 2200);
  }

  function getMagicEffectVolume() {
    const value = Number(state.magicEffectVolume);
    return Math.max(0, Math.min(3, value / 100));
  }

  function setMagicEffectVolume(value, publish = true) {
    state.magicEffectVolume = Math.max(0, Math.min(300, Number(value) || 0));
    if (publish) publishSessionState();
  }

  function setMagicEffectSound(enabled, publish = true) {
    state.magicEffectSound = !!enabled;
    if (publish) publishSessionState();
  }

  function setMagicEffectIntensity(value, publish = true) {
    state.magicEffectIntensity = ['low', 'normal', 'grand'].includes(value) ? value : 'grand';
    if (publish) publishSessionState();
  }

  function playMagicEffectSound(effectId) {
    if (!state.magicEffectSound) return;
    primePresentationAudio();
    if (!state.audioContext) return;
    const ctx = state.audioContext;
    if (ctx.state === 'suspended') ctx.resume().catch(() => {});

    const now = ctx.currentTime + 0.02;
    const volume = getMagicEffectVolume();
    if (volume <= 0) return;

    const effectMixMap = {
      drumroll: 1.28,
      confetti: 1.1,
      micdrop: 1.34,
      curtain: 1.06,
      bubbles: 0.92,
      blur: 0.95,
      quiet: 0.78,
      applause: 1.02,
      spotlight: 0.82,
      correct: 0.96,
      wrong: 0.9,
      timesup: 1.0,
      sparkle: 0.88,
      stars: 0.92,
      hype: 1.08,
      freeze: 0.9
    };
    const effectMix = effectMixMap[effectId] || 1;

    // One compressed mix bus per effect. This lets the effects sound much bigger
    // without turning into painful clipping/distortion on laptop speakers.
    const master = ctx.createGain();
    const compressor = ctx.createDynamicsCompressor();
    try {
      compressor.threshold.setValueAtTime(-20, now);
      compressor.knee.setValueAtTime(28, now);
      compressor.ratio.setValueAtTime(6, now);
      compressor.attack.setValueAtTime(0.004, now);
      compressor.release.setValueAtTime(0.28, now);
    } catch (error) {}
    master.gain.setValueAtTime(Math.min(1.32, (0.68 + volume * 0.16) * effectMix), now);
    master.connect(compressor).connect(ctx.destination);

    const clampGain = (gainValue) => Math.max(0.0001, Math.min(0.92, gainValue * volume * Math.min(1.14, 0.94 + effectMix * 0.14)));

    const disconnectLater = (node, delay = 3.5) => {
      window.setTimeout(() => {
        try { node.disconnect(); } catch (error) {}
      }, delay * 1000);
    };

    const tone = (time, freq, duration = 0.16, type = 'sine', gainValue = 0.12, endFreq = null) => {
      try {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        const level = clampGain(gainValue);
        osc.type = type;
        osc.frequency.setValueAtTime(freq, time);
        if (endFreq) osc.frequency.exponentialRampToValueAtTime(Math.max(1, endFreq), time + Math.max(0.025, duration * 0.88));
        gain.gain.setValueAtTime(0.0001, time);
        gain.gain.exponentialRampToValueAtTime(level, time + 0.012);
        gain.gain.exponentialRampToValueAtTime(0.0001, time + duration);
        osc.connect(gain).connect(master);
        osc.start(time);
        osc.stop(time + duration + 0.04);
      } catch (error) {}
    };

    const noise = (time, duration = 0.18, gainValue = 0.12, filterType = 'bandpass', frequency = 900, q = 1.2) => {
      try {
        const length = Math.max(1, Math.floor(ctx.sampleRate * duration));
        const buffer = ctx.createBuffer(1, length, ctx.sampleRate);
        const data = buffer.getChannelData(0);
        for (let i = 0; i < length; i++) data[i] = (Math.random() * 2 - 1) * (1 - i / length * 0.35);
        const source = ctx.createBufferSource();
        const filter = ctx.createBiquadFilter();
        const gain = ctx.createGain();
        filter.type = filterType;
        filter.frequency.setValueAtTime(frequency, time);
        filter.Q.setValueAtTime(q, time);
        gain.gain.setValueAtTime(0.0001, time);
        gain.gain.exponentialRampToValueAtTime(clampGain(gainValue), time + 0.01);
        gain.gain.exponentialRampToValueAtTime(0.0001, time + duration);
        source.buffer = buffer;
        source.connect(filter).connect(gain).connect(master);
        source.start(time);
        source.stop(time + duration + 0.02);
      } catch (error) {}
    };

    const chord = (time, freqs, duration = 0.24, type = 'triangle', gainValue = 0.08) => {
      freqs.forEach((freq, index) => tone(time + index * 0.006, freq, duration, type, gainValue, null));
    };

    const boom = (time = now, strength = 1) => {
      tone(time, 145, 0.34, 'sine', 0.25 * strength, 42);
      tone(time + 0.012, 72, 0.42, 'triangle', 0.24 * strength, 32);
      noise(time + 0.01, 0.32, 0.16 * strength, 'lowpass', 210, 0.75);
    };

    const cymbal = (time = now, strength = 1) => {
      noise(time, 0.52, 0.13 * strength, 'highpass', 4200, 0.55);
      noise(time + 0.02, 0.38, 0.06 * strength, 'bandpass', 8200, 0.9);
    };

    const whoosh = (time = now, duration = 0.55, strength = 1) => {
      try {
        const length = Math.max(1, Math.floor(ctx.sampleRate * duration));
        const buffer = ctx.createBuffer(1, length, ctx.sampleRate);
        const data = buffer.getChannelData(0);
        for (let i = 0; i < length; i++) data[i] = Math.random() * 2 - 1;
        const source = ctx.createBufferSource();
        const filter = ctx.createBiquadFilter();
        const gain = ctx.createGain();
        filter.type = 'bandpass';
        filter.frequency.setValueAtTime(260, time);
        filter.frequency.exponentialRampToValueAtTime(4600, time + duration * 0.88);
        filter.Q.setValueAtTime(0.85, time);
        gain.gain.setValueAtTime(0.0001, time);
        gain.gain.exponentialRampToValueAtTime(clampGain(0.18 * strength), time + duration * 0.35);
        gain.gain.exponentialRampToValueAtTime(0.0001, time + duration);
        source.buffer = buffer;
        source.connect(filter).connect(gain).connect(master);
        source.start(time);
        source.stop(time + duration + 0.04);
      } catch (error) {}
    };

    const sparkleRun = (time = now, count = 7, base = 880) => {
      for (let i = 0; i < count; i++) tone(time + i * 0.055, base * Math.pow(2, (i % 5) / 12), 0.12, 'sine', 0.065);
    };

    const fanfare = (time = now, heroic = false) => {
      const seq = heroic
        ? [[392, 494, 587], [523, 659, 784], [659, 784, 1046], [784, 988, 1318]]
        : [[523, 659, 784], [659, 784, 1046], [784, 988, 1318]];
      seq.forEach((freqs, index) => chord(time + index * 0.16, freqs, 0.22, 'triangle', 0.075));
      tone(time + seq.length * 0.16 + 0.02, heroic ? 1568 : 1318, 0.34, 'sine', 0.11);
      cymbal(time + seq.length * 0.16, heroic ? 1.15 : 0.75);
    };

    switch (effectId) {
      case 'drumroll': {
        for (let i = 0; i < 88; i++) {
          const t = now + i * 0.029;
          const swell = 0.82 + i / 120;
          noise(t, 0.05, (i % 2 ? 0.16 : 0.13) * swell, 'bandpass', i % 2 ? 1420 : 940, 2.35);
          tone(t, i % 2 ? 152 : 112, 0.052, 'square', 0.052 * swell, i % 2 ? 122 : 96);
        }
        for (let i = 0; i < 10; i++) {
          tone(now + 2.18 + i * 0.045, 174 + i * 16, 0.085, 'triangle', 0.085, 96 + i * 8);
        }
        boom(now + 2.72, 1.18);
        cymbal(now + 2.76, 1.2);
        break;
      }
      case 'confetti':
        fanfare(now, true);
        sparkleRun(now + 0.12, 13, 1046);
        cymbal(now + 0.24, 1.0);
        break;
      case 'correct':
        chord(now, [523, 659, 784], 0.18, 'triangle', 0.1);
        chord(now + 0.16, [659, 784, 1046], 0.28, 'triangle', 0.11);
        tone(now + 0.34, 1568, 0.25, 'sine', 0.09);
        cymbal(now + 0.22, 0.6);
        break;
      case 'sparkle':
        sparkleRun(now, 12, 988);
        noise(now + 0.12, 0.52, 0.05, 'highpass', 5600, 0.65);
        break;
      case 'stars':
        sparkleRun(now, 15, 880);
        chord(now + 0.16, [659, 784, 988], 0.24, 'triangle', 0.06);
        noise(now + 0.12, 0.5, 0.045, 'highpass', 5200, 0.55);
        break;
      case 'hype':
        whoosh(now, 0.48, 1.08);
        boom(now + 0.3, 1.15);
        fanfare(now + 0.34, true);
        cymbal(now + 0.56, 0.9);
        break;
      case 'applause':
        for (let i = 0; i < 34; i++) noise(now + i * 0.034 + Math.random() * 0.022, 0.045, 0.12, 'bandpass', 1400 + Math.random() * 2600, 0.9);
        sparkleRun(now + 0.18, 8, 988);
        break;
      case 'wrong':
        tone(now, 190, 0.26, 'sawtooth', 0.21, 128);
        tone(now + 0.18, 142, 0.36, 'sawtooth', 0.23, 88);
        noise(now, 0.44, 0.08, 'lowpass', 500, 0.8);
        break;
      case 'timesup':
        for (let i = 0; i < 5; i++) tone(now + i * 0.18, 1046, 0.08, 'square', 0.09);
        tone(now + 0.12, 1318, 0.08, 'square', 0.08);
        tone(now + 0.3, 1318, 0.08, 'square', 0.08);
        tone(now + 0.48, 1318, 0.08, 'square', 0.08);
        boom(now + 0.88, 0.8);
        cymbal(now + 0.92, 0.75);
        break;
      case 'quiet':
        whoosh(now, 0.42, 0.42);
        noise(now + 0.04, 0.78, 0.09, 'highpass', 2600, 0.46);
        if ('speechSynthesis' in window) {
          try {
            window.speechSynthesis.cancel();
            const utter = new SpeechSynthesisUtterance('Be Quiet');
            const voices = window.speechSynthesis.getVoices ? window.speechSynthesis.getVoices() : [];
            const preferred = voices.find(v => /en|us|uk|ph/i.test((v.lang || '') + ' ' + (v.name || ''))) || voices[0];
            if (preferred) utter.voice = preferred;
            utter.rate = 0.72;
            utter.pitch = 0.8;
            utter.volume = Math.min(1, 0.95 * volume);
            window.speechSynthesis.speak(utter);
          } catch (error) {}
        }
        break;
      case 'micdrop':
        whoosh(now, 0.54, 0.95);
        tone(now + 0.08, 520, 0.16, 'triangle', 0.10, 220);
        tone(now + 0.22, 300, 0.18, 'sawtooth', 0.10, 124);
        noise(now + 0.52, 0.22, 0.06, 'bandpass', 1100, 0.7);
        boom(now + 0.98, 1.62);
        tone(now + 1.0, 92, 0.42, 'sine', 0.26, 36);
        noise(now + 1.02, 0.54, 0.20, 'lowpass', 185, 0.65);
        noise(now + 1.12, 0.36, 0.10, 'bandpass', 720, 0.55);
        tone(now + 1.22, 148, 0.18, 'square', 0.055, 82);
        break;
      case 'curtain':
        whoosh(now, 0.95, 1.12);
        chord(now + 0.18, [196, 247, 330], 0.42, 'triangle', 0.08);
        noise(now + 0.2, 0.72, 0.08, 'bandpass', 620, 0.7);
        chord(now + 0.72, [392, 494, 659], 0.34, 'triangle', 0.07);
        break;
      case 'bubbles':
        [698, 784, 880, 988, 1174, 1318, 1568, 1760, 1975].forEach((freq, i) => {
          const t = now + i * 0.21;
          tone(t, freq, 0.16, 'sine', 0.052, freq * 1.1);
          noise(t + 0.035, 0.12, 0.026, 'highpass', 4700 + i * 210, 0.46);
        });
        noise(now + 0.44, 1.55, 0.028, 'highpass', 5600, 0.36);
        break;
      case 'blur':
        whoosh(now, 0.62, 0.78);
        noise(now + 0.08, 0.72, 0.07, 'lowpass', 420, 0.5);
        tone(now + 0.1, 220, 0.42, 'triangle', 0.08, 118);
        tone(now + 0.34, 146, 0.34, 'sine', 0.05, 96);
        break;
      case 'spotlight':
        whoosh(now, 0.44, 0.78);
        tone(now + 0.18, 520, 0.18, 'triangle', 0.1, 760);
        tone(now + 0.34, 1046, 0.24, 'sine', 0.08);
        break;
      case 'freeze':
        [1174, 988, 880, 740, 622].forEach((freq, i) => tone(now + i * 0.06, freq, 0.16, 'sine', 0.065));
        noise(now + 0.16, 0.42, 0.06, 'highpass', 6200, 0.78);
        tone(now + 0.42, 330, 0.3, 'triangle', 0.08, 210);
        break;
      default:
        fanfare(now, false);
        break;
    }

    const disconnectDelayMap = { drumroll: 4.4, bubbles: 4.4, curtain: 3.8, confetti: 3.6, spotlight: 3.6, micdrop: 3.6, applause: 3.6, hype: 3.6, blur: 3.4 };
    disconnectLater(master, disconnectDelayMap[effectId] || 2.9);
  }

  function handleKeyboard(event) {
    if (!els.viewerView || els.viewerView.classList.contains('hidden')) return;

    const key = event.key;
    const magicEffect = getMagicEffectByShortcut(key);
    const isPresentationKey = key === 'ArrowRight' || key === 'PageDown' || key === ' ' || key === 'ArrowLeft' || key === 'PageUp' || key.toLowerCase() === 'f' || key === 'Escape' || !!magicEffect;
    if (!isPresentationKey) return;

    if (magicEffect && !shouldIgnoreKeyboardForTyping()) {
      event.preventDefault();
      event.stopPropagation();
      triggerMagicEffect(magicEffect.id);
      return;
    }

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
    return [
      'slide-transition',
      'transition-fade',
      'transition-slide-left',
      'transition-slide-right',
      'transition-slide-up',
      'transition-slide-down',
      'transition-zoom-in',
      'transition-zoom-out',
      'transition-soft-blur',
      'transition-bounce-pop',
      'transition-happy-pop',
      'transition-elastic-pop',
      'transition-sparkle-zoom',
      'transition-page-flip',
      'transition-paper-turn',
      'transition-circle-reveal',
      'transition-corner-reveal',
      'transition-bright-wipe',
      'transition-ribbon-wipe',
      'transition-split-open',
      'transition-drop-in',
      'transition-float-up',
      'transition-diagonal-rise',
      'transition-diagonal-fall',
      'transition-smooth-swing',
      'transition-push-left',
      'transition-push-right',
      'transition-carousel'
    ];
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
    target.style.setProperty('--ph-transition-ms', `${state.transitionDuration}ms`);
    target.classList.add('slide-transition', `transition-${effect}`);
    window.setTimeout(() => {
      target.classList.remove(...classes);
      target.style.removeProperty('--ph-transition-ms');
    }, state.transitionDuration + 180);
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
    if (!state.firebaseReady || (!state.activeFile && !state.mediaMode)) return;
    if (state.unsubscribeSession) state.unsubscribeSession();
    state.sessionId = state.sessionId || makeSessionId();
    state.sessionRef = state.firebaseDb.collection(SESSION_COLLECTION).doc(state.sessionId);
    state.lastCommandId = null;
    if (state.remoteCommandPollTimer) clearInterval(state.remoteCommandPollTimer);
    state.processedRemoteCommands = new Set();
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
      processRemoteCommand(data && data.command, 'snapshot');
      processRemoteCommand(data && data.lastMagicCommand, 'snapshot-magic');
      processMediaCastSignal(data && data.mediaCast);
    }, (error) => {
      console.warn('Remote live listener failed; command polling fallback remains active.', error);
    });
    startRemoteCommandFallbackPoll();
  }

  function processRemoteCommand(command, source = 'unknown') {
    if (!command || !command.id) return;
    if (state.processedRemoteCommands && state.processedRemoteCommands.has(command.id)) return;
    state.lastCommandId = command.id;
    if (!state.processedRemoteCommands) state.processedRemoteCommands = new Set();
    state.processedRemoteCommands.add(command.id);
    // Keep the Set small so the app does not grow memory during long sessions.
    if (state.processedRemoteCommands.size > 80) {
      state.processedRemoteCommands = new Set(Array.from(state.processedRemoteCommands).slice(-30));
    }
    applyRemoteCommand(command);
  }

  function startRemoteCommandFallbackPoll() {
    if (!state.sessionRef) return;
    if (state.remoteCommandPollTimer) clearInterval(state.remoteCommandPollTimer);
    // Firebase onSnapshot should handle commands instantly. This fallback catches
    // cases where the realtime listener silently stops, the page resumes from
    // sleep, or the browser throttles the tab. It is intentionally light.
    state.remoteCommandPollTimer = setInterval(async () => {
      try {
        if (!state.sessionRef) return;
        const snap = await state.sessionRef.get();
        if (!snap.exists) return;
        const data = snap.data();
        processRemoteCommand(data && data.command, 'poll');
        processRemoteCommand(data && data.lastMagicCommand, 'poll-magic');
        processMediaCastSignal(data && data.mediaCast);
      } catch (error) {}
    }, 700);
  }


  function mediaCastSupported() {
    return typeof RTCPeerConnection !== 'undefined';
  }

  function getMediaKindFromType(type = '', name = '') {
    const mime = String(type || '').toLowerCase();
    const fileName = String(name || '').toLowerCase();
    if (mime.startsWith('image/') || /\.(png|jpe?g|webp|gif|bmp|svg)$/i.test(fileName)) return 'image';
    if (mime.startsWith('audio/') || /\.(mp3|wav|m4a|aac|ogg)$/i.test(fileName)) return 'audio';
    if (mime.startsWith('video/') || /\.(mp4|webm|mov|m4v)$/i.test(fileName)) return 'video';
    return 'file';
  }

  function formatMediaBytes(bytes = 0) {
    const n = Number(bytes) || 0;
    if (n < 1024) return `${Math.round(n)} B`;
    if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
    if (n < 1024 * 1024 * 1024) return `${(n / 1024 / 1024).toFixed(1)} MB`;
    return `${(n / 1024 / 1024 / 1024).toFixed(1)} GB`;
  }

  function cleanMediaMeta(raw = {}) {
    const name = String(raw.name || 'Media file').slice(0, 140);
    const type = String(raw.type || '').slice(0, 120);
    const size = Math.max(0, Number(raw.size) || 0);
    return { name, type, size, kind: getMediaKindFromType(type, name) };
  }

  function waitForIceGatheringComplete(pc, timeout = 5200) {
    return new Promise((resolve) => {
      if (!pc || pc.iceGatheringState === 'complete') return resolve();
      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        try { pc.removeEventListener('icegatheringstatechange', onChange); } catch (error) {}
        resolve();
      };
      const onChange = () => {
        if (pc.iceGatheringState === 'complete') finish();
      };
      pc.addEventListener('icegatheringstatechange', onChange);
      setTimeout(finish, timeout);
    });
  }

  async function updateMediaCastSignal(patch = {}) {
    if (!state.sessionRef) return;
    const dotted = {};
    Object.keys(patch).forEach((key) => { dotted[`mediaCast.${key}`] = patch[key]; });
    try {
      await state.sessionRef.update(dotted);
    } catch (error) {
      try { await state.sessionRef.set({ mediaCast: patch }, { merge: true }); } catch (inner) {}
    }
  }

  function processMediaCastSignal(signal) {
    if (isRemoteMode || !signal || !signal.id) return;
    if (signal.status === 'stop') {
      if (state.mediaCastReceiverId === signal.id) closeMediaCastOverlay(false);
      return;
    }
    if (!signal.offer || state.mediaCastReceiverId === signal.id) return;
    if (signal.createdAt && Date.now() - Number(signal.createdAt) > 15 * 60 * 1000) return;
    startMediaCastReceiver(signal).catch((error) => {
      console.warn('Media Cast receiver failed:', error);
      updateMediaCastSignal({ receiverStatus: 'error', receiverMessage: 'Desktop could not receive media.', receiverUpdatedAt: Date.now() });
    });
  }

  async function startMediaCastReceiver(signal) {
    if (!mediaCastSupported() || !state.sessionRef) return;
    if (state.mediaCastReceiver && state.mediaCastReceiver.pc) {
      try { state.mediaCastReceiver.pc.close(); } catch (error) {}
    }
    state.mediaCastReceiverId = signal.id;
    const meta = cleanMediaMeta(signal.meta || {});
    const transfer = { id: signal.id, meta, chunks: [], received: 0, pc: null, startedAt: Date.now() };
    state.mediaCastReceiver = transfer;
    showMediaCastReceiving(meta, 0);
    await updateMediaCastSignal({ receiverStatus: 'connecting', receiverMessage: 'Desktop preparing media cast...', receiverUpdatedAt: Date.now() });

    const pc = new RTCPeerConnection(MEDIA_CAST_ICE_SERVERS);
    transfer.pc = pc;
    pc.ondatachannel = (event) => {
      const channel = event.channel;
      channel.binaryType = 'arraybuffer';
      channel.onopen = () => updateMediaCastSignal({ receiverStatus: 'receiving', receiverMessage: 'Receiving media...', receiverUpdatedAt: Date.now() });
      channel.onerror = () => updateMediaCastSignal({ receiverStatus: 'error', receiverMessage: 'Media channel error.', receiverUpdatedAt: Date.now() });
      channel.onmessage = async (event) => {
        await handleMediaCastReceiverMessage(transfer, event.data);
      };
    };
    pc.onconnectionstatechange = () => {
      const status = pc.connectionState;
      if (status === 'failed' || status === 'disconnected') {
        updateMediaCastSignal({ receiverStatus: status, receiverMessage: 'Media cast connection was interrupted.', receiverUpdatedAt: Date.now() });
      }
    };
    await pc.setRemoteDescription(new RTCSessionDescription(signal.offer));
    const answer = await pc.createAnswer();
    await pc.setLocalDescription(answer);
    await waitForIceGatheringComplete(pc);
    await updateMediaCastSignal({ answer: pc.localDescription.toJSON ? pc.localDescription.toJSON() : { type: pc.localDescription.type, sdp: pc.localDescription.sdp }, receiverStatus: 'ready', receiverMessage: 'Desktop ready. Sending file...', receiverUpdatedAt: Date.now() });
  }

  async function handleMediaCastReceiverMessage(transfer, data) {
    if (!transfer || transfer.id !== state.mediaCastReceiverId) return;
    if (typeof data === 'string') {
      if (data.startsWith('PH_MEDIA_META:')) {
        try { transfer.meta = cleanMediaMeta(JSON.parse(data.slice('PH_MEDIA_META:'.length))); } catch (error) {}
        showMediaCastReceiving(transfer.meta, 0);
        return;
      }
      if (data === 'PH_MEDIA_DONE') {
        await finishMediaCastReceive(transfer);
        return;
      }
      if (data.startsWith('PH_MEDIA_ABORT')) {
        showMediaCastStatus('Media cast stopped by phone.', 'warning');
        await updateMediaCastSignal({ receiverStatus: 'stopped', receiverMessage: 'Media cast stopped.', receiverUpdatedAt: Date.now() });
      }
      return;
    }
    let chunk = data;
    if (data instanceof Blob) chunk = await data.arrayBuffer();
    if (!chunk || !chunk.byteLength) return;
    transfer.chunks.push(chunk);
    transfer.received += chunk.byteLength;
    showMediaCastReceiving(transfer.meta, transfer.received);
    if (!transfer.lastStatusAt || Date.now() - transfer.lastStatusAt > 700) {
      transfer.lastStatusAt = Date.now();
      await updateMediaCastSignal({ receiverStatus: 'receiving', receiverReceived: transfer.received, receiverMessage: `Received ${formatMediaBytes(transfer.received)} of ${formatMediaBytes(transfer.meta.size)}`, receiverUpdatedAt: Date.now() });
    }
  }

  async function finishMediaCastReceive(transfer) {
    const meta = cleanMediaMeta(transfer.meta || {});
    if (state.mediaCastBlobUrl) {
      try { URL.revokeObjectURL(state.mediaCastBlobUrl); } catch (error) {}
    }
    const blob = new Blob(transfer.chunks, { type: meta.type || 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    state.mediaCastBlobUrl = url;
    showMediaCastPlayer(url, meta);
    await updateMediaCastSignal({ receiverStatus: 'complete', receiverReceived: blob.size, receiverMessage: 'Media ready on desktop.', receiverUpdatedAt: Date.now() });
    try { if (transfer.pc) transfer.pc.close(); } catch (error) {}
  }

  function ensureMediaCastLayer() {
    let layer = document.getElementById('mediaCastLayer');
    const host = document.fullscreenElement || document.body;
    if (!layer) {
      layer = document.createElement('div');
      layer.id = 'mediaCastLayer';
      layer.className = 'media-cast-layer hidden';
    }
    if (layer.parentElement !== host) host.appendChild(layer);
    return layer;
  }

  function showMediaCastReceiving(meta = {}, received = 0) {
    const layer = ensureMediaCastLayer();
    const total = Number(meta.size) || 0;
    const pct = total ? Math.max(0, Math.min(100, Math.round((received / total) * 100))) : 0;
    layer.className = 'media-cast-layer show receiving';
    layer.innerHTML = `
      <div class="media-cast-card receiving-card">
        <div class="media-cast-spinner"></div>
        <strong>Receiving media from phone...</strong>
        <span>${escapeHtml(meta.name || 'Media file')}</span>
        <div class="media-cast-progress"><i style="width:${pct}%"></i></div>
        <small>${formatMediaBytes(received)} / ${formatMediaBytes(total)}</small>
      </div>
    `;
  }

  function showMediaCastStatus(message, tone = '') {
    const layer = ensureMediaCastLayer();
    layer.className = `media-cast-layer show ${tone}`;
    layer.innerHTML = `<div class="media-cast-card"><strong>${escapeHtml(message || 'Media Cast')}</strong></div>`;
  }

  function showMediaCastPlayer(url, meta = {}) {
    const safe = cleanMediaMeta(meta);
    const layer = ensureMediaCastLayer();
    const title = escapeHtml(safe.name || 'Media');
    let mediaHtml = '';
    if (safe.kind === 'image') {
      mediaHtml = `<img class="media-cast-image" src="${url}" alt="${title}">`;
    } else if (safe.kind === 'video') {
      mediaHtml = `<video id="mediaCastPlayer" class="media-cast-player" src="${url}" controls playsinline preload="metadata"></video>`;
    } else if (safe.kind === 'audio') {
      mediaHtml = `<div class="media-cast-audio-card"><div class="media-cast-audio-icon">♪</div><strong>${title}</strong><audio id="mediaCastPlayer" class="media-cast-player" src="${url}" controls preload="metadata"></audio></div>`;
    } else {
      mediaHtml = `<div class="media-cast-audio-card"><div class="media-cast-audio-icon">FILE</div><strong>${title}</strong><small>This file type cannot be previewed.</small></div>`;
    }
    layer.className = `media-cast-layer show media-kind-${safe.kind}`;
    layer.innerHTML = `
      <div class="media-cast-frame">
        <button id="mediaCastClose" type="button" class="media-cast-close">×</button>
        <div class="media-cast-title">${title}</div>
        <div class="media-cast-media-wrap">${mediaHtml}</div>
      </div>
    `;
    const close = document.getElementById('mediaCastClose');
    if (close) close.addEventListener('click', () => closeMediaCastOverlay(true));
    const player = document.getElementById('mediaCastPlayer');
    if (player) {
      player.volume = 0.92;
      player.addEventListener('play', () => layer.classList.add('is-playing'));
      player.addEventListener('pause', () => layer.classList.remove('is-playing'));
    }
  }

  function closeMediaCastOverlay(publishStop = false) {
    const layer = document.getElementById('mediaCastLayer');
    if (layer) {
      layer.classList.add('hidden');
      layer.classList.remove('show');
      layer.innerHTML = '';
    }
    if (state.mediaCastBlobUrl) {
      try { URL.revokeObjectURL(state.mediaCastBlobUrl); } catch (error) {}
      state.mediaCastBlobUrl = '';
    }
    if (state.mediaCastReceiver && state.mediaCastReceiver.pc) {
      try { state.mediaCastReceiver.pc.close(); } catch (error) {}
    }
    state.mediaCastReceiver = null;
    if (publishStop) updateMediaCastSignal({ receiverStatus: 'closed', receiverMessage: 'Media closed on desktop.', receiverUpdatedAt: Date.now() });
  }

  function showMediaCastGesturePrompt(player) {
    const layer = ensureMediaCastLayer();
    let prompt = document.getElementById('mediaCastGesturePrompt');
    if (!prompt) {
      prompt = document.createElement('div');
      prompt.id = 'mediaCastGesturePrompt';
      prompt.className = 'media-cast-gesture';
      prompt.innerHTML = '<strong>Tap desktop once to play sound/video</strong><button type="button">Play now</button>';
      layer.appendChild(prompt);
    }
    const button = prompt.querySelector('button');
    if (button) button.onclick = () => {
      try { player.play().then(() => prompt.remove()).catch(() => {}); } catch (error) {}
    };
  }


  function normalizeOnlineMediaUrl(value = '') {
    const raw = String(value || '').trim();
    if (!raw) return null;
    try {
      const url = new URL(raw);
      if (!/^https?:$/i.test(url.protocol)) return null;
      return url;
    } catch (error) {
      return null;
    }
  }

  function extractYouTubeId(url) {
    if (!url) return '';
    const host = url.hostname.replace(/^www\./i, '').toLowerCase();
    if (host === 'youtu.be') return (url.pathname.split('/').filter(Boolean)[0] || '').slice(0, 24);
    if (host.endsWith('youtube.com') || host.endsWith('youtube-nocookie.com')) {
      if (url.searchParams.get('v')) return url.searchParams.get('v').slice(0, 24);
      const parts = url.pathname.split('/').filter(Boolean);
      const embedIndex = parts.findIndex((part) => ['embed', 'shorts', 'live'].includes(part));
      if (embedIndex >= 0 && parts[embedIndex + 1]) return parts[embedIndex + 1].slice(0, 24);
    }
    return '';
  }

  function extractTikTokVideoId(url) {
    if (!url) return '';
    const parts = url.pathname.split('/').filter(Boolean);
    const videoIndex = parts.indexOf('video');
    if (videoIndex >= 0 && parts[videoIndex + 1]) return parts[videoIndex + 1].replace(/[^0-9]/g, '').slice(0, 32);
    return '';
  }

  function detectOnlineMediaKind(url) {
    const host = url.hostname.replace(/^www\./i, '').toLowerCase();
    const path = url.pathname.toLowerCase();
    if (host === 'youtu.be' || host.endsWith('youtube.com') || host.endsWith('youtube-nocookie.com')) return 'youtube';
    if (host.endsWith('tiktok.com') || host === 'vm.tiktok.com' || host === 'vt.tiktok.com') return 'tiktok';
    if (/\.(png|jpe?g|webp|gif|bmp|svg)(\?.*)?$/i.test(`${url.pathname}${url.search}`)) return 'image';
    if (/\.(mp3|wav|m4a|aac|ogg)(\?.*)?$/i.test(`${url.pathname}${url.search}`)) return 'audio';
    if (/\.(mp4|webm|mov|m4v)(\?.*)?$/i.test(`${url.pathname}${url.search}`)) return 'video';
    return 'web';
  }

  function postYouTubeCommand(func, args = []) {
    const frame = document.getElementById('mediaCastLinkFrame');
    if (!frame || !frame.contentWindow || frame.dataset.provider !== 'youtube') return;
    try {
      frame.contentWindow.postMessage(JSON.stringify({ event: 'command', func, args }), '*');
    } catch (error) {}
  }

  function showOnlineMediaCast(payload = {}) {
    const parsed = normalizeOnlineMediaUrl(payload.url || payload.link || '');
    if (!parsed) { showMediaCastStatus('Invalid media link.', 'warning'); return; }
    const kind = detectOnlineMediaKind(parsed);
    const rawVolume = Number(payload.volume);
    const volume = Math.max(0, Math.min(1, Number.isFinite(rawVolume) ? rawVolume : (state.mediaLinkVolume || 0.9)));
    state.mediaLinkVolume = volume;
    const layer = ensureMediaCastLayer();
    const rawUrl = parsed.toString();
    const safeUrl = escapeHtml(rawUrl);
    const safeTitle = escapeHtml(kind === 'youtube' ? 'YouTube Link' : kind === 'tiktok' ? 'TikTok Link' : 'Online Media Link');
    let content = '';

    if (kind === 'youtube') {
      const id = extractYouTubeId(parsed);
      if (!id) { showMediaCastStatus('Could not read the YouTube video ID.', 'warning'); return; }
      const embed = `https://www.youtube.com/embed/${encodeURIComponent(id)}?enablejsapi=1&autoplay=1&playsinline=1&rel=0&modestbranding=1`;
      content = `<iframe id="mediaCastLinkFrame" class="media-cast-link-frame" data-provider="youtube" src="${escapeHtml(embed)}" title="YouTube video" allow="autoplay; encrypted-media; picture-in-picture; fullscreen" allowfullscreen></iframe>`;
    } else if (kind === 'tiktok') {
      const id = extractTikTokVideoId(parsed);
      const embed = id ? `https://www.tiktok.com/embed/v2/${encodeURIComponent(id)}` : rawUrl;
      content = `<iframe id="mediaCastLinkFrame" class="media-cast-link-frame tiktok" data-provider="tiktok" src="${escapeHtml(embed)}" title="TikTok video" allow="autoplay; encrypted-media; fullscreen; picture-in-picture" allowfullscreen></iframe>`;
    } else if (kind === 'image') {
      content = `<img class="media-cast-image" src="${safeUrl}" alt="Online image">`;
    } else if (kind === 'video') {
      content = `<video id="mediaCastPlayer" class="media-cast-player" src="${safeUrl}" controls playsinline autoplay preload="auto"></video>`;
    } else if (kind === 'audio') {
      content = `<div class="media-cast-audio-card"><div class="media-cast-audio-icon">♪</div><strong>Online Audio</strong><audio id="mediaCastPlayer" class="media-cast-player" src="${safeUrl}" controls autoplay preload="auto"></audio></div>`;
    } else {
      content = `<iframe id="mediaCastLinkFrame" class="media-cast-link-frame" data-provider="web" src="${safeUrl}" title="Online media page" allow="autoplay; encrypted-media; fullscreen; picture-in-picture" allowfullscreen></iframe>`;
    }

    layer.className = `media-cast-layer show media-kind-${kind} media-link-mode`;
    layer.innerHTML = `
      <div class="media-cast-frame media-cast-link-frame-wrap">
        <button id="mediaCastClose" type="button" class="media-cast-close">×</button>
        <div class="media-cast-title">${safeTitle}</div>
        <div class="media-cast-media-wrap">${content}</div>
        <div class="media-cast-link-note">${kind === 'tiktok' ? 'TikTok autoplay/volume depends on the embedded player permission.' : 'Remote controls are synced to this desktop screen.'}</div>
      </div>
    `;
    const close = document.getElementById('mediaCastClose');
    if (close) close.addEventListener('click', () => closeMediaCastOverlay(true));
    const player = document.getElementById('mediaCastPlayer');
    if (player) {
      try { player.volume = volume; player.muted = volume <= 0; } catch (error) {}
      if (payload.autoplay !== false) {
        try { player.play().catch(() => showMediaCastGesturePrompt(player)); } catch (error) { showMediaCastGesturePrompt(player); }
      }
    }
    if (kind === 'youtube') {
      window.setTimeout(() => {
        postYouTubeCommand('setVolume', [Math.round(volume * 100)]);
        if (payload.autoplay !== false) postYouTubeCommand('playVideo');
      }, 1400);
    }
  }

  function controlMediaCast(raw = {}) {
    const action = typeof raw === 'string' ? raw : String(raw.type || raw.action || '');
    const layer = document.getElementById('mediaCastLayer');
    const player = document.getElementById('mediaCastPlayer');
    const frame = document.getElementById('mediaCastLinkFrame');
    const provider = frame ? String(frame.dataset.provider || '') : '';
    if (action === 'stop') { closeMediaCastOverlay(false); return; }
    if (action === 'hide') { if (layer) layer.classList.add('media-cast-minimized'); return; }
    if (action === 'show') { if (layer) layer.classList.remove('media-cast-minimized', 'hidden'); return; }
    if (provider === 'youtube') {
      if (action === 'play') { postYouTubeCommand('playVideo'); return; }
      if (action === 'pause') { postYouTubeCommand('pauseVideo'); return; }
      if (action === 'toggle') { postYouTubeCommand('playVideo'); return; }
      if (action === 'volume') {
        const value = Math.max(0, Math.min(1, Number(raw.value) || 0));
        state.mediaLinkVolume = value;
        postYouTubeCommand('setVolume', [Math.round(value * 100)]);
        if (value <= 0) postYouTubeCommand('mute'); else postYouTubeCommand('unMute');
        return;
      }
    }
    if (!player) return;
    if (action === 'play') {
      try { player.play().catch(() => showMediaCastGesturePrompt(player)); } catch (error) { showMediaCastGesturePrompt(player); }
      return;
    }
    if (action === 'pause') { try { player.pause(); } catch (error) {} return; }
    if (action === 'toggle') {
      try { if (player.paused) player.play().catch(() => showMediaCastGesturePrompt(player)); else player.pause(); } catch (error) {}
      return;
    }
    if (action === 'volume') {
      const value = Math.max(0, Math.min(1, Number(raw.value) || 0));
      state.mediaLinkVolume = value;
      try { player.volume = value; player.muted = value <= 0; } catch (error) {}
    }
  }

  function setRemoteMediaStatus(message, tone = '') {
    const node = $('remoteMediaStatus');
    if (!node) return;
    node.textContent = message || 'Ready to cast media.';
    node.classList.remove('ok', 'warn', 'error', 'busy');
    if (tone) node.classList.add(tone);
  }

  function syncRemoteMediaStatus(signal = {}) {
    const current = window.__phRemoteMediaCast || null;
    if (!signal || !signal.id) return;
    if (current && current.id && signal.id !== current.id) return;
    const status = signal.receiverStatus || signal.status || '';
    if (status === 'ready') setRemoteMediaStatus('Desktop ready. Sending media...', 'busy');
    else if (status === 'receiving') setRemoteMediaStatus(signal.receiverMessage || 'Desktop is receiving media...', 'busy');
    else if (status === 'complete') setRemoteMediaStatus('Media is now showing on desktop.', 'ok');
    else if (status === 'error' || status === 'failed' || status === 'disconnected') setRemoteMediaStatus(signal.receiverMessage || 'Media Cast failed. Try same WiFi or a smaller file.', 'error');
    else if (status === 'connecting') setRemoteMediaStatus('Desktop is preparing receiver...', 'busy');
  }

  function waitForDataChannelOpen(channel, timeout = 30000) {
    return new Promise((resolve, reject) => {
      if (!channel) return reject(new Error('No media data channel.'));
      if (channel.readyState === 'open') return resolve();
      const timer = setTimeout(() => reject(new Error('Timed out waiting for media channel.')), timeout);
      channel.onopen = () => { clearTimeout(timer); resolve(); };
      channel.onerror = () => { clearTimeout(timer); reject(new Error('Media data channel failed.')); };
    });
  }

  function waitForBufferedAmountLow(channel) {
    return new Promise((resolve) => {
      if (!channel || channel.bufferedAmount < 2 * 1024 * 1024) return resolve();
      const prev = channel.onbufferedamountlow;
      channel.bufferedAmountLowThreshold = 512 * 1024;
      channel.onbufferedamountlow = (...args) => {
        channel.onbufferedamountlow = prev || null;
        resolve(...args);
      };
      setTimeout(resolve, 900);
    });
  }

  async function sendFileOverDataChannel(channel, file) {
    const meta = cleanMediaMeta({ name: file.name, type: file.type, size: file.size });
    channel.send(`PH_MEDIA_META:${JSON.stringify(meta)}`);
    let offset = 0;
    const statusEvery = 280;
    let lastStatus = 0;
    while (offset < file.size) {
      if (channel.readyState !== 'open') throw new Error('Media channel closed.');
      await waitForBufferedAmountLow(channel);
      const chunk = await file.slice(offset, offset + MEDIA_CAST_CHUNK_SIZE).arrayBuffer();
      channel.send(chunk);
      offset += chunk.byteLength;
      const now = Date.now();
      if (now - lastStatus > statusEvery || offset >= file.size) {
        lastStatus = now;
        const pct = file.size ? Math.round((offset / file.size) * 100) : 100;
        setRemoteMediaStatus(`Sending ${pct}% • ${formatMediaBytes(offset)} / ${formatMediaBytes(file.size)}`, 'busy');
      }
    }
    channel.send('PH_MEDIA_DONE');
  }

  async function startRemoteMediaCast(ref, file) {
    if (!ref || !file) return;
    if (!mediaCastSupported()) {
      setRemoteMediaStatus('This browser does not support Media Cast.', 'error');
      return;
    }
    if (!/^(image|audio|video)\//i.test(file.type || '') && !/\.(png|jpe?g|webp|gif|mp3|wav|m4a|aac|ogg|mp4|webm|mov|m4v)$/i.test(file.name || '')) {
      setRemoteMediaStatus('Please choose a picture, audio, or video file.', 'warn');
      return;
    }
    if (file.size > MEDIA_CAST_MAX_FILE_BYTES) {
      setRemoteMediaStatus(`File is too large for beta cast (${formatMediaBytes(file.size)}). Try below ${formatMediaBytes(MEDIA_CAST_MAX_FILE_BYTES)}.`, 'warn');
      return;
    }
    if (window.__phRemoteMediaCast && window.__phRemoteMediaCast.pc) {
      try { window.__phRemoteMediaCast.pc.close(); } catch (error) {}
      try { if (window.__phRemoteMediaCast.unsubscribe) window.__phRemoteMediaCast.unsubscribe(); } catch (error) {}
    }
    const id = `media-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const meta = cleanMediaMeta({ name: file.name, type: file.type, size: file.size });
    setRemoteMediaStatus(`Preparing ${meta.kind} • ${formatMediaBytes(file.size)}`, 'busy');
    const pc = new RTCPeerConnection(MEDIA_CAST_ICE_SERVERS);
    const channel = pc.createDataChannel('presentationHubMediaCast', { ordered: true });
    channel.binaryType = 'arraybuffer';
    channel.bufferedAmountLowThreshold = 512 * 1024;
    let answerApplied = false;
    const unsubscribe = ref.onSnapshot(async (snap) => {
      try {
        const data = snap.exists ? snap.data() : {};
        const signal = data && data.mediaCast;
        if (!signal || signal.id !== id) return;
        syncRemoteMediaStatus(signal);
        if (signal.answer && !answerApplied) {
          answerApplied = true;
          await pc.setRemoteDescription(new RTCSessionDescription(signal.answer));
        }
      } catch (error) {
        console.warn('Could not apply Media Cast answer:', error);
      }
    });
    window.__phRemoteMediaCast = { id, pc, channel, unsubscribe };
    try {
      const offer = await pc.createOffer();
      await pc.setLocalDescription(offer);
      await waitForIceGatheringComplete(pc);
      await ref.set({ mediaCast: { id, status: 'offer', meta, offer: pc.localDescription.toJSON ? pc.localDescription.toJSON() : { type: pc.localDescription.type, sdp: pc.localDescription.sdp }, createdAt: Date.now(), receiverStatus: 'waiting', receiverMessage: 'Waiting for desktop receiver...' } }, { merge: true });
      setRemoteMediaStatus('Connecting to desktop...', 'busy');
      await waitForDataChannelOpen(channel, 36000);
      await sendFileOverDataChannel(channel, file);
      await wait(450);
      setRemoteMediaStatus('Media sent. It should appear on desktop now.', 'ok');
      try { channel.close(); } catch (error) {}
      try { pc.close(); } catch (error) {}
      try { unsubscribe(); } catch (error) {}
    } catch (error) {
      console.warn('Media Cast send failed:', error);
      setRemoteMediaStatus('Media Cast failed. Try same WiFi, smaller file, or reconnect QR.', 'error');
      try { await ref.set({ mediaCast: { id, status: 'error', receiverStatus: 'error', receiverMessage: 'Phone could not send media.' } }, { merge: true }); } catch (inner) {}
      try { unsubscribe(); } catch (inner) {}
      try { pc.close(); } catch (inner) {}
    }
  }

  function attachRemoteMediaCastControls(ref, isHost) {
    const fileInput = $('remoteMediaFile');
    const sendBtn = $('remoteMediaSend');
    const fileName = $('remoteMediaFileName');
    if (!fileInput || !sendBtn) return;
    let selectedFile = null;
    fileInput.addEventListener('change', () => {
      selectedFile = fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
      if (fileName) fileName.textContent = selectedFile ? `${selectedFile.name} • ${formatMediaBytes(selectedFile.size)}` : 'No media selected';
      sendBtn.disabled = !selectedFile || !isHost;
      if (selectedFile) setRemoteMediaStatus('Ready to cast. Tap Cast to Screen.', 'ok');
    });
    sendBtn.addEventListener('click', () => {
      if (!isHost || !selectedFile) return;
      startRemoteMediaCast(ref, selectedFile);
    });
    const linkInput = $('remoteMediaLink');
    const linkBtn = $('remoteMediaLinkCast');
    if (linkBtn && linkInput) {
      linkBtn.addEventListener('click', () => {
        if (!isHost) return;
        const url = String(linkInput.value || '').trim();
        if (!url) { setRemoteMediaStatus('Paste a YouTube, TikTok, or media link first.', 'warn'); return; }
        let parsed = null;
        try { parsed = new URL(url); } catch (error) {}
        if (!parsed || !/^https?:$/i.test(parsed.protocol)) { setRemoteMediaStatus('Use a valid http/https media link.', 'warn'); return; }
        const volumeNode = $('remoteMediaVolume');
        const volumeValue = volumeNode ? Math.max(0, Math.min(1, Number(volumeNode.value) / 100)) : 0.9;
        sendRemoteCommand(ref, 'mediaLinkCast', { url, volume: volumeValue, autoplay: true });
        setRemoteMediaStatus('Link sent. Desktop should open and play it now.', 'busy');
      });
    }
    const command = (type, value) => {
      if (!isHost) return;
      sendRemoteCommand(ref, 'mediaCastControl', { type, value });
      if (type === 'stop') {
        try { ref.set({ mediaCast: { id: window.__phRemoteMediaCast ? window.__phRemoteMediaCast.id : `media-${Date.now()}`, status: 'stop', receiverStatus: 'stopped', receiverMessage: 'Media stopped from phone.', stoppedAt: Date.now() } }, { merge: true }); } catch (error) {}
      }
    };
    const bind = (id, type) => { const node = $(id); if (node) node.addEventListener('click', () => command(type)); };
    bind('remoteMediaPlay', 'play');
    bind('remoteMediaPause', 'pause');
    bind('remoteMediaShow', 'show');
    bind('remoteMediaHide', 'hide');
    bind('remoteMediaStop', 'stop');
    const volume = $('remoteMediaVolume');
    if (volume) volume.addEventListener('input', () => {
      const value = Math.max(0, Math.min(1, Number(volume.value) / 100));
      command('volume', value);
      const label = $('remoteMediaVolumeLabel');
      if (label) label.textContent = `${Math.round(value * 100)}%`;
    });
  }

  function makeSessionId() {
    return Math.random().toString(36).slice(2, 6).toUpperCase() + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
  }

  async function publishSessionState(force = false) {
    if (!state.sessionRef || (!state.activeFile && !state.mediaMode) || (state.publishLock && !force)) return;
    state.publishLock = true;
    setTimeout(() => { state.publishLock = false; }, force ? 0 : 120);

    // Keep the hot remote-control path light: write slide/control state first,
    // then upload the heavier thumbnail preview in a short deferred task.
    const payload = {
      fileName: state.activeFile ? state.activeFile.name : 'Media Viewing',
      type: state.activeFile ? state.activeFile.type : 'media',
      currentPage: state.activeFile ? state.currentPage : 0,
      totalPages: state.activeFile ? state.totalPages : 0,
      mediaMode: !state.activeFile && !!state.mediaMode,
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
      magicEffectVolume: state.magicEffectVolume,
      magicEffectSound: !!state.magicEffectSound,
      magicEffectIntensity: state.magicEffectIntensity || 'grand',
      classroom: classroomRemoteSnapshot(),
      updatedAt: Date.now(),
    };

    try {
      await state.sessionRef.set(payload, { merge: true });
      if (state.activeFile) scheduleRemotePreviewPublish(force ? 60 : 220);
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
      case 'magicEffect':
      case 'triggerMagicEffect':
      case 'remoteMagicEffect':
        triggerMagicEffect(command.value || 'confetti');
        break;
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
      case 'setMagicEffectVolume':
        setMagicEffectVolume(command.value, false);
        publishSessionState();
        break;
      case 'setMagicEffectSound':
        setMagicEffectSound(command.value !== false && command.value !== 'false' && command.value !== 0, false);
        publishSessionState();
        break;
      case 'setMagicEffectIntensity':
        setMagicEffectIntensity(command.value || 'grand', false);
        publishSessionState();
        break;
      case 'testMagicEffect':
        triggerMagicEffect(command.value || 'confetti');
        break;
      case 'mediaCastControl':
        controlMediaCast(command.value || {});
        break;
      case 'mediaLinkCast':
        showOnlineMediaCast(command.value || {});
        break;
      case 'classroomSetSection':
        state.classroom.activeSectionId = String(command.value || '');
        state.classroom.pickedIds = [];
        state.classroom.rolledGroups = [];
        renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true);
        break;
      case 'classroomSetRepeatMode':
        state.classroom.removePicked = String(command.value || '') !== 'allow-repeat';
        renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true);
        break;
      case 'classroomSetGroupMode': {
        const mode = ['balanced', 'no-repeat', 'free'].includes(String(command.value || '')) ? String(command.value) : 'balanced';
        state.classroom.groupRollMode = mode;
        state.classroom.balanced = mode === 'balanced';
        state.classroom.rolledGroups = [];
        renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true);
        break;
      }
      case 'classroomPick': pickRandomStudent(); break;
      case 'classroomRoll': rollGroupDice(); break;
      case 'classroomReset': resetClassroomRound(); break;
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
    if (!state.activeFile && !state.mediaMode) return;
    if (!state.firebaseReady) {
      els.qrHelp.textContent = 'Phone remote across devices needs Firebase config. You can still present locally. Click Remote setup on the home screen to see where to add your Firebase keys.';
    } else {
      await setupRemoteSessionIfPossible();
      els.qrHelp.textContent = state.mediaMode && !state.activeFile ? `Media Viewing session ${state.sessionId} is live. Scan the CONTROL QR, then cast files or paste links from the phone.` : `Session ${state.sessionId} is live. Scan the CONTROL QR for buttons. Viewer QR is preview-only / view-only.`;
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
    let remoteMagicVolumeTimer = null;

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

        <section class="remote-section remote-premium-section remote-media-cast-section" data-host-only="true">
          <div class="remote-section-title"><span>Media Cast Beta</span><small>Pictures, audio, or MP4 from this phone</small></div>
          <div class="remote-media-cast-card">
            <label class="remote-media-picker">Choose media
              <input id="remoteMediaFile" type="file" accept="image/*,audio/*,video/*,.mp4,.webm,.mov,.m4v,.mp3,.wav,.m4a,.aac,.ogg,.png,.jpg,.jpeg,.webp,.gif" data-host-only="true">
            </label>
            <div id="remoteMediaFileName" class="remote-media-file-name">No media selected</div>
            <button id="remoteMediaSend" type="button" class="remote-media-send" disabled data-host-only="true">Cast to Screen</button>
            <div class="remote-media-link-card">
              <label>Online media link
                <input id="remoteMediaLink" type="url" inputmode="url" placeholder="Paste YouTube, TikTok, MP4, MP3, or image link" data-host-only="true">
              </label>
              <button id="remoteMediaLinkCast" type="button" class="remote-media-link-send" data-host-only="true">Play Link on Screen</button>
            </div>
            <div class="remote-media-controls">
              <button id="remoteMediaPlay" type="button" data-host-only="true">Play</button>
              <button id="remoteMediaPause" type="button" data-host-only="true">Pause</button>
              <button id="remoteMediaShow" type="button" data-host-only="true">Show</button>
              <button id="remoteMediaHide" type="button" data-host-only="true">Hide</button>
              <button id="remoteMediaStop" type="button" data-host-only="true">Stop</button>
            </div>
            <label class="remote-media-volume">Volume <span id="remoteMediaVolumeLabel">90%</span>
              <input id="remoteMediaVolume" type="range" min="0" max="100" step="1" value="90" data-host-only="true">
            </label>
            <div id="remoteMediaStatus" class="remote-media-status">No permanent upload. File is sent only for this live session.</div>
          </div>
        </section>

        <section class="remote-section remote-premium-section remote-magic-section" data-host-only="true">
          <div class="remote-section-title"><span>Magic Effects</span><small>Hidden until you open the list</small></div>
          <div class="remote-magic-settings-card">
            <div class="remote-magic-settings-head">
              <strong>Magic Settings</strong>
              <button id="remoteMagicTest" type="button" class="remote-mini-pill" data-host-only="true">Test</button>
            </div>
            <label class="remote-magic-switch"><input id="remoteMagicSound" type="checkbox" checked data-host-only="true"> Effects sound</label>
            <label>Effect volume <span id="remoteMagicVolumeLabel">200%</span>
              <input id="remoteMagicVolume" type="range" min="0" max="300" step="5" value="200" data-host-only="true">
            </label>
            <label>Effect intensity
              <select id="remoteMagicIntensity" data-host-only="true">
                <option value="low">Low</option>
                <option value="normal">Normal</option>
                <option value="grand" selected>Grand</option>
              </select>
            </label>
          </div>
          <button id="remoteMagicToggle" type="button" class="remote-magic-toggle" aria-expanded="false" data-host-only="true">✨ Show Magic Effects</button>
          <div id="remoteMagicGrid" class="remote-magic-grid remote-magic-collapsed" hidden aria-hidden="true">
            ${MAGIC_EFFECTS.map((effect) => `<button type="button" class="remote-magic-btn" data-magic-effect="${effect.id}" data-host-only="true"><span>${effect.emoji}</span><strong>${effect.label}</strong><small>${effect.shortcut}</small></button>`).join('')}
          </div>
        </section>

        <section class="remote-section remote-premium-section remote-classroom-section" data-host-only="true">
          <div class="remote-section-title"><span>Classroom Randomizer</span><small>Choose section, pick name, then roll group</small></div>
          <select id="remoteClassSection" data-host-only="true"><option value="">No saved section</option></select>
          <div class="remote-class-options">
            <label>Name picks
              <select id="remoteNameRepeat" data-host-only="true">
                <option value="no-repeat">No repeat</option>
                <option value="allow-repeat">Allow repeat</option>
              </select>
            </label>
            <label>Group rolls
              <select id="remoteGroupMode" data-host-only="true">
                <option value="balanced">Balanced</option>
                <option value="no-repeat">No repeat</option>
                <option value="free">Free repeat</option>
              </select>
            </label>
          </div>
          <div id="remoteClassResult" class="remote-class-result">Ready</div>
          <div class="remote-class-actions">
            <button id="remotePickName" data-host-only="true">Pick Name</button>
            <button id="remoteRollGroup" data-host-only="true">Roll Group</button>
            <button id="remoteClassReset" data-host-only="true">Reset</button>
          </div>
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
              <option value="slide-left">Slide left transition</option>
              <option value="slide-right">Slide right transition</option>
              <option value="slide-up">Slide up transition</option>
              <option value="slide-down">Slide down transition</option>
              <option value="zoom-in">Zoom in transition</option>
              <option value="zoom-out">Zoom out transition</option>
              <option value="soft-blur">Soft blur transition</option>
              <option value="bounce-pop">Bounce pop transition</option>
              <option value="happy-pop">Happy pop transition</option>
              <option value="elastic-pop">Elastic pop transition</option>
              <option value="sparkle-zoom">Sparkle zoom transition</option>
              <option value="page-flip">Page flip transition</option>
              <option value="paper-turn">Paper turn transition</option>
              <option value="circle-reveal">Circle reveal transition</option>
              <option value="corner-reveal">Corner reveal transition</option>
              <option value="bright-wipe">Bright wipe transition</option>
              <option value="ribbon-wipe">Ribbon wipe transition</option>
              <option value="split-open">Split open transition</option>
              <option value="drop-in">Drop in transition</option>
              <option value="float-up">Float up transition</option>
              <option value="diagonal-rise">Diagonal rise transition</option>
              <option value="diagonal-fall">Diagonal fall transition</option>
              <option value="smooth-swing">Smooth swing transition</option>
              <option value="push-left">Push left transition</option>
              <option value="push-right">Push right transition</option>
              <option value="carousel">Carousel transition</option>
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
        <div class="remote-draw-toolbar" data-host-only="true" aria-label="Preview zoom and drawing tools. Pinch and drag zooms or pans unless a drawing tool is active.">
          <button id="remotePreviewFit" class="remote-preview-zoom-btn" type="button">Fit</button>
          <button id="remotePreviewZoomOut" class="remote-preview-zoom-btn" type="button">−</button>
          <span id="remotePreviewZoomLabel" class="remote-preview-zoom-label">100%</span>
          <button id="remotePreviewZoomIn" class="remote-preview-zoom-btn" type="button">+</button>
          <span class="remote-tool-divider" aria-hidden="true"></span>
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
        <div class="remote-full-help">Pinch to zoom. Drag to pan when zoomed. Tap Pen/Highlight/Erase only when you want to draw.</div>
      </div>
    `;

    if (!isHost) {
      els.remoteApp.querySelectorAll('[data-host-only="true"]').forEach((node) => {
        node.setAttribute('aria-disabled', 'true');
        if ('disabled' in node) node.disabled = true;
      });
    }

    const remoteClassSection = $('remoteClassSection');
    const remoteNameRepeat = $('remoteNameRepeat');
    const remoteGroupMode = $('remoteGroupMode');
    const remoteRef = () => state.firebaseDb && sessionId ? state.firebaseDb.collection(SESSION_COLLECTION).doc(sessionId) : null;
    if (remoteClassSection) remoteClassSection.addEventListener('change', () => { const ref = remoteRef(); if (ref) sendRemoteCommand(ref, 'classroomSetSection', remoteClassSection.value); });
    if (remoteNameRepeat) remoteNameRepeat.addEventListener('change', () => { const ref = remoteRef(); if (ref) sendRemoteCommand(ref, 'classroomSetRepeatMode', remoteNameRepeat.value); });
    if (remoteGroupMode) remoteGroupMode.addEventListener('change', () => { const ref = remoteRef(); if (ref) sendRemoteCommand(ref, 'classroomSetGroupMode', remoteGroupMode.value); });
    const bindClassCommand = (id, action) => { const node = $(id); if (node) node.addEventListener('click', () => { const ref = remoteRef(); if (ref) sendRemoteCommand(ref, action); }); };
    bindClassCommand('remotePickName', 'classroomPick');
    bindClassCommand('remoteRollGroup', 'classroomRoll');
    bindClassCommand('remoteClassReset', 'classroomReset');

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
        if ($('remoteClassSection') && data.classroom) {
          const cs = data.classroom;
          $('remoteClassSection').innerHTML = (cs.sections || []).map((s) => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)} (${s.count})</option>`).join('') || '<option value="">No saved section</option>';
          $('remoteClassSection').value = cs.activeSectionId || '';
          if ($('remoteNameRepeat')) $('remoteNameRepeat').value = cs.removePicked === false ? 'allow-repeat' : 'no-repeat';
          if ($('remoteGroupMode')) $('remoteGroupMode').value = ['balanced', 'no-repeat', 'free'].includes(cs.groupRollMode) ? cs.groupRollMode : (cs.balanced === false ? 'free' : 'balanced');
          const bits = [];
          if (cs.revealMode === 'name' && cs.lastStudent) bits.push(cs.lastStudent);
          if (cs.revealMode === 'group' && cs.lastGroup) bits.push(`Group ${cs.lastGroup}`);
          if (!bits.length && Number.isFinite(cs.available)) bits.push(`${cs.available} available`);
          $('remoteClassResult').textContent = bits.join(' • ') || 'Ready';
          const classBusy = !!cs.busy;
          if ($('remotePickName')) $('remotePickName').disabled = !isHost || classBusy;
          if ($('remoteRollGroup')) $('remoteRollGroup').disabled = !isHost || classBusy;
          if ($('remoteClassReset')) $('remoteClassReset').disabled = !isHost || classBusy;
        }
        window.__phRemoteCurrentPage = Number(data.currentPage) || 1;
        window.__phRemoteInkStrokes = Array.isArray(data.inkStrokes) ? data.inkStrokes : [];
        $('remoteStatusPill').textContent = isHost ? 'Host' : 'Viewer';
        const isMediaSession = data.mediaMode || data.type === 'media';
        $('remoteSlideLabel').textContent = isMediaSession ? 'Media Viewing Mode' : `${data.type === 'pdf' ? 'Page' : 'Slide'} ${data.currentPage || '--'} / ${data.totalPages || '--'}`;
        $('remoteFullLabel').textContent = isMediaSession ? 'Media Preview' : `${data.type === 'pdf' ? 'Page' : 'Slide'} ${data.currentPage || '--'} / ${data.totalPages || '--'}`;
        $('remoteFileLabel').textContent = isMediaSession ? 'Cast local files or paste online links from this phone.' : (data.fileName || 'Active presentation');
        if (data.thumb) latestThumb = data.thumb;
        const preview = $('remotePreview');
        preview.innerHTML = isMediaSession ? '<span>Media screen ready. Use Media Cast below.</span>' : (latestThumb ? `<img src="${latestThumb}" alt="Current slide preview">` : '<span>No preview yet</span>');
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
        if ($('remoteMagicVolume')) $('remoteMagicVolume').value = data.magicEffectVolume ?? 200;
        if ($('remoteMagicVolumeLabel')) $('remoteMagicVolumeLabel').textContent = `${data.magicEffectVolume ?? 200}%`;
        if ($('remoteMagicSound')) $('remoteMagicSound').checked = data.magicEffectSound !== false;
        if ($('remoteMagicIntensity')) $('remoteMagicIntensity').value = data.magicEffectIntensity || 'grand';
        syncRemoteMediaStatus(data && data.mediaCast);
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

      els.remoteApp.querySelectorAll('[data-magic-effect]').forEach((button) => {
        button.addEventListener('click', () => {
          if (!isHost) return;
          sendRemoteCommand(ref, 'magicEffect', button.dataset.magicEffect);
        });
      });
      const magicToggle = $('remoteMagicToggle');
      const magicGrid = $('remoteMagicGrid');
      if (magicGrid) {
        magicGrid.hidden = true;
        magicGrid.classList.add('remote-magic-collapsed');
        magicGrid.setAttribute('aria-hidden', 'true');
        magicGrid.style.display = 'none';
      }
      if (magicToggle && magicGrid) {
        magicToggle.addEventListener('click', () => {
          const open = magicGrid.hidden;
          magicGrid.hidden = !open;
          magicGrid.classList.toggle('remote-magic-collapsed', !open);
          magicGrid.setAttribute('aria-hidden', open ? 'false' : 'true');
          magicGrid.style.display = open ? 'grid' : 'none';
          magicToggle.setAttribute('aria-expanded', String(open));
          magicToggle.textContent = open ? '✨ Hide Magic Effects' : '✨ Show Magic Effects';
        });
      }
      if ($('remoteMagicVolume')) {
        const sendMagicVolume = () => {
          if (!isHost) return;
          const value = Number($('remoteMagicVolume').value) || 0;
          sendRemoteCommand(ref, 'setMagicEffectVolume', value);
        };
        $('remoteMagicVolume').addEventListener('input', () => {
          if (!isHost) return;
          const value = Number($('remoteMagicVolume').value) || 0;
          if ($('remoteMagicVolumeLabel')) $('remoteMagicVolumeLabel').textContent = `${value}%`;
          clearTimeout(remoteMagicVolumeTimer);
          remoteMagicVolumeTimer = setTimeout(sendMagicVolume, 120);
        });
        $('remoteMagicVolume').addEventListener('change', sendMagicVolume);
      }
      if ($('remoteMagicSound')) {
        $('remoteMagicSound').addEventListener('change', () => {
          if (!isHost) return;
          sendRemoteCommand(ref, 'setMagicEffectSound', $('remoteMagicSound').checked);
        });
      }
      if ($('remoteMagicIntensity')) {
        $('remoteMagicIntensity').addEventListener('change', () => {
          if (!isHost) return;
          sendRemoteCommand(ref, 'setMagicEffectIntensity', $('remoteMagicIntensity').value);
        });
      }
      if ($('remoteMagicTest')) {
        $('remoteMagicTest').addEventListener('click', () => {
          if (!isHost) return;
          sendRemoteCommand(ref, 'testMagicEffect', 'confetti');
        });
      }

      attachRemoteMediaCastControls(ref, isHost);

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
    const fullImg = $('remoteFullImg');
    const fitBtn = $('remotePreviewFit');
    const zoomOutBtn = $('remotePreviewZoomOut');
    const zoomInBtn = $('remotePreviewZoomIn');
    const zoomLabel = $('remotePreviewZoomLabel');
    if (!openBtn || !full || !stage) return;

    const normalizeViewport = (value = {}) => ({
      zoom: Math.min(4, Math.max(1, Number(value.zoom) || 1)),
      centerX: Math.max(0, Math.min(1, Number.isFinite(Number(value.centerX)) ? Number(value.centerX) : 0.5)),
      centerY: Math.max(0, Math.min(1, Number.isFinite(Number(value.centerY)) ? Number(value.centerY) : 0.5)),
    });
    const safeViewport = (value) => clampViewportToRemoteScreen(normalizeViewport(value));

    let localViewport = normalizeViewport(getViewport());
    let rafId = 0;
    let gestureSettleTimer = null;
    let resizeFrame = 0;

    const sendViewportLive = throttle((viewport) => {
      if (!isHost) return;
      sendRemoteCommand(ref, 'setViewport', safeViewport(viewport));
    }, 145);

    function updateZoomLabel() {
      if (zoomLabel) zoomLabel.textContent = `${Math.round((localViewport.zoom || 1) * 100)}%`;
    }

    function renderLocal(viewport) {
      localViewport = safeViewport(viewport);
      setLocalViewport(localViewport);
      updateZoomLabel();
      if (rafId) return;
      rafId = requestAnimationFrame(() => {
        rafId = 0;
        applyRemoteFullPreviewTransform(localViewport);
      });
    }

    function commitViewport(viewport = localViewport) {
      localViewport = safeViewport(viewport);
      renderLocal(localViewport);
      if (isHost) sendRemoteCommand(ref, 'setViewport', localViewport);
    }

    function beginGesture() {
      window.__presentationHubRemoteGestureActive = true;
      clearTimeout(gestureSettleTimer);
    }

    function endGesture(delay = 110) {
      clearTimeout(gestureSettleTimer);
      gestureSettleTimer = setTimeout(() => {
        window.__presentationHubRemoteGestureActive = false;
        commitViewport(localViewport);
      }, delay);
    }

    function fitPreview() {
      beginGesture();
      renderLocal({ zoom: 1, centerX: 0.5, centerY: 0.5 });
      endGesture(40);
    }

    function stepZoom(multiplier) {
      beginGesture();
      const current = safeViewport(localViewport || getViewport());
      const nextZoom = Math.min(4, Math.max(1, current.zoom * multiplier));
      renderLocal({ ...current, zoom: nextZoom });
      endGesture(40);
    }

    function refreshPreviewLayout() {
      if (full.classList.contains('hidden')) return;
      cancelAnimationFrame(resizeFrame);
      resizeFrame = requestAnimationFrame(() => {
        renderLocal(localViewport);
        requestAnimationFrame(() => renderLocal(localViewport));
      });
    }

    function openFullPreview() {
      localViewport = normalizeViewport(getViewport());
      full.classList.remove('hidden');
      document.body.classList.add('remote-fullscreen-open');
      requestAnimationFrame(() => {
        renderLocal(localViewport);
        requestAnimationFrame(() => renderLocal(localViewport));
      });
      if (fullImg && !fullImg.complete) fullImg.addEventListener('load', refreshPreviewLayout, { once: true });
      if (full.requestFullscreen) full.requestFullscreen().catch(() => {});
    }

    function closeFullPreview() {
      clearTimeout(gestureSettleTimer);
      window.__presentationHubRemoteGestureActive = false;
      full.classList.add('hidden');
      document.body.classList.remove('remote-fullscreen-open');
      document.documentElement.classList.add('presentation-hub-remote-mode');
      document.body.classList.add('presentation-hub-remote-mode');
      if (document.fullscreenElement === full) document.exitFullscreen().catch(() => {});
      requestAnimationFrame(() => document.body.classList.remove('remote-fullscreen-open'));
    }

    openBtn.addEventListener('click', openFullPreview);
    if (closeBtn) closeBtn.addEventListener('click', closeFullPreview);
    if (backBtn) backBtn.addEventListener('click', closeFullPreview);
    if (fitBtn) fitBtn.addEventListener('click', fitPreview);
    if (zoomOutBtn) zoomOutBtn.addEventListener('click', () => stepZoom(0.8));
    if (zoomInBtn) zoomInBtn.addEventListener('click', () => stepZoom(1.25));
    document.addEventListener('fullscreenchange', () => {
      if (!document.fullscreenElement && !full.classList.contains('hidden')) closeFullPreview();
      else refreshPreviewLayout();
    });
    window.addEventListener('resize', refreshPreviewLayout, { passive: true });
    window.addEventListener('orientationchange', () => setTimeout(refreshPreviewLayout, 160), { passive: true });
    if (window.ResizeObserver) {
      const screen = stage.querySelector('.remote-monitor-screen') || stage;
      new ResizeObserver(refreshPreviewLayout).observe(screen);
    }
    if (fullImg) fullImg.addEventListener('load', refreshPreviewLayout);

    if (!isHost) return;

    let startDistance = 0;
    let startZoom = 1;
    let startCenter = { centerX: 0.5, centerY: 0.5 };
    let pinchAnchor = { x: 0.5, y: 0.5 };
    let startPoint = null;
    let touchMoved = false;
    let lastTapAt = 0;

    const distance = (touches) => {
      const dx = touches[0].clientX - touches[1].clientX;
      const dy = touches[0].clientY - touches[1].clientY;
      return Math.hypot(dx, dy);
    };

    const midpointPixels = (touches, rect) => ({
      x: ((touches[0].clientX + touches[1].clientX) / 2) - rect.left,
      y: ((touches[0].clientY + touches[1].clientY) / 2) - rect.top,
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

    function startPanGesture(touch) {
      const current = safeViewport(localViewport || getViewport());
      startZoom = current.zoom;
      startCenter = { centerX: current.centerX, centerY: current.centerY };
      startPoint = { x: touch.clientX, y: touch.clientY };
      startDistance = 0;
      touchMoved = false;
    }

    function startPinchGesture(touches) {
      const current = safeViewport(localViewport || getViewport());
      const geometry = screenGeometry(current.zoom);
      const mid = midpointPixels(touches, geometry.rect);
      startZoom = current.zoom;
      startCenter = { centerX: current.centerX, centerY: current.centerY };
      startDistance = Math.max(1, distance(touches));
      pinchAnchor = {
        x: current.centerX + ((mid.x - geometry.rect.width / 2) / Math.max(1, geometry.contentW)),
        y: current.centerY + ((mid.y - geometry.rect.height / 2) / Math.max(1, geometry.contentH)),
      };
      startPoint = null;
      touchMoved = false;
    }

    // No separate Move tool: the normal state is zoom/pan.
    // Tap Pen/Highlight/Erase to draw; tap the same tool again to return to zoom/pan.
    let activeTool = null;
    let currentStroke = null;
    const toolButtons = {
      pen: $('remoteToolPen'),
      highlighter: $('remoteToolHighlighter'),
      erase: $('remoteToolErase'),
    };
    const isDrawingTool = (tool) => tool === 'pen' || tool === 'highlighter' || tool === 'erase';

    function setActiveTool(tool) {
      activeTool = activeTool === tool ? null : tool;
      Object.entries(toolButtons).forEach(([name, button]) => {
        if (button) button.classList.toggle('active', name === activeTool);
      });
      const screen = stage.querySelector('.remote-monitor-screen');
      if (screen) {
        if (activeTool) screen.dataset.tool = activeTool;
        else delete screen.dataset.tool;
      }
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
      if (isDrawingTool(activeTool) && event.touches.length === 1) {
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
      if (event.touches.length >= 2) startPinchGesture(event.touches);
      else if (event.touches.length === 1) startPanGesture(event.touches[0]);
    }, { passive: false });

    stage.addEventListener('touchmove', (event) => {
      if (event.touches.length !== 1 && event.touches.length !== 2) return;
      event.preventDefault();
      if (isDrawingTool(activeTool) && event.touches.length === 1) {
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
      let next = safeViewport(localViewport || getViewport());

      if (event.touches.length === 2) {
        if (!startDistance) startPinchGesture(event.touches);
        const nextZoom = Math.min(4, Math.max(1, startZoom * (distance(event.touches) / Math.max(1, startDistance))));
        const geometry = screenGeometry(nextZoom);
        const mid = midpointPixels(event.touches, geometry.rect);
        next = safeViewport({
          zoom: nextZoom,
          centerX: pinchAnchor.x - ((mid.x - geometry.rect.width / 2) / Math.max(1, geometry.contentW)),
          centerY: pinchAnchor.y - ((mid.y - geometry.rect.height / 2) / Math.max(1, geometry.contentH)),
        });
        touchMoved = true;
      } else if (event.touches.length === 1) {
        if (!startPoint) startPanGesture(event.touches[0]);
        const dx = event.touches[0].clientX - startPoint.x;
        const dy = event.touches[0].clientY - startPoint.y;
        if (Math.hypot(dx, dy) > 4) touchMoved = true;
        if (startZoom > 1.001) {
          const geometry = screenGeometry(startZoom);
          next = safeViewport({
            zoom: startZoom,
            centerX: startCenter.centerX - dx / Math.max(1, geometry.contentW),
            centerY: startCenter.centerY - dy / Math.max(1, geometry.contentH),
          });
        } else {
          next = { zoom: 1, centerX: 0.5, centerY: 0.5 };
        }
      }

      renderLocal(next);
      sendViewportLive(next);
    }, { passive: false });

    stage.addEventListener('touchend', (event) => {
      if (currentStroke) {
        const finished = currentStroke;
        currentStroke = null;
        window.__phRemoteTempStroke = null;
        if (finished.points && finished.points.length > 1) sendRemoteCommand(ref, 'addInkStroke', finished);
        renderRemoteInkPreview(window.__phRemoteInkStrokes || [], window.__phRemoteCurrentPage || 1, localViewport);
        return;
      }
      if (event.touches && event.touches.length === 1) {
        startPanGesture(event.touches[0]);
        touchMoved = true;
        return;
      }
      if (!touchMoved && !activeTool) {
        const now = Date.now();
        if (now - lastTapAt < 320) {
          lastTapAt = 0;
          fitPreview();
          return;
        }
        lastTapAt = now;
      }
      startDistance = 0;
      startPoint = null;
      endGesture();
    }, { passive: true });

    stage.addEventListener('touchcancel', () => {
      currentStroke = null;
      window.__phRemoteTempStroke = null;
      renderRemoteInkPreview(window.__phRemoteInkStrokes || [], window.__phRemoteCurrentPage || 1, localViewport);
      startDistance = 0;
      startPoint = null;
      endGesture();
    }, { passive: true });

    stage.addEventListener('dblclick', (event) => {
      if (activeTool) return;
      event.preventDefault();
      fitPreview();
    });

    stage.addEventListener('wheel', (event) => {
      event.preventDefault();
      beginGesture();
      const current = safeViewport(localViewport || getViewport());
      const geometry = screenGeometry(current.zoom);
      const pointer = { x: event.clientX - geometry.rect.left, y: event.clientY - geometry.rect.top };
      const anchor = {
        x: current.centerX + ((pointer.x - geometry.rect.width / 2) / Math.max(1, geometry.contentW)),
        y: current.centerY + ((pointer.y - geometry.rect.height / 2) / Math.max(1, geometry.contentH)),
      };
      const nextZoom = Math.min(4, Math.max(1, current.zoom * (event.deltaY < 0 ? 1.12 : 0.89)));
      const nextGeometry = screenGeometry(nextZoom);
      const next = safeViewport({
        zoom: nextZoom,
        centerX: anchor.x - ((pointer.x - nextGeometry.rect.width / 2) / Math.max(1, nextGeometry.contentW)),
        centerY: anchor.y - ((pointer.y - nextGeometry.rect.height / 2) / Math.max(1, nextGeometry.contentH)),
      });
      renderLocal(next);
      sendViewportLive(next);
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
    const id = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    const command = { id, action, value: value ?? null, issuedAt: Date.now(), v: 12 };
    const mirror = action === 'magicEffect' || action === 'testMagicEffect'
      ? { lastMagicCommand: command, lastMagicCommandAt: Date.now() }
      : {};
    return ref.set({ command, ...mirror }, { merge: true });
  }


  function loadClassroomLocal() {
    try {
      const saved = JSON.parse(localStorage.getItem('presentationHubClassroom') || 'null');
      if (saved && Array.isArray(saved.sections)) {
        state.classroom = { ...state.classroom, ...saved };
        state.classroom.groupRollMode = ['balanced', 'no-repeat', 'free'].includes(state.classroom.groupRollMode)
          ? state.classroom.groupRollMode
          : (state.classroom.balanced === false ? 'free' : 'balanced');
        state.classroom.balanced = state.classroom.groupRollMode === 'balanced';
        state.classroom.rolledGroups = Array.isArray(state.classroom.rolledGroups) ? state.classroom.rolledGroups : [];
      }
    } catch (_) {}
  }

  function saveClassroomLocal() { localStorage.setItem('presentationHubClassroom', JSON.stringify(state.classroom)); }
  function scheduleClassroomSave() {
    saveClassroomLocal();
    clearTimeout(state.classroomSyncTimer);
    state.classroomSyncTimer = setTimeout(saveClassroomCloud, 350);
  }
  async function saveClassroomCloud() {
    const user = state.firebaseUser;
    if (!state.firebaseReady || !user || user.isAnonymous) return;
    try { await state.firebaseDb.collection('presentationHubUsers').doc(user.uid).collection('data').doc('classroom').set({ ...state.classroom, updatedAt: Date.now() }); } catch (e) { console.warn('Classroom cloud save failed', e); }
  }
  async function loadClassroomCloud() {
    const user = state.firebaseUser;
    if (!state.firebaseReady || !user || user.isAnonymous) return;
    try {
      const snap = await state.firebaseDb.collection('presentationHubUsers').doc(user.uid).collection('data').doc('classroom').get();
      if (snap.exists) state.classroom = { ...state.classroom, ...snap.data() };
      else await saveClassroomCloud();
      state.classroom.groupRollMode = ['balanced', 'no-repeat', 'free'].includes(state.classroom.groupRollMode)
        ? state.classroom.groupRollMode
        : (state.classroom.balanced === false ? 'free' : 'balanced');
      state.classroom.balanced = state.classroom.groupRollMode === 'balanced';
      state.classroom.rolledGroups = Array.isArray(state.classroom.rolledGroups) ? state.classroom.rolledGroups : [];
      saveClassroomLocal(); renderClassroomTools(); publishSessionState(true);
    } catch (e) { console.warn('Classroom cloud load failed', e); }
  }
  async function signInClassroomAccount() {
    if (!window.firebase || !hasFirebaseConfig()) return alert('Firebase config is needed first.');
    try {
      const provider = new firebase.auth.GoogleAuthProvider();
      const current = firebase.auth().currentUser;
      let result;
      if (current && current.isAnonymous) {
        try { result = await current.linkWithPopup(provider); }
        catch (e) { if (e.code === 'auth/credential-already-in-use') result = await firebase.auth().signInWithCredential(e.credential); else throw e; }
      } else result = await firebase.auth().signInWithPopup(provider);
      state.firebaseUser = result.user;
      await loadClassroomCloud(); updateClassroomAuthUI();
    } catch (e) { console.warn(e); alert('Google sign-in did not finish. Check Firebase Authentication > Google provider and authorized domains.'); }
  }
  function updateClassroomAuthUI() {
    if (!els.classroomAuthStatus || !els.classroomAuthBtn) return;
    const user = state.firebaseUser;
    const signed = user && !user.isAnonymous;
    els.classroomAuthStatus.textContent = signed ? `Synced as ${user.displayName || user.email || 'Google account'}` : 'Local save active. Sign in to sync across devices.';
    els.classroomAuthBtn.textContent = signed ? 'Account Synced' : 'Sign in with Google';
    els.classroomAuthBtn.disabled = !!signed;
  }
  function openClassroomTools() { renderClassroomTools(); updateClassroomAuthUI(); showModal(els.classroomModal); }
  function activeClassSection() { return state.classroom.sections.find(s => s.id === state.classroom.activeSectionId) || null; }
  function addClassroomSection() {
    const name = (els.newSectionName.value || '').trim(); if (!name) return;
    const id = `sec-${Date.now().toString(36)}`;
    state.classroom.sections.push({ id, name, students: [] }); state.classroom.activeSectionId = id; els.newSectionName.value = '';
    renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true);
  }
  function normalizeClassHeader(value) {
    return String(value || '').replace(/^\uFEFF/, '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
  }
  function cleanStudentName(value) {
    return String(value == null ? '' : value)
      .replace(/[\r\n\t]+/g, ' ')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }
  function cleanClassValue(value) {
    return String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ').replace(/\s{2,}/g, ' ').trim();
  }
  function findClassColumn(row, aliases) {
    const keys = Object.keys(row || {});
    const wanted = aliases.map(normalizeClassHeader);
    return keys.find((key) => wanted.includes(normalizeClassHeader(key))) || '';
  }
  function makeStudentId(sectionId, studentId, name, index) {
    const source = String(studentId || name || index).trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
    return `${sectionId}-${source || index}`;
  }
  function rowsFromClassSheet(sheet) {
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });
    const aliases = new Set(['studentid','studentnumber','studentno','id','lrn','name','studentname','fullname','learnername','firstname','middlename','lastname','surname','gender','sex','section','class','classname','gradeandsection']);
    let headerIndex = matrix.findIndex((row) => Array.isArray(row) && row.some((cell) => aliases.has(normalizeClassHeader(cell))));
    if (headerIndex < 0) headerIndex = matrix.findIndex((row) => Array.isArray(row) && row.some((cell) => cleanClassValue(cell)));
    if (headerIndex < 0) return [];
    const headerCounts = {};
    const headers = matrix[headerIndex].map((cell, i) => {
      const base = cleanClassValue(cell) || `Column${i + 1}`;
      headerCounts[base] = (headerCounts[base] || 0) + 1;
      return headerCounts[base] === 1 ? base : `${base}__${headerCounts[base]}`;
    });
    return matrix.slice(headerIndex + 1).map((cells) => {
      const row = {};
      headers.forEach((header, i) => { row[header] = cells[i] == null ? '' : cells[i]; });
      return row;
    }).filter((row) => Object.values(row).some((value) => cleanClassValue(value)));
  }
  async function importClassListFile(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    if (els.classImportStatus) els.classImportStatus.textContent = 'Reading class list...';
    try {
      if (!window.XLSX) throw new Error('Excel reader is not available. Connect to the internet once, then reload the app.');
      let rows = [];
      const lower = file.name.toLowerCase();
      if (lower.endsWith('.csv')) {
        const text = await file.text();
        const workbook = XLSX.read(text, { type: 'string' });
        for (const sheetName of workbook.SheetNames) rows.push(...rowsFromClassSheet(workbook.Sheets[sheetName]));
      } else {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });
        for (const sheetName of workbook.SheetNames) rows.push(...rowsFromClassSheet(workbook.Sheets[sheetName]));
      }
      if (!rows.length) throw new Error('No student rows were found. Make sure the file has a header row.');
      const sample = rows.find(row => Object.values(row || {}).some(value => cleanClassValue(value))) || {};
      const idKey = findClassColumn(sample, ['student id','student number','student no','id','lrn']);
      const nameKey = findClassColumn(sample, ['name','student name','full name','learner name']);
      const firstNameKey = findClassColumn(sample, ['first name','firstname','given name']);
      const middleNameKey = findClassColumn(sample, ['middle name','middlename','middle initial']);
      const lastNameKey = findClassColumn(sample, ['last name','lastname','surname','family name']);
      const genderKey = findClassColumn(sample, ['gender','sex']);
      const sectionKey = findClassColumn(sample, ['section','class','class name','grade and section']);
      if (!nameKey && !firstNameKey && !lastNameKey) throw new Error('A Name column is required. Use: Student ID | Name | Gender | Section');
      const grouped = new Map();
      const importedRowKeys = new Set();
      for (const row of rows) {
        // When a dedicated Name/Full Name column exists, read that cell only.
        // Never join neighboring generic columns because that can merge two learners.
        let name = nameKey ? cleanStudentName(row[nameKey]) : '';
        if (!name && !nameKey) {
          name = [firstNameKey && row[firstNameKey], middleNameKey && row[middleNameKey], lastNameKey && row[lastNameKey]]
            .map(cleanStudentName).filter(Boolean).join(' ');
        }
        name = cleanStudentName(name);
        if (!name) continue;
        const sectionName = cleanClassValue((sectionKey && row[sectionKey]) || 'Imported Section') || 'Imported Section';
        const studentId = idKey ? cleanClassValue(row[idKey]) : '';
        const dedupeKey = studentId
          ? `${sectionName.toLowerCase()}|id:${studentId.toLowerCase()}`
          : `${sectionName.toLowerCase()}|name:${name.toLowerCase()}`;
        // Some school workbooks repeat the same class list on multiple sheets.
        if (importedRowKeys.has(dedupeKey)) continue;
        importedRowKeys.add(dedupeKey);
        if (!grouped.has(sectionName)) grouped.set(sectionName, []);
        grouped.get(sectionName).push({ studentId, name, gender: genderKey ? cleanClassValue(row[genderKey]) : '' });
      }
      if (!grouped.size) throw new Error('No complete student names were found in the file.');
      let importedStudents = 0;
      let updatedStudents = 0;
      for (const [sectionName, students] of grouped.entries()) {
        let section = state.classroom.sections.find(s => cleanClassValue(s.name).toLowerCase() === sectionName.toLowerCase());
        if (!section) {
          section = { id: `sec-${Date.now().toString(36)}-${Math.random().toString(36).slice(2,7)}`, name: sectionName, students: [] };
          state.classroom.sections.push(section);
        }
        section.students = (section.students || []).map(s => ({ ...s, name: cleanStudentName(s.name) }));
        const byKey = new Map();
        section.students.forEach((s) => {
          const idKeyValue = cleanClassValue(s.studentId).toLowerCase();
          const nameKeyValue = cleanStudentName(s.name).toLowerCase();
          if (idKeyValue) byKey.set(`id:${idKeyValue}`, s);
          if (nameKeyValue) byKey.set(`name:${nameKeyValue}`, s);
        });
        for (const student of students) {
          const idLookup = student.studentId ? `id:${student.studentId.toLowerCase()}` : '';
          const nameLookup = `name:${student.name.toLowerCase()}`;
          const existing = (idLookup && byKey.get(idLookup)) || byKey.get(nameLookup);
          if (existing) {
            existing.name = student.name;
            if (student.studentId) existing.studentId = student.studentId;
            if (student.gender) existing.gender = student.gender;
            updatedStudents++;
          } else {
            const record = { id: makeStudentId(section.id, student.studentId, student.name, section.students.length), ...student };
            section.students.push(record);
            if (student.studentId) byKey.set(`id:${student.studentId.toLowerCase()}`, record);
            byKey.set(nameLookup, record);
            importedStudents++;
          }
        }
        if (!state.classroom.activeSectionId) state.classroom.activeSectionId = section.id;
      }
      const firstImported = state.classroom.sections.find(s => grouped.has(s.name));
      if (firstImported) state.classroom.activeSectionId = firstImported.id;
      state.classroom.pickedIds = [];
      state.classroom.assignments = {};
      renderClassroomTools();
      scheduleClassroomSave();
      publishSessionState(true);
      if (els.classImportStatus) els.classImportStatus.textContent = `Imported ${importedStudents} new student${importedStudents === 1 ? '' : 's'}${updatedStudents ? ` and updated ${updatedStudents}` : ''} across ${grouped.size} section${grouped.size === 1 ? '' : 's'}.`;
    } catch (error) {
      console.warn('Class list import failed', error);
      if (els.classImportStatus) els.classImportStatus.textContent = error.message || 'Could not import this file.';
      alert(error.message || 'Could not import this class list.');
    } finally {
      event.target.value = '';
    }
  }
  function saveClassroomNames() {
    const sec = activeClassSection(); if (!sec) return alert('Create or select a section first.');
    const names = (els.studentNamesInput.value || '').split(/\r?\n/).map(cleanStudentName).filter(Boolean);
    sec.students = names.map((name, i) => ({ id: `${sec.id}-${i}-${name.toLowerCase().replace(/[^a-z0-9]+/g,'-')}`, name }));
    state.classroom.pickedIds = []; state.classroom.assignments = {}; renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true);
  }
  function renderClassroomTools() {
    if (!els.sectionSelect) return;
    els.sectionSelect.innerHTML = state.classroom.sections.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)} (${s.students.length})</option>`).join('') || '<option value="">Create a section</option>';
    if (!state.classroom.activeSectionId && state.classroom.sections[0]) state.classroom.activeSectionId = state.classroom.sections[0].id;
    els.sectionSelect.value = state.classroom.activeSectionId || '';
    const sec = activeClassSection();
    els.studentNamesInput.value = sec ? sec.students.map(s => cleanStudentName(s.name)).join('\n') : '';
    els.groupCountInput.value = state.classroom.groupCount;
    els.removePickedInput.checked = state.classroom.removePicked;
    state.classroom.groupRollMode = ['balanced', 'no-repeat', 'free'].includes(state.classroom.groupRollMode)
      ? state.classroom.groupRollMode
      : (state.classroom.balanced === false ? 'free' : 'balanced');
    state.classroom.balanced = state.classroom.groupRollMode === 'balanced';
    if (els.groupRollModeSelect) els.groupRollModeSelect.value = state.classroom.groupRollMode;
    if (els.balancedGroupsInput) els.balancedGroupsInput.checked = state.classroom.balanced;
    els.classroomResult.innerHTML = state.classroom.revealMode === 'name' && state.classroom.lastStudent ? `<strong>${escapeHtml(state.classroom.lastStudent.name)}</strong>` : state.classroom.revealMode === 'group' && state.classroom.lastGroup ? `<span>GROUP ${state.classroom.lastGroup}</span>` : '<span>Ready to pick</span>';
    [els.pickNameBtn, els.rollGroupBtn, els.resetPicksBtn].filter(Boolean).forEach(btn => { btn.disabled = !!state.classroom.busy; });
    renderClassroomPresentation();
    els.classroomRoster.innerHTML = sec ? sec.students.map(s => `<div><span title="${escapeHtml(cleanStudentName(s.name))}">${escapeHtml(cleanStudentName(s.name))}</span><b>${state.classroom.assignments[s.id] ? 'G'+state.classroom.assignments[s.id] : ''}</b></div>`).join('') : '';
  }
  function classroomRemoteSnapshot() {
    const sec = activeClassSection();
    return {
      sections: state.classroom.sections.map(s => ({ id:s.id, name:s.name, count:s.students.length })),
      activeSectionId: state.classroom.activeSectionId,
      groupCount: state.classroom.groupCount,
      removePicked: !!state.classroom.removePicked,
      balanced: !!state.classroom.balanced,
      groupRollMode: ['balanced', 'no-repeat', 'free'].includes(state.classroom.groupRollMode) ? state.classroom.groupRollMode : (state.classroom.balanced ? 'balanced' : 'free'),
      revealMode: state.classroom.revealMode || '',
      lastStudent: state.classroom.lastStudent && state.classroom.lastStudent.name,
      lastGroup: state.classroom.lastGroup,
      busy: !!state.classroom.busy,
      available: sec ? Math.max(0, sec.students.length - state.classroom.pickedIds.length) : 0
    };
  }
  function classroomNameSize(name, mode = 'roulette') {
    const length = cleanStudentName(name).length;
    if (mode === 'stage') {
      if (length > 42) return 'clamp(1.15rem, 2.7vw, 2.7rem)';
      if (length > 34) return 'clamp(1.35rem, 3vw, 3.2rem)';
      if (length > 26) return 'clamp(1.65rem, 3.6vw, 4rem)';
      return 'clamp(2rem, 5vw, 5.6rem)';
    }
    if (mode === 'final') {
      if (length > 42) return 'clamp(1.25rem, 3.2vw, 3.2rem)';
      if (length > 34) return 'clamp(1.45rem, 3.8vw, 4rem)';
      if (length > 26) return 'clamp(1.7rem, 4.8vw, 4.9rem)';
      return 'clamp(2rem, 6vw, 6.2rem)';
    }
    if (length > 42) return 'clamp(1.1rem, 2.9vw, 2.7rem)';
    if (length > 34) return 'clamp(1.25rem, 3.5vw, 3.4rem)';
    if (length > 26) return 'clamp(1.45rem, 4.2vw, 4.2rem)';
    return 'clamp(1.8rem, 5.2vw, 5.4rem)';
  }
  function classroomNameStyle(name, mode = 'roulette') {
    return `--name-size:${classroomNameSize(name, mode)}`;
  }
  function pickRoulettePoolName(pool, exclude = []) {
    const blocked = new Set((exclude || []).filter(Boolean));
    const filtered = pool.filter((entry) => entry && !blocked.has(entry));
    const source = filtered.length ? filtered : pool;
    return source[Math.floor(Math.random() * source.length)] || pool[0] || '';
  }
  function renderRouletteMovingStrip(track, pool, currentName, nextName) {
    const prevTwo = pickRoulettePoolName(pool, [currentName, nextName]);
    const prevOne = pickRoulettePoolName(pool, [currentName, nextName, prevTwo]);
    const afterOne = nextName || pickRoulettePoolName(pool, [currentName, prevTwo, prevOne]);
    const afterTwo = pickRoulettePoolName(pool, [currentName, nextName, prevTwo, prevOne, afterOne]);
    const afterThree = pickRoulettePoolName(pool, [currentName, nextName, prevTwo, prevOne, afterOne, afterTwo]);
    const rows = [prevTwo, prevOne, currentName, afterOne, afterTwo, afterThree];
    track.innerHTML = `<div class="roulette-strip">${rows.map((name, index) => `<div class="roulette-row ${index === 2 ? 'row-current' : index === 3 ? 'row-next' : ''}"><div class="roulette-name" style="${classroomNameStyle(name, 'roulette')}">${escapeHtml(name)}</div></div>`).join('')}</div>`;
    return track.firstElementChild;
  }
  function renderRouletteFinalStrip(track, pool, finalName) {
    const aboveTwo = pickRoulettePoolName(pool, [finalName]);
    const aboveOne = pickRoulettePoolName(pool, [finalName, aboveTwo]);
    const belowOne = pickRoulettePoolName(pool, [finalName, aboveTwo, aboveOne]);
    const belowTwo = pickRoulettePoolName(pool, [finalName, aboveTwo, aboveOne, belowOne]);
    const rows = [aboveTwo, aboveOne, finalName, belowOne, belowTwo];
    track.innerHTML = `<div class="roulette-strip final-strip">${rows.map((name, index) => `<div class="roulette-row ${index === 2 ? 'row-final-center' : ''}"><div class="roulette-name ${index === 2 ? 'roulette-final' : ''}" style="${classroomNameStyle(name, index === 2 ? 'final' : 'roulette')}">${escapeHtml(name)}</div></div>`).join('')}</div>`;
    return track.firstElementChild;
  }
  async function pickRandomStudent() {
    if (state.classroom.busy) return;
    const sec = activeClassSection(); if (!sec || !sec.students.length) return alert('Add student names first.');
    let pool = sec.students.filter(s => !state.classroom.removePicked || !state.classroom.pickedIds.includes(s.id));
    if (!pool.length) { state.classroom.pickedIds = []; pool = sec.students.slice(); }
    const chosen = pool[Math.floor(Math.random()*pool.length)];
    state.classroom.busy = true;
    state.classroom.revealMode = 'name';
    state.classroom.lastGroup = null;
    renderClassroomTools(); publishSessionState(true);
    await showNameRoulette(sec.students.map(s => cleanStudentName(s.name)), cleanStudentName(chosen.name));
    state.classroom.lastStudent = chosen;
    if (state.classroom.removePicked && !state.classroom.pickedIds.includes(chosen.id)) state.classroom.pickedIds.push(chosen.id);
    state.classroom.busy = false;
    renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true); return chosen;
  }
  function chooseGroupRoll() {
    const n = Math.max(2, Number(state.classroom.groupCount) || 6);
    const mode = ['balanced', 'no-repeat', 'free'].includes(state.classroom.groupRollMode)
      ? state.classroom.groupRollMode
      : (state.classroom.balanced ? 'balanced' : 'free');
    if (mode === 'no-repeat') {
      let used = Array.isArray(state.classroom.rolledGroups)
        ? Array.from(new Set(state.classroom.rolledGroups.map(Number).filter(g => g >= 1 && g <= n)))
        : [];
      if (used.length >= n) used = [];
      const choices = Array.from({ length: n }, (_, index) => index + 1).filter(group => !used.includes(group));
      const group = choices[Math.floor(Math.random() * choices.length)] || 1;
      state.classroom.rolledGroups = [...used, group];
      return group;
    }
    if (mode === 'free') return 1 + Math.floor(Math.random() * n);
    const counts = Array(n).fill(0);
    Object.values(state.classroom.assignments).forEach(g => { if (g >= 1 && g <= n) counts[g - 1]++; });
    const min = Math.min(...counts);
    const choices = counts.map((count, index) => count === min ? index + 1 : null).filter(Boolean);
    return choices[Math.floor(Math.random() * choices.length)];
  }
  async function rollGroupDice() {
    if (state.classroom.busy) return;
    const student = state.classroom.lastStudent || null;
    const group = chooseGroupRoll();
    state.classroom.busy = true;
    state.classroom.revealMode = 'group';
    state.classroom.lastGroup = null;
    renderClassroomTools(); publishSessionState(true);
    await showRollingGroupDice(group);
    state.classroom.lastGroup = group;
    if (student) state.classroom.assignments[student.id] = group;
    state.classroom.busy = false;
    renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true); return group;
  }
  function resetClassroomRound() { if (state.classroom.busy) return; state.classroom.pickedIds=[]; state.classroom.rolledGroups=[]; state.classroom.assignments={}; state.classroom.lastStudent=null; state.classroom.lastGroup=null; state.classroom.revealMode=''; renderClassroomTools(); scheduleClassroomSave(); publishSessionState(true); }
  function getClassroomPresentationLayer() {
    let layer = document.getElementById('classroomPresentationLayer');
    if (layer) return layer;
    layer = document.createElement('div');
    layer.id = 'classroomPresentationLayer';
    layer.className = 'classroom-presentation-layer hidden';
    layer.innerHTML = `
      <div class="classroom-stage-aurora aurora-a"></div>
      <div class="classroom-stage-aurora aurora-b"></div>
      <div class="classroom-stage-grid"></div>
      <div class="classroom-stage-stars"></div>
      <button class="classroom-stage-close" type="button" aria-label="Exit classroom presentation">×</button>
      <div class="classroom-stage-content">
        <div class="classroom-stage-kicker">CLASSROOM RANDOMIZER</div>
        <div class="classroom-stage-section">Select a section</div>
        <div class="classroom-stage-main">
          <div class="classroom-stage-icon">✦</div>
          <h1>Ready for the next pick</h1>
          <p>Use the phone remote or the classroom controls.</p>
        </div>
      </div>`;
    document.body.appendChild(layer);
    layer.querySelector('.classroom-stage-close').addEventListener('click', closeClassroomPresentationMode);
    return layer;
  }
  function renderClassroomPresentation() {
    const layer = document.getElementById('classroomPresentationLayer');
    if (!layer) return;
    const sec = activeClassSection();
    const sectionEl = layer.querySelector('.classroom-stage-section');
    const main = layer.querySelector('.classroom-stage-main');
    if (sectionEl) sectionEl.textContent = sec ? sec.name : 'Select a section';
    if (!main) return;
    if (state.classroom.revealMode === 'name' && state.classroom.lastStudent) {
      const shownName = cleanStudentName(state.classroom.lastStudent.name);
      main.innerHTML = `<div class="classroom-stage-label">SELECTED STUDENT</div><h1 class="classroom-single-line-name" style="${classroomNameStyle(shownName, 'stage')}">${escapeHtml(shownName)}</h1><p>Ready to roll a group.</p>`;
    } else if (state.classroom.revealMode === 'group' && state.classroom.lastGroup) {
      main.innerHTML = `<div class="classroom-stage-label">GROUP RESULT</div><div class="classroom-stage-final-die">${state.classroom.lastGroup}</div><h2>GROUP ${state.classroom.lastGroup}</h2>`;
    } else {
      main.innerHTML = '<div class="classroom-stage-icon">✦</div><h1>Ready for the next pick</h1><p>Use the phone remote or the classroom controls.</p>';
    }
  }
  async function openClassroomPresentationMode() {
    const layer = getClassroomPresentationLayer();
    renderClassroomPresentation();
    layer.classList.remove('hidden');
    hideModal(els.classroomModal);
    try { if (!document.fullscreenElement && layer.requestFullscreen) await layer.requestFullscreen(); } catch (e) { console.warn('Classroom fullscreen was blocked', e); }
  }
  function closeClassroomPresentationMode() {
    const layer = document.getElementById('classroomPresentationLayer');
    if (document.fullscreenElement === layer) document.exitFullscreen().catch(() => {});
    if (layer) layer.classList.add('hidden');
  }
  document.addEventListener('fullscreenchange', () => {
    const layer = document.getElementById('classroomPresentationLayer');
    if (layer && !document.fullscreenElement && !layer.classList.contains('hidden')) layer.classList.add('hidden');
  });
  function getClassroomRevealLayer(mode) {
    const fullscreenHost = document.fullscreenElement || document.body;
    let layer = document.getElementById('classroomRevealLayer');
    if (!layer) { layer = document.createElement('div'); layer.id = 'classroomRevealLayer'; }
    if (layer.parentElement !== fullscreenHost) fullscreenHost.appendChild(layer);
    const sparkles = Array.from({length:18}, (_,i)=>`<i style="--i:${i};--x:${8+Math.random()*84}%;--y:${8+Math.random()*80}%;--d:${(Math.random()*1.3).toFixed(2)}s"></i>`).join('');
    layer.className = `classroom-reveal-layer show ${mode}-reveal`;
    layer.innerHTML = `<div class="classroom-reveal-bg"></div><div class="classroom-reveal-orbit orbit-one"></div><div class="classroom-reveal-orbit orbit-two"></div><div class="classroom-reveal-sparkles">${sparkles}</div><canvas class="classroom-reveal-confetti"></canvas><div class="classroom-reveal-card"></div>`;
    return layer;
  }
  function wait(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
  function playClassroomTone(kind = 'tick', step = 0) {
    try {
      primePresentationAudio();
      const ctx = state.audioContext;
      if (!ctx) return;
      if (ctx.state === 'suspended') ctx.resume().catch(() => {});
      const now = ctx.currentTime + 0.01;
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      const filter = ctx.createBiquadFilter();
      const compressor = ctx.createDynamicsCompressor();
      filter.type = 'lowpass';
      filter.frequency.setValueAtTime(kind === 'dice' ? 2600 : 3400, now);
      try {
        compressor.threshold.setValueAtTime(-18, now);
        compressor.knee.setValueAtTime(20, now);
        compressor.ratio.setValueAtTime(5, now);
        compressor.attack.setValueAtTime(0.003, now);
        compressor.release.setValueAtTime(0.18, now);
      } catch (error) {}
      osc.type = kind === 'dice' ? 'square' : kind === 'land' ? 'triangle' : 'sine';
      const base = kind === 'select' ? 720 : kind === 'land' ? 260 : kind === 'dice' ? 150 + (step % 5) * 34 : 500 + (step % 7) * 34;
      osc.frequency.setValueAtTime(base, now);
      if (kind === 'select') osc.frequency.exponentialRampToValueAtTime(1220, now + 0.24);
      if (kind === 'land') osc.frequency.exponentialRampToValueAtTime(95, now + 0.34);
      const peak = kind === 'select' ? 0.32 : kind === 'land' ? 0.36 : kind === 'dice' ? 0.13 : 0.1;
      gain.gain.setValueAtTime(0.0001, now);
      gain.gain.exponentialRampToValueAtTime(peak, now + 0.008);
      gain.gain.exponentialRampToValueAtTime(0.0001, now + (kind === 'select' ? 0.32 : kind === 'land' ? 0.42 : 0.085));
      osc.connect(filter).connect(gain).connect(compressor).connect(ctx.destination);
      osc.start(now);
      osc.stop(now + (kind === 'select' ? 0.35 : kind === 'land' ? 0.45 : 0.105));
    } catch (error) {}
  }
  function fitClassroomNames(root) {
    if (!root) return;
    const applyFit = () => {
      const rootWidth = Math.max(320, root.clientWidth || root.getBoundingClientRect().width || window.innerWidth * 0.9);
      root.querySelectorAll('.roulette-name,.classroom-single-line-name').forEach((node) => {
        const holder = node.closest('.name-roulette-track') || node.closest('.classroom-stage-main') || node.parentElement;
        const available = Math.max(260, Math.min(rootWidth - 28, (holder && (holder.clientWidth || holder.getBoundingClientRect().width)) || rootWidth) - 20);
        const isFinal = node.classList.contains('roulette-final') || node.classList.contains('classroom-single-line-name');
        const isCenter = !!node.closest('.row-current,.row-next,.row-final-center');
        const baseSize = isFinal ? Math.min(76, Math.max(42, available / 14)) : isCenter ? Math.min(54, Math.max(32, available / 22)) : Math.min(38, Math.max(26, available / 30));
        node.style.fontSize = `${baseSize}px`;
        node.style.transform = '';
        node.style.transformOrigin = 'center center';
        node.style.display = 'inline-block';
        node.style.width = 'max-content';
        const naturalWidth = node.scrollWidth || node.getBoundingClientRect().width;
        if (naturalWidth > available) {
          const scale = Math.max(isFinal ? 0.68 : 0.74, available / naturalWidth);
          node.style.transform = `scaleX(${scale})`;
        }
      });
    };
    requestAnimationFrame(() => requestAnimationFrame(applyFit));
  }
  async function showNameRoulette(allNames, finalName) {
    const layer = getClassroomRevealLayer('name');
    const card = layer.querySelector('.classroom-reveal-card');
    card.innerHTML = `<div class="classroom-reveal-badge">RANDOM NAME</div><div class="name-roulette-window"><div class="roulette-window-sheen"></div><div class="roulette-window-fade top"></div><div class="roulette-window-fade bottom"></div><div class="roulette-center-band"></div><div class="name-roulette-track"></div></div><div class="classroom-reveal-line"></div><p>Choosing...</p>`;
    const track = card.querySelector('.name-roulette-track');
    const status = card.querySelector('p');
    const safePool = Array.from(new Set((allNames.length ? allNames : [finalName]).map(cleanStudentName).filter(Boolean)));
    if (!safePool.includes(finalName)) safePool.push(finalName);

    // Fast at the start, then progressively slower until it settles.
    // The cubic easing creates a clear roulette-style deceleration instead of
    // running at one speed and suddenly stopping.
    const totalSteps = 36;
    let current = pickRoulettePoolName(safePool, [finalName]) || finalName;
    for (let i = 0; i < totalSteps; i++) {
      const isLastStep = i === totalSteps - 1;
      const next = isLastStep ? finalName : pickRoulettePoolName(safePool, [current, finalName]);
      const progress = i / Math.max(1, totalSteps - 1);
      const eased = progress * progress * progress;
      const delay = Math.round(42 + eased * 500);
      const strip = renderRouletteMovingStrip(track, safePool, current, next);
      fitClassroomNames(track);
      playClassroomTone('tick', i);
      strip.style.setProperty('--roulette-duration', `${Math.max(38, Math.round(delay * 0.82))}ms`);
      void strip.offsetHeight;
      strip.classList.add('spinning');
      await wait(delay);
      current = next;
    }

    renderRouletteFinalStrip(track, safePool, finalName);
    fitClassroomNames(track);
    status.textContent = 'Selected!';
    playClassroomTone('select');
    startClassroomRevealConfetti(layer, false);
    await wait(3000);
    layer.classList.remove('show');
    renderClassroomPresentation();
  }
  function buildDicePipMarkup(value) {
    const n = Math.max(1, Number(value) || 1);
    const classic = {
      1: ['c'],
      2: ['tl', 'br'],
      3: ['tl', 'c', 'br'],
      4: ['tl', 'tr', 'bl', 'br'],
      5: ['tl', 'tr', 'c', 'bl', 'br'],
      6: ['tl', 'tr', 'ml', 'mr', 'bl', 'br']
    };
    if (n <= 6) {
      return `<div class="die-pips classic value-${n}">${classic[n].map(pos => `<i class="pip pip-${pos}"></i>`).join('')}</div>`;
    }
    const cols = n <= 9 ? 3 : 4;
    const rows = Math.ceil(n / cols);
    return `<div class="die-pips pip-grid cols-${cols} rows-${rows}">${Array.from({ length: n }, () => '<i class="pip"></i>').join('')}</div>`;
  }
  function paintDiceFaces(faces, frontValue, maxValue) {
    faces.forEach((face, index) => {
      const faceValue = index === 0 ? Number(frontValue) : 1 + ((Number(frontValue) + index - 1) % maxValue);
      face.innerHTML = buildDicePipMarkup(faceValue);
      face.setAttribute('data-face-value', String(faceValue));
    });
  }
  async function showRollingGroupDice(finalGroup) {
    const layer = getClassroomRevealLayer('group');
    const card = layer.querySelector('.classroom-reveal-card');
    card.innerHTML = `<div class="classroom-reveal-badge">GROUP DICE</div><div class="rolling-dice-scene"><div class="rolling-die rolling-die-cube"><div class="die-face die-front"></div><div class="die-face die-back"></div><div class="die-face die-right"></div><div class="die-face die-left"></div><div class="die-face die-top"></div><div class="die-face die-bottom"></div></div><div class="dice-floor-shadow"></div></div><h3 class="rolling-group-label">ROLLING...</h3>`;
    const die = card.querySelector('.rolling-die');
    const faces = Array.from(die.querySelectorAll('.die-face'));
    const max = Math.max(2, Number(state.classroom.groupCount) || 6);
    const started = performance.now();
    let tick = 0;
    while (performance.now() - started < 4900) {
      const value = 1 + Math.floor(Math.random() * max);
      paintDiceFaces(faces, value, max);
      die.style.setProperty('--rx', `${540 + tick*97}deg`);
      die.style.setProperty('--ry', `${720 + tick*83}deg`);
      die.style.setProperty('--rz', `${90 + tick*37}deg`);
      die.style.setProperty('--jump', `${-18 - (tick % 3) * 9}px`);
      playClassroomTone('dice', tick);
      tick++;
      await wait(Math.min(170, 52 + tick*6));
    }
    paintDiceFaces(faces, finalGroup, max);
    playClassroomTone('land');
    die.classList.add('landed');
    card.querySelector('.rolling-group-label').textContent = `GROUP ${finalGroup}`;
    startClassroomRevealConfetti(layer, true);
    await wait(3000);
    layer.classList.remove('show');
    renderClassroomPresentation();
  }
  function startClassroomRevealConfetti(layer, groupMode) {
    const canvas = layer.querySelector('.classroom-reveal-confetti'); if (!canvas) return;
    const ctx = canvas.getContext('2d'); const host = layer.getBoundingClientRect();
    const dpr = Math.min(window.devicePixelRatio || 1, 2); canvas.width = host.width*dpr; canvas.height = host.height*dpr; ctx.scale(dpr,dpr);
    const colors=['#fde047','#fb7185','#60a5fa','#34d399','#c084fc','#ffffff'];
    const count = groupMode ? 150 : 100;
    const particles=Array.from({length:count},(_,i)=>({x:host.width/2+(Math.random()-.5)*100,y:host.height*.48,vx:(Math.random()-.5)*12,vy:-5-Math.random()*10,g:.16+Math.random()*.12,w:5+Math.random()*8,h:7+Math.random()*13,r:Math.random()*6.28,vr:(Math.random()-.5)*.3,c:colors[i%colors.length],life:1}));
    const start=performance.now();
    function frame(now){ const elapsed=now-start; ctx.clearRect(0,0,host.width,host.height); particles.forEach(p=>{p.x+=p.vx;p.y+=p.vy;p.vy+=p.g;p.r+=p.vr;p.life=Math.max(0,1-elapsed/3000);ctx.save();ctx.globalAlpha=p.life;ctx.translate(p.x,p.y);ctx.rotate(p.r);ctx.fillStyle=p.c;ctx.fillRect(-p.w/2,-p.h/2,p.w,p.h);ctx.restore();}); if(elapsed<3000 && layer.classList.contains('show')) requestAnimationFrame(frame); }
    requestAnimationFrame(frame);
  }
  function startCanvasMagicEffect(layer, type) {
    const canvas=layer.querySelector('canvas'); if(!canvas) return; const ctx=canvas.getContext('2d'); let raf=0; const start=performance.now();
    const resize=()=>{ const d=Math.min(devicePixelRatio||1,2); canvas.width=innerWidth*d; canvas.height=innerHeight*d; canvas.style.width=innerWidth+'px'; canvas.style.height=innerHeight+'px'; ctx.setTransform(d,0,0,d,0,0); }; resize();
    const count=type==='confetti'?150:62; const colors=['#ff4d6d','#ffd166','#06d6a0','#4cc9f0','#8338ec','#ffffff'];
    const parts=Array.from({length:count},(_,i)=> type==='confetti'
      ? {x:Math.random()*innerWidth,y:-20-Math.random()*innerHeight*.35,vx:(Math.random()-.5)*4,vy:3+Math.random()*5,r:3+Math.random()*6,h:7+Math.random()*12,rot:Math.random()*6,vr:(Math.random()-.5)*.25,c:colors[i%colors.length]}
      : {x:Math.random()*innerWidth,y:innerHeight*.62+Math.random()*(innerHeight*.58),r:10+Math.random()*26,vx:(Math.random()-.5)*.42,vy:0.86+Math.random()*1.08,a:.45+Math.random()*.42,wobble:Math.random()*Math.PI*2,wobbleSpeed:.00075+Math.random()*.00105});
    function frame(now){ const t=now-start; ctx.clearRect(0,0,innerWidth,innerHeight); for(const p of parts){ if(type==='confetti'){p.x+=p.vx;p.y+=p.vy;p.vy+=.035;p.rot+=p.vr;ctx.save();ctx.translate(p.x,p.y);ctx.rotate(p.rot);ctx.fillStyle=p.c;ctx.fillRect(-p.r,-p.h/2,p.r*2,p.h);ctx.restore();}else{p.x+=p.vx+Math.sin(t*p.wobbleSpeed+p.wobble)*.24;p.y-=p.vy;const fadeIn=Math.min(1,(innerHeight+80-p.y)/150);const fadeOut=Math.max(0,Math.min(1,(p.y+p.r*2)/150));const alpha=Math.max(0,Math.min(1,fadeIn*fadeOut));ctx.save();ctx.globalAlpha=alpha;ctx.beginPath();ctx.arc(p.x,p.y,p.r,0,Math.PI*2);const g=ctx.createRadialGradient(p.x-p.r*.38,p.y-p.r*.4,1,p.x,p.y,p.r);g.addColorStop(0,'rgba(255,255,255,.96)');g.addColorStop(.18,'rgba(255,255,255,.36)');g.addColorStop(.58,'rgba(125,211,252,.19)');g.addColorStop(.82,'rgba(167,139,250,.12)');g.addColorStop(1,'rgba(59,130,246,.025)');ctx.fillStyle=g;ctx.fill();ctx.strokeStyle=`rgba(255,255,255,${p.a})`;ctx.lineWidth=1.25;ctx.stroke();ctx.beginPath();ctx.arc(p.x-p.r*.28,p.y-p.r*.28,p.r*.22,Math.PI*1.05,Math.PI*1.72);ctx.strokeStyle='rgba(255,255,255,.65)';ctx.lineWidth=1.05;ctx.stroke();ctx.restore();}}
      if(t<(type==='confetti'?3100:8100) && layer.classList.contains('show')) raf=requestAnimationFrame(frame); }
    raf=requestAnimationFrame(frame);
  }

  function registerServiceWorker() {
    if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }
})();
