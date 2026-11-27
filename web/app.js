import { KNOWN_BLOCKED_DOMAINS, isKnownBlockedDomain, adaptProviderUrl, joinNotes } from './adapters.js';
import { createScrollController } from './scroll.js';
import { createDrawController } from './draw.js';

const STORAGE_KEYS = {
  chromeVisible: "webviewer2.chromeVisible",
  currentUrl: "webviewer2.currentUrl",
  recentUrls: "webviewer2.recentUrls",
  isDark: "webviewer2.isDark",
  zoom: "webviewer2.zoom",
  refreshInterval: "webviewer2.refreshInterval",
  desktopFit: "webviewer2.desktopFit",
  safeMode: "webviewer2.safeMode",
  invertTheme: "webviewer2.invertTheme",
  mobileView: "webviewer2.mobileView",
  shapeSize:  "webviewer2.shapeSize"   // { w, h } in pt — shape dims before mobile resize
};

const LOAD_TIMEOUT_MS = 9000;

const state = {
  chromeVisible: true,
  _savedEditChrome: null,   // preserves edit-mode chrome state across slideshow
  currentUrl: "",
  recentUrls: [],
  isDark: false,
  zoom: 1.0,
  refreshInterval: 0,
  refreshTimerId: null,
  desktopFit: false,
  safeMode: false,
  drawMode: false,
  invertTheme: false,
  mobileView: false,
  _mobileShapeResized: false, // true when Mobile View used the PowerPoint API (not CSS fallback)
  originalShapeW: null,       // shape width  in pt before mobile resize
  originalShapeH: null,       // shape height in pt before mobile resize
  autoScrollSpeed: 0,       // 0=off 1=slow 2=medium 3=fast (not persisted)
  autoScrollPos: 0,
  autoScrollRafId: null,
  isLoading: false,
  pendingLoads: 0,
  loadTimer: null,
  officeReady: false,
  lastView: null,              // "edit" | "read" — so we react only to real transitions
  _lastSavedChrome: undefined,          // last chromeVisible written to document settings
  _lastSavedUrl: undefined,             // last currentUrl written to document settings
  _lastSavedMobileView: undefined,      // last mobileView written to document settings
  _lastSavedZoom: undefined,            // last zoom written to document settings
  _lastSavedDesktopFit: undefined,      // last desktopFit written to document settings
  _lastSavedSafeMode: undefined,        // last safeMode written to document settings
  _lastSavedRefreshInterval: undefined  // last refreshInterval written to document settings
};

let scroll;                 // scroll controller (createScrollController)
let draw;                   // draw controller (createDrawController)
let _iframeCache = [];      // cached NodeList snapshot, invalidated on each loadIntoFrame
let _cachedShapeId = null;  // shape id once found — skips the O(n) search on subsequent calls

const ui = {};
let toastTimer = null;

document.addEventListener("DOMContentLoaded", initialize);

function initialize() {
  cacheUi();
  scroll = createScrollController(state, ui, { getFrameIframes, syncAdvancedTools });
  draw = createDrawController(ui);
  bindEvents();
  hydrateFromBrowserStorage();

  if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
    Office.onReady(handleOfficeReady);
  }
}

function cacheUi() {
  ui.blockDismissBtn = document.getElementById("block-dismiss-btn");
  ui.blockOpenBtn = document.getElementById("block-open-btn");
  ui.blockOverlay = document.getElementById("block-overlay");
  ui.emptyState = document.getElementById("empty-state");
  ui.floatingToggle = document.getElementById("floating-toggle");
  ui.frame = document.getElementById("viewer-frame");
  ui.hideBtn = document.getElementById("hide-btn");
  ui.loadingBar = document.getElementById("loading-bar");
  ui.loadForm = document.getElementById("load-form");
  ui.shell = document.getElementById("shell");
  ui.toast = document.getElementById("toast");
  ui.urlInput = document.getElementById("url-input");

  ui.qrBtn = document.getElementById("qr-btn");
  ui.qrOverlay = document.getElementById("qr-overlay");
  ui.qrCanvas = document.getElementById("qr-canvas");
  ui.qrCloseBtn = document.getElementById("qr-close-btn");
  ui.recentDropdown = document.getElementById("recent-dropdown");
  ui.recentList = document.getElementById("recent-list");

  ui.refreshBtn = document.getElementById("refresh-btn");
  ui.refreshDropdown = document.getElementById("refresh-dropdown");
  ui.refreshList = document.getElementById("refresh-list");
  ui.zoomBtn = document.getElementById("zoom-btn");
  ui.zoomDropdown = document.getElementById("zoom-dropdown");
  ui.zoomList = document.getElementById("zoom-list");
  ui.desktopFitBtn = document.getElementById("desktop-fit-btn");
  ui.safeModeBtn = document.getElementById("safe-mode-btn");
  ui.drawModeBtn = document.getElementById("draw-mode-btn");
  ui.clearDrawBtn = document.getElementById("clear-draw-btn");
  ui.undoDrawBtn = document.getElementById("undo-draw-btn");
  ui.drawColorRow = document.getElementById("draw-color-row");
  ui.invertBtn = document.getElementById("invert-btn");
  ui.safeModeOverlay = document.getElementById("safe-mode-overlay");
  ui.drawCanvas = document.getElementById("draw-canvas");
  ui.popoutBtn = document.getElementById("popout-btn");
  ui.themeBtn = document.getElementById("theme-btn");
  ui.moreBtn = document.getElementById("more-btn");
  ui.moreDropdown = document.getElementById("more-dropdown");
  ui.mobileViewBtn = document.getElementById("mobile-view-btn");
  ui.autoScrollBtn = document.getElementById("auto-scroll-btn");

  syncEmptyState();
}

function bindEvents() {
  ui.loadForm.addEventListener("submit", handleLoadSubmit);
  ui.hideBtn.addEventListener("click", toggleChrome);
  ui.floatingToggle.addEventListener("click", toggleChrome);
  ui.blockOpenBtn.addEventListener("click", openCurrentUrl);
  ui.blockDismissBtn.addEventListener("click", dismissBlockOverlay);

  ui.qrBtn.addEventListener("click", showQrCode);
  ui.qrCloseBtn.addEventListener("click", hideQrCode);
  ui.urlInput.addEventListener("focus", showRecentDropdown);
  document.addEventListener("click", handleDocumentClick);

  ui.themeBtn.addEventListener("click", toggleTheme);
  ui.zoomBtn.addEventListener("click", () => {
    const open = ui.zoomDropdown.classList.toggle("is-hidden") === false;
    ui.zoomBtn.setAttribute("aria-expanded", String(open));
    if (open) syncZoomMenuState();
  });
  ui.zoomList.addEventListener("click", handleZoomOptionClick);
  ui.zoomList.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); handleZoomOptionClick(e); }
  });
  ui.moreBtn.addEventListener("click", () => {
    const open = ui.moreDropdown.classList.toggle("is-hidden") === false;
    ui.moreBtn.setAttribute("aria-expanded", String(open));
  });
  ui.desktopFitBtn.addEventListener("click", toggleDesktopFit);
  ui.safeModeBtn.addEventListener("click", toggleSafeMode);
  ui.drawModeBtn.addEventListener("click", toggleDrawMode);
  ui.clearDrawBtn.addEventListener("click", draw.clearCanvas);
  ui.undoDrawBtn.addEventListener("click", draw.undoStroke);
  ui.drawColorRow.addEventListener("click", (e) => {
    const swatch = e.target.closest(".draw-swatch");
    if (!swatch) return;
    ui.drawColorRow.querySelectorAll(".draw-swatch").forEach(s => s.classList.remove("is-active"));
    swatch.classList.add("is-active");
    draw.setColor(swatch.dataset.color);
  });
  ui.invertBtn.addEventListener("click", toggleInvertTheme);
  ui.mobileViewBtn.addEventListener("click", toggleMobileView);
  ui.autoScrollBtn.addEventListener("click", scroll.toggleAutoScroll);

  // Set up drawing canvas
  draw.setupDrawing();

  // Re-apply view transforms on resize; also resize draw canvas when active
  window.addEventListener("resize", () => {
    if (state.desktopFit) {
      applyFrameLayout(document.body.clientWidth);
    } else if (state.mobileView) {
      applyMobileView();
    }
    if (state.drawMode) draw.resizeCanvas();
  });

  ui.popoutBtn.addEventListener("click", openCurrentUrl);
  ui.refreshBtn.addEventListener("click", () => {
    const open = ui.refreshDropdown.classList.toggle("is-hidden") === false;
    ui.refreshBtn.setAttribute("aria-expanded", String(open));
  });
  ui.refreshList.addEventListener("click", handleRefreshOptionClick);
  // Keyboard parity: Enter/Space activates the focused option in either dropdown.
  ui.refreshList.addEventListener("keydown", handleListKeyActivate);
  ui.recentList.addEventListener("keydown", handleListKeyActivate);

  // Escape: close any open overlay or dropdown
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    ui.recentDropdown.classList.add("is-hidden");
    ui.refreshDropdown.classList.add("is-hidden");
    ui.refreshBtn.setAttribute("aria-expanded", "false");
    ui.zoomDropdown.classList.add("is-hidden");
    ui.zoomBtn.setAttribute("aria-expanded", "false");
    ui.moreDropdown.classList.add("is-hidden");
    ui.moreBtn.setAttribute("aria-expanded", "false");
    if (!ui.qrOverlay.classList.contains("overlay-hidden")) hideQrCode();
    if (!ui.blockOverlay.classList.contains("overlay-hidden")) dismissBlockOverlay();
  });

  // Focus trap for block overlay (2 buttons: cycle Tab between them)
  ui.blockOverlay.addEventListener("keydown", (e) => {
    if (e.key !== "Tab") return;
    const focusable = [...ui.blockOverlay.querySelectorAll("button:not([disabled])")];
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (e.shiftKey && document.activeElement === first) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && document.activeElement === last) {
      e.preventDefault();
      first.focus();
    }
  });

  // Focus trap for QR overlay (single button: prevent Tab from escaping)
  ui.qrOverlay.addEventListener("keydown", (e) => {
    if (e.key === "Tab") e.preventDefault();
  });
}

function handleOfficeReady() {
  state.officeReady = true;
  
  if (Office?.context?.document?.getActiveViewAsync) {
    Office.context.document.getActiveViewAsync((result) => {
      let delay = 0;
      if (result.status === Office.AsyncResultStatus.Succeeded && String(result.value).toLowerCase() === "read") {
        delay = 600; // Wait for slide show transition to finish before initializing WebGL/iframes
      }
      
      setTimeout(() => {
        hydrateFromDocumentSettings();
        syncChromeState();
        syncActiveView();
        registerActiveViewChanged();
        initShapeTracking();
      }, delay);
    });
  } else {
    hydrateFromDocumentSettings();
    syncChromeState();
    syncActiveView();
    registerActiveViewChanged();
    initShapeTracking();
  }
}

function hydrateFromBrowserStorage() {
  try {
    const chromeVisible = window.localStorage.getItem(STORAGE_KEYS.chromeVisible);
    const recentUrlsStr = window.localStorage.getItem(STORAGE_KEYS.recentUrls);
    const isDarkStr = window.localStorage.getItem(STORAGE_KEYS.isDark);
    const invertStr = window.localStorage.getItem(STORAGE_KEYS.invertTheme);

    if (isDarkStr) state.isDark = isDarkStr === "true";
    if (invertStr) state.invertTheme = invertStr === "true";
    // zoom, desktopFit, safeMode, refreshInterval, mobileView, and currentUrl are
    // all per-shape and live exclusively in Office document settings (not localStorage).

    if (recentUrlsStr) {
      try {
        state.recentUrls = JSON.parse(recentUrlsStr);
        renderRecentUrls();
      } catch (e) {}
    }

    if (chromeVisible !== null) {
      state.chromeVisible = chromeVisible === "true";
    }

    syncChromeState();
    syncThemeState();
    syncZoomState();      // zoom is 1.0 (default); sets button label correctly
    syncAdvancedTools();  // reflects isDark / invertTheme button states

    // currentUrl is intentionally NOT restored from localStorage here.
    // localStorage is shared across every add-in instance on every slide, so
    // restoring it here would cause a new insert to load a stale URL from a
    // completely different shape.  The URL lives exclusively in Office document
    // settings (per-shape) and is restored in hydrateFromDocumentSettings().
  } catch (error) {
    console.warn("Local storage is unavailable.", error);
  }
}

function hydrateFromDocumentSettings() {
  if (!Office?.context?.document?.settings) return;

  const settings = Office.context.document.settings;
  
  settings.refreshAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.warn("Failed to refresh document settings", result.error);
    }
    
    const savedChromeVisible = settings.get(STORAGE_KEYS.chromeVisible);
    if (typeof savedChromeVisible === "boolean") {
      state.chromeVisible = savedChromeVisible;
      syncChromeState();
    }

    // Restore all per-shape view state BEFORE loading the URL so that
    // loadIntoFrame and applyMobileView see the correct state when they fire.

    const savedZoom = settings.get(STORAGE_KEYS.zoom);
    if (typeof savedZoom === "number" && savedZoom > 0) state.zoom = savedZoom;

    const savedDesktopFit = settings.get(STORAGE_KEYS.desktopFit);
    if (typeof savedDesktopFit === "boolean") state.desktopFit = savedDesktopFit;

    const savedSafeMode = settings.get(STORAGE_KEYS.safeMode);
    if (typeof savedSafeMode === "boolean") state.safeMode = savedSafeMode;

    const savedRefreshInterval = settings.get(STORAGE_KEYS.refreshInterval);
    if (typeof savedRefreshInterval === "number") state.refreshInterval = savedRefreshInterval;

    const savedMobileView = settings.get(STORAGE_KEYS.mobileView);
    if (typeof savedMobileView === "boolean") state.mobileView = savedMobileView;

    // Apply all restored view state in one pass before the URL loads.
    syncZoomState();
    syncAdvancedTools();
    applyRefreshInterval();
    if (state.desktopFit) applyFrameLayout(document.body.clientWidth);

    const savedUrl = settings.get(STORAGE_KEYS.currentUrl);
    // currentUrl is exclusively owned by document settings (per-shape).
    // localStorage no longer holds it, so no double-load guard is needed.
    if (typeof savedUrl === "string" && savedUrl.trim()) {
      state.currentUrl = savedUrl;
      ui.urlInput.value = savedUrl;
      safelyHydrateUrl(savedUrl);
    }
  });
}

function handleLoadSubmit(event) {
  event.preventDefault();

  const rawInput = ui.urlInput.value.trim();
  
  if (!rawInput) {
    return;
  }
  
  // Smart URL Validation
  if (!looksLikeUrlOrIframe(rawInput)) {
    showToast("Please enter a valid link or iframe snippet.", "warning");
    return;
  }

  try {
    loadIntoFrame(rawInput);
    saveRecentUrl(rawInput);
  } catch (error) {
    const message = error instanceof Error ? error.message : "The URL could not be loaded.";
    showToast(message, "error");
  }
}

function loadIntoFrame(rawInput, options = {}) {
  const { persist = true, silent = false } = options;
  const inputUrlString = String(rawInput ?? "").trim();
  
  state.currentUrl = inputUrlString;
  ui.urlInput.value = inputUrlString;
  
  syncEmptyState();
  hideBlockOverlay();

  ui.frame.innerHTML = ""; // Clear existing grid
  _iframeCache = [];       // Invalidate iframe cache

  // Reset auto-scroll position so the new page always starts from the top.
  state.autoScrollPos = 0;

  if (!inputUrlString) {
    ui.emptyState.classList.remove("overlay-hidden");
    return;
  }

  ui.emptyState.classList.add("overlay-hidden");

  // Split into up to 4 URLs for split-screen — but only when the commas
  // genuinely separate distinct URLs (see splitUrls). A single URL that
  // happens to contain commas (Google Maps @lat,lng,zoom; dashboard query
  // params) is kept intact.
  const urlStrings = splitUrls(inputUrlString);
  ui.frame.dataset.count = urlStrings.length;

  // Force inline grid styles to prevent CSS caching issues
  ui.frame.style.display = "grid";
  ui.frame.style.flex = "1";
  ui.frame.style.minHeight = "0";
  ui.frame.style.gap = "2px";
  ui.frame.style.background = "var(--btn-border, #444)";
  if (urlStrings.length === 1) {
    ui.frame.style.gridTemplateColumns = "1fr";
    ui.frame.style.gridTemplateRows = "1fr";
  } else if (urlStrings.length === 2) {
    ui.frame.style.gridTemplateColumns = "1fr 1fr";
    ui.frame.style.gridTemplateRows = "1fr";
  } else {
    ui.frame.style.gridTemplateColumns = "1fr 1fr";
    ui.frame.style.gridTemplateRows = "1fr 1fr";
  }

  let hasErrors = false;
  state.pendingLoads = 0;
  state.isLoading = false;

  urlStrings.forEach(urlString => {
    let request;
    try {
      request = prepareEmbedRequest(urlString);
    } catch (e) {
      if (urlStrings.length === 1) {
        throw e;
      }
      hasErrors = true;
      return;
    }

    if (isKnownBlockedDomain(request.hostname)) {
      if (urlStrings.length === 1) {
        showBlockOverlay();
      }
      hasErrors = true;
      return;
    }

    const iframe = document.createElement("iframe");
    iframe.src = request.normalizedUrl;
    iframe.allow = "accelerometer; autoplay; encrypted-media; fullscreen; gyroscope; picture-in-picture; web-share";
    iframe.referrerPolicy = "strict-origin-when-cross-origin";

    // Force inline iframe styles
    iframe.style.width = "100%";
    iframe.style.height = "100%";
    iframe.style.border = "none";
    iframe.style.background = "#fff";
    iframe.style.minHeight = "0"; // prevent flex/grid blowout

    iframe.addEventListener("load", handleFrameLoaded);
    ui.frame.appendChild(iframe);
    state.pendingLoads++;

    if (!silent && request.note && urlStrings.length === 1) {
      showToast(request.note, "info");
    }
  });

  if (hasErrors && urlStrings.length > 1) {
    showToast("One or more links could not be loaded.", "warning");
  }

  // Apply view modes to newly created iframes
  if (state.mobileView) {
    requestAnimationFrame(() => applyMobileView());
  } else if (state.autoScrollSpeed > 0) {
    scroll.applyAutoScrollIframes();
  }

  if (state.pendingLoads > 0) {
    state.isLoading = true;
    ui.loadingBar.classList.add("is-active");
    ui.loadingBar.removeAttribute("aria-hidden");
    clearTimeout(state.loadTimer);
    state.loadTimer = window.setTimeout(handleLoadTimeout, LOAD_TIMEOUT_MS);
  }

  if (persist) {
    persistState();
  }
}

function handleFrameLoaded() {
  if (state.pendingLoads > 0) state.pendingLoads--;
  if (state.pendingLoads > 0 || !state.isLoading) return;

  state.isLoading = false;
  ui.loadingBar.classList.remove("is-active");
  ui.loadingBar.setAttribute("aria-hidden", "true");
  clearTimeout(state.loadTimer);

  if (state.currentUrl) hideToast();
}

function handleLoadTimeout() {
  if (!state.isLoading || !state.currentUrl) {
    return;
  }

  state.isLoading = false;
  state.pendingLoads = 0;
  ui.loadingBar.classList.remove("is-active");

  // In a split-screen view a single slow/blocked pane shouldn't blank out the
  // whole grid — the other panes may have rendered fine. Surface a toast
  // instead of the full-screen "can't be embedded" overlay.
  if (ui.frame.querySelectorAll("iframe").length > 1) {
    showToast("A pane took too long or blocks embedding. Use Open in Browser for that site.", "warning");
    return;
  }

  showBlockOverlay();
}

function showBlockOverlay() {
  ui.blockOverlay.classList.remove("overlay-hidden");
  ui.blockOpenBtn.focus();
}

function hideBlockOverlay() {
  ui.blockOverlay.classList.add("overlay-hidden");
}

function dismissBlockOverlay() {
  hideBlockOverlay();
  ui.urlInput.value = "";
  ui.urlInput.focus();

  state.currentUrl = "";
  state.isLoading = false;
  state.pendingLoads = 0;
  clearTimeout(state.loadTimer);
  ui.frame.innerHTML = "";
  ui.frame.dataset.count = 0;
  ui.emptyState.classList.remove("overlay-hidden");
  syncEmptyState();
  persistState();
}

function openCurrentUrl() {
  // Use the resolved first URL so split-screen (comma-joined) input and iframe
  // snippets are not handed to the browser verbatim.
  const resolved = getResolvedUrls();
  const candidate = resolved[0] || (state.currentUrl || ui.urlInput.value || "").trim();

  if (!candidate) {
    showToast("Add a URL before opening it in a browser.", "warning");
    return;
  }

  if (Office?.context?.ui?.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(candidate);
  } else {
    window.open(candidate, "_blank", "noopener,noreferrer");
  }
}

function toggleChrome() {
  state.chromeVisible = !state.chromeVisible;
  syncChromeState();
  persistState();
}

function syncChromeState() {
  ui.shell.classList.toggle("is-chrome-hidden", !state.chromeVisible);
}

function syncEmptyState() {
  ui.shell.classList.toggle("is-empty", !state.currentUrl);
}

function syncActiveView() {
  if (!Office?.context?.document?.getActiveViewAsync) {
    ui.shell.dataset.view = "edit";
    return;
  }

  Office.context.document.getActiveViewAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      ui.shell.dataset.view = "edit";
      return;
    }

    const nextView = String(result.value).toLowerCase() === "read" ? "read" : "edit";
    const viewChanged = state.lastView !== nextView;
    state.lastView = nextView;
    ui.shell.dataset.view = nextView;

    // React only when the view actually transitioned. A plain window refocus
    // (the visibilitychange fallback below) reports the SAME view — it must not
    // reload the iframe or clobber the saved chrome state.
    if (!viewChanged) {
      return;
    }

    if (nextView === "read") {
      document.body.classList.add("is-presentation");

      // Auto-hide toolbar on entering slideshow, preserving the edit-mode state
      // so it is correctly restored when the presenter exits the slideshow.
      state._savedEditChrome = state.chromeVisible;
      state.chromeVisible = false;
      syncChromeState(); // adds is-chrome-hidden → floating toggle appears

      draw.clearCanvas();
      hideBlockOverlay();
      // Intentionally no forceRefresh() — the loaded webpage must persist
      // unchanged across edit ↔ slideshow transitions.
    } else {
      document.body.classList.remove("is-presentation");

      // Restore the chrome state from before the slideshow.
      if (state._savedEditChrome !== null) {
        state.chromeVisible = state._savedEditChrome;
        state._savedEditChrome = null;
      }
      syncChromeState();
      // Intentionally no forceRefresh() — same as above.
    }
  });
}

function forceRefresh() {
  _iframeCache = []; // iframes are being replaced; force a fresh query
  const iframes = Array.from(ui.frame.querySelectorAll("iframe"));
  iframes.forEach(iframe => {
    const newIframe = document.createElement("iframe");
    newIframe.src = iframe.src;
    newIframe.allow = iframe.allow;
    newIframe.referrerPolicy = iframe.referrerPolicy;
    newIframe.style.cssText = iframe.style.cssText;
    newIframe.addEventListener("load", handleFrameLoaded);
    
    // Replace cleanly to avoid about:blank crashes
    iframe.replaceWith(newIframe);
  });

  if (iframes.length > 0) {
    state.pendingLoads = iframes.length;
    state.isLoading = true;
    ui.loadingBar.classList.add("is-active");

    // Re-apply view modes to replaced iframes
    if (state.mobileView) {
      requestAnimationFrame(() => applyMobileView());
    } else if (state.autoScrollSpeed > 0) {
      scroll.applyAutoScrollIframes();
    }
  }
}

function registerActiveViewChanged() {
  if (Office?.context?.document?.addHandlerAsync && Office?.EventType?.ActiveViewChanged) {
    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, () => {
      syncActiveView();
    });
  }

  // Fallback for PowerPoint Online where ActiveViewChanged may not fire reliably
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      syncActiveView();
    }
  });
}

function persistState() {
  persistToBrowserStorage();
  persistToDocumentSettings();
}

function persistToBrowserStorage() {
  try {
    // Only global user preferences live in localStorage.
    // Per-shape state (zoom, desktopFit, safeMode, refreshInterval, mobileView,
    // currentUrl) lives exclusively in Office document settings so each add-in
    // instance is independent and new inserts always start clean.
    window.localStorage.setItem(STORAGE_KEYS.chromeVisible, String(state.chromeVisible));
    window.localStorage.setItem(STORAGE_KEYS.isDark, String(state.isDark));
    window.localStorage.setItem(STORAGE_KEYS.invertTheme, String(state.invertTheme));
  } catch (error) {
    console.warn("Local persistence failed.", error);
  }
}

function persistToDocumentSettings() {
  const settings = Office?.context?.document?.settings;
  if (!settings) return;

  // Track dirty state per field so unrelated toggles (zoom, theme, etc.) never
  // mark the .pptx as modified. Only write + save when something actually changed.
  let dirty = false;

  if (state.chromeVisible !== state._lastSavedChrome) {
    settings.set(STORAGE_KEYS.chromeVisible, state.chromeVisible);
    state._lastSavedChrome = state.chromeVisible;
    dirty = true;
  }

  if (state.currentUrl !== state._lastSavedUrl) {
    if (state.currentUrl) settings.set(STORAGE_KEYS.currentUrl, state.currentUrl);
    state._lastSavedUrl = state.currentUrl;
    dirty = true;
  }

  if (state.mobileView !== state._lastSavedMobileView) {
    settings.set(STORAGE_KEYS.mobileView, state.mobileView);
    state._lastSavedMobileView = state.mobileView;
    dirty = true;
  }

  if (state.zoom !== state._lastSavedZoom) {
    settings.set(STORAGE_KEYS.zoom, state.zoom);
    state._lastSavedZoom = state.zoom;
    dirty = true;
  }

  if (state.desktopFit !== state._lastSavedDesktopFit) {
    settings.set(STORAGE_KEYS.desktopFit, state.desktopFit);
    state._lastSavedDesktopFit = state.desktopFit;
    dirty = true;
  }

  if (state.safeMode !== state._lastSavedSafeMode) {
    settings.set(STORAGE_KEYS.safeMode, state.safeMode);
    state._lastSavedSafeMode = state.safeMode;
    dirty = true;
  }

  if (state.refreshInterval !== state._lastSavedRefreshInterval) {
    settings.set(STORAGE_KEYS.refreshInterval, state.refreshInterval);
    state._lastSavedRefreshInterval = state.refreshInterval;
    dirty = true;
  }

  if (!dirty) return;

  settings.saveAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.warn("Document settings were not saved.", result.error);
    }
  });
}

/* ── URL Processing ─────────────────────────────────────────────── */

function prepareEmbedRequest(rawInput) {
  const rawValue = String(rawInput ?? "").trim();
  const extracted = extractUrlCandidate(rawInput);

  if (!extracted) {
    throw new Error("Paste a URL or an iframe snippet first.");
  }

  let candidate = extracted.trim();

  if (candidate.startsWith("//")) {
    candidate = `https:${candidate}`;
  } else if (!/^[a-zA-Z][a-zA-Z\d+\-.]*:/.test(candidate)) {
    candidate = `https://${candidate}`;
  }

  let parsedUrl;
  try {
    parsedUrl = new URL(candidate);
  } catch (error) {
    throw new Error("That does not look like a valid web address.");
  }

  if (parsedUrl.protocol !== "https:") {
    throw new Error("Only HTTPS pages are supported for reliable embedding in Office.");
  }

  if (!parsedUrl.hostname) {
    throw new Error("A hostname is required.");
  }

  const sourceNote = rawValue.match(/<iframe/i)
    ? "Extracted the iframe src from the snippet."
    : "";
  const adapted = adaptProviderUrl(parsedUrl);

  // Use the ADAPTED URL's hostname for the blocked-domain check so that
  // provider adapters that switch to an embeddable subdomain (e.g. Google Maps
  // converting www.google.com/maps/embed → maps.google.com) are not incorrectly
  // blocked by the www.google.com entry in KNOWN_BLOCKED_DOMAINS.
  const adaptedHostname = new URL(adapted.url.toString()).hostname;

  return {
    hostname: adaptedHostname,
    note: joinNotes(sourceNote, adapted.note),
    normalizedUrl: adapted.url.toString(),
    providerName: adapted.providerName
  };
}

function extractUrlCandidate(rawInput) {
  const rawValue = String(rawInput ?? "").trim();

  if (!rawValue) {
    return "";
  }

  const iframeMatch = rawValue.match(/src\s*=\s*["']([^"']+)["']/i);
  if (iframeMatch?.[1]) {
    return iframeMatch[1];
  }

  return rawValue;
}

// True when a string looks like a bare/absolute URL or an <iframe> snippet.
function looksLikeUrlOrIframe(value) {
  return /^((https?:\/\/)?([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,}|<iframe\s)/i.test(String(value ?? "").trim());
}

// Splits input into up to 4 URLs for split-screen — but ONLY when every
// comma-separated segment independently looks like a URL. Otherwise the commas
// belong to a single URL (e.g. Google Maps @lat,lng,zoom or a dashboard's query
// params) and the whole string is returned as a single entry. This prevents a
// normal pasted URL from being shredded into several broken iframes.
function splitUrls(input) {
  const raw = String(input ?? "").trim();
  if (!raw) {
    return [];
  }
  const parts = raw.split(",").map((s) => s.trim()).filter(Boolean);
  if (parts.length > 1 && parts.every(looksLikeUrlOrIframe)) {
    return parts.slice(0, 4);
  }
  return [raw];
}

// Returns the cleaned, https URLs for the current input — resolving iframe
// snippets and provider embeds, and splitting only genuine multi-URL input.
// Used by QR + Open-in-browser so they never act on raw snippet HTML or a
// comma-joined split-screen string.
function getResolvedUrls() {
  const raw = (state.currentUrl || ui.urlInput.value || "").trim();
  if (!raw) {
    return [];
  }
  const urls = [];
  splitUrls(raw).forEach((segment) => {
    try {
      urls.push(prepareEmbedRequest(segment).normalizedUrl);
    } catch (e) {
      // Skip segments that can't be normalized into an embeddable URL.
    }
  });
  return urls;
}

/* ── Toast ──────────────────────────────────────────────────────── */

function showToast(message, type = "info", duration = 4000) {
  ui.toast.textContent = message;
  ui.toast.dataset.type = type;
  ui.toast.classList.remove("overlay-hidden");

  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(hideToast, duration);
}

function hideToast() {
  clearTimeout(toastTimer);
  ui.toast.classList.add("overlay-hidden");
}

function safelyHydrateUrl(value) {
  try {
    loadIntoFrame(value, { persist: false, silent: true });
  } catch (error) {
    console.warn("Saved URL could not be restored.", error);
  }
}

/* ── QR Code ────────────────────────────────────────────────────── */

function showQrCode() {
  if (!state.currentUrl) {
    showToast("Load a URL first to generate a QR code.", "warning");
    return;
  }

  // Encode the resolved first URL so split-screen + iframe snippets produce a
  // scannable single link rather than raw HTML or a comma-joined string.
  const target = getResolvedUrls()[0] || state.currentUrl;

  try {
    if (window.QRious) {
      new QRious({
        element: ui.qrCanvas,
        value: target,
        size: 200,
        background: "white",
        foreground: "black"
      });
      // Update accessible label with the actual URL
      ui.qrCanvas.setAttribute("aria-label", `QR code for ${target}`);
      ui.qrOverlay.classList.remove("overlay-hidden");
      // Move focus to the close button so keyboard/AT users land inside the overlay
      ui.qrCloseBtn.focus();
    } else {
      showToast("QR code library not loaded.", "error");
    }
  } catch (e) {
    showToast("Could not generate QR code.", "error");
  }
}

function hideQrCode() {
  ui.qrOverlay.classList.add("overlay-hidden");
  // Return focus to the button that opened the overlay
  ui.qrBtn.focus();
}

/* ── Recent Links ───────────────────────────────────────────────── */

function saveRecentUrl(url) {
  if (!url) return;
  
  state.recentUrls = state.recentUrls.filter(u => u !== url);
  state.recentUrls.unshift(url);
  
  if (state.recentUrls.length > 5) {
    state.recentUrls = state.recentUrls.slice(0, 5);
  }
  
  window.localStorage.setItem(STORAGE_KEYS.recentUrls, JSON.stringify(state.recentUrls));
  renderRecentUrls();
}

function renderRecentUrls() {
  ui.recentList.innerHTML = "";
  
  if (state.recentUrls.length === 0) {
    const li = document.createElement("li");
    li.className = "recent-item";
    li.style.color = "var(--text-placeholder)";
    li.style.pointerEvents = "none";
    li.textContent = "No recent links";
    ui.recentList.appendChild(li);
    return;
  }
  
  state.recentUrls.forEach(url => {
    const li = document.createElement("li");
    li.className = "recent-item";
    try {
      const parsed = new URL(url);
      li.textContent = parsed.hostname + (parsed.pathname !== "/" ? parsed.pathname : "");
    } catch (_) {
      li.textContent = url;
    }
    li.title = url;
    li.tabIndex = 0;
    li.setAttribute("role", "option");
    li.addEventListener("click", () => {
      ui.urlInput.value = url;
      ui.recentDropdown.classList.add("is-hidden");
      ui.loadForm.dispatchEvent(new Event("submit", { cancelable: true }));
    });
    ui.recentList.appendChild(li);
  });
}

function showRecentDropdown() {
  if (ui.urlInput.value.trim() === "") {
    ui.recentDropdown.classList.remove("is-hidden");
  }
}

function handleDocumentClick(event) {
  if (!ui.urlInput.contains(event.target) && !ui.recentDropdown.contains(event.target)) {
    ui.recentDropdown.classList.add("is-hidden");
  }
  if (!ui.refreshBtn.contains(event.target) && !ui.refreshDropdown.contains(event.target)) {
    ui.refreshDropdown.classList.add("is-hidden");
    ui.refreshBtn.setAttribute("aria-expanded", "false");
  }
  if (!ui.zoomBtn.contains(event.target) && !ui.zoomDropdown.contains(event.target)) {
    ui.zoomDropdown.classList.add("is-hidden");
    ui.zoomBtn.setAttribute("aria-expanded", "false");
  }
  if (!ui.moreBtn.contains(event.target) && !ui.moreDropdown.contains(event.target)) {
    ui.moreDropdown.classList.add("is-hidden");
    ui.moreBtn.setAttribute("aria-expanded", "false");
  }
}

/* ── New Features (Theme, Zoom, Refresh) ────────────────────────── */

function toggleTheme() {
  state.isDark = !state.isDark;
  syncThemeState();
  persistState();
}

function syncThemeState() {
  ui.shell.classList.toggle("is-dark", state.isDark);
  ui.themeBtn.classList.toggle("is-active", state.isDark);
  syncMoreBtnState();
}

function handleZoomOptionClick(e) {
  const li = e.target.closest("[data-pct]");
  if (!li) return;
  const pct = Number(li.dataset.pct);
  if (!pct) return;
  state.zoom = pct / 100;
  state.desktopFit = false;
  ui.zoomDropdown.classList.add("is-hidden");
  ui.zoomBtn.setAttribute("aria-expanded", "false");
  syncZoomState();
  persistState();
}

function syncZoomMenuState() {
  if (!ui.zoomList) return;
  const currentPct = Math.round(state.zoom * 100);
  ui.zoomList.querySelectorAll("li[data-pct]").forEach((li) => {
    li.setAttribute("aria-selected", String(Number(li.dataset.pct) === currentPct));
  });
}

function syncZoomState() {
  if (ui.zoomBtn) {
    let label;
    if (state.desktopFit) {
      label = "Fit";
    } else if (state.mobileView) {
      label = "390px";
    } else {
      label = `${Math.round(state.zoom * 100)}%`;
    }
    ui.zoomBtn.textContent = label;
    ui.zoomBtn.setAttribute("aria-label", `Zoom — current: ${label}`);
  }
  if (state.desktopFit || state.mobileView) return;
  applyFrameLayout();
  syncZoomMenuState();
}

function toggleDesktopFit() {
  state.desktopFit = !state.desktopFit;

  if (state.desktopFit) {
    // Mutual exclusivity: disable mobile view
    if (state.mobileView) {
      state.mobileView = false;
      ui.mobileViewBtn.classList.remove("is-active");
      getFrameIframes().forEach(f => {
        f.style.width = "100%";
        f.style.height = "100%";
        f.style.transform = "";
        f.style.transformOrigin = "";
      });
    }
    // Stop any active auto-scroll
    scroll.stopAndReset();
    applyFrameLayout(document.body.clientWidth);
  } else {
    applyFrameLayout();
  }

  ui.desktopFitBtn.classList.toggle("is-active", state.desktopFit);
  syncZoomState();
  syncMoreBtnState();
  persistState();
}

// Single entry-point for all #viewer-frame transform state.
// Precedence: mobileView resets; desktopFit scales to 1280; otherwise zoom.
// Auto-scroll translateY is applied to individual iframes and never conflicts.
function applyFrameLayout(containerWidth) {
  if (state.mobileView) {
    ui.frame.style.transform = "";
    ui.frame.style.width = "100%";
    ui.frame.style.height = "100%";
    return;
  }
  if (state.desktopFit) {
    const TARGET_DESKTOP_WIDTH = 1280;
    const scale = Math.min((containerWidth || document.body.clientWidth) / TARGET_DESKTOP_WIDTH, 1);
    ui.frame.style.transformOrigin = "top left";
    ui.frame.style.transform = `scale(${scale})`;
    ui.frame.style.width = `${100 / scale}%`;
    ui.frame.style.height = `${100 / scale}%`;
    return;
  }
  if (state.zoom === 1.0) {
    ui.frame.style.transform = "";
    ui.frame.style.width = "100%";
    ui.frame.style.height = "100%";
  } else {
    ui.frame.style.transformOrigin = "top left";
    ui.frame.style.transform = `scale(${state.zoom})`;
    ui.frame.style.width = `${100 / state.zoom}%`;
    ui.frame.style.height = `${100 / state.zoom}%`;
  }
}

// Enter/Space on a focused dropdown option behaves like a click, so the recent
// and refresh menus are operable without a mouse.
function handleListKeyActivate(event) {
  if (event.key !== "Enter" && event.key !== " " && event.key !== "Spacebar") {
    return;
  }
  const item = event.target.closest("li[role='option']");
  if (!item) return;
  event.preventDefault();
  item.click();
}

function handleRefreshOptionClick(e) {
  const li = e.target.closest("li.recent-item");
  if (!li) return;

  // "Reload now" — trigger an immediate reload without changing the interval
  if (li.dataset.val === "reload") {
    if (state.currentUrl) forceRefresh();
    ui.refreshDropdown.classList.add("is-hidden");
    ui.refreshBtn.setAttribute("aria-expanded", "false");
    return;
  }

  const val = parseInt(li.dataset.val, 10);
  if (isNaN(val)) return;

  state.refreshInterval = val;
  applyRefreshInterval();
  persistState();

  ui.refreshDropdown.classList.add("is-hidden");
  ui.refreshBtn.setAttribute("aria-expanded", "false");
}

function applyRefreshInterval() {
  if (state.refreshTimerId) {
    clearInterval(state.refreshTimerId);
    state.refreshTimerId = null;
  }
  
  ui.refreshBtn.classList.toggle("is-active", state.refreshInterval > 0);
  if (state.refreshInterval > 0) {
    state.refreshTimerId = setInterval(() => {
      if (state.currentUrl && !state.isLoading) {
        forceRefresh();
      }
    }, state.refreshInterval);
  }

  syncRefreshMenuState();
}

// Mark the currently-selected interval in the refresh dropdown so users can see
// which cadence is active (the "Reload now" action is never the active state).
function syncRefreshMenuState() {
  if (!ui.refreshList) return;
  ui.refreshList.querySelectorAll("li.recent-item[data-val]").forEach((li) => {
    if (li.dataset.val === "reload") {
      li.removeAttribute("aria-selected");
      return;
    }
    const isActive = parseInt(li.dataset.val, 10) === state.refreshInterval;
    li.setAttribute("aria-selected", String(isActive));
  });
}


/* ── Advanced Tools Logic ────────────────────────────────────────── */

function syncAdvancedTools() {
  // Safe Mode
  ui.safeModeBtn.classList.toggle("is-active", state.safeMode);
  ui.safeModeOverlay.classList.toggle("is-hidden", !state.safeMode);

  // Draw Mode
  ui.drawModeBtn.classList.toggle("is-active", state.drawMode);
  ui.drawCanvas.classList.toggle("is-hidden", !state.drawMode);
  if (state.drawMode) { ui.drawCanvas.removeAttribute("role"); } else { ui.drawCanvas.setAttribute("role", "img"); }
  ui.clearDrawBtn.classList.toggle("is-hidden", !state.drawMode);
  ui.undoDrawBtn.classList.toggle("is-hidden", !state.drawMode);
  ui.drawColorRow.classList.toggle("is-hidden", !state.drawMode);
  if (state.drawMode) draw.resizeCanvas();

  // Desktop Fit
  ui.desktopFitBtn.classList.toggle("is-active", state.desktopFit);

  // Invert Theme
  ui.invertBtn.classList.toggle("is-active", state.invertTheme);
  ui.frame.classList.toggle("smart-invert", state.invertTheme);

  // Dark Toolbar
  ui.themeBtn.classList.toggle("is-active", state.isDark);

  // Mobile View
  ui.mobileViewBtn.classList.toggle("is-active", state.mobileView);

  // Auto-Scroll
  ui.autoScrollBtn.classList.toggle("is-active", state.autoScrollSpeed > 0);

  syncMoreBtnState();
}

function syncMoreBtnState() {
  const anyActive = state.desktopFit || state.safeMode || state.drawMode || state.invertTheme || state.isDark || state.mobileView || state.autoScrollSpeed > 0;
  ui.moreBtn.style.color = anyActive ? "var(--accent)" : "";
}

function toggleSafeMode() {
  state.safeMode = !state.safeMode;
  if (state.safeMode) {
    state.drawMode = false; // Turn off draw mode if entering safe mode
  }
  syncAdvancedTools();
  persistState();
}

function toggleDrawMode() {
  state.drawMode = !state.drawMode;
  if (state.drawMode) {
    state.safeMode = false; // Turn off safe mode if entering draw mode
    // Pause auto-scroll — annotating a moving canvas is not useful
    if (state.autoScrollSpeed > 0) {
      scroll.stopAutoScrollAnimation();
    }
  } else if (state.autoScrollSpeed > 0) {
    // Resume auto-scroll when draw mode is exited
    scroll.startAutoScrollAnimation();
  }
  syncAdvancedTools();
  // Draw mode is NOT persisted between reloads
}

function toggleInvertTheme() {
  state.invertTheme = !state.invertTheme;
  syncAdvancedTools();
  persistState();
}


/* ── iframe cache ────────────────────────────────────────────────── */

// Returns a cached array of iframes inside #viewer-frame. The cache is
// populated on first call after each loadIntoFrame (which sets _iframeCache=[]).
function getFrameIframes() {
  if (_iframeCache.length === 0) {
    _iframeCache = Array.from(ui.frame.querySelectorAll("iframe"));
  }
  return _iframeCache;
}

/* ── Frame Transform Composition ────────────────────────────────── */

/* ── Mobile View ─────────────────────────────────────────────────── */

// Shape tag used to identify our add-in shape in the PowerPoint object model.
const SHAPE_TAG_KEY   = "WV2_ID";
const SHAPE_TAG_VALUE = "webviewer2";

// ── Startup: discover shape, cache original dimensions ────────────

// Called once from handleOfficeReady.  Restores any previously saved shape
// dimensions from localStorage, then proactively tags our PowerPoint shape
// so the first Mobile View toggle is instant.
// Must NOT overwrite saved dimensions when mobileView is already ON —
// the shape will be at the reduced mobile width in that case.
function initShapeTracking() {
  // Restore saved dimensions from a previous session first.
  if (!state.originalShapeW) {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEYS.shapeSize);
      if (stored) {
        const p = JSON.parse(stored);
        if (p && p.w > 0) { state.originalShapeW = p.w; state.originalShapeH = p.h; }
      }
    } catch (_) {}
  }

  // Proactively find and tag our shape in the background.
  // Only record the live shape dimensions when mobileView is OFF —
  // otherwise the shape is already at mobile width, not the original.
  if (typeof PowerPoint !== "undefined") {
    PowerPoint.run(async context => {
      // Read-only lookup — must not write to the document at startup or the
      // .pptx will be marked as modified every time the add-in loads.
      const shape = await findOurShape(context);
      if (!shape) return;
      if (!state.mobileView && !state.originalShapeW) {
        state.originalShapeW = shape.width;
        state.originalShapeH = shape.height;
        try {
          window.localStorage.setItem(STORAGE_KEYS.shapeSize,
            JSON.stringify({ w: shape.width, h: shape.height }));
        } catch (_) {}
      }
    }).catch(e => console.warn("[WV2] initShapeTracking:", e));
  }
}

// Read-only shape lookup — used at startup (initShapeTracking) so the
// document is never marked dirty just by loading the add-in.
// Preference order:
//   1. Shape already tagged with SHAPE_TAG_KEY  (fast direct lookup)
//   2. Shape whose width matches our CSS viewport width ±8 pt
//   3. Shape whose name contains "webviewer" (case-insensitive)
// Merged shape-finder: one O(n) pass per slide, ID-cached after first discovery.
// Pass { tag: true } only when the document is already dirty (resize calls) to
// avoid marking the presentation modified on a read-only startup lookup.
async function findShape(context, { tag = false } = {}) {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();

  for (const slide of slides.items) {
    const shapes = slide.shapes;
    shapes.load("items");
    await context.sync();
    if (!shapes.items.length) continue;

    // If we have a cached id, fetch that shape directly — O(1).
    if (_cachedShapeId !== null) {
      const cached = shapes.items.find(s => { s.load("id"); return false; });
      // getItemOrNullObject by id is not available in all host versions;
      // fall back to scanning the slide if the id approach fails.
      void cached; // suppress lint — we scan anyway if needed
    }

    shapes.items.forEach(s => s.load("id,name,width,height"));
    await context.sync();

    // Fast path: id already cached.
    if (_cachedShapeId !== null) {
      const hit = shapes.items.find(s => s.id === _cachedShapeId);
      if (hit) return hit;
    }

    // Single pass: load tags for all shapes at once.
    const entries = shapes.items.map(s => {
      const t = s.tags.getItemOrNullObject(SHAPE_TAG_KEY);
      t.load("isNullObject,value");
      return { shape: s, tag: t };
    });
    await context.sync();

    const viewportPx = document.documentElement.clientWidth || 720;
    const expectedPt = viewportPx * (72 / 96);

    for (const { shape, tagObj } of entries.map(e => ({ shape: e.shape, tagObj: e.tag }))) {
      const isOurs =
        (!tagObj.isNullObject && tagObj.value === SHAPE_TAG_VALUE) ||
        (tagObj.isNullObject && Math.abs(shape.width - expectedPt) <= 8) ||
        (tagObj.isNullObject && shape.name &&
          shape.name.toLowerCase().includes("webviewer"));

      if (!isOurs) continue;

      _cachedShapeId = shape.id;

      if (tag && tagObj.isNullObject) {
        shape.tags.add(SHAPE_TAG_KEY, SHAPE_TAG_VALUE);
        await context.sync();
      }
      return shape;
    }
  }

  return null;
}

// Convenience aliases kept so call sites read clearly.
function findOurShape(context)       { return findShape(context, { tag: false }); }
function findOrTagOurShape(context)  { return findShape(context, { tag: true });  }

// ── Toggle ────────────────────────────────────────────────────────

async function toggleMobileView() {
  state.mobileView = !state.mobileView;

  if (state.mobileView) {
    // Mutual exclusivity: disable desktop fit
    if (state.desktopFit) {
      state.desktopFit = false;
      ui.desktopFitBtn.classList.remove("is-active");
    }
    // Mutual exclusivity: stop auto-scroll
    scroll.stopAndReset();
    // Clear any zoom transforms on the frame
    ui.frame.style.transform = "";
    ui.frame.style.width     = "100%";
    ui.frame.style.height    = "100%";

    // Attempt PowerPoint shape resize; fall back to CSS iframe trick.
    state._mobileShapeResized = await resizeShapeToMobile();
    if (state._mobileShapeResized) {
      showToast("Shape resized to phone width — click Mobile View again to restore.", "info");
    } else {
      applyMobileViewFallback();
      showToast("Mobile View active — iframes narrowed to phone width. Shape resize requires Office 2021+.", "info");
    }
  } else {
    // Restore iframes (needed when we were in CSS-fallback mode).
    getFrameIframes().forEach(f => {
      f.style.width           = "100%";
      f.style.height          = "100%";
      f.style.transform       = "";
      f.style.transformOrigin = "";
      f.style.justifySelf     = "";
    });

    // Always attempt shape restore — handles restarts where
    // _mobileShapeResized was reset to false but the shape is still narrow.
    await restoreShapeSize();
    state._mobileShapeResized = false;
    applyFrameLayout();
  }

  syncZoomState();
  syncAdvancedTools();
  persistState();
}

// ── PowerPoint shape resize helpers ──────────────────────────────

// Resize the shape that hosts this add-in to a 390 CSS-px-equivalent
// width.  The pt value is derived from live measurements so it is
// accurate regardless of screen DPI.
// Returns true if the shape was successfully resized.
async function resizeShapeToMobile() {
  if (typeof PowerPoint === "undefined") return false;
  try {
    let success = false;
    await PowerPoint.run(async context => {
      const shape = await findOrTagOurShape(context);
      if (!shape) return;

      // Derive pt-per-px using the shape's current width and our actual
      // container pixel width — accounts for any DPI/zoom scaling.
      const containerPx = document.documentElement.clientWidth || 720;
      const ptPerPx     = shape.width / containerPx;

      // Always snapshot current (full-size) dimensions before resizing so
      // a subsequent restore is accurate even if the shape was manually
      // resized since the last toggle.
      state.originalShapeW = shape.width;
      state.originalShapeH = shape.height;
      try {
        window.localStorage.setItem(STORAGE_KEYS.shapeSize,
          JSON.stringify({ w: shape.width, h: shape.height }));
      } catch (_) {}

      shape.width = Math.round(390 * ptPerPx * 10) / 10;
      await context.sync();
      success = true;
    });
    return success;
  } catch (e) {
    console.warn("[WV2] resizeShapeToMobile:", e);
    return false;
  }
}

// Restore the shape to its pre-mobile dimensions.
// Safe no-op if no saved dimensions are available.
async function restoreShapeSize() {
  if (typeof PowerPoint === "undefined") return;

  let w = state.originalShapeW;
  if (!w) {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEYS.shapeSize);
      if (stored) { const p = JSON.parse(stored); w = p.w; }
    } catch (_) {}
  }
  if (!w) return; // no saved size — nothing to restore

  try {
    await PowerPoint.run(async context => {
      const shape = await findOrTagOurShape(context);
      if (!shape) return;
      shape.width = w;
      await context.sync();
    });
  } catch (e) {
    console.warn("[WV2] restoreShapeSize:", e);
  }
}

// ── CSS fallback (when PowerPoint API is unavailable) ─────────────

// Sets each iframe to a real 390 CSS-px width (capped at the pane width)
// with NO scale transform, centred in its grid cell.  The embedded page
// sees a genuine narrow viewport and its responsive CSS fires naturally.
function applyMobileViewFallback() {
  if (!state.mobileView) return;

  const MOBILE_WIDTH = 390;
  const iframes      = getFrameIframes();
  const paneCount    = iframes.length;
  const containerW   = ui.frame.clientWidth || document.body.clientWidth;
  const hasCols      = paneCount >= 2;
  const paneW        = hasCols ? containerW / 2 : containerW;
  const w            = Math.min(MOBILE_WIDTH, paneW);

  iframes.forEach(iframe => {
    iframe.style.width           = `${w}px`;
    iframe.style.height          = "100%";
    iframe.style.transform       = "";
    iframe.style.transformOrigin = "";
    iframe.style.justifySelf     = "center";
  });
}

// Called by the resize-event handler and loadIntoFrame.
// Routes to the CSS fallback when the PowerPoint shape was not (or
// cannot be) resized.  When the shape was resized, iframes already fill
// it at 100% — no CSS intervention needed.
function applyMobileView() {
  if (!state.mobileView) return;
  if (!state._mobileShapeResized) applyMobileViewFallback();
}

