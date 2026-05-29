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
  mobileView: "webviewer2.mobileView"
};

const LOAD_TIMEOUT_MS = 9000;

// Sites we know will absolutely refuse to load in an iframe due to X-Frame-Options or CSP frame-ancestors.
// This saves the user 9 seconds of waiting for a timeout.
const KNOWN_BLOCKED_DOMAINS = [
  "google", "google.com", "www.google.com",
  "facebook", "facebook.com", "www.facebook.com",
  "twitter", "twitter.com", "x.com",
  "github", "github.com",
  "reddit", "reddit.com", "www.reddit.com",
  "amazon", "amazon.com", "www.amazon.com",
  "apple", "apple.com", "www.apple.com",
  "microsoft", "microsoft.com", "www.microsoft.com",
  "linkedin", "linkedin.com", "www.linkedin.com",
  "instagram", "instagram.com", "www.instagram.com"
];

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
  autoScrollSpeed: 0,       // 0=off 1=slow 2=medium 3=fast (not persisted)
  autoScrollPos: 0,
  autoScrollRafId: null,
  isLoading: false,
  pendingLoads: 0,
  loadTimer: null,
  officeReady: false
};

let _autoScrollLast = null; // rAF timestamp tracking
let _iframeCache = [];      // cached NodeList snapshot, invalidated on each loadIntoFrame

const ui = {};
let toastTimer = null;

document.addEventListener("DOMContentLoaded", initialize);

function initialize() {
  cacheUi();
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
  ui.zoomInBtn = document.getElementById("zoom-in-btn");
  ui.zoomOutBtn = document.getElementById("zoom-out-btn");
  ui.zoomLevelBtn = document.getElementById("zoom-level-btn");
  ui.desktopFitBtn = document.getElementById("desktop-fit-btn");
  ui.safeModeBtn = document.getElementById("safe-mode-btn");
  ui.drawModeBtn = document.getElementById("draw-mode-btn");
  ui.clearDrawBtn = document.getElementById("clear-draw-btn");
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
  ui.zoomInBtn.addEventListener("click", zoomIn);
  ui.zoomOutBtn.addEventListener("click", zoomOut);
  ui.zoomLevelBtn.addEventListener("click", resetZoom);
  ui.moreBtn.addEventListener("click", () => {
    const open = ui.moreDropdown.classList.toggle("is-hidden") === false;
    ui.moreBtn.setAttribute("aria-expanded", String(open));
  });
  ui.desktopFitBtn.addEventListener("click", toggleDesktopFit);
  ui.safeModeBtn.addEventListener("click", toggleSafeMode);
  ui.drawModeBtn.addEventListener("click", toggleDrawMode);
  ui.clearDrawBtn.addEventListener("click", clearCanvas);
  ui.invertBtn.addEventListener("click", toggleInvertTheme);
  ui.mobileViewBtn.addEventListener("click", toggleMobileView);
  ui.autoScrollBtn.addEventListener("click", toggleAutoScroll);

  // Set up drawing canvas
  setupDrawing();

  // Re-apply view transforms on resize
  window.addEventListener("resize", () => {
    if (state.desktopFit) {
      applyDesktopFitScale(document.body.clientWidth);
    } else if (state.mobileView) {
      applyMobileView();
    }
  });

  ui.popoutBtn.addEventListener("click", openCurrentUrl);
  ui.refreshBtn.addEventListener("click", () => {
    const open = ui.refreshDropdown.classList.toggle("is-hidden") === false;
    ui.refreshBtn.setAttribute("aria-expanded", String(open));
  });
  ui.refreshList.addEventListener("click", handleRefreshOptionClick);

  // Escape: close any open overlay or dropdown
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    ui.recentDropdown.classList.add("is-hidden");
    ui.refreshDropdown.classList.add("is-hidden");
    ui.refreshBtn.setAttribute("aria-expanded", "false");
    ui.moreDropdown.classList.add("is-hidden");
    ui.moreBtn.setAttribute("aria-expanded", "false");
    if (!ui.qrOverlay.classList.contains("is-hidden")) hideQrCode();
    if (!ui.blockOverlay.classList.contains("is-hidden")) dismissBlockOverlay();
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
      }, delay);
    });
  } else {
    hydrateFromDocumentSettings();
    syncChromeState();
    syncActiveView();
    registerActiveViewChanged();
  }
}

function hydrateFromBrowserStorage() {
  try {
    const chromeVisible = window.localStorage.getItem(STORAGE_KEYS.chromeVisible);
    const currentUrl = window.localStorage.getItem(STORAGE_KEYS.currentUrl);
    const recentUrlsStr = window.localStorage.getItem(STORAGE_KEYS.recentUrls);
    const isDarkStr = window.localStorage.getItem(STORAGE_KEYS.isDark);
    const zoomStr = window.localStorage.getItem(STORAGE_KEYS.zoom);
    const refreshStr = window.localStorage.getItem(STORAGE_KEYS.refreshInterval);
    const desktopFitStr = window.localStorage.getItem(STORAGE_KEYS.desktopFit);
    const safeModeStr = window.localStorage.getItem(STORAGE_KEYS.safeMode);
    const invertStr = window.localStorage.getItem(STORAGE_KEYS.invertTheme);
    const mobileViewStr = window.localStorage.getItem(STORAGE_KEYS.mobileView);

    if (isDarkStr) state.isDark = isDarkStr === "true";
    if (zoomStr) state.zoom = parseFloat(zoomStr) || 1.0;
    if (refreshStr) state.refreshInterval = parseInt(refreshStr, 10) || 0;
    if (desktopFitStr) state.desktopFit = desktopFitStr === "true";
    if (safeModeStr) state.safeMode = safeModeStr === "true";
    if (invertStr) state.invertTheme = invertStr === "true";
    if (mobileViewStr) state.mobileView = mobileViewStr === "true";

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
    if (state.desktopFit) {
      applyDesktopFitScale(document.body.clientWidth);
    }
    syncZoomState();
    syncAdvancedTools();
    applyRefreshInterval();

    if (currentUrl) {
      ui.urlInput.value = currentUrl;
      safelyHydrateUrl(currentUrl);
    }
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

    const savedUrl = settings.get(STORAGE_KEYS.currentUrl);
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
  const looksLikeUrl = /^((https?:\/\/)?([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,}|<iframe\s)/i.test(rawInput);
  if (!looksLikeUrl) {
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
    ui.emptyState.classList.remove("is-hidden");
    return;
  }
  
  ui.emptyState.classList.add("is-hidden");

  // Split by comma, up to 4 URLs
  const urlStrings = inputUrlString.split(',').map(s => s.trim()).filter(Boolean).slice(0, 4);
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
    applyAutoScrollIframes();
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
  showBlockOverlay();
}

function isKnownBlockedDomain(hostname) {
  return KNOWN_BLOCKED_DOMAINS.includes(hostname.toLowerCase());
}

function showBlockOverlay() {
  ui.blockOverlay.classList.remove("is-hidden");
}

function hideBlockOverlay() {
  ui.blockOverlay.classList.add("is-hidden");
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
  ui.emptyState.classList.remove("is-hidden");
  syncEmptyState();
  persistState();
}

function openCurrentUrl() {
  const candidate = state.currentUrl || ui.urlInput.value.trim();

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

function syncActiveView(options = {}) {
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
    ui.shell.dataset.view = nextView;
    
    if (nextView === "read") {
      document.body.classList.add("is-presentation");

      // Auto-hide toolbar on entering slideshow, preserving the edit-mode state
      // so it is correctly restored when the presenter exits the slideshow.
      state._savedEditChrome = state.chromeVisible;
      state.chromeVisible = false;
      syncChromeState(); // adds is-chrome-hidden → floating toggle appears

      clearCanvas();
      hideBlockOverlay();

      if (!options.isInitial) {
        forceRefresh();
      }
    } else {
      document.body.classList.remove("is-presentation");

      // Restore the chrome state from before the slideshow.
      if (state._savedEditChrome !== null) {
        state.chromeVisible = state._savedEditChrome;
        state._savedEditChrome = null;
      }
      syncChromeState();

      if (!options.isInitial) {
        forceRefresh();
      }
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
      applyAutoScrollIframes();
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
    window.localStorage.setItem(STORAGE_KEYS.chromeVisible, String(state.chromeVisible));
    window.localStorage.setItem(STORAGE_KEYS.isDark, String(state.isDark));
    window.localStorage.setItem(STORAGE_KEYS.zoom, String(state.zoom));
    window.localStorage.setItem(STORAGE_KEYS.refreshInterval, String(state.refreshInterval));
    window.localStorage.setItem(STORAGE_KEYS.desktopFit, String(state.desktopFit));
    window.localStorage.setItem(STORAGE_KEYS.safeMode, String(state.safeMode));
    window.localStorage.setItem(STORAGE_KEYS.invertTheme, String(state.invertTheme));
    window.localStorage.setItem(STORAGE_KEYS.mobileView, String(state.mobileView));
    if (state.currentUrl) {
      window.localStorage.setItem(STORAGE_KEYS.currentUrl, state.currentUrl);
    }
  } catch (error) {
    console.warn("Local persistence failed.", error);
  }
}

function persistToDocumentSettings() {
  const settings = Office?.context?.document?.settings;

  if (!settings) {
    return;
  }

  settings.set(STORAGE_KEYS.chromeVisible, state.chromeVisible);
  if (state.currentUrl) {
    settings.set(STORAGE_KEYS.currentUrl, state.currentUrl);
  }

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

  return {
    hostname: parsedUrl.hostname,
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

/* ── Provider Adapters ──────────────────────────────────────────── */

function adaptProviderUrl(url) {
  const hostname = normalizeHostname(url.hostname);

  if (hostname === "youtu.be" || hostname.endsWith("youtube.com") || hostname.endsWith("youtube-nocookie.com")) {
    return adaptYouTubeUrl(url);
  }

  if (hostname === "vimeo.com" || hostname === "player.vimeo.com") {
    return adaptVimeoUrl(url);
  }

  if (hostname === "loom.com") {
    return adaptLoomUrl(url);
  }

  if (hostname === "docs.google.com") {
    return adaptGoogleDocsUrl(url);
  }

  if (hostname === "app.powerbi.com") {
    return adaptPowerBiUrl(url);
  }

  if (hostname === "forms.office.com" || hostname === "forms.microsoft.com") {
    return {
      note: "",
      providerName: "Microsoft Forms",
      url
    };
  }

  if (hostname.endsWith(".sharepoint.com") || hostname.endsWith(".sharepoint.cn")) {
    return adaptSharePointUrl(url);
  }

  if (hostname === "figma.com" || hostname === "www.figma.com") {
    return adaptFigmaUrl(url);
  }

  if (hostname === "miro.com") {
    return adaptMiroUrl(url);
  }

  if (hostname === "canva.com" || hostname === "www.canva.com") {
    return adaptCanvaUrl(url);
  }

  if (hostname === "canva.link") {
    throw new Error("Canva short links cannot be embedded directly. Open it in your browser and copy the full link.");
  }

  if (hostname === "airtable.com") {
    return adaptAirtableUrl(url);
  }

  if (hostname === "codepen.io") {
    return adaptCodePenUrl(url);
  }

  if (hostname === "open.spotify.com") {
    return adaptSpotifyUrl(url);
  }

  return {
    note: "",
    providerName: formatProviderName(hostname),
    url
  };
}

function adaptYouTubeUrl(url) {
  const hostname = normalizeHostname(url.hostname);
  const pathParts = url.pathname.split("/").filter(Boolean);
  let videoId = "";

  if (hostname === "youtu.be") {
    videoId = pathParts[0] ?? "";
  } else if (pathParts[0] === "watch") {
    videoId = url.searchParams.get("v") ?? "";
  } else if (["embed", "shorts", "live"].includes(pathParts[0] ?? "")) {
    videoId = pathParts[1] ?? "";
  }

  if (!videoId) {
    return {
      note: "",
      providerName: "YouTube",
      url
    };
  }

  const embedUrl = new URL(`https://www.youtube-nocookie.com/embed/${videoId}`);
  const startSeconds = extractStartSeconds(url.searchParams.get("t") ?? url.searchParams.get("start"));
  const playlistId = url.searchParams.get("list");

  if (startSeconds > 0) {
    embedUrl.searchParams.set("start", String(startSeconds));
  }

  if (playlistId) {
    embedUrl.searchParams.set("list", playlistId);
  }

  const transformed = embedUrl.toString() !== url.toString();

  return {
    note: transformed ? "Converted to YouTube embed URL." : "",
    providerName: "YouTube",
    url: transformed ? embedUrl : url
  };
}

function adaptVimeoUrl(url) {
  const pathParts = url.pathname.split("/").filter(Boolean);
  const numericId =
    pathParts[0] === "video"
      ? pathParts[1]
      : pathParts.find((segment) => /^\d+$/.test(segment));

  if (!numericId) {
    return {
      note: "",
      providerName: "Vimeo",
      url
    };
  }

  const embedUrl = new URL(`https://player.vimeo.com/video/${numericId}`);
  const transformed = embedUrl.toString() !== url.toString();

  return {
    note: transformed ? "Converted to Vimeo embed URL." : "",
    providerName: "Vimeo",
    url: transformed ? embedUrl : url
  };
}

function adaptLoomUrl(url) {
  const pathParts = url.pathname.split("/").filter(Boolean);
  const loomId = pathParts[1] ?? "";

  if (!loomId || !["share", "embed"].includes(pathParts[0] ?? "")) {
    return {
      note: "",
      providerName: "Loom",
      url
    };
  }

  const embedUrl = new URL(`https://www.loom.com/embed/${loomId}`);
  const transformed = embedUrl.toString() !== url.toString();

  return {
    note: transformed ? "Converted to Loom embed URL." : "",
    providerName: "Loom",
    url: transformed ? embedUrl : url
  };
}

function adaptGoogleDocsUrl(url) {
  const pathParts = url.pathname.split("/").filter(Boolean);
  const docType = pathParts[0] ?? "";
  const documentId = pathParts[2] ?? "";

  // URL doesn't have the standard /type/d/ID structure — pass through
  if (pathParts[1] !== "d" || !documentId) {
    return { note: "", providerName: "Google Workspace", url };
  }

  if (docType === "presentation") {
    return {
      note: "Converted to Google Slides embed.",
      providerName: "Google Slides",
      url: new URL(`https://docs.google.com/presentation/d/${documentId}/embed`)
    };
  }

  if (docType === "document") {
    return {
      note: "Converted to Google Docs preview.",
      providerName: "Google Docs",
      url: new URL(`https://docs.google.com/document/d/${documentId}/preview`)
    };
  }

  if (docType === "spreadsheets") {
    return {
      note: "Converted to Google Sheets preview.",
      providerName: "Google Sheets",
      url: new URL(`https://docs.google.com/spreadsheets/d/${documentId}/preview`)
    };
  }

  if (docType === "forms") {
    // ?embedded=true removes the Google header bar so the form fills the iframe
    const embedUrl = new URL(url.toString());
    if (!embedUrl.searchParams.has("embedded")) {
      embedUrl.searchParams.set("embedded", "true");
    }
    const transformed = embedUrl.toString() !== url.toString();
    return {
      note: transformed ? "Converted to Google Forms embed." : "",
      providerName: "Google Forms",
      url: embedUrl
    };
  }

  return { note: "", providerName: "Google Workspace", url };
}

function adaptPowerBiUrl(url) {
  const isAuthoringUrl = /\/groups\/.+\/reports\//i.test(url.pathname) && !/reportEmbed/i.test(url.pathname);

  return {
    note: isAuthoringUrl
      ? "This Power BI link may need a publish-to-web embed URL."
      : "",
    providerName: "Power BI",
    url
  };
}
function adaptSharePointUrl(url) {
  if (url.pathname.includes("_layouts/15/Doc.aspx")) {
    const embedUrl = new URL(url.toString());
    embedUrl.searchParams.set("action", "embedview");
    embedUrl.searchParams.set("wdAllowInteractivity", "True");
    
    const transformed = embedUrl.toString() !== url.toString();
    
    return {
      note: transformed ? "Converted to SharePoint embed view." : "",
      providerName: "SharePoint",
      url: transformed ? embedUrl : url
    };
  }
  
  return {
    note: "",
    providerName: "SharePoint",
    url
  };
}

function adaptFigmaUrl(url) {
  // Only convert design or prototype links that aren't already embeds
  if ((url.pathname.startsWith("/design/") || url.pathname.startsWith("/file/") || url.pathname.startsWith("/proto/")) && url.pathname !== "/embed") {
    const embedUrl = new URL("https://www.figma.com/embed");
    embedUrl.searchParams.set("embed_host", "share");
    embedUrl.searchParams.set("url", url.toString());
    
    return {
      note: "Converted to Figma embed.",
      providerName: "Figma",
      url: embedUrl
    };
  }
  
  return {
    note: "",
    providerName: "Figma",
    url
  };
}

function adaptMiroUrl(url) {
  if (url.pathname.startsWith("/app/board/")) {
    const embedUrl = new URL(url.toString().replace("/app/board/", "/app/live-embed/"));
    return {
      note: "Converted to Miro embed.",
      providerName: "Miro",
      url: embedUrl
    };
  }
  return { note: "", providerName: "Miro", url };
}

function adaptCanvaUrl(url) {
  if (url.pathname.startsWith("/design/")) {
    let embedUrl = new URL(url.toString());
    
    if (url.pathname.endsWith("/edit") || url.pathname.endsWith("/watch")) {
      embedUrl.pathname = url.pathname.replace(/\/(edit|watch)$/, "/view");
    }
    
    if (embedUrl.pathname.endsWith("/view")) {
      embedUrl.searchParams.set("embed", "");
      return {
        note: "Converted to Canva embed.",
        providerName: "Canva",
        url: embedUrl
      };
    }
  }
  return { note: "", providerName: "Canva", url };
}

function adaptAirtableUrl(url) {
  const pathParts = url.pathname.split("/").filter(Boolean);
  if (pathParts[0] && pathParts[0].startsWith("shr")) {
    const embedUrl = new URL(`https://airtable.com/embed/${pathParts[0]}`);
    embedUrl.search = url.search;
    return {
      note: "Converted to Airtable embed.",
      providerName: "Airtable",
      url: embedUrl
    };
  }
  return { note: "", providerName: "Airtable", url };
}

function adaptCodePenUrl(url) {
  if (url.pathname.includes("/pen/")) {
    const embedUrl = new URL(url.toString().replace("/pen/", "/embed/"));
    return {
      note: "Converted to CodePen embed.",
      providerName: "CodePen",
      url: embedUrl
    };
  }
  return { note: "", providerName: "CodePen", url };
}

function adaptSpotifyUrl(url) {
  const pathParts = url.pathname.split("/").filter(Boolean);
  if (pathParts.length >= 2 && !pathParts.includes("embed")) {
    const type = pathParts[0]; // track, album, playlist, episode
    const id = pathParts[1];
    const embedUrl = new URL(`https://open.spotify.com/embed/${type}/${id}`);
    return {
      note: "Converted to Spotify embed.",
      providerName: "Spotify",
      url: embedUrl
    };
  }
  return { note: "", providerName: "Spotify", url };
}

/* ── Utilities ──────────────────────────────────────────────────── */

function normalizeHostname(hostname) {
  return String(hostname ?? "").toLowerCase().replace(/^www\./, "").replace(/^m\./, "");
}

function formatProviderName(hostname) {
  const providerName = hostname.replace(/^www\./, "");
  return providerName ? providerName : "Website";
}

function extractStartSeconds(value) {
  if (!value) {
    return 0;
  }

  if (/^\d+$/.test(value)) {
    return Number.parseInt(value, 10);
  }

  const match = value.match(/^(?:(\d+)h)?(?:(\d+)m)?(?:(\d+)s)?$/i);
  if (!match) {
    return 0;
  }

  const hours = Number.parseInt(match[1] ?? "0", 10);
  const minutes = Number.parseInt(match[2] ?? "0", 10);
  const seconds = Number.parseInt(match[3] ?? "0", 10);
  return (hours * 3600) + (minutes * 60) + seconds;
}

function joinNotes(...notes) {
  return notes.filter(Boolean).join(" ").trim();
}

/* ── Toast ──────────────────────────────────────────────────────── */

function showToast(message, type = "info", duration = 4000) {
  ui.toast.textContent = message;
  ui.toast.dataset.type = type;
  ui.toast.classList.remove("is-hidden");

  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(hideToast, duration);
}

function hideToast() {
  clearTimeout(toastTimer);
  ui.toast.classList.add("is-hidden");
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

  try {
    if (window.QRious) {
      new QRious({
        element: ui.qrCanvas,
        value: state.currentUrl,
        size: 200,
        background: "white",
        foreground: "black"
      });
      // Update accessible label with the actual URL
      ui.qrCanvas.setAttribute("aria-label", `QR code for ${state.currentUrl}`);
      ui.qrOverlay.classList.remove("is-hidden");
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
  ui.qrOverlay.classList.add("is-hidden");
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
    li.textContent = url;
    li.title = url;
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

function zoomIn() {
  state.zoom = Math.min(state.zoom + 0.1, 3.0);
  syncZoomState();
  persistState();
}

function zoomOut() {
  state.zoom = Math.max(state.zoom - 0.1, 0.25);
  syncZoomState();
  persistState();
}

function resetZoom() {
  state.zoom = 1.0;
  syncZoomState();
  persistState();
}

function syncZoomState() {
  if (ui.zoomLevelBtn) {
    let label;
    if (state.desktopFit) {
      label = "Fit";
    } else if (state.mobileView) {
      label = "390px";
    } else {
      label = `${Math.round(state.zoom * 100)}%`;
    }
    ui.zoomLevelBtn.textContent = label;
    // Keep aria-label in sync so screen readers announce the correct value
    ui.zoomLevelBtn.setAttribute("aria-label", `Reset zoom — current: ${label}`);
  }
  if (state.desktopFit || state.mobileView) return;
  applyFrameTransform();
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
    stopAutoScrollAnimation();
    state.autoScrollSpeed = 0;
    state.autoScrollPos = 0;
    const scrollSpan = ui.autoScrollBtn.querySelector("span");
    if (scrollSpan) scrollSpan.textContent = AUTO_SCROLL_LABELS[0];
    resetAutoScrollIframes();
    applyDesktopFitScale(document.body.clientWidth);
  } else {
    ui.frame.style.transform = "";
    ui.frame.style.width = "100%";
    ui.frame.style.height = "100%";
  }

  ui.desktopFitBtn.classList.toggle("is-active", state.desktopFit);
  syncZoomState();
  syncMoreBtnState();
  persistState();
}

function applyDesktopFitScale(containerWidth) {
  const TARGET_DESKTOP_WIDTH = 1280;
  const scale = Math.min(containerWidth / TARGET_DESKTOP_WIDTH, 1);
  ui.frame.style.transformOrigin = "top left";
  ui.frame.style.transform = `scale(${scale})`;
  ui.frame.style.width = `${100 / scale}%`;
  ui.frame.style.height = `${100 / scale}%`;
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
}


/* ── Advanced Tools Logic ────────────────────────────────────────── */

function syncAdvancedTools() {
  // Safe Mode
  ui.safeModeBtn.classList.toggle("is-active", state.safeMode);
  ui.safeModeOverlay.classList.toggle("is-hidden", !state.safeMode);

  // Draw Mode
  ui.drawModeBtn.classList.toggle("is-active", state.drawMode);
  ui.drawCanvas.classList.toggle("is-hidden", !state.drawMode);
  ui.clearDrawBtn.classList.toggle("is-hidden", !state.drawMode);
  if (state.drawMode) resizeCanvas();

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
      stopAutoScrollAnimation();
    }
  } else if (state.autoScrollSpeed > 0) {
    // Resume auto-scroll when draw mode is exited
    startAutoScrollAnimation();
  }
  syncAdvancedTools();
  // Draw mode is NOT persisted between reloads
}

function toggleInvertTheme() {
  state.invertTheme = !state.invertTheme;
  syncAdvancedTools();
  persistState();
}

// Drawing Logic
let ctx, isDrawing = false, lastX = 0, lastY = 0;

function setupDrawing() {
  ctx = ui.drawCanvas.getContext("2d");
  
  window.addEventListener('resize', resizeCanvas);
  
  ui.drawCanvas.addEventListener('mousedown', startDrawing);
  ui.drawCanvas.addEventListener('mousemove', draw);
  ui.drawCanvas.addEventListener('mouseup', stopDrawing);
  ui.drawCanvas.addEventListener('mouseout', stopDrawing);
  
  // Touch support
  ui.drawCanvas.addEventListener('touchstart', handleTouchStart, {passive: false});
  ui.drawCanvas.addEventListener('touchmove', handleTouchMove, {passive: false});
  ui.drawCanvas.addEventListener('touchend', stopDrawing);
}

function resizeCanvas() {
  if (!ui.drawCanvas) return;
  const rect = ui.shell.getBoundingClientRect();
  ui.drawCanvas.width = rect.width;
  ui.drawCanvas.height = rect.height;
  ctx.strokeStyle = "#e05252"; // Red ink
  ctx.lineWidth = 4;
  ctx.lineCap = "round";
  ctx.lineJoin = "round";
}

function getPos(e) {
  const rect = ui.drawCanvas.getBoundingClientRect();
  // Use nullish coalescing so x=0 / y=0 (left/top edge) stays correct
  const touch = e.touches?.[0];
  const clientX = touch ? touch.clientX : e.clientX;
  const clientY = touch ? touch.clientY : e.clientY;
  return {
    x: clientX - rect.left,
    y: clientY - rect.top
  };
}

function startDrawing(e) {
  isDrawing = true;
  const pos = getPos(e);
  lastX = pos.x;
  lastY = pos.y;
}

function draw(e) {
  if (!isDrawing) return;
  e.preventDefault();
  const pos = getPos(e);
  ctx.beginPath();
  ctx.moveTo(lastX, lastY);
  ctx.lineTo(pos.x, pos.y);
  ctx.stroke();
  lastX = pos.x;
  lastY = pos.y;
}

function handleTouchStart(e) {
  e.preventDefault();
  startDrawing(e);
}

function handleTouchMove(e) {
  e.preventDefault();
  draw(e);
}

function stopDrawing() {
  isDrawing = false;
}

function clearCanvas() {
  if (ctx && ui.drawCanvas) {
    ctx.clearRect(0, 0, ui.drawCanvas.width, ui.drawCanvas.height);
  }
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

// Applies zoom scale to #viewer-frame. Auto-scroll translateY is applied
// directly to each <iframe> so the two transforms never conflict.
function applyFrameTransform() {
  if (state.desktopFit || state.mobileView) return;

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

/* ── Mobile View ─────────────────────────────────────────────────── */

function toggleMobileView() {
  state.mobileView = !state.mobileView;

  if (state.mobileView) {
    // Mutual exclusivity: disable desktop fit
    if (state.desktopFit) {
      state.desktopFit = false;
      ui.desktopFitBtn.classList.remove("is-active");
    }
    // Mutual exclusivity: stop auto-scroll
    stopAutoScrollAnimation();
    state.autoScrollSpeed = 0;
    state.autoScrollPos = 0;
    const scrollSpan = ui.autoScrollBtn.querySelector("span");
    if (scrollSpan) scrollSpan.textContent = AUTO_SCROLL_LABELS[0];
    resetAutoScrollIframes();
    // Reset frame-level transforms set by zoom
    ui.frame.style.transform = "";
    ui.frame.style.width = "100%";
    ui.frame.style.height = "100%";
    applyMobileView();
  } else {
    // Restore normal iframe sizing
    getFrameIframes().forEach(f => {
      f.style.width = "100%";
      f.style.height = "100%";
      f.style.transform = "";
      f.style.transformOrigin = "";
    });
    applyFrameTransform();
  }

  syncZoomState();
  syncAdvancedTools();
  persistState();
}

function applyMobileView() {
  if (!state.mobileView) return;

  const MOBILE_WIDTH = 390; // iPhone-class viewport width in CSS px
  const iframes = getFrameIframes();
  const paneCount = iframes.length;
  const containerW = ui.frame.clientWidth || document.body.clientWidth;
  const containerH = ui.frame.clientHeight || document.body.clientHeight;

  // In a 2-column layout (2–4 panes), each cell is half the container width/height.
  const hasRows = paneCount >= 3;
  const hasCols = paneCount >= 2;
  const paneW = hasCols ? containerW / 2 : containerW;
  const paneH = hasRows ? containerH / 2 : containerH;
  const scale = paneW / MOBILE_WIDTH;

  iframes.forEach(iframe => {
    iframe.style.width = `${MOBILE_WIDTH}px`;
    iframe.style.height = `${paneH / scale}px`;
    iframe.style.transformOrigin = "top left";
    iframe.style.transform = `scale(${scale})`;
  });
}

/* ── Auto-Scroll ─────────────────────────────────────────────────── */

// px/sec for each speed level (0=off, 1=slow, 2=medium, 3=fast)
const AUTO_SCROLL_SPEEDS = [0, 20, 50, 120];
const AUTO_SCROLL_LABELS = ["Auto-Scroll", "Slow Scroll", "Medium Scroll", "Fast Scroll"];

function toggleAutoScroll() {
  // Cycle: off → slow → medium → fast → off
  state.autoScrollSpeed = (state.autoScrollSpeed + 1) % 4;

  const span = ui.autoScrollBtn.querySelector("span");
  if (span) span.textContent = AUTO_SCROLL_LABELS[state.autoScrollSpeed];

  if (state.autoScrollSpeed === 0) {
    stopAutoScrollAnimation();
    state.autoScrollPos = 0;
    resetAutoScrollIframes();
  } else {
    applyAutoScrollIframes(); // extend iframe heights before animating
    startAutoScrollAnimation();
  }

  syncAdvancedTools();
}

// Extend each iframe to 3× the container height so there is scrollable
// content to reveal. The translateY in autoScrollTick pans through it.
function applyAutoScrollIframes() {
  const containerH = ui.frame.clientHeight || 400;
  getFrameIframes().forEach(iframe => {
    iframe.style.height = `${containerH * 3}px`;
  });
}

// Restore iframes to their normal full-height sizing.
function resetAutoScrollIframes() {
  getFrameIframes().forEach(iframe => {
    iframe.style.height = "100%";
    iframe.style.transform = "";
    iframe.style.transformOrigin = "";
    iframe.style.opacity = "";
    iframe.style.transition = "";
  });
}

function startAutoScrollAnimation() {
  stopAutoScrollAnimation();
  _autoScrollLast = null;
  state.autoScrollRafId = requestAnimationFrame(autoScrollTick);
}

function stopAutoScrollAnimation() {
  if (state.autoScrollRafId) {
    cancelAnimationFrame(state.autoScrollRafId);
    state.autoScrollRafId = null;
  }
  _autoScrollLast = null;
}

function autoScrollTick(timestamp) {
  if (!_autoScrollLast) _autoScrollLast = timestamp;
  const delta = (timestamp - _autoScrollLast) / 1000; // convert ms → seconds
  _autoScrollLast = timestamp;

  const speed = AUTO_SCROLL_SPEEDS[state.autoScrollSpeed] || 0;
  if (speed > 0) {
    const containerH = ui.frame.clientHeight || 400;
    const maxScroll = containerH * 2; // 3× iframe height − 1× visible = 2× scrollable
    const newPos = (state.autoScrollPos + speed * delta) % maxScroll;
    const wrapped = newPos < state.autoScrollPos; // true when the modulo reset fired
    state.autoScrollPos = newPos;

    const iframes = getFrameIframes();
    iframes.forEach(iframe => {
      iframe.style.transformOrigin = "top left";
      iframe.style.transform = `translateY(-${state.autoScrollPos}px)`;
    });

    // On wrap, fade out instantly then fade back in so the jump to the top
    // reads as an intentional reset rather than a visual glitch.
    if (wrapped) {
      iframes.forEach(f => {
        f.style.transition = "none";
        f.style.opacity = "0";
        requestAnimationFrame(() => {
          f.style.transition = "opacity 0.5s ease";
          f.style.opacity = "1";
        });
      });
    }
  }

  if (state.autoScrollSpeed > 0) {
    state.autoScrollRafId = requestAnimationFrame(autoScrollTick);
  } else {
    state.autoScrollRafId = null;
  }
}
