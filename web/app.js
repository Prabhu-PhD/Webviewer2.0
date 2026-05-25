const STORAGE_KEYS = {
  chromeVisible: "webviewer2.chromeVisible",
  currentUrl: "webviewer2.currentUrl",
  recentUrls: "webviewer2.recentUrls",
  isDark: "webviewer2.isDark",
  zoom: "webviewer2.zoom",
  refreshInterval: "webviewer2.refreshInterval",
  desktopFit: "webviewer2.desktopFit",
  safeMode: "webviewer2.safeMode",
  invertTheme: "webviewer2.invertTheme"
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
  isLoading: false,
  loadTimer: null,
  officeReady: false
};

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
  ui.desktopFitBtn = document.getElementById("desktop-fit-btn");
  ui.safeModeBtn = document.getElementById("safe-mode-btn");
  ui.drawModeBtn = document.getElementById("draw-mode-btn");
  ui.clearDrawBtn = document.getElementById("clear-draw-btn");
  ui.invertBtn = document.getElementById("invert-btn");
  ui.safeModeOverlay = document.getElementById("safe-mode-overlay");
  ui.drawCanvas = document.getElementById("draw-canvas");
  ui.popoutBtn = document.getElementById("popout-btn");
  ui.themeBtn = document.getElementById("theme-btn");

  syncEmptyState();
}

function bindEvents() {
  ui.loadForm.addEventListener("submit", handleLoadSubmit);
  ui.hideBtn.addEventListener("click", toggleChrome);
  ui.floatingToggle.addEventListener("click", toggleChrome);
  ui.frame.addEventListener("load", handleFrameLoaded);
  ui.blockOpenBtn.addEventListener("click", openCurrentUrl);
  ui.blockDismissBtn.addEventListener("click", dismissBlockOverlay);

  ui.qrBtn.addEventListener("click", showQrCode);
  ui.qrCloseBtn.addEventListener("click", hideQrCode);
  ui.urlInput.addEventListener("focus", showRecentDropdown);
  document.addEventListener("click", handleDocumentClick);

  ui.themeBtn.addEventListener("click", toggleTheme);
  ui.zoomInBtn.addEventListener("click", zoomIn);
  ui.zoomOutBtn.addEventListener("click", zoomOut);
  ui.desktopFitBtn.addEventListener("click", toggleDesktopFit);
  ui.safeModeBtn.addEventListener("click", toggleSafeMode);
  ui.drawModeBtn.addEventListener("click", toggleDrawMode);
  ui.clearDrawBtn.addEventListener("click", clearCanvas);
  ui.invertBtn.addEventListener("click", toggleInvertTheme);
  
  // Set up drawing canvas
  setupDrawing();
  ui.popoutBtn.addEventListener("click", openCurrentUrl);
  ui.refreshBtn.addEventListener("click", () => ui.refreshDropdown.classList.toggle("is-hidden"));
  ui.refreshList.addEventListener("click", handleRefreshOptionClick);
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

    if (isDarkStr) state.isDark = isDarkStr === "true";
    if (zoomStr) state.zoom = parseFloat(zoomStr) || 1.0;
    if (refreshStr) state.refreshInterval = parseInt(refreshStr, 10) || 0;
    if (desktopFitStr) state.desktopFit = desktopFitStr === "true";
    if (safeModeStr) state.safeMode = safeModeStr === "true";
    if (invertStr) state.invertTheme = invertStr === "true";

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
      ui.desktopFitBtn.style.color = "var(--accent)";
      applyDesktopFitScale(document.body.clientWidth, document.body.clientHeight);
    } else {
      syncZoomState();
    syncAdvancedTools();
    }
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
  const settings = Office?.context?.document?.settings;

  if (!settings) {
    return;
  }

  const savedUrl = settings.get(STORAGE_KEYS.currentUrl);
  const savedChromeVisible = settings.get(STORAGE_KEYS.chromeVisible);

  if (typeof savedChromeVisible === "boolean") {
    state.chromeVisible = savedChromeVisible;
    syncChromeState();
  }

  if (typeof savedUrl === "string" && savedUrl.trim()) {
    ui.urlInput.value = savedUrl;
    safelyHydrateUrl(savedUrl);
  }
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
  state.isLoading = true;
  ui.loadingBar.classList.add("is-active");

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
    iframe.allow = "accelerometer; autoplay; clipboard-write; encrypted-media; fullscreen; gyroscope; picture-in-picture; web-share";
    iframe.referrerPolicy = "strict-origin-when-cross-origin";
    
    // Force inline iframe styles
    iframe.style.width = "100%";
    iframe.style.height = "100%";
    iframe.style.border = "none";
    iframe.style.background = "#fff";
    iframe.style.minHeight = "0"; // prevent flex/grid blowout
    
    iframe.addEventListener("load", handleFrameLoaded);
    ui.frame.appendChild(iframe);

    if (!silent && request.note && urlStrings.length === 1) {
      showToast(request.note, "info");
    }
  });

  if (hasErrors && urlStrings.length > 1) {
    showToast("One or more links could not be loaded.", "warning");
  }

  clearTimeout(state.loadTimer);
  state.loadTimer = window.setTimeout(handleLoadTimeout, LOAD_TIMEOUT_MS);

  if (persist) {
    persistState();
  }
}

function handleFrameLoaded() {
  if (!state.isLoading) return;
  
  state.isLoading = false;
  ui.loadingBar.classList.remove("is-active");
  clearTimeout(state.loadTimer);

  if (!state.currentUrl) {
    return;
  }

  hideToast();
}

function handleLoadTimeout() {
  if (!state.isLoading || !state.currentUrl) {
    return;
  }

  state.isLoading = false;
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
  syncEmptyState();
  ui.frame.src = "about:blank";
  ui.emptyState.classList.remove("is-hidden");
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
      ui.chrome.classList.add("is-hidden");
      ui.toolbar.classList.add("is-hidden");
      
      if (typeof clearCanvas === "function") clearCanvas();
      hideBlockOverlay();
      
      if (!options.isInitial) {
        forceRefresh();
      }
    } else {
      document.body.classList.remove("is-presentation");
      ui.toolbar.classList.remove("is-hidden");
      syncChromeState();
      
      if (!options.isInitial) {
        forceRefresh();
      }
    }
  });
}

function forceRefresh() {
  const iframes = ui.frame.querySelectorAll("iframe");
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
    ui.loadingBar.classList.add("is-active");
  }
}

function registerActiveViewChanged() {
  if (!Office?.context?.document?.addHandlerAsync || !Office?.EventType?.ActiveViewChanged) {
    return;
  }

  Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, () => {
    syncActiveView();
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

  if (pathParts[1] !== "d" || !documentId) {
  
  if (docType === "forms" && pathParts[2] === "viewform") {
    return {
      note: "Converted to Google Forms embed.",
      providerName: "Google Forms",
      url: new URL(`https://docs.google.com/forms/d/${documentId}/viewform?embedded=true`)
    };
  }

  return {
    note: "",
    providerName: "Google Workspace",
    url
  };
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

  return {
    note: "",
    providerName: "Google Workspace",
    url
  };
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
        background: 'white',
        foreground: 'black'
      });
      ui.qrOverlay.classList.remove("is-hidden");
    } else {
      showToast("QR code library not loaded.", "error");
    }
  } catch (e) {
    showToast("Could not generate QR code.", "error");
  }
}

function hideQrCode() {
  ui.qrOverlay.classList.add("is-hidden");
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

function syncZoomState() {
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

function handleRefreshOptionClick(e) {
  const li = e.target.closest("li.recent-item");
  if (!li) return;
  
  const val = parseInt(li.dataset.val, 10);
  if (isNaN(val)) return;
  
  state.refreshInterval = val;
  applyRefreshInterval();
  persistState();
  
  ui.refreshDropdown.classList.add("is-hidden");
}

function applyRefreshInterval() {
  if (state.refreshTimerId) {
    clearInterval(state.refreshTimerId);
    state.refreshTimerId = null;
  }
  
  if (state.refreshInterval > 0) {
    ui.refreshBtn.style.color = "var(--accent)";
    state.refreshTimerId = setInterval(() => {
      if (state.currentUrl && !state.isLoading) {
        forceRefresh();
      }
    }, state.refreshInterval);
  } else {
    ui.refreshBtn.style.color = "";
  }
}


/* ── Advanced Tools Logic ────────────────────────────────────────── */

function syncAdvancedTools() {
  // Safe Mode
  if (state.safeMode) {
    ui.safeModeBtn.style.color = "var(--accent)";
    ui.safeModeOverlay.classList.remove("is-hidden");
  } else {
    ui.safeModeBtn.style.color = "";
    ui.safeModeOverlay.classList.add("is-hidden");
  }

  // Draw Mode
  if (state.drawMode) {
    ui.drawModeBtn.style.color = "var(--accent)";
    ui.drawCanvas.classList.remove("is-hidden");
    ui.clearDrawBtn.classList.remove("is-hidden");
    resizeCanvas();
  } else {
    ui.drawModeBtn.style.color = "";
    ui.drawCanvas.classList.add("is-hidden");
    ui.clearDrawBtn.classList.add("is-hidden");
  }

  // Invert Theme
  if (state.invertTheme) {
    ui.invertBtn.style.color = "var(--accent)";
    ui.frame.classList.add("smart-invert");
  } else {
    ui.invertBtn.style.color = "";
    ui.frame.classList.remove("smart-invert");
  }
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
  const clientX = e.clientX || (e.touches && e.touches[0].clientX);
  const clientY = e.clientY || (e.touches && e.touches[0].clientY);
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
