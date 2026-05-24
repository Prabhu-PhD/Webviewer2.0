const STORAGE_KEYS = {
  chromeVisible: "webviewer2.chromeVisible",
  currentUrl: "webviewer2.currentUrl",
  recentUrls: "webviewer2.recentUrls"
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
}

function handleOfficeReady() {
  state.officeReady = true;
  hydrateFromDocumentSettings();
  syncChromeState();
  syncActiveView();
  registerActiveViewChanged();
}

function hydrateFromBrowserStorage() {
  try {
    const chromeVisible = window.localStorage.getItem(STORAGE_KEYS.chromeVisible);
    const currentUrl = window.localStorage.getItem(STORAGE_KEYS.currentUrl);
    const recentUrlsStr = window.localStorage.getItem(STORAGE_KEYS.recentUrls);

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
  const request = prepareEmbedRequest(rawInput);

  state.currentUrl = request.normalizedUrl;
  ui.urlInput.value = request.normalizedUrl;
  
  syncEmptyState();
  ui.emptyState.classList.add("is-hidden");
  hideBlockOverlay();

  // Instant blocklist check
  if (isKnownBlockedDomain(request.hostname)) {
    showBlockOverlay();
    return;
  }

  state.isLoading = true;
  ui.loadingBar.classList.add("is-active");
  ui.frame.src = request.normalizedUrl;

  clearTimeout(state.loadTimer);
  state.loadTimer = window.setTimeout(handleLoadTimeout, LOAD_TIMEOUT_MS);

  if (!silent && request.note) {
    showToast(request.note, "info");
  }

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
    ui.shell.dataset.view = nextView;
  });
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
  if (url.pathname.startsWith("/design/") && url.pathname.endsWith("/view")) {
    const embedUrl = new URL(url.toString());
    embedUrl.searchParams.set("embed", "");
    return {
      note: "Converted to Canva embed.",
      providerName: "Canva",
      url: embedUrl
    };
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
}
