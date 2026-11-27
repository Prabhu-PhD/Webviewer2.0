// Provider URL adapters — pure functions, no DOM or Office dependencies.
// All functions accept and return standard URL objects (or throw Error).

export const KNOWN_BLOCKED_DOMAINS = [
  "google.com", "www.google.com",
  "facebook.com", "www.facebook.com",
  "twitter.com", "x.com",
  "github.com", "www.github.com",
  "reddit.com", "www.reddit.com",
  "amazon.com", "www.amazon.com",
  "apple.com", "www.apple.com",
  "microsoft.com", "www.microsoft.com",
  "linkedin.com", "www.linkedin.com",
  "instagram.com", "www.instagram.com"
];

export function isKnownBlockedDomain(hostname) {
  return KNOWN_BLOCKED_DOMAINS.includes(hostname.toLowerCase());
}

export function normalizeHostname(hostname) {
  return String(hostname ?? "").toLowerCase().replace(/^www\./, "").replace(/^m\./, "");
}

export function formatProviderName(hostname) {
  const providerName = hostname.replace(/^www\./, "");
  return providerName ? providerName : "Website";
}

export function extractStartSeconds(value) {
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

export function joinNotes(...notes) {
  return notes.filter(Boolean).join(" ").trim();
}

export function adaptProviderUrl(url) {
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

  // Google Maps — must come after the generic google.com block so the adapter
  // can switch the hostname to maps.google.com (which is not in KNOWN_BLOCKED_DOMAINS).
  if (hostname === "google.com" && url.pathname.startsWith("/maps")) {
    return adaptGoogleMapsUrl(url);
  }

  if (hostname === "maps.google.com") {
    return adaptGoogleMapsUrl(url);
  }

  // Google Maps short links (mobile share) — can't resolve without a server
  if (hostname === "maps.app.goo.gl" || (hostname === "goo.gl" && url.pathname.startsWith("/maps"))) {
    throw new Error("Google Maps short links can't be embedded directly. In Google Maps, tap Share → Embed a map and paste the iframe code instead.");
  }

  // Any other shortened link can't be resolved client-side.
  if (hostname === "goo.gl") {
    throw new Error("Shortened links can't be resolved here — paste the full destination URL instead.");
  }

  if (hostname === "lookerstudio.google.com" || hostname === "datastudio.google.com") {
    return adaptLookerStudioUrl(url);
  }

  if (hostname === "mentimeter.com") {
    return adaptMentimeterUrl(url);
  }

  if (hostname === "padlet.com") {
    return adaptPadletUrl(url);
  }

  if (hostname === "public.tableau.com") {
    return adaptTableauPublicUrl(url);
  }

  if (hostname === "openstreetmap.org") {
    return adaptOpenStreetMapUrl(url);
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
    const embedUrl = new URL(url.toString());

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
  if (pathParts[0]?.startsWith("shr")) {
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

function adaptGoogleMapsUrl(url) {
  // Already an embed path — switch hostname to maps.google.com so the
  // known-blocked-domain check on www.google.com doesn't fire.
  if (url.pathname.startsWith("/maps/embed")) {
    const embedUrl = new URL(url.toString());
    embedUrl.hostname = "maps.google.com";
    return { note: "", providerName: "Google Maps", url: embedUrl };
  }

  // Place/search URL — extract coordinates and build a simple embed.
  const coordMatch = url.pathname.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (coordMatch) {
    const lat = coordMatch[1];
    const lng = coordMatch[2];
    const zoomMatch = url.pathname.match(/@-?\d+\.\d+,-?\d+\.\d+,(\d+)z/);
    const zoom = zoomMatch ? zoomMatch[1] : "14";
    const embedUrl = new URL("https://maps.google.com/maps");
    embedUrl.searchParams.set("q", `${lat},${lng}`);
    embedUrl.searchParams.set("z", zoom);
    embedUrl.searchParams.set("output", "embed");
    return {
      note: "Converted to Google Maps embed.",
      providerName: "Google Maps",
      url: embedUrl
    };
  }

  // Can't auto-convert — advise using the Maps share dialog.
  return {
    note: "For best results: in Google Maps use Share → Embed a map and paste the iframe code here.",
    providerName: "Google Maps",
    url: new URL(url.toString().replace(/^https?:\/\/www\.google\.com/, "https://maps.google.com"))
  };
}

function adaptLookerStudioUrl(url) {
  // From: /reporting/ID/page/P_ID
  // To:   /embed/reporting/ID/page/P_ID
  if (url.pathname.startsWith("/reporting/") && !url.pathname.startsWith("/embed/")) {
    const embedUrl = new URL(url.toString());
    embedUrl.pathname = `/embed${url.pathname}`;
    return {
      note: "Converted to Looker Studio embed.",
      providerName: "Looker Studio",
      url: embedUrl
    };
  }
  return { note: "", providerName: "Looker Studio", url };
}

function adaptMentimeterUrl(url) {
  // From: /s/HASH/SLIDE_ID
  // To:   /embed/HASH/SLIDE_ID
  const match = url.pathname.match(/^\/s\/([^/]+)(?:\/([^/]+))?/);
  if (match) {
    const part1 = match[1];
    const part2 = match[2];
    const embedPath = part2 ? `/embed/${part1}/${part2}` : `/embed/${part1}`;
    const embedUrl = new URL(`https://www.mentimeter.com${embedPath}`);
    return {
      note: "Converted to Mentimeter embed.",
      providerName: "Mentimeter",
      url: embedUrl
    };
  }
  return { note: "", providerName: "Mentimeter", url };
}

function adaptPadletUrl(url) {
  // Append /embed if not already present
  if (!url.pathname.endsWith("/embed")) {
    const embedUrl = new URL(url.toString());
    embedUrl.pathname = url.pathname.replace(/\/?$/, "/embed");
    return {
      note: "Converted to Padlet embed.",
      providerName: "Padlet",
      url: embedUrl
    };
  }
  return { note: "", providerName: "Padlet", url };
}

function adaptTableauPublicUrl(url) {
  // Add embed query params to Tableau Public view URLs
  if (!url.searchParams.has(":embed")) {
    const embedUrl = new URL(url.toString());
    embedUrl.searchParams.set(":embed", "y");
    embedUrl.searchParams.set(":showVizHome", "no");
    return {
      note: "Converted to Tableau Public embed.",
      providerName: "Tableau Public",
      url: embedUrl
    };
  }
  return { note: "", providerName: "Tableau Public", url };
}

function adaptOpenStreetMapUrl(url) {
  // Already the embeddable export endpoint — pass through.
  if (url.pathname.startsWith("/export/embed.html")) {
    return { note: "", providerName: "OpenStreetMap", url };
  }

  // A standard OSM URL carries the view in the hash: #map=ZOOM/LAT/LNG
  const match = url.hash.match(/map=(\d+(?:\.\d+)?)\/(-?\d+\.\d+)\/(-?\d+\.\d+)/);
  if (match) {
    const zoom = Number.parseFloat(match[1]);
    const lat = Number.parseFloat(match[2]);
    const lng = Number.parseFloat(match[3]);
    // Approximate a bounding box around the centre; span shrinks as zoom grows.
    const lonSpan = 360 / 2 ** zoom;
    const latSpan = lonSpan / 2;
    const bbox = [
      (lng - lonSpan).toFixed(5),
      (lat - latSpan).toFixed(5),
      (lng + lonSpan).toFixed(5),
      (lat + latSpan).toFixed(5)
    ].join(",");
    const embedUrl = new URL("https://www.openstreetmap.org/export/embed.html");
    embedUrl.searchParams.set("bbox", bbox);
    embedUrl.searchParams.set("layer", "mapnik");
    embedUrl.searchParams.set("marker", `${lat},${lng}`);
    return {
      note: "Converted to OpenStreetMap embed.",
      providerName: "OpenStreetMap",
      url: embedUrl
    };
  }

  // Couldn't parse a view — point the user at the embed export tool.
  return {
    note: "For a map, use OpenStreetMap's Share → HTML embed link.",
    providerName: "OpenStreetMap",
    url
  };
}
