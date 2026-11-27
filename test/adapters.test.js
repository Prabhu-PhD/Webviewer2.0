import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  KNOWN_BLOCKED_DOMAINS,
  isKnownBlockedDomain,
  normalizeHostname,
  formatProviderName,
  extractStartSeconds,
  joinNotes,
  adaptProviderUrl,
} from "../web/adapters.js";

// ── Helpers ──────────────────────────────────────────────────────

function url(str) {
  return new URL(str);
}

// ── normalizeHostname ────────────────────────────────────────────

describe("normalizeHostname", () => {
  it("lowercases the hostname", () => {
    assert.equal(normalizeHostname("YouTube.COM"), "youtube.com");
  });
  it("strips www. prefix", () => {
    assert.equal(normalizeHostname("www.example.com"), "example.com");
  });
  it("strips m. prefix", () => {
    assert.equal(normalizeHostname("m.youtube.com"), "youtube.com");
  });
  it("handles empty string", () => {
    assert.equal(normalizeHostname(""), "");
  });
  it("handles null/undefined", () => {
    assert.equal(normalizeHostname(null), "");
    assert.equal(normalizeHostname(undefined), "");
  });
});

// ── formatProviderName ───────────────────────────────────────────

describe("formatProviderName", () => {
  it("returns Website for empty hostname", () => {
    assert.equal(formatProviderName(""), "Website");
  });
  it("strips www. prefix", () => {
    assert.equal(formatProviderName("www.example.com"), "example.com");
  });
  it("leaves bare hostname untouched", () => {
    assert.equal(formatProviderName("example.com"), "example.com");
  });
});

// ── extractStartSeconds ──────────────────────────────────────────

describe("extractStartSeconds", () => {
  it("returns 0 for falsy input", () => {
    assert.equal(extractStartSeconds(null), 0);
    assert.equal(extractStartSeconds(""), 0);
  });
  it("parses plain integer seconds", () => {
    assert.equal(extractStartSeconds("90"), 90);
  });
  it("parses h/m/s format", () => {
    assert.equal(extractStartSeconds("1h2m3s"), 3723);
    assert.equal(extractStartSeconds("2m30s"), 150);
    assert.equal(extractStartSeconds("45s"), 45);
    assert.equal(extractStartSeconds("1h"), 3600);
  });
  it("returns 0 for unrecognised format", () => {
    assert.equal(extractStartSeconds("abc"), 0);
  });
});

// ── joinNotes ────────────────────────────────────────────────────

describe("joinNotes", () => {
  it("joins non-empty strings with a space", () => {
    assert.equal(joinNotes("a", "b"), "a b");
  });
  it("drops empty/falsy parts", () => {
    assert.equal(joinNotes("a", "", null, "b"), "a b");
  });
  it("returns empty string when all parts are falsy", () => {
    assert.equal(joinNotes("", null, undefined), "");
  });
});

// ── isKnownBlockedDomain ─────────────────────────────────────────

describe("isKnownBlockedDomain", () => {
  it("matches a blocked domain", () => {
    assert.equal(isKnownBlockedDomain("google.com"), true);
  });
  it("is case-insensitive", () => {
    assert.equal(isKnownBlockedDomain("GOOGLE.COM"), true);
  });
  it("does not match a supported subdomain", () => {
    assert.equal(isKnownBlockedDomain("docs.google.com"), false);
    assert.equal(isKnownBlockedDomain("lookerstudio.google.com"), false);
  });
  it("KNOWN_BLOCKED_DOMAINS contains expected entries", () => {
    assert.ok(KNOWN_BLOCKED_DOMAINS.includes("facebook.com"));
    assert.ok(KNOWN_BLOCKED_DOMAINS.includes("twitter.com"));
  });
});

// ── YouTube ──────────────────────────────────────────────────────

describe("adaptProviderUrl — YouTube", () => {
  it("converts watch URL to nocookie embed", () => {
    const result = adaptProviderUrl(url("https://www.youtube.com/watch?v=dQw4w9WgXcQ"));
    assert.equal(result.providerName, "YouTube");
    assert.ok(result.url.hostname.includes("youtube-nocookie.com"));
    assert.ok(result.url.pathname.includes("dQw4w9WgXcQ"));
    assert.ok(result.note.includes("Converted"));
  });

  it("converts youtu.be short link", () => {
    const result = adaptProviderUrl(url("https://youtu.be/dQw4w9WgXcQ"));
    assert.ok(result.url.hostname.includes("youtube-nocookie.com"));
  });

  it("preserves ?t= timestamp in seconds", () => {
    const result = adaptProviderUrl(url("https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=90"));
    assert.equal(result.url.searchParams.get("start"), "90");
  });

  it("parses h/m/s timestamp", () => {
    const result = adaptProviderUrl(url("https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=1m30s"));
    assert.equal(result.url.searchParams.get("start"), "90");
  });

  it("passes through already-embedded URL without note", () => {
    const embedUrl = "https://www.youtube-nocookie.com/embed/dQw4w9WgXcQ";
    const result = adaptProviderUrl(url(embedUrl));
    assert.equal(result.note, "");
    assert.equal(result.url.toString(), embedUrl);
  });

  it("handles /shorts/ path", () => {
    const result = adaptProviderUrl(url("https://www.youtube.com/shorts/dQw4w9WgXcQ"));
    assert.ok(result.url.pathname.includes("dQw4w9WgXcQ"));
  });
});

// ── Vimeo ────────────────────────────────────────────────────────

describe("adaptProviderUrl — Vimeo", () => {
  it("converts standard vimeo URL to player embed", () => {
    const result = adaptProviderUrl(url("https://vimeo.com/123456789"));
    assert.equal(result.providerName, "Vimeo");
    assert.equal(result.url.hostname, "player.vimeo.com");
    assert.ok(result.url.pathname.includes("123456789"));
  });

  it("passes through already-embedded URL", () => {
    const embedUrl = "https://player.vimeo.com/video/123456789";
    const result = adaptProviderUrl(url(embedUrl));
    assert.equal(result.note, "");
  });

  it("passes through URL with no numeric ID", () => {
    const result = adaptProviderUrl(url("https://vimeo.com/channels/staffpicks"));
    assert.equal(result.note, "");
  });
});

// ── Loom ─────────────────────────────────────────────────────────

describe("adaptProviderUrl — Loom", () => {
  it("converts share URL to embed URL", () => {
    const result = adaptProviderUrl(url("https://www.loom.com/share/abc123def456"));
    assert.equal(result.providerName, "Loom");
    assert.ok(result.url.pathname.startsWith("/embed/"));
    assert.ok(result.note.includes("Converted"));
  });

  it("passes through already-embedded URL", () => {
    const embedUrl = "https://www.loom.com/embed/abc123def456";
    const result = adaptProviderUrl(url(embedUrl));
    assert.equal(result.note, "");
  });
});

// ── Google Docs ──────────────────────────────────────────────────

describe("adaptProviderUrl — Google Docs", () => {
  it("converts presentation to /embed", () => {
    const result = adaptProviderUrl(url("https://docs.google.com/presentation/d/DOC_ID/edit"));
    assert.equal(result.providerName, "Google Slides");
    assert.ok(result.url.pathname.endsWith("/embed"));
  });

  it("converts document to /preview", () => {
    const result = adaptProviderUrl(url("https://docs.google.com/document/d/DOC_ID/edit"));
    assert.equal(result.providerName, "Google Docs");
    assert.ok(result.url.pathname.endsWith("/preview"));
  });

  it("converts spreadsheet to /preview", () => {
    const result = adaptProviderUrl(url("https://docs.google.com/spreadsheets/d/SHEET_ID/edit"));
    assert.equal(result.providerName, "Google Sheets");
    assert.ok(result.url.pathname.endsWith("/preview"));
  });

  it("adds ?embedded=true to forms URL", () => {
    const result = adaptProviderUrl(url("https://docs.google.com/forms/d/FORM_ID/viewform"));
    assert.equal(result.providerName, "Google Forms");
    assert.equal(result.url.searchParams.get("embedded"), "true");
  });

  it("does not duplicate embedded param if already set", () => {
    const result = adaptProviderUrl(url("https://docs.google.com/forms/d/FORM_ID/viewform?embedded=true"));
    assert.equal(result.note, "");
  });
});

// ── Google Maps ──────────────────────────────────────────────────

describe("adaptProviderUrl — Google Maps", () => {
  it("converts coordinate URL to embed", () => {
    const result = adaptProviderUrl(url("https://www.google.com/maps/@37.7749,-122.4194,14z"));
    assert.equal(result.providerName, "Google Maps");
    assert.equal(result.url.hostname, "maps.google.com");
    assert.equal(result.url.searchParams.get("output"), "embed");
  });

  it("passes through existing embed path with hostname fix", () => {
    const result = adaptProviderUrl(url("https://www.google.com/maps/embed?pb=abc"));
    assert.equal(result.url.hostname, "maps.google.com");
    assert.equal(result.note, "");
  });

  it("throws for Maps short links", () => {
    assert.throws(
      () => adaptProviderUrl(url("https://maps.app.goo.gl/abc123")),
      /short links/
    );
  });
});

// ── Figma ────────────────────────────────────────────────────────

describe("adaptProviderUrl — Figma", () => {
  it("wraps /design/ URL in embed", () => {
    const result = adaptProviderUrl(url("https://www.figma.com/design/abc/My-Design?node-id=0"));
    assert.equal(result.providerName, "Figma");
    assert.ok(result.url.searchParams.get("url")?.includes("figma.com/design/"));
    assert.ok(result.note.includes("Converted"));
  });

  it("wraps /proto/ URL in embed", () => {
    const result = adaptProviderUrl(url("https://www.figma.com/proto/abc/proto?node-id=0"));
    assert.ok(result.url.searchParams.get("url")?.includes("/proto/"));
  });
});

// ── Miro ─────────────────────────────────────────────────────────

describe("adaptProviderUrl — Miro", () => {
  it("converts /app/board/ to /app/live-embed/", () => {
    const result = adaptProviderUrl(url("https://miro.com/app/board/abc123/"));
    assert.equal(result.providerName, "Miro");
    assert.ok(result.url.pathname.startsWith("/app/live-embed/"));
  });
});

// ── Canva ────────────────────────────────────────────────────────

describe("adaptProviderUrl — Canva", () => {
  it("converts /design/.../edit to embed view", () => {
    const result = adaptProviderUrl(url("https://www.canva.com/design/ABCD/abc/edit"));
    assert.equal(result.providerName, "Canva");
    assert.ok(result.url.pathname.endsWith("/view"));
    assert.ok(result.url.searchParams.has("embed"));
  });

  it("converts /design/.../watch to embed view", () => {
    const result = adaptProviderUrl(url("https://www.canva.com/design/ABCD/abc/watch"));
    assert.ok(result.url.pathname.endsWith("/view"));
  });
});

// ── Airtable ─────────────────────────────────────────────────────

describe("adaptProviderUrl — Airtable", () => {
  it("converts /shrXXX path to /embed/shrXXX", () => {
    const result = adaptProviderUrl(url("https://airtable.com/shrABCDEF123"));
    assert.equal(result.providerName, "Airtable");
    assert.ok(result.url.pathname.startsWith("/embed/shr"));
  });

  it("passes through non-shr paths", () => {
    const result = adaptProviderUrl(url("https://airtable.com/appXXX/tblYYY"));
    assert.equal(result.note, "");
  });
});

// ── CodePen ──────────────────────────────────────────────────────

describe("adaptProviderUrl — CodePen", () => {
  it("converts /pen/ to /embed/", () => {
    const result = adaptProviderUrl(url("https://codepen.io/user/pen/abcdef"));
    assert.equal(result.providerName, "CodePen");
    assert.ok(result.url.pathname.includes("/embed/"));
  });
});

// ── Spotify ──────────────────────────────────────────────────────

describe("adaptProviderUrl — Spotify", () => {
  it("converts track URL to embed", () => {
    const result = adaptProviderUrl(url("https://open.spotify.com/track/4cOdK2wGLETKBW3PvgPWqT"));
    assert.equal(result.providerName, "Spotify");
    assert.ok(result.url.pathname.startsWith("/embed/track/"));
  });

  it("passes through already-embedded URL", () => {
    const result = adaptProviderUrl(url("https://open.spotify.com/embed/track/4cOdK2wGLETKBW3PvgPWqT"));
    assert.equal(result.note, "");
  });
});

// ── Looker Studio ────────────────────────────────────────────────

describe("adaptProviderUrl — Looker Studio", () => {
  it("prepends /embed to /reporting/ path", () => {
    const result = adaptProviderUrl(url("https://lookerstudio.google.com/reporting/abc/page/xyz"));
    assert.equal(result.providerName, "Looker Studio");
    assert.ok(result.url.pathname.startsWith("/embed/reporting/"));
  });

  it("also handles datastudio.google.com", () => {
    const result = adaptProviderUrl(url("https://datastudio.google.com/reporting/abc/page/xyz"));
    assert.ok(result.url.pathname.startsWith("/embed/"));
  });

  it("passes through already-embedded URL", () => {
    const result = adaptProviderUrl(url("https://lookerstudio.google.com/embed/reporting/abc"));
    assert.equal(result.note, "");
  });
});

// ── Mentimeter ───────────────────────────────────────────────────

describe("adaptProviderUrl — Mentimeter", () => {
  it("converts /s/HASH to /embed/HASH", () => {
    const result = adaptProviderUrl(url("https://www.mentimeter.com/s/abc123"));
    assert.equal(result.providerName, "Mentimeter");
    assert.ok(result.url.pathname.startsWith("/embed/"));
  });

  it("includes slide ID in embed path", () => {
    const result = adaptProviderUrl(url("https://www.mentimeter.com/s/abc123/slide456"));
    assert.ok(result.url.pathname.includes("slide456"));
  });
});

// ── Padlet ───────────────────────────────────────────────────────

describe("adaptProviderUrl — Padlet", () => {
  it("appends /embed to padlet URL", () => {
    const result = adaptProviderUrl(url("https://padlet.com/user/board"));
    assert.equal(result.providerName, "Padlet");
    assert.ok(result.url.pathname.endsWith("/embed"));
  });

  it("passes through already-embedded URL", () => {
    const result = adaptProviderUrl(url("https://padlet.com/user/board/embed"));
    assert.equal(result.note, "");
  });
});

// ── Tableau Public ───────────────────────────────────────────────

describe("adaptProviderUrl — Tableau Public", () => {
  it("adds :embed and :showVizHome params", () => {
    const result = adaptProviderUrl(url("https://public.tableau.com/views/MyViz/Sheet1"));
    assert.equal(result.providerName, "Tableau Public");
    assert.equal(result.url.searchParams.get(":embed"), "y");
    assert.equal(result.url.searchParams.get(":showVizHome"), "no");
  });

  it("passes through already-embedded URL", () => {
    const result = adaptProviderUrl(url("https://public.tableau.com/views/MyViz/Sheet1?:embed=y"));
    assert.equal(result.note, "");
  });
});

// ── OpenStreetMap ────────────────────────────────────────────────

describe("adaptProviderUrl — OpenStreetMap", () => {
  it("converts hash-based URL to embed with bbox", () => {
    const result = adaptProviderUrl(url("https://www.openstreetmap.org/#map=14/51.5074/-0.1278"));
    assert.equal(result.providerName, "OpenStreetMap");
    assert.ok(result.url.pathname.startsWith("/export/embed.html"));
    assert.ok(result.url.searchParams.has("bbox"));
  });

  it("passes through existing embed endpoint", () => {
    const result = adaptProviderUrl(url("https://www.openstreetmap.org/export/embed.html?bbox=0,0,1,1"));
    assert.equal(result.note, "");
  });
});

// ── SharePoint ───────────────────────────────────────────────────

describe("adaptProviderUrl — SharePoint", () => {
  it("adds embedview params to Doc.aspx URL", () => {
    const result = adaptProviderUrl(url("https://contoso.sharepoint.com/sites/mysite/_layouts/15/Doc.aspx?sourcedoc=abc"));
    assert.equal(result.providerName, "SharePoint");
    assert.equal(result.url.searchParams.get("action"), "embedview");
    assert.equal(result.url.searchParams.get("wdAllowInteractivity"), "True");
  });

  it("passes through non-Doc.aspx SharePoint URLs", () => {
    const result = adaptProviderUrl(url("https://contoso.sharepoint.com/sites/mysite/Shared%20Documents/file.pdf"));
    assert.equal(result.note, "");
  });
});

// ── Power BI ─────────────────────────────────────────────────────

describe("adaptProviderUrl — Power BI", () => {
  it("adds note for authoring URL", () => {
    const result = adaptProviderUrl(url("https://app.powerbi.com/groups/me/reports/abc-123"));
    assert.equal(result.providerName, "Power BI");
    assert.ok(result.note.includes("publish-to-web"));
  });

  it("no note for reportEmbed URL", () => {
    const result = adaptProviderUrl(url("https://app.powerbi.com/reportEmbed?reportId=abc"));
    assert.equal(result.note, "");
  });
});

// ── Microsoft Forms ──────────────────────────────────────────────

describe("adaptProviderUrl — Microsoft Forms", () => {
  it("passes through forms.office.com URLs", () => {
    const result = adaptProviderUrl(url("https://forms.office.com/r/abc123"));
    assert.equal(result.providerName, "Microsoft Forms");
    assert.equal(result.note, "");
  });
});

// ── Unknown/pass-through ─────────────────────────────────────────

describe("adaptProviderUrl — unknown domain", () => {
  it("returns providerName from hostname for unknown domain", () => {
    const result = adaptProviderUrl(url("https://example.com/page"));
    assert.equal(result.providerName, "example.com");
    assert.equal(result.note, "");
    assert.equal(result.url.hostname, "example.com");
  });

  it("throws for Canva short link", () => {
    assert.throws(
      () => adaptProviderUrl(url("https://canva.link/abc123")),
      /short links/
    );
  });

  it("throws for goo.gl shortened link", () => {
    assert.throws(
      () => adaptProviderUrl(url("https://goo.gl/maps/abc123")),
      /short links/
    );
  });
});
