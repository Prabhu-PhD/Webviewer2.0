# Webviewer2

Webviewer2 is a fresh foundation for a PowerPoint content add-in that embeds live web content directly on a slide.

This repo starts from the problem the original Microsoft `Web Viewer` left behind:

- It was retired on December 2, 2024.
- Users relied on it for dashboards, teaching aids, timers, status boards, and mixed slide plus web layouts.
- It frequently failed without clear feedback when a site refused iframe embedding.

This first cut gives us a clean, production-oriented base:

- A PowerPoint content add-in manifest using the add-in-only XML format.
- A slide-embedded viewer shell with URL normalization and document-persisted state.
- Provider-aware normalization for common share links such as YouTube, Vimeo, Loom, and Google Docs or Slides.
- Better feedback when a site likely blocks embedding.
- A tiny local web server, a localhost certificate workflow, and manifest validation helpers.

## Why Content Add-In First

PowerPoint supports both task pane add-ins and content add-ins. We are using a content add-in because the original value proposition was seeing web content inside the slide itself, not beside the deck.

## Project Layout

- `manifest.localhost.xml`: local development manifest that points at `https://localhost:3000`
- `manifest.hosted.xml`: hosted manifest template for a real deployment URL
- `web/content.html`: the embedded slide surface
- `web/app.js`: app logic for loading, persisting, and presenting web content
- `web/styles.css`: UI styling for the embedded viewer
- `web/support.html`: lightweight support page used by the manifest
- `scripts/serve.mjs`: no-dependency static server
- `scripts/new-dev-cert.ps1`: generates and trusts a localhost development certificate
- `scripts/start-local.ps1`: starts the local server with the default development certificate when present
- `scripts/dev-doctor.ps1`: checks the manifest, Node runtime, and local certificate setup
- `scripts/validate-manifest.ps1`: pragmatic manifest sanity checks
- `docs/local-dev.md`: local HTTPS, sideloading, and Windows desktop troubleshooting notes
- `docs/research.md`: dated notes on retirement, platform constraints, and architecture choices

## Local Development

1. Generate a trusted localhost certificate.
2. Run the local doctor script.
3. Start the local server over HTTPS.
4. Sideload `manifest.localhost.xml` into PowerPoint.

Commands:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/new-dev-cert.ps1
powershell -ExecutionPolicy Bypass -File scripts/dev-doctor.ps1
powershell -ExecutionPolicy Bypass -File scripts/start-local.ps1
powershell -ExecutionPolicy Bypass -File scripts/validate-manifest.ps1
```

The server supports both HTTP and HTTPS and now auto-discovers the default development certificate in `certs/`:

- For real PowerPoint testing in Office, use HTTPS.
- If no certificate is configured, the server falls back to HTTP so the UI can still be previewed in a normal browser.

Optional environment variables for HTTPS:

```powershell
$env:SSL_KEY_FILE="C:\path\to\localhost.key"
$env:SSL_CERT_FILE="C:\path\to\localhost.crt"
node scripts/serve.mjs
```

Or with a PFX bundle:

```powershell
$env:SSL_PFX_FILE="C:\path\to\localhost.pfx"
$env:SSL_PFX_PASSPHRASE="your-passphrase"
node scripts/serve.mjs
```

If PowerPoint desktop shows `We can't open this add-in from localhost`, Microsoft’s current guidance is to add the Desktop App Web Viewer loopback exemption from an elevated prompt:

```powershell
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

The practical local workflow is documented in [docs/local-dev.md](/C:/Users/prabh/source/repos/Webviewer2/docs/local-dev.md).

## What This MVP Handles

- Accepts a plain URL or an iframe snippet and extracts the first usable URL.
- Normalizes hostnames to HTTPS when the protocol is missing.
- Rewrites several common share URLs into embed-friendly URLs when the provider has a predictable embed format.
- Persists the last loaded URL in document settings so the slide remembers it.
- Detects edit vs slide-show view where the host exposes that information.
- Gives a clearer explanation when a page is probably blocked by `frame-ancestors` or `X-Frame-Options`.

## Important Constraint

No PowerPoint add-in can force an arbitrary website to allow embedding. If a site sends restrictive CSP or `X-Frame-Options` headers, the viewer must respect that. The right long-term product answer is better diagnostics, better vendor-specific embed guidance, and optional provider integrations, not brittle hacks.

## Suggested Next Milestones

1. Add provider presets and templates for common live-slide use cases such as timers, Power BI, maps, and internal dashboards.
2. Add a companion task pane for bookmarks, slide presets, and failure diagnostics.
3. Add telemetry and structured diagnostics for blocked embeds and load failures.
4. Add a deployment path for AppSource submission, including branded assets and support/privacy pages.

## Research Notes

The architecture choices in this repo are summarized in [docs/research.md](/C:/Users/prabh/source/repos/Webviewer2/docs/research.md).
