# Webviewer 2.0

A free PowerPoint content add-in that embeds live web pages directly on a slide — the open-source replacement for the Microsoft Web Viewer add-in retired in December 2024.

## Features

- **Paste any URL** and it appears as a live frame on the slide
- **15+ provider adapters** — YouTube, Vimeo, Loom, Google Slides, Google Docs, Google Sheets, Google Forms, Google Maps, Looker Studio, Power BI, Figma, Miro, Canva, Airtable, CodePen, Spotify, Mentimeter, Padlet, Tableau Public, OpenStreetMap, SharePoint — share URLs are auto-converted to embed URLs
- **Split-screen** — separate up to 4 URLs with commas for side-by-side panes
- **Auto-refresh** — keep live dashboards current during a presentation (30 s / 1 m / 5 m)
- **Draw mode** — annotate over live content with a stylus or mouse
- **Mobile View** — resizes the PowerPoint shape to phone width (390 px) so the embedded page renders its genuine responsive mobile layout
- **Desktop Fit** — scales the embed to simulate a 1280 px desktop viewport
- **Zoom controls** — zoom in/out with persistent per-shape zoom level
- **Auto-scroll** — slow / medium / fast continuous pan through the page (for kiosk or display use)
- **QR code** — generate a scannable link for the current URL in one click
- **Dark toolbar** — matches dark-themed slides
- **Invert theme** — inverts embedded page colours for dark-mode compatibility
- **Safe Mode** — loads the page in a sandboxed overlay for untrusted content
- **Presentation mode** — toolbar auto-hides when the slideshow starts; floating toggle to show/hide it; embedded page stays loaded across edit ↔ slideshow transitions without reloading

## Per-shape state

Every add-in instance on every slide is fully independent. The following settings are stored in Office document settings (scoped to the specific shape in the .pptx file, not shared between instances):

- Loaded URL
- Zoom level
- Desktop Fit / Mobile View
- Safe Mode
- Auto-refresh interval
- Chrome visibility (toolbar shown/hidden)

Global user preferences (dark toolbar, invert theme) are stored in browser localStorage and shared across all instances.

## Project layout

| Path | Purpose |
|---|---|
| `manifest.prod.xml` | Production manifest pointing at GitHub Pages |
| `manifest.localhost.xml` | Local development manifest (HTTPS localhost:3000) |
| `manifest.hosted.xml` | Hosted manifest template for custom deployments |
| `web/content.html` | Add-in surface loaded inside the PowerPoint shape |
| `web/app.js` | All app logic — URL processing, state, Office API, UI |
| `web/styles.css` | UI styling |
| `web/support.html` | Support page linked from the manifest |
| `web/privacy.html` | Privacy policy |
| `web/terms.html` | Terms of use |
| `scripts/serve.mjs` | Local HTTPS dev server (no dependencies) |
| `scripts/new-dev-cert.ps1` | Generates and trusts a localhost certificate |
| `scripts/start-local.ps1` | Starts the local server |
| `scripts/dev-doctor.ps1` | Checks manifest, Node, and certificate setup |
| `docs/local-dev.md` | Local HTTPS, sideloading, and Windows troubleshooting |
| `docs/research.md` | Architecture decisions and platform constraint notes |

## Local development

```powershell
powershell -ExecutionPolicy Bypass -File scripts/new-dev-cert.ps1
powershell -ExecutionPolicy Bypass -File scripts/dev-doctor.ps1
powershell -ExecutionPolicy Bypass -File scripts/start-local.ps1
```

Then sideload `manifest.localhost.xml` into PowerPoint Desktop via **Insert → Get Add-ins → Manage My Add-ins → Upload My Add-in**.

If PowerPoint shows `We can't open this add-in from localhost`, run this once from an elevated prompt:

```powershell
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

Optional HTTPS environment variables:

```powershell
$env:SSL_KEY_FILE  = "C:\path\to\localhost.key"
$env:SSL_CERT_FILE = "C:\path\to\localhost.crt"
node scripts/serve.mjs
```

Full details in [docs/local-dev.md](docs/local-dev.md).

## Constraints

**Embedding** — no add-in can force a site to allow iframe embedding. If a site sends restrictive `X-Frame-Options` or `frame-ancestors` CSP headers, the viewer shows a block overlay with an Open-in-Browser and QR fallback.

**Runtime** — requires the Chromium-based WebView2 runtime that ships with Microsoft 365 / Office 2021 and newer. Not compatible with the legacy IE-based WebView in some Office 2016/2019 installations.

**Mobile Office** — content add-ins (`xsi:type="ContentApp"`) are a desktop-only feature. Not supported on Office for iOS or Android.

**Slideshow scroll** — when the cursor is over the add-in during a slideshow, mouse-wheel events are captured by PowerPoint at the OS level for slide navigation and never reach the iframe. Use Auto-scroll for hands-free panning during presentations.
