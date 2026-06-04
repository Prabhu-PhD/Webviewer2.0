# Changelog

All notable changes to Webviewer 2.0 are documented here.

---

## 2.0.6 — June 2026 (v17)

### Fixed — correctness
- **New inserts no longer inherit a stale URL.** Previously, inserting a fresh add-in on a new slide would load the last URL used by any other Webviewer instance (they shared a browser localStorage entry). Each instance is now fully independent — a new insert always starts blank.
- **Slideshow mode no longer reloads the webpage.** Entering or exiting the slideshow used to force-reload the iframe. The embedded page now stays loaded and unchanged across edit ↔ slideshow transitions. Only the toolbar visibility changes.
- **All per-shape settings are now isolated per instance.** Zoom level, Desktop Fit, Safe Mode, Auto-refresh interval, Mobile View, and the loaded URL are stored in Office document settings (scoped to each individual shape in the .pptx file). Previously several of these were shared in browser storage, so changing one add-in's settings would bleed into others on different slides.

### Changed — Mobile View
- **Mobile View now resizes the PowerPoint shape** to a genuine phone-width viewport instead of applying a CSS scale transform. The embedded page renders its real responsive layout at 1:1 scale — it actually sees a 390 px wide browser window, not a zoomed-out desktop layout. The shape is restored to its original size when Mobile View is toggled off.

---

## 2.0.5 — May 2026 (v15 / v16)

### Added — providers
- **7 new auto-converters:** Airtable, CodePen, Spotify, Mentimeter, Padlet, Tableau Public, OpenStreetMap — share URLs are automatically rewritten to their embed form.
- **Google Maps** unblocked — `google.com/maps` and `maps.google.com` URLs are now converted to the embeddable format. Short links (`maps.app.goo.gl`) show a clear error explaining how to get the embed code instead.

### Added — features
- **Auto-Scroll** — slow / medium / fast continuous scroll through the embedded page, useful for kiosk or ambient display use cases.
- **Mobile View** — preview any site at phone width directly on the slide.
- **Active auto-refresh indicator** — the selected refresh interval is shown with a checkmark in the dropdown.

### Fixed
- Single URLs containing commas (Google Maps coordinates, query parameters with comma-separated values) are no longer incorrectly split into multiple panes.
- Split-screen: a slow or blocked pane now shows a toast warning instead of a full-screen block overlay that would hide the other panes.
- Alt-tab away from PowerPoint and back no longer triggers an unnecessary page reload.
- OpenStreetMap: the `#map=zoom/lat/lng` fragment format is now correctly converted to the embed URL.
- Keyboard users can now operate the Recent Links and Auto-Refresh dropdowns with Enter/Space; focus rings are visible throughout.
- Startup double-load guard: returning to a presentation no longer loads the URL twice.

---

## 2.0.4 — May 2026 (v14)

### Added
- **QR code button** — generates a scannable link for the current URL in one click.
- **Recent Links** — the address bar shows up to 5 previously loaded URLs.
- **Dark toolbar** — matches dark-themed slides.
- **Invert Theme** — inverts the embedded page's colours for dark-mode compatibility.
- **Safe Mode** — loads the page in a sandboxed overlay.
- **Desktop Fit** — scales the embed to simulate a 1280 px desktop viewport.
- **Zoom in / Zoom out / Reset** controls with a persistent zoom level.
- **Draw Mode** — annotate over live content with a stylus or mouse; clear annotations with one tap.
- **Floating toolbar toggle** — a small handle in the corner lets the presenter show/hide the toolbar during a slideshow without needing keyboard access.
- **Split-screen** — separate up to 4 URLs with commas for side-by-side panes.
- **Auto-Refresh** — keep live dashboards current during a presentation (30 s / 1 m / 5 m intervals).
- **More menu** — secondary tools grouped cleanly to keep the primary toolbar uncluttered.

### Fixed
- Toolbar correctly auto-hides when a slideshow starts and restores its state when the presenter exits.
- Numerous accessibility improvements: ARIA roles, labels, keyboard navigation.
- Dark mode contrast and button state indicators.
- Performance: iframe cache, debounced resize handling.

---

## 2.0.1 — May 2026

### Added
- **13 provider auto-converters:** YouTube (with timestamp and playlist support), Vimeo, Loom, Google Slides, Google Docs, Google Sheets, Google Forms, Looker Studio / Data Studio, Power BI (publish-to-web), Figma, Miro, Canva, SharePoint — paste any share link and the correct embed URL is used automatically.
- Microsoft Forms and Google Forms pass through without conversion (they embed directly).
- `canva.link` short links show a clear error with instructions.

### Fixed
- Manifest schema ordering for Microsoft Admin Center validation.
- Requirements element added for correct Office version gating.
- State hydration in Slide Show mode.

---

## 2.0.0 — May 2026

Initial public release.

- Paste any HTTPS URL or `<iframe>` snippet — the src is extracted automatically.
- Provider-unrecognised URLs pass through as-is.
- Known embedding-blocked domains (Google, GitHub, Reddit, etc.) show an immediate error overlay instead of waiting for a 9-second timeout.
- Loaded URL is saved in Office document settings so the slide remembers it across saves and reloads.
- Slideshow mode detection via `Office.context.document.getActiveViewAsync`.
- Open in Browser button and fallback for sites that block embedding.
