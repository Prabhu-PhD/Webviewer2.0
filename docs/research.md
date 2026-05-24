# Research Notes

Date: April 21, 2026

## What Happened To The Original Web Viewer

- Microsoft's original `Web Viewer` add-in existed as a PowerPoint content add-in.
- Microsoft retired it on December 2, 2024.
- Public-facing explanations pointed to low usage, but community posts show some users depended on it heavily, especially in teaching and live presentation workflows.

## What Users Needed It For

Patterns repeated across user reports:

- countdown timers
- classroom dashboards
- embedded homework/status pages
- live dashboards and reports
- mixed PowerPoint plus webpage layouts during presentations

That usage pattern is exactly why a content add-in is still the right base architecture.

## Historical Failure Modes We Should Design Around

1. Silent embed failures
   Many sites never rendered because they blocked iframe embedding with `Content-Security-Policy: frame-ancestors` or `X-Frame-Options`.

2. Fragile trust and certificate issues
   Users reported periods where the add-in failed because of certificate/signing issues.

3. Slide-show inconsistencies
   Community issues around PowerPoint content add-ins show view-switch behavior and presentation-mode rendering have had edge cases on some clients.

4. Weak diagnostics
   The old experience often looked blank without telling the user whether the problem was the URL, the site, host policy, or platform behavior.

## Architecture Decisions For This Repo

### 1. PowerPoint content add-in

Reason:

- The value is content inside the slide, not a side panel.
- Microsoft Learn still documents content add-ins as the correct model for embedded HTML content in PowerPoint.

### 2. Add-in-only XML manifest

Reason:

- Microsoft Learn states that the unified manifest remains preview-only for production PowerPoint add-ins.
- The add-in-only manifest is still the stable production choice for PowerPoint.

### 3. Plain web app with minimal runtime assumptions

Reason:

- The add-in must run in multiple host surfaces.
- Windows Office desktop uses Edge WebView2.
- Mac and iOS use WKWebView.
- PowerPoint on the web uses the browser itself and iframe hosting.

### 4. Better diagnostics instead of brittle workarounds

Reason:

- Arbitrary websites cannot be force-embedded if they explicitly forbid it.
- The product needs to explain this clearly and guide users toward embed URLs when available.

## Source Snapshot

The following sources informed these decisions:

- Microsoft Learn: PowerPoint add-ins
  https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/powerpoint-add-ins

- Microsoft Learn: Office Add-ins manifest
  https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests

- Microsoft Learn: Content Office Add-ins
  https://learn.microsoft.com/en-us/office/dev/add-ins/design/content-add-ins

- Microsoft Learn: Browsers and webview controls used by Office Add-ins
  https://learn.microsoft.com/en-us/office/dev/add-ins/testing/ie-11-testing

- Microsoft Q&A thread discussing retirement concern
  https://learn.microsoft.com/en-us/answers/questions/5390831/please-continue-to-support-the-web-viewer-add-in

- Public rollout summary of the retirement notice referencing Microsoft message center item `MC921112`
  https://mwpro.co.uk/blog/2024/10/29/microsoft-powerpoint-web-viewer-add-in-retires-december-2024-mc921112/

- MDN on iframe embedding restrictions via `frame-ancestors`
  https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Headers/Content-Security-Policy/frame-ancestors

## Working Assumptions

- We are optimizing first for PowerPoint on Windows and PowerPoint on the web.
- We will not promise arbitrary-site embedding success.
- We should design the product around common successful cases plus excellent failure messaging.
