# Local Development

Date: April 21, 2026

This project is intentionally lightweight, so the local workflow is based on PowerShell scripts instead of a large toolchain.

## Recommended Flow

1. Create and trust a localhost certificate.

```powershell
powershell -ExecutionPolicy Bypass -File scripts/new-dev-cert.ps1
```

2. Check local readiness.

```powershell
powershell -ExecutionPolicy Bypass -File scripts/dev-doctor.ps1
```

3. Start the local server.

```powershell
powershell -ExecutionPolicy Bypass -File scripts/start-local.ps1
```

4. Open the preview directly in a browser if you want a fast UI-only check.

```text
https://localhost:3000/content.html
```

## Sideloading Paths

### PowerPoint on the web

This is the fastest manual loop for an add-in-only manifest.

According to Microsoft Learn:

- Open Office on the web.
- Open a PowerPoint file.
- Select `Home > Add-ins > More Settings`.
- Select `Upload My Add-in`.
- Choose `manifest.localhost.xml`.

Reference:

- https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing

### PowerPoint desktop on Windows

For add-in-only manifests, Microsoft documents two realistic routes:

- publish the manifest through an organization-managed deployment flow such as an app catalog or Microsoft 365 admin center
- use a trusted shared-folder catalog on Windows for local or preproduction testing

The shared-folder route is the most practical desktop path when you want to test this content add-in locally without building a larger Yeoman-based toolchain around it.

Reference:

- https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins

## Localhost Troubleshooting On Windows

If PowerPoint desktop reports `We can't open this add-in from localhost`, Microsoft’s current troubleshooting article says one common cause is that Desktop App Web Viewer is missing a loopback exemption.

Run this command from an elevated command prompt:

```powershell
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

Reference:

- https://learn.microsoft.com/en-us/troubleshoot/office/office-suite-issues/cannot-open-add-in-from-localhost

## Notes

- Office add-ins should use HTTPS during development.
- Self-signed certificates are acceptable for development and testing as long as the certificate is trusted on the local machine.
- This repo keeps the default dev certificate material under `certs/`, which is already gitignored.
