# Polarion Requirement Assistant

A browser-based requirement review app with a Node.js cloud service. Users do not install VS Code or a VSIX; they open the hosted site, sign in with GitHub, and use their own GitHub Copilot access.

## What it does

1. Import a SharePoint link-list `.xlsx` or `.csv` with `File Title` and `Hyperlink` columns.
2. Select a locally synced folder containing PDF, DOCX, TXT, or Markdown documents.
3. Sign in with GitHub, provide the allowed Scope values, and run Copilot analysis in the cloud.
4. Review and edit the rows, then export a four-column `.xlsx`: `File Title`, `Hyperlink`, `Priority`, `Scope`.

The website does not connect to SharePoint. Users must import their own verified link list. Scanned PDFs without extractable text are not supported by this cloud version.

## Deploy the cloud service

GitHub Pages only hosts static files and cannot run the Copilot backend. Deploy this repository as a **Docker web service** (the `Dockerfile` is at the repository root); the Node service serves both the web page and the API from the same origin. Any Docker-capable host can be used, subject to your organization's data and hosting policies.

1. Create a GitHub OAuth App. Set its callback URL to `https://YOUR-SERVICE-HOST/auth/callback`.
2. Deploy the repository from its root using the included `Dockerfile`.
3. Configure these environment variables in the host's secret/configuration settings:
   - `APP_BASE_URL`: the public HTTPS origin, with no path (for example `https://requirements.example.com`)
   - `GITHUB_CLIENT_ID`
   - `GITHUB_CLIENT_SECRET`
   - `GITHUB_ALLOWED_USERS`: comma-separated GitHub logins permitted to use this service
   - `PORT`: usually supplied by the hosting provider; the service defaults to `3000`
4. Open the public service URL. Users sign in with GitHub and must have Copilot access enabled for their accounts.

For local development, use Node.js `22.12` or newer, copy `server/.env.example` to `server/.env`, fill in the OAuth app values and use a local callback such as `http://localhost:3000/auth/callback`. Then run `npm ci` and `npm start` from `server/`.

Do not commit `.env` files or place OAuth secrets in browser code. Do not configure a shared `COPILOT_GITHUB_TOKEN`; the SDK receives each signed-in user's token for that user's analysis session.

## Data, privacy, and service limits

- The selected document is uploaded to this cloud service. The service extracts its text and sends that text to GitHub Copilot under the signed-in user's account. Copilot usage is subject to that user's subscription/organization entitlement and usage limits.
- The HTTP handler does not write uploaded documents to disk, but the Copilot SDK runtime may store session state/transcripts on the service host. Idle runtime sessions are configured for cleanup after 15 minutes; do not treat this prototype as zero-retention. Review Copilot, host logging, retention, and company data policies before uploading real customer documents.
- OAuth tokens and app sessions are held in process memory, expire after inactivity, and are lost on service restart. This initial deployment is intended for a single service instance; production multi-instance deployments need a shared session store, stronger operational monitoring, and an organization-approved retention design.
- Files are limited to 10 MB and extracted text to 60,000 characters per document. Scanned PDFs need OCR and are not supported here.
- This app does not fetch files from SharePoint or write directly to Polarion.

GitHub Pages URLs remain a static-only version and cannot perform cloud Copilot analysis. The hosted Node service URL is the no-install entry point.

## Third-party code

`vendor/xlsx.full.min.js` is SheetJS Community Edition 0.18.5, licensed under Apache 2.0. The license text is in `vendor/LICENSE-Apache-2.0.txt`.
