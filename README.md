# Home Projects & Groceries — Mechanicus Data-Slate

An installable Progressive Web App (PWA) for managing shared household projects and department-grouped groceries backed by Google Sheets and Google Apps Script, secured with server-side passcode authentication.

## Features

- **Views**:
  - **Today**: Overdue and due-today tasks where status is not `Done`.
  - **All Projects**: Filterable by Category and Modifier tags, with inline category management (add/rename/remove), clickable Due Date badges, inline Modifier tag editor (`± Tags`), and expandable/editable Notes.
  - **Grocery List**: Grouped by store department with a one-tap **Clear Checked** button.
- **Security**: Server-side `APP_PASSCODE` validation in `Code.gs` on every `GET` and `POST` request; full-screen Lock Screen modal on the front end with `localStorage` persistence and a **Lock** button.
- **Mobile PWA**: Includes `manifest.json`, `sw.js` (shell caching with live network-only Google Sheets sync), custom Mechanicus cogwheel icons, and automatic refresh on app resume so two people stay in sync.
- **Print Mode**: Clean black-and-white `@media print` checklist formatting with `☐` / `☑` checkboxes.

## Setup & Deployment

### 1. Google Apps Script Backend (`Code.gs`)
1. Open your **Home Projects & Groceries** Google Sheet → **Extensions → Apps Script**.
2. Paste `Code.gs`, replace `'CHANGE_ME_BEFORE_DEPLOY'` with your shared household passphrase, and save.
3. Click **Deploy → Manage deployments → Edit → Version: New version → Deploy** (*Execute as: Me*, *Who has access: Anyone*).

### 2. GitHub Pages Hosting
1. Push this repository to GitHub.
2. Go to **Settings → Pages** → set Source to **Deploy from a branch** → select `main` / `/ (root)` → **Save**.
3. Open your GitHub Pages HTTPS URL on both phones:
   - **iOS (Safari)**: **Share → Add to Home Screen**
   - **Android (Chrome)**: Tap **+ Install** in the header (or menu **⋮ → Install app**)
4. Enter your shared passphrase once on each device to unlock.
