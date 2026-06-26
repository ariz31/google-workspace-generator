# Development Guide

## Prerequisites

- Node.js
- Google Apps Script access
- clasp, installed with `npm install`

## Local setup

```bash
npm install
npm run clasp:login
npx clasp create --type sheets --title "Google Workspace Generator" --rootDir .
npm run push
npm run open
```

For an existing Apps Script project, copy `.clasp.json.example` to `.clasp.json`, replace `YOUR_SCRIPT_ID_HERE`, then run `npm run push`.

## Common commands

```bash
npm run push
npm run pull
npm run open
npm run status
npm run deploy
```

## Root files

The canonical Apps Script files are now at the repository root:

- `Code.gs`
- `Sidebar.html`
- `Index.html`
- `appsscript.json`

Do not use a `src/` root for clasp unless you intentionally restructure the project again.

## Manual deployment

For spreadsheet-bound use, bind or copy the root files into a Google Sheet Apps Script project. Reload the spreadsheet and use the **Productivity Suite** menu.

For web-app use, deploy the same project as a web app from Apps Script. `Index.html` is served by `doGet()`.

## Testing checklist

- Spreadsheet reload shows the menu.
- Setup creates `R-DOC-GEN`, `C-DOC-GEN`, `EMAIL`, and `Log`.
- QR generation inserts an image from the selected cell.
- Sidebar detects row selections in `R-DOC-GEN`.
- Sidebar detects column selections in `C-DOC-GEN`.
- Separate Doc, Slide, and Sheet generation works.
- PDF conversion works for Docs and Slides.
- Combined Doc, Slide, and Sheet generation works for same-type selections.
- Email sending works when recipient and template data exist.
- Cancel stops a running batch.
- Standalone web app generation works from `Index.html`.

## Code style

- Keep root files as the source of truth.
- Keep HTML-called server functions public.
- Keep helpers private with trailing underscores.
- Keep logging non-blocking.
- Avoid unnecessary OAuth scopes.
