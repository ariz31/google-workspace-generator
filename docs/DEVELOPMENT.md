# Development Guide

This project is intended to be installed directly inside Google Apps Script. There is no local tool setup requirement.

## Prerequisites

- Google account with Apps Script access
- A Google Sheet for the spreadsheet-bound generator
- Access to the Drive templates and destination folders you plan to use

## Manual setup

1. Create or open a Google Sheet.
2. Open **Extensions > Apps Script**.
3. Copy the root `Code.gs` into the Apps Script `Code.gs` file.
4. Add an HTML file named `Sidebar` and copy `Sidebar.html` into it.
5. Add an HTML file named `Index` and copy `Index.html` into it.
6. Add or update the Apps Script manifest with `appsscript.json`.
7. Save the project.
8. Run `setupGeneratorSheets()` once and approve permissions.
9. Reload the spreadsheet and use the **Productivity Suite** menu.

## Root files

The canonical Apps Script files are at the repository root:

- `Code.gs`
- `Sidebar.html`
- `Index.html`
- `appsscript.json`

Keep backend helper functions inside `Code.gs` unless the project is intentionally restructured later.

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
