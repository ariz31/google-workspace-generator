# Google Workspace Generator

Google Workspace Generator is a Google Apps Script productivity suite for creating and batch-generating Google Workspace files.

This repository is documented with a **non-technical, copy-paste setup first**, similar to the peer-evaluation project. The existing app features remain intact.

## Files to copy into Apps Script

For the simplest setup, copy these root files into one Google Apps Script project:

- `Code.gs` - main backend logic and helper functions
- `Sidebar.html` - spreadsheet-bound generator sidebar
- `Index.html` - standalone web app UI and user guide
- `appsscript.json` - required Apps Script manifest and permissions

No extra `.gs` helper files are required. The backend logic stays in `Code.gs`.

## What the app supports

### Spreadsheet-bound template generator

Use this workflow from a Google Sheet through `Sidebar.html`.

- Generate Google Docs, Slides, Sheets, or PDFs from template IDs.
- Use `R-DOC-GEN` row mode or `C-DOC-GEN` column mode.
- Replace placeholders such as `{{Name}}`, `{{Name}u}}`, `{{Name}b}}`, `{{Name}ub}}`, `{{Name}1}}`, and `{{Name}1b}}`.
- Preserve formatting behavior supported by the current generator.
- Generate separate files or combined Docs, Slides, and Sheets.
- Send generated files by email when recipient data exists.
- Log generated file links, folder links, status, and email status to `Log`.
- Insert QR codes from the active cell.
- Cancel/reset long-running generation state when needed.

### Standalone web app workspace generator

Use this workflow from `Index.html` after deploying the project as a web app.

- Create a Drive folder.
- Generate basic Docs, Sheets, and Slides from the form or advanced JSON.
- Return folder and file links in the page.

## Non-technical setup

1. Create or open a Google Sheet.
2. Open **Extensions > Apps Script**.
3. Replace the default `Code.gs` with this repository's `Code.gs`.
4. Add an HTML file named `Sidebar` and paste `Sidebar.html`.
5. Add an HTML file named `Index` and paste `Index.html`.
6. Add or replace the manifest with this repository's `appsscript.json`.
7. Save the Apps Script project.
8. Run `setupGeneratorSheets()` once from the Apps Script editor and approve permissions.
9. Reload the Google Sheet.
10. Use the **Productivity Suite** menu.

## Daily spreadsheet workflow

1. Open the Google Sheet.
2. Use **Productivity Suite > Setup Generator Sheets** if the setup sheets do not exist yet.
3. Fill in `R-DOC-GEN` or `C-DOC-GEN`.
4. Select the rows or columns to process.
5. Open **Productivity Suite > Open Generator Sidebar**.
6. Choose output options.
7. Click **Generate**.
8. Review links and statuses in the sidebar or the `Log` sheet.

## Row mode format: `R-DOC-GEN`

| Column | Purpose |
| --- | --- |
| A | Template file ID |
| B | Destination Drive folder ID |
| C | Recipient email, optional |
| D | Output name |
| E+ | Placeholder values, with headers like `{{Name}}` |

Use row mode when each selected row should create one output file.

## Column mode format: `C-DOC-GEN`

| Row | Purpose |
| --- | --- |
| 1 | Template file ID per item column |
| 2 | Destination Drive folder ID |
| 3 | Recipient email, optional |
| 4 | Output name |
| 5+ | Placeholder values, with row labels like `{{Name}}` |

Use column mode when each selected column should create one output file.

## Email setup

The `EMAIL` sheet controls optional email delivery.

- Add the email subject and body template there.
- Put recipient emails in `R-DOC-GEN` column C or `C-DOC-GEN` row 3.
- Enable email delivery in the sidebar.
- The app logs email status in the `Log` sheet.

## Standalone web app setup

Deploy the same project as a web app:

1. Open Apps Script.
2. Click **Deploy > New deployment**.
3. Select **Web app**.
4. Execute as **Me**.
5. Choose the access level appropriate for your workspace.

`Index.html` calls `getDefaultWorkspaceConfig()` and `generateWorkspace(config)` from `Code.gs`.

## Notes

- `appsscript.json` includes Drive, Docs, Sheets, Slides, Mail, external request, UI, and trigger scopes because the full spreadsheet generator uses them.
- The web app manifest currently uses `ANYONE` access. Change it before deployment if you need domain-only or private access.
- Keep `Code.gs`, `Sidebar.html`, `Index.html`, and `appsscript.json` at the repository root.
- Do not split helper functions into extra `.gs` files unless the project is intentionally restructured later.

## License

MIT.
