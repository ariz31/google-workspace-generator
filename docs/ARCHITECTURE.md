# Architecture

Google Workspace Generator is a root-level Google Apps Script productivity suite. It has two user interfaces backed by one `Code.gs` file.

## Runtime

- Backend: `Code.gs`
- Spreadsheet sidebar UI: `Sidebar.html`
- Standalone web app UI: `Index.html`
- Manifest: `appsscript.json`
- Platform: Google Apps Script V8 runtime
- Output: generated Google Drive, Docs, Sheets, Slides, and PDF files

## Workflows

### Spreadsheet-bound template generator

1. The spreadsheet opens and `onOpen()` adds the **Productivity Suite** menu.
2. The user runs **Setup Generator Sheets** or manually prepares `R-DOC-GEN`, `C-DOC-GEN`, and `EMAIL`.
3. The user selects rows or columns.
4. `Sidebar.html` calls `getSelectionDetails()` to inspect the active selection.
5. The user chooses output options and calls `runGenerator(options, sheetName)`.
6. `Code.gs` stores generation state in `PropertiesService`.
7. `continueGeneration()` processes batches and schedules follow-up triggers when needed.
8. The sidebar polls `getGenerationStatus()` until the run completes.
9. Outputs and errors are written to the sidebar and to the `Log` sheet.

### Standalone web app generator

1. The deployed web app loads `Index.html` from `doGet()`.
2. The UI calls `getDefaultWorkspaceConfig()` for sample JSON.
3. The user submits guided input or advanced JSON.
4. `generateWorkspace(config)` creates a Drive folder and basic Docs, Sheets, and Slides.
5. Resource links are returned to the browser.

## Spreadsheet data contracts

### `R-DOC-GEN`

- Column A: template file ID
- Column B: destination folder ID
- Column C: optional recipient email
- Column D: output name
- Column E onward: placeholder values with headers like `{{Name}}`

### `C-DOC-GEN`

- Row 1: template file ID per item column
- Row 2: destination folder ID
- Row 3: optional recipient email
- Row 4: output name
- Row 5 onward: placeholder values with labels like `{{Name}}` in column A

## Placeholder behavior

Supported placeholder forms:

- `{{Name}}` plain replacement
- `{{Name}u}}` uppercase
- `{{Name}b}}` bold
- `{{Name}ub}}` uppercase and bold
- `{{Name}1}}` title case
- `{{Name}1b}}` title case and bold

Docs and Slides also attempt to replace image placeholders when the value is an image URL or a Drive image file ID.

## Safety and reliability choices

- Generation runs in batches to reduce Apps Script timeout risk.
- State is stored per user through `PropertiesService.getUserProperties()`.
- Triggers are cleaned up before a new run and after completion or cancellation.
- Logging failures are swallowed so a bad log sheet does not break file generation.
- `Sidebar.html` escapes rendered output links and messages.
- Root files are the canonical clasp source; the old `src/` scaffold has been removed.

## Extension points

Useful future improvements:

1. Add automated tests for pure validation and placeholder helpers.
2. Add optional role-based sharing controls for generated files.
3. Add a preset template library.
4. Add a status dashboard sheet for long-running jobs.
5. Add CI that verifies root Apps Script files are present before deployment.
