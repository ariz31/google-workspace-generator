# Architecture

Google Workspace Generator is a small Google Apps Script web app.

## Runtime

- Frontend: `src/Index.html`
- Backend: `src/Code.gs`
- Platform: Google Apps Script V8 runtime
- Output: generated files in Google Drive

## Request flow

1. The user opens the deployed web app.
2. `Index.html` renders the form.
3. The UI calls `generateWorkspace(config)` with `google.script.run`.
4. `Code.gs` validates and normalizes the input.
5. A Drive folder is created.
6. Docs, Sheets, and Slides files are created.
7. Files are moved into the generated folder.
8. Resource links are returned to the UI.

## Backend responsibilities

`Code.gs` handles:

- The web app entry point through `doGet()`.
- Sample configuration through `getDefaultConfig()`.
- Workspace creation through `generateWorkspace(config)`.
- Input normalization and validation.
- Google Drive, Docs, Sheets, and Slides file creation.
- Returning a structured result object to the UI.

## Frontend responsibilities

`Index.html` handles:

- Guided workspace configuration.
- Advanced JSON editing.
- Client-side validation.
- Backend calls with `google.script.run`.
- Rendering generated file links.

## Configuration shape

```json
{
  "workspaceName": "Generated Workspace",
  "description": "Created with Google Workspace Generator.",
  "documents": [
    { "name": "Project Brief", "body": "Document body" }
  ],
  "spreadsheets": [
    {
      "name": "Tracker",
      "sheets": [
        { "name": "Tasks", "values": [["Task", "Owner"], ["", ""]] }
      ]
    }
  ],
  "presentations": [
    { "name": "Overview Deck", "title": "Generated Workspace", "subtitle": "Created with Google Workspace Generator" }
  ]
}
```

## Safety limits

The generator uses limits in `Code.gs` to prevent accidental quota-heavy runs:

- File count limits by type and total.
- File name length limits.
- Description and document body length limits.
- Spreadsheet row and column limits.

## Extension points

Useful next improvements:

1. Add template duplication support with optional template IDs.
2. Add optional sharing rules for generated files.
3. Add audit logging to a spreadsheet.
4. Add preset workspace types for classes, projects, reports, or events.
