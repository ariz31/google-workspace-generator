# Google Workspace Generator

A lightweight Google Apps Script web app for generating a structured Google Drive workspace with optional Google Docs, Sheets, and Slides resources.

This repository is intentionally simple: it uses Apps Script directly, keeps the UI in one `Index.html` file, and can be deployed with [`clasp`](https://github.com/google/clasp) or copied into the Apps Script editor.

## What it does

- Creates a dedicated Google Drive folder for a workspace.
- Generates configurable Google Docs, Sheets, and Slides files inside that folder.
- Provides a browser-based Apps Script UI for non-technical users.
- Supports both a guided form and advanced JSON configuration.
- Returns shareable links to all generated resources.
- Includes validation and safe defaults to avoid accidental large file generation.

## Repository structure

```text
.
├── README.md
├── package.json
├── .clasp.json.example
├── .editorconfig
├── .gitignore
├── docs/
│   ├── ARCHITECTURE.md
│   └── DEVELOPMENT.md
└── src/
    ├── Code.gs
    ├── Index.html
    └── appsscript.json
```

## Quick start with clasp

1. Install dependencies:

   ```bash
   npm install
   ```

2. Log in to Google Apps Script:

   ```bash
   npm run clasp:login
   ```

3. Create a new Apps Script project:

   ```bash
   npx clasp create --type webapp --title "Google Workspace Generator" --rootDir src
   ```

4. Push the source:

   ```bash
   npm run push
   ```

5. Open the Apps Script project:

   ```bash
   npm run open
   ```

6. Deploy as a web app from Apps Script:
   - Click **Deploy > New deployment**.
   - Select **Web app**.
   - Execute as: **Me**.
   - Who has access: choose the appropriate setting for your workspace.

## Quick start without clasp

1. Go to [script.google.com](https://script.google.com/).
2. Create a new Apps Script project.
3. Copy `src/Code.gs` into `Code.gs`.
4. Create an HTML file named `Index` and copy `src/Index.html` into it.
5. Copy settings from `src/appsscript.json` into the Apps Script manifest.
6. Deploy as a web app.

## Example advanced configuration

```json
{
  "workspaceName": "Physics 101 - Quarter 1",
  "description": "Generated class workspace",
  "documents": [
    { "name": "Syllabus", "body": "Course overview and expectations." },
    { "name": "Lesson Plan Template", "body": "Objectives\nMaterials\nActivities\nAssessment" }
  ],
  "spreadsheets": [
    {
      "name": "Gradebook",
      "sheets": [
        { "name": "Scores", "values": [["Student", "Activity", "Score"], ["", "", ""]] }
      ]
    }
  ],
  "presentations": [
    { "name": "Class Orientation", "title": "Welcome", "subtitle": "Generated with Google Workspace Generator" }
  ]
}
```

## Safety limits

The app includes conservative defaults:

- Maximum generated Docs: 20
- Maximum generated Sheets: 20
- Maximum generated Slides files: 20
- Maximum total generated files: 50
- Maximum file name length: 120 characters

These limits can be adjusted in `src/Code.gs` if needed.

## Recommended next improvements

- Add template duplication support for existing Docs, Sheets, and Slides.
- Add Google Drive sharing rules per generated workspace.
- Add audit logging to a central spreadsheet.
- Add unit tests with `gas-local` or a small pure JavaScript validation module.
- Add a CI workflow for linting and formatting.

## License

MIT. Update this section if you plan to use a different license.
