# Google Workspace Generator

Google Workspace Generator is a Google Apps Script productivity suite for creating and batch-generating Google Workspace files.

It supports two workflows:

1. **Spreadsheet-bound template generator** through `Sidebar.html`
   - Select rows in `R-DOC-GEN` or columns in `C-DOC-GEN`.
   - Generate Docs, Slides, Sheets, or PDFs from template IDs.
   - Replace placeholders like `{{Name}}`, `{{Name}u}}`, `{{Name}b}}`, `{{Name}ub}}`, `{{Name}1}}`, and `{{Name}1b}}`.
   - Process separate files or combined Docs, Slides, and Sheets.
   - Send generated files by email when recipient data exists.
   - Log activity to a `Log` sheet.
   - Insert QR codes from the active cell.

2. **Standalone web app workspace generator** through `Index.html`
   - Create a Drive folder.
   - Generate basic Docs, Sheets, and Slides from a form or JSON configuration.
   - Return links to the generated resources.

## Repository structure

```text
.
├── Code.gs              # Main Apps Script backend
├── Sidebar.html         # Spreadsheet-bound generator sidebar
├── Index.html           # Standalone web app UI
├── appsscript.json      # Apps Script manifest and scopes
├── package.json         # clasp helper scripts
├── .clasp.json.example  # clasp config template
├── docs/
│   ├── ARCHITECTURE.md
│   └── DEVELOPMENT.md
└── LICENSE
```

## Quick start with clasp

```bash
npm install
npm run clasp:login
npx clasp create --type sheets --title "Google Workspace Generator" --rootDir .
npm run push
npm run open
```

For an existing Apps Script project, copy `.clasp.json.example` to `.clasp.json`, replace `YOUR_SCRIPT_ID_HERE`, then run:

```bash
npm run push
```

## Spreadsheet generator setup

1. Push or copy the root files into a Google Apps Script project bound to a Google Sheet.
2. Reload the spreadsheet.
3. Open **Productivity Suite > Setup Generator Sheets**.
4. Fill in the generated sheets:
   - `R-DOC-GEN`: one output per selected row.
   - `C-DOC-GEN`: one output per selected column.
   - `EMAIL`: optional subject/body templates.
5. Select the rows or columns to process.
6. Open **Productivity Suite > Open Generator Sidebar**.
7. Choose options and click **Generate**.

## Row mode format: `R-DOC-GEN`

| Column | Purpose |
| --- | --- |
| A | Template file ID |
| B | Destination Drive folder ID |
| C | Recipient email, optional |
| D | Output name |
| E+ | Placeholder values, with headers like `{{Name}}` |

## Column mode format: `C-DOC-GEN`

| Row | Purpose |
| --- | --- |
| 1 | Template file ID per item column |
| 2 | Destination Drive folder ID |
| 3 | Recipient email, optional |
| 4 | Output name |
| 5+ | Placeholder values, with row labels like `{{Name}}` |

## Standalone web app setup

Deploy the same project as a web app:

1. Open Apps Script.
2. Click **Deploy > New deployment**.
3. Select **Web app**.
4. Execute as **Me**.
5. Choose the access level appropriate for your workspace.

`Index.html` calls `getDefaultWorkspaceConfig()` and `generateWorkspace(config)` from `Code.gs`.

## Example standalone JSON

```json
{
  "workspaceName": "Physics 101 - Quarter 1",
  "description": "Generated class workspace",
  "documents": [
    { "name": "Syllabus", "body": "Course overview and expectations." }
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

## Notes

- `appsscript.json` includes Drive, Docs, Sheets, Slides, Mail, external request, UI, and trigger scopes because the full spreadsheet generator uses all of them.
- The web app manifest currently uses `ANYONE` access. Change it before deployment if you need domain-only or private access.
- Keep `Code.gs`, `Sidebar.html`, `Index.html`, and `appsscript.json` at the repository root. The root is now the canonical clasp source.

## License

MIT.
