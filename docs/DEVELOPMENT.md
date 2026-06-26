# Development Guide

## Prerequisites

- Node.js installed locally.
- A Google account with Apps Script access.
- The clasp CLI, installed through this repo with `npm install`.

## Local setup

```bash
npm install
npm run clasp:login
```

Create or connect an Apps Script project:

```bash
npx clasp create --type webapp --title "Google Workspace Generator" --rootDir src
```

Or copy `.clasp.json.example` to `.clasp.json` and replace `YOUR_SCRIPT_ID_HERE` with an existing Apps Script project ID.

## Common commands

```bash
npm run push      # Push src/ to Apps Script
npm run pull      # Pull remote Apps Script files into src/
npm run open      # Open the Apps Script editor
npm run status    # Compare local and remote files
npm run deploy    # Create a clasp deployment
```

## Manual deployment

Inside the Apps Script editor:

1. Click **Deploy > New deployment**.
2. Choose **Web app**.
3. Set **Execute as** to **Me**.
4. Choose the appropriate access setting for your environment.
5. Click **Deploy** and authorize the required Google Workspace scopes.

## Testing checklist

Before deploying broadly, verify these flows:

- The web app loads without console errors.
- The guided form can create one Doc, one Sheet, and one Slides file.
- Advanced JSON can create custom file names and spreadsheet values.
- Invalid JSON shows a useful error message.
- Setting all file counts to zero is rejected.
- Generated files are moved into the generated Drive folder.
- Links in the result panel open the expected files.

## Code style

- Keep Apps Script functions small and explicit.
- Keep public server functions limited to UI entry points.
- Keep helper functions private by using the trailing underscore naming convention.
- Avoid requesting scopes that the app does not use.
- Prefer plain Apps Script services before introducing advanced services.

## Suggested future refactor

For a larger version, split pure validation logic into a shared JavaScript module and add tests around it. Apps Script itself is harder to test locally, so the best test target is the pure normalization and validation layer.
