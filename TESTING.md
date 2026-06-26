# Testing Checklist

Use this checklist after pushing or copying the root files into Google Apps Script.

## 1. Basic deployment

- Push the root files with `npm run push`, or copy these files manually:
  - `Code.gs`
  - `Runtime.gs`
  - `Sidebar.html`
  - `Index.html`
  - `appsscript.json`
- Open the Apps Script editor and confirm there are no syntax errors.
- Reload the bound Google Sheet.
- Confirm the **Productivity Suite** menu appears.

## 2. One-click demo setup

- Click **Productivity Suite > Setup Complete Demo Workspace**.
- Confirm the script creates or resets:
  - `R-DOC-GEN`
  - `C-DOC-GEN`
  - `EMAIL`
  - `Log`
- Confirm a Google Drive demo output folder was created.
- Confirm `R-DOC-GEN` row 1 contains formula-generated placeholders from row 2 readable labels.
- Confirm `C-DOC-GEN` column A contains formula-generated placeholders from column B readable labels.
- Confirm the generated sample rows/columns contain actual Template IDs and Folder IDs.

## 3. Health check

- Click **Productivity Suite > Run Health Check**.
- Expected result: health check passes, or shows only warnings that are understandable.
- If errors appear, fix them before generating.

## 4. Row-mode generation

- Go to `R-DOC-GEN`.
- Edit row 2 readable labels if needed; row 1 placeholders should update automatically.
- Select rows 3:5.
- Open **Productivity Suite > Open Generator Sidebar**.
- Confirm the sidebar detects selected templates.
- Generate in separate-file mode.
- Confirm output links appear in the sidebar.
- Confirm generated files appear in the demo Drive folder.
- Confirm the `Log` sheet contains success records.

## 5. Column-mode generation

- Go to `C-DOC-GEN`.
- Edit column B readable field names if needed; column A placeholders should update automatically.
- Select item columns C:E. Column B is only the readable field-name helper.
- Open the sidebar.
- Generate in separate-file mode.
- Confirm output links, Drive files, and log records.

## 6. Combined generation

Test these separately:

- Select only Doc template rows or item columns, then choose combined Doc/Slide mode.
- Select only Slide template rows or item columns, then choose combined Doc/Slide mode.
- Select only Sheet template rows or item columns, then choose combined Sheet mode.

Expected result: combined files are created without mixing incompatible template types.

## 7. PDF conversion

- Select Doc or Slide templates.
- Set format to PDF.
- Generate.
- Confirm the output is a PDF file in Drive.

## 8. Preflight validation

Test intentional bad data:

- Remove a Template ID.
- Remove a Folder ID.
- Use an invalid Template ID.
- Use an invalid Folder ID.
- Mix Docs and Slides in combined Doc/Slide mode.
- Select only column B in `C-DOC-GEN`; generation should reject it because column B is not an item column.

Expected result: generation should stop before creating files and show grouped errors.

## 9. Email sending

- Put your own email address in the Recipient Email field.
- Check **Send email when recipient exists** in the sidebar.
- Generate one file.
- Confirm the email is received and the `Log` sheet records the email status.

## 10. Reset and cancel

- Start a run with multiple selected rows or item columns.
- Click Cancel in the sidebar or use **Productivity Suite > Cancel Current Generation**.
- Then click **Productivity Suite > Reset Generator State**.
- Confirm a new generation can start normally afterward.

## 11. Standalone web app

- Deploy as a web app.
- Open the deployment URL.
- Generate one Doc, one Sheet, and one Slide.
- Confirm the output folder and links are returned in the UI.

## Expected production baseline

Before using with real templates, all of these should pass:

- Demo setup completes.
- Formula-generated placeholders update from readable labels.
- Health check passes.
- Row mode works.
- Column mode works.
- Validation blocks bad data before generation.
- Generated files are logged.
- Reset clears stuck state.
