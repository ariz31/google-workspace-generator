/**
 * Deployment hardening layer.
 * Public functions here intentionally override earlier public functions so a
 * deployed project can set itself up, validate itself, and start generation
 * without the lock handoff bug in the original Code.gs runGenerator.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Productivity Suite')
    .addItem('Open Generator Sidebar', 'showSidebar')
    .addItem('Create QR Code from Active Cell', 'insertQRCode')
    .addSeparator()
    .addItem('Setup Complete Demo Workspace', 'setupGeneratorSheets')
    .addItem('Run Health Check', 'runHealthCheck')
    .addItem('Reset Generator State', 'resetGeneratorState')
    .addItem('Cancel Current Generation', 'cancelGeneration')
    .addToUi();
}

function setupGeneratorSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this from a Google Sheet to set up generator sheets.');

  resetGeneratorState();
  const folder = DriveApp.createFolder(`${APP_NAME} Demo Output - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')}`);
  const templates = _createDemoTemplatesRuntime_(folder);

  _setupRowModeSheetRuntime_(ss, folder, templates);
  _setupColumnModeSheetRuntime_(ss, folder, templates);
  _setupEmailSheetRuntime_(ss);
  _getLogSheet();
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Complete demo workspace created.\n\n' +
    `Output folder: ${folder.getUrl()}\n\n` +
    'Next: select rows 3:5 in R-DOC-GEN or item columns C:E in C-DOC-GEN, open the sidebar, and click Generate.'
  );
}

function runHealthCheck() {
  const report = getHealthCheckReport();
  try {
    SpreadsheetApp.getUi().alert(_formatHealthReportRuntime_(report));
  } catch (error) {}
  return report;
}

function getHealthCheckReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = { ok: true, errors: [], warnings: [], checks: [] };

  _addHealthCheckRuntime_(report, Boolean(ss), 'Spreadsheet is active', 'Open this project from a bound Google Sheet.');
  if (!ss) return report;

  [ROW_MODE_SHEET, COLUMN_MODE_SHEET, EMAIL_SHEET, LOG_SHEET].forEach((name) => {
    _addHealthCheckRuntime_(report, Boolean(ss.getSheetByName(name)), `Sheet exists: ${name}`, `Missing sheet: ${name}. Run setup first.`);
  });

  _validateSheetContractRuntime_(report, ss, ROW_MODE_SHEET);
  _validateSheetContractRuntime_(report, ss, COLUMN_MODE_SHEET);

  const activeSheet = ss.getActiveSheet();
  if (activeSheet && (activeSheet.getName() === ROW_MODE_SHEET || activeSheet.getName() === COLUMN_MODE_SHEET)) {
    const selected = activeSheet.getName() === ROW_MODE_SHEET
      ? Array.from(_getSelectedRows(ss, ROW_MODE_SHEET))
      : Array.from(_getSelectedColumns(ss, COLUMN_MODE_SHEET));
    if (selected.length) {
      const validation = _validateGenerationRequestRuntime_({}, activeSheet.getName());
      validation.errors.forEach((message) => report.errors.push(message));
      validation.warnings.forEach((message) => report.warnings.push(message));
      report.ok = report.ok && validation.errors.length === 0;
      report.checks.push({ ok: validation.errors.length === 0, label: 'Current selection is generation-ready' });
    } else {
      report.warnings.push(activeSheet.getName() === COLUMN_MODE_SHEET
        ? 'No item columns selected. Select C:E or later in C-DOC-GEN.'
        : 'No active row selection to validate.');
    }
  } else {
    report.warnings.push(`Active sheet is not ${ROW_MODE_SHEET} or ${COLUMN_MODE_SHEET}.`);
  }

  try {
    const response = UrlFetchApp.fetch('https://quickchart.io/qr?text=health-check&size=80', { muteHttpExceptions: true });
    _addHealthCheckRuntime_(report, response.getResponseCode() >= 200 && response.getResponseCode() < 300, 'External request scope works', 'External request test failed. QR/image URLs may not work.');
  } catch (error) {
    report.warnings.push(`External request test failed: ${error.message}`);
  }

  report.ok = report.ok && report.errors.length === 0;
  return report;
}

function resetGeneratorState() {
  _deleteTrigger();
  PropertiesService.getUserProperties().deleteAllProperties();
  return { status: 'reset', message: 'Generator state and pending triggers were cleared.' };
}

function runGenerator(rawOptions, sheetName) {
  _deleteTrigger();
  PropertiesService.getUserProperties().deleteProperty(STATE_KEY);

  const validation = _validateGenerationRequestRuntime_(rawOptions || {}, sheetName);
  if (validation.errors.length) {
    throw new Error('Fix these issues before generating:\n- ' + validation.errors.join('\n- '));
  }

  _configureCombinedRunIfNeeded(validation.state, validation.itemTypes, validation.sheetData);
  PropertiesService.getUserProperties().setProperty(STATE_KEY, JSON.stringify(validation.state));

  continueGeneration();
  const updatedState = getGenerationStatus();
  return { status: 'started', totalItems: updatedState.totalItems, currentIndex: updatedState.currentIndex, state: updatedState, warnings: validation.warnings };
}

function getGenerationStatus() {
  const raw = PropertiesService.getUserProperties().getProperty(STATE_KEY);
  if (!raw) return { status: 'idle', currentIndex: 0, totalItems: 0, results: [], folders: [], progress: 0 };
  const state = JSON.parse(raw);
  state.progress = state.totalItems ? Math.round((state.currentIndex / state.totalItems) * 100) : 0;
  return state;
}

function getSelectionDetails() {
  const base = _getSelectionDetailsBaseRuntime_();
  if (!base.selectedIndices || !base.selectedIndices.length) return base;

  try {
    const validation = _validateGenerationRequestRuntime_({}, base.sheetName);
    base.errors = validation.errors;
    base.warnings = validation.warnings;
    if (validation.errors.length) base.message = `Selection has ${validation.errors.length} issue(s).`;
    else if (validation.warnings.length) base.message = `Selection is usable with ${validation.warnings.length} warning(s).`;
    else base.message = `Ready: ${base.selectedIndices.length} selected item(s).`;
  } catch (error) {
    base.errors = [error.message];
    base.message = error.message;
  }

  return base;
}

function _validateGenerationRequestRuntime_(rawOptions, requestedSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { errors: ['No active spreadsheet found.'], warnings: [], state: null, itemTypes: null, sheetData: null };

  const sheetName = requestedSheetName || ss.getActiveSheet().getName();
  const sheet = ss.getSheetByName(sheetName);
  const errors = [];
  const warnings = [];

  if (!sheet) errors.push(`Sheet not found: ${sheetName}.`);
  if (sheetName !== ROW_MODE_SHEET && sheetName !== COLUMN_MODE_SHEET) errors.push(`Use ${ROW_MODE_SHEET} or ${COLUMN_MODE_SHEET}.`);
  if (errors.length) return { errors, warnings, state: null, itemTypes: null, sheetData: null };

  const options = _normalizeRunOptions(rawOptions || {});
  const selectedItems = sheetName === ROW_MODE_SHEET
    ? Array.from(_getSelectedRows(ss, sheetName)).sort((a, b) => a - b)
    : Array.from(_getSelectedColumns(ss, sheetName)).sort((a, b) => a - b);

  if (!selectedItems.length) {
    errors.push(sheetName === COLUMN_MODE_SHEET
      ? `No valid item columns selected in ${sheetName}. Select column C or later; column B is only for readable field names.`
      : `No valid selection found in ${sheetName}.`);
  }

  const sheetData = sheet.getDataRange().getDisplayValues();
  const placeholderMap = sheetName === ROW_MODE_SHEET ? _buildPlaceholderMap(sheetData[0] || [], 4) : _buildColumnPlaceholderMap(sheetData, 4);
  if (!Object.keys(placeholderMap).length) errors.push('No placeholders found. Use readable field labels so the setup formulas can create {{Name}} labels.');

  const itemTypes = { docs: [], slides: [], sheets: [], unsupported: [] };
  const relevantData = [];
  const folders = [];

  selectedItems.forEach((itemNum) => {
    const label = sheetName === ROW_MODE_SHEET ? `Row ${itemNum}` : `Column ${_columnNumberToLetter(itemNum)}`;
    const itemData = sheetName === ROW_MODE_SHEET ? sheetData[itemNum - 1] || [] : sheetData.map((row) => row[itemNum - 1] || '');
    relevantData.push(itemData);

    const templateId = _safeTrim(itemData[0]);
    const folderId = _safeTrim(itemData[1]);
    const recipient = _safeTrim(itemData[2]);

    if (!templateId) errors.push(`${label}: missing Template ID.`);
    if (!folderId) errors.push(`${label}: missing Folder ID.`);
    if (recipient && !/^\S+@\S+\.\S+$/.test(recipient)) warnings.push(`${label}: recipient email looks invalid: ${recipient}.`);

    if (templateId) {
      try {
        const file = DriveApp.getFileById(templateId);
        const mimeType = file.getMimeType();
        if (mimeType === APP_MIME.GOOGLE_DOCS) itemTypes.docs.push(itemNum);
        else if (mimeType === APP_MIME.GOOGLE_SLIDES) itemTypes.slides.push(itemNum);
        else if (mimeType === APP_MIME.GOOGLE_SHEETS) itemTypes.sheets.push(itemNum);
        else {
          itemTypes.unsupported.push(itemNum);
          errors.push(`${label}: unsupported template type: ${mimeType}.`);
        }
      } catch (error) {
        errors.push(`${label}: template ID is not accessible or invalid.`);
      }
    }

    if (folderId) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        folders.push({ id: folderId, url: folder.getUrl() });
      } catch (error) {
        errors.push(`${label}: folder ID is not accessible or invalid.`);
      }
    }
  });

  if (options.docSlideOutputMode === 'combined' && itemTypes.docs.length && itemTypes.slides.length) errors.push('Combined Doc/Slide mode cannot mix Docs and Slides. Select only one type.');
  if (options.sheetOutputMode === 'combined' && (itemTypes.docs.length || itemTypes.slides.length) && itemTypes.sheets.length) warnings.push('Combined Sheet mode will only combine selected Sheet templates; Doc/Slide items will be ignored by the combined run.');
  if (options.format === 'PDF' && itemTypes.sheets.length && !itemTypes.docs.length && !itemTypes.slides.length) warnings.push('PDF conversion applies to Docs and Slides only. Sheet templates will stay as Sheets.');

  const state = {
    status: 'running',
    startedAt: new Date().toISOString(),
    items: selectedItems,
    currentIndex: 0,
    totalItems: selectedItems.length,
    options,
    sheetName,
    results: [],
    folders: _dedupeFoldersRuntime_(folders),
    placeholderMap,
    emailTemplates: _getEmailTemplates(options.sendEmail),
    relevantData,
    itemTypes,
    isCombinedRun: false,
    preflightWarnings: warnings,
  };

  return { errors, warnings, state, itemTypes, sheetData };
}

function _getSelectionDetailsBaseRuntime_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { sheetName: '', mimeTypes: [], selectedIndices: [], invalidItems: [], message: 'No active spreadsheet found.' };
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  if (sheetName !== ROW_MODE_SHEET && sheetName !== COLUMN_MODE_SHEET) return { sheetName, mimeTypes: [], selectedIndices: [], invalidItems: [], message: `Switch to ${ROW_MODE_SHEET} or ${COLUMN_MODE_SHEET}.` };

  const selectedIndices = sheetName === ROW_MODE_SHEET ? Array.from(_getSelectedRows(ss, sheetName)).sort((a, b) => a - b) : Array.from(_getSelectedColumns(ss, sheetName)).sort((a, b) => a - b);
  const data = sheet.getDataRange().getDisplayValues();
  const mimeTypes = new Set();
  const invalidItems = [];

  selectedIndices.forEach((itemNum) => {
    const templateId = sheetName === ROW_MODE_SHEET ? _safeTrim(data[itemNum - 1] && data[itemNum - 1][0]) : _safeTrim(data[0] && data[0][itemNum - 1]);
    if (!templateId) {
      invalidItems.push(itemNum);
      return;
    }
    try {
      mimeTypes.add(DriveApp.getFileById(templateId).getMimeType());
    } catch (error) {
      invalidItems.push(itemNum);
    }
  });

  return { sheetName, mimeTypes: Array.from(mimeTypes), selectedIndices, invalidItems, message: selectedIndices.length ? '' : sheetName === COLUMN_MODE_SHEET ? 'Select one or more item columns starting at C. Column B is only for readable field names.' : 'Select one or more valid rows to process.' };
}

function _getSelectedColumns(ss, sheetName) {
  const selectedColumns = new Set();
  const rangeList = ss.getActiveRangeList();
  if (!rangeList) return selectedColumns;
  const minimumColumn = sheetName === COLUMN_MODE_SHEET ? 3 : 2;

  rangeList.getRanges().forEach((range) => {
    if (range.getSheet().getName() !== sheetName) return;
    for (let i = 0; i < range.getNumColumns(); i++) {
      const column = range.getColumn() + i;
      if (column >= minimumColumn) selectedColumns.add(column);
    }
  });

  return selectedColumns;
}

function _createDemoTemplatesRuntime_(folder) {
  const doc = DocumentApp.create('Demo Certificate Template');
  doc.getBody().clear();
  doc.getBody().appendParagraph('Certificate for {{Name}1b}}').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  doc.getBody().appendParagraph('Activity: {{Activity}}');
  doc.getBody().appendParagraph('Score: {{Score}}');
  doc.getBody().appendParagraph('Remarks: {{Remarks}}');
  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(folder);

  const presentation = SlidesApp.create('Demo Slide Template');
  const slide = presentation.getSlides()[0];
  slide.getPageElements().forEach((element) => element.remove());
  slide.insertTextBox('Report for {{Name}1b}}', 60, 80, 600, 70);
  slide.insertTextBox('Activity: {{Activity}}\nScore: {{Score}}\nRemarks: {{Remarks}}', 60, 170, 600, 120);
  presentation.saveAndClose();
  const slideFile = DriveApp.getFileById(presentation.getId());
  slideFile.moveTo(folder);

  const spreadsheet = SpreadsheetApp.create('Demo Sheet Template');
  const templateSheet = spreadsheet.getSheets()[0];
  templateSheet.setName('Report');
  templateSheet.getRange(1, 1, 5, 2).setValues([
    ['Name', '{{Name}}'],
    ['Activity', '{{Activity}}'],
    ['Score', '{{Score}}'],
    ['Remarks', '{{Remarks}}'],
    ['Generated', new Date()],
  ]);
  templateSheet.autoResizeColumns(1, 2);
  SpreadsheetApp.flush();
  const sheetFile = DriveApp.getFileById(spreadsheet.getId());
  sheetFile.moveTo(folder);

  return { docId: doc.getId(), slideId: presentation.getId(), sheetId: spreadsheet.getId() };
}

function _setupRowModeSheetRuntime_(ss, folder, templates) {
  const sheet = _getOrCreateSheet(ss, ROW_MODE_SHEET);
  sheet.clear();
  sheet.getRange(1, 1, 1, 8).setValues([['Template ID', 'Folder ID', 'Recipient Email', 'Output Name', '=IF(LEN(TRIM(E2)),"{{"&TRIM(E2)&"}}","")', '=IF(LEN(TRIM(F2)),"{{"&TRIM(F2)&"}}","")', '=IF(LEN(TRIM(G2)),"{{"&TRIM(G2)&"}}","")', '=IF(LEN(TRIM(H2)),"{{"&TRIM(H2)&"}}","")']]);
  sheet.getRange(2, 1, 1, 8).setValues([['Template ID', 'Folder ID', 'Recipient Email', 'Output Name', 'Name', 'Activity', 'Score', 'Remarks']]);
  sheet.getRange(3, 1, 3, 8).setValues([
    [templates.docId, folder.getId(), '', 'Demo Document Output', 'Juan Dela Cruz', 'Activity 1', '95', 'Excellent work'],
    [templates.slideId, folder.getId(), '', 'Demo Slides Output', 'Maria Santos', 'Activity 2', '92', 'Ready for presentation'],
    [templates.sheetId, folder.getId(), '', 'Demo Sheet Output', 'Pedro Reyes', 'Activity 3', '88', 'For review'],
  ]);
  sheet.getRange('E1:H1').setNote('Generated from readable labels in row 2. Edit row 2 labels to update placeholders automatically.');
  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, 8);
}

function _setupColumnModeSheetRuntime_(ss, folder, templates) {
  const sheet = _getOrCreateSheet(ss, COLUMN_MODE_SHEET);
  sheet.clear();
  sheet.getRange(1, 1, 8, 5).setValues([
    ['Template ID', 'Readable Field Name', templates.docId, templates.slideId, templates.sheetId],
    ['Folder ID', '', folder.getId(), folder.getId(), folder.getId()],
    ['Recipient Email', '', '', '', ''],
    ['Output Name', '', 'Column Demo Document', 'Column Demo Slides', 'Column Demo Sheet'],
    ['=IF(LEN(TRIM(B5)),"{{"&TRIM(B5)&"}}","")', 'Name', 'Juan Dela Cruz', 'Maria Santos', 'Pedro Reyes'],
    ['=IF(LEN(TRIM(B6)),"{{"&TRIM(B6)&"}}","")', 'Activity', 'Column Activity 1', 'Column Activity 2', 'Column Activity 3'],
    ['=IF(LEN(TRIM(B7)),"{{"&TRIM(B7)&"}}","")', 'Score', '95', '92', '88'],
    ['=IF(LEN(TRIM(B8)),"{{"&TRIM(B8)&"}}","")', 'Remarks', 'Excellent work', 'Ready for presentation', 'For review'],
  ]);
  sheet.getRange('A5:A8').setNote('Generated from readable labels in column B. Edit column B labels to update placeholders automatically.');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.autoResizeColumns(1, 5);
}

function _setupEmailSheetRuntime_(ss) {
  const sheet = _getOrCreateSheet(ss, EMAIL_SHEET);
  sheet.clear();
  sheet.getRange(1, 1, 3, 2).setValues([
    ['Email Template Field', 'Value'],
    ['Subject', 'Generated file for {{Name}1}}'],
    ['Body', '<p>Hello {{Name}1}},</p><p>Your generated file for <b>{{Activity}}</b> is attached.</p>'],
  ]);
  sheet.autoResizeColumns(1, 2);
}

function _validateSheetContractRuntime_(report, ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const data = sheet.getDataRange().getDisplayValues();
  if (sheetName === ROW_MODE_SHEET) {
    _addHealthCheckRuntime_(report, data[0] && data[0][0] === 'Template ID', 'R-DOC-GEN has Template ID header', 'R-DOC-GEN column A should be Template ID.');
    _addHealthCheckRuntime_(report, data[0] && data[0][1] === 'Folder ID', 'R-DOC-GEN has Folder ID header', 'R-DOC-GEN column B should be Folder ID.');
    _addHealthCheckRuntime_(report, Object.keys(_buildPlaceholderMap(data[0] || [], 4)).length > 0, 'R-DOC-GEN has formula-generated placeholders', 'Add readable labels in row 2 from column E onward.');
  } else if (sheetName === COLUMN_MODE_SHEET) {
    _addHealthCheckRuntime_(report, data[0] && data[0][0] === 'Template ID', 'C-DOC-GEN has Template ID row', 'C-DOC-GEN row 1 should be Template ID.');
    _addHealthCheckRuntime_(report, data[1] && data[1][0] === 'Folder ID', 'C-DOC-GEN has Folder ID row', 'C-DOC-GEN row 2 should be Folder ID.');
    _addHealthCheckRuntime_(report, data[0] && data[0][1] === 'Readable Field Name', 'C-DOC-GEN has readable field-name helper column', 'C-DOC-GEN column B should contain readable field names.');
    _addHealthCheckRuntime_(report, Object.keys(_buildColumnPlaceholderMap(data, 4)).length > 0, 'C-DOC-GEN has formula-generated placeholders', 'Add readable labels in column B from row 5 downward.');
  }
}

function _addHealthCheckRuntime_(report, ok, label, errorMessage) {
  report.checks.push({ ok: Boolean(ok), label });
  if (!ok) {
    report.ok = false;
    report.errors.push(errorMessage || label);
  }
}

function _formatHealthReportRuntime_(report) {
  const lines = [];
  lines.push(report.ok ? 'Health check passed.' : 'Health check found issues.');
  lines.push('');
  lines.push(`Checks passed: ${report.checks.filter((check) => check.ok).length}/${report.checks.length}`);
  if (report.errors.length) lines.push('\nErrors:\n- ' + report.errors.join('\n- '));
  if (report.warnings.length) lines.push('\nWarnings:\n- ' + report.warnings.join('\n- '));
  return lines.join('\n');
}

function _dedupeFoldersRuntime_(folders) {
  const seen = {};
  return folders.filter((folder) => {
    if (!folder || !folder.id || seen[folder.id]) return false;
    seen[folder.id] = true;
    return true;
  });
}
