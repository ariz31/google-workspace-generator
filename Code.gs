/**
 * Google Workspace Generator / Productivity Suite
 *
 * Root Apps Script backend for:
 * - QR code insertion from the active spreadsheet cell.
 * - Spreadsheet-bound template generation from R-DOC-GEN and C-DOC-GEN sheets.
 * - Async batched generation for Docs, Slides, Sheets, PDF conversion, email, logging, and combined output.
 * - Standalone web-app workspace generation through Index.html.
 *
 * Required scopes are declared in appsscript.json.
 */

const APP_NAME = 'Google Workspace Generator';
const STATE_KEY = 'generationState';
const TRIGGER_HANDLER = 'continueGeneration';
const ROW_MODE_SHEET = 'R-DOC-GEN';
const COLUMN_MODE_SHEET = 'C-DOC-GEN';
const EMAIL_SHEET = 'EMAIL';
const LOG_SHEET = 'Log';

const APP_MIME = Object.freeze({
  GOOGLE_DOCS: 'application/vnd.google-apps.document',
  GOOGLE_SLIDES: 'application/vnd.google-apps.presentation',
  GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet',
});

const GENERATOR_LIMITS = Object.freeze({
  minBatchSize: 1,
  maxBatchSize: 50,
  triggerDelayMs: 60 * 1000,
  batchTimeLimitMs: 270 * 1000,
  maxSimpleDocs: 20,
  maxSimpleSheets: 20,
  maxSimpleSlides: 20,
  maxSimpleFiles: 50,
  maxNameLength: 120,
  maxDescriptionLength: 500,
  maxDocumentBodyLength: 20000,
  maxSimpleSheetRows: 500,
  maxSimpleSheetColumns: 50,
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Productivity Suite')
    .addItem('Open Generator Sidebar', 'showSidebar')
    .addItem('Create QR Code from Active Cell', 'insertQRCode')
    .addSeparator()
    .addItem('Setup Generator Sheets', 'setupGeneratorSheets')
    .addItem('Cancel Current Generation', 'cancelGeneration')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Generator Suite')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function setupGeneratorSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this from a Google Sheet to set up generator sheets.');

  const rowSheet = _getOrCreateSheet(ss, ROW_MODE_SHEET);
  rowSheet.clear();
  rowSheet.getRange(1, 1, 1, 5).setValues([['Template ID', 'Folder ID', 'Recipient Email', 'Output Name', '{{Sample}}']]);
  rowSheet.getRange(2, 1, 1, 5).setValues([['Paste a Docs/Slides/Sheets template file ID below', 'Paste target Drive folder ID below', 'optional@email.com', 'Generated file name', 'Replacement value']]);
  rowSheet.getRange(3, 1, 1, 5).setValues([['', '', '', 'Example Output', 'Hello world']]);
  rowSheet.setFrozenRows(2);
  rowSheet.autoResizeColumns(1, 5);

  const columnSheet = _getOrCreateSheet(ss, COLUMN_MODE_SHEET);
  columnSheet.clear();
  columnSheet.getRange(1, 1, 5, 3).setValues([
    ['Field', 'Item 1', 'Item 2'],
    ['Template ID', '', ''],
    ['Folder ID', '', ''],
    ['Recipient Email', '', ''],
    ['Output Name', 'Example Output 1', 'Example Output 2'],
  ]);
  columnSheet.getRange(6, 1, 1, 3).setValues([['{{Sample}}', 'Hello item 1', 'Hello item 2']]);
  columnSheet.setFrozenRows(1);
  columnSheet.setFrozenColumns(1);
  columnSheet.autoResizeColumns(1, 3);

  const emailSheet = _getOrCreateSheet(ss, EMAIL_SHEET);
  emailSheet.clear();
  emailSheet.getRange(1, 1, 3, 2).setValues([
    ['Email Template Field', 'Value'],
    ['Subject', 'Your generated file: {{Sample}}'],
    ['Body', '<p>Hello,</p><p>Your generated file is attached.</p>'],
  ]);
  emailSheet.autoResizeColumns(1, 2);

  _getLogSheet();
  SpreadsheetApp.getUi().alert('Generator sheets created. Fill in template IDs, folder IDs, replacement values, then select rows or columns and open the sidebar.');
}

function insertQRCode() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const data = cell.getDisplayValue().trim();

  if (!data) {
    SpreadsheetApp.getUi().alert('The selected cell is empty.');
    return;
  }

  const size = 300;
  const urls = [
    `https://quickchart.io/qr?size=${size}&text=${encodeURIComponent(data)}`,
    `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodeURIComponent(data)}`,
  ];

  let lastError = '';
  for (let i = 0; i < urls.length; i++) {
    try {
      const response = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true });
      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        const blob = response.getBlob().setName('qrcode.png');
        sheet.insertImage(blob, cell.getColumn(), cell.getRow());
        return;
      }
      lastError = `HTTP ${response.getResponseCode()}`;
    } catch (error) {
      lastError = error.message;
    }
  }

  SpreadsheetApp.getUi().alert(`Failed to generate QR code: ${lastError}`);
}

function runGenerator(rawOptions, sheetName) {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);

  try {
    _deleteTrigger();
    PropertiesService.getUserProperties().deleteProperty(STATE_KEY);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('This generator must run from a bound Google Sheet.');

    const targetSheetName = sheetName || ss.getActiveSheet().getName();
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error(`Sheet "${targetSheetName}" not found.`);
    if (targetSheetName !== ROW_MODE_SHEET && targetSheetName !== COLUMN_MODE_SHEET) {
      throw new Error(`Use ${ROW_MODE_SHEET} or ${COLUMN_MODE_SHEET}, then select the rows or columns to process.`);
    }

    const options = _normalizeRunOptions(rawOptions || {});
    const selectedItems = targetSheetName === ROW_MODE_SHEET
      ? Array.from(_getSelectedRows(ss, targetSheetName)).sort((a, b) => a - b)
      : Array.from(_getSelectedColumns(ss, targetSheetName)).sort((a, b) => a - b);

    if (selectedItems.length === 0) {
      throw new Error(`No valid selection found in ${targetSheetName}. Select data rows in ${ROW_MODE_SHEET} or data columns in ${COLUMN_MODE_SHEET}.`);
    }

    const sheetData = sheet.getDataRange().getDisplayValues();
    const placeholderMap = targetSheetName === ROW_MODE_SHEET
      ? _buildPlaceholderMap(sheetData[0] || [], 4)
      : _buildColumnPlaceholderMap(sheetData, 4);

    if (Object.keys(placeholderMap).length === 0) {
      throw new Error('No placeholders found. Use headers like {{Name}} starting at column E for R-DOC-GEN or row 5 for C-DOC-GEN.');
    }

    const emailTemplates = _getEmailTemplates(options.sendEmail);
    const relevantData = targetSheetName === ROW_MODE_SHEET
      ? selectedItems.map((rowNum) => sheetData[rowNum - 1] || [])
      : selectedItems.map((colNum) => sheetData.map((row) => row[colNum - 1] || ''));

    const itemTypes = _getItemTypes(selectedItems, sheet, sheetData);
    const baseState = {
      status: 'running',
      startedAt: new Date().toISOString(),
      items: selectedItems,
      currentIndex: 0,
      totalItems: selectedItems.length,
      options,
      sheetName: targetSheetName,
      results: [],
      folders: [],
      placeholderMap,
      emailTemplates,
      relevantData,
      itemTypes,
      isCombinedRun: false,
    };

    _configureCombinedRunIfNeeded(baseState, itemTypes, sheetData);
    PropertiesService.getUserProperties().setProperty(STATE_KEY, JSON.stringify(baseState));

    continueGeneration();
    const state = getGenerationStatus();
    return { status: 'started', totalItems: state.totalItems, currentIndex: state.currentIndex, state };
  } finally {
    lock.releaseLock();
  }
}

function continueGeneration() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30000)) return;

  try {
    const userProperties = PropertiesService.getUserProperties();
    const stateProperty = userProperties.getProperty(STATE_KEY);
    if (!stateProperty) return;

    const state = JSON.parse(stateProperty);
    if (state.status !== 'running') return;

    const started = Date.now();
    let processed = 0;
    const batchSize = _clampNumber(state.options.batchSize, GENERATOR_LIMITS.minBatchSize, GENERATOR_LIMITS.maxBatchSize);
    const usedSheetNames = state.isCombinedRun ? new Set(state.usedSheetNames || []) : null;

    while (state.currentIndex < state.totalItems) {
      if (Date.now() - started > GENERATOR_LIMITS.batchTimeLimitMs) break;
      if (processed >= batchSize) break;

      const itemNum = state.items[state.currentIndex];
      const itemData = state.relevantData[state.currentIndex];

      if (state.isCombinedRun) {
        _processCombinedItem(state, itemNum, itemData, usedSheetNames);
      } else {
        _processSeparateItem(state, itemNum, itemData);
      }

      state.currentIndex += 1;
      processed += 1;
    }

    if (state.currentIndex < state.totalItems) {
      if (state.isCombinedRun) state.usedSheetNames = Array.from(usedSheetNames);
      state.updatedAt = new Date().toISOString();
      userProperties.setProperty(STATE_KEY, JSON.stringify(state));
      _createTrigger();
    } else {
      if (state.isCombinedRun) state.usedSheetNames = Array.from(usedSheetNames);
      _finalizeCombinedRun(state);
      state.status = 'complete';
      state.completedAt = new Date().toISOString();
      _deleteTrigger();
      userProperties.setProperty(STATE_KEY, JSON.stringify(state));
    }
  } catch (error) {
    _deleteTrigger();
    const failedState = getGenerationStatus();
    failedState.status = 'error';
    failedState.error = error.message;
    failedState.updatedAt = new Date().toISOString();
    PropertiesService.getUserProperties().setProperty(STATE_KEY, JSON.stringify(failedState));
    _logActivity({ generator: 'System', status: 'Error', details: error.message });
  } finally {
    lock.releaseLock();
  }
}

function getGenerationStatus() {
  const stateProperty = PropertiesService.getUserProperties().getProperty(STATE_KEY);
  if (!stateProperty) return { status: 'idle', currentIndex: 0, totalItems: 0, results: [], folders: [] };

  const state = JSON.parse(stateProperty);
  state.progress = state.totalItems ? Math.round((state.currentIndex / state.totalItems) * 100) : 0;
  return state;
}

function cancelGeneration() {
  _deleteTrigger();
  PropertiesService.getUserProperties().deleteProperty(STATE_KEY);
  return { status: 'cancelled', currentIndex: 0, totalItems: 0, results: [], folders: [] };
}

function getSelectionDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { sheetName: '', mimeTypes: [], selectedIndices: [], message: 'No active spreadsheet found.' };

  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  if (sheetName !== ROW_MODE_SHEET && sheetName !== COLUMN_MODE_SHEET) {
    return {
      sheetName,
      mimeTypes: [],
      selectedIndices: [],
      message: `Switch to ${ROW_MODE_SHEET} or ${COLUMN_MODE_SHEET}.`,
    };
  }

  const selectedIndices = sheetName === ROW_MODE_SHEET
    ? Array.from(_getSelectedRows(ss, sheetName)).sort((a, b) => a - b)
    : Array.from(_getSelectedColumns(ss, sheetName)).sort((a, b) => a - b);

  const data = activeSheet.getDataRange().getDisplayValues();
  const mimeTypes = new Set();
  const invalidItems = [];

  selectedIndices.forEach((itemNum) => {
    const templateId = sheetName === ROW_MODE_SHEET
      ? _safeTrim(data[itemNum - 1] && data[itemNum - 1][0])
      : _safeTrim(data[0] && data[0][itemNum - 1]);

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

  return {
    sheetName,
    mimeTypes: Array.from(mimeTypes),
    selectedIndices,
    invalidItems,
    message: selectedIndices.length ? '' : 'Select one or more valid rows or columns to process.',
  };
}

function _normalizeRunOptions(options) {
  return {
    format: options.format === 'PDF' ? 'PDF' : 'DOC',
    sendEmail: Boolean(options.sendEmail),
    batchSize: _clampNumber(options.batchSize || 5, GENERATOR_LIMITS.minBatchSize, GENERATOR_LIMITS.maxBatchSize),
    sheetTarget: _safeTrim(options.sheetTarget),
    newSheetName: _sanitizeSheetName(_safeTrim(options.newSheetName)),
    sheetOutputMode: options.sheetOutputMode === 'combined' ? 'combined' : 'separate',
    docSlideOutputMode: options.docSlideOutputMode === 'combined' ? 'combined' : 'separate',
    combinedSheetName: _sanitizeFileName(options.combinedSheetName || 'Combined Spreadsheet'),
    combinedDocSlideName: _sanitizeFileName(options.combinedDocSlideName || 'Combined Output'),
  };
}

function _configureCombinedRunIfNeeded(state, itemTypes, sheetData) {
  const hasSheets = itemTypes.sheets.length > 0;
  const hasDocs = itemTypes.docs.length > 0;
  const hasSlides = itemTypes.slides.length > 0;

  if (state.options.sheetOutputMode === 'combined' && hasSheets) {
    _initializeCombinedRun('sheets', itemTypes.sheets, state, sheetData);
    return;
  }

  if (state.options.docSlideOutputMode === 'combined' && (hasDocs || hasSlides)) {
    if (hasDocs && hasSlides) {
      throw new Error('Cannot combine Docs and Slides in one master file. Select only Docs or only Slides for combined mode.');
    }
    _initializeCombinedRun(hasDocs ? 'docs' : 'slides', hasDocs ? itemTypes.docs : itemTypes.slides, state, sheetData);
  }
}

function _initializeCombinedRun(type, itemsToCombine, state, sheetData) {
  if (!itemsToCombine.length) return;

  const firstItemNum = itemsToCombine[0];
  const firstData = state.sheetName === ROW_MODE_SHEET
    ? sheetData[firstItemNum - 1]
    : sheetData.map((row) => row[firstItemNum - 1]);

  const firstTemplateId = _safeTrim(firstData && firstData[0]);
  const folderId = _safeTrim(firstData && firstData[1]);
  if (!firstTemplateId) throw new Error('The first selected item for combining is missing a Template ID.');
  if (!folderId) throw new Error('The first selected item for combining is missing a Folder ID.');

  const destinationFolder = DriveApp.getFolderById(folderId);
  let combinedFile;

  if (type === 'sheets') {
    const spreadsheet = SpreadsheetApp.create(state.options.combinedSheetName || 'Combined Spreadsheet');
    combinedFile = DriveApp.getFileById(spreadsheet.getId());
    combinedFile.moveTo(destinationFolder);
  } else {
    const templateFile = DriveApp.getFileById(firstTemplateId);
    combinedFile = templateFile.makeCopy(state.options.combinedDocSlideName || 'Combined Output', destinationFolder);
    if (type === 'docs') _prepareCombinedDocument(combinedFile.getId());
    if (type === 'slides') _prepareCombinedPresentation(combinedFile.getId());
  }

  const itemSet = new Set(itemsToCombine);
  state.items = itemsToCombine;
  state.totalItems = itemsToCombine.length;
  state.relevantData = state.sheetName === ROW_MODE_SHEET
    ? state.items.map((rowNum) => sheetData[rowNum - 1] || [])
    : state.items.map((colNum) => sheetData.map((row) => row[colNum - 1] || ''));
  state.folders = [{ id: folderId, url: destinationFolder.getUrl() }];
  state.isCombinedRun = true;
  state.combinedRunType = type;
  state.combinedFileId = combinedFile.getId();
  state.usedSheetNames = [];
  state.skippedItems = state.items.filter((item) => !itemSet.has(item));
}

function _processSeparateItem(state, itemNum, itemData) {
  const itemLabel = state.sheetName === ROW_MODE_SHEET ? `Row ${itemNum}` : `Column ${_columnNumberToLetter(itemNum)}`;
  const templateId = _safeTrim(itemData && itemData[0]);

  if (!templateId) {
    state.results.push({ item: itemLabel, status: '❌ Missing Template ID' });
    return;
  }

  const result = _processSingleItem(itemLabel, templateId, itemData, state.placeholderMap, state.options, state.emailTemplates.subject, state.emailTemplates.body);
  state.results.push(result.result);

  if (result.folderId) {
    const exists = state.folders.some((folder) => folder.id === result.folderId);
    if (!exists) state.folders.push({ id: result.folderId, url: `https://drive.google.com/drive/folders/${result.folderId}` });
  }
}

function _processSingleItem(itemLabel, templateId, data, placeholderMap, options, subjectTemplate, bodyTemplate) {
  const folderId = _safeTrim(data && data[1]);
  const recipientEmail = _safeTrim(data && data[2]);
  const outputName = _sanitizeFileName(data && data[3] ? data[3] : itemLabel);
  const logData = { fileName: outputName, recipient: recipientEmail || 'N/A' };

  if (!folderId) {
    _logActivity({ ...logData, generator: 'Template', status: 'Error', details: 'Missing Folder ID' });
    return { result: { item: itemLabel, status: '❌ Missing Folder ID' }, folderId: null };
  }

  try {
    const templateFile = DriveApp.getFileById(templateId);
    const mimeType = templateFile.getMimeType();
    const destinationFolder = DriveApp.getFolderById(folderId);
    const newFileName = `${templateFile.getName()} - ${outputName}`;
    const newFile = templateFile.makeCopy(newFileName, destinationFolder);
    let finalFile = newFile;
    let status = '✅ Success';
    let generator = 'Template';

    if (mimeType === APP_MIME.GOOGLE_DOCS || mimeType === APP_MIME.GOOGLE_SLIDES) {
      generator = mimeType === APP_MIME.GOOGLE_DOCS ? 'Doc' : 'Slide';
      _generateDocOrSlide(newFile, mimeType, placeholderMap, data);
      if (options.format === 'PDF') {
        finalFile = _convertFileToPdf(newFile, destinationFolder);
        newFile.setTrashed(true);
        status = '✅ Success (PDF)';
      } else {
        status = '✅ Success (Original format)';
      }
    } else if (mimeType === APP_MIME.GOOGLE_SHEETS) {
      generator = 'Sheet';
      _generateSheet(SpreadsheetApp.openById(newFile.getId()), templateId, placeholderMap, data, options.sheetTarget, options.newSheetName, false, null);
      status = '✅ Success (Sheet)';
    } else {
      throw new Error(`Unsupported template type: ${mimeType}`);
    }

    let emailStatus = 'Not sent';
    if (recipientEmail && options.sendEmail) {
      const emailResult = _sendEmailNotification(finalFile, recipientEmail, subjectTemplate, bodyTemplate, data, placeholderMap);
      status = emailResult.status;
      emailStatus = emailResult.logStatus;
    }

    _logActivity({
      generator,
      fileName: finalFile.getName(),
      fileUrl: finalFile.getUrl(),
      folderUrl: destinationFolder.getUrl(),
      recipient: recipientEmail || 'N/A',
      emailStatus,
      status: status.startsWith('❌') ? 'Error' : 'Success',
      details: status.replace(/^[✅⚠️❌]\s*/, ''),
    });

    return { result: { item: itemLabel, status, url: finalFile.getUrl() }, folderId };
  } catch (error) {
    _logActivity({ ...logData, generator: 'Template', status: 'Error', details: error.message });
    return { result: { item: itemLabel, status: `❌ ${error.message}` }, folderId };
  }
}

function _processCombinedItem(state, itemNum, itemData, usedSheetNames) {
  const itemLabel = state.sheetName === ROW_MODE_SHEET ? `Row ${itemNum}` : `Column ${_columnNumberToLetter(itemNum)}`;
  const templateId = _safeTrim(itemData && itemData[0]);
  let status = '✅ Success (Combined)';

  try {
    if (!templateId) throw new Error('Missing Template ID');

    if (state.combinedRunType === 'sheets') {
      const combinedSpreadsheet = SpreadsheetApp.openById(state.combinedFileId);
      const sheetNameBase = _sanitizeSheetName(itemData && itemData[3] ? itemData[3] : `${itemLabel} Sheet`);
      _generateSheet(combinedSpreadsheet, templateId, state.placeholderMap, itemData, state.options.sheetTarget, sheetNameBase, true, usedSheetNames);
    } else {
      const destinationFolder = DriveApp.getFolderById(state.folders[0].id);
      const templateFile = DriveApp.getFileById(templateId);
      const tempFile = templateFile.makeCopy(`temp_${itemLabel}_${Date.now()}`, destinationFolder);

      try {
        const mimeType = templateFile.getMimeType();
        _generateDocOrSlide(tempFile, mimeType, state.placeholderMap, itemData);

        if (state.combinedRunType === 'docs') {
          const masterDoc = DocumentApp.openById(state.combinedFileId);
          const sourceDoc = DocumentApp.openById(tempFile.getId());
          _appendDocContent(masterDoc, sourceDoc, state.currentIndex > 0);
          masterDoc.saveAndClose();
        } else if (state.combinedRunType === 'slides') {
          const masterPresentation = SlidesApp.openById(state.combinedFileId);
          const sourcePresentation = SlidesApp.openById(tempFile.getId());
          _appendSlides(masterPresentation, sourcePresentation);
          masterPresentation.saveAndClose();
        }
      } finally {
        tempFile.setTrashed(true);
      }
    }
  } catch (error) {
    status = `❌ Error: ${error.message}`;
  }

  const combinedFile = DriveApp.getFileById(state.combinedFileId);
  state.results.push({ item: itemLabel, status, url: combinedFile.getUrl() });
}

function _finalizeCombinedRun(state) {
  if (!state.isCombinedRun || !state.combinedFileId) return;

  const combinedFile = DriveApp.getFileById(state.combinedFileId);
  let finalFile = combinedFile;

  if (state.combinedRunType === 'sheets') {
    const spreadsheet = SpreadsheetApp.openById(state.combinedFileId);
    const defaultSheet = spreadsheet.getSheetByName('Sheet1');
    if (defaultSheet && spreadsheet.getSheets().length > 1) spreadsheet.deleteSheet(defaultSheet);
  }

  if (state.combinedRunType === 'slides') {
    _removeEmptyFirstSlide(state.combinedFileId);
  }

  if ((state.combinedRunType === 'docs' || state.combinedRunType === 'slides') && state.options.format === 'PDF') {
    const folder = DriveApp.getFolderById(state.folders[0].id);
    finalFile = _convertFileToPdf(combinedFile, folder);
    combinedFile.setTrashed(true);
  }

  _logActivity({
    generator: `Combined ${state.combinedRunType}`,
    fileName: finalFile.getName(),
    fileUrl: finalFile.getUrl(),
    folderUrl: state.folders.length ? state.folders[0].url : '',
    status: 'Success',
    details: `Combined ${state.results.filter((result) => result.status.startsWith('✅')).length} of ${state.totalItems} items.`,
  });

  state.combinedFinalUrl = finalFile.getUrl();
}

function _generateDocOrSlide(file, mimeType, placeholderMap, rowData) {
  if (mimeType === APP_MIME.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(file.getId());
    _replacePlaceholdersInDocument(doc, placeholderMap, rowData);
    doc.saveAndClose();
    return;
  }

  if (mimeType === APP_MIME.GOOGLE_SLIDES) {
    const presentation = SlidesApp.openById(file.getId());
    _replaceInPresentation(presentation, placeholderMap, rowData);
    presentation.saveAndClose();
    return;
  }

  throw new Error(`Unsupported document or slide type: ${mimeType}`);
}

function _generateSheet(targetSpreadsheet, templateId, placeholderMap, rowData, sheetTarget, newSheetName, isCombined, usedSheetNames) {
  const replacementMap = _buildReplacementMap(placeholderMap, rowData);

  if (isCombined) {
    const templateSpreadsheet = SpreadsheetApp.openById(templateId);
    const templateSheets = sheetTarget ? [templateSpreadsheet.getSheetByName(sheetTarget)] : templateSpreadsheet.getSheets();
    if (templateSheets.some((sheet) => !sheet)) throw new Error(`Sheet "${sheetTarget}" not found in template.`);

    templateSheets.forEach((templateSheet) => {
      const copiedSheet = templateSheet.copyTo(targetSpreadsheet);
      const preferredName = newSheetName || templateSheet.getName();
      const uniqueName = _getUniqueSheetName(preferredName, usedSheetNames);
      usedSheetNames.add(uniqueName);
      copiedSheet.setName(uniqueName);
      replacePlaceholdersInSheet(copiedSheet, replacementMap);
    });
    return;
  }

  const sheetsToProcess = [];
  if (sheetTarget) {
    const targetSheet = targetSpreadsheet.getSheetByName(sheetTarget);
    if (!targetSheet) throw new Error(`Sheet "${sheetTarget}" not found in copied spreadsheet.`);
    targetSpreadsheet.getSheets().forEach((sheet) => {
      if (sheet.getSheetId() !== targetSheet.getSheetId()) targetSpreadsheet.deleteSheet(sheet);
    });
    if (newSheetName) targetSheet.setName(_getUniqueSheetNameInSpreadsheet(targetSpreadsheet, newSheetName, targetSheet));
    sheetsToProcess.push(targetSheet);
  } else {
    targetSpreadsheet.getSheets().forEach((sheet) => sheetsToProcess.push(sheet));
    if (newSheetName && sheetsToProcess.length === 1) {
      sheetsToProcess[0].setName(_getUniqueSheetNameInSpreadsheet(targetSpreadsheet, newSheetName, sheetsToProcess[0]));
    }
  }

  sheetsToProcess.forEach((sheet) => replacePlaceholdersInSheet(sheet, replacementMap));
  SpreadsheetApp.flush();
}

function replacePlaceholdersInSheet(sheet, replacementMap) {
  replacementMap.forEach((value, placeholder) => {
    sheet.createTextFinder(placeholder).matchCase(false).replaceAllWith(value == null ? '' : String(value));
  });
}

function _replacePlaceholdersInDocument(doc, placeholderMap, rowData) {
  [doc.getHeader(), doc.getBody(), doc.getFooter()].forEach((section) => {
    if (!section) return;
    Object.keys(placeholderMap).forEach((base) => {
      _replaceInSection(section, base, rowData[placeholderMap[base]]);
    });
  });
}

function _replaceInSection(section, base, rawValue) {
  const regex = _placeholderRegexString(base);
  let found = section.findText(regex);

  while (found) {
    const textElement = found.getElement().asText();
    const start = found.getStartOffset();
    const end = found.getEndOffsetInclusive();
    const matchedText = textElement.getText().substring(start, end + 1);
    const modifier = _extractModifier(base, matchedText);
    const existingUrl = _safeGetTextLink(textElement, start);
    const imageBlob = _getImageBlob(rawValue);

    textElement.deleteText(start, end);

    if (imageBlob) {
      _insertImageAfterTextElement(textElement, imageBlob);
    } else {
      const replacementText = _applyModifiers(rawValue, modifier);
      if (replacementText) {
        textElement.insertText(start, replacementText);
        const replacementEnd = start + replacementText.length - 1;
        if (modifier.includes('b')) textElement.setBold(start, replacementEnd, true);
        if (existingUrl) textElement.setLinkUrl(start, replacementEnd, existingUrl);
        else if (/^https?:\/\//i.test(replacementText)) textElement.setLinkUrl(start, replacementEnd, replacementText);
      }
    }

    found = section.findText(regex, found);
  }
}

function _replaceInPresentation(presentation, placeholderMap, rowData) {
  presentation.getSlides().forEach((slide) => {
    Object.keys(placeholderMap).forEach((base) => {
      const rawValue = rowData[placeholderMap[base]];
      if (!_replacePlaceholderAsImageInSlide(slide, base, rawValue)) {
        _replacePlaceholderAsTextInSlide(slide, base, rawValue);
      }
    });
  });
}

function _replacePlaceholderAsImageInSlide(slide, base, rawValue) {
  const blob = _getImageBlob(rawValue);
  if (!blob) return false;

  const canonicalPlaceholder = `{{${base}}}`;
  const shapes = slide.getShapes();

  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    try {
      if (!shape.getText) continue;
      if (shape.getText().asString().trim() !== canonicalPlaceholder) continue;

      const image = slide.insertImage(blob);
      image.setLeft(shape.getLeft()).setTop(shape.getTop()).setWidth(shape.getWidth()).setHeight(shape.getHeight());
      shape.remove();
      return true;
    } catch (error) {
      // Keep searching other shapes.
    }
  }

  return false;
}

function _replacePlaceholderAsTextInSlide(slide, base, rawValue) {
  const regex = _placeholderRegex(base);
  const elements = _getSlideTextElements(slide);

  elements.forEach((element) => {
    try {
      const textRange = element.getText();
      const originalText = textRange.asString();
      const matches = [];
      let match;
      while ((match = regex.exec(originalText)) !== null) matches.push(match);

      for (let i = matches.length - 1; i >= 0; i--) {
        const currentMatch = matches[i];
        const placeholder = currentMatch[0];
        const modifier = currentMatch[1] || '';
        const replacementText = _applyModifiers(rawValue, modifier);
        const start = currentMatch.index;
        const end = start + placeholder.length;
        const placeholderRange = textRange.getRange(start, end);
        let existingUrl = null;

        try {
          const link = placeholderRange.getTextStyle().getLink();
          if (link) existingUrl = link.getUrl();
        } catch (error) {}

        placeholderRange.setText(replacementText);

        if (replacementText) {
          const replacementRange = textRange.getRange(start, start + replacementText.length);
          if (modifier.includes('b')) replacementRange.getTextStyle().setBold(true);
          if (existingUrl) replacementRange.getTextStyle().setLinkUrl(existingUrl);
          else if (/^https?:\/\//i.test(replacementText)) replacementRange.getTextStyle().setLinkUrl(replacementText);
        }
      }
    } catch (error) {
      // Skip unsupported slide elements.
    }
  });
}

function _getSlideTextElements(slide) {
  const elements = [];
  slide.getShapes().forEach((shape) => elements.push(shape));
  slide.getTables().forEach((table) => {
    for (let row = 0; row < table.getNumRows(); row++) {
      for (let col = 0; col < table.getNumColumns(); col++) {
        elements.push(table.getCell(row, col));
      }
    }
  });
  return elements;
}

function _sendEmailNotification(file, recipient, subjectTemplate, bodyTemplate, rowData, placeholderMap) {
  try {
    file.addViewer(recipient);
    const subject = _replacePlaceholdersInEmail(subjectTemplate || `Generated file: ${file.getName()}`, rowData, placeholderMap);
    const htmlBody = _replacePlaceholdersInEmail(bodyTemplate || '<p>Your generated file is attached.</p>', rowData, placeholderMap);
    MailApp.sendEmail({ to: recipient, subject, htmlBody, attachments: [file.getBlob()] });
    return { status: `✅ Emailed to ${recipient}`, logStatus: 'Sent' };
  } catch (error) {
    return { status: `⚠️ Success, but email failed: ${error.message}`, logStatus: `Failed: ${error.message}` };
  }
}

function _replacePlaceholdersInEmail(template, rowData, placeholderMap) {
  let result = template || '';
  Object.keys(placeholderMap).forEach((base) => {
    result = result.replace(_placeholderRegex(base), (match, modifier) => {
      let value = _escapeHtml(_applyModifiers(rowData[placeholderMap[base]], modifier || ''));
      if ((modifier || '').includes('b')) value = `<b>${value}</b>`;
      return value;
    });
  });
  return result;
}

function _applyModifiers(rawValue, modifier) {
  if (rawValue === null || typeof rawValue === 'undefined') return '';
  let value = String(rawValue);

  if (!/[a-zA-Z]/.test(value)) {
    const maybeDate = new Date(value);
    if (!isNaN(maybeDate.getTime()) && /^\d{1,4}[\-/]\d{1,2}[\-/]\d{1,4}|\d{4}-\d{2}-\d{2}T/.test(value)) {
      value = Utilities.formatDate(maybeDate, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
    }
  }

  if (modifier === 'u' || modifier === 'ub') value = value.toUpperCase();
  if (modifier === '1' || modifier === '1b') value = value.toLowerCase().replace(/\b\w/g, (char) => char.toUpperCase());
  return value;
}

function _buildReplacementMap(placeholderMap, rowData) {
  const replacements = new Map();
  Object.keys(placeholderMap).forEach((base) => {
    const value = rowData[placeholderMap[base]];
    ['', 'u', 'ub', 'b', '1', '1b'].forEach((modifier) => {
      const placeholder = modifier ? `{{${base}}${modifier}}`.replace('}}', '}') + '}' : `{{${base}}}`;
      replacements.set(placeholder, _applyModifiers(value, modifier));
    });
  });
  return replacements;
}

function _getEmailTemplates(sendEmail) {
  if (!sendEmail) return { subject: '', body: '' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emailSheet = ss && ss.getSheetByName(EMAIL_SHEET);
    if (!emailSheet) return { subject: '', body: '' };

    const values = emailSheet.getDataRange().getValues();
    const subject = values[1] ? values[1][1] || values[1][0] || '' : '';
    const body = values[2] ? values[2][1] || values[2][0] || '' : '';
    return { subject, body };
  } catch (error) {
    return { subject: '', body: '' };
  }
}

function _getSelectedRows(ss, sheetName) {
  const selectedRows = new Set();
  const rangeList = ss.getActiveRangeList();
  if (!rangeList) return selectedRows;

  rangeList.getRanges().forEach((range) => {
    if (range.getSheet().getName() !== sheetName) return;
    for (let i = 0; i < range.getNumRows(); i++) {
      const row = range.getRow() + i;
      if (row >= 3) selectedRows.add(row);
    }
  });

  return selectedRows;
}

function _getSelectedColumns(ss, sheetName) {
  const selectedColumns = new Set();
  const rangeList = ss.getActiveRangeList();
  if (!rangeList) return selectedColumns;

  rangeList.getRanges().forEach((range) => {
    if (range.getSheet().getName() !== sheetName) return;
    for (let i = 0; i < range.getNumColumns(); i++) {
      const column = range.getColumn() + i;
      if (column >= 2) selectedColumns.add(column);
    }
  });

  return selectedColumns;
}

function _buildPlaceholderMap(headerRow, startIndex) {
  const placeholderMap = {};
  headerRow.forEach((header, columnIndex) => {
    const match = _safeTrim(header).match(/^\{\{(.+?)\}\}$/);
    if (match && columnIndex >= startIndex) placeholderMap[match[1]] = columnIndex;
  });
  return placeholderMap;
}

function _buildColumnPlaceholderMap(data, startIndex) {
  const placeholderMap = {};
  data.forEach((row, rowIndex) => {
    const match = _safeTrim(row && row[0]).match(/^\{\{(.+?)\}\}$/);
    if (match && rowIndex >= startIndex) placeholderMap[match[1]] = rowIndex;
  });
  return placeholderMap;
}

function _getItemTypes(selectedItems, sheet, data) {
  const types = { docs: [], slides: [], sheets: [], unsupported: [] };

  selectedItems.forEach((itemNum) => {
    const templateId = sheet.getName() === ROW_MODE_SHEET
      ? _safeTrim(data[itemNum - 1] && data[itemNum - 1][0])
      : _safeTrim(data[0] && data[0][itemNum - 1]);

    try {
      if (!templateId) return;
      const mimeType = DriveApp.getFileById(templateId).getMimeType();
      if (mimeType === APP_MIME.GOOGLE_DOCS) types.docs.push(itemNum);
      else if (mimeType === APP_MIME.GOOGLE_SLIDES) types.slides.push(itemNum);
      else if (mimeType === APP_MIME.GOOGLE_SHEETS) types.sheets.push(itemNum);
      else types.unsupported.push(itemNum);
    } catch (error) {
      types.unsupported.push(itemNum);
    }
  });

  return types;
}

function _createTrigger() {
  _deleteTrigger();
  ScriptApp.newTrigger(TRIGGER_HANDLER).timeBased().after(GENERATOR_LIMITS.triggerDelayMs).create();
}

function _deleteTrigger() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === TRIGGER_HANDLER) ScriptApp.deleteTrigger(trigger);
  });
}

function _getLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;

  const logSheet = _getOrCreateSheet(ss, LOG_SHEET);
  if (logSheet.getLastRow() === 0) {
    const headers = ['Timestamp', 'Generator', 'File Name', 'File URL', 'Folder URL', 'Recipient', 'Email Status', 'Status', 'Details'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }
  return logSheet;
}

function _logActivity(logData) {
  try {
    const logSheet = _getLogSheet();
    if (!logSheet) return;

    const entry = {
      generator: 'N/A',
      fileName: 'N/A',
      fileUrl: 'N/A',
      folderUrl: 'N/A',
      recipient: 'N/A',
      emailStatus: 'N/A',
      status: 'Unknown',
      details: '',
      ...logData,
    };

    logSheet.appendRow([
      new Date(),
      entry.generator,
      entry.fileName,
      entry.fileUrl,
      entry.folderUrl,
      entry.recipient,
      entry.emailStatus,
      entry.status,
      entry.details,
    ]);
  } catch (error) {
    // Logging must never break generation.
  }
}

function _prepareCombinedDocument(fileId) {
  const doc = DocumentApp.openById(fileId);
  doc.getBody().clear();
  doc.saveAndClose();
}

function _prepareCombinedPresentation(fileId) {
  const presentation = SlidesApp.openById(fileId);
  const slides = presentation.getSlides();
  for (let i = slides.length - 1; i > 0; i--) slides[i].remove();
  presentation.getSlides()[0].getPageElements().forEach((element) => element.remove());
  presentation.saveAndClose();
}

function _appendDocContent(masterDoc, sourceDoc, addPageBreak) {
  const masterBody = masterDoc.getBody();
  const sourceBody = sourceDoc.getBody();

  if (addPageBreak && masterBody.getNumChildren() > 0) masterBody.appendPageBreak();

  for (let i = 0; i < sourceBody.getNumChildren(); i++) {
    const child = sourceBody.getChild(i).copy();
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) masterBody.appendParagraph(child);
    else if (type === DocumentApp.ElementType.TABLE) masterBody.appendTable(child);
    else if (type === DocumentApp.ElementType.LIST_ITEM) masterBody.appendListItem(child);
    else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) masterBody.appendHorizontalRule();
    else masterBody.appendParagraph(child.asText ? child.asText().getText() : '');
  }
}

function _appendSlides(masterPresentation, sourcePresentation) {
  sourcePresentation.getSlides().forEach((slide) => {
    masterPresentation.insertSlide(masterPresentation.getSlides().length, slide);
  });
}

function _removeEmptyFirstSlide(fileId) {
  const presentation = SlidesApp.openById(fileId);
  const slides = presentation.getSlides();
  if (slides.length > 1 && slides[0].getPageElements().length === 0) slides[0].remove();
  presentation.saveAndClose();
}

function _convertFileToPdf(file, folder) {
  const pdfBlob = file.getBlob().getAs(MimeType.PDF).setName(`${file.getName()}.pdf`);
  return folder.createFile(pdfBlob);
}

function _getImageBlob(rawValue) {
  const value = _safeTrim(rawValue);
  if (!value) return null;

  if (/^https?:\/\//i.test(value)) {
    try {
      const response = UrlFetchApp.fetch(value, { muteHttpExceptions: true });
      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        const blob = response.getBlob();
        return blob.getContentType().startsWith('image/') ? blob : null;
      }
    } catch (error) {
      return null;
    }
  }

  if (/^[a-zA-Z0-9_-]{25,}$/.test(value)) {
    try {
      const file = DriveApp.getFileById(value);
      return file.getMimeType().startsWith('image/') ? file.getBlob() : null;
    } catch (error) {
      return null;
    }
  }

  return null;
}

function _insertImageAfterTextElement(textElement, blob) {
  try {
    const parent = textElement.getParent();
    if (parent && parent.insertInlineImage && parent.getChildIndex) {
      parent.insertInlineImage(parent.getChildIndex(textElement) + 1, blob);
    }
  } catch (error) {
    // If image insertion fails, leave the placeholder removed instead of crashing generation.
  }
}

function _safeGetTextLink(textElement, offset) {
  try {
    return textElement.getLinkUrl(offset);
  } catch (error) {
    return null;
  }
}

function _placeholderRegex(base) {
  return new RegExp(_placeholderRegexString(base), 'g');
}

function _placeholderRegexString(base) {
  return `\\{\\{${_escapeRegExp(base)}(?:\\}(u|ub|b|1|1b))?\\}\\}`;
}

function _extractModifier(base, matchedText) {
  const match = matchedText.match(new RegExp(`^${_placeholderRegexString(base)}$`));
  return match && match[1] ? match[1] : '';
}

function _getUniqueSheetName(baseName, usedNamesSet) {
  let safeBase = _sanitizeSheetName(baseName || 'Sheet');
  if (!usedNamesSet.has(safeBase)) return safeBase;

  let index = 2;
  while (usedNamesSet.has(`${safeBase} ${index}`)) index += 1;
  return `${safeBase} ${index}`;
}

function _getUniqueSheetNameInSpreadsheet(spreadsheet, preferredName, currentSheet) {
  const baseName = _sanitizeSheetName(preferredName || 'Sheet');
  const existing = spreadsheet.getSheets()
    .filter((sheet) => !currentSheet || sheet.getSheetId() !== currentSheet.getSheetId())
    .map((sheet) => sheet.getName());

  if (existing.indexOf(baseName) === -1) return baseName;
  let index = 2;
  while (existing.indexOf(`${baseName} ${index}`) !== -1) index += 1;
  return `${baseName} ${index}`;
}

function _getOrCreateSheet(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function _columnNumberToLetter(columnNumber) {
  let number = columnNumber;
  let letter = '';
  while (number > 0) {
    const remainder = (number - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    number = Math.floor((number - 1) / 26);
  }
  return letter;
}

function _safeTrim(value) {
  return value === null || typeof value === 'undefined' ? '' : String(value).trim();
}

function _sanitizeFileName(value) {
  const cleaned = _safeTrim(value)
    .replace(/[\u0000-\u001f]/g, '')
    .replace(/[\\/]/g, '-')
    .replace(/\s+/g, ' ')
    .slice(0, GENERATOR_LIMITS.maxNameLength);
  return cleaned || 'Generated File';
}

function _sanitizeSheetName(value) {
  const cleaned = _safeTrim(value)
    .replace(/[\\/?*\[\]:]/g, '-')
    .slice(0, 90);
  return cleaned || 'Sheet';
}

function _clampNumber(value, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return min;
  return Math.max(min, Math.min(max, Math.floor(number)));
}

function _escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function _escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// =================================================================================================
// Standalone web app workspace generator used by Index.html
// =================================================================================================

function getDefaultWorkspaceConfig() {
  return {
    workspaceName: 'Generated Workspace',
    description: 'Created with Google Workspace Generator.',
    documents: [
      { name: 'Project Brief', body: 'Purpose\n\nScope\n\nDeliverables\n\nTimeline' },
    ],
    spreadsheets: [
      {
        name: 'Tracker',
        sheets: [
          {
            name: 'Tasks',
            values: [
              ['Task', 'Owner', 'Status', 'Due Date'],
              ['Example task', '', 'Not started', ''],
            ],
          },
        ],
      },
    ],
    presentations: [
      { name: 'Overview Deck', title: 'Generated Workspace', subtitle: 'Created with Google Workspace Generator' },
    ],
  };
}

function generateWorkspace(rawConfig) {
  const config = _normalizeWorkspaceConfig(rawConfig);
  const errors = _validateWorkspaceConfig(config);
  if (errors.length) throw new Error(errors.join(' '));

  const folder = DriveApp.createFolder(config.workspaceName);
  if (config.description) folder.setDescription(config.description);

  const resources = [];
  config.documents.forEach((docConfig) => resources.push(_createSimpleDocument(docConfig, folder)));
  config.spreadsheets.forEach((sheetConfig) => resources.push(_createSimpleSpreadsheet(sheetConfig, folder)));
  config.presentations.forEach((slideConfig) => resources.push(_createSimplePresentation(slideConfig, folder)));

  return {
    appName: APP_NAME,
    createdAt: new Date().toISOString(),
    workspace: _folderToResource(folder),
    resources,
    summary: {
      totalFiles: resources.length,
      documents: config.documents.length,
      spreadsheets: config.spreadsheets.length,
      presentations: config.presentations.length,
    },
  };
}

function _normalizeWorkspaceConfig(rawConfig) {
  let input = rawConfig || {};
  if (typeof input === 'string') input = JSON.parse(input);
  const defaults = getDefaultWorkspaceConfig();

  return {
    workspaceName: _sanitizeFileName(input.workspaceName || defaults.workspaceName),
    description: _safeTrim(input.description).slice(0, GENERATOR_LIMITS.maxDescriptionLength),
    documents: _asArray(input.documents).map((item, index) => ({
      name: _sanitizeFileName(typeof item === 'string' ? item : item && item.name || `Document ${index + 1}`),
      body: _safeTrim(item && item.body).slice(0, GENERATOR_LIMITS.maxDocumentBodyLength),
    })),
    spreadsheets: _asArray(input.spreadsheets).map((item, index) => ({
      name: _sanitizeFileName(typeof item === 'string' ? item : item && item.name || `Spreadsheet ${index + 1}`),
      sheets: _asArray(item && item.sheets).length ? _asArray(item.sheets).map((sheet, sheetIndex) => ({
        name: _sanitizeSheetName(sheet && sheet.name || `Sheet ${sheetIndex + 1}`),
        values: _normalizeSimpleSheetValues(sheet && sheet.values),
      })) : [{ name: 'Sheet1', values: [] }],
    })),
    presentations: _asArray(input.presentations).map((item, index) => ({
      name: _sanitizeFileName(typeof item === 'string' ? item : item && item.name || `Presentation ${index + 1}`),
      title: _safeTrim(item && item.title || item && item.name || `Presentation ${index + 1}`).slice(0, GENERATOR_LIMITS.maxNameLength),
      subtitle: _safeTrim(item && item.subtitle).slice(0, GENERATOR_LIMITS.maxDescriptionLength),
    })),
  };
}

function _validateWorkspaceConfig(config) {
  const errors = [];
  const total = config.documents.length + config.spreadsheets.length + config.presentations.length;
  if (!config.workspaceName) errors.push('Workspace name is required.');
  if (total < 1) errors.push('Generate at least one file.');
  if (config.documents.length > GENERATOR_LIMITS.maxSimpleDocs) errors.push(`Maximum documents allowed: ${GENERATOR_LIMITS.maxSimpleDocs}.`);
  if (config.spreadsheets.length > GENERATOR_LIMITS.maxSimpleSheets) errors.push(`Maximum spreadsheets allowed: ${GENERATOR_LIMITS.maxSimpleSheets}.`);
  if (config.presentations.length > GENERATOR_LIMITS.maxSimpleSlides) errors.push(`Maximum presentations allowed: ${GENERATOR_LIMITS.maxSimpleSlides}.`);
  if (total > GENERATOR_LIMITS.maxSimpleFiles) errors.push(`Maximum total files allowed: ${GENERATOR_LIMITS.maxSimpleFiles}.`);
  return errors;
}

function _createSimpleDocument(config, folder) {
  const doc = DocumentApp.create(config.name);
  const body = doc.getBody();
  body.clear();
  body.appendParagraph(config.name).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  (config.body || 'Generated by Google Workspace Generator.').split('\n').forEach((line) => body.appendParagraph(line));
  doc.saveAndClose();
  return _fileToResource(_moveFileToFolder(doc.getId(), folder), 'document');
}

function _createSimpleSpreadsheet(config, folder) {
  const spreadsheet = SpreadsheetApp.create(config.name);
  const existingSheets = spreadsheet.getSheets();

  config.sheets.forEach((sheetConfig, index) => {
    const sheet = index === 0 ? existingSheets[0] : spreadsheet.insertSheet();
    sheet.setName(_getUniqueSheetNameInSpreadsheet(spreadsheet, sheetConfig.name, sheet));
    if (sheetConfig.values.length) {
      const width = Math.max(1, Math.max.apply(null, sheetConfig.values.map((row) => row.length)));
      const values = sheetConfig.values.map((row) => {
        const normalized = row.slice(0, width);
        while (normalized.length < width) normalized.push('');
        return normalized;
      });
      sheet.getRange(1, 1, values.length, width).setValues(values);
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, width);
    } else {
      sheet.getRange(1, 1).setValue('Generated by Google Workspace Generator');
    }
  });

  SpreadsheetApp.flush();
  return _fileToResource(_moveFileToFolder(spreadsheet.getId(), folder), 'spreadsheet');
}

function _createSimplePresentation(config, folder) {
  const presentation = SlidesApp.create(config.name);
  const slide = presentation.getSlides()[0];
  slide.getPageElements().forEach((element) => element.remove());
  slide.insertTextBox(config.title || config.name, 60, 80, 600, 70).getText().getTextStyle().setFontSize(32).setBold(true);
  slide.insertTextBox(config.subtitle || 'Generated by Google Workspace Generator', 60, 170, 600, 80).getText().getTextStyle().setFontSize(16);
  presentation.saveAndClose();
  return _fileToResource(_moveFileToFolder(presentation.getId(), folder), 'presentation');
}

function _normalizeSimpleSheetValues(values) {
  if (!Array.isArray(values)) return [];
  return values.slice(0, GENERATOR_LIMITS.maxSimpleSheetRows).map((row) => {
    const sourceRow = Array.isArray(row) ? row : [row];
    return sourceRow.slice(0, GENERATOR_LIMITS.maxSimpleSheetColumns).map((cell) => {
      if (cell === null || typeof cell === 'undefined') return '';
      if (['string', 'number', 'boolean'].indexOf(typeof cell) !== -1) return cell;
      return String(cell);
    });
  });
}

function _moveFileToFolder(fileId, folder) {
  const file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
  return file;
}

function _fileToResource(file, type) {
  return { id: file.getId(), type, name: file.getName(), url: file.getUrl(), mimeType: file.getMimeType() };
}

function _folderToResource(folder) {
  return { id: folder.getId(), type: 'folder', name: folder.getName(), url: folder.getUrl() };
}

function _asArray(value) {
  return Array.isArray(value) ? value : [];
}
