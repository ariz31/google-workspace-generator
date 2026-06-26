/**
 * @OnlyCurrentDoc
 *
 * This script combines a QR Code generator and a unified Document, Slide, and Sheet generator
 * into a single productivity suite.
 *
 * SCRIPT AUTHORIZATION SCOPES:
 * @AuthScope https://www.googleapis.com/auth/spreadsheets
 * @AuthScope https://www.googleapis.com/auth/drive
 * @AuthScope https://www.googleapis.com/auth/script.container.ui
 * @AuthScope https://www.googleapis.com/auth/userinfo.email
 * @AuthScope https://www.googleapis.com/auth/script.send_mail
 * @AuthScope https://www.googleapis.com/auth/documents
 * @AuthScope https://www.googleapis.com/auth/presentations
 * @AuthScope https://www.googleapis.com/auth/script.external_request
 * @AuthScope https://www.googleapis.com/auth/script.scriptapp
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Productivity Suite')
    .addItem('Create QR Code', 'insertQRCode')
    .addSeparator()
    .addItem('Open Generator', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Generator Suite').setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function insertQRCode() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  // Fixed to getDisplayValue so numeric formats aren't stripped before QR generation
  const data = cell.getDisplayValue().trim();
  if (!data) {
    SpreadsheetApp.getUi().alert('The selected cell is empty.');
    return;
  }
  const size = 300;
  const googleUrl = `https://chart.googleapis.com/chart?cht=qr&chs=${size}x${size}&chl=${encodeURIComponent(data)}`;
  try {
    let response = UrlFetchApp.fetch(googleUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      const qrserverUrl = `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodeURIComponent(data)}`;
      response = UrlFetchApp.fetch(qrserverUrl);
    }
    const blob = response.getBlob().setName('qrcode.png');
    sheet.insertImage(blob, cell.getColumn(), cell.getRow());
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Failed to generate QR code: ${e.message}`);
  }
}

//==================================================================================================
// REGION: ASYNCHRONOUS GENERATOR ENGINE
//==================================================================================================

const MimeType = {
    GOOGLE_DOCS: 'application/vnd.google-apps.document',
    GOOGLE_SLIDES: 'application/vnd.google-apps.presentation',
    GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet',
    PDF: 'application/pdf'
};

function runGenerator(options, sheetName) {
  _deleteTrigger();
  PropertiesService.getUserProperties().deleteAllProperties();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const allSelectedItems = (sheetName === 'R-DOC-GEN') 
    ? Array.from(_getSelectedRows(ss, sheetName)).sort((a,b) => a - b)
    : Array.from(_getSelectedColumns(ss, sheetName)).sort((a,b) => a - b);

  if (allSelectedItems.length === 0) {
    throw new Error(`No items selected in '${sheetName}'. Please highlight rows/columns to process.`);
  }
  
  // CRITICAL FIX: getDisplayValues() pulls numbers/dates exactly as formatted, bypassing numeric crashes
  const sheetData = sheet.getDataRange().getDisplayValues();
  const placeholderMap = (sheetName === 'R-DOC-GEN') 
      ? _buildPlaceholderMap(sheetData[0], 4) 
      : _buildColumnPlaceholderMap(sheetData, 4);
      
  const emailSheet = ss.getSheetByName('EMAIL');
  const emailTemplates = {
      subject: options.sendEmail && emailSheet ? emailSheet.getRange('A2').getValue() : '',
      body: options.sendEmail && emailSheet ? emailSheet.getRange('A3').getValue() : ''
  };

  const relevantData = (sheetName === 'R-DOC-GEN')
    ? allSelectedItems.map(rowNum => sheetData[rowNum - 1])
    : allSelectedItems.map(colNum => sheetData.map(row => row[colNum - 1]));

  const baseState = {
    items: allSelectedItems,
    currentIndex: 0,
    totalItems: allSelectedItems.length,
    options: options,
    sheetName: sheetName,
    results: [],
    folders: [],
    status: 'running',
    placeholderMap: placeholderMap,
    emailTemplates: emailTemplates,
    relevantData: relevantData
  };
  
  const itemTypes = _getItemTypes(allSelectedItems, sheet, sheetData);
  let isCombinedRun = false;

  if (options.sheetOutputMode === 'combined' && itemTypes.sheets.length > 0) {
      _initializeCombinedRun('sheets', itemTypes, baseState, sheetData);
      isCombinedRun = true;
  } else if (options.docSlideOutputMode === 'combined' && (itemTypes.docs.length > 0 || itemTypes.slides.length > 0)) {
      if (itemTypes.docs.length > 0 && itemTypes.slides.length > 0) {
          throw new Error("Cannot combine Docs and Slides. Please select only one type for combined mode.");
      }
      const combineType = itemTypes.docs.length > 0 ? 'docs' : 'slides';
      _initializeCombinedRun(combineType, itemTypes, baseState, sheetData);
      isCombinedRun = true;
  }

  if (!isCombinedRun) {
      PropertiesService.getUserProperties().setProperty('generationState', JSON.stringify(baseState));
  }
  
  continueGeneration();
  const finalState = JSON.parse(PropertiesService.getUserProperties().getProperty('generationState'));
  return { status: 'started', totalItems: finalState.totalItems };
}

function getGenerationStatus() {
  const stateProperty = PropertiesService.getUserProperties().getProperty('generationState');
  return stateProperty ? JSON.parse(stateProperty) : { status: 'idle' };
}

function cancelGeneration() {
  _deleteTrigger();
  PropertiesService.getUserProperties().deleteAllProperties();
  return { status: 'cancelled' };
}

function continueGeneration() {
  const userProperties = PropertiesService.getUserProperties();
  const stateProperty = userProperties.getProperty('generationState');
  if (!stateProperty) return;

  let state = JSON.parse(stateProperty);
  const BATCH_START_TIME = new Date().getTime();
  const TIME_LIMIT_MS = 280000; 
  const batchSize = state.options.batchSize || 5;
  let itemsProcessedThisRun = 0;
  
  const usedSheetNames = state.isCombinedRun ? new Set(state.usedSheetNames) : null;

  while (state.currentIndex < state.totalItems) {
    if (new Date().getTime() - BATCH_START_TIME > TIME_LIMIT_MS) break; 
    if (itemsProcessedThisRun >= batchSize) break;

    const itemNum = state.items[state.currentIndex];
    const itemData = state.relevantData[state.currentIndex];

    if (state.isCombinedRun) {
      _processCombinedItem(state, itemNum, itemData, usedSheetNames);
    } else {
      _processSeparateItem(state, itemNum, itemData);
    }

    state.currentIndex++;
    itemsProcessedThisRun++;
  }

  if (state.currentIndex < state.totalItems) {
    if (state.isCombinedRun) {
        state.usedSheetNames = Array.from(usedSheetNames);
    }
    userProperties.setProperty('generationState', JSON.stringify(state));
    _createTrigger();
  } else {
    _deleteTrigger();
    state.status = 'complete';
    if (state.isCombinedRun) {
        state.usedSheetNames = Array.from(usedSheetNames);
    }
    _finalizeCombinedRun(state);
    userProperties.setProperty('generationState', JSON.stringify(state));
  }
}

function _createTrigger() {
  _deleteTrigger();
  ScriptApp.newTrigger('continueGeneration')
      .timeBased()
      .after(1000) 
      .create();
}

function _deleteTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueGeneration') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

//==================================================================================================
// REGION: HELPER FUNCTIONS
//==================================================================================================

function getSelectionDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  if (sheetName !== 'R-DOC-GEN' && sheetName !== 'C-DOC-GEN') {
    return { sheetName: sheetName, mimeTypes: [], selectedIndices: [] };
  }
  
  const uniqueMimeTypes = new Set();
  const selectedIndices = new Set(); 
  const rangeList = ss.getActiveRangeList();
  if (!rangeList) return { sheetName, mimeTypes: [], selectedIndices: [] };
  
  const ranges = rangeList.getRanges().filter(r => r.getSheet().getName() === sheetName);
  if (ranges.length === 0) return { sheetName, mimeTypes: [], selectedIndices: [] };

  try {
    if (sheetName === 'R-DOC-GEN') {
      ranges.forEach(range => {
        for (let i = 0; i < range.getNumRows(); i++) {
          const row = range.getRow() + i;
          if (row < 3) continue; 
          selectedIndices.add(row);
          const templateId = activeSheet.getRange(row, 1).getValue();
          if (templateId) {
            uniqueMimeTypes.add(DriveApp.getFileById(templateId.toString().trim()).getMimeType());
          }
        }
      });
    } else { 
       ranges.forEach(range => {
        for (let i = 0; i < range.getNumColumns(); i++) {
          const col = range.getColumn() + i;
          if (col < 2) continue; 
          selectedIndices.add(col);
          const templateId = activeSheet.getRange(1, col).getValue();
           if (templateId) {
            uniqueMimeTypes.add(DriveApp.getFileById(templateId.toString().trim()).getMimeType());
          }
        }
      });
    }
    return { 
      sheetName: sheetName, 
      mimeTypes: Array.from(uniqueMimeTypes),
      selectedIndices: Array.from(selectedIndices).sort((a, b) => a - b) 
    };
  } catch (e) {
    return { sheetName: sheetName, mimeTypes: [], selectedIndices: [] };
  }
}

function _processSingleItem(itemIdentifier, templateId, data, placeholderMap, options, subjectTemplate, bodyTemplate) {
  const { format, sheetTarget, newSheetName, sendEmail } = options;
  
  const folderId = (data[1] || '').trim();
  const recipientEmail = (data[2] || '').trim();
  const fileNameValue = (data[3] && data[3].trim() !== '' ? data[3] : itemIdentifier).trim();

  let logData = { fileName: fileNameValue, recipient: recipientEmail || 'N/A' };

  if (!templateId || !folderId) {
    const errorMsg = !templateId ? 'Missing Template ID' : 'Missing Folder ID';
    _logActivity({ ...logData, generator: 'N/A', status: 'Error', details: errorMsg });
    return { result: { item: itemIdentifier, status: `❌ ${errorMsg}` }, folderId: null };
  }

  try {
    const templateFile = DriveApp.getFileById(templateId);
    const mimeType = templateFile.getMimeType();
    const destinationFolder = DriveApp.getFolderById(folderId);
    logData.folderUrl = destinationFolder.getUrl();
    
    const newFileName = `${templateFile.getName()} - ${fileNameValue}`;
    logData.fileName = newFileName;
    const newFile = templateFile.makeCopy(newFileName, destinationFolder);
    
    let finalFile = newFile;
    let status;

    if (mimeType === MimeType.GOOGLE_DOCS || mimeType === MimeType.GOOGLE_SLIDES) {
      logData.generator = 'Doc/Slide';
      const result = _generateDocOrSlide(newFile, mimeType, placeholderMap, data, format, folderId);
      finalFile = result.file;
      status = result.status;
    } else if (mimeType === MimeType.GOOGLE_SHEETS) {
      logData.generator = 'Sheet';
      const newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
      _generateSheet(newSpreadsheet, templateId, placeholderMap, data, sheetTarget, newSheetName);
      status = '✅ Success';
    } else {
      throw new Error(`Unsupported template type.`);
    }
    
    logData.fileUrl = finalFile.getUrl();
    logData.emailStatus = 'Not Applicable';

    if (recipientEmail && sendEmail) {
      const emailResult = _sendEmailNotification(finalFile, recipientEmail, subjectTemplate, bodyTemplate, data, placeholderMap);
      status = emailResult.status;
      logData.emailStatus = emailResult.logStatus;
    }
    
    _logActivity({ ...logData, status: 'Success', details: status.substring(2) });
    return { result: { item: itemIdentifier, status, url: finalFile.getUrl() }, folderId: folderId };

  } catch (e) {
    _logActivity({ ...logData, generator: 'N/A', status: 'Error', details: e.message });
    return { result: { item: itemIdentifier, status: `❌ ${e.message}` }, folderId: folderId };
  }
}

function _generateDocOrSlide(newFile, mimeType, placeholderMap, row, format, folderId) {
  const copyId = newFile.getId();
  if (mimeType === MimeType.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(copyId);
    [doc.getHeader(), doc.getBody(), doc.getFooter()].forEach(sec => {
      _replacePlaceholdersInSection(sec, placeholderMap, row);
    });
    doc.saveAndClose();
  } else { 
    const pres = SlidesApp.openById(copyId);
    _replaceInPresentation(pres, placeholderMap, row);
    pres.saveAndClose();
  }

  let finalFile = newFile;
  if (format === 'PDF') {
    const pdfBlob = newFile.getBlob().getAs(MimeType.PDF);
    const pdfFile = DriveApp.getFolderById(folderId).createFile(pdfBlob).setName(newFile.getName() + '.pdf');
    finalFile = pdfFile;
    newFile.setTrashed(true);
  }
  return { file: finalFile, status: `✅ Success (${format})` };
}

function _generateSheet(targetSpreadsheet, templateId, placeholderMap, rowData, sheetTarget, newSheetName, isCombined = false, usedSheetNames = null) {
  const templateSpreadsheet = SpreadsheetApp.openById(templateId);
  
  let sheetsToProcess = [];
  if (sheetTarget) {
    const targetSheet = templateSpreadsheet.getSheetByName(sheetTarget);
    if (!targetSheet) throw new Error(`Sheet "${sheetTarget}" not found in template.`);
    sheetsToProcess.push(targetSheet);
  } else {
    sheetsToProcess = templateSpreadsheet.getSheets();
  }
  
  const originalSheetNames = isCombined ? [] : targetSpreadsheet.getSheets().map(s => s.getName());
  
  const rowPlaceholderMap = new Map();
  Object.keys(placeholderMap).forEach(base => {
    const placeholder = `{{${base}}}`;
    const index = placeholderMap[base];
    rowPlaceholderMap.set(placeholder, rowData[index]);
  });
  
  sheetsToProcess.forEach(templateSheet => {
    const copiedSheet = templateSheet.copyTo(targetSpreadsheet);
    let finalSheetName = templateSheet.getName();

    if (isCombined) {
        const baseName = newSheetName || templateSheet.getName();
        finalSheetName = _getUniqueSheetName(baseName, usedSheetNames);
        usedSheetNames.add(finalSheetName);
    } else if (sheetsToProcess.length === 1 && newSheetName) {
        finalSheetName = newSheetName;
    }
    
    copiedSheet.setName(finalSheetName);
    replacePlaceholdersInSheet(copiedSheet, rowPlaceholderMap);
  });
  
  if (!isCombined) {
    originalSheetNames.forEach(sheetName => {
        const sheetToDelete = targetSpreadsheet.getSheetByName(sheetName);
        if (sheetToDelete) targetSpreadsheet.deleteSheet(sheetToDelete);
    });
  }
}

function _getUniqueSheetName(baseName, usedNamesSet) {
  if (!usedNamesSet.has(baseName)) return baseName;
  let i = 1;
  while (true) {
    const newName = `${baseName} (${i})`;
    if (!usedNamesSet.has(newName)) return newName;
    i++;
  }
}

function _sendEmailNotification(file, recipient, subjectTemplate, bodyTemplate, rowData, placeholderMap) {
  try {
    file.addEditor(recipient);
    const subject = _replacePlaceholdersInEmail(subjectTemplate, rowData, placeholderMap);
    const body = _replacePlaceholdersInEmail(bodyTemplate, rowData, placeholderMap);
    MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: body, attachments: [file.getBlob()] });
    return { status: `✅ Emailed to ${recipient}`, logStatus: 'Sent' };
  } catch (e) {
    return { status: `⚠️ Success, but email failed: ${e.message}`, logStatus: `Failed: ${e.message}` };
  }
}

function _getLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheetName = 'Log';
  let logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
    const headers = ['Timestamp', 'Generator', 'File Name', 'File URL', 'Folder URL', 'Recipient', 'Email Status', 'Status', 'Details'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }
  return logSheet;
}

function _logActivity(logData) {
  const logSheet = _getLogSheet();
  const timestamp = new Date();
  const logEntry = {
    generator: 'N/A', fileName: 'N/A', fileUrl: 'N/A', folderUrl: 'N/A',
    recipient: 'N/A', emailStatus: 'N/A', status: 'Unknown', details: '', ...logData
  };
  logSheet.appendRow([
    timestamp, logEntry.generator, logEntry.fileName, logEntry.fileUrl,
    logEntry.folderUrl, logEntry.recipient, logEntry.emailStatus,
    logEntry.status, logEntry.details
  ]);
}

function _applyModifiers(rawValue, mod) {
  if (rawValue === null || rawValue === undefined || rawValue === '') return '';
  
  let str = String(rawValue); // Safely convert any type to string
  const hasLetters = /[a-zA-Z]/.test(str);
  
  if (!hasLetters) {
    const dateLikeRegex = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(\.\d+)?Z?)?$|^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}$/;
    if (dateLikeRegex.test(str)) {
      const parsed = new Date(str);
      if (!isNaN(parsed.getTime())) {
        return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
      }
    }
  }
  
  const proper = s => s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
  switch (mod) {
    case 'u': case 'ub': str = str.toUpperCase(); break;
    case '1': case '1b': str = proper(str); break;
  }
  return str;
}

function _isDriveFileId(str) {
  return typeof str === 'string' && /^[a-zA-Z0-9_-]{25,}$/.test(str);
}

// ------------------------------------------------------------------
// FIXED: _replaceInSection (Handles image validation and URL linking)
// ------------------------------------------------------------------
function _replaceInSection(section, base, rawValue) {
  if (!section) return;
  const modPattern = 'u|ub|b|1|1b';
  const regexStr = `\\{\\{${base}(?:\\}(${modPattern}))?\\}\\}`;
  let found = section.findText(regexStr);
  
  while (found) {
    const el = found.getElement().asText();
    const start = found.getStartOffset();
    const end = found.getEndOffsetInclusive();
    const elText = el.getText();
    
    // Ensure the placeholder match doesn't span across multiple boundaries.
    if (start < 0 || end >= elText.length || start > end) {
        found = section.findText(regexStr, found);
        continue;
    }

    const matchedText = elText.substring(start, end + 1);
    const modMatch = matchedText.match(new RegExp(`^${regexStr}$`));
    const mod = (modMatch && modMatch[1]) || '';
    
    let isImage = false, blob;
    if (typeof rawValue === 'string') {
      if (/^https?:\/\//i.test(rawValue)) {
        try { 
          const response = UrlFetchApp.fetch(rawValue, {muteHttpExceptions: true});
          if (response.getResponseCode() === 200) {
             const tempBlob = response.getBlob();
             // Prevent crash: Ensure the URL actually points to an image
             if (tempBlob.getContentType().startsWith('image/')) {
                 blob = tempBlob; 
                 isImage = true;
             }
          }
        } catch (_) {}
      } else if (_isDriveFileId(rawValue)) {
        try { 
          const file = DriveApp.getFileById(rawValue);
          // Prevent crash: Ensure the Drive file is actually an image
          if (file.getMimeType().startsWith('image/')) {
             blob = file.getBlob(); 
             isImage = true; 
          }
        } catch (_) {}
      }
    }
    
    // Capture any hyperlink already applied to the template placeholder
    const existingUrl = el.getLinkUrl(start);
    el.deleteText(start, end);
    
    if (isImage && blob) {
      const parent = el.getParent().asParagraph();
      const idx = parent.getChildIndex(el);
      const full = el.getText();
      const before = full.substring(0, start);
      const after = full.substring(start);
      parent.removeChild(el);
      if (before) parent.insertText(idx, before);
      parent.insertInlineImage(idx + (before ? 1 : 0), blob);
      if (after) parent.insertText(idx + (before ? 2 : 1), after);
    } else {
      const txt = _applyModifiers(rawValue, mod);
      el.insertText(start, txt);
      
      if (txt.length > 0) {
        // Re-apply Bold
        if (mod.includes('b')) {
          el.setBold(start, start + txt.length - 1, true);
        }
        // Restore template link, OR auto-link if the value is a standard web address
        if (existingUrl) {
          el.setLinkUrl(start, start + txt.length - 1, existingUrl);
        } else if (/^https?:\/\//i.test(txt)) {
          el.setLinkUrl(start, start + txt.length - 1, txt);
        }
      }
    }
    found = section.findText(regexStr, found);
  }
}

function _replaceInPresentation(presentation, placeholderMap, rowData) {
  presentation.getSlides().forEach(slide => {
    Object.keys(placeholderMap).forEach(base => {
      const rawValue = rowData[placeholderMap[base]];
      if (rawValue === null || rawValue === undefined || rawValue === '') return;
      const replacedAsImage = _replacePlaceholderAsImageInSlide(slide, base, rawValue);
      if (!replacedAsImage) {
        _replacePlaceholderAsTextInSlide(slide, base, rawValue);
      }
    });
  });
}

// ------------------------------------------------------------------
// FIXED: _replacePlaceholderAsImageInSlide 
// ------------------------------------------------------------------
function _replacePlaceholderAsImageInSlide(slide, base, rawValue) {
  let blob;
  if (typeof rawValue === 'string') {
    if (/^https?:\/\//i.test(rawValue)) {
      try { 
        const response = UrlFetchApp.fetch(rawValue, {muteHttpExceptions: true});
        if (response.getResponseCode() === 200) {
          const tempBlob = response.getBlob();
          if (tempBlob.getContentType().startsWith('image/')) {
             blob = tempBlob;
          }
        }
      } catch (e) {}
    } else if (_isDriveFileId(rawValue)) {
      try { 
        const file = DriveApp.getFileById(rawValue);
        if (file.getMimeType().startsWith('image/')) {
           blob = file.getBlob();
        }
      } catch (e) {}
    }
  }
  
  if (!blob) return false;
  
  const placeholder = `{{${base}}}`;
  let shapeToReplace = null;
  slide.getShapes().forEach(shape => {
    if (shape.getText && shape.getText().asString().trim() === placeholder) {
      shapeToReplace = shape;
    }
  });
  
  if (shapeToReplace) {
    const img = slide.insertImage(blob);
    img.setWidth(shapeToReplace.getWidth());
    img.setHeight(shapeToReplace.getHeight());
    img.setLeft(shapeToReplace.getLeft());
    img.setTop(shapeToReplace.getTop());
    shapeToReplace.remove();
    return true;
  }
  return false;
}

// ------------------------------------------------------------------
// FIXED: _replacePlaceholderAsTextInSlide 
// ------------------------------------------------------------------
function _replacePlaceholderAsTextInSlide(slide, base, rawValue) {
  const modPattern = 'u|ub|b|1|1b';
  const regex = new RegExp(`\\{\\{${base}(?:\\}(${modPattern}))?\\}\\}`, 'g');
  const elements = [...slide.getShapes(), ...slide.getTables().flatMap(table => {
    const cells = [];
    for (let r = 0; r < table.getNumRows(); r++) {
      for (let c = 0; c < table.getNumColumns(); c++) {
        cells.push(table.getCell(r, c));
      }
    }
    return cells;
  })];
  
  elements.forEach(element => {
    try {
      if (!element.getText || !element.getText().asString()) return;
      const textRange = element.getText();
      const matches = [];
      let match;
      while ((match = regex.exec(textRange.asString())) !== null) {
        matches.push(match);
      }
      
      for (let i = matches.length - 1; i >= 0; i--) {
        match = matches[i];
        const placeholder = match[0];
        const mod = match[1] || '';
        const replacementText = _applyModifiers(rawValue, mod);
        const startIndex = match.index;
        const endIndex = startIndex + placeholder.length;
        
        const specificRange = textRange.getRange(startIndex, endIndex);
        
        // Capture existing link attached to the Slide placeholder
        let existingUrl = null;
        try {
          const link = specificRange.getTextStyle().getLink();
          if (link) existingUrl = link.getUrl();
        } catch(e) {}

        specificRange.setText(replacementText);
        
        if (replacementText.length > 0) {
          const newRange = textRange.getRange(startIndex, startIndex + replacementText.length);
          
          if (mod.includes('b')) {
            newRange.getTextStyle().setBold(true);
          }
          
          // Apply link back to text
          if (existingUrl) {
            newRange.getTextStyle().setLinkUrl(existingUrl);
          } else if (/^https?:\/\//i.test(replacementText)) {
            newRange.getTextStyle().setLinkUrl(replacementText);
          }
        }
      }
    } catch (e) {}
  });
}

function _replacePlaceholdersInEmail(template, rowData, placeholderMap) {
  if (!template) return '';
  let result = template;
  Object.keys(placeholderMap).forEach(base => {
    const modPattern = 'u|ub|b|1|1b';
    const regex = new RegExp(`\\{\\{${base}(?:\\}(${modPattern}))?\\}\\}`, 'g');
    result = result.replace(regex, (match, mod = '') => {
      const rawValue = rowData[placeholderMap[base]];
      let processedValue = _applyModifiers(rawValue, mod);
      if (mod.includes('b')) {
        processedValue = `<b>${processedValue}</b>`;
      }
      return processedValue;
    });
  });
  return result;
}

function replacePlaceholdersInSheet(sheet, placeholderMap) {
  placeholderMap.forEach((value, placeholder) => {
    let displayValue = value != null ? String(value) : '';
    sheet.createTextFinder(placeholder).replaceAllWith(displayValue);
  });
}

function _getSelectedRows(ss, sheetName) {
    const selectedRows = new Set();
    const rl = ss.getActiveRangeList();
    if (rl) {
        rl.getRanges().forEach(r => {
            if (r.getSheet().getName() === sheetName) {
                for (let i = 0; i < r.getNumRows(); i++) {
                    const rn = r.getRow() + i;
                    if (rn >= 3) selectedRows.add(rn);
                }
            }
        });
    }
    return selectedRows;
}

function _getSelectedColumns(ss, sheetName) {
    const selectedCols = new Set();
    const rl = ss.getActiveRangeList();
    if (rl) {
        rl.getRanges().forEach(r => {
            if (r.getSheet().getName() === sheetName) {
                for (let i = 0; i < r.getNumColumns(); i++) {
                    const cn = r.getColumn() + i;
                    if (cn >= 2) selectedCols.add(cn);
                }
            }
        });
    }
    return selectedCols;
}

function _buildPlaceholderMap(headerRow, startIndex) {
    const placeholderMap = {};
    headerRow.forEach((h, c) => {
        const m = (h || '').toString().trim().match(/^\{\{(.+?)\}\}$/);
        if (m && c >= startIndex) placeholderMap[m[1]] = c;
    });
    return placeholderMap;
}

function _buildColumnPlaceholderMap(data, startIndex) {
    const placeholderRowMap = {};
    data.forEach((row, r) => {
        const m = (row[0] || '').toString().trim().match(/^\{\{(.+?)\}\}$/);
        if (m && r >= startIndex) placeholderRowMap[m[1]] = r;
    });
    return placeholderRowMap;
}

function _getItemTypes(selectedItems, sheet, data) {
    const types = { docs: [], slides: [], sheets: [] };
    selectedItems.forEach(itemNum => {
        let templateId;
        if (sheet.getName() === 'R-DOC-GEN') {
            templateId = data[itemNum - 1][0];
        } else { 
            templateId = data[0][itemNum - 1];
        }
        try {
            if (templateId) {
                const mimeType = DriveApp.getFileById(templateId).getMimeType();
                if (mimeType === MimeType.GOOGLE_DOCS) types.docs.push(itemNum);
                else if (mimeType === MimeType.GOOGLE_SLIDES) types.slides.push(itemNum);
                else if (mimeType === MimeType.GOOGLE_SHEETS) types.sheets.push(itemNum);
            }
        } catch(e) {}
    });
    return types;
}

function _initializeCombinedRun(type, itemTypes, baseState, data) {
    const itemsToCombine = (type === 'docs') ? itemTypes.docs : (type === 'slides') ? itemTypes.slides : itemTypes.sheets;
    const firstItemNum = itemsToCombine[0];
    
    let folderId, firstTemplateId, firstItemData;
    if (baseState.sheetName === 'R-DOC-GEN') {
        firstTemplateId = data[firstItemNum - 1][0];
        folderId = data[firstItemNum - 1][1];
        firstItemData = data[firstItemNum - 1];
    } else { 
        firstTemplateId = data[0][firstItemNum - 1];
        folderId = data[1][firstItemNum - 1];
        firstItemData = data.map(row => row[firstItemNum - 1]);
    }
    if (!folderId) throw new Error("The first selected item for combining is missing a Folder ID.");
    
    const destinationFolder = DriveApp.getFolderById(folderId);

    if (type === 'sheets') {
        const newFileName = baseState.options.combinedSheetName || 'Combined Spreadsheet';
        const newSpreadsheet = SpreadsheetApp.create(newFileName);
        const newFile = DriveApp.getFileById(newSpreadsheet.getId());
        newFile.moveTo(destinationFolder);
        baseState.combinedFileId = newFile.getId();
    } else { 
        const templateFile = DriveApp.getFileById(firstTemplateId);
        const newFileName = baseState.options.combinedDocSlideName || 'Combined Document';
        const newFile = templateFile.makeCopy(newFileName, destinationFolder);
        baseState.combinedFileId = newFile.getId();
        
        if (type === 'docs') {
            const mDoc = DocumentApp.openById(newFile.getId());
            mDoc.getBody().clear(); 
            [mDoc.getHeader(), mDoc.getFooter()].forEach(sec => {
                 _replacePlaceholdersInSection(sec, baseState.placeholderMap, firstItemData);
            });
            mDoc.saveAndClose();
        } else {
            const mPres = SlidesApp.openById(newFile.getId());
            mPres.getSlides().forEach(s => s.remove()); 
            mPres.saveAndClose();
        }
    }

    baseState.items = itemsToCombine;
    baseState.totalItems = itemsToCombine.length;
    baseState.folders = [{id: folderId, url: destinationFolder.getUrl()}];
    baseState.isCombinedRun = true;
    baseState.combinedRunType = type;
    baseState.usedSheetNames = [];

    const itemSet = new Set(itemsToCombine);
    if(baseState.sheetName === 'R-DOC-GEN'){
        baseState.relevantData = Array.from(_getSelectedRows(SpreadsheetApp.getActive(), baseState.sheetName))
                                   .filter(rowNum => itemSet.has(rowNum))
                                   .map(rowNum => data[rowNum-1]);
    } else {
        baseState.relevantData = Array.from(_getSelectedColumns(SpreadsheetApp.getActive(), baseState.sheetName))
                                   .filter(colNum => itemSet.has(colNum))
                                   .map(colNum => data.map(row => row[colNum-1]));
    }

    PropertiesService.getUserProperties().setProperty('generationState', JSON.stringify(baseState));
}

function _processCombinedItem(state, itemNum, itemData, usedSheetNames) {
    let status = '❌ Failed';
    let itemIdentifier = '';
    
    try {
        let templateId;
        if (state.sheetName === 'R-DOC-GEN') {
            itemIdentifier = `Row ${itemNum}`;
            templateId = itemData[0];
        } else { 
            itemIdentifier = `Column ${String.fromCharCode(65 + itemNum -1)}`;
            templateId = itemData[0];
        }

        if (state.combinedRunType === 'sheets') {
            const combinedSpreadsheet = SpreadsheetApp.openById(state.combinedFileId);
            const sheetNameBase = itemData[3] || `${itemIdentifier}_Sheet`;
            _generateSheet(combinedSpreadsheet, templateId, state.placeholderMap, itemData, state.options.sheetTarget, sheetNameBase, true, usedSheetNames);
        } else { 
            const templateFile = DriveApp.getFileById(templateId);
            const tempDriveFile = templateFile.makeCopy(`temp_${itemIdentifier}`);
            const tempFileId = tempDriveFile.getId();
            
            try {
                if (state.combinedRunType === 'docs') {
                    const tempDoc = DocumentApp.openById(tempFileId);
                    [tempDoc.getHeader(), tempDoc.getBody(), tempDoc.getFooter()].forEach(sec => {
                        _replacePlaceholdersInSection(sec, state.placeholderMap, itemData);
                    });
                    tempDoc.saveAndClose();
                    
                    const masterDoc = DocumentApp.openById(state.combinedFileId);
                    _appendDocContent(masterDoc, DocumentApp.openById(tempFileId), state.currentIndex > 0);
                    masterDoc.saveAndClose();
                    
                } else { 
                    const tempPres = SlidesApp.openById(tempFileId);
                    _replaceInPresentation(tempPres, state.placeholderMap, itemData);
                    tempPres.saveAndClose();
                    
                    const masterPres = SlidesApp.openById(state.combinedFileId);
                    _appendSlides(masterPres, SlidesApp.openById(tempFileId));
                    masterPres.saveAndClose();
                }
            } finally {
                tempDriveFile.setTrashed(true);
            }
        }
        status = `✅ Success (Combined)`;
    } catch(e) {
        status = `❌ Error: ${e.message}`;
    }
    
    const combinedFile = DriveApp.getFileById(state.combinedFileId);
    state.results.push({ item: itemIdentifier, status: status, url: combinedFile.getUrl() });
}

function _processSeparateItem(state, itemNum, itemData) {
    let result;
    if (state.sheetName === 'R-DOC-GEN') {
        const templateId = (itemData[0] || '').toString().trim();
        if (templateId) {
            result = _processSingleItem(`Row ${itemNum}`, templateId, itemData, state.placeholderMap, state.options, state.emailTemplates.subject, state.emailTemplates.body);
        }
    } else { 
        const templateId = (itemData[0] || '').toString().trim();
        if (templateId) {
            result = _processSingleItem(`Column ${String.fromCharCode(65 + itemNum -1)}`, templateId, itemData, state.placeholderMap, state.options, state.emailTemplates.subject, state.emailTemplates.body);
        }
    }
    if (result) {
        state.results.push(result.result);
        if(result.folderId) {
            const folderSet = new Set(state.folders.map(f => f.id));
            if(!folderSet.has(result.folderId)){
                state.folders.push({id: result.folderId, url: `https://drive.google.com/drive/folders/${result.folderId}`});
            }
        }
    }
}

function _finalizeCombinedRun(state) {
  if (!state.isCombinedRun) return;

  const combinedFile = DriveApp.getFileById(state.combinedFileId);
  let finalFileUrl = combinedFile.getUrl();

  if (state.combinedRunType === 'sheets') {
      const combinedSpreadsheet = SpreadsheetApp.openById(state.combinedFileId);
      const defaultSheet = combinedSpreadsheet.getSheetByName('Sheet1');
      if (defaultSheet && combinedSpreadsheet.getSheets().length > 1) {
          combinedSpreadsheet.deleteSheet(defaultSheet);
      }
  } else if (state.combinedRunType === 'slides') {
      const combinedPresentation = SlidesApp.openById(state.combinedFileId);
      if (combinedPresentation.getSlides().length > 1) {
          const defaultSlide = combinedPresentation.getSlides()[0];
          if (defaultSlide.getShapes().length === 0 && defaultSlide.getMasters()[0].getShapes().length === 0) {
             defaultSlide.remove();
          }
      }
  }

  if ((state.combinedRunType === 'docs' || state.combinedRunType === 'slides') && state.options.format === 'PDF') {
      const pdfBlob = combinedFile.getBlob().getAs(MimeType.PDF);
      const pdfFile = DriveApp.getFolderById(state.folders[0].id).createFile(pdfBlob).setName(combinedFile.getName() + '.pdf');
      finalFileUrl = pdfFile.getUrl();
      combinedFile.setTrashed(true); 
  }
  
   _logActivity({
      generator: `Combined ${state.combinedRunType}`, 
      fileName: state.options.combinedDocSlideName || state.options.combinedSheetName || 'Combined Output', 
      fileUrl: finalFileUrl,
      folderUrl: state.folders.length > 0 ? state.folders[0].url : '', 
      status: 'Success',
      details: `Combined ${state.results.filter(r => r.status.startsWith('✅')).length} of ${state.totalItems} items.`
  });
}

function _replacePlaceholdersInSection(section, placeholderMap, data) {
    if (!section) return;
    Object.keys(placeholderMap).forEach(base => {
        const index = placeholderMap[base];
        _replaceInSection(section, base, data[index]);
    });
}

function _appendDocContent(masterDoc, sourceDoc, isNotFirstItem) {
    const masterBody = masterDoc.getBody();
    if (isNotFirstItem) {
        masterBody.appendPageBreak();
    }
    
    const sourceBody = sourceDoc.getBody();
    const numChildren = sourceBody.getNumChildren();
    for (let i = 0; i < numChildren; i++) {
        const child = sourceBody.getChild(i);
        const type = child.getType();
        
        if (type === DocumentApp.ElementType.PARAGRAPH) {
            masterBody.appendParagraph(child.copy());
        } else if (type === DocumentApp.ElementType.TABLE) {
            masterBody.appendTable(child.copy());
        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
            masterBody.appendListItem(child.copy());
        }
    }

    if (!isNotFirstItem && masterBody.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && masterBody.getChild(0).asText().getText() === '') {
        masterBody.removeChild(masterBody.getChild(0));
    }
}

function _appendSlides(masterPres, sourcePres) {
    const slides = sourcePres.getSlides();
    for (let i = 0; i < slides.length; i++) {
        masterPres.insertSlide(masterPres.getSlides().length, slides[i]);
    }
}
