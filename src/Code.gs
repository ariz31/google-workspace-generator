/**
 * Google Workspace Generator
 *
 * Apps Script backend for creating a Google Drive workspace with optional
 * Google Docs, Sheets, and Slides resources.
 */

const APP_NAME = 'Google Workspace Generator';

const LIMITS = Object.freeze({
  maxDocuments: 20,
  maxSpreadsheets: 20,
  maxPresentations: 20,
  maxTotalFiles: 50,
  maxNameLength: 120,
  maxDescriptionLength: 500,
  maxDocumentBodyLength: 20000,
  maxSheetRows: 500,
  maxSheetColumns: 50,
});

/**
 * Serves the web app UI.
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns a safe sample configuration for the UI.
 * @return {Object}
 */
function getDefaultConfig() {
  return {
    workspaceName: 'Generated Workspace',
    description: 'Created with Google Workspace Generator.',
    documents: [
      {
        name: 'Project Brief',
        body: 'Purpose\n\nScope\n\nDeliverables\n\nTimeline',
      },
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
      {
        name: 'Overview Deck',
        title: 'Generated Workspace',
        subtitle: 'Created with Google Workspace Generator',
      },
    ],
  };
}

/**
 * Creates the configured workspace and resources.
 * This function is called by google.script.run from the UI.
 *
 * @param {Object|string} rawConfig Configuration object or JSON string.
 * @return {Object} Creation result with resource URLs.
 */
function generateWorkspace(rawConfig) {
  const config = normalizeConfig_(rawConfig);
  const errors = validateConfig_(config);

  if (errors.length > 0) {
    throw new Error('Please fix the following: ' + errors.join(' '));
  }

  const startedAt = new Date();
  const folder = DriveApp.createFolder(config.workspaceName);

  if (config.description) {
    folder.setDescription(config.description);
  }

  const resources = [];

  config.documents.forEach((documentConfig) => {
    resources.push(createDocument_(documentConfig, folder));
  });

  config.spreadsheets.forEach((spreadsheetConfig) => {
    resources.push(createSpreadsheet_(spreadsheetConfig, folder));
  });

  config.presentations.forEach((presentationConfig) => {
    resources.push(createPresentation_(presentationConfig, folder));
  });

  return {
    appName: APP_NAME,
    createdAt: startedAt.toISOString(),
    workspace: folderToResource_(folder),
    resources,
    summary: {
      totalFiles: resources.length,
      documents: config.documents.length,
      spreadsheets: config.spreadsheets.length,
      presentations: config.presentations.length,
    },
  };
}

/**
 * Normalizes raw user input into the internal configuration shape.
 * @param {Object|string} rawConfig
 * @return {Object}
 * @private
 */
function normalizeConfig_(rawConfig) {
  let input = rawConfig || {};

  if (typeof input === 'string') {
    try {
      input = JSON.parse(input);
    } catch (error) {
      throw new Error('Advanced JSON is invalid: ' + error.message);
    }
  }

  const defaults = getDefaultConfig();

  return {
    workspaceName: sanitizeName_(input.workspaceName, defaults.workspaceName),
    description: sanitizeText_(input.description || '', LIMITS.maxDescriptionLength),
    documents: normalizeDocuments_(input.documents || []),
    spreadsheets: normalizeSpreadsheets_(input.spreadsheets || []),
    presentations: normalizePresentations_(input.presentations || []),
  };
}

/**
 * @param {Array} items
 * @return {Array<Object>}
 * @private
 */
function normalizeDocuments_(items) {
  return ensureArray_(items).map((item, index) => {
    if (typeof item === 'string') {
      return {
        name: sanitizeName_(item, `Document ${index + 1}`),
        body: '',
      };
    }

    const source = item || {};
    return {
      name: sanitizeName_(source.name, `Document ${index + 1}`),
      body: sanitizeText_(source.body || '', LIMITS.maxDocumentBodyLength),
    };
  });
}

/**
 * @param {Array} items
 * @return {Array<Object>}
 * @private
 */
function normalizeSpreadsheets_(items) {
  return ensureArray_(items).map((item, index) => {
    const source = typeof item === 'string' ? { name: item } : item || {};
    const sheets = ensureArray_(source.sheets && source.sheets.length ? source.sheets : [{ name: 'Sheet1', values: [] }]);

    return {
      name: sanitizeName_(source.name, `Spreadsheet ${index + 1}`),
      sheets: sheets.map((sheet, sheetIndex) => ({
        name: sanitizeName_(sheet && sheet.name, `Sheet ${sheetIndex + 1}`),
        values: normalizeSheetValues_(sheet && sheet.values),
      })),
    };
  });
}

/**
 * @param {Array} items
 * @return {Array<Object>}
 * @private
 */
function normalizePresentations_(items) {
  return ensureArray_(items).map((item, index) => {
    if (typeof item === 'string') {
      return {
        name: sanitizeName_(item, `Presentation ${index + 1}`),
        title: item,
        subtitle: '',
      };
    }

    const source = item || {};
    return {
      name: sanitizeName_(source.name, `Presentation ${index + 1}`),
      title: sanitizeText_(source.title || source.name || `Presentation ${index + 1}`, LIMITS.maxNameLength),
      subtitle: sanitizeText_(source.subtitle || '', LIMITS.maxDescriptionLength),
    };
  });
}

/**
 * Validates normalized configuration.
 * @param {Object} config
 * @return {Array<string>}
 * @private
 */
function validateConfig_(config) {
  const errors = [];
  const totalFiles = config.documents.length + config.spreadsheets.length + config.presentations.length;

  if (!config.workspaceName) {
    errors.push('Workspace name is required.');
  }

  if (config.documents.length > LIMITS.maxDocuments) {
    errors.push(`Maximum documents allowed: ${LIMITS.maxDocuments}.`);
  }

  if (config.spreadsheets.length > LIMITS.maxSpreadsheets) {
    errors.push(`Maximum spreadsheets allowed: ${LIMITS.maxSpreadsheets}.`);
  }

  if (config.presentations.length > LIMITS.maxPresentations) {
    errors.push(`Maximum presentations allowed: ${LIMITS.maxPresentations}.`);
  }

  if (totalFiles < 1) {
    errors.push('Generate at least one file.');
  }

  if (totalFiles > LIMITS.maxTotalFiles) {
    errors.push(`Maximum total files allowed: ${LIMITS.maxTotalFiles}.`);
  }

  config.spreadsheets.forEach((spreadsheet) => {
    spreadsheet.sheets.forEach((sheet) => {
      if (sheet.values.length > LIMITS.maxSheetRows) {
        errors.push(`Sheet "${sheet.name}" exceeds the ${LIMITS.maxSheetRows}-row limit.`);
      }

      if (sheet.values.some((row) => row.length > LIMITS.maxSheetColumns)) {
        errors.push(`Sheet "${sheet.name}" exceeds the ${LIMITS.maxSheetColumns}-column limit.`);
      }
    });
  });

  return errors;
}

/**
 * Creates a Google Doc.
 * @param {Object} config
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @return {Object}
 * @private
 */
function createDocument_(config, folder) {
  const document = DocumentApp.create(config.name);
  const body = document.getBody();

  body.clear();
  body.appendParagraph(config.name).setHeading(DocumentApp.ParagraphHeading.HEADING1);

  if (config.body) {
    config.body.split('\n').forEach((line) => body.appendParagraph(line));
  } else {
    body.appendParagraph('Generated by Google Workspace Generator.');
  }

  document.saveAndClose();
  const file = moveFileToFolder_(document.getId(), folder);

  return fileToResource_(file, 'document');
}

/**
 * Creates a Google Sheet.
 * @param {Object} config
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @return {Object}
 * @private
 */
function createSpreadsheet_(config, folder) {
  const spreadsheet = SpreadsheetApp.create(config.name);
  const existingSheets = spreadsheet.getSheets();

  config.sheets.forEach((sheetConfig, index) => {
    const sheet = index === 0 ? existingSheets[0] : spreadsheet.insertSheet();
    sheet.setName(makeUniqueSheetName_(spreadsheet, sheetConfig.name, sheet));
    writeValuesToSheet_(sheet, sheetConfig.values);
  });

  spreadsheet.getSheets().forEach((sheet) => {
    if (sheet.getLastColumn() > 0) {
      sheet.autoResizeColumns(1, sheet.getLastColumn());
    }
  });

  SpreadsheetApp.flush();
  const file = moveFileToFolder_(spreadsheet.getId(), folder);

  return fileToResource_(file, 'spreadsheet');
}

/**
 * Creates a Google Slides presentation.
 * @param {Object} config
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @return {Object}
 * @private
 */
function createPresentation_(config, folder) {
  const presentation = SlidesApp.create(config.name);
  const slide = presentation.getSlides()[0];

  slide.getPageElements().forEach((element) => element.remove());
  slide.insertTextBox(config.title || config.name, 60, 80, 600, 70)
    .getText()
    .getTextStyle()
    .setFontSize(32)
    .setBold(true);

  slide.insertTextBox(config.subtitle || 'Generated by Google Workspace Generator', 60, 170, 600, 80)
    .getText()
    .getTextStyle()
    .setFontSize(16);

  presentation.saveAndClose();
  const file = moveFileToFolder_(presentation.getId(), folder);

  return fileToResource_(file, 'presentation');
}

/**
 * Moves a Drive file into the generated workspace folder.
 * @param {string} fileId
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @return {GoogleAppsScript.Drive.File}
 * @private
 */
function moveFileToFolder_(fileId, folder) {
  const file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
  return file;
}

/**
 * Writes a rectangular value matrix into a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<*>>} values
 * @private
 */
function writeValuesToSheet_(sheet, values) {
  if (!values || values.length === 0) {
    sheet.getRange(1, 1).setValue('Generated by Google Workspace Generator');
    return;
  }

  const columnCount = Math.max(1, Math.min(
    LIMITS.maxSheetColumns,
    Math.max.apply(null, values.map((row) => row.length))
  ));

  const rectangularValues = values.map((row) => {
    const normalizedRow = row.slice(0, columnCount);
    while (normalizedRow.length < columnCount) {
      normalizedRow.push('');
    }
    return normalizedRow;
  });

  sheet.getRange(1, 1, rectangularValues.length, columnCount).setValues(rectangularValues);
  sheet.setFrozenRows(1);
}

/**
 * Ensures sheet names do not collide inside one spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} preferredName
 * @param {GoogleAppsScript.Spreadsheet.Sheet} currentSheet
 * @return {string}
 * @private
 */
function makeUniqueSheetName_(spreadsheet, preferredName, currentSheet) {
  const baseName = sanitizeSheetName_(preferredName || 'Sheet');
  const existingNames = spreadsheet.getSheets()
    .filter((sheet) => sheet.getSheetId() !== currentSheet.getSheetId())
    .map((sheet) => sheet.getName());

  if (existingNames.indexOf(baseName) === -1) {
    return baseName;
  }

  let suffix = 2;
  let nextName = `${baseName} ${suffix}`;

  while (existingNames.indexOf(nextName) !== -1) {
    suffix += 1;
    nextName = `${baseName} ${suffix}`;
  }

  return nextName;
}

/**
 * @param {*} value
 * @return {Array}
 * @private
 */
function ensureArray_(value) {
  return Array.isArray(value) ? value : [];
}

/**
 * @param {*} values
 * @return {Array<Array<string|number|boolean>>}
 * @private
 */
function normalizeSheetValues_(values) {
  if (!Array.isArray(values)) {
    return [];
  }

  return values
    .slice(0, LIMITS.maxSheetRows)
    .map((row) => {
      const sourceRow = Array.isArray(row) ? row : [row];
      return sourceRow.slice(0, LIMITS.maxSheetColumns).map((cell) => {
        if (cell === null || typeof cell === 'undefined') {
          return '';
        }
        if (['string', 'number', 'boolean'].indexOf(typeof cell) !== -1) {
          return cell;
        }
        return String(cell);
      });
    });
}

/**
 * @param {*} value
 * @param {string} fallback
 * @return {string}
 * @private
 */
function sanitizeName_(value, fallback) {
  const cleaned = String(value || fallback || '')
    .replace(/[\u0000-\u001f]/g, '')
    .replace(/[\\/]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

  return cleaned.slice(0, LIMITS.maxNameLength);
}

/**
 * @param {*} value
 * @param {number} limit
 * @return {string}
 * @private
 */
function sanitizeText_(value, limit) {
  return String(value || '')
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
    .slice(0, limit);
}

/**
 * @param {string} name
 * @return {string}
 * @private
 */
function sanitizeSheetName_(name) {
  const cleaned = String(name || 'Sheet')
    .replace(/[\\/?*\[\]:]/g, '-')
    .trim()
    .slice(0, 90);

  return cleaned || 'Sheet';
}

/**
 * @param {GoogleAppsScript.Drive.File} file
 * @param {string} type
 * @return {Object}
 * @private
 */
function fileToResource_(file, type) {
  return {
    id: file.getId(),
    type,
    name: file.getName(),
    url: file.getUrl(),
    mimeType: file.getMimeType(),
  };
}

/**
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @return {Object}
 * @private
 */
function folderToResource_(folder) {
  return {
    id: folder.getId(),
    type: 'folder',
    name: folder.getName(),
    url: folder.getUrl(),
  };
}
