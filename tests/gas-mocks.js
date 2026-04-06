/**
 * GAS グローバルサービスのモック
 */

export function createMockSheet(headers = [], data = []) {
  const allValues = [headers, ...data];
  const sheet = {
    _name: 'MockSheet',
    _data: allValues,
    getDataRange: () => ({
      getValues: () => [...allValues.map(r => [...r])],
    }),
    getRange: (...args) => {
      if (args.length === 4) {
        const [row, col, numRows, numCols] = args;
        return {
          setValues: (vals) => {
            for (let r = 0; r < numRows; r++) {
              while (allValues.length <= row - 1 + r) allValues.push(new Array(headers.length).fill(''));
              for (let c = 0; c < numCols; c++) {
                allValues[row - 1 + r][col - 1 + c] = vals[r][c];
              }
            }
          },
          clearContent: () => {
            for (let r = 0; r < numRows; r++) {
              if (allValues[row - 1 + r]) {
                for (let c = 0; c < numCols; c++) {
                  allValues[row - 1 + r][col - 1 + c] = '';
                }
              }
            }
          },
          sort: () => {},
        };
      }
      if (args.length === 2) {
        const [row, col] = args;
        return {
          getValue: () => allValues[row - 1]?.[col - 1] ?? '',
        };
      }
      return { setValues: () => {}, clearContent: () => {}, sort: () => {} };
    },
    getLastRow: () => allValues.length,
    getLastColumn: () => headers.length,
    getName: () => sheet._name,
    getSheetId: () => 0,
    getParent: () => ({ getId: () => 'mock-ss-id' }),
    getFormUrl: () => null,
    appendRow: (row) => { allValues.push([...row]); },
    clearContents: () => { allValues.length = 0; },
    setFrozenRows: () => {},
    createTextFinder: (_str) => ({
      replaceAllWith: () => {},
    }),
    copyTo: () => sheet,
    hideSheet: () => sheet,
    showSheet: () => sheet,
    activate: () => {},
  };
  return sheet;
}

export function createMockSpreadsheetApp(sheets = {}) {
  const ss = {
    getSheetByName: (name) => sheets[name] || null,
    insertSheet: (name) => {
      const newSheet = createMockSheet([]);
      newSheet._name = name;
      sheets[name] = newSheet;
      return newSheet;
    },
    getSheets: () => Object.values(sheets),
    moveActiveSheet: () => {},
  };
  return {
    getActiveSpreadsheet: () => ss,
    getUi: () => ({
      createMenu: () => ({ addItem: () => ({ addItem: () => ({ addToUi: () => {} }) }), addToUi: () => {} }),
      alert: () => {},
    }),
    getActiveSheet: () => Object.values(sheets)[0] || createMockSheet([]),
    flush: () => {},
    openByUrl: () => ss,
  };
}

export function createMockUrlFetchApp() {
  const responses = [];
  return {
    _addResponse: (code, body) => { responses.push({ code, body }); },
    fetch: () => {
      const resp = responses.shift() || { code: 200, body: '{}' };
      return {
        getResponseCode: () => resp.code,
        getContentText: () => typeof resp.body === 'string' ? resp.body : JSON.stringify(resp.body),
      };
    },
  };
}

export function createMockUtilities() {
  return {
    sleep: () => {},
    formatDate: (date, _tz, _fmt) => new Date(date).toISOString(),
  };
}

export function createMockContentService() {
  return {
    createTextOutput: (text) => ({
      _text: text,
      setMimeType: function () { return this; },
    }),
    MimeType: { JSON: 'JSON' },
  };
}

// HtmlService モックは廃止（GAS WebApp → GitHub Pages移行済み）
