// https://github.com/TomoBossi/upsertable

const CONFIG = {
  SPREADSHEET_LOCATOR_LEADING_CHARACTER: "\\@", // Because it will be used inside a RegExp, the leading character may need to be escaped (\\)
  QUERY_STRING_LEADING_CHARACTER: "?",
  QUERY_STRING_PARAMETER_SEPARATOR: "%",
  QUERY_STRING_PARAMETER_ASSIGNMENT_OPERATOR: "=",
  QUERY_STRING_INTERFILTER_SEPARATOR: "$",
  QUERY_STRING_INTRAFILTER_SEPARATOR: ":",
  QUERY_STRING_FILTER_VARIABLE: "x",
  DEFAULT_FONT_FAMILY: "Roboto Slab", // Set to undefined to copy font family from Spreadsheet (doesn't always work properly)
  DEFAULT_FONT_SIZE: 9 // Set to undefined to copy font size from Spreadsheet (doesn't always work properly)
}


function getTabs(doc) {
  return getNestedTabs(doc.getTabs());
}


function getNestedTabs(tabs) {
  if (tabs) {
    for (let t of tabs) {
      tabs = tabs.concat(getNestedTabs(t.getChildTabs()));
    }
  }
  return tabs;
}


function getElementBodyIndex(body, element) {
  let index;
  if (element.getParent().getParent().getType() === DocumentApp.ElementType.TABLE_CELL) { // Pre-inserted table
    index = body.getChildIndex(element.getParent().getParent().getParent().getParent()); // TEXT > PARAGRAPH > TABLE_CELL > TABLE_ROW > TABLE
  } else { // Standalone paragraph
    index = body.getChildIndex(element.getParent()); // TEXT > PARAGRAPH
  }
  return index;
}


function splitOnce(str, separator) {
  const [first, ...second] = str.split(separator);
  return [first, second.join(separator)];
}


function parseSpreadsheetLocator(spreadsheetLocator) {
  const [atSpreadsheetId, queryString] = splitOnce(spreadsheetLocator, CONFIG.QUERY_STRING_LEADING_CHARACTER);
  const {sheet, range, filters, fontFamily, fontSize} = parseQueryString(queryString);
  return {
    "id": atSpreadsheetId.slice(1), 
    "sheetId": sheet, 
    "range": range,
    "filters": filters,
    "fontFamily": fontFamily,
    "fontSize": fontSize
  };
}


function parseQueryString(queryString) {
  if (queryString) return queryString.split(CONFIG.QUERY_STRING_PARAMETER_SEPARATOR).reduce((acc, p) => {
    const [name, value] = splitOnce(p, CONFIG.QUERY_STRING_PARAMETER_ASSIGNMENT_OPERATOR);
    acc[name] = value;
    return acc;
  }, {});
}


function parseFilters(filters, columns, startCol) {
  if (filters) return filters.split(CONFIG.QUERY_STRING_INTERFILTER_SEPARATOR).reduce((acc, f) => {
    const [column, filter] = splitOnce(f, CONFIG.QUERY_STRING_INTRAFILTER_SEPARATOR);
    const i = /^[A-Z]+$/.test(column) ? getColumnIndex(column) : columns.indexOf(column) + startCol;
    acc[i] = eval(`(${CONFIG.QUERY_STRING_FILTER_VARIABLE}) => ${filter}`);
    return acc;
  }, {});
}


function getColumns(sheet, startRow, startCol) {
  const columns = getNonEmptySubarray(sheet.getDataRange().getDisplayValues()[startRow], startCol);
  return {
    "columns": columns.array,
    "startCol": columns.startIndex
  };
}


function getNonEmptySubarray(arr, startIndex) {
  let endIndex = startIndex;
  while (startIndex > 0 && arr[startIndex - 1] !== "") {
    startIndex--;
  }
  while (endIndex < arr.length && arr[endIndex] !== "") {
    endIndex++;
  }
  return {
    "array": arr.slice(startIndex, endIndex),
    "startIndex": startIndex
  };
}


function getColumnIndex(column) {
  return Array.from(column).reverse().reduce((acc, c, i) => acc + (c.charCodeAt(0) - 64)*26**i, 0) - 1; // Man, I love fbase26
}


function getExcludedRows(sheet, spreadsheetData, filterMap) {
  let excludedRows = [];
  const startRow = spreadsheetData.getRow() - 1;
  const numRows = spreadsheetData.getNumRows();
  const values = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < numRows; i++) {
    let included = true;
    for (let column in filterMap) {
      included &= filterMap[column](values[i + startRow][column]);
    }
    if (!included) excludedRows.push(i);
  }
  return excludedRows;
}


function getSmallestDataRange(sheet) {
  const defaultDataRangeValues = sheet.getDataRange().getDisplayValues();
  let startRow = 0;
  const numRows = defaultDataRangeValues.length;
  while (startRow < numRows && defaultDataRangeValues[startRow].every(item => item === "")) {
    startRow++;
  }
  let startCol = 0;
  const numCols = defaultDataRangeValues[0].length;
  while (startCol < numCols && defaultDataRangeValues.every(row => row[startCol] === "")) {
    startCol++;
  }
  return sheet.getRange(startRow + 1, startCol + 1, numRows - startRow, numCols - startCol);
}


function getIntersectionRange(sheet, r1, r2) {
  const startRow = Math.max(r1.getRow(), r2.getRow());
  const startCol = Math.max(r1.getColumn(), r2.getColumn());
  const lastRow = Math.min(r1.getLastRow(), r2.getLastRow());
  const lastCol = Math.min(r1.getLastColumn(), r2.getLastColumn());
  return sheet.getRange(startRow, startCol, lastRow - startRow + 1, lastCol - startCol + 1);
}


function getSpreadsheetData(spreadsheetLocator) {
  const {id, sheetId, range, filters, fontFamily, fontSize} = parseSpreadsheetLocator(spreadsheetLocator);
  const spreadsheet = SpreadsheetApp.openById(id);
  const sheet = sheetId !== undefined ? spreadsheet.getSheetById(sheetId) : spreadsheet.getSheets()[0];
  let spreadsheetData = getSmallestDataRange(sheet);
  if (range) {
    spreadsheetData = getIntersectionRange(sheet, spreadsheetData, sheet.getRange(range));
  }
  const {columns, startCol} = getColumns(sheet, spreadsheetData.getRow() - 1, spreadsheetData.getColumn() - 1);
  return {
    "spreadsheetData": spreadsheetData, 
    "spreadsheetUrl": `${spreadsheet.getUrl()}?gid=${sheet.getSheetId()}`,
    "excludedRows": getExcludedRows(sheet, spreadsheetData, parseFilters(filters, columns, startCol)),
    "fontFamily": fontFamily,
    "fontSize": fontSize
  };
}


function createTable(body, index, spreadsheetLocator) {
  const {spreadsheetData, spreadsheetUrl, excludedRows, fontFamily, fontSize} = getSpreadsheetData(spreadsheetLocator);
  const table = body.insertTable(index, spreadsheetData.getDisplayValues());
  applySpreadsheetStyle(
    spreadsheetData, 
    table, 
    fontFamily ? fontFamily : CONFIG.DEFAULT_FONT_FAMILY, 
    fontSize ? fontSize : CONFIG.DEFAULT_FONT_SIZE
  );
  // applySpreadsheetMerge(spreadsheetData, table);
  removeRows(table, excludedRows);
  addSpreadsheetLocator(table, spreadsheetLocator, spreadsheetUrl);
}


function applySpreadsheetStyle(data, table, fontFamily, fontSize) {
  const fontSizes = data.getFontSizes();
  const fontFamilies = data.getFontFamilies();
  const fontWeights = data.getFontWeights();
  const backgroundColors = data.getBackgrounds();
  const fontColors = data.getFontColors();
  const richTextValues = data.getRichTextValues();
  const rows = data.getNumRows();
  const cols = data.getNumColumns();
  for (let i = 0; i < rows; i++) {
    const row = table.getRow(i);
    for (let j = 0; j < cols; j++) {
      const cell = row.getCell(j);
      const text = cell.getChild(0);
      const cellRichText = richTextValues[i][j];
      if (cellRichText) {
        cell.setLinkUrl(cellRichText.getLinkUrl());
      }
      text.setBold(fontWeights[i][j] === "bold");
      cell.setBackgroundColor(backgroundColors[i][j]);
      cell.setForegroundColor(fontColors[i][j]);
      fontFamily ? cell.setFontFamily(fontFamily) : cell.setFontFamily(fontFamilies[i][j]);
      fontSize ? cell.setFontSize(fontSize) : cell.setFontSize(fontSizes[i][j]);
    }
  }
}


function applySpreadsheetMerge(data, table) {
  const mergedRanges = data.getMergedRanges();
  for (let mergedRange of mergedRanges) {
    const startRow = mergedRange.getRow() - 1;
    const startCol = mergedRange.getColumn() - 1;
    const numRows = mergedRange.getNumRows();
    const numCols = mergedRange.getNumColumns();
    
    // As of last commit the Apps Script Docs API provides no methods for merging cells both horizontally and vertically
    // There are also no methods to customize individual cell borders to give the impression of cells being merged
    // As a result, it is currently impossible to merge an arbitrary range of cells using Apps Script alone :'( (I'm crying)
  }
}


function removeRows(table, rows) {
  for (let row of rows.sort((a, b) => b - a)) {
    table.removeRow(row);
  }
}


function addSpreadsheetLocator(table, spreadsheetLocator, spreadsheetUrl, fontFamily="Courier New", fontSize=1) {
  const header = table.insertTableRow(0);
  const cell = header.insertTableCell(0);
  cell.editAsText().setText(spreadsheetLocator);
  cell.setLinkUrl(spreadsheetUrl);
  cell.setFontFamily(fontFamily);
  cell.setFontSize(fontSize);
}


function upsert(doc) {
  const spreadsheetRegExp = `^${CONFIG.SPREADSHEET_LOCATOR_LEADING_CHARACTER}([a-zA-Z0-9-_^~]{44}).*`;
  const tabs = getTabs(doc);
  for (let t of tabs) {
    const body = t.asDocumentTab().getBody();
    let match = body.findText(spreadsheetRegExp);
    while (match) {
      const element = match.getElement();
      // Get the matched element's index in the current tab's body
      const matchIndex = getElementBodyIndex(body, element);
      // Get sheet locator (@<spreadsheet UID><querystring>)
      const spreadsheetLocator = element.asText().getText();
      // Create and insert the updated table at the matched element's index
      createTable(body, matchIndex, spreadsheetLocator);
      // Look for next match in the current tab
      match = body.findText(spreadsheetRegExp, match);
      // Remove the matched element, now found at the following index
      if (matchIndex + 1 < body.getNumChildren() - 1) {
        body.removeChild(body.getChild(matchIndex + 1));
      } else {
        element.setText(""); // Delete paragraph content
      }
    }
  }
}
