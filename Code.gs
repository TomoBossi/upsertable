const CONFIG = {
  SPREADSHEET_LOCATOR_LEADING_CHARACTER: "\\@", // Because it will be used inside a RegExp, the leading character may need to be escaped (\\)
  QUERY_STRING_LEADING_CHARACTER: "?",
  QUERY_STRING_PARAMETER_SEPARATOR: "%",
  QUERY_STRING_PARAMETER_ASSIGNMENT_OPERATOR: "=",
  QUERY_STRING_INTERFILTER_SEPARATOR: "$",
  QUERY_STRING_INTRAFILTER_SEPARATOR: ":",
  QUERY_STRING_FILTER_VARIABLE: "x",
  DEFAULT_FONT_FAMILY: "Roboto Slab", // Set to undefined to copy font family from Spreadsheet (doesn't always work properly)
  DEFAULT_FONT_SIZE: 10 // Set to undefined to copy font size from Spreadsheet (doesn't always work properly)
}


function getTabs(doc) {
  const unnestedTabs = doc.getTabs();
  return getNestedTabs(unnestedTabs);
}


function getNestedTabs(unnestedTabs) {
  if (!unnestedTabs) {
    return [];
  }
  let tabs = unnestedTabs;
  for (let t of unnestedTabs) {
    tabs = tabs.concat(getNestedTabs(t.getChildTabs()));
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


function parseSpreadsheetLocator(spreadsheetLocator) {
  const [atSpreadsheetId, ...queryString] = spreadsheetLocator.split(CONFIG.QUERY_STRING_LEADING_CHARACTER);
  const {sheet, range, filters, fontFamily, fontSize} = parseQueryString(queryString.join(CONFIG.QUERY_STRING_LEADING_CHARACTER));
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
    const [name, ...value] = p.split(CONFIG.QUERY_STRING_PARAMETER_ASSIGNMENT_OPERATOR);
    acc[name] = value.join(CONFIG.QUERY_STRING_PARAMETER_ASSIGNMENT_OPERATOR);
    return acc;
  }, {});
}


function parseFilters(filters, columns) {
  if (filters) return filters.split(CONFIG.QUERY_STRING_INTERFILTER_SEPARATOR).reduce((acc, f) => {
    const [column, ...filter] = f.split(CONFIG.QUERY_STRING_INTRAFILTER_SEPARATOR);
    const i = /^[A-Z]+$/.test(column) ? getColumnIndex(column) : columns.indexOf(column);
    acc[i] = eval(`(${CONFIG.QUERY_STRING_FILTER_VARIABLE}) => ${filter.join(CONFIG.QUERY_STRING_INTRAFILTER_SEPARATOR)}`);
    return acc;
  }, {});
}


function getColumns(sheet) {
  return sheet.getDataRange().getDisplayValues()[0];
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


function getSpreadsheetData(spreadsheetLocator) {
  const {id, sheetId, range, filters, fontFamily, fontSize} = parseSpreadsheetLocator(spreadsheetLocator);
  const spreadsheet = SpreadsheetApp.openById(id);
  const sheet = sheetId !== undefined ? spreadsheet.getSheetById(sheetId) : spreadsheet.getSheets()[0];
  let spreadsheetData = sheet.getDataRange();
  if (range) {
    const selection = sheet.getRange(range);
    const rangeFirstRow = Math.max(spreadsheetData.getRow(), selection.getRow());
    const rangeFirstCol = Math.max(spreadsheetData.getColumn(), selection.getColumn());
    const rangeLastRow = Math.min(spreadsheetData.getLastRow(), selection.getLastRow());
    const rangeLastCol = Math.min(spreadsheetData.getLastColumn(), selection.getLastColumn());
    spreadsheetData = sheet.getRange(rangeFirstRow, rangeFirstCol, rangeLastRow - rangeFirstRow + 1, rangeLastCol - rangeFirstCol + 1);
  }
  return {
    "spreadsheetData": spreadsheetData, 
    "spreadsheetUrl": `${spreadsheet.getUrl()}?gid=${sheet.getSheetId()}`,
    "excludedRows": getExcludedRows(sheet, spreadsheetData, parseFilters(filters, getColumns(sheet))),
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
