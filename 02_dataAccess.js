/**
 * Data Access Module for MOAA Chart Tools
 *
 * Purpose: Handles all direct interactions with Google Sheets
 *
 * Dependencies:
 * - Utils: For data formatting and error handling
 *
 * Dependent Modules:
 * - ConfigManager: For accessing spreadsheet data
 * - WorksheetService: For data retrieval
 * - UIManager: For accessing spreadsheet data
 *
 * Initialization Order: Should be loaded after Utils
 */
var DataAccess = (function () {
  // Constants for common sheet and column names
  const CONSTANTS = {
    SHEETS: {
      WORKSHOP_PLANNER: "Workshop Planner",
      TABLE_SETTINGS: "Table Settings",
      PARTNER_LIST: "Partner List",
      COLUMN_SELECTOR: "Column Selector",
    },
    COLUMNS: {
      WORKSHEET_NAME: "Worksheet Name",
      TYPE: "Type",
      CHART_TYPE: "Chart Type",
      CHART_COLUMNS: "Chart Columns",
      CHART_GROUPS: "Chart Groups",
      WORKSHOP_ID: "ID",
      WORKSHOP_DISPLAY_NAME: "Display Name",
      PARTNER_ID: "ID",
      PARTNER_DISPLAY_NAME: "Full Name",
    },
  };

  /**
   * Gets current active spreadsheet
   * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  function getActiveSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  /**
   * Gets a sheet by name
   * @param {string} sheetName - Name of the sheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Spreadsheet (uses active if not provided)
   * @return {GoogleAppsScript.Spreadsheet.Sheet|null} Sheet object or null if not found
   */
  function getSheetByName(sheetName, ss) {
    if (!ss) ss = getActiveSpreadsheet();
    return ss.getSheetByName(sheetName);
  }

  /**
   * Gets full data range from a sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to get data from
   * @param {boolean} [includeHeaders=true] - Whether to include headers in the result
   * @return {SheetData} Object with headers array and data array
   * @typedef {Object} SheetData
   * @property {string[]} headers - Array of header values
   * @property {Array<Array>} data - 2D array of data values
   */
  function getSheetData(sheet, includeHeaders = false) {
    const numColumns = sheet.getLastColumn();
    const numRows = Utils.getLastRowWithContent(sheet, 1);

    if (numRows <= (includeHeaders ? 1 : 0) || numColumns === 0) {
      return { headers: [], data: [] };
    }

    const range = sheet.getRange(1, 1, numRows, numColumns);
    const values = range.getValues();

    const headers = values[0];
    const data = includeHeaders ? values : values.slice(1);

    return { headers, data };
  }

  /**
   * Gets list of worksheet configurations from Table Settings
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Spreadsheet (uses active if not provided)
   * @return {Array<WorksheetConfig>} Array of worksheet configuration objects
   * @typedef {Object} WorksheetConfig
   * @property {string} worksheetName - Name of the worksheet
   * @property {string} type - Type of worksheet (static, dynamic)
   * @property {string} [chartType] - Type of chart to generate
   * @property {string} [chartColumns] - Columns to use for chart data
   * @property {string} [chartGroups] - Columns to use for chart groups
   */
  function getWorksheetsList(ss) {
    if (!ss) ss = getActiveSpreadsheet();
    const tableSettingsSheet = getSheetByName(
      CONSTANTS.SHEETS.TABLE_SETTINGS,
      ss
    );

    if (!tableSettingsSheet) {
      console.error("Table Settings sheet not found");
      return [];
    }

    const { headers, data } = getSheetData(tableSettingsSheet);

    // Find column indices
    const worksheetNameIdx = headers.indexOf(CONSTANTS.COLUMNS.WORKSHEET_NAME);
    const typeIdx = headers.indexOf(CONSTANTS.COLUMNS.TYPE);
    const chartTypeIdx = headers.indexOf(CONSTANTS.COLUMNS.CHART_TYPE);
    const chartColumnsIdx = headers.indexOf(CONSTANTS.COLUMNS.CHART_COLUMNS);
    const chartGroupsIdx = headers.indexOf(CONSTANTS.COLUMNS.CHART_GROUPS);

    if (worksheetNameIdx === -1 || typeIdx === -1) {
      console.error("Required columns not found in Table Settings");
      return [];
    }

    const worksheets = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[worksheetNameIdx]) continue;

      const worksheet = {
        worksheetName: row[worksheetNameIdx],
        type: row[typeIdx],
      };

      // Add optional properties if columns exist
      if (chartTypeIdx !== -1) worksheet.chartType = row[chartTypeIdx];
      if (chartColumnsIdx !== -1) worksheet.chartColumns = row[chartColumnsIdx];
      if (chartGroupsIdx !== -1) worksheet.chartGroups = row[chartGroupsIdx];

      worksheets.push(worksheet);
    }

    return worksheets;
  }

  /**
   * Gets list of workshops from Workshop Planner
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Spreadsheet (uses active if not provided)
   * @return {Array<Object>} Array of workshop objects with value and label properties
   */
  function getWorkshopList(ss) {
    if (!ss) ss = getActiveSpreadsheet();
    const workshopPlannerSheet = getSheetByName(
      CONSTANTS.SHEETS.WORKSHOP_PLANNER,
      ss
    );

    if (!workshopPlannerSheet) {
      console.error("Workshop Planner sheet not found");
      return [];
    }

    const { headers, data } = getSheetData(workshopPlannerSheet);

    // Find column indices
    const idIdx = headers.indexOf(CONSTANTS.COLUMNS.WORKSHOP_ID);
    const displayNameIdx = headers.indexOf(
      CONSTANTS.COLUMNS.WORKSHOP_DISPLAY_NAME
    );

    if (idIdx === -1 || displayNameIdx === -1) {
      console.error("Required columns not found in Workshop Planner");
      return [];
    }

    const workshops = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[idIdx]) continue;

      workshops.push({
        value: row[idIdx],
        label: row[displayNameIdx],
      });
    }

    return workshops;
  }

  /**
   * Gets list of partners from Partner List
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Spreadsheet (uses active if not provided)
   * @return {Array<Object>} Array of partner objects with value and label properties
   */
  function getPartnerList(ss) {
    if (!ss) ss = getActiveSpreadsheet();
    const partnerListSheet = getSheetByName(CONSTANTS.SHEETS.PARTNER_LIST, ss);

    if (!partnerListSheet) {
      console.error("Partner List sheet not found");
      return [];
    }

    const { headers, data } = getSheetData(partnerListSheet);

    // Find column indices
    const idIdx = headers.indexOf(CONSTANTS.COLUMNS.PARTNER_ID);
    const displayNameIdx = headers.indexOf(
      CONSTANTS.COLUMNS.PARTNER_DISPLAY_NAME
    );

    if (idIdx === -1 || displayNameIdx === -1) {
      console.error("Required columns not found in Partner List");
      return [];
    }

    // Skip header row
    const partners = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[idIdx]) continue;

      partners.push({
        value: row[idIdx],
        label: row[displayNameIdx],
      });
    }

    return partners;
  }

  /**
   * Gets data from a specific worksheet, optionally filtered by ID
   * @param {string} worksheetName - Name of the worksheet
   * @param {string} [filterId] - Optional ID to filter by (first column)
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Spreadsheet (uses active if not provided)
   * @return {Object} Object with headers and filtered data
   */
  function getWorksheetData(worksheetName, filterId, ss) {
    if (!ss) ss = getActiveSpreadsheet();
    const sheet = getSheetByName(worksheetName, ss);

    if (!sheet) {
      console.error(`Worksheet ${worksheetName} not found`);
      return { headers: [], data: [] };
    }

    const { headers, data } = getSheetData(sheet);

    // If no filter ID, return all data
    if (!filterId) {
      return { headers, data };
    }

    // Filter data by ID (assuming ID is in first column)
    const filteredData = data.filter((row) => row[0] === filterId);
    return { headers, data: filteredData };
  }

  /**
   * Writes values to a range in a sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to write to
   * @param {number} startRow - Starting row (1-based)
   * @param {number} startCol - Starting column (1-based)
   * @param {Array<Array>} values - 2D array of values to write
   * @return {boolean} True if successful
   */
  function writeToRange(sheet, startRow, startCol, values) {
    if (!sheet || !values || values.length === 0) return false;

    try {
      const numRows = values.length;
      const numCols = values[0].length;
      sheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
      return true;
    } catch (e) {
      console.error(`Error writing to range: ${e.message}`);
      return false;
    }
  }

  // Public API
  return {
    CONSTANTS: CONSTANTS,
    getActiveSpreadsheet: getActiveSpreadsheet,
    getSheetByName: getSheetByName,
    getSheetData: getSheetData,
    getWorksheetsList: getWorksheetsList,
    getWorkshopList: getWorkshopList,
    getPartnerList: getPartnerList,
    getWorksheetData: getWorksheetData,
    writeToRange: writeToRange,
  };
})();
