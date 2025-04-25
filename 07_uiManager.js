/**
 * UI Manager Module for MOAA Column Selector
 *
 * Purpose: Handles user interface interactions and sheet UI updates
 *
 * Dependencies:
 * - DataAccess: For spreadsheet operations
 * - ConfigManager: For managing column selections
 * - WorksheetService: For data operations
 *
 * Dependent Modules:
 * - Global UI functions: For menu operations
 * - main.js: For onOpen handler
 *
 * Initialization Order: Should be loaded after all other modules
 *
 * Note: Provides global functions for UI menu callbacks
 */
var UIManager = (function () {
  // Constants for UI elements
  const CONSTANTS = {
    MENU_NAME: "MOAA Chart Tools",
    SELECTOR_SHEET_NAME: "Column Selector",
    SELECTOR_HEADER_COLS: [
      "Worksheet Name",
      "Column Name",
      "Include in Chart",
      "Chart Group",
    ],
    COLUMN_GROUPS: ["Group 1", "Group 2", "Group 3", "Group 4", "Group 5"],
    BUTTON_TYPES: {
      OK: "OK",
      YES_NO: "YES_NO",
      YES_NO_CANCEL: "YES_NO_CANCEL",
      OK_CANCEL: "OK_CANCEL",
    },
    // Alert types for consistent messaging
    ALERT_TYPES: {
      INFO: "INFO",
      WARNING: "WARNING",
      ERROR: "ERROR",
      SUCCESS: "SUCCESS",
    },
  };

  /**
   * Creates the custom menu in the spreadsheet
   */
  function createCustomMenu() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu(CONSTANTS.MENU_NAME)
      .addItem("Generate Column Selector", "generateColumnSelector")
      .addItem("Apply Column Selections", "applyColumnSelections")
      .addSeparator()
      .addItem("Check Column Changes", "showColumnChanges")
      .addToUi();
  }

  function isButtonYes(response) {
    const ui = SpreadsheetApp.getUi();
    return response === ui.Button.YES;
  }

  /**
   * Gets the appropriate ButtonSet based on string type
   * @param {string} buttonType - Type from CONSTANTS.BUTTON_TYPES
   * @return {GoogleAppsScript.Base.ButtonSet} The ButtonSet object
   */
  function getButtonSet(buttonType) {
    const ui = SpreadsheetApp.getUi();
    switch (buttonType) {
      case CONSTANTS.BUTTON_TYPES.YES_NO:
        return ui.ButtonSet.YES_NO;
      case CONSTANTS.BUTTON_TYPES.YES_NO_CANCEL:
        return ui.ButtonSet.YES_NO_CANCEL;
      case CONSTANTS.BUTTON_TYPES.OK_CANCEL:
        return ui.ButtonSet.OK_CANCEL;
      case CONSTANTS.BUTTON_TYPES.OK:
      default:
        return ui.ButtonSet.OK;
    }
  }

  /**
   * Shows UI alert with custom message
   * @param {string} title - Alert title
   * @param {string} message - Alert message
   * @param {string} [buttonType=CONSTANTS.BUTTON_TYPES.OK] - Button type from CONSTANTS.BUTTON_TYPES
   * @param {string} [alertType=CONSTANTS.ALERT_TYPES.INFO] - Alert type from CONSTANTS.ALERT_TYPES
   * @return {GoogleAppsScript.Base.Button | null} Button that was clicked
   */
  function showAlert(
    title,
    message,
    buttonType = CONSTANTS.BUTTON_TYPES.OK,
    alertType = CONSTANTS.ALERT_TYPES.INFO
  ) {
    try {
      const ui = SpreadsheetApp.getUi();
      const buttonSet = getButtonSet(buttonType);
      // Optionally format message based on alertType
      let formattedMessage = message;
      return ui.alert(title, formattedMessage, buttonSet);
    } catch (error) {
      console.error("Error in showAlert:", error);
      return null;
    }
  }

  /**
   * Shows dialog to confirm column changes and optionally regenerate selector
   * @param {string[]} changedWorksheets - List of worksheets with changes
   * @return {boolean} True if user chose to regenerate
   */
  function confirmColumnChanges(changedWorksheets) {
    if (changedWorksheets.length === 0) {
      showAlert(
        "No Column Changes",
        "All worksheets match their stored signatures. The Column Selector is up to date.",
        CONSTANTS.BUTTON_TYPES.OK
      );
      return false;
    }

    const response = showAlert(
      "Column Changes Detected",
      `The following worksheets have column changes:\n\n${changedWorksheets.join(
        ", "
      )}\n\nWould you like to regenerate the Column Selector?`,
      CONSTANTS.BUTTON_TYPES.YES_NO
    );

    return isButtonYes(response);
  }

  /**
   * Creates or clears the Column Selector sheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet
   * @return {GoogleAppsScript.Spreadsheet.Sheet} The selector sheet
   */
  function prepareColumnSelectorSheet(ss) {
    let selectorSheet = ss.getSheetByName(CONSTANTS.SELECTOR_SHEET_NAME);
    if (!selectorSheet) {
      selectorSheet = ss.insertSheet(CONSTANTS.SELECTOR_SHEET_NAME);
    } else {
      selectorSheet.clear();
    }
    return selectorSheet;
  }

  /**
   * Applies formatting to the Column Selector sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} selectorSheet - The selector sheet
   * @param {number} dataRowCount - Number of data rows
   */
  function formatColumnSelector(selectorSheet, dataRowCount) {
    // Format header row
    selectorSheet.getRange("A2:D2").setFontWeight("bold");

    // Add header message
    ConfigManager.setColumnSelectorHeader(selectorSheet, "generated");

    // Freeze header rows
    selectorSheet.setFrozenRows(2);

    // Resize columns for better readability
    selectorSheet.autoResizeColumns(1, 4);
  }

  /**
   * Applies data validation to the Chart Group column
   * @param {GoogleAppsScript.Spreadsheet.Sheet} selectorSheet - The selector sheet
   * @param {number[]} validationRanges - Row indices for validation
   */
  function applyGroupValidation(selectorSheet, validationRanges) {
    if (validationRanges.length === 0) return;

    const groups = CONSTANTS.COLUMN_GROUPS;
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(groups, true)
      .build();

    // Apply validation in batches of continuous ranges
    let currentStart = validationRanges[0];
    let currentEnd = validationRanges[0];

    for (let i = 1; i <= validationRanges.length; i++) {
      // If we're at the end or found a discontinuity
      if (
        i === validationRanges.length ||
        validationRanges[i] !== currentEnd + 1
      ) {
        // Apply validation to the continuous range
        selectorSheet
          .getRange(currentStart, 4, currentEnd - currentStart + 1, 1)
          .setDataValidation(validation);

        // Start a new range if we're not at the end
        if (i < validationRanges.length) {
          currentStart = validationRanges[i];
          currentEnd = validationRanges[i];
        }
      } else {
        // Continue the current range
        currentEnd = validationRanges[i];
      }
    }
  }

  /**
   * Applies checkboxes to the Include in Chart column
   * @param {GoogleAppsScript.Spreadsheet.Sheet} selectorSheet - The selector sheet
   * @param {number[]} checkboxRanges - Row indices for checkboxes
   */
  function applyCheckboxes(selectorSheet, checkboxRanges) {
    if (checkboxRanges.length === 0) return;

    // Find continuous ranges to minimize API calls
    let currentStart = checkboxRanges[0];
    let currentEnd = checkboxRanges[0];

    for (let i = 1; i <= checkboxRanges.length; i++) {
      // If we're at the end or found a discontinuity
      if (i === checkboxRanges.length || checkboxRanges[i] !== currentEnd + 1) {
        // Apply checkboxes to the continuous range
        selectorSheet
          .getRange(currentStart, 3, currentEnd - currentStart + 1, 1)
          .insertCheckboxes();

        // Start a new range if we're not at the end
        if (i < checkboxRanges.length) {
          currentStart = checkboxRanges[i];
          currentEnd = checkboxRanges[i];
        }
      } else {
        // Continue the current range
        currentEnd = checkboxRanges[i];
      }
    }
  }

  /**
   * Generates the Column Selector sheet
   */
  function generateColumnSelector() {
    // Capture existing selections before regenerating
    const savedSelections = ConfigManager.captureColumnSelections();

    const ss = DataAccess.getActiveSpreadsheet();
    const worksheetData = DataAccess.getWorksheetsList(ss);

    // Create or clear the Column Selector sheet
    const selectorSheet = prepareColumnSelectorSheet(ss);

    // Prepare all data in memory first
    const allData = [];
    const checkboxRanges = [];
    const validationRanges = [];

    // Add blank row for instructions
    allData.push(["", "", "", ""]);

    // Add header row
    allData.push(CONSTANTS.SELECTOR_HEADER_COLS);

    let dataRowIndex = 2; // Starting index for data rows (after header)

    // Process each worksheet
    worksheetData.forEach(function (worksheet) {
      const worksheetName = worksheet.worksheetName;
      const sheet = DataAccess.getSheetByName(worksheetName, ss);
      if (!sheet) return;

      // Get column names
      const columnNames = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0];

      let addedRowsForWorksheet = false;

      // Add each column as a row in the selector
      columnNames.forEach(function (columnName, index) {
        if (columnName && index > 0) {
          // Skip first column assuming it's ID
          // Add data row
          allData.push([worksheetName, columnName, "", ""]);

          // Track checkbox and validation ranges
          checkboxRanges.push(dataRowIndex + 2); // +2 for instruction and header
          validationRanges.push(dataRowIndex + 2);

          dataRowIndex++;
          addedRowsForWorksheet = true;
        }
      });

      // Add blank row between worksheets for readability
      if (addedRowsForWorksheet) {
        allData.push(["", "", "", ""]);
        dataRowIndex++;
      }
    });

    // Write all data at once - more efficient than individual writes
    if (allData.length > 0) {
      selectorSheet.getRange(1, 1, allData.length, 4).setValues(allData);
    }

    // Apply checkboxes and validation
    applyCheckboxes(selectorSheet, checkboxRanges);
    applyGroupValidation(selectorSheet, validationRanges);

    // Format the sheet
    formatColumnSelector(selectorSheet, allData.length);

    // Store signatures and restore selections
    ConfigManager.storeColumnSignatures(ss);
    ConfigManager.restoreColumnSelections(savedSelections);

    // Update header timestamp
    ConfigManager.setColumnSelectorHeader(selectorSheet, "updated");
  }

  /**
   * Shows UI for column changes detection
   */
  function showColumnChanges() {
    const ss = DataAccess.getActiveSpreadsheet();
    const changedWorksheets = ConfigManager.detectColumnChanges(ss);

    const shouldRegenerate = confirmColumnChanges(changedWorksheets);

    if (shouldRegenerate) {
      generateColumnSelector();
      showAlert(
        "Column Selector Updated",
        "The Column Selector has been regenerated with the latest column structures.",
        CONSTANTS.BUTTON_TYPES.OK
      );
    }
  }

  /**
   * Updates the Table Settings worksheet with column selections and clears values for worksheets with no selections.
   * This function takes the user's selections from the Column Selector sheet and applies them to the Table Settings
   * worksheet, updating chart columns, chart groups, and chart types. It efficiently handles batch updates and
   * ensures that worksheets with no selected columns have their values cleared.
   *
   * @param {GoogleAppsScript.Spreadsheet.Sheet} tableSettingsSheet - The Table Settings worksheet to be updated
   * @param {Object.<string, Object.<string, string[]>>} worksheetSelections - Object containing selected columns by worksheet and group
   *    Format: {
   *      "WorksheetName1": {
   *        "Group1": ["Column1", "Column2"],
   *        "Group2": ["Column3"]
   *      },
   *      "WorksheetName2": {
   *        "Group1": ["Column1"]
   *      }
   *    }
   * @returns {void}
   *
   * @example
   * const tableSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Table Settings");
   * const selections = {
   *   "Demographics": {
   *     "Group 1": ["Annual Revenue", "Number of Leadership Team"]
   *   },
   *   "Current & Needed Integrator": {
   *     "Group 1": ["Currently have Operational Integrator", "Currently have Conductor Integrator"],
   *     "Group 2": ["Need Conductor Integrator", "Need Executive Integrator"]
   *   }
   * };
   * updateTableSettings(tableSettings, selections);
   */
  function updateTableSettings(tableSettingsSheet, worksheetSelections) {
    const lastRow = tableSettingsSheet.getLastRow();
    const lastColumn = tableSettingsSheet.getLastColumn();
    const data = tableSettingsSheet
      .getRange(1, 1, lastRow, lastColumn)
      .getValues();
    const headers = data[0];

    // Find column indexes
    const worksheetNameColIdx = headers.indexOf(
      DataAccess.CONSTANTS.COLUMNS.WORKSHEET_NAME
    );
    const chartColumnsColIdx = headers.indexOf(
      DataAccess.CONSTANTS.COLUMNS.CHART_COLUMNS
    );
    const chartGroupsColIdx = headers.indexOf(
      DataAccess.CONSTANTS.COLUMNS.CHART_GROUPS
    );
    const chartTypeColIdx = headers.indexOf(
      DataAccess.CONSTANTS.COLUMNS.CHART_TYPE
    );

    // Ensure required columns exist
    if (
      worksheetNameColIdx === -1 ||
      chartColumnsColIdx === -1 ||
      chartGroupsColIdx === -1
    ) {
      showAlert(
        "Missing Columns",
        "Required columns not found in Table Settings sheet.",
        CONSTANTS.BUTTON_TYPES.OK
      );
      return;
    }

    // Prepare arrays for batch updates
    const columnsValues = [];
    const groupsValues = [];
    const typeValues = [];

    for (let i = 1; i < data.length; i++) {
      const worksheetName = data[i][worksheetNameColIdx];

      if (!worksheetName) continue;

      // Default to empty values
      let columnsValue = "";
      let groupsValue = "";
      let typeValue =
        chartTypeColIdx !== -1 ? data[i][chartTypeColIdx] || "" : "";

      // If this worksheet has selections, update the values
      if (worksheetSelections[worksheetName]) {
        const selections = worksheetSelections[worksheetName];
        const groups = Object.keys(selections);

        if (groups.length > 0) {
          // Process selections
          const allColumns = [];
          groups.forEach(function (group) {
            selections[group].forEach(function (column) {
              allColumns.push(column);
            });
          });

          columnsValue = JSON.stringify(allColumns);
          groupsValue = JSON.stringify(selections);

          // Set default chart type if needed
          if (chartTypeColIdx !== -1 && !typeValue) {
            typeValue =
              groups.length > 1
                ? ChartGenerator.CONSTANTS.CHART_TYPES.PIE
                : ChartGenerator.CONSTANTS.CHART_TYPES.BAR;
          }
        }
      }

      columnsValues.push([columnsValue]);
      groupsValues.push([groupsValue]);
      if (chartTypeColIdx !== -1) {
        typeValues.push([typeValue]);
      }
    }

    // Apply batch updates
    if (columnsValues.length > 0) {
      tableSettingsSheet
        .getRange(2, chartColumnsColIdx + 1, columnsValues.length, 1)
        .setValues(columnsValues);

      tableSettingsSheet
        .getRange(2, chartGroupsColIdx + 1, groupsValues.length, 1)
        .setValues(groupsValues);

      if (chartTypeColIdx !== -1) {
        tableSettingsSheet
          .getRange(2, chartTypeColIdx + 1, typeValues.length, 1)
          .setValues(typeValues);
      }
    }
  }

  /**
   * Applies column selections to the Table Settings sheet
   */
  function applyColumnSelections() {
    const ss = DataAccess.getActiveSpreadsheet();
    const selectorSheet = DataAccess.getSheetByName(
      CONSTANTS.SELECTOR_SHEET_NAME,
      ss
    );
    const tableSettings = DataAccess.getSheetByName(
      DataAccess.CONSTANTS.SHEETS.TABLE_SETTINGS,
      ss
    );

    if (!selectorSheet || !tableSettings) {
      showAlert(
        "Missing Sheets",
        "Column Selector or Table Settings sheet not found",
        CONSTANTS.BUTTON_TYPES.OK
      );
      return;
    }

    // Check for column changes
    const changedWorksheets = ConfigManager.detectColumnChanges(ss);

    if (changedWorksheets.length > 0) {
      const response = showAlert(
        "Column Changes Detected",
        `The following worksheets have column changes since the selector was generated:\n\n${changedWorksheets.join(
          ", "
        )}\n\nWould you like to regenerate the Column Selector before applying changes?`,
        CONSTANTS.BUTTON_TYPES.YES_NO
      );

      if (isButtonYes(response)) {
        generateColumnSelector();
        showAlert(
          "Column Selector Updated",
          "Please review and select columns again before applying changes.",
          CONSTANTS.BUTTON_TYPES.OK
        );
        return;
      }
    }

    // Get all data from the selector sheet
    const lastRow = selectorSheet.getLastRow();
    const selectorData = selectorSheet
      .getRange(2, 1, lastRow - 1, 4)
      .getValues();

    // Organize selections by worksheet
    const worksheetSelections = {};

    selectorData.forEach(function (row) {
      const worksheetName = row[0];
      const columnName = row[1];
      const isIncluded = row[2]; // Boolean for checkbox
      const chartGroup = row[3];

      // Skip headers, empty rows, or unselected columns
      if (!worksheetName || !columnName || !isIncluded) return;

      // Initialize worksheet entry if needed
      if (!worksheetSelections[worksheetName]) {
        worksheetSelections[worksheetName] = {};
      }

      // Initialize group if needed
      if (!worksheetSelections[worksheetName][chartGroup]) {
        worksheetSelections[worksheetName][chartGroup] = [];
      }

      // Add column to group
      worksheetSelections[worksheetName][chartGroup].push(columnName);
    });

    // Update Table Settings with selections
    updateTableSettings(tableSettings, worksheetSelections);

    // Store updated column signatures
    ConfigManager.storeColumnSignatures(ss);

    // Show success message
    showAlert(
      "Success",
      "Column selections applied to Table Settings",
      CONSTANTS.BUTTON_TYPES.OK
    );
  }

  // Public API
  return {
    CONSTANTS: CONSTANTS,
    createCustomMenu: createCustomMenu,
    showAlert: showAlert,
    generateColumnSelector: generateColumnSelector,
    applyColumnSelections: applyColumnSelections,
    showColumnChanges: showColumnChanges,
    updateTableSettings: updateTableSettings,
  };
})();

// Expose functions directly or with a factory
function getGlobalFunctions() {
  return {
    generateColumnSelector: UIManager.generateColumnSelector,
    applyColumnSelections: UIManager.applyColumnSelections,
    showColumnChanges: UIManager.showColumnChanges,
  };
}

// Then use reflection to create globals
Object.assign(this, getGlobalFunctions());
