/**
 * Configuration Manager Module for MOAA dashboard
 *
 * Purpose: Handles storage and retrieval of application state and configuration
 *
 * Dependencies:
 * - DataAccess: For accessing spreadsheet data
 *
 * Dependent Modules:
 * - UIManager: For managing column selections
 *
 * Initialization Order: Should be loaded after DataAccess
 */
var ConfigManager = (function () {
  // Constants for property keys and configuration
  const CONSTANTS = {
    PROPERTY_KEYS: {
      COLUMN_SIGNATURES: "columnSignatures",
      COLUMN_SELECTIONS: "columnSelections",
    },
    SELECTOR_SHEET: "Column Selector",
    SELECTOR_HEADER_RANGE: "A1:D1",
  };

  /**
   * Gets the document properties service
   * @return {GoogleAppsScript.Properties.Properties} Properties service
   */
  function getDocumentProperties() {
    return PropertiesService.getDocumentProperties();
  }

  /**
   * Generates a signature for a worksheet's columns
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The worksheet
   * @param {string[]} [columnNames] - Optional specific column names to include
   * @return {string} JSON string representing column structure
   */
  function generateColumnSignature(sheet, columnNames = []) {
    if (!sheet) return "{}";

    // If column names not provided, get them from the sheet
    if (columnNames.length === 0) {
      columnNames = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0];
    }

    // Filter out empty columns
    const signature = columnNames.filter((col) => col !== "");

    // Create a string representation
    return JSON.stringify(signature);
  }

  /**
   * Stores column signatures for all relevant worksheets
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet
   * @return {Object} The generated signatures object
   */
  function storeColumnSignatures(spreadsheet) {
    const worksheetData = DataAccess.getWorksheetsList(spreadsheet);
    const signatures = {};

    worksheetData.forEach(function (worksheetInfo) {
      const worksheetName = worksheetInfo.worksheetName;
      const ws = DataAccess.getSheetByName(worksheetName, spreadsheet);

      if (ws) {
        signatures[worksheetName] = generateColumnSignature(ws);
      }
    });

    // Store in document properties
    getDocumentProperties().setProperty(
      CONSTANTS.PROPERTY_KEYS.COLUMN_SIGNATURES,
      JSON.stringify(signatures)
    );

    return signatures;
  }

  /**
   * Retrieves stored column signatures
   * @return {Object} The stored signatures object or empty object if none
   */
  function getStoredColumnSignatures() {
    try {
      const signaturesJson = getDocumentProperties().getProperty(
        CONSTANTS.PROPERTY_KEYS.COLUMN_SIGNATURES
      );
      return signaturesJson ? JSON.parse(signaturesJson) : {};
    } catch (e) {
      console.error("Error retrieving column signatures:", e);
      return {};
    }
  }

  /**
   * Checks if worksheet columns have changed since signatures were last stored
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet
   * @return {string[]} Names of worksheets with column changes
   */
  function detectColumnChanges(spreadsheet) {
    const worksheetData = DataAccess.getWorksheetsList(spreadsheet);
    const storedSignatures = getStoredColumnSignatures();
    const changedWorksheets = [];

    worksheetData.forEach(function (worksheetInfo) {
      const worksheetName = worksheetInfo.worksheetName;
      const ws = DataAccess.getSheetByName(worksheetName, spreadsheet);

      if (ws) {
        const currentSignature = generateColumnSignature(ws);

        if (
          storedSignatures[worksheetName] &&
          storedSignatures[worksheetName] !== currentSignature
        ) {
          changedWorksheets.push(worksheetName);
        }
      }
    });

    return changedWorksheets;
  }

  /**
   * Captures current column selections from selector sheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [spreadsheet] - The spreadsheet (uses active if not provided)
   * @return {Object} Object with saved selections or empty object if none
   */
  function captureColumnSelections(spreadsheet) {
    if (!spreadsheet) spreadsheet = DataAccess.getActiveSpreadsheet();

    const selectorSheet = spreadsheet.getSheetByName(CONSTANTS.SELECTOR_SHEET);
    if (!selectorSheet) return {};

    const lastRow = selectorSheet.getLastRow();
    // Skip header rows
    if (lastRow <= 2) return {};

    // Get all data in one operation to minimize API calls
    const data = selectorSheet.getRange(3, 1, lastRow - 2, 4).getValues();

    const selections = {};

    data.forEach(function (row) {
      const worksheet = row[0];
      const column = row[1];
      const isChecked = row[2];
      const group = row[3];

      // Skip empty rows
      if (!worksheet || !column) return;

      // Initialize nested objects if needed
      if (!selections[worksheet]) {
        selections[worksheet] = {};
      }

      selections[worksheet][column] = {
        checked: Boolean(isChecked),
        group: group || "",
      };
    });

    // Store in document properties
    getDocumentProperties().setProperty(
      CONSTANTS.PROPERTY_KEYS.COLUMN_SELECTIONS,
      JSON.stringify(selections)
    );

    return selections;
  }

  /**
   * Restores saved column selections to the selector sheet
   * @param {Object} [savedSelections] - Selections to restore (retrieves from properties if not provided)
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [spreadsheet] - The spreadsheet (uses active if not provided)
   */
  function restoreColumnSelections(savedSelections = null, spreadsheet = null) {
    if (!spreadsheet) spreadsheet = DataAccess.getActiveSpreadsheet();

    const selectorSheet = spreadsheet.getSheetByName(CONSTANTS.SELECTOR_SHEET);
    if (!selectorSheet) return;

    // If no selections passed, try to get from document properties
    if (!savedSelections) {
      try {
        const selectionsJson = getDocumentProperties().getProperty(
          CONSTANTS.PROPERTY_KEYS.COLUMN_SELECTIONS
        );
        if (!selectionsJson) return;
        savedSelections = JSON.parse(selectionsJson);
      } catch (e) {
        console.error("Error retrieving saved selections:", e);
        return;
      }
    }

    // Get current selector data
    const lastRow = selectorSheet.getLastRow();
    if (lastRow <= 2) return; // Just header

    const data = selectorSheet.getRange(3, 1, lastRow - 2, 4).getValues();

    // Create ranges for batch updates
    const checkboxUpdates = [];
    const groupUpdates = [];

    data.forEach(function (row, index) {
      const rowIndex = index + 3; // Adjust for header rows
      const worksheet = row[0];
      const column = row[1];

      // Skip empty rows
      if (!worksheet || !column) return;

      // Check if we have saved selection for this worksheet/column
      if (savedSelections[worksheet] && savedSelections[worksheet][column]) {
        const selection = savedSelections[worksheet][column];

        // Store checkbox updates for batch operations
        if (selection.checked) {
          checkboxUpdates.push(rowIndex);
        }

        // Store group values for batch updates
        if (selection.group) {
          groupUpdates.push({
            row: rowIndex,
            value: selection.group,
          });
        }
      }
    });

    // Apply checkbox updates in batches
    if (checkboxUpdates.length > 0) {
      // We can update in continuous ranges for efficiency
      let start = checkboxUpdates[0];
      let end = checkboxUpdates[0];

      for (let i = 1; i <= checkboxUpdates.length; i++) {
        if (i === checkboxUpdates.length || checkboxUpdates[i] !== end + 1) {
          // Update this range
          selectorSheet.getRange(start, 3, end - start + 1, 1).check();

          // Start new range if needed
          if (i < checkboxUpdates.length) {
            start = checkboxUpdates[i];
            end = checkboxUpdates[i];
          }
        } else {
          // Continue the range
          end = checkboxUpdates[i];
        }
      }
    }

    // Apply group updates (need to set cell by cell for different values)
    groupUpdates.forEach(function (update) {
      selectorSheet.getRange(update.row, 4).setValue(update.value);
    });
  }

  /**
   * Resets stored column signatures (for troubleshooting)
   * @return {boolean} True if reset successful
   */
  function resetColumnSignatures() {
    try {
      getDocumentProperties().deleteProperty(
        CONSTANTS.PROPERTY_KEYS.COLUMN_SIGNATURES
      );
      return true;
    } catch (e) {
      console.error("Error resetting column signatures:", e);
      return false;
    }
  }

  /**
   * Sets the header message in the selector sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} selectorSheet - The selector sheet to update
   * @param {string} [actionType="generated"] - Action description ("generated", "updated", etc.)
   */
  function setColumnSelectorHeader(selectorSheet, actionType = "generated") {
    if (!selectorSheet) return;

    selectorSheet.getRange(CONSTANTS.SELECTOR_HEADER_RANGE).merge();
    selectorSheet
      .getRange("A1")
      .setValue(
        "Column Selector " +
          actionType +
          " on " +
          new Date().toLocaleString() +
          ". Select columns to include in charts and assign them to groups."
      )
      .setFontStyle("italic");
  }

  // Public API
  return {
    CONSTANTS: CONSTANTS,
    generateColumnSignature: generateColumnSignature,
    storeColumnSignatures: storeColumnSignatures,
    getStoredColumnSignatures: getStoredColumnSignatures,
    detectColumnChanges: detectColumnChanges,
    captureColumnSelections: captureColumnSelections,
    restoreColumnSelections: restoreColumnSelections,
    resetColumnSignatures: resetColumnSignatures,
    setColumnSelectorHeader: setColumnSelectorHeader,
  };
})();
