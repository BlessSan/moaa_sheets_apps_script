/**
 * Utilities Module for MOAA dashboard
 *
 * Purpose: Provides common helper functions for the application
 *
 * Dependencies: None
 *
 * Dependent Modules:
 * - DataProcessor: For number formatting and data extraction
 * - ChartGenerator: For data formatting
 * - DataAccess: For helper utilities
 * - WorksheetService: Indirectly via other modules
 * - UIManager: Indirectly via other modules
 *
 * Initialization Order: Should be loaded after ServiceLocator
 */
var Utils = (function () {
  /**
   * Formats a value as US currency
   * @param {number|string} value - Value to format
   * @return {string} Formatted currency string
   */
  function formatCurrency(value) {
    if (!value && value !== 0) return "";
    const numValue = parseFloat(value);
    if (isNaN(numValue)) return "";

    return new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
    }).format(numValue);
  }

  /**
   * Converts specified columns in a data array to currency format
   * @param {string[][] | object[]} data - Data to process (2D array or array of objects)
   * @param {number[]|string[]} columns - Column indices or keys to format as currency
   * @param {boolean} modifyInPlace - Whether to modify the original data (default: true)
   * @return {string[][] | object[]} - The formatted data (same as input if modifyInPlace=true)
   */
  function formatCurrencyColumns(data, columns, modifyInPlace = true) {
    if (!data || !columns || columns.length === 0) return data;

    const formatter = new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
    });

    // Clone the data if not modifying in place
    const result = modifyInPlace ? data : JSON.parse(JSON.stringify(data));

    // Handle array of arrays (2D array)
    if (Array.isArray(data) && Array.isArray(data[0])) {
      for (let i = 0; i < result.length; i++) {
        for (let j = 0; j < columns.length; j++) {
          const colIndex = columns[j];
          const value = result[i][colIndex];
          if (value !== "" && !isNaN(value)) {
            result[i][colIndex] = formatter.format(parseInt(Number(value)));
          }
        }
      }
    }
    // Handle array of objects
    else if (Array.isArray(data) && typeof data[0] === "object") {
      for (let i = 0; i < result.length; i++) {
        for (let j = 0; j < columns.length; j++) {
          const key = columns[j];
          const value = result[i][key];
          if (value !== "" && !isNaN(value)) {
            result[i][key] = formatter.format(parseInt(Number(value)));
          }
        }
      }
    }

    return result;
  }

  /**
   * Extracts numeric value from different data formats
   * @param {any} value - The value to extract number from
   * @return {number} The extracted numeric value
   */
  function extractNumericValue(value) {
    // Handle numeric values directly
    if (typeof value === "number" || !isNaN(value)) {
      return Number(parseFloat(value).toFixed(2));
    }
    // Handle string values
    else if (typeof value === "string") {
      // Format: "X (Y%)" - extract X
      const countMatch = value.match(/^(\d+)/);
      if (countMatch) return parseFloat(countMatch[1]);

      // Format: "$X,XXX.XX" - extract X,XXX.XX
      const currencyMatch = value.match(/[\$]?([\d,]+(\.\d+)?)/);
      if (currencyMatch) {
        return parseFloat(currencyMatch[1].replace(/,/g, ""));
      }

      // Try direct number parsing
      const numValue = parseFloat(value).toFixed(2);
      if (!isNaN(numValue)) return numValue;
    }

    return 0;
  }

  /**
   * Finds the last non-empty row in a column range
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to check
   * @param {number} column - Column index (1-based)
   * @return {number} The last non-empty row
   */
  function getLastRowWithContent(sheet, column = 1) {
    const values = sheet.getRange(1, column, sheet.getMaxRows(), 1).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][0] !== "") {
        return i + 1;
      }
    }
    return 1;
  }

  /**
   * Extract value for custom format of X (Y%)
   * @typedef {Object} ExtractedNumbers
   * @property {number} x - The extracted value X
   * @property {number} y - The extracted percentage Y
   * @param {string} value - The value to extract from
   * @return {ExtractedNumbers|null} The extracted numbers or null if not found
   */
  function extractCustomFormatValue(value) {
    const regex =
      /(\d+(?:,\d{3})*)(?:\.\d+)?\s*\((\d+(?:,\d{3})*(?:\.\d+)?)%\)/;

    const match = value.match(regex);
    if (match) {
      const x = parseInt(match[1].replace(/,/g, ""), 10);
      const y = parseFloat(match[2].replace(/,/g, ""));
      return { x: x, y: y };
    }
    return null;
  }

  /**
   * Logs errors with consistent formatting
   * @param {string} source - Source of the error (function name)
   * @param {Error|string} error - The error object or message
   */
  function logError(source, error) {
    console.error(`[${source}] ${error.message || error}`);
    if (error.stack) {
      console.error(error.stack);
    }
  }

  /**
   * Safely executes a function with error handling
   * @param {Function} fn - Function to execute
   * @param {Array} args - Arguments to pass to the function
   * @param {string} source - Source for error logging
   * @return {any} Result of the function or null if error
   */
  function safeExecute(fn, args, source) {
    try {
      return fn.apply(null, args);
    } catch (e) {
      logError(source, e);
      return null;
    }
  }

  // Public API
  return {
    formatCurrency: formatCurrency,
    formatCurrencyColumns: formatCurrencyColumns,
    extractNumericValue: extractNumericValue,
    getLastRowWithContent: getLastRowWithContent,
    extractCustomFormatValue: extractCustomFormatValue,
    logError: logError,
    safeExecute: safeExecute,
  };
})();
