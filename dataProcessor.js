/**
 * Data Processing Module for MOAA Chart Tools
 * Handles transformation of raw spreadsheet data into structured formats
 */
var DataProcessor = (function () {
  // Constants for data processing
  const CONSTANTS = {
    // Column identifiers
    WORKSHOP_ID_COLUMN_INDEX: 0,
    AGGREGATE_INDICATOR_TEXT: "Aggregate",

    // Column types
    CURRENCY_COLUMN_NAMES: ["Annual Revenue", "Revenue per Leadership Team"],

    // Processing configuration
    FIRST_ROW_IS_AGGREGATE: true,
    FORMAT_PRECISION: 2,
  };

  /**
   * Transforms raw 2D array data into array of objects with named properties
   * @param {string[]} columns - Column names (headers)
   * @param {Array<Array>} values - 2D array of values
   * @param {string} type - Data type (e.g., "workshop" to skip ID column)
   * @return {Array<Object>} Array of objects with column names as keys
   */
  function transformData(columns, values, type) {
    if (!columns || !values || values.length === 0) return [];

    const result = [];
    const colStart = type === "workshop" ? 1 : 0;

    for (let row = 0; row < values.length; row++) {
      const resultObj = {};

      for (let col = colStart; col < columns.length; col++) {
        if (columns[col]) {
          // Skip empty column names
          resultObj[columns[col]] = values[row][col];
        }
      }

      result.push(resultObj);
    }

    return result;
  }

  /**
   * Calculates summary for a non-numeric column from filtered data
   * @param {number} columnIndex - Index of the column
   * @param {Array<Array>} filteredRows - Filtered data rows
   * @return {string} Count and percentage (e.g., "5 (50.00%)")
   */
  function calculateColumnNaN(columnIndex, filteredRows) {
    let count = 0;

    for (let i = 0; i < filteredRows.length; i++) {
      const rowValue = filteredRows[i][columnIndex];
      if (rowValue) {
        count += 1;
      }
    }

    const averagePercent = (count / filteredRows.length) * 100;
    return `${count} (${averagePercent.toFixed(CONSTANTS.FORMAT_PRECISION)}%)`;
  }

  /**
   * Calculates average for a numeric column from filtered data
   * @param {number} columnIndex - Index of the column
   * @param {Array<Array>} filteredRows - Filtered data rows
   * @return {string} Formatted average value
   */
  function calculateColumnAverage(columnIndex, filteredRows) {
    let sum = 0;

    for (let i = 0; i < filteredRows.length; i++) {
      const rowValue = filteredRows[i][columnIndex];
      if (rowValue) {
        sum += parseFloat(rowValue);
      }
    }

    const avg = sum / filteredRows.length;
    return avg.toFixed(CONSTANTS.FORMAT_PRECISION);
  }

  /**
   * Generates summary data for filtered results
   * @param {string[]} columns - Column names
   * @param {Array} aggregateDataRow - First row with aggregate data
   * @param {Array<Array>} filteredRows - Filtered data rows
   * @param {number[]|null} currencyColIndex - Indices of currency columns
   * @return {Object} Summary data object with column names as keys
   */
  function getFilteredDataColumnSummary(
    columns,
    aggregateDataRow,
    filteredRows,
    currencyColIndex = null
  ) {
    let result = {};

    // Process each column
    for (let index = 0; index < aggregateDataRow.length; index++) {
      const aggregateDataValue = aggregateDataRow[index];

      if (aggregateDataValue) {
        // Format numeric aggregate values
        if (aggregateDataValue !== "" && !isNaN(aggregateDataValue)) {
          aggregateDataRow[index] = aggregateDataValue.toFixed(
            CONSTANTS.FORMAT_PRECISION
          );
        }

        // Calculate column summaries if we have filtered rows
        if (filteredRows.length > 0) {
          if (isNaN(aggregateDataValue)) {
            // Non-numeric column - count occurrences
            const columnSummary = calculateColumnNaN(index, filteredRows);
            result[columns[index]] = columnSummary;
          } else {
            // Numeric column - calculate average
            let columnSummary = calculateColumnAverage(index, filteredRows);

            // Format as currency if needed
            if (currencyColIndex && currencyColIndex.includes(index)) {
              columnSummary = Utils.formatCurrency(columnSummary);
            }

            result[columns[index]] = columnSummary;
          }
        }
      }
    }

    return result;
  }

  /**
   * Detects currency columns based on column names
   * @param {string[]} columns - Column names
   * @return {number[]} Array of indices for currency columns
   */
  function detectCurrencyColumns(columns) {
    if (!columns) return [];

    const currencyIndices = [];
    for (let i = 0; i < columns.length; i++) {
      if (CONSTANTS.CURRENCY_COLUMN_NAMES.includes(columns[i])) {
        currencyIndices.push(i);
      }
    }

    return currencyIndices;
  }

  /**
   * Processes worksheet data with filtering and aggregation
   * @param {Object} rawData - Raw worksheet data { headers, data }
   * @param {string} worksheetName - Name of the worksheet
   * @param {string} worksheetType - Type of worksheet (static, dynamic)
   * @param {string|null} filterId - ID to filter by (null for all data)
   * @return {Object} Processed worksheet data object
   */
  function processWorksheetData(
    rawData,
    worksheetName,
    worksheetType,
    filterId
  ) {
    const { headers, data } = rawData;

    if (!headers || !data || data.length === 0) {
      return {
        worksheet: worksheetName,
        type: worksheetType,
        isWorkshopTable: false,
        data: [],
        columnsSummaryData: {},
      };
    }

    // Detect currency columns
    const currencyColIndices = detectCurrencyColumns(headers);

    // If no filter ID, return just the aggregate data
    if (!filterId) {
      // note to self: the use of spread operator here is to create a copy of the array so that the original data array is not modified
      // details: though spread operator creates a shallow copy, it is enough for this use case since aggregateRow is modified in a way that its primitive value changes, therefore does not affect the original data array
      const aggregateRow = [...data[0]];

      // Format currency values
      if (currencyColIndices.length > 0) {
        Utils.formatCurrencyColumns([aggregateRow], currencyColIndices);
      }

      const transformedData = transformData(
        headers,
        [aggregateRow],
        worksheetType
      );

      return {
        worksheet: `Aggregate ${worksheetName}`,
        type: "static",
        isWorkshopTable: false,
        data: transformedData,
        columnsSummaryData: {},
      };
    }

    // Filter rows by ID
    const filteredRows = data.filter(
      (row) => row[CONSTANTS.WORKSHOP_ID_COLUMN_INDEX] === filterId
    );

    // Get aggregate data row
    const aggregateDataRow = [...data[0]];

    // Calculate summaries for filtered data
    const columnSummaryData = getFilteredDataColumnSummary(
      headers,
      aggregateDataRow,
      filteredRows,
      currencyColIndices
    );

    // Combine aggregate row with filtered rows
    const resultDataRows = [aggregateDataRow, ...filteredRows];

    // Format currency columns
    if (currencyColIndices.length > 0) {
      Utils.formatCurrencyColumns(resultDataRows, currencyColIndices);
    }

    // Transform to array of objects
    const transformedData = transformData(
      headers,
      resultDataRows,
      worksheetType
    );

    return {
      worksheet: worksheetName,
      type: worksheetType,
      isWorkshopTable: true,
      data: transformedData,
      columnsSummaryData: columnSummaryData,
    };
  }

  // Public API
  return {
    CONSTANTS: CONSTANTS,
    transformData: transformData,
    calculateColumnNaN: calculateColumnNaN,
    calculateColumnAverage: calculateColumnAverage,
    getFilteredDataColumnSummary: getFilteredDataColumnSummary,
    detectCurrencyColumns: detectCurrencyColumns,
    processWorksheetData: processWorksheetData,
  };
})();
