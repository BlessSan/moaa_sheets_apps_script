/**
 * Worksheet Service Module for MOAA Chart Tools
 *
 * Purpose: Orchestrates worksheet data processing and chart generation
 *
 * Dependencies:
 * - DataProcessor: For data transformation
 * - ChartGenerator: For chart data creation
 * - DataAccess: For spreadsheet data access
 *
 * Dependent Modules:
 * - main.js: For HTTP request handling
 * - UIManager: For worksheet operations
 *
 * Initialization Order: Should be loaded after its dependencies
 */
var WorksheetService = (function () {
  // Constants for worksheet processing
  const CONSTANTS = {
    // Workshop ID column is always the first column (index 0)
    WORKSHOP_ID_COLUMN: 0,

    // Processing options
    DEFAULT_CHART_TYPE: "bar",
    DEFAULT_WORKSHEET_TYPE: "dynamic",
  };

  /**
   * Gets data from all worksheets, optionally filtered by workshop ID
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet to process
   * @param {string} [workshopId] - Optional workshop ID to filter data
   * @param {Array<Object>} [worksheetList] - Optional list of worksheet configurations (fetched if not provided)
   * @return {Array<Object>} Array of processed worksheet data objects
   */
  function getWorksheetsData(spreadsheet, workshopId, worksheetList) {
    // Get worksheet list if not provided
    if (!worksheetList) {
      worksheetList = DataAccess.getWorksheetsList(spreadsheet);
    }

    // Process each worksheet
    const results = [];

    worksheetList.forEach(function (worksheetConfig) {
      // Process the worksheet
      const worksheetData = processWorksheetWithFilter(
        spreadsheet,
        worksheetConfig,
        workshopId
      );

      // Add to results if we got data
      if (worksheetData) {
        results.push(worksheetData);
      }
    });

    return results;
  }

  /**
   * Processes a single worksheet with optional filtering
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
   * @param {Object} worksheetConfig - Configuration for the worksheet
   * @param {string} [workshopId] - Optional workshop ID to filter data
   * @return {Object|null} Processed worksheet data or null if error
   */
  function processWorksheetWithFilter(
    spreadsheet,
    worksheetConfig,
    workshopId
  ) {
    const worksheetName = worksheetConfig.worksheetName;
    const worksheetType =
      worksheetConfig.type || CONSTANTS.DEFAULT_WORKSHEET_TYPE;

    // Get the worksheet
    // TODO: evalute if this is necessary. It only checks if the sheet exists but adds additional read operations
    // const sheet = DataAccess.getSheetByName(worksheetName, spreadsheet);
    // if (!sheet) {
    //   console.error(`Worksheet ${worksheetName} not found`);
    //   return null;
    // }

    // Get raw worksheet data
    const rawData = DataAccess.getWorksheetData(
      worksheetName,
      null,
      spreadsheet
    );
    if (!rawData.headers || rawData.headers.length === 0) {
      console.error(`No headers found in worksheet ${worksheetName}`);
      return null;
    }

    // Process the worksheet data
    const processedData = DataProcessor.processWorksheetData(
      rawData,
      worksheetName,
      worksheetType,
      workshopId
    );

    // Generate chart data if applicable
    if (workshopId && worksheetConfig.chartGroups) {
      const chartData = prepareChartDataForWorksheet(
        processedData,
        worksheetConfig,
        workshopId
      );

      if (chartData) {
        processedData.chartData = chartData;
      }
    }

    return processedData;
  }

  /**
   * Prepares chart data for a worksheet based on configuration
   * @param {Object} processedData - Processed worksheet data
   * @param {Object} worksheetConfig - Configuration for the worksheet
   * @param {string} workshopId - Workshop ID for the filtered data
   * @return {Object|null} Chart data object or null if generation fails
   */
  function prepareChartDataForWorksheet(
    processedData,
    worksheetConfig,
    workshopId
  ) {
    const chartType = worksheetConfig.chartType || CONSTANTS.DEFAULT_CHART_TYPE;
    const chartGroups = worksheetConfig.chartGroups;
    const chartTitle = worksheetConfig.chartTitle;

    if (!chartGroups) {
      return null;
    }

    return ChartGenerator.generateChartData(
      processedData,
      chartTitle,
      chartType,
      chartGroups,
      workshopId
    );
  }

  //TODO: evalue if this function is necessary, seem to be unsed now
  /**
   * Gets processed data for a single worksheet
   * @param {string} worksheetName - Name of the worksheet
   * @param {string} [workshopId] - Optional workshop ID to filter by
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [spreadsheet] - Spreadsheet (uses active if not provided)
   * @return {Object|null} Processed worksheet data or null if error
   */
  function getWorksheetData(worksheetName, workshopId, spreadsheet) {
    if (!spreadsheet) spreadsheet = DataAccess.getActiveSpreadsheet();

    // Get worksheet configuration
    const worksheetList = DataAccess.getWorksheetsList(spreadsheet);
    const worksheetConfig = worksheetList.find(
      (config) => config.worksheetName === worksheetName
    );

    if (!worksheetConfig) {
      console.error(`No configuration found for worksheet ${worksheetName}`);
      return null;
    }

    return processWorksheetWithFilter(spreadsheet, worksheetConfig, workshopId);
  }

  /**
   * Gets an overview of all worksheets in the spreadsheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [spreadsheet] - Spreadsheet (uses active if not provided)
   * @return {Array<Object>} Summary information for each worksheet
   */
  function getWorksheetsSummary(spreadsheet) {
    if (!spreadsheet) spreadsheet = DataAccess.getActiveSpreadsheet();

    const worksheetList = DataAccess.getWorksheetsList(spreadsheet);
    return worksheetList.map(function (config) {
      const sheet = DataAccess.getSheetByName(
        config.worksheetName,
        spreadsheet
      );

      return {
        name: config.worksheetName,
        type: config.type || CONSTANTS.DEFAULT_WORKSHEET_TYPE,
        rows: sheet ? sheet.getLastRow() : 0,
        columns: sheet ? sheet.getLastColumn() : 0,
        hasChartConfig: Boolean(config.chartGroups),
      };
    });
  }

  // Public API
  return {
    CONSTANTS: CONSTANTS,
    getWorksheetsData: getWorksheetsData,
    processWorksheetWithFilter: processWorksheetWithFilter,
    prepareChartDataForWorksheet: prepareChartDataForWorksheet,
    getWorksheetsSummary: getWorksheetsSummary,
  };
})();
