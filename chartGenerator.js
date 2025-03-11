/**
 * Chart Generator Module for MOAA Chart Tools
 * Handles creation of chart data structures for visualization
 */
var ChartGenerator = (function () {
  // Constants for chart generation
  const CONSTANTS = {
    // Chart types
    CHART_TYPES: {
      BAR: "bar",
      PIE: "pie",
      LINE: "line",
    },

    // Dataset labels
    AGGREGATE_LABEL: "Aggregate",

    // Default values
    DEFAULT_CHART_TYPE: "bar",
    FALLBACK_VALUE: 0,
  };

  /**
   * Generates chart data structure for visualization
   * @param {Object} worksheetData - Processed worksheet data
   * @param {string} chartType - Type of chart (bar, pie, line, etc.)
   * @param {string} chartGroupsJSON - JSON string of column groupings
   * @param {string} workshopId - ID of the current workshop
   * @return {Object|null} Chart data object or null if generation fails
   */
  function generateChartData(
    worksheetData,
    chartType,
    chartGroupsJSON,
    workshopId
  ) {
    // Validate inputs
    if (!worksheetData || !chartGroupsJSON) {
      console.error("Missing required data for chart generation");
      return null;
    }

    // Parse chart groups
    let chartGroups;
    try {
      chartGroups = JSON.parse(chartGroupsJSON);
    } catch (e) {
      console.error("Error parsing chart groups:", e);
      return null;
    }

    // Initialize chart data structure
    const chartData = {
      type: chartType || CONSTANTS.DEFAULT_CHART_TYPE,
      data: [],
    };

    // Get data sources
    const aggregateRow = worksheetData.data[0];
    const workshopData = worksheetData.columnsSummaryData;

    if (!aggregateRow || !workshopData) {
      console.error("Missing aggregate or workshop data");
      return null;
    }

    // Process each chart group
    Object.keys(chartGroups).forEach(function (groupName) {
      const columns = chartGroups[groupName];
      if (!columns || columns.length === 0) return;

      // Create data structure for this group
      const groupData = {
        labels: columns, // Column names become labels
        datasets: [
          {
            label: CONSTANTS.AGGREGATE_LABEL,
            data: [],
          },
          {
            label: workshopId,
            data: [],
          },
        ],
      };

      // Extract data for each column
      columns.forEach(function (column) {
        // Check if column exists in data sources
        if (
          aggregateRow.hasOwnProperty(column) &&
          workshopData.hasOwnProperty(column)
        ) {
          // Add aggregate data
          groupData.datasets[0].data.push(
            Utils.extractNumericValue(aggregateRow[column])
          );

          // Add workshop-specific data
          groupData.datasets[1].data.push(
            Utils.extractNumericValue(workshopData[column])
          );
        } else {
          // Push default value for missing columns
          groupData.datasets[0].data.push(CONSTANTS.FALLBACK_VALUE);
          groupData.datasets[1].data.push(CONSTANTS.FALLBACK_VALUE);
          console.warn(`Column "${column}" not found in dataset`);
        }
      });

      chartData.data.push(groupData);
    });

    return chartData;
  }

  /**
   * Generates a pie chart data structure
   * @param {Object} worksheetData - Processed worksheet data
   * @param {string[]} columns - Columns to include in the chart
   * @param {string} workshopId - ID of the current workshop
   * @return {Object} Pie chart data structure
   */
  function generatePieChartData(worksheetData, columns, workshopId) {
    // Create fake group structure for reuse with main generator
    const chartGroups = { group1: columns };
    return generateChartData(
      worksheetData,
      CONSTANTS.CHART_TYPES.PIE,
      JSON.stringify(chartGroups),
      workshopId
    );
  }

  /**
   * Generates a bar chart data structure
   * @param {Object} worksheetData - Processed worksheet data
   * @param {string[]} columns - Columns to include in the chart
   * @param {string} workshopId - ID of the current workshop
   * @return {Object} Bar chart data structure
   */
  function generateBarChartData(worksheetData, columns, workshopId) {
    // Create fake group structure for reuse with main generator
    const chartGroups = { group1: columns };
    return generateChartData(
      worksheetData,
      CONSTANTS.CHART_TYPES.BAR,
      JSON.stringify(chartGroups),
      workshopId
    );
  }

  /**
   * Detects if the specified columns are appropriate for a certain chart type
   * @param {string} chartType - Type of chart to validate for
   * @param {Object} worksheetData - Processed worksheet data
   * @param {string[]} columns - Columns to check
   * @return {boolean} True if columns are appropriate for the chart type
   */
  function validateColumnsForChartType(chartType, worksheetData, columns) {
    // This is a placeholder for future validation logic
    // For example, pie charts might require percentage data
    // Bar charts might need certain data distribution

    if (!worksheetData || !columns || columns.length === 0) {
      return false;
    }

    return true;
  }

  // Public API
  return {
    CONSTANTS: CONSTANTS,
    generateChartData: generateChartData,
    generatePieChartData: generatePieChartData,
    generateBarChartData: generateBarChartData,
    validateColumnsForChartType: validateColumnsForChartType,
  };
})();
