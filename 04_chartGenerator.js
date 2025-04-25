/**
 * Chart Generator Module for MOAA Chart Tools
 *
 * Purpose: Handles creation of chart data structures for visualization
 *
 * Dependencies:
 * - Utils: For number extraction and formatting
 *
 * Dependent Modules:
 * - WorksheetService: For chart data generation
 *
 * Initialization Order: Should be loaded after Utils
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
   * @param {string} chartTitle - Title of the chart
   * @param {string} chartType - Type of chart (bar, pie, line, etc.)
   * @param {string} chartGroupsJSON - JSON string of column groupings
   * @param {string} workshopId - ID of the current workshop
   * @return {Object|null} Chart data object or null if generation fails
   */
  function generateChartData(
    worksheetData,
    chartTitle,
    chartType,
    chartGroupsJSON,
    workshopId,
    isAggregateOnly = false
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
      title: chartTitle || "",
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

      // Create data structure for this group, where aggregate only is one dataset
      const groupData = {
        labels: columns, // Column names become labels
        datasets: isAggregateOnly
          ? [
              // For aggregate-only, create just one dataset
              {
                label: CONSTANTS.AGGREGATE_LABEL,
                data: [],
                customLabels: [],
              },
            ]
          : [
              // For comparison view, create two datasets
              {
                label: CONSTANTS.AGGREGATE_LABEL,
                data: [],
                customLabels: [],
              },
              {
                label: workshopId,
                data: [],
                customLabels: [],
              },
            ],
      };

      // Extract data for each column
      columns.forEach(function (column) {
        if (aggregateRow.hasOwnProperty(column)) {
          // Add aggregate data
          const aggregateDataPoint = aggregateRow[column];
          const aggregateValue = Utils.extractNumericValue(aggregateDataPoint);
          groupData.datasets[0].data.push(aggregateValue);
          groupData.datasets[0].customLabels.push(aggregateDataPoint);

          // Add workshop-specific data only if not in aggregate-only mode
          if (!isAggregateOnly && workshopData.hasOwnProperty(column)) {
            const workshopDataPoint = workshopData[column];
            groupData.datasets[1].data.push(
              Utils.extractNumericValue(workshopDataPoint)
            );
            groupData.datasets[1].customLabels.push(workshopDataPoint);
          }
        } else {
          // Push default value for missing columns
          groupData.datasets[0].data.push(CONSTANTS.FALLBACK_VALUE);
          groupData.datasets[0].customLabels.push(CONSTANTS.FALLBACK_VALUE);
          groupData.datasets[1].data.push(CONSTANTS.FALLBACK_VALUE);
          groupData.datasets[1].customLabels.push(CONSTANTS.FALLBACK_VALUE);
          console.warn(`Column "${column}" not found in dataset`);
        }
      });

      // Special handling for single data points in pie charts
      //! IMPORTANT NOTE: This implementation assumes that pie chart will always handle data with format of X (Y%)
      if (columns.length === 1 && chartType === CONSTANTS.CHART_TYPES.PIE) {
        // add datapoint for remainder values in Others category

        //try to extract value for custom format of X (Y%)
        const extractedAggregateValue = Utils.extractCustomFormatValue(
          aggregateRow[columns[0]]
        );
        // For aggregate-only mode, only process aggregate data
        if (isAggregateOnly) {
          if (extractedAggregateValue !== null) {
            // Add Others category for aggregate data only
            columns.push("Others");
            const aggregateRemainder = getRemainderValue(
              extractedAggregateValue
            );

            groupData.datasets[0].data.push(aggregateRemainder.x);
            groupData.datasets[0].customLabels.push(
              `${aggregateRemainder.x} (${aggregateRemainder.y}%)`
            );
          }
        }
        // For comparison mode (with workshop data)
        else {
          const extractedWorkshopValue = Utils.extractCustomFormatValue(
            workshopData[columns[0]]
          );

          if (
            extractedAggregateValue !== null &&
            extractedWorkshopValue !== null
          ) {
            // Add Others category
            columns.push("Others");
            const aggregateRemainder = getRemainderValue(
              extractedAggregateValue
            );
            const workshopRemainder = getRemainderValue(extractedWorkshopValue);

            groupData.datasets[0].data.push(aggregateRemainder.x);
            groupData.datasets[0].customLabels.push(
              `${aggregateRemainder.x} (${aggregateRemainder.y}%)`
            );

            groupData.datasets[1].data.push(workshopRemainder.x);
            groupData.datasets[1].customLabels.push(
              `${workshopRemainder.x} (${workshopRemainder.y}%)`
            );
          }
        }
      }

      chartData.data.push(groupData);
    });

    return chartData;
  }

  /**
   * @description Gets the remainder value for a pie chart with format of X (Y%)
   *
   * Assumes that `x` and `y` are numeric, and `y` represents a percentage value.
   *
   * NOTE: This function could be more complete with handling remainder values of different formats
   * @typedef {Object} ExtractedNumbers
   * @property {number} x - The extracted value X
   * @property {number} y - The extracted percentage Y
   * @param {ExtractedNumbers} data - Data object for the chart
   * @return {ExtractedNumbers} Remainder value x and its percentage y
   */
  function getRemainderValue({ x, y }) {
    const remainder = (x * 100) / y - x;
    return { x: Math.round(remainder), y: parseFloat((100 - y).toFixed(2)) };
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
