/* 
worksheets will be a list of object
{
  worksheetName:string,
  type: "static" | "dynamic"
}
*/
function getWorksheetsData(sheet, id, worksheetList) {
  // accepted format:
  // workshop id will always in 1st column (index 0)
  const workshopIdColumnIndex = 0;

  // get values of registered workshop id under "Workshop Planner"
  // a list of {value:workshopId, label:workshopLabel}
  const recordedWorkshopIdList = getWorkshopList(sheet);

  let worksheetsData = [];

  // loop over each worksheet
  worksheetList.forEach(function (worksheet, index) {
    const worksheetName = worksheet.worksheetName;
    const worksheetType = worksheet.type;
    const ws = sheet.getSheetByName(worksheetName);

    const lastColumn = ws.getLastColumn();
    const lastRow = ws.getLastRow();

    const aVals = ws.getRange("A1:A" + lastRow).getValues();
    const aLast = lastRow - aVals.reverse().findIndex((c) => c[0] != "");

    const worksheetRange = ws.getRange(1, 1, aLast, lastColumn);

    // columns: [col1, col1, col3]
    // values: [[row1 col1, row1, col2], [rowN col1, rowN col2, rowN col3]]
    const [columns, ...values] = worksheetRange.getValues();

    // check if this worksheet contains workshop id
    // check if first column values exist in recordedWorkshopIdList
    /**
     * For now each worksheet is a workshop
    const isWorkshopSheet = values.some(function(arrayOfValues){
      return recordedWorkshopIdList.some(function (workshopIdObj){
        // each value at col1 will be checked for each value in recordedWorkshopList
        return arrayOfValues[workshopIdColumnIndex] === workshopIdObj.value
      })
    })
     */

    const isWorkshopSheet = true;

    // get first row that contains summary value of each column (aggregate data)
    const aggregateDataRow = values[0];

    let currencyIndex = null;

    //TODO: check if value should be formatted as US currency
    //check the index of column "Annual Revenue" and "Revenue per Leadership Team"

    if (
      columns.includes("Annual Revenue") ||
      columns.includes("Revenue per Leadership Team")
    ) {
      currencyIndex = [
        columns.indexOf("Annual Revenue"),
        columns.indexOf("Revenue per Leadership Team"),
      ];
    }

    // if its a worksheet that has workshop data to process, do
    // check if id exist
    if (isWorkshopSheet) {
      // if Id exist, do
      // filter by id
      // calculate each filtered column based on first row value
      // if first row value isNaN => do count(count/total)
      // if first row value !isNaN => do average of that column
      if (id) {
        const filteredRows = values.filter(
          (row) => row[workshopIdColumnIndex] === id
        );

        // result rows will be a combination of
        // first row which contains summary values
        // rest of row that is filtered by id

        const columnSummaryData = getFilteredDataColumnSummary(
          columns,
          aggregateDataRow,
          filteredRows,
          currencyIndex
        );
        let resultDataRow = [aggregateDataRow, ...filteredRows];

        if (currencyIndex) {
          //this function modifies the passed array, does not return new array
          convertToCurrency(resultDataRow, currencyIndex);
        }

        const transformedData = transformData(
          columns,
          resultDataRow,
          "workshop"
        );

        let worksheetResult = {
          worksheet: worksheetName,
          type: worksheetType,
          isWorkshopTable: isWorkshopSheet,
          data: transformedData,
          columnsSummaryData: columnSummaryData,
        };

        if (id && worksheet.chartGroups) {
          const chartData = generateChartData(
            { data: transformedData, columnsSummaryData: columnSummaryData },
            worksheet.chartType,
            worksheet.chartGroups,
            id
          );

          // Only add the chartData property if it was successfully generated
          if (chartData) {
            worksheetResult.chartData = chartData;
          }
        }

        worksheetsData.push(worksheetResult);
      }
      // if there is no id, get only the first row
      else {
        // since aggregate data row is a 1D array, convert it to 2D
        let emptyColsIndex = [];
        aggregateDataRow.forEach(function (value, index, array) {
          if (value !== "" && !isNaN(value)) {
            if (currencyIndex && currencyIndex.includes(index)) {
              array[index] = new Intl.NumberFormat("en-US", {
                style: "currency",
                currency: "USD",
              }).format(parseInt(value));
            } else {
              array[index] = value.toFixed(2);
            }
          }
        });
        // filter empty cells
        const filteredAggregate = aggregateDataRow.filter(function (
          value,
          index
        ) {
          if (value === "") {
            emptyColsIndex.push(index);
          }
          return value !== "";
        });
        // columns that have empty cell will be removed
        const columnsFiltered = columns.filter(function (column, index) {
          return !emptyColsIndex.includes(index);
        });
        const transformedData = transformData(columnsFiltered, [
          filteredAggregate,
        ]);
        worksheetsData.push({
          worksheet: `Aggregate ${worksheetName}`,
          type: "static",
          isWorkshopTable: false,
          data: transformedData,
          columnsSummaryData: {},
        });
      }
    }
    // if not a workshop sheet, then get all data
    else {
      const transformedData = transformData(columns, values, worksheetType);
      worksheetsData.push({
        worksheet: worksheetName,
        type: worksheetType,
        isWorkshopTable: isWorkshopSheet,
        data: transformedData,
        columnsSummaryData: {},
      });
    }
  });

  return worksheetsData;
}

/**
 * do a calculation for the filtered data for each worksheet, and adds "aggregate" string one cell before aggregate data
 * (for the purpose of displaying a descriptive text on aggregate data row)
 * @param {string[]} columns - an array of column names
 * @param {string[]} aggregateDataRow - the first row that has calculation for the unfiltered data (aggregate data)
 * @param {string[][]} filteredRows - the data to process that has been filtered by id
 * @param {number[] | null} currencyColIndex - index of column where value should be converted into currency
 * @typedef {{string:string|number}} ColumnsSummary
 * @return {ColumnsSummary} an object with column name as key and its summary as value
 */
function getFilteredDataColumnSummary(
  columns,
  aggregateRows,
  filteredRows,
  currencyColIndex = null
) {
  let result = {};

  // record empty cell index (preferrably one cell before any value appear)
  let emptyCellIndex = -1;

  aggregateRows.forEach(function (aggregateDataValue, index) {
    // index will indicate the column
    // aggregateDataValue will be empty if its under a column that does not need calcs
    if (aggregateDataValue) {
      // if have not recorded then record the index
      // not using lasIndexOf because considering an empty value in the middle of the row
      if (emptyCellIndex < 0) {
        emptyCellIndex = index - 1;
      }

      if (aggregateDataValue !== "" && !isNaN(aggregateDataValue)) {
        aggregateRows[index] = aggregateDataValue.toFixed(2);
      }

      // only check if filteredRows is not empty
      if (filteredRows.length > 0) {
        // if value is not a number, do the custom calcs
        if (isNaN(aggregateDataValue)) {
          const columnSummary = calculateColumnNaN(index, filteredRows);
          result[columns[index]] = columnSummary;
        } else {
          let columnSummary = calculateColumnAverage(index, filteredRows);
          if (currencyColIndex && currencyColIndex.includes(index)) {
            columnSummary = new Intl.NumberFormat("en-US", {
              style: "currency",
              currency: "USD",
            }).format(columnSummary);
          }
          result[columns[index]] = columnSummary;
        }
      }
    }
  });

  //BE AWARE THIS MODIFY THE ORIGINAL ARRAY
  if (emptyCellIndex >= 0) {
    aggregateRows[emptyCellIndex] = "Aggregate";
  }

  return result;
}

/**
 * do the custom calculation for a column
 * @param {number} columnIndex - the index of the column
 * @param {string[][]} filteredRows - the data to process
 * @return {string} the result value of that column
 */
function calculateColumnNaN(columnIndex, filteredRows) {
  let count = 0;

  filteredRows.forEach(function (row) {
    const rowValue = row[columnIndex];
    // count the number of "Yes" in the column, or the number of cell that has value
    if (rowValue) {
      count += 1;
    }
  });

  const averagePrecent = (count / filteredRows.length) * 100;
  return `${count} (${averagePrecent.toFixed(2)}%)`;
}

function calculateColumnAverage(columnIndex, filteredRows) {
  let sum = 0;

  filteredRows.forEach(function (row) {
    const rowValue = row[columnIndex];
    if (rowValue) {
      sum += parseFloat(rowValue);
    }
  });

  const avg = sum / filteredRows.length;
  return avg.toFixed(2);
}

/**
 * transform data from
 * [col1, col2, col3, ...]
 * [row01, row02, row03, ...] <-- row 0
 * [row11, row12, row13, ...] <-- row 1
 *
 * to
 * {
 *  col1:row01
 *  col2:row02
 *  col3:row03
 * },
 * {
 *  col1:row11
 *  col2:row12
 *  col3:row13
 * }
 * @param {string[][]} values - 2D array of rows
 * @param {string} type - a type of workshop ignore first column which contains workshop Id
 */
function transformData(columns, values, type) {
  let result = [];
  const colStart = type === "workshop" ? 1 : 0;
  for (let row = 0; row < values.length; row++) {
    let resultObj = {};
    //TODO consied starting col from 1, since at index 0 its the workshop ID
    // static type does not have ID, start at col 0
    for (let col = colStart; col < columns.length; col++) {
      resultObj[columns[col]] = values[row][col];
    }
    result.push(resultObj);
  }
  return result;
}

function getWorksheetsList(sheet) {
  const worksheetSettingsSheetName = "Table Settings";
  const worksheetNameColumn = "Worksheet Name";
  const typeColumn = "Type";
  const chartTypeColumn = "Chart Type";
  const chartColumnsColumn = "Chart Columns";
  const chartGroupsColumn = "Chart Groups";

  var worksheetList = [];
  const ws = sheet.getSheetByName(worksheetSettingsSheetName);
  const lastColumn = ws.getLastColumn();
  const lastRow = ws.getLastRow();

  //https://stackoverflow.com/questions/17632165/determining-the-last-row-in-a-single-column
  const aVals = ws.getRange("A1:A" + lastRow).getValues();
  const aLast = lastRow - aVals.reverse().findIndex((c) => c[0] != "");

  const range = ws.getRange(1, 1, aLast, lastColumn);

  // range values will be a 2D array [column][row]
  const [columnNames, ...values] = range.getValues();
  const worksheetNameColumnPosition = columnNames.indexOf(worksheetNameColumn);
  const worksheetTypeColumnPosition = columnNames.indexOf(typeColumn);
  const chartTypeColumnPosition = columnNames.indexOf(chartTypeColumn);
  const chartColumnsColumnPosition = columnNames.indexOf(chartColumnsColumn);
  const chartGroupsColumnPosition = columnNames.indexOf(chartGroupsColumn);

  for (let i = 0; i < values.length; i++) {
    worksheetList.push({
      worksheetName: values[i][worksheetNameColumnPosition],
      type: values[i][worksheetTypeColumnPosition],
      chartType:
        chartTypeColumnPosition >= 0
          ? values[i][chartTypeColumnPosition]
          : null,
      chartColumns:
        chartColumnsColumnPosition >= 0
          ? values[i][chartColumnsColumnPosition]
          : null,
      chartGroups:
        chartGroupsColumnPosition >= 0
          ? values[i][chartGroupsColumnPosition]
          : null,
    });
  }

  Logger.log(worksheetList);
  return worksheetList;
}

function getWorkshopList(sheet) {
  const workshopListWorkSheetName = "Workshop Planner";
  const workshopIdColumn = "ID";
  const workshopDisplayNameColumn = "Display Name";

  var workshopList = [];
  // get value from column with value of id

  const ws = sheet.getSheetByName(workshopListWorkSheetName);
  const lastColumn = ws.getLastColumn();
  const lastRow = ws.getLastRow();

  //https://stackoverflow.com/questions/17632165/determining-the-last-row-in-a-single-column
  const aVals = ws.getRange("A1:A" + lastRow).getValues();
  const aLast = lastRow - aVals.reverse().findIndex((c) => c[0] != "");

  const range = ws.getRange(1, 1, aLast, lastColumn);

  // range values will be a 2D array [column][row]
  const [columnNames, ...values] = range.getValues();
  const workshopIdPosition = columnNames.indexOf(workshopIdColumn);
  const workshopDisplayNameColumnPosition = columnNames.indexOf(
    workshopDisplayNameColumn
  );
  for (let i = 0; i < values.length; i++) {
    workshopList.push({
      value: values[i][workshopIdPosition],
      label: values[i][workshopDisplayNameColumnPosition],
    });
  }

  return workshopList;
}

function getPartnerList(sheet) {
  const partnerListWorkSheetName = "Partner List";
  const partnerIdColumn = "ID";
  const partnerDisplayNameColumn = "Full Name";

  var partnerList = [];
  // get value from column with value of id

  const ws = sheet.getSheetByName(partnerListWorkSheetName);
  const lastColumn = ws.getLastColumn();
  const lastRow = ws.getLastRow();

  //https://stackoverflow.com/questions/17632165/determining-the-last-row-in-a-single-column
  const aVals = ws.getRange("A1:A" + lastRow).getValues();
  const aLast = lastRow - aVals.reverse().findIndex((c) => c[0] != "");

  const range = ws.getRange(1, 1, aLast, lastColumn);

  // range values will be a 2D array [column][row]
  const [columnNames, ...values] = range.getValues();
  const partnerIdPosition = columnNames.indexOf(partnerIdColumn);
  const partnerDisplayNameColumnPosition = columnNames.indexOf(
    partnerDisplayNameColumn
  );
  for (let i = 0; i < values.length; i++) {
    partnerList.push({
      value: values[i][partnerIdPosition],
      label: values[i][partnerDisplayNameColumnPosition],
    });
  }

  return partnerList;
}

/**
 * converts cell values into currency
 * @param {string[][] | {string:string}} data - data to process, could be a 2D array or object of key value pair
 * @param {Number[]} colIndex - array containing the column's index which the value should be formatted
 */
function convertToCurrency(data, currencyColIndex = null) {
  const usCurrency = new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  });

  if (Array.isArray(data)) {
    data.forEach(function (row, rowIndex, array) {
      row.forEach(function (cell, colIndex) {
        if (currencyColIndex && currencyColIndex.includes(colIndex)) {
          array[rowIndex][colIndex] = usCurrency.format(parseInt(cell));
        }
      });
    });
  }

  return;
}

/**
 * Generates a signature for a worksheet's columns
 * @param {Sheet} sheet - The worksheet to generate signature for
 * @param {string[]} columnNames - The names of columns to include in the signature
 * @return {string} JSON string representing column structure
 */
function generateColumnSignature(sheet, columnNames = []) {
  // Get column headers

  if (columnNames.length === 0) {
    columnNames = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  const signature = columnNames.filter((col) => col !== "");

  // Create a string representation (e.g., column names joined with a delimiter)
  return JSON.stringify(signature);
}

/**
 * Stores column signatures for all relevant worksheets
 * @param {Spreadsheet} sheet - The active spreadsheet
 * @return {Object} The generated signatures object
 */
function storeColumnSignatures(sheet) {
  const worksheetData = getWorksheetsList(sheet);
  const signatures = {};

  worksheetData.forEach((worksheetInfo) => {
    const worksheetName = worksheetInfo.worksheetName;
    const ws = sheet.getSheetByName(worksheetName);

    if (ws) {
      signatures[worksheetName] = generateColumnSignature(ws);
    }
  });

  // Store in document properties (associated with this spreadsheet)
  PropertiesService.getDocumentProperties().setProperty(
    "columnSignatures",
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
    const signaturesJson =
      PropertiesService.getDocumentProperties().getProperty("columnSignatures");
    return signaturesJson ? JSON.parse(signaturesJson) : {};
  } catch (e) {
    console.error("Error retrieving column signatures:", e);
    return {};
  }
}

/**
 * Checks if worksheet columns have changed since signatures were last stored
 * @param {Spreadsheet} sheet - The active spreadsheet
 * @return {Array} Names of worksheets with column changes
 */
function detectColumnChanges(sheet) {
  const worksheetData = getWorksheetsList(sheet);
  const storedSignatures = getStoredColumnSignatures();
  const changedWorksheets = [];

  worksheetData.forEach((worksheetInfo) => {
    const worksheetName = worksheetInfo.worksheetName;
    const ws = sheet.getSheetByName(worksheetName);

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
 * Sets the header message in the selector sheet
 * @param {SpreadsheetApp.Sheet} selectorSheet - The selector sheet to update
 * @param {string} [actionType="generated"] - Action description ("generated", "updated", etc.)
 */
function setColumnSelectorHeader(selectorSheet, actionType = "generated") {
  selectorSheet.getRange("A1:D1").merge();
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

/**
 * Resets stored column signatures (for troubleshooting)
 */
function resetColumnSignatures() {
  PropertiesService.getDocumentProperties().deleteProperty("columnSignatures");
  SpreadsheetApp.getUi().alert(
    "Signatures Reset",
    "Column signatures have been reset. Next checks will establish new baselines.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function captureColumnSelections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectorSheet = ss.getSheetByName("Column Selector");
  if (!selectorSheet) return {};

  const lastRow = selectorSheet.getLastRow();
  // Skip header rows
  if (lastRow <= 2) return {};

  // Get all data in one operation to minimize API calls
  const data = selectorSheet.getRange(3, 1, lastRow - 2, 4).getValues();

  const selections = {};

  data.forEach((row) => {
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
  PropertiesService.getDocumentProperties().setProperty(
    "columnSelections",
    JSON.stringify(selections)
  );

  return selections;
}

function restoreColumnSelections(savedSelections = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectorSheet = ss.getSheetByName("Column Selector");
  if (!selectorSheet) return;

  // If no selections passed, try to get from document properties
  if (!savedSelections) {
    try {
      const selectionsJson =
        PropertiesService.getDocumentProperties().getProperty(
          "columnSelections"
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

  data.forEach((row, index) => {
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
  groupUpdates.forEach((update) => {
    selectorSheet.getRange(update.row, 4).setValue(update.value);
  });
}

/**
 * Generates chart data for a worksheet based on chart settings
 * @param {Object} worksheetData - The processed worksheet data object
 * @param {string} chartType - Type of chart (pie, bar, etc.)
 * @param {string} chartGroupsJSON - JSON string containing column groupings
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
    return null;
  }

  let chartGroups;
  try {
    chartGroups = JSON.parse(chartGroupsJSON);
  } catch (e) {
    console.error("Error parsing chart groups:", e);
    return null;
  }

  // Initialize chart data structure
  const chartData = {
    type: chartType || "bar",
    data: [],
  };

  // Get data sources
  const aggregateRow = worksheetData.data[0];
  const workshopData = worksheetData.columnsSummaryData;

  // Process each chart group
  Object.keys(chartGroups).forEach((groupName) => {
    const columns = chartGroups[groupName];
    if (!columns || columns.length === 0) return;

    // Create data structure for this group
    const groupData = {
      labels: columns, // Column names become labels
      datasets: [
        {
          label: "Aggregate",
          data: [],
        },
        {
          label: workshopId, // Will be replaced with workshop name in frontend
          data: [],
        },
      ],
    };

    // Extract data for each column
    columns.forEach((column) => {
      // Check if column exists in data sources
      if (
        aggregateRow.hasOwnProperty(column) &&
        workshopData.hasOwnProperty(column)
      ) {
        // Add aggregate data
        groupData.datasets[0].data.push(
          extractNumericValue(aggregateRow[column])
        );

        // Add workshop-specific data
        groupData.datasets[1].data.push(
          extractNumericValue(workshopData[column])
        );
      } else {
        // Push 0 for missing columns
        groupData.datasets[0].data.push(0);
        groupData.datasets[1].data.push(0);
        console.warn(`Column "${column}" not found in dataset`);
      }
    });

    chartData.data.push(groupData);
  });

  return chartData;
}

/**
 * Extracts numeric value from different data formats
 * @param {any} value - The value to extract number from
 * @return {number} The extracted numeric value
 */
function extractNumericValue(value) {
  // Handle numeric values directly
  if (typeof value === "number") return value;
  if (!value) return 0;

  // Handle string values
  if (typeof value === "string") {
    // Format: "X (Y%)" - extract X
    const countMatch = value.match(/^(\d+)/);
    if (countMatch) return parseFloat(countMatch[1]);

    // Format: "$X,XXX.XX" - extract X,XXX.XX
    const currencyMatch = value.match(/[\$]?([\d,]+(\.\d+)?)/);
    if (currencyMatch) {
      return parseFloat(currencyMatch[1].replace(/,/g, ""));
    }

    // Try direct number parsing
    const numValue = parseFloat(value);
    if (!isNaN(numValue)) return numValue;
  }

  return 0;
}
