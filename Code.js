const getWorkshopListAction = "getWorkshops";
const getWorkshopResultsAction = "getWorkshopResults";
const getPartnerListAction = "getPartners";

//name of worskheet where the list of workshop exist
const workshopListWorkSheetName = "Workshop Planner";
//const worksheets = ["Workshop Result TEST", "Workshop Result TEST 2"]

var retryGetValueTimes = 3;

function doGet(request) {
  console.log(request);

  const params = request?.parameter;

  Logger.log(params);
  let responseValue = {};

  // sheet variable is of type Sheet https://developers.google.com/apps-script/reference/spreadsheet/sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  const action = params?.action;
  switch (action) {
    case getWorkshopListAction:
      // will return a [workshopID_1, workshopID_2, ...]
      const resultArr = getWorkshopList(sheet);
      responseValue = { data: resultArr };
      break;
    case getWorkshopResultsAction:
      try {
        const id = params.workshop_id;
        const worksheetsList = getWorksheetsList(sheet);
        const resultArr = getWorksheetsData(sheet, id, worksheetsList);
        responseValue = { data: resultArr };
      } catch (e) {
        responseValue = { error: e.message };
      }
      break;
    case getPartnerListAction:
      const resultPartnerArr = getPartnerList(sheet);
      responseValue = { data: resultPartnerArr };
      break;
    default:
      const error = new Error(
        "action does not exist or action parameter missing"
      );
      responseValue = { error: error.message };
  }

  /** 
  if(id) {
    id = decodeURIComponent(id)
    try{
      resultArr = retry_(getReturnValue, retryGetValueTimes, id, console.log)
      responseValue = {data:resultArr}
    }catch(e){
      responseValue = {error: e.message}
    }
  }else{
    responseValue = {error: 'missing id'}
  }
*/

  return ContentService.createTextOutput(
    JSON.stringify(responseValue)
  ).setMimeType(ContentService.MimeType.JSON);
}

function getReturnValue(id) {
  var resultArr = [];
  // get value from column with value of id
  // sheet variable is of type Sheet https://developers.google.com/apps-script/reference/spreadsheet/sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    workshopListWorkSheetName
  );
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, 1, lastRow, lastColumn);

  // range values will be a 2D array [column][row]
  var [columnNames, ...values] = range.getValues();
  Logger.log(range.getValues());
  // get the index for the column that we want
  var idColumnPosition = columnNames.indexOf(idColumnName);

  var latestValueRow = -1;

  // find result with specified quiz_id
  // assume newest record is at the end
  // loop from the end, if duplicate found by email then it must be old record = ignore
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][idColumnPosition]) === id) {
      const email = values[i][emailColumnPosition];
      const assessmentResult = values[i][valueColumnPosition];
      if (!resultArr.find((result) => result?.email === email)) {
        resultArr.push({ email: email, assessmentResult: assessmentResult });
      }
    }
  }

  /*
  if(latestValueRow >= 0){
    var encodedUri = encodeURIComponent(values[latestValueRow][valueColumnPosition])
    result = {value: encodedUri}
  }
*/
  if (resultArr.length == 0) {
    throw new Error("That id does not exist");
  }

  console.log(resultArr);
  return resultArr;
}

/**
 * Retry with truncated exponential backoff
 * Mashup of https://gist.github.com/peterherrmann/2700284 exponential backoff and partials
 * @param {function} the function to retry
 * @param {numeric} number of times to retry before throwing an error
 * @param {string} optional function arguments
 * @param {function} optional logging function
 * @return {function applied} response of function
 * @private
 */
function retry_(func, numRetries, optKey, optLoggerFunction) {
  for (var n = 0; n <= numRetries; n++) {
    try {
      if (optKey) {
        var response = func(optKey);
      } else {
        var response = func();
      }
      return response;
    } catch (e) {
      if (optLoggerFunction) {
        optLoggerFunction("Retry " + n + ": " + e);
      }
      if (n == numRetries) {
        throw e;
      }
      Utilities.sleep(Math.pow(2, n) * 1000 + Math.round(Math.random() * 1000));
    }
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("MOAA Chart Tools")
    .addItem("Generate Column Selector", "generateColumnSelector")
    .addItem("Apply Column Selector", "applyColumnSelections")
    .addToUi();
}

/**
 * Generates a Column Selector worksheet with all columns from relevant worksheets
 */
function generateColumnSelector() {
  const signatures = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const worksheetData = getWorksheetsList(ss);

  // Create or clear the Column Selector sheet
  let selectorSheet = ss.getSheetByName("Column Selector");
  if (!selectorSheet) {
    selectorSheet = ss.insertSheet("Column Selector");
  } else {
    selectorSheet.clear();
  }

  // Prepare all data in memory first
  const allData = [];
  const checkboxRanges = [];
  const validationRanges = [];

  allData.push([
    "Select columns to include in charts and assign them to groups. Then use 'Chart Tools > Apply Column Selections' to update settings.",
    "",
    "",
    "",
  ]);

  // Add header row
  allData.push([
    "Worksheet Name",
    "Column Name",
    "Include in Chart",
    "Chart Group",
  ]);

  let dataRowIndex = 2; // Starting index for data rows (after header)

  worksheetData.forEach((worksheet) => {
    const worksheetName = worksheet.worksheetName;
    const sheet = ss.getSheetByName(worksheetName);
    if (!sheet) return;

    // Get column names
    const columnNames = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    let addedRowsForWorksheet = false;

    columnNames.forEach((columnName, index) => {
      if (columnName && index > 0) {
        // Skip first column assuming it's ID
        // Add data row
        allData.push([worksheetName, columnName, "", ""]);

        // Track checkbox and validation ranges
        // +2 because we have instruction row and header row
        checkboxRanges.push(dataRowIndex + 2);
        validationRanges.push(dataRowIndex + 2);

        dataRowIndex++;
        addedRowsForWorksheet = true;
      }
    });

    // Add blank row between worksheets for readability if we added rows
    if (addedRowsForWorksheet) {
      allData.push(["", "", "", ""]);
      dataRowIndex++;
    }
  });

  // Write all data at once - more efficient than individual writes
  if (allData.length > 0) {
    selectorSheet.getRange(1, 1, allData.length, 4).setValues(allData);
  }

  // Apply merged cell for instructions
  selectorSheet.getRange("A1:D1").merge().setFontStyle("italic");

  // Format header row
  selectorSheet.getRange("A2:D2").setFontWeight("bold");

  // Apply checkboxes for all "Include in Chart" cells in column C
  // Note: We need to batch these by continuous ranges if possible
  if (checkboxRanges.length > 0) {
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

  if (validationRanges.length > 0) {
    const groups = ["Group 1", "Group 2", "Group 3", "Group 4", "Group 5"];
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

  // Final formatting
  selectorSheet.setFrozenRows(2); // Freeze both instruction and header rows
  selectorSheet.autoResizeColumns(1, 4);
}

/**
 * Reads selections from Column Selector and updates Table Settings
 */
function applyColumnSelections() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectorSheet = ss.getSheetByName("Column Selector");
  const tableSettings = ss.getSheetByName("Table Settings");

  if (!selectorSheet || !tableSettings) {
    SpreadsheetApp.getUi().alert(
      "Column Selector or Table Settings sheet not found"
    );
    return;
  }

  // Get all data from the selector sheet
  const lastRow = selectorSheet.getLastRow();
  const selectorData = selectorSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // Organize selections by worksheet
  const worksheetSelections = {};

  selectorData.forEach((row) => {
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

  updateTableSettings(tableSettings, worksheetSelections);
  SpreadsheetApp.getUi().alert("Column selections applied to Table Settings");
}

/**
 * Updates the Table Settings worksheet with column selections
 */
function updateTableSettings(tableSettingsSheet, worksheetSelections) {
  const lastRow = tableSettingsSheet.getLastRow();
  const lastColumn = tableSettingsSheet.getLastColumn();

  // Get current table settings
  const data = tableSettingsSheet
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();
  const headers = data[0];

  // Find column indexes
  const worksheetNameColIdx = headers.indexOf("Worksheet Name");
  const chartColumnsColIdx = headers.indexOf("Chart Columns");
  const chartGroupsColIdx = headers.indexOf("Chart Groups");
  const chartTypeColIdx = headers.indexOf("Chart Type");

  // Ensure required columns exist
  if (worksheetNameColIdx === -1) {
    // Need to add Chart Columns and Chart Groups columns if they don't exist
    if (chartColumnsColIdx === -1) {
      tableSettingsSheet.getRange(1, lastColumn + 1).setValue("Chart Columns");
      chartColumnsColIdx = lastColumn;
      lastColumn++;
    }
    if (chartGroupsColIdx === -1) {
      tableSettingsSheet.getRange(1, lastColumn + 1).setValue("Chart Groups");
      chartGroupsColIdx = lastColumn;
    }
  }

  // Update each worksheet's chart settings
  for (let i = 1; i < data.length; i++) {
    const worksheetName = data[i][worksheetNameColIdx];

    if (worksheetName && worksheetSelections[worksheetName]) {
      // Get the selections for this worksheet
      const selections = worksheetSelections[worksheetName];
      const groups = Object.keys(selections);

      // create chart columns string - all selected columns combined
      const allColumns = [];
      groups.forEach((group) => {
        selections[group].forEach((selectedColumnName) => {
          allColumns.push(selectedColumnName);
        });
      });

      // Create groups object for JSON storage
      const groupsObject = {};
      groups.forEach((group) => {
        groupsObject[group] = selections[group];
      });

      tableSettingsSheet
        .getRange(i + 1, chartColumnsColIdx + 1)
        .setValue(JSON.stringify(allColumns));
      tableSettingsSheet
        .getRange(i + 1, chartGroupsColIdx + 1)
        .setValue(JSON.stringify(groupsObject));

      // Make sure the worksheet has a chart type if one doesn't exist
      if (chartTypeColIdx !== -1 && !data[i][chartTypeColIdx]) {
        // Set default chart type based on number of groups
        const defaultChartType = groups.length > 1 ? "pie" : "bar";
        tableSettingsSheet
          .getRange(i + 1, chartTypeColIdx + 1)
          .setValue(defaultChartType);
      }
    }
  }
  // When reading back later:
}

function generateColumnSignature(sheet, columnNames = []) {
  // Get column headers

  if (columnNames.length === 0) {
    columnNames = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  const signature = columnNames.filter((col) => col !== "");

  // Create a string representation (e.g., column names joined with a delimiter)
  return JSON.stringify(signature);
}

function storeColumnSignatures(signatures) {
  PropertiesService.getScriptProperties().setProperty(
    "columnSignatures",
    JSON.stringify(signatures)
  );
}
