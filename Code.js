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
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const worksheetData = getWorksheetsList(ss);

  // Create or clear the Column Selector sheet
  let selectorSheet = ss.getSheetByName("Column Selector");
  if (!selectorSheet) {
    selectorSheet = ss.insertSheet("Column Selector");
  } else {
    selectorSheet.clear();
  }

  // Set up header row
  selectorSheet
    .getRange("A1:D1")
    .setValues([
      ["Worksheet Name", "Column Name", "Include in Chart", "Chart Group"],
    ]);
  selectorSheet.getRange("A1:D1").setFontWeight("bold");

  let rowIndex = 2;
  worksheetData.forEach((worksheet) => {
    const worksheetName = worksheet.worksheetName;
    const sheet = ss.getSheetByName(worksheetName);
    if (!sheet) return;

    // Get column names (assuming they're in row 1)
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const columnNames = headerRange.getValues()[0];

    columnNames.forEach((columnName, index) => {
      if (columnName && index >> 0) {
        // Skip first column assuming it's ID

        // Add worksheet name, column name, and checkboxes
        selectorSheet
          .getRange(rowIndex, 1, 1, 2)
          .setValues([[worksheetName, columnName]]);

        // Add checkbox for "Include in Chart"
        selectorSheet.getRange(rowIndex, 3).insertCheckboxes();

        // Add dropdown for "Chart Group"
        const groups = ["Group 1", "Group 2", "Group 3", "Group 4", "Group 5"];
        const validation = SpreadsheetApp.newDataValidation()
          .requireValueInList(groups, true)
          .build();
        selectorSheet.getRange(rowIndex, 4).setDataValidation(validation);

        rowIndex++;
      }
    });

    // Add a blank row between worksheets for readability
    selectorSheet.getRange(rowIndex, 1).setValue("");
    rowIndex++;
  });

  // Format the sheet
  selectorSheet.setFrozenRows(1);
  selectorSheet.autoResizeColumns(1, 4);

  // Add instructions at the top
  selectorSheet.insertRowBefore(1);
  selectorSheet.getRange("A1:D1").merge();
  selectorSheet
    .getRange("A1")
    .setValue(
      "Select columns to include in charts and assign them to groups. Then use 'Chart Tools > Apply Column Selections' to update settings."
    );
  selectorSheet.getRange("A1").setFontStyle("italic");
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
