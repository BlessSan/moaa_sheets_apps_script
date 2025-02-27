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
}
