/**
 * Main Module for MOAA Chart Tools
 * Entry point for HTTP requests and application initialization
 */

// Action constants for HTTP requests
const ACTION_TYPES = {
  GET_WORKSHOPS: "getWorkshops",
  GET_WORKSHOP_RESULTS: "getWorkshopResults",
  GET_PARTNERS: "getPartners",
};

/**
 * Handles HTTP GET requests to the web app
 * @param {Object} request - The HTTP request object
 * @return {GoogleAppsScript.Content.TextOutput} JSON response
 */
function doGet(request) {
  const params = request?.parameter;
  let responseValue = {};

  try {
    const action = params?.action;
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    switch (action) {
      case ACTION_TYPES.GET_WORKSHOPS:
        // Return list of available workshops
        const workshopList = DataAccess.getWorkshopList(spreadsheet);
        responseValue = { data: workshopList };
        break;

      case ACTION_TYPES.GET_WORKSHOP_RESULTS:
        // Return data for a specific workshop
        const workshopId = params.workshop_id;
        if (!workshopId) {
          throw new Error("Missing workshop_id parameter");
        }

        const worksheetsList = DataAccess.getWorksheetsList(spreadsheet);
        const worksheetsData = WorksheetService.getWorksheetsData(
          spreadsheet,
          workshopId,
          worksheetsList
        );

        responseValue = { data: worksheetsData };
        break;

      case ACTION_TYPES.GET_PARTNERS:
        // Return list of partners
        const partnerList = DataAccess.getPartnerList(spreadsheet);
        responseValue = { data: partnerList };
        break;

      default:
        throw new Error("Action does not exist or action parameter missing");
    }
  } catch (error) {
    // Log the error for debugging
    console.error("Error processing request:", error);
    responseValue = { error: error.message };
  }

  // Return the response as JSON
  return ContentService.createTextOutput(
    JSON.stringify(responseValue)
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Runs when the spreadsheet is opened, setting up the custom menu
 */
function onOpen() {
  UIManager.createCustomMenu();
}

/**
 * Global handler for retry operations with exponential backoff
 * @param {Function} func - Function to retry
 * @param {number} numRetries - Number of retry attempts
 * @param {any} optKey - Optional parameter to pass to the function
 * @param {Function} optLoggerFunction - Optional logging function
 * @return {any} Result of the function
 */
function retryWithBackoff(func, numRetries, optKey, optLoggerFunction) {
  for (let n = 0; n <= numRetries; n++) {
    try {
      return optKey ? func(optKey) : func();
    } catch (e) {
      if (optLoggerFunction) {
        optLoggerFunction(`Retry ${n}: ${e}`);
      }

      if (n === numRetries) {
        throw e;
      }

      // Exponential backoff with random jitter
      const delay = Math.pow(2, n) * 1000 + Math.round(Math.random() * 1000);
      Utilities.sleep(delay);
    }
  }
}
