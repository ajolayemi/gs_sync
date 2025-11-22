import {google, sheets_v4 as SheetsV4} from "googleapis";

import {logger} from "firebase-functions/v2";
import {SheetInfo, SheetSize} from "../interfaces/gs_interfaces";

/**
 * A wrapper to handle deletion request in google sheets
 * @param {number} sheetIdToDeleteFrom The unique id of the worksheet to
 * delete from. It's the "gid" value in a google worksheet url
 * @param {number} startIndex The row to start deleting from
 * @param {number} endIndex The row to end the deletion request
 * @param {string} spreadsheetId The unique id of the spreadsheet
 * @return {Promise<SheetsV4.Schema$BatchUpdateSpreadsheetResponse | null>}
 */
const deleteRowFromSheet = async (
  sheetIdToDeleteFrom: number,
  startIndex: number,
  endIndex: number,
  spreadsheetId: string
): Promise<SheetsV4.Schema$BatchUpdateSpreadsheetResponse | null> => {
  try {
    const sheets = getSpreadsheets();

    const deleteResponse = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId,

      requestBody: {
        requests: [
          {
            deleteDimension: {
              range: {
                sheetId: sheetIdToDeleteFrom,
                dimension: "ROWS",
                startIndex: startIndex,
                endIndex: endIndex,
              },
            },
          },
        ],
      },
    });
    if (deleteResponse.status == 200) {
      const deleteRequestResponseData = deleteResponse.data;
      logger.info(
        `Delete request done successfully: ${deleteRequestResponseData}`
      );
      return deleteRequestResponseData;
    }
    logger.error(`Request to delete the rows with index from 
    ${startIndex} to ${endIndex} in worksheet with id: ${sheetIdToDeleteFrom}
    failed with status ${deleteResponse.statusText}`);
    return null;
  } catch (error) {
    logger.error(`An error: ${error} occurred while trying
    to delete the rows with index from ${startIndex} to ${endIndex} in
    worksheet with id: ${sheetIdToDeleteFrom}`);
    return null;
  }
};

/**
 * A wrapper around google sheet api func to insert new values to google sheet
 * @param {string} worksheetName The name of the worksheet where the insertion
 * should be done
 * @param {Array<Array<string | number>>} dataToInsert The new values
 * to be inserted
 * @param {string} spreadsheetId The unique id of the spreadsheet
 * @return {Promise<SheetsV4.Schema$UpdateValuesResponse | null>}
 * The result of the insert request if it was successful or a null value
 */
const insertDataInSheet = async (
  worksheetName: string,
  dataToInsert: Array<Array<string | number>>,
  spreadsheetId: string
): Promise<SheetsV4.Schema$AppendValuesResponse | null> => {
  try {
    const sheets = getSpreadsheets();

    const insertResponse = await sheets.spreadsheets.values.append({
      spreadsheetId: spreadsheetId,
      range: worksheetName,
      valueInputOption: "RAW",
      requestBody: {
        values: dataToInsert,
      },
    });

    if (insertResponse.status == 200) {
      const responseData = insertResponse.data;
      logger.info(`Insert request done successfully: ${responseData.updates}`);
      return responseData;
    }
    logger.error(`The request to insert data in worksheet: ${worksheetName} 
    with provided values ${dataToInsert} failed with status
    ${insertResponse.statusText}`);
    return null;
  } catch (error) {
    logger.error(`An error: ${error} occurred while 
    trying to insert data in worksheet: ${worksheetName} 
    with provided values ${dataToInsert}`);
    return null;
  }
};

/**
 * A wrapper around google sheet api func to update values
 * @param {string} rangeToUpdate The range to be updated
 * Written in form <sheetName>!A1notation range
 * @param {Array<Array<string | number>>} updateRequestBodyValue The new values
 * to be set in the specified range
 * @param {string} spreadsheetId The unique id of the spreadsheet
 * @return {Promise<SheetsV4.Schema$UpdateValuesResponse | null>}
 * The result of the update request if it was successful or a null value
 */
const updateDataInSheet = async (
  rangeToUpdate: string,
  updateRequestBodyValue: Array<Array<string | number>>,
  spreadsheetId: string
): Promise<SheetsV4.Schema$UpdateValuesResponse | null> => {
  try {
    const sheets = getSpreadsheets();

    const updateResponse = await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: rangeToUpdate,
      valueInputOption: "RAW",
      requestBody: {
        values: updateRequestBodyValue,
      },
    });

    if (updateResponse.status == 200) {
      const responseData = updateResponse.data;
      const updatedCols = responseData.updatedColumns ?? 0;
      const updatedRows = responseData.updatedRows ?? 0;
      logger.info(
        `Update request done successfully with ${updatedRows} rows and ${updatedCols} columns updated`
      );
      return responseData;
    }
    logger.error(`The request to update the values in range: ${rangeToUpdate} 
    with provided values ${updateRequestBodyValue} failed with status
    ${updateResponse.statusText}`);
    return null;
  } catch (error) {
    logger.error(`An error: ${error} occurred while 
    trying to update the values in range: ${rangeToUpdate} 
    with provided values ${updateRequestBodyValue}`);
    return null;
  }
};

/**
 * A wrapper around google sheet api sheet function that
 * makes request to read data from spreadsheet.
 * It's so to avoid code duplication
 * @param {string} range  The range to get from spreadsheet.
 * The said range has to be an A1 notation range type
 * @param {string} spreadsheetId The unique id of the spreadsheet
 * @return {Promise<Array<Array<string | number>>>}
 * Returns either a nested list of the values read from spreadsheet or
 * an empty list if the reading wasn't successful
 */
const readDataFromSpreadsheet = async (
  range: string,
  spreadsheetId: string
): Promise<Array<Array<string | number>>> => {
  try {
    const sheet = getSpreadsheets();
    const readResponse = await sheet.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: range,
      valueRenderOption: "UNFORMATTED_VALUE",
    });

    // If request response status is 200
    if (readResponse.status == 200) {
      return readResponse.data.values ?? [];
    }
    logger.error(
      `The request to read data in range: ${range} in sheet with id: 
      ${spreadsheetId} failed with status ${readResponse.statusText}`
    );
    return [];
  } catch (error) {
    logger.error(
      `An error: ${error} occurred while trying to load range: 
      ${range} in sheet with id: ${spreadsheetId}`
    );
    throw error;
  }
};

/** *
 * Given a worksheet in a spreadsheet, it ensures it has the required
 * number of rows and columns. If not, it resizes it to the required size
 * @param {string} spreadsheetId The unique id of the spreadsheet
 * @param {number} sheetId The unique id of the worksheet in the spreadsheet
 * @param {number} requiredRowCount The required number of rows
 * @param {number} requiredColumnCount The required number of columns
 * @return {Promise<void>}
 */
const ensureSheetSize = async (
  spreadsheetId: string,
  sheetId: number,
  requiredRowCount: number,
  requiredColumnCount: number
): Promise<void> => {
  const {rowCount, columnCount} = await getSheetSize(spreadsheetId, sheetId);
  const requests: SheetsV4.Schema$Request[] = [];

  if (rowCount < requiredRowCount) {
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId: sheetId,
          gridProperties: {
            rowCount: requiredRowCount,
          },
        },
        fields: "gridProperties.rowCount",
      },
    });
  }

  if (columnCount < requiredColumnCount) {
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId: sheetId,
          gridProperties: {
            columnCount: requiredColumnCount,
          },
        },
        fields: "gridProperties.columnCount",
      },
    });
  }

  if (requests.length > 0) {
    const sheets = getSpreadsheets();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId,
      requestBody: {
        requests: requests,
      },
    });
    logger.info(
      `Resized sheet with id: ${sheetId} in spreadsheet with id: 
      ${spreadsheetId} to rows: ${requiredRowCount}, columns: ${requiredColumnCount}`
    );
  } else {
    logger.info(
      `Sheet with id: ${sheetId} in spreadsheet with id: 
      ${spreadsheetId} already has required size rows: ${requiredRowCount}, 
      columns: ${requiredColumnCount}`
    );
  }
};

/**
 * Given a spreadsheet and a particular worksheet in it, it returns the size
 * of the worksheet in terms of number of rows and columns
 * @param {string} spreadsheetId The unique id of the spreadsheet
 * @param {number} worksheetId The unique id of the worksheet in the spreadsheet
 * @return {Promise<SheetSize | null>}
 * Returns either an object with rowCount and columnCount or null
 * if the request wasn't successful
 */
const getSheetSize = async (
  spreadsheetId: string,
  worksheetId: number
): Promise<SheetSize> => {
  try {
    const sheets = getSpreadsheets();
    const res = await sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId,
      includeGridData: false,
    });

    const sheet = res.data.sheets?.find(
      (s) => s.properties?.sheetId === worksheetId
    );
    const grid = sheet?.properties?.gridProperties;

    logger.info(`Check destination has data; ${sheet?.data}`)

    return {
      rowCount: grid?.rowCount ?? 0,
      columnCount: grid?.columnCount ?? 0,
    };
  } catch (error) {
    logger.error(
      `An error: ${error} occurred while trying to get the size of 
      worksheet with id: ${worksheetId} in spreadsheet with id: ${spreadsheetId}`
    );
    return {
      rowCount: 0,
      columnCount: 0,
    };
  }
};

/**
 * Handles the logic behind spreadsheet authorization
 * @return {SheetsV4.Sheets} The loaded spreadsheet
 */
const getSpreadsheets = (): SheetsV4.Sheets => {
  const auth = new google.auth.GoogleAuth({
    keyFilename: "keys/gs-sync-dev.json",
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  const options: SheetsV4.Options = {auth: auth, version: "v4"};
  return google.sheets(options);
};

/**
 * A class implementation of sheet info
 */
class SheetInfoImpl {
  /**
   * @param {SheetInfo} sheetInfo An interface of information pertaining to a
   * Google Worksheet
   * @return {string} returns an A1 notation of a google sheet info
   * This info is mainly used when reading data from Google Worksheet
   */
  static getReadA1NotationRange(sheetInfo: SheetInfo): string {
    const firstRow = sheetInfo.firstRow ?? "";
    const lastRow = sheetInfo.lastRow ?? "";
    return `${sheetInfo.name}!${sheetInfo.firstColumn}${firstRow}:${sheetInfo.lastColumn}${lastRow}`;
  }
}

export {
  SheetInfoImpl,
  getSheetSize,
  getSpreadsheets,
  readDataFromSpreadsheet,
  insertDataInSheet,
  updateDataInSheet,
  deleteRowFromSheet,
  ensureSheetSize,
};
