import {logger} from "firebase-functions";
import {sheets_v4 as SheetsV4} from "googleapis";
import {getSheetSize, getSpreadsheets} from "./gs_utils";
import {SyncDataRequestBody} from "../interfaces/request";

/**
 * Applies batch updates to the destination Google Sheet
 * @param {sheetsV4.Schema$Request[]} requests - The batch update requests
 * @param {SyncDataRequestBody} body - The request body containing worksheet info
 * @return {Promise<void>}
 */
const applyBatchUpdates = async (
  requests: SheetsV4.Schema$Request[],
  body: SyncDataRequestBody
): Promise<void> => {
  const logTag = {TAG: "applyBatchUpdatesLog"};

  if (requests.length === 0) {
    logger.info(logTag, "No batch update requests to apply. Exiting function.");
    return;
  }
  // eslint-disable-next-line max-len
  logger.info(
    logTag,
    `Applying batch updates with requests to destination sheet with name : ${body.destinationWorksheetName} from spreadsheet: ${body.destinationSpreadsheetName}`
  );
  const sheets = getSpreadsheets();

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: body.destinationSpreadsheetId,
    requestBody: {
      requests: requests,
    },
  });
  // eslint-disable-next-line max-len
  logger.info(
    logTag,
    `Batch update response sent to destination sheet with name : ${body.destinationWorksheetName} from spreadsheet: ${body.destinationSpreadsheetName}`
  );
};

const buildBatchUpdates = async (
  body: SyncDataRequestBody,
  rowCount: number,
  colCount: number
): Promise<SheetsV4.Schema$Request[]> => {
  const logTag = {TAG: "buildBatchUpdatesLog"};
  // Get the size of the destination sheet
  const sheetSize = await getSheetSize(
    body.destinationSpreadsheetId,
    body.destinationWorksheetId
  );

  const hasEnoughRows = sheetSize.rowCount >= rowCount;
  const hasEnoughCols = sheetSize.columnCount >= colCount;

  logger.info(
    logTag,
    `Current destination sheet size: ${JSON.stringify(sheetSize)}`
  );
  logger.info("Building initial batch updates...");

  const batchUpdates: SheetsV4.Schema$Request[] = [];

  if (!hasEnoughRows) {
    // eslint-disable-next-line max-len
    logger.info(
      logTag,
      `Destination sheet has insufficient rows. Current: ${sheetSize.rowCount}, Required: ${rowCount}, add batch update to resize.`
    );
    batchUpdates.push({
      updateSheetProperties: {
        properties: {
          sheetId: body.destinationWorksheetId,
          gridProperties: {
            rowCount: rowCount,
          },
        },
        fields: "gridProperties(rowCount)",
      },
    });
  }

  if (!hasEnoughCols) {
    // eslint-disable-next-line max-len
    logger.info(
      logTag,
      `Destination sheet has insufficient columns. Current: ${sheetSize.columnCount}, Required: ${colCount}, add batch update to resize.`
    );
    batchUpdates.push({
      updateSheetProperties: {
        properties: {
          sheetId: body.destinationWorksheetId,
          gridProperties: {
            columnCount: colCount,
          },
        },
        fields: "gridProperties(columnCount)",
      },
    });
  }

  if (hasEnoughCols && hasEnoughRows) {
    logger.info(
      logTag,
      "Destination sheet has sufficient rows and columns. No resize needed."
    );
  }

  return batchUpdates;
};

export {buildBatchUpdates, applyBatchUpdates};
