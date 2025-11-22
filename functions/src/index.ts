import {setGlobalOptions, https, logger} from "firebase-functions";
import {
  getSheetSize,
  getSpreadsheets,
  readDataFromSpreadsheet,
  SheetInfoImpl,
  updateDataInSheet,
} from "./utils/gs_utils";
import {SyncDataRequestBody} from "./interfaces/request";
import {sheets_v4 as SheetV4} from "googleapis/build/src/apis/sheets/v4";

setGlobalOptions({maxInstances: 10});

/**
 * Cloud function that helps to sync data between two google sheets
 * using google sheets api
 *
 */
exports.gsSyncFunction = https.onRequest(async (req, res) => {
  const logTag = {TAG: "gsSyncFunctionLog"};
  try {
    const originalBody = req.body;
    logger.info(logTag, "Original request body:", originalBody);
    const body: SyncDataRequestBody = Buffer.isBuffer(originalBody)
      ? JSON.parse(Buffer.from(originalBody).toString())
      : originalBody;

    logger.info(logTag, `Parsed request body: ${JSON.stringify(body)}`);
    logger.info(logTag, `Origin Sheet ID: ${body.originWorksheetId}`);

    const readRange = SheetInfoImpl.getReadA1NotationRange({
      name: body.originWorksheetName,
      id: body.originWorksheetId,
      firstColumn: body.originWorksheetFirstColumn,
      lastColumn: body.originWorksheetLastColumn,
      firstRow: body.originWorksheetFirstRow,
      lastRow: body.originWorksheetLastRow,
    });

    if (!body.originSpreadsheetId || !body.destinationSpreadsheetId) {
      logger.error(
        logTag,
        "No originSpreadsheetId or destinationSpreadsheetId provided in the request body."
      );
      res
        .status(400)
        .send(
          "Bad Request: Missing originSpreadsheetId or destinationSpreadsheetId"
        );
      return;
    }

    logger.info(`range from which data is read from origin: ${readRange}`);

    // 1. Read data from the origin google sheet
    const dataFromGs = await readDataFromSpreadsheet(
      readRange,
      body.originSpreadsheetId
    );
    logger.info(logTag, `Data from Origin Sheet : ${dataFromGs.length}`);

    const sheets = getSpreadsheets();
    const rowCount = dataFromGs.length;
    const colCount = dataFromGs[0]?.length ?? 0;

    logger.info(
      logTag,
      `Building batch updates for ${rowCount} rows and ${colCount} columns.`
    );

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

    const batchUpdates: SheetV4.Schema$Request[] = [
      // 1️⃣ Clear all existing values only (keep formatting)
      {
        updateCells: {
          range: {
            sheetId: body.destinationWorksheetId,
            startRowIndex: 0,
            startColumnIndex: 0,
            endColumnIndex: colCount,
          },
          fields: "userEnteredValue", // clear only values
        },
      },
    ];

    if (!hasEnoughRows) {
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

    // Apply initial batch updates to clear and resize
    logger.info(
      logTag,
      "Applying initial batch updates to clear and eventually resize the destination sheet."
    );
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: body.destinationSpreadsheetId,
      requestBody: {
        requests: batchUpdates,
      },
    });

    const aiNotation = `${body.destinationWorksheetName}!${body.originWorksheetFirstColumn}:${body.originWorksheetLastColumn}`;

    logger.info(
      logTag,
      `Updating data at AI Notation: ${aiNotation} with ${dataFromGs.length} rows in destination spreadsheet ${body.destinationSpreadsheetName}.`
    );
    await updateDataInSheet(
      aiNotation,
      dataFromGs,
      body.destinationSpreadsheetId
    );

    res.send("Completed");
  } catch (error) {
    logger.error(logTag, "An error occurred in gsSyncFunction:", error);
    res.status(500).send("Internal Server Error");
  }
});
