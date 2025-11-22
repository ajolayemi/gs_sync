import {setGlobalOptions, https, logger} from "firebase-functions";
import {
  readDataFromSpreadsheet,
  SheetInfoImpl,
  updateDataInSheet,
} from "./utils/gs_utils";
import {SyncDataRequestBody} from "./interfaces/request";
import {sheets_v4 as SheetV4} from "googleapis/build/src/apis/sheets/v4";
import {applyBatchUpdates, buildBatchUpdates} from "./utils/data_utils";
import {generateHash} from "./utils/hash_utils";

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

    const originReadRange = SheetInfoImpl.getReadA1NotationRange({
      name: body.originWorksheetName,
      id: body.originWorksheetId,
      firstColumn: body.originWorksheetFirstColumn,
      lastColumn: body.originWorksheetLastColumn,
      firstRow: body.originWorksheetFirstRow,
      lastRow: body.originWorksheetLastRow,
    });

    const destinationReadRange = SheetInfoImpl.getReadA1NotationRange({
      name: body.destinationWorksheetName,
      id: body.destinationWorksheetId,
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

    logger.info(
      `range from which data is read from origin: ${originReadRange}`
    );

    // 1. Read data from the origin google sheet
    const dataFromOrigin = await readDataFromSpreadsheet(
      originReadRange,
      body.originSpreadsheetId
    );
    logger.info(
      logTag,
      `Data from origin sheet has length: ${dataFromOrigin.length}`
    );
    const hashedDataFromOrigin = await generateHash(dataFromOrigin);
    logger.info(logTag, `Hashed data from origin: ${hashedDataFromOrigin}`);

    // 2. Fetch data from destination sheet to determine if data needs updating
    const dataFromDestination = await readDataFromSpreadsheet(
      destinationReadRange,
      body.destinationSpreadsheetId
    );
    logger.info(
      logTag,
      `Data from destination sheet has length: ${dataFromDestination.length}`
    );
    const hashedDataFromDestination = await generateHash(dataFromDestination);
    logger.info(
      logTag,
      `Hashed data from destination: ${hashedDataFromDestination}`
    );

    // 3. Compare hashes to determine if update is needed
    if (hashedDataFromOrigin === hashedDataFromDestination) {
      logger.info(
        logTag,
        "Data in destination sheet is up-to-date. No update needed."
      );
      res.send("No update needed; data is already synchronized.");
      return;
    }

    logger.info(logTag, "Data mismatch detected; proceeding with update.");

    const rowCount = dataFromOrigin.length;
    const colCount = dataFromOrigin[0]?.length ?? 0;

    logger.info(
      logTag,
      `Building batch updates for ${rowCount} rows and ${colCount} columns.`
    );

    const batchUpdates: SheetV4.Schema$Request[] = await buildBatchUpdates(
      body,
      rowCount,
      colCount
    );
    // Apply initial batch updates to clear and resize
    logger.info(
      logTag,
      "Applying initial batch updates to clear and eventually resize the destination sheet."
    );

    // When data differs and the destination sheet has no data to begin with
    // execute batch updates to resize sheet if necessary and then do a complete data update
    if (dataFromDestination.length === 0) {
      logger.info(
        logTag,
        "Destination sheet is empty; applying batch updates for eventual resize."
      );
      await applyBatchUpdates(batchUpdates, body);

      logger.info(
        logTag,
        `Updating data at range: ${destinationReadRange} with ${dataFromOrigin.length} rows in destination spreadsheet ${body.destinationSpreadsheetName}.`
      );
      await updateDataInSheet(
        destinationReadRange,
        dataFromOrigin,
        body.destinationSpreadsheetId
      );
    }
    // Otherwise, when data differs, use row-wise hash to check which rows need updating
    // and only update those rows
    else {
      logger.info(
        logTag,
        "Destination sheet has existing data; determining row-wise updates."
      );

      const dataFromOriginLen = dataFromOrigin.length;
      const dataFromDestinationLen = dataFromDestination.length;

      const expectedRowCount =
        dataFromOriginLen > dataFromDestinationLen
          ? dataFromOriginLen
          : dataFromDestinationLen;

      const rowsToUpdate: number[] = [];
      for (let i = 0; i < expectedRowCount; i++) {
        const originRow = dataFromOrigin[i] || [];
        const destinationRow = dataFromDestination[i] || [];
        const originRowHash = await generateHash([originRow]);
        const destinationRowHash = await generateHash([destinationRow]);

        if (originRowHash !== destinationRowHash) {
          rowsToUpdate.push(i);
        }
      }

      logger.info(logTag, `Rows identified for update: ${rowsToUpdate.length}`);

      // If there are rows to update, ensure the sheet is resized appropriately first
      if (rowsToUpdate.length) {
        logger.info(
          logTag,
          "applying batch updates for eventual resize before updating rows."
        );
        await applyBatchUpdates(batchUpdates, body);
      }
      // Update only the rows that have changed
      for (const rowIndex of rowsToUpdate) {
        const rowData = dataFromOrigin[rowIndex] || [];
        const updateRange = `${body.destinationWorksheetName}!${
          body.originWorksheetFirstColumn
        }${rowIndex + 1}:${body.originWorksheetLastColumn}${rowIndex + 1}`;

        logger.info(
          logTag,
          `Updating row ${rowIndex + 1} at range: ${updateRange}`
        );

        await updateDataInSheet(
          updateRange,
          [rowData],
          body.destinationSpreadsheetId
        );
      }
    }

    res.send("Completed");
  } catch (error) {
    logger.error(logTag, "An error occurred in gsSyncFunction:", error);
    res.status(500).send("Internal Server Error");
  }
});
