interface SyncDataRequestBody {
  originWorksheetId: number;
  originWorksheetName: string;
  originSpreadsheetName: string;
  originWorksheetFirstColumn: string;
  originWorksheetLastColumn: string;
  originWorksheetFirstRow: number;
  originWorksheetLastRow: number;
  originSpreadsheetId: string;
  destinationSpreadsheetId: string;
  destinationSpreadsheetName: string;
  destinationWorksheetId: number;
  destinationWorksheetName: string;
}

export {SyncDataRequestBody};
