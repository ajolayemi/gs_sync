/**
 * A representation of the needed info of a Google Worksheet
 */
interface SheetInfo {
  /**
   * The name of a Google Worksheet
   */
  name: string;

  /**
   * The unique id for a google worksheet
   */
  id: number;

  /**
   * A letter specifying the first column to read
   */
  firstColumn: string;

  /**
   * A letter specifying the last column to read
   */
  lastColumn: string;

  /**
   * A number specifying the first row to read
   */
  firstRow: number;

  /**
   * A number specifying the last row to read
   */
  lastRow: number;
}

interface SheetSize {
  rowCount: number;
  columnCount: number;
}

/**
 * Info read from google sheet to know what files needs to be imported
 * and there respective destination paths
 */
interface ImportDataInfo {
  originSpreadsheetFileId: string;
  destinationWorksheetId: number;
  destinationSpreadsheetFileId: string;
  destinationSpreadsheetWorksheetName: string;
}

export {SheetSize, SheetInfo, ImportDataInfo};
