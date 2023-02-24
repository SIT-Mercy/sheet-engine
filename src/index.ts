import xlsx from 'node-xlsx'
interface CellRange {
  row: number
  column: number
}
interface XlsxDocument {
  name2Sheet: Map<string, XlsxSheet>
}
export interface XlsxDocumentLoaderPorvider {
  name: string
  create: (context: any) => XlsxDocumentLoader
}
export interface XlsxDocumentLoader {
  load: (document: XlsxDocument) => Promise<void>
}
/**
 * Starts with 1
 */
export type RowSymbol = number
/**
 * Starts with 1 or "A".
 */
export type ColumnSymbol = number | string

export interface XlsxSheet {
  name: string
  start: CellRange | null
  end: CellRange | null
  columnLength: number
  rowLength: number
  /**
   * @param row the number of row, starts with 1
   * @param column the number or name of column, starts with 1 or "A"
   * @returns the content of cell
   */
  at: (row: RowSymbol, column: ColumnSymbol) => string
  /**
   * 
   * @param row the number of row, starts with 1
   * @returns a row of cells
   */
  onRow: (row: RowSymbol) => string[]
}

class XlsxSheetImpl implements XlsxSheet {
  name: string
  start: CellRange | null
  end: CellRange | null
  grid: string[][]
  columnLength: number
  rowLength: number
  constructor(
    name: string,
    grid: string[][],
    start: CellRange | null,
    end: CellRange | null,
  ) {
    this.name = name
    this.grid = grid
    this.start = start
    this.end = end
    this.rowLength = grid.length
    this.columnLength = grid.length > 0 ? grid[0].length : 0
  }
  at(row: RowSymbol, column: ColumnSymbol): string {
    if (typeof column === "string") {
      column = parseColumnNameToIndex(column)
    }
    return this.grid[row - 1][column]
  }
  onRow(row: RowSymbol): string[] {
    return this.grid[row - 1]
  }
}

export interface XlsxSheetLoaderPorvider {
  name: string
  create: (context: any) => XlsxSheetLoaderLoader
}
export interface XlsxSheetLoaderLoader {
  load: (sheet: XlsxSheet) => Promise<void>
}
/**
 * @param column the name of column, such as "A", "B", "AC"
 * @returns the corresponding index
 */
export function parseColumnNameToIndex(column: string): number {
  var result = 0
  for (var i = 0; i < column.length; i++) {
    var charCode = column.charCodeAt(i) - 64
    result = result * 26 + charCode
  }
  return result - 1
}

export function parseXlsxDocument(source: any): XlsxDocument {
  const rawMeta = xlsx.parseMetadata(source)
  const rawDocument = xlsx.parse(source)
  const name2Sheet = new Map<string, XlsxSheet>()
  if (rawMeta.length !== rawDocument.length) {
    throw new XlsxDocumentParseError(XlsxDocumentErrorType.metaNotMatch)
  }
  for (let i = 0; i < rawMeta.length; i++) {
    const meta = rawMeta[i]
    const doc: unknown = rawDocument[i]
    const rangeInfo = meta.data
    const { start, end } = parseXlsxRangeInfo(rangeInfo)
    name2Sheet.set(meta.name, new XlsxSheetImpl(meta.name, doc as string[][], start, end))
  }

  return {
    name2Sheet
  }
}
function parseXlsxRangeInfo(rangeInfo: any): {
  start: CellRange | null,
  end: CellRange | null
} {
  return {
    start: rangeInfo ? {
      row: rangeInfo.s.r,
      column: rangeInfo.s.c,
    } : null,
    end: rangeInfo ? {
      row: rangeInfo.e.r,
      column: rangeInfo.e.c,
    } : null,
  }
}

export enum XlsxDocumentErrorType {
  metaNotMatch = "metaNotMatch",
}

export class XlsxDocumentParseError extends Error {

  constructor(message: XlsxDocumentErrorType) {
    super(message);
    this.name = "XlsxDocumentParseError";
  }
}