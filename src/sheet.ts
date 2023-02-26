export interface CellRange {
  row: number
  column: number
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
  at: (row: RowSymbol, column: ColumnSymbol) => string | undefined
  /**
   * 
   * @param row the number of row, starts with 1
   * @returns a row of cells
   */
  onRow: (row: RowSymbol) => string[]
}

export class XlsxGridSheet implements XlsxSheet {
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

  at(row: RowSymbol, column: ColumnSymbol): string | undefined {
    if (typeof column === "string") {
      column = parseColumnNameToIndex(column)
    }
    return this.grid[row - 1][column]
  }

  onRow(row: RowSymbol): string[] {
    return this.grid[row - 1]
  }
}
/**
 * @param column the name of column, such as "A", "B", "AC"
 * @returns the corresponding index
 */
export function parseColumnNameToIndex(column: string): number {
  let result = 0
  for (let i = 0; i < column.length; i++) {
    const charCode = column.charCodeAt(i) - 64
    result = result * 26 + charCode
  }
  return result - 1
}
