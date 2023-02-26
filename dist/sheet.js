export class XlsxGridSheet {
    constructor(name, grid, start, end) {
        this.name = name;
        this.grid = grid;
        this.start = start;
        this.end = end;
        this.rowLength = grid.length;
        this.columnLength = grid.length > 0 ? grid[0].length : 0;
    }
    at(row, column) {
        if (typeof column === "string") {
            column = parseColumnNameToIndex(column);
        }
        return this.grid[row - 1][column];
    }
    onRow(row) {
        return this.grid[row - 1];
    }
}
export function parseColumnNameToIndex(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
        const charCode = column.charCodeAt(i) - 64;
        result = result * 26 + charCode;
    }
    return result - 1;
}
//# sourceMappingURL=sheet.js.map