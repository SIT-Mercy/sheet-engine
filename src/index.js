import xlsx from 'node-xlsx';
class XlsxSheetImpl {
    constructor(name, grid, start, end) {
        this.name = name;
        this.grid = grid;
        this.start = start;
        this.end = end;
        this.columnLength = grid.length;
        this.rowLength = grid.length > 0 ? grid[0].length : 0;
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
    var result = 0;
    for (var i = 0; i < column.length; i++) {
        var charCode = column.charCodeAt(i) - 64;
        result = result * 26 + charCode;
    }
    return result - 1;
}
export function parseXlsxDocument(source) {
    const rawMeta = xlsx.parseMetadata(source);
    const rawDocument = xlsx.parse(source);
    const name2Sheet = new Map();
    if (rawMeta.length !== rawDocument.length) {
        throw new XlsxDocumentParseError(XlsxDocumentErrorType.metaNotMatch);
    }
    for (let i = 0; i < rawMeta.length; i++) {
        const meta = rawMeta[i];
        const doc = rawDocument[i];
        const rangeInfo = meta.data;
        const { start, end } = parseXlsxRangeInfo(rangeInfo);
        name2Sheet.set(meta.name, new XlsxSheetImpl(meta.name, doc, start, end));
    }
    return {
        name2Sheet
    };
}
function parseXlsxRangeInfo(rangeInfo) {
    return {
        start: rangeInfo ? {
            row: rangeInfo.s.r,
            column: rangeInfo.s.c,
        } : null,
        end: rangeInfo ? {
            row: rangeInfo.e.r,
            column: rangeInfo.e.c,
        } : null,
    };
}
export var XlsxDocumentErrorType;
(function (XlsxDocumentErrorType) {
    XlsxDocumentErrorType["metaNotMatch"] = "metaNotMatch";
})(XlsxDocumentErrorType || (XlsxDocumentErrorType = {}));
export class XlsxDocumentParseError extends Error {
    constructor(message) {
        super(message);
        this.name = "XlsxDocumentParseError";
    }
}
//# sourceMappingURL=index.js.map