import xlsx from 'node-xlsx';
import { XlsxGridSheet } from './sheet.js';
export function parseXlsxDocument(source) {
    const rawMeta = xlsx.parseMetadata(source);
    const rawDocument = xlsx.parse(source);
    const name2Sheet = new Map();
    if (rawMeta.length !== rawDocument.length) {
        throw new XlsxDocumentParseError(XlsxDocumentErrorType.metaNotMatch);
    }
    for (let i = 0; i < rawMeta.length; i++) {
        const meta = rawMeta[i];
        const doc = rawDocument[i].data;
        const rangeInfo = meta.data;
        const { start, end } = parseXlsxRangeInfo(rangeInfo);
        name2Sheet.set(meta.name, new XlsxGridSheet(meta.name, doc, start, end));
    }
    return name2Sheet;
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
//# sourceMappingURL=loader.js.map