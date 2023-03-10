import xlsx from "node-xlsx";
import { XlsxGridSheet } from "./sheet.js";
import * as fs from "fs";
import path from "path";
import { promisify } from "util";
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
        start: rangeInfo
            ? {
                row: rangeInfo.s.r,
                column: rangeInfo.s.c,
            }
            : null,
        end: rangeInfo
            ? {
                row: rangeInfo.e.r,
                column: rangeInfo.e.c,
            }
            : null,
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
export async function loadSheetProvider(path) {
    const module = await import(path);
    const provider = module.default;
    if ("name" in provider && "type" in provider && "create" in provider) {
        return provider;
    }
    return null;
}
export async function loadSheetProviderInDir(folder, onError = null) {
    if (!fs.existsSync(folder))
        return [];
    const readdir = promisify(fs.readdir);
    const files = await readdir(folder, { withFileTypes: true });
    const providers = [];
    for (const file of files) {
        const fileName = file.name;
        const ext = path.extname(fileName);
        if (ext === ".js" || ext === ".mjs") {
            const filePath = path.join(folder, fileName);
            const pathUri = filePathToUri(filePath);
            try {
                const provider = await loadSheetProvider(pathUri);
                if (provider) {
                    providers.push(Object.assign(Object.assign({}, provider), { filePath,
                        pathUri }));
                }
            }
            catch (e) {
                onError === null || onError === void 0 ? void 0 : onError(e);
            }
        }
    }
    return providers;
}
function filePathToUri(filePath) {
    let pathName = path.resolve(filePath).replace(/\\/g, "/");
    if (pathName[0] !== "/") {
        pathName = `/${pathName}`;
    }
    return encodeURI(`file://${pathName}`);
}
//# sourceMappingURL=loader.js.map