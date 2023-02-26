import xlsx from 'node-xlsx'
import { XlsxGridSheet, XlsxSheet, CellRange } from './sheet.js'
export interface XlsxSheetLoaderPorvider {
  name: string
  type: string
  create: (context: any) => XlsxSheetLoaderLoader
}
export interface XlsxSheetLoaderLoader {
  load: (sheet: XlsxSheet) => Promise<any>
}

export function parseXlsxDocument(source: any): Map<string, XlsxSheet> {
  const rawMeta = xlsx.parseMetadata(source)
  const rawDocument = xlsx.parse(source)
  const name2Sheet = new Map<string, XlsxSheet>()
  if (rawMeta.length !== rawDocument.length) {
    throw new XlsxDocumentParseError(XlsxDocumentErrorType.metaNotMatch)
  }

  for (let i = 0; i < rawMeta.length; i++) {
    const meta = rawMeta[i]
    const doc = rawDocument[i].data
    const rangeInfo = meta.data
    const { start, end } = parseXlsxRangeInfo(rangeInfo)
    name2Sheet.set(meta.name, new XlsxGridSheet(meta.name, doc as string[][], start, end))
  }

  return name2Sheet
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