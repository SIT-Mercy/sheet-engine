import xlsx from "node-xlsx"
import { XlsxGridSheet, type XlsxSheet, type CellRange } from "./sheet.js"
import * as fs from "fs"
import path from "path"
import { promisify } from "util"

export class XlsxSheetLoaderEntry {
  provider: XlsxSheetLoaderPorvider
  filePath: string
  constructor(provider: XlsxSheetLoaderPorvider, filePath: string) {
    this.provider = provider
    this.filePath = filePath
  }

  get name(): string {
    return this.provider.name
  }

  get type(): string {
    return this.provider.type
  }

  get pathUrl(): string {
    return fileUrl(this.filePath)
  }
}

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
  start: CellRange | null
  end: CellRange | null
} {
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
  }
}

export enum XlsxDocumentErrorType {
  metaNotMatch = "metaNotMatch",
}

export class XlsxDocumentParseError extends Error {
  constructor(message: XlsxDocumentErrorType) {
    super(message)
    this.name = "XlsxDocumentParseError"
  }
}

export async function loadSheetProvider(path: string): Promise<XlsxSheetLoaderPorvider | null> {
  const module = await import(path)
  const provider = module.default
  if ("name" in provider && "type" in provider && "create" in provider) {
    return provider
  }
  return null
}

/**
 * 
 * @param folder the folder where contains SheetLoaderProvider scripts
 * @returns name to provider
 */
export async function loadSheetProviderInDir(folder: string, onError: ((e: any) => any) | null = null): Promise<XlsxSheetLoaderEntry[]> {
  const readdir = promisify(fs.readdir)
  const files = await readdir(folder, { withFileTypes: true })
  const providers: XlsxSheetLoaderEntry[] = []
  for (const file of files) {
    const fileName = file.name
    const ext = path.extname(fileName)
    if (ext === ".js" || ext === ".mjs") {
      const filePath = path.join(folder, fileName)
      const pathUrl = fileUrl(filePath)
      try {
        const provider = await loadSheetProvider(pathUrl)
        if (provider) {
          providers.push(new XlsxSheetLoaderEntry(provider, filePath))
        }
      } catch (e) {
        onError?.(e)
      }
    }
  }
  return providers
}

function fileUrl(filePath: string): string {
  let pathName: string = path.resolve(filePath).replace(/\\/g, "/")

  // Windows drive letter must be prefixed with a slash
  if (pathName[0] !== "/") {
    pathName = `/${pathName}`
  }

  return encodeURI(`file://${pathName}`)
}
