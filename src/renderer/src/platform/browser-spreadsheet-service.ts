import type { ChatAttachment, OpenWorkbookResult, SheetPreview, WorkbookSummary } from "@shared/types"

export class BrowserSpreadsheetService {
  async connectLiveWorkbook(): Promise<OpenWorkbookResult> {
    return this.getActiveWorkbookData()
  }

  async createSheet(sheetName: string): Promise<WorkbookSummary> {
    await Excel.run(async (context: any) => {
      context.workbook.worksheets.add(sheetName)
      await context.sync()
    })
    return this.getWorkbookSummary()
  }

  async writeCell(cellAddress: string, value: string, sheetName?: string): Promise<WorkbookSummary> {
    await Excel.run(async (context: any) => {
      const worksheet = getWorksheet(context, sheetName)
      worksheet.getRange(cellAddress).values = [[value]]
      await context.sync()
    })
    return this.getWorkbookSummary()
  }

  async deleteCell(cellAddress: string, sheetName?: string): Promise<WorkbookSummary> {
    await Excel.run(async (context: any) => {
      const worksheet = getWorksheet(context, sheetName)
      worksheet.getRange(cellAddress).clear(Excel.ClearApplyTo.contents)
      await context.sync()
    })
    return this.getWorkbookSummary()
  }

  async writeTable(rows: string[][], startCell?: string, sheetName?: string): Promise<{ workbook: WorkbookSummary; startCell: string }> {
    const anchor = startCell?.trim() ? normalizeTableAnchor(startCell) : "A1"
    await this.writeRange(anchor, rows, sheetName)
    return { workbook: await this.getWorkbookSummary(), startCell: anchor }
  }

  async writeRange(range: string, rows: string[][], sheetName?: string): Promise<string> {
    validateRangeCanFitRows(range, rows)
    await Excel.run(async (context: any) => {
      const worksheet = getWorksheet(context, sheetName)
      worksheet.getRange(normalizeRangeAddress(range)).values = rows
      await context.sync()
    })
    return `${sheetName ?? "현재 시트"}!${normalizeRangeAddress(range)}에 ${rows.length}행을 기록했습니다.`
  }

  async appendRows(rows: string[][], sheetName?: string): Promise<{ workbook: WorkbookSummary; startCell: string }> {
    const workbook = await this.getWorkbookSummary()
    const startCell = `A${Math.max(workbook.activeSheetName ? 2 : 1, 1)}`
    await this.writeRange(startCell, rows, sheetName)
    return { workbook: await this.getWorkbookSummary(), startCell }
  }

  async deleteRow(rowNumber: number, sheetName?: string): Promise<WorkbookSummary> {
    await Excel.run(async (context: any) => {
      const worksheet = getWorksheet(context, sheetName)
      worksheet.getRange(`${rowNumber}:${rowNumber}`).delete(Excel.DeleteShiftDirection.up)
      await context.sync()
    })
    return this.getWorkbookSummary()
  }

  async readRange(range: string, sheetName?: string): Promise<{ workbook: WorkbookSummary; attachment: ChatAttachment }> {
    if (isWholeSheetRange(range)) {
      return this.readUsedRange(sheetName)
    }

    const normalized = normalizeRangeAddress(range)
    const rows = await this.readRangeValues(normalized, sheetName)
    return {
      workbook: await this.getWorkbookSummary(),
      attachment: { kind: "table", title: `${sheetName ?? "현재 시트"}!${normalized}`, rows },
    }
  }

  async readUsedRange(sheetName?: string): Promise<{ workbook: WorkbookSummary; attachment: ChatAttachment }> {
    const result = await Excel.run(async (context: any) => {
      const worksheet = getWorksheet(context, sheetName)
      const usedRange = worksheet.getUsedRangeOrNullObject()
      await context.sync()
      if (usedRange.isNullObject) {
        return [] as string[][]
      }
      usedRange.load(["values", "address"])
      await context.sync()
      return normalizeRows(usedRange.values)
    })

    return {
      workbook: await this.getWorkbookSummary(),
      attachment: { kind: "table", title: `${sheetName ?? "현재 시트"}!used range`, rows: result },
    }
  }

  async getActiveWorkbookData(): Promise<OpenWorkbookResult> {
    return {
      workbook: await this.getWorkbookSummary(),
      preview: await this.getPreview(),
    }
  }

  getObservation(toolContext: {
    userAttachment?: ChatAttachment
    lastWebSearchTable?: ChatAttachment
    lastReadRangeTable?: ChatAttachment
  }) {
    return {
      workbook: undefined,
      preview: [],
      availableTables: [
        toAvailableTable("user_attachment", toolContext.userAttachment),
        toAvailableTable("last_web_search", toolContext.lastWebSearchTable),
        toAvailableTable("last_read_range", toolContext.lastReadRangeTable),
      ].filter((value): value is NonNullable<typeof value> => Boolean(value)),
    }
  }

  private async readRangeValues(range: string, sheetName?: string) {
    return Excel.run(async (context: any) => {
      const worksheet = getWorksheet(context, sheetName)
      const target = worksheet.getRange(range)
      target.load(["values"])
      await context.sync()
      return normalizeRows(target.values)
    })
  }

  private async getWorkbookSummary(): Promise<WorkbookSummary> {
    return Excel.run(async (context: any) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet()
      worksheet.load(["name"])
      context.workbook.worksheets.load(["items/name"])
      await context.sync()
      return {
        mode: "live",
        name: "Excel Workbook",
        path: "",
        format: "xlsx",
        sheetNames: context.workbook.worksheets.items.map((sheet: any) => sheet.name),
        activeSheetName: worksheet.name,
      }
    })
  }

  private async getPreview(): Promise<SheetPreview[]> {
    return Excel.run(async (context: any) => {
      const sheets = context.workbook.worksheets
      sheets.load(["items/name"])
      await context.sync()
      const previews: SheetPreview[] = []
      for (const sheet of sheets.items.slice(0, 3)) {
        const usedRange = sheet.getUsedRangeOrNullObject()
        await context.sync()
        if (usedRange.isNullObject) {
          previews.push({ sheetName: sheet.name, rows: [] })
          continue
        }
        usedRange.load(["values"])
        await context.sync()
        previews.push({ sheetName: sheet.name, rows: normalizeRows(usedRange.values).slice(0, 7).map((row) => row.slice(0, 7)) })
      }
      return previews
    })
  }
}

function getWorksheet(context: any, sheetName?: string) {
  return sheetName ? context.workbook.worksheets.getItem(sheetName) : context.workbook.worksheets.getActiveWorksheet()
}

function normalizeRows(values: unknown) {
  if (!Array.isArray(values)) return [] as string[][]
  return values.map((row) => (Array.isArray(row) ? row.map((cell) => stringifyCellValue(cell)) : [stringifyCellValue(row)]))
}

function stringifyCellValue(value: unknown) {
  if (value == null) return ""
  if (typeof value === "object") return JSON.stringify(value)
  return String(value)
}

function isWholeSheetRange(range: string) {
  const normalized = range.trim().toLowerCase().replace(/\s+/g, "")
  return ["전체", "전체범위", "all", "allsheet", "wholesheet", "usedrange", "currentsheet"].includes(normalized)
}

function normalizeTableAnchor(startCell: string) {
  return normalizeRangeAddress(startCell).split(":")[0] ?? "A1"
}

function normalizeRangeAddress(range: string) {
  return range
    .trim()
    .replace(/[()\[\]]/g, "")
    .replace(/\$/g, "")
    .replace(/\s+/g, "")
    .replace(/^([^!:]+)!([A-Za-z]+\d+:[A-Za-z]+\d+)$/i, "$2")
    .replace(/^([^!:]+)!([A-Za-z]+\d+)$/i, "$2")
    .toUpperCase()
}

function validateRangeCanFitRows(range: string, rows: string[][]) {
  if (isWholeSheetRange(range)) return
  if (!range.includes(":")) return
  const rowCount = rows.length
  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0)
  if (rowCount === 0 || columnCount === 0) return
}

function toAvailableTable(
  source: "user_attachment" | "last_web_search" | "last_read_range",
  attachment?: ChatAttachment,
) {
  if (!attachment) return undefined
  return {
    source,
    title: attachment.title,
    rows: attachment.rows.length,
    columns: attachment.rows[0]?.length ?? 0,
  }
}
