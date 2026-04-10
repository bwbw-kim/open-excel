import type { ChatAttachment, WorkbookSummary, SheetPreview } from "@/shared/types"

export class ExcelService {
  async getWorkbookInfo(): Promise<WorkbookSummary | null> {
    return Excel.run(async (context) => {
      const workbook = context.workbook
      const sheets = workbook.worksheets
      const activeSheet = workbook.worksheets.getActiveWorksheet()

      sheets.load("items/name")
      activeSheet.load("name")
      workbook.load("name")

      await context.sync()

      const sheetNames = sheets.items.map((sheet) => sheet.name)

      return {
        mode: "live" as const,
        name: workbook.name || "Workbook",
        path: workbook.name || "Workbook",
        format: "xlsx" as const,
        sheetNames,
        activeSheetName: activeSheet.name,
      }
    })
  }

  async getSheetPreview(maxSheets = 3, maxRows = 7): Promise<SheetPreview[]> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets
      sheets.load("items/name")

      await context.sync()

      const previews: SheetPreview[] = []
      const sheetsToPreview = sheets.items.slice(0, maxSheets)

      for (const sheet of sheetsToPreview) {
        const usedRange = sheet.getUsedRangeOrNullObject()
        usedRange.load("values, isNullObject")
      }

      await context.sync()

      for (let i = 0; i < sheetsToPreview.length; i++) {
        const sheet = sheetsToPreview[i]
        const usedRange = sheet.getUsedRangeOrNullObject()

        if (usedRange.isNullObject) {
          previews.push({ sheetName: sheet.name, rows: [[""]] })
        } else {
          const rows = (usedRange.values as unknown[][])
            .slice(0, maxRows)
            .map((row) => row.map((cell) => this.stringifyCell(cell)))
          previews.push({ sheetName: sheet.name, rows })
        }
      }

      return previews
    })
  }

  async readRange(range: string, sheetName?: string): Promise<ChatAttachment> {
    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const targetRange = sheet.getRange(range)
      targetRange.load("values, address")
      sheet.load("name")

      await context.sync()

      const rows = (targetRange.values as unknown[][]).map((row) =>
        row.map((cell) => this.stringifyCell(cell)),
      )

      return {
        kind: "table" as const,
        title: `${sheet.name}!${this.normalizeAddress(targetRange.address)}`,
        rows,
      }
    })
  }

  async readUsedRange(sheetName?: string): Promise<ChatAttachment> {
    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const usedRange = sheet.getUsedRangeOrNullObject()
      usedRange.load("values, address, isNullObject")
      sheet.load("name")

      await context.sync()

      if (usedRange.isNullObject) {
        return {
          kind: "table" as const,
          title: `${sheet.name}!A1`,
          rows: [[""]],
        }
      }

      const rows = (usedRange.values as unknown[][]).map((row) =>
        row.map((cell) => this.stringifyCell(cell)),
      )

      return {
        kind: "table" as const,
        title: `${sheet.name}!${this.normalizeAddress(usedRange.address)}`,
        rows,
      }
    })
  }

  async writeRange(range: string, rows: string[][], sheetName?: string): Promise<string> {
    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const targetRange = sheet.getRange(range)
      targetRange.values = rows
      targetRange.format.autofitColumns()

      sheet.load("name")
      await context.sync()

      return `${sheet.name}!${range}에 ${rows.length}행을 작성했습니다.`
    })
  }

  async writeCell(cell: string, value: string, sheetName?: string): Promise<string> {
    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const targetCell = sheet.getRange(cell)
      targetCell.values = [[value]]

      sheet.load("name")
      await context.sync()

      return `${sheet.name}!${cell}에 값을 작성했습니다.`
    })
  }

  async deleteCell(cell: string, sheetName?: string): Promise<string> {
    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const targetCell = sheet.getRange(cell)
      targetCell.clear(Excel.ClearApplyTo.contents)

      sheet.load("name")
      await context.sync()

      return `${sheet.name}!${cell}의 내용을 삭제했습니다.`
    })
  }

  async appendRows(rows: string[][], sheetName?: string): Promise<string> {
    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const usedRange = sheet.getUsedRangeOrNullObject()
      usedRange.load("rowCount, isNullObject")

      await context.sync()

      const startRow = usedRange.isNullObject ? 1 : usedRange.rowCount + 1
      const startCell = `A${startRow}`
      const endCell = `${this.columnToLetter(rows[0].length)}${startRow + rows.length - 1}`
      const targetRange = sheet.getRange(`${startCell}:${endCell}`)

      targetRange.values = rows
      targetRange.format.autofitColumns()

      sheet.load("name")
      await context.sync()

      return `${sheet.name}에 ${rows.length}행을 추가했습니다.`
    })
  }

  async deleteRow(rowNumber: number, sheetName?: string): Promise<string> {
    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const row = sheet.getRange(`${rowNumber}:${rowNumber}`)
      row.delete(Excel.DeleteShiftDirection.up)

      sheet.load("name")
      await context.sync()

      return `${sheet.name}의 ${rowNumber}행을 삭제했습니다.`
    })
  }

  async createSheet(sheetName: string): Promise<string> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets
      const existingSheet = sheets.getItemOrNullObject(sheetName)
      existingSheet.load("isNullObject")

      await context.sync()

      if (!existingSheet.isNullObject) {
        throw new Error(`이미 "${sheetName}" 시트가 존재합니다.`)
      }

      const newSheet = sheets.add(sheetName)
      newSheet.activate()

      await context.sync()

      return `"${sheetName}" 시트를 생성했습니다.`
    })
  }

  getObservation(toolContext: { lastReadRangeTable?: ChatAttachment }): SpreadsheetAgentObservation {
    return {
      mode: "live",
      connectedWorkbook: null,
      lastReadRangeTable: toolContext.lastReadRangeTable,
    }
  }

  private stringifyCell(value: unknown): string {
    if (value == null) return ""
    return String(value)
  }

  private normalizeAddress(address: string): string {
    const parts = address.split("!")
    return parts.length > 1 ? parts[1].replace(/\$/g, "") : address.replace(/\$/g, "")
  }

  private columnToLetter(column: number): string {
    let letter = ""
    while (column > 0) {
      const remainder = (column - 1) % 26
      letter = String.fromCharCode(65 + remainder) + letter
      column = Math.floor((column - 1) / 26)
    }
    return letter
  }
}

export interface SpreadsheetAgentObservation {
  mode: "live" | "file"
  connectedWorkbook: WorkbookSummary | null
  lastReadRangeTable?: ChatAttachment
}
