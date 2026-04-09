import ExcelJS from "exceljs"
import type { OpenWorkbookResult } from "@shared/types"

export class XlsxAdapter {
  supports(format: "xlsx" | "csv" | "numbers") {
    return format === "xlsx"
  }

  async openWorkbook(filePath: string): Promise<OpenWorkbookResult> {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(filePath)

    const sheetNames = workbook.worksheets.map((sheet) => sheet.name)
    const preview = workbook.worksheets.slice(0, 3).map((sheet) => ({
      sheetName: sheet.name,
      rows: sheet.getSheetValues().slice(1, 8).map((row) => normalizeRow(row)),
    }))

    return {
      workbook: {
        mode: "file",
        name: filePath.split(/[/\\]/).at(-1) ?? filePath,
        path: filePath,
        format: "xlsx",
        sheetNames,
        activeSheetName: workbook.worksheets[0]?.name,
      },
      preview,
    }
  }
}

function normalizeRow(row: unknown) {
  if (!Array.isArray(row)) return []
  return row.slice(1, 8).map((cell) => stringifyCell(cell))
}

function stringifyCell(value: unknown) {
  if (value == null) return ""
  if (typeof value === "object") return JSON.stringify(value)
  return String(value)
}
