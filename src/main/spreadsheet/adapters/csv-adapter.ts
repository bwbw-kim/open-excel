import ExcelJS from "exceljs"
import type { OpenWorkbookResult } from "@shared/types"

export class CsvAdapter {
  supports(format: "xlsx" | "csv" | "numbers") {
    return format === "csv"
  }

  async openWorkbook(filePath: string): Promise<OpenWorkbookResult> {
    const workbook = new ExcelJS.Workbook()
    const worksheet = await workbook.csv.readFile(filePath)

    const rows = worksheet.getSheetValues().slice(1, 8).map((row) => normalizeRow(row))

    return {
      workbook: {
        mode: "file",
        name: filePath.split(/[/\\]/).at(-1) ?? filePath,
        path: filePath,
        format: "csv",
        sheetNames: [worksheet.name],
        activeSheetName: worksheet.name,
      },
      preview: [
        {
          sheetName: worksheet.name,
          rows,
        },
      ],
    }
  }
}

function normalizeRow(row: unknown) {
  if (!Array.isArray(row)) return []
  return row.slice(1, 8).map((cell) => (cell == null ? "" : String(cell)))
}
