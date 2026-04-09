import path from "node:path"
import ExcelJS from "exceljs"
import type { ChatAttachment, OpenWorkbookResult, SpreadsheetFormat, WorkbookSummary } from "@shared/types"
import type { Logger } from "../logging/logger"
import { ExcelLiveService } from "../live/excel-live-service"

interface ActiveWorkbookState {
  path: string
  format: Exclude<SpreadsheetFormat, "numbers">
  workbook: ExcelJS.Workbook
}

export class SpreadsheetService {
  private activeWorkbook?: ActiveWorkbookState
  private activeLiveWorkbook?: WorkbookSummary
  private activeMode: "file" | "live" = "file"
  private readonly liveService: ExcelLiveService

  constructor(private readonly logger: Logger) {
    this.liveService = new ExcelLiveService(logger)
  }

  async openWorkbook(filePath: string): Promise<OpenWorkbookResult> {
    const format = detectFormat(filePath)
    if (format === "numbers") {
      throw new Error("Numbers support is not included in the MVP build yet.")
    }

    const workbook = new ExcelJS.Workbook()

    this.logger.info("opening workbook", { filePath, format })

    if (format === "xlsx") {
      await workbook.xlsx.readFile(filePath)
    } else {
      await workbook.csv.readFile(filePath)
    }

    this.activeMode = "file"
    this.activeLiveWorkbook = undefined
    this.activeWorkbook = { path: filePath, format, workbook }
    return this.getActiveWorkbookData()
  }

  async connectLiveWorkbook(): Promise<OpenWorkbookResult> {
    const workbook = await this.liveService.connect()
    this.activeMode = "live"
    this.activeWorkbook = undefined
    this.activeLiveWorkbook = workbook
    return {
      workbook,
      preview: [],
    }
  }

  getActiveWorkbook() {
    if (this.activeMode === "live") {
      return this.activeLiveWorkbook
    }
    return this.activeWorkbook ? toWorkbookSummary(this.activeWorkbook) : undefined
  }

  getActiveWorkbookData(): OpenWorkbookResult {
    if (this.activeMode === "live") {
      const workbook = this.requireLiveWorkbook()
      return { workbook, preview: [] }
    }

    const activeWorkbook = this.requireActiveWorkbook()
    return {
      workbook: toWorkbookSummary(activeWorkbook),
      preview: createPreview(activeWorkbook.workbook),
    }
  }

  getObservation(toolContext: {
    userAttachment?: ChatAttachment
    lastWebSearchTable?: ChatAttachment
    lastReadRangeTable?: ChatAttachment
  }) {
    const activeWorkbook = this.getActiveWorkbook()
    return {
      workbook: activeWorkbook,
      preview: this.activeMode === "file" && this.activeWorkbook ? createPreview(this.activeWorkbook.workbook) : [],
      availableTables: [
        toAvailableTable("user_attachment", toolContext.userAttachment),
        toAvailableTable("last_web_search", toolContext.lastWebSearchTable),
        toAvailableTable("last_read_range", toolContext.lastReadRangeTable),
      ].filter((value): value is NonNullable<typeof value> => Boolean(value)),
    }
  }

  async createSheet(sheetName: string) {
    if (this.activeMode === "live") {
      this.activeLiveWorkbook = await this.liveService.createSheet(sheetName)
      return `${sheetName} 시트를 만들었습니다.`
    }

    const activeWorkbook = this.requireActiveWorkbook()
    if (activeWorkbook.format === "csv") {
      throw new Error("CSV 파일은 새 시트를 만들 수 없습니다.")
    }

    if (activeWorkbook.workbook.getWorksheet(sheetName)) {
      throw new Error(`이미 ${sheetName} 시트가 있습니다.`)
    }

    activeWorkbook.workbook.addWorksheet(sheetName)
    await this.saveActiveWorkbook()
    return `${sheetName} 시트를 만들었습니다.`
  }

  async writeCell(cellAddress: string, value: string, sheetName?: string) {
    if (this.activeMode === "live") {
      const workbook = await this.liveService.writeCell(cellAddress, value, sheetName)
      this.activeLiveWorkbook = workbook
      return `${workbook.activeSheetName ?? sheetName ?? "현재 시트"}!${cellAddress} 셀에 값을 입력했습니다.`
    }

    const worksheet = this.getWorksheet(sheetName)
    worksheet.getCell(cellAddress).value = value
    await this.saveActiveWorkbook()
    return `${worksheet.name}!${cellAddress} 셀에 값을 입력했습니다.`
  }

  async deleteCell(cellAddress: string, sheetName?: string) {
    if (this.activeMode === "live") {
      const workbook = await this.liveService.deleteCell(cellAddress, sheetName)
      this.activeLiveWorkbook = workbook
      return `${workbook.activeSheetName ?? sheetName ?? "현재 시트"}!${cellAddress} 셀 내용을 삭제했습니다.`
    }

    const worksheet = this.getWorksheet(sheetName)
    worksheet.getCell(cellAddress).value = null
    await this.saveActiveWorkbook()
    return `${worksheet.name}!${cellAddress} 셀 내용을 삭제했습니다.`
  }

  async writeTable(rows: string[][], startCell?: string, sheetName?: string) {
    if (this.activeMode === "live") {
      const result = await this.liveService.writeTable(rows, startCell, sheetName)
      this.activeLiveWorkbook = result.workbook
      return `${result.workbook.activeSheetName ?? sheetName ?? "현재 시트"}!${result.startCell}부터 ${rows.length}행 표를 기록했습니다.`
    }

    const worksheet = this.getWorksheet(sheetName)
    const actualStartCell = startCell?.trim() ? normalizeTableAnchor(startCell) : findDefaultWriteCell(worksheet)
    const { row, col } = decodeCellAddress(actualStartCell)

    rows.forEach((rowValues, rowOffset) => {
      rowValues.forEach((value, colOffset) => {
        worksheet.getCell(row + rowOffset, col + colOffset).value = value
      })
    })

    await this.saveActiveWorkbook()
    return `${worksheet.name}!${actualStartCell}부터 ${rows.length}행 표를 기록했습니다.`
  }

  async writeRange(range: string, rows: string[][], sheetName?: string) {
    validateRangeCanFitRows(range, rows)
    return this.writeTable(rows, range, sheetName)
  }

  async appendRows(rows: string[][], sheetName?: string) {
    if (this.activeMode === "live") {
      const result = await this.liveService.appendRows(rows, sheetName)
      this.activeLiveWorkbook = result.workbook
      return `${result.workbook.activeSheetName ?? sheetName ?? "현재 시트"} 시트에 ${rows.length}개 행을 추가했습니다.`
    }

    const worksheet = this.getWorksheet(sheetName)
    rows.forEach((row) => worksheet.addRow(row))
    await this.saveActiveWorkbook()
    return `${worksheet.name} 시트에 ${rows.length}개 행을 추가했습니다.`
  }

  async deleteRow(rowNumber: number, sheetName?: string) {
    if (this.activeMode === "live") {
      const workbook = await this.liveService.deleteRow(rowNumber, sheetName)
      this.activeLiveWorkbook = workbook
      return `${workbook.activeSheetName ?? sheetName ?? "현재 시트"} 시트에서 ${rowNumber}행을 삭제했습니다.`
    }

    const worksheet = this.getWorksheet(sheetName)
    worksheet.spliceRows(rowNumber, 1)
    await this.saveActiveWorkbook()
    return `${worksheet.name} 시트에서 ${rowNumber}행을 삭제했습니다.`
  }

  async readRange(range: string, sheetName?: string): Promise<ChatAttachment> {
    if (this.activeMode === "live") {
      const result = await this.liveService.readRange(range, sheetName)
      this.activeLiveWorkbook = result.workbook
      return result.attachment
    }

    if (isWholeSheetRange(range)) {
      return this.readUsedRange(sheetName)
    }

    const worksheet = this.getWorksheet(sheetName)
    const { start, end } = parseRange(range)
    const rows: string[][] = []

    for (let rowIndex = start.row; rowIndex <= end.row; rowIndex += 1) {
      const currentRow: string[] = []
      for (let colIndex = start.col; colIndex <= end.col; colIndex += 1) {
        currentRow.push(stringifyCellValue(worksheet.getCell(rowIndex, colIndex).value))
      }
      rows.push(currentRow)
    }

    return {
      kind: "table",
      title: `${worksheet.name}!${range}`,
      rows,
    }
  }

  async readUsedRange(sheetName?: string): Promise<ChatAttachment> {
    if (this.activeMode === "live") {
      const result = await this.liveService.readUsedRange(sheetName)
      this.activeLiveWorkbook = result.workbook
      return result.attachment
    }

    const worksheet = this.getWorksheet(sheetName)
    const rowCount = Math.max(worksheet.rowCount, 1)
    const colCount = Math.max(worksheet.actualColumnCount, 1)
    const endColumn = columnNumberToName(colCount)
    const range = `A1:${endColumn}${rowCount}`
    return this.readRange(range, sheetName)
  }

  private getWorksheet(sheetName?: string) {
    const activeWorkbook = this.requireActiveWorkbook()
    const worksheet = sheetName
      ? activeWorkbook.workbook.getWorksheet(sheetName)
      : activeWorkbook.workbook.worksheets[0]

    if (!worksheet) {
      throw new Error(sheetName ? `${sheetName} 시트를 찾지 못했습니다.` : "활성 시트를 찾지 못했습니다.")
    }

    return worksheet
  }

  private requireActiveWorkbook() {
    if (!this.activeWorkbook) {
      throw new Error("먼저 Excel에 연결해 주세요.")
    }
    return this.activeWorkbook
  }

  private requireLiveWorkbook() {
    if (!this.activeLiveWorkbook) {
      throw new Error("먼저 실행 중인 Excel에 연결해 주세요.")
    }
    return this.activeLiveWorkbook
  }

  private async saveActiveWorkbook() {
    const activeWorkbook = this.requireActiveWorkbook()
    if (activeWorkbook.format === "xlsx") {
      await activeWorkbook.workbook.xlsx.writeFile(activeWorkbook.path)
      return
    }

    await activeWorkbook.workbook.csv.writeFile(activeWorkbook.path, {
      sheetName: activeWorkbook.workbook.worksheets[0]?.name,
      formatterOptions: {
        writeBOM: true,
      },
    })
  }
}

function detectFormat(filePath: string): SpreadsheetFormat {
  const extension = path.extname(filePath).toLowerCase()
  if (extension === ".xlsx") return "xlsx"
  if (extension === ".csv") return "csv"
  if (extension === ".numbers") return "numbers"
  throw new Error(`Unsupported file extension: ${extension}`)
}

function toWorkbookSummary(activeWorkbook: ActiveWorkbookState): WorkbookSummary {
  return {
    mode: "file",
    name: path.basename(activeWorkbook.path),
    path: activeWorkbook.path,
    format: activeWorkbook.format,
    sheetNames: activeWorkbook.workbook.worksheets.map((sheet) => sheet.name),
    activeSheetName: activeWorkbook.workbook.worksheets[0]?.name,
  }
}

function createPreview(workbook: ExcelJS.Workbook) {
  return workbook.worksheets.slice(0, 3).map((sheet) => ({
    sheetName: sheet.name,
    rows: Array.from({ length: Math.min(sheet.rowCount, 7) }, (_value, index) => {
      const row = sheet.getRow(index + 1)
      return Array.from({ length: Math.min(row.cellCount, 7) }, (_cell, cellIndex) => {
        return stringifyCellValue(row.getCell(cellIndex + 1).value)
      })
    }),
  }))
}

function decodeCellAddress(address: string) {
  const normalized = normalizeCellAddress(address)
  const match = normalized.match(/^([A-Z]+)(\d+)$/)
  if (!match) {
    throw new Error(`유효하지 않은 셀 주소입니다: ${address}`)
  }

  return {
    col: columnNameToNumber(match[1]),
    row: Number(match[2]),
  }
}

function parseRange(range: string) {
  const normalizedRange = normalizeRangeAddress(range)
  const [startRaw, endRaw] = normalizedRange.split(":")
  const start = decodeCellAddress(startRaw)
  const end = decodeCellAddress(endRaw ?? startRaw)
  return { start, end }
}

function normalizeCellAddress(address: string) {
  return address
    .trim()
    .replace(/[()\[\]]/g, "")
    .replace(/\$/g, "")
    .replace(/^.*!/g, "")
    .replace(/\s+/g, "")
    .toUpperCase()
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

function normalizeTableAnchor(startCell: string) {
  if (isWholeSheetRange(startCell)) {
    return "A1"
  }

  const normalized = normalizeRangeAddress(startCell)
  return normalized.split(":")[0] ?? normalized
}

function isWholeSheetRange(range: string) {
  const normalized = range.trim().toLowerCase().replace(/\s+/g, "")
  return ["전체", "전체범위", "all", "allsheet", "wholesheet", "usedrange", "currentsheet"].includes(normalized)
}

function validateRangeCanFitRows(range: string, rows: string[][]) {
  if (isWholeSheetRange(range)) {
    return
  }

  const normalizedRange = normalizeRangeAddress(range)
  const [startRaw, endRaw] = normalizedRange.split(":")
  if (!endRaw) {
    return
  }

  const start = decodeCellAddress(startRaw)
  const end = decodeCellAddress(endRaw)
  const maxRows = end.row - start.row + 1
  const maxColumns = end.col - start.col + 1
  const rowCount = rows.length
  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0)

  if (rowCount > maxRows || columnCount > maxColumns) {
    throw new Error(`${range} 범위(${maxRows}x${maxColumns})에 ${rowCount}x${columnCount} 데이터를 모두 쓸 수 없습니다.`)
  }
}

function columnNameToNumber(columnName: string) {
  return columnName.split("").reduce((sum, character) => sum * 26 + character.charCodeAt(0) - 64, 0)
}

function columnNumberToName(columnNumber: number) {
  let value = columnNumber
  let result = ""
  while (value > 0) {
    const remainder = (value - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    value = Math.floor((value - 1) / 26)
  }
  return result || "A"
}

function stringifyCellValue(value: ExcelJS.CellValue) {
  if (value == null) return ""
  if (typeof value === "object") return JSON.stringify(value)
  return String(value)
}

function findDefaultWriteCell(worksheet: ExcelJS.Worksheet) {
  const hasAnyValue = worksheet.rowCount > 1 || worksheet.getCell("A1").value != null
  if (!hasAnyValue) {
    return "A1"
  }

  return `A${worksheet.rowCount + 1}`
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
