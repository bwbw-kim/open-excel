import { execFile } from "node:child_process"
import { promisify } from "node:util"
import type { ChatAttachment, SpreadsheetFormat, WorkbookSummary } from "@shared/types"
import type { Logger } from "../logging/logger"

const execFileAsync = promisify(execFile)

interface LiveWorkbookResult {
  workbook: WorkbookSummary
}

interface LiveWriteTableResult extends LiveWorkbookResult {
  startCell: string
}

interface LiveReadRangeResult extends LiveWorkbookResult {
  rows: string[][]
  title: string
}

export class ExcelLiveService {
  constructor(private readonly logger: Logger) {}

  async connect(): Promise<WorkbookSummary> {
    const result = await this.run<LiveWorkbookResult>({ action: "connect" })
    return result.workbook
  }

  async createSheet(sheetName: string): Promise<WorkbookSummary> {
    const result = await this.run<LiveWorkbookResult>({ action: "create_sheet", sheetName })
    return result.workbook
  }

  async writeCell(cell: string, value: string, sheetName?: string): Promise<WorkbookSummary> {
    const result = await this.run<LiveWorkbookResult>({ action: "write_cell", cell, value, sheetName })
    return result.workbook
  }

  async deleteCell(cell: string, sheetName?: string): Promise<WorkbookSummary> {
    const result = await this.run<LiveWorkbookResult>({ action: "delete_cell", cell, sheetName })
    return result.workbook
  }

  async writeTable(rows: string[][], startCell?: string, sheetName?: string): Promise<LiveWriteTableResult> {
    return this.run<LiveWriteTableResult>({ action: "write_table", rows, startCell, sheetName })
  }

  async appendRows(rows: string[][], sheetName?: string): Promise<LiveWriteTableResult> {
    return this.run<LiveWriteTableResult>({ action: "append_rows", rows, sheetName })
  }

  async deleteRow(rowNumber: number, sheetName?: string): Promise<WorkbookSummary> {
    const result = await this.run<LiveWorkbookResult>({ action: "delete_row", rowNumber, sheetName })
    return result.workbook
  }

  async readRange(range: string, sheetName?: string): Promise<{ workbook: WorkbookSummary; attachment: ChatAttachment }> {
    const result = await this.run<LiveReadRangeResult>({ action: "read_range", range, sheetName })
    return {
      workbook: result.workbook,
      attachment: {
        kind: "table",
        title: result.title,
        rows: result.rows,
      },
    }
  }

  async readUsedRange(sheetName?: string): Promise<{ workbook: WorkbookSummary; attachment: ChatAttachment }> {
    const result = await this.run<LiveReadRangeResult>({ action: "read_range", sheetName })
    return {
      workbook: result.workbook,
      attachment: {
        kind: "table",
        title: result.title,
        rows: result.rows,
      },
    }
  }

  private async run<T>(payload: Record<string, unknown>): Promise<T> {
    if (process.platform !== "win32") {
      throw new Error("Excel Live Mode는 현재 Windows에서만 지원됩니다.")
    }

    const payloadBase64 = Buffer.from(JSON.stringify(payload), "utf8").toString("base64")
    const script = buildPowerShellScript(payloadBase64)
    const encodedCommand = Buffer.from(script, "utf16le").toString("base64")

    this.logger.info("running excel live command", { action: payload.action })

    let stdout = ""
    let stderr = ""

    try {
      const result = await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-EncodedCommand",
        encodedCommand,
      ], { encoding: "buffer" })
      stdout = result.stdout.toString("utf8")
      stderr = result.stderr.toString("utf8")
    } catch (error) {
      const execError = error as NodeJS.ErrnoException & { stdout?: Buffer | string; stderr?: Buffer | string }
      const stderrText = bufferToUtf8(execError.stderr)
      const stdoutText = bufferToUtf8(execError.stdout)
      const errorText = cleanPowerShellError(stderrText.trim() || stdoutText.trim() || execError.message)
      throw new Error(`Excel 연결에 실패했습니다. ${errorText}`)
    }

    if (stderr.trim()) {
      this.logger.info("excel live stderr", stderr)
    }

    try {
      return JSON.parse(stdout.trim()) as T
    } catch {
      throw new Error(`Excel 응답을 해석하지 못했습니다. ${stdout.trim() || "응답 없음"}`)
    }
  }
}

function cleanPowerShellError(raw: string) {
  return raw
    .replace(/^#<\s*CLIXML\s*/i, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/_x000D__x000A_/g, " ")
    .replace(/\s+/g, " ")
    .trim()
}

function bufferToUtf8(value: Buffer | string | undefined) {
  if (!value) return ""
  return typeof value === "string" ? value : value.toString("utf8")
}

function buildPowerShellScript(payloadBase64: string) {
  return `
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
chcp 65001 > $null
$payloadJson = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('${payloadBase64}'))
$payload = $payloadJson | ConvertFrom-Json

function Get-ExcelApp {
  try {
    return [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    try {
      Add-Type -AssemblyName Microsoft.VisualBasic
      return [Microsoft.VisualBasic.Interaction]::GetObject('', 'Excel.Application')
    } catch {
      throw '실행 중인 Excel 인스턴스를 찾지 못했습니다. Excel과 workbook을 먼저 열고, Open Excel과 같은 권한 수준으로 실행해 주세요.'
    }
  }
}

function Get-WorkbookInfo($excel, $workbook) {
  if ($null -eq $workbook) {
    throw '열려 있는 workbook을 찾지 못했습니다. Excel에서 workbook을 먼저 열어 주세요.'
  }

  $sheetNames = @()
  foreach ($sheet in $workbook.Worksheets) {
    $sheetNames += [string]$sheet.Name
  }

  $fullName = [string]$workbook.FullName
  $name = [string]$workbook.Name
  $extension = [System.IO.Path]::GetExtension($fullName).ToLowerInvariant()
  $format = 'xlsx'
  if ($extension -eq '.csv') { $format = 'csv' }
  elseif ($extension -eq '.numbers') { $format = 'numbers' }

  return @{
    workbook = @{
      mode = 'live'
      name = $name
      path = $(if ($fullName) { $fullName } else { $name })
      format = $format
      sheetNames = $sheetNames
      activeSheetName = [string]$excel.ActiveSheet.Name
    }
  }
}

function Get-WorkbookAndSheet($excel, $sheetName) {
  $workbook = $excel.ActiveWorkbook
  if ($null -eq $workbook) {
    if ($excel.Workbooks.Count -gt 0) {
      $workbook = $excel.Workbooks.Item(1)
    } else {
      throw '현재 활성 workbook이 없습니다. Excel에서 workbook을 먼저 선택해 주세요.'
    }
  }

  $sheet = $(if ([string]::IsNullOrWhiteSpace($sheetName)) { $workbook.ActiveSheet } else { $workbook.Worksheets.Item($sheetName) })
  if ($null -eq $sheet) {
    throw '요청한 시트를 찾지 못했습니다.'
  }

  return @{ workbook = $workbook; sheet = $sheet }
}

function ColumnToNumber($column) {
  $sum = 0
  foreach ($character in $column.ToCharArray()) {
    $sum = ($sum * 26) + ([int][char]$character - [int][char]'A' + 1)
  }
  return $sum
}

function DecodeCell($address) {
  $normalized = Normalize-CellAddress([string]$address)
  $match = [regex]::Match($normalized, '^([A-Z]+)(\d+)$')
  if (-not $match.Success) {
    throw "유효하지 않은 셀 주소입니다: $address"
  }

  return @{
    Row = [int]$match.Groups[2].Value
    Col = ColumnToNumber($match.Groups[1].Value)
  }
}

function NumberToColumn($number) {
  $column = ''
  while ($number -gt 0) {
    $remainder = ($number - 1) % 26
    $column = [char]([int][char]'A' + $remainder) + $column
    $number = [math]::Floor(($number - 1) / 26)
  }
  return $column
}

function Get-DefaultWriteAnchor($sheet) {
  $used = Get-SafeUsedRange $sheet
  $isEmpty = $null -eq $used -or ($used.Rows.Count -eq 1 -and $used.Columns.Count -eq 1 -and [string]::IsNullOrEmpty([string]$sheet.Cells.Item(1,1).Text))
  if ($isEmpty) {
    return @{ Row = 1; Col = 1 }
  }

  return @{ Row = $used.Row + $used.Rows.Count; Col = 1 }
}

function Write-TableToSheet($sheet, $rows, $startRow, $startCol) {
  $rowCount = $rows.Count
  if ($rowCount -le 0) {
    throw '쓰기할 rows가 비어 있습니다.'
  }

  $colCount = 0
  for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
    $cells = $rows[$rowIndex]
    if ($null -ne $cells -and $cells.Count -gt $colCount) {
      $colCount = $cells.Count
    }
  }

  if ($colCount -le 0) {
    throw '쓰기할 컬럼이 비어 있습니다.'
  }

  $matrix = New-Object 'object[,]' $rowCount, $colCount
  for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
    $cells = $rows[$rowIndex]
    for ($colIndex = 0; $colIndex -lt $colCount; $colIndex++) {
      $value = ''
      if ($null -ne $cells -and $colIndex -lt $cells.Count -and $null -ne $cells[$colIndex]) {
        $value = [string]$cells[$colIndex]
      }
      $matrix[$rowIndex, $colIndex] = $value
    }
  }

  $targetRange = $sheet.Range(
    $sheet.Cells.Item($startRow, $startCol),
    $sheet.Cells.Item($startRow + $rowCount - 1, $startCol + $colCount - 1)
  )
  $targetRange.Value2 = $matrix

  return @{ Row = $startRow; Col = $startCol }
}

function Normalize-CellAddress($address) {
  return ([string]$address).Trim().ToUpperInvariant().Replace('$', '') -replace '[\(\)\[\]]', '' -replace '^.*!', '' -replace '\s+', ''
}

function Normalize-RangeAddress($range) {
  $normalized = ([string]$range).Trim().ToUpperInvariant().Replace('$', '') -replace '[\(\)\[\]]', '' -replace '\s+', ''
  if ($normalized -match '^[^!:]+!([A-Z]+\d+:[A-Z]+\d+)$') {
    return $matches[1]
  }
  if ($normalized -match '^[^!:]+!([A-Z]+\d+)$') {
    return $matches[1]
  }
  return $normalized
}

function ToRows($rangeValue) {
  if ($null -eq $rangeValue) { return @(@('')) }
  if ($rangeValue -isnot [System.Array]) { return @(@([string]$rangeValue)) }

  $rows = @()
  $rowLower = $rangeValue.GetLowerBound(0)
  $rowUpper = $rangeValue.GetUpperBound(0)
  $colLower = $rangeValue.GetLowerBound(1)
  $colUpper = $rangeValue.GetUpperBound(1)

  for ($row = $rowLower; $row -le $rowUpper; $row++) {
    $current = @()
    for ($col = $colLower; $col -le $colUpper; $col++) {
      $value = $rangeValue[$row, $col]
      $current += $(if ($null -eq $value) { '' } else { [string]$value })
    }
    $rows += ,$current
  }

  return $rows
}

function Get-RangeAddress($range) {
  if ($null -eq $range) {
    return 'A1'
  }
  return $range.Address($false, $false)
}

function Get-SafeUsedRange($sheet) {
  $used = $sheet.UsedRange
  if ($null -eq $used) {
    return $sheet.Range('A1')
  }
  return $used
}

$excel = Get-ExcelApp

switch ([string]$payload.action) {
  'connect' {
    $result = Get-WorkbookInfo $excel $excel.ActiveWorkbook
  }
  'create_sheet' {
    $context = Get-WorkbookAndSheet $excel $null
    $existing = $null
    try { $existing = $context.workbook.Worksheets.Item([string]$payload.sheetName) } catch {}
    if ($null -ne $existing) { throw '이미 같은 이름의 시트가 있습니다.' }
    [void]$context.workbook.Worksheets.Add()
    $excel.ActiveSheet.Name = [string]$payload.sheetName
    $result = Get-WorkbookInfo $excel $context.workbook
  }
  'write_cell' {
    $context = Get-WorkbookAndSheet $excel $payload.sheetName
    $context.sheet.Range((Normalize-CellAddress([string]$payload.cell))).Value2 = [string]$payload.value
    $result = Get-WorkbookInfo $excel $context.workbook
  }
  'delete_cell' {
    $context = Get-WorkbookAndSheet $excel $payload.sheetName
    $context.sheet.Range((Normalize-CellAddress([string]$payload.cell))).ClearContents()
    $result = Get-WorkbookInfo $excel $context.workbook
  }
  'write_table' {
    $context = Get-WorkbookAndSheet $excel $payload.sheetName
    $start = $(if ([string]::IsNullOrWhiteSpace([string]$payload.startCell)) { Get-DefaultWriteAnchor $context.sheet } else { DecodeCell((Normalize-CellAddress([string]$payload.startCell))) })
    $rows = $payload.rows
    $written = Write-TableToSheet $context.sheet $rows $start.Row $start.Col
    $result = Get-WorkbookInfo $excel $context.workbook
    $result.startCell = (NumberToColumn($written.Col)) + [string]$written.Row
  }
  'append_rows' {
    $context = Get-WorkbookAndSheet $excel $payload.sheetName
    $start = Get-DefaultWriteAnchor $context.sheet
    $rows = $payload.rows
    $written = Write-TableToSheet $context.sheet $rows $start.Row $start.Col
    $result = Get-WorkbookInfo $excel $context.workbook
    $result.startCell = (NumberToColumn($written.Col)) + [string]$written.Row
  }
  'delete_row' {
    $context = Get-WorkbookAndSheet $excel $payload.sheetName
    $context.sheet.Rows.Item([int]$payload.rowNumber).Delete()
    $result = Get-WorkbookInfo $excel $context.workbook
  }
  'read_range' {
    $context = Get-WorkbookAndSheet $excel $payload.sheetName
    $normalizedRange = [string]$payload.range
    $range = $(if ([string]::IsNullOrWhiteSpace($normalizedRange)) { Get-SafeUsedRange $context.sheet } else { $context.sheet.Range((Normalize-RangeAddress($normalizedRange))) })
    $info = Get-WorkbookInfo $excel $context.workbook
    $info.rows = ToRows($range.Value2)
    $info.title = ([string]$context.sheet.Name) + '!' + (Get-RangeAddress $range)
    $result = $info
  }
  default {
    throw "지원하지 않는 live action입니다: $($payload.action)"
  }
}

$result | ConvertTo-Json -Depth 20 -Compress
`
}
