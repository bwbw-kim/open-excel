# Excel Live Mode 테스트 스크립트
# 사용법: powershell -ExecutionPolicy Bypass -File test-excel-live.ps1

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Get-ExcelApp {
  try {
    return [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    try {
      Add-Type -AssemblyName Microsoft.VisualBasic
      return [Microsoft.VisualBasic.Interaction]::GetObject('', 'Excel.Application')
    } catch {
      throw '실행 중인 Excel 인스턴스를 찾지 못했습니다.'
    }
  }
}

function Get-WorkbookAndSheet($excel, $sheetName) {
  $workbook = $excel.ActiveWorkbook
  if ($null -eq $workbook) {
    if ($excel.Workbooks.Count -gt 0) {
      $workbook = $excel.Workbooks.Item(1)
    } else {
      throw '현재 활성 workbook이 없습니다.'
    }
  }

  $sheet = $(if ([string]::IsNullOrWhiteSpace($sheetName)) { $workbook.ActiveSheet } else { $workbook.Worksheets.Item($sheetName) })
  if ($null -eq $sheet) {
    throw '요청한 시트를 찾지 못했습니다.'
  }

  return @{ workbook = $workbook; sheet = $sheet }
}

function ToRows($rangeValue) {
  if ($null -eq $rangeValue) { return ,@('') }
  if ($rangeValue -isnot [System.Array]) { return ,@([string]$rangeValue) }

  if ($rangeValue.Rank -lt 2) {
    $current = @()
    $lower = $rangeValue.GetLowerBound(0)
    $upper = $rangeValue.GetUpperBound(0)

    for ($index = $lower; $index -le $upper; $index++) {
      $value = $rangeValue.GetValue($index)
      $current += $(if ($null -eq $value) { '' } else { [string]$value })
    }

    return ,$current
  }

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

function Write-TableToSheet($sheet, $rows, $startRow, $startCol) {
  $rowCount = $rows.Count
  if ($rowCount -le 0) { throw '쓰기할 rows가 비어 있습니다.' }

  $colCount = 0
  for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
    $cells = $rows[$rowIndex]
    if ($null -ne $cells -and $cells.Count -gt $colCount) {
      $colCount = $cells.Count
    }
  }

  if ($colCount -le 0) { throw '쓰기할 컬럼이 비어 있습니다.' }

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

function DecodeCell($address) {
  $normalized = ([string]$address).Trim().ToUpperInvariant().Replace('$', '') -replace '[\(\)\[\]]', '' -replace '^.*!', '' -replace '\s+', ''
  $match = [regex]::Match($normalized, '^([A-Z]+)(\d+)$')
  if (-not $match.Success) { throw "유효하지 않은 셀 주소입니다: $address" }

  $col = 0
  foreach ($character in $match.Groups[1].Value.ToCharArray()) {
    $col = ($col * 26) + ([int][char]$character - [int][char]'A' + 1)
  }

  return @{ Row = [int]$match.Groups[2].Value; Col = $col }
}

function Get-RangeAddress($range) {
  if ($null -eq $range) { return 'A1' }
  $address = [string]$range.Address()
  return $address -replace '\$', ''
}

function Get-SafeUsedRange($sheet) {
  $used = $sheet.UsedRange
  if ($null -eq $used) { return $sheet.Range('A1') }
  return $used
}

# 테스트 시작
Write-Host "=== Excel Live Mode 테스트 ===" -ForegroundColor Cyan

try {
  $excel = Get-ExcelApp
  Write-Host "Excel 연결 성공" -ForegroundColor Green
  
  $context = Get-WorkbookAndSheet $excel $null
  $sheet = $context.sheet
  Write-Host "현재 시트: $($sheet.Name)" -ForegroundColor Yellow
  
  # 테스트 1: read_range - A1:B5
  Write-Host "`n--- 테스트 1: read_range A1:B5 ---" -ForegroundColor Cyan
  $range = $sheet.Range("A1:B5")
  $rows = ToRows($range.Value2)
  Write-Host "읽은 데이터:"
  $rows | ForEach-Object { Write-Host ($_ -join ", ") }
  
  # 테스트 2: read_range - 전체 사용 범위
  Write-Host "`n--- 테스트 2: read_range 전체 ---" -ForegroundColor Cyan
  $usedRange = Get-SafeUsedRange $sheet
  $allRows = ToRows($usedRange.Value2)
  Write-Host "읽은 데이터:"
  $allRows | ForEach-Object { Write-Host ($_ -join ", ") }
  
  # 테스트 3: write_range - D1:E3에 쓰기
  Write-Host "`n--- 테스트 3: write_range D1:E3 ---" -ForegroundColor Cyan
  $testRows = @(
    @( " name", "age"),
    @( "kim", "25"),
    @( "lee", "30")
  )
  $written = Write-TableToSheet $sheet $testRows 1 4  # D1부터
  Write-Host "쓰기 완료: $($written.Row), $($written.Col)" -ForegroundColor Green
  
  # 테스트 4: 읽기 확인
  Write-Host "`n--- 테스트 4: 쓰기 후 확인 ---" -ForegroundColor Cyan
  $verifyRange = $sheet.Range("D1:E3")
  $verifyRows = ToRows($verifyRange.Value2)
  Write-Host "읽은 데이터:"
  $verifyRows | ForEach-Object { Write-Host ($_ -join ", ") }
  
  Write-Host "`n=== 모든 테스트 완료 ===" -ForegroundColor Green

} catch {
  Write-Host "오류: $($_.Exception.Message)" -ForegroundColor Red
}