import type { ChatAttachment, SheetPreview, WorkbookSummary } from "@/shared/types"

export type SpreadsheetAgentActionType =
  | "read_range"
  | "write_range"
  | "write_cell"
  | "delete_cell"
  | "write_table"
  | "append_rows"
  | "delete_row"
  | "create_sheet"
  | "web_search"
  | "answer"

export interface SpreadsheetAgentAction {
  action: SpreadsheetAgentActionType
  sheetName?: string
  cell?: string
  rowNumber?: number
  range?: string
  value?: string
  rows?: string[][]
  source?: "user_attachment" | "last_web_search"
  query?: string
  answer?: string
}

export interface SpreadsheetAgentObservation {
  mode: "live" | "file"
  connectedWorkbook: WorkbookSummary | null
  preview?: SheetPreview[]
  availableTables?: Array<{
    source: "user_attachment" | "last_web_search" | "last_read_range"
    title: string
    rows: number
    columns: number
  }>
  lastReadRangeTable?: ChatAttachment
}

export interface SpreadsheetAgentStepRecord {
  step: number
  action: SpreadsheetAgentAction
  result: string
}

export interface ConversationTurn {
  role: "user" | "assistant"
  message: string
}

export interface ToolExecutionContext {
  userAttachment?: ChatAttachment
  lastWebSearchTable?: ChatAttachment
  lastReadRangeTable?: ChatAttachment
}
