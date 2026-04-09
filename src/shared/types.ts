export type MessageRole = "user" | "assistant" | "system"

export interface ChatMessage {
  id: string
  role: MessageRole
  content: string
  createdAt: number
  attachment?: ChatAttachment
}

export interface ChatAttachment {
  kind: "table"
  title: string
  rows: string[][]
}

export interface ChatSessionState {
  sessionId: string
  messages: ChatMessage[]
  activeWorkbook?: WorkbookSummary
  preview?: SheetPreview[]
}

export interface WorkbookSummary {
  mode: "file" | "live"
  name: string
  path: string
  format: SpreadsheetFormat
  sheetNames: string[]
  activeSheetName?: string
}

export type SpreadsheetFormat = "xlsx" | "csv" | "numbers"

export interface SheetPreview {
  sheetName: string
  rows: string[][]
}

export interface OpenWorkbookResult {
  workbook: WorkbookSummary
  preview: SheetPreview[]
}

export interface ParsedClipboardTable {
  rows: string[][]
  plainText: string
}

export interface AuthState {
  authenticated: boolean
  accountId?: string
  expiresAt?: number
}

export interface SendMessageInput {
  message: string
  attachment?: ChatAttachment
}

export interface DevState {
  isDev: boolean
}
