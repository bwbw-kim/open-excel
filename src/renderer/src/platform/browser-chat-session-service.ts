import type { ChatAttachment, ChatMessage, ChatSessionState, OpenWorkbookResult, SendMessageInput } from "@shared/types"
import type { SpreadsheetAgentObservation } from "@main/agent/types"
import { BrowserSpreadsheetService } from "./browser-spreadsheet-service"
import { planNextSpreadsheetAction, writeSpreadsheetFinalAnswer } from "./codex-client"

const MAX_STEPS = 20
const MAX_EXECUTION_MS = 3 * 60 * 1000

export class BrowserChatSessionService {
  private readonly state: ChatSessionState = {
    sessionId: crypto.randomUUID(),
    messages: [createMessage("assistant", "안녕하세요. 현재 열려 있는 Excel에 연결한 뒤 원하는 변경이나 분석을 요청해 주세요.")],
  }

  constructor(private readonly spreadsheetService: BrowserSpreadsheetService) {}

  async getState(): Promise<ChatSessionState> {
    return this.state
  }

  attachWorkbook(result: OpenWorkbookResult) {
    this.state.activeWorkbook = result.workbook
    this.state.preview = result.preview
  }

  async sendMessage(input: SendMessageInput): Promise<ChatSessionState> {
    const trimmedRequest = input.message.trim()
    if (!trimmedRequest) {
      return this.state
    }

    this.state.messages.push(createMessage("user", trimmedRequest, input.attachment))

    const history: Array<{ step: number; action: unknown; result: string }> = []
    const toolContext: { userAttachment?: ChatAttachment; lastWebSearchTable?: ChatAttachment; lastReadRangeTable?: ChatAttachment } = {
      userAttachment: input.attachment,
    }
    let lastError: string | undefined
    const startedAt = Date.now()

    for (let step = 1; step <= MAX_STEPS; step += 1) {
      if (Date.now() - startedAt > MAX_EXECUTION_MS) {
        this.state.messages.push(createMessage("assistant", "작업 시간이 초과되었습니다."))
        return this.state
      }

      const observation: SpreadsheetAgentObservation = {
        workbook: this.state.activeWorkbook,
        preview: this.state.preview ?? [],
        availableTables: [
          toolContext.userAttachment ? { source: "user_attachment" as const, title: toolContext.userAttachment.title, rows: toolContext.userAttachment.rows.length, columns: toolContext.userAttachment.rows[0]?.length ?? 0 } : undefined,
          toolContext.lastWebSearchTable ? { source: "last_web_search" as const, title: toolContext.lastWebSearchTable.title, rows: toolContext.lastWebSearchTable.rows.length, columns: toolContext.lastWebSearchTable.rows[0]?.length ?? 0 } : undefined,
          toolContext.lastReadRangeTable ? { source: "last_read_range" as const, title: toolContext.lastReadRangeTable.title, rows: toolContext.lastReadRangeTable.rows.length, columns: toolContext.lastReadRangeTable.rows[0]?.length ?? 0 } : undefined,
        ].filter((value): value is NonNullable<typeof value> => Boolean(value)),
      }
      const planned = await planNextSpreadsheetAction({
        userRequest: trimmedRequest,
        observation,
        history,
        conversation: this.state.messages
          .filter((message) => message.role === "user" || message.role === "assistant")
          .map((message) => ({ role: message.role as "user" | "assistant", message: message.content })),
        userAttachment: input.attachment,
        lastReadRangeTable: toolContext.lastReadRangeTable,
        lastWebSearchTable: toolContext.lastWebSearchTable,
        lastError,
      })

      if (planned.action.action === "read_range") {
        const attachment = planned.action.range?.trim()
          ? await this.spreadsheetService.readRange(planned.action.range, planned.action.sheetName)
          : await this.spreadsheetService.readUsedRange(planned.action.sheetName)
        toolContext.lastReadRangeTable = attachment.attachment
        this.state.activeWorkbook = attachment.workbook
        history.push({ step, action: planned.action, result: `${attachment.attachment.title} 범위를 읽었습니다.` })
        lastError = undefined
        continue
      }

      if (planned.action.action === "write_range") {
        const range = planned.action.range?.trim()
        if (!range) {
          lastError = "range 값이 필요합니다."
          history.push({ step, action: planned.action, result: `ERROR: ${lastError}` })
          continue
        }
        const rows = planned.action.rows?.length ? planned.action.rows : input.attachment?.rows ?? []
        const resultText = await this.spreadsheetService.writeRange(range, rows, planned.action.sheetName)
        history.push({ step, action: planned.action, result: resultText })
        const workbook = await this.spreadsheetService.getActiveWorkbookData()
        this.state.activeWorkbook = workbook.workbook
        this.state.preview = workbook.preview
        lastError = undefined
        continue
      }

      if (planned.action.action === "answer") {
        const reply =
          planned.action.answer?.trim() ||
          (await writeSpreadsheetFinalAnswer({
            userRequest: trimmedRequest,
            observation,
            history,
            conversation: this.state.messages
              .filter((message) => message.role === "user" || message.role === "assistant")
              .map((message) => ({ role: message.role as "user" | "assistant", message: message.content })),
            lastReadRangeTable: toolContext.lastReadRangeTable,
            lastWebSearchTable: toolContext.lastWebSearchTable,
          }))
        this.state.messages.push(createMessage("assistant", reply, toolContext.lastReadRangeTable ?? toolContext.lastWebSearchTable))
        return this.state
      }
    }

    this.state.messages.push(createMessage("assistant", "작업을 완료했습니다."))
    return this.state
  }
}

function createMessage(role: ChatMessage["role"], content: string, attachment?: ChatAttachment): ChatMessage {
  return { id: crypto.randomUUID(), role, content, createdAt: Date.now(), attachment }
}
