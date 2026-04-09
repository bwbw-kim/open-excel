import type { ChatAttachment, ChatMessage, ChatSessionState, OpenWorkbookResult, SendMessageInput } from "@shared/types"
import type { AuthService } from "../auth/auth-service"
import { SpreadsheetAgentService } from "../agent/spreadsheet-agent-service"
import type { Logger } from "../logging/logger"
import type { SpreadsheetService } from "../spreadsheet/spreadsheet-service"

export class ChatSessionService {
  private readonly state: ChatSessionState = {
    sessionId: crypto.randomUUID(),
    messages: [
      createMessage(
        "assistant",
        "안녕하세요. 실행 중인 Excel에 연결한 뒤 원하는 변경이나 분석을 요청해 주세요.",
      ),
    ],
  }

  private readonly spreadsheetAgentService: SpreadsheetAgentService

  constructor(
    private readonly logger: Logger,
    authService: AuthService,
    spreadsheetService: SpreadsheetService,
  ) {
    this.spreadsheetAgentService = new SpreadsheetAgentService(authService, spreadsheetService, logger)
  }

  async getState(): Promise<ChatSessionState> {
    return this.state
  }

  attachWorkbook(result: OpenWorkbookResult) {
    this.state.activeWorkbook = result.workbook
    this.state.preview = result.preview
  }

  async sendMessage(input: SendMessageInput): Promise<ChatSessionState> {
    const userMessage = createMessage("user", input.message, input.attachment)
    this.state.messages.push(userMessage)

    try {
      const result = await this.spreadsheetAgentService.handleUserMessage(input, this.toConversation())
      this.state.messages.push(createMessage("assistant", result.reply, result.attachment))
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      this.logger.error("chat message handling failed", { message })
      this.state.messages.push(createMessage("assistant", message))
    }

    return this.state
  }

  syncWorkbookPreview(result: OpenWorkbookResult) {
    this.state.activeWorkbook = result.workbook
    this.state.preview = result.preview
  }

  private toConversation() {
    return this.state.messages.map((message) => ({
      role: message.role === "system" ? "assistant" : message.role,
      message: message.content,
    }))
  }
}

function createMessage(role: ChatMessage["role"], content: string, attachment?: ChatAttachment): ChatMessage {
  return {
    id: crypto.randomUUID(),
    role,
    content,
    createdAt: Date.now(),
    attachment,
  }
}
