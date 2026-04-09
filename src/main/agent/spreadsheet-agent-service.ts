import type { ChatAttachment, SendMessageInput } from "@shared/types"
import type { AuthService } from "../auth/auth-service"
import type { Logger } from "../logging/logger"
import { planNextSpreadsheetAction, writeSpreadsheetFinalAnswer } from "../llm/spreadsheet-llm"
import type {
  ConversationTurn,
  SpreadsheetAgentAction,
  SpreadsheetAgentStepRecord,
  ToolExecutionContext,
} from "./types"
import type { SpreadsheetService } from "../spreadsheet/spreadsheet-service"
import { WebSearchService } from "../web/web-search-service"

const MAX_STEPS = 20
const MAX_EXECUTION_MS = 3 * 60 * 1000

interface AgentResult {
  reply: string
  attachment?: ChatAttachment
}

export class SpreadsheetAgentService {
  private readonly webSearchService = new WebSearchService()

  constructor(
    private readonly authService: AuthService,
    private readonly spreadsheetService: SpreadsheetService,
    private readonly logger: Logger,
  ) {}

  async handleUserMessage(input: SendMessageInput, conversation: ConversationTurn[]): Promise<AgentResult> {
    const trimmedRequest = input.message.trim()
    if (!trimmedRequest) {
      return { reply: "입력이 비어 있습니다." }
    }

    const history: SpreadsheetAgentStepRecord[] = []
    const toolContext: ToolExecutionContext = {
      userAttachment: input.attachment,
    }
    let lastError: string | undefined
    const startedAt = Date.now()

    for (let step = 1; step <= MAX_STEPS; step += 1) {
      if (Date.now() - startedAt > MAX_EXECUTION_MS) {
        return {
          reply: await writeSpreadsheetFinalAnswer({
            authService: this.authService,
            userRequest: trimmedRequest,
            observation: this.spreadsheetService.getObservation(toolContext),
            history: [
              ...history,
              {
                step,
                action: { action: "answer", answer: "" },
                result: "TIMEOUT: Agent execution exceeded 3 minutes.",
              },
            ],
            conversation,
            lastReadRangeTable: toolContext.lastReadRangeTable,
            lastWebSearchTable: toolContext.lastWebSearchTable,
          }),
          attachment: toolContext.lastWebSearchTable,
        }
      }

      const observation = this.spreadsheetService.getObservation(toolContext)
      const planned = await planNextSpreadsheetAction({
        authService: this.authService,
        userRequest: trimmedRequest,
        observation,
        history,
        conversation,
        userAttachment: input.attachment,
        lastReadRangeTable: toolContext.lastReadRangeTable,
        lastWebSearchTable: toolContext.lastWebSearchTable,
        lastError,
      })

      this.logger.info("planner raw response", planned.raw)

      if (planned.action.action === "read_range" && isRedundantReadAction(planned.action, toolContext.lastReadRangeTable)) {
        return {
          reply: await writeSpreadsheetFinalAnswer({
            authService: this.authService,
            userRequest: trimmedRequest,
            observation,
            history,
            conversation,
            lastReadRangeTable: toolContext.lastReadRangeTable,
            lastWebSearchTable: toolContext.lastWebSearchTable,
          }),
          attachment: toolContext.lastReadRangeTable,
        }
      }

      if (planned.action.action === "answer") {
        return {
          reply:
            planned.action.answer?.trim() ||
            (await writeSpreadsheetFinalAnswer({
              authService: this.authService,
              userRequest: trimmedRequest,
              observation,
              history,
              conversation,
              lastReadRangeTable: toolContext.lastReadRangeTable,
              lastWebSearchTable: toolContext.lastWebSearchTable,
            })),
          attachment: toolContext.lastReadRangeTable ?? toolContext.lastWebSearchTable,
        }
      }

      try {
        if (shouldPreReadBeforeMutation(planned.action, toolContext)) {
          const attachment = await this.spreadsheetService.readUsedRange(planned.action.sheetName)
          toolContext.lastReadRangeTable = attachment
          history.push({
            step,
            action: { action: "read_range", sheetName: planned.action.sheetName },
            result: `${attachment.title} 범위를 선행 확인했습니다.`,
          })
          lastError = undefined
          continue
        }

        const result = await this.executeAction(planned.action, toolContext)
        this.logger.info("tool action executed", {
          step,
          action: planned.action.action,
          result: result.resultText,
        })
        history.push({ step, action: planned.action, result: result.resultText })
        if (planned.action.action === "read_range" && result.generatedAttachment) {
          toolContext.lastReadRangeTable = result.generatedAttachment
        }
        if (planned.action.action === "web_search" && result.generatedAttachment) {
          toolContext.lastWebSearchTable = result.generatedAttachment
        }
        lastError = undefined
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error)
        history.push({ step, action: planned.action, result: `ERROR: ${message}` })
        lastError = message
      }
    }

    return {
      reply: await writeSpreadsheetFinalAnswer({
        authService: this.authService,
        userRequest: trimmedRequest,
        observation: this.spreadsheetService.getObservation(toolContext),
        history,
        conversation,
        lastReadRangeTable: toolContext.lastReadRangeTable,
        lastWebSearchTable: toolContext.lastWebSearchTable,
      }),
      attachment: toolContext.lastReadRangeTable ?? toolContext.lastWebSearchTable,
    }
  }

  private async executeAction(action: SpreadsheetAgentAction, toolContext: ToolExecutionContext) {
    switch (action.action) {
      case "read_range": {
        const attachment = action.range?.trim()
          ? await this.spreadsheetService.readRange(action.range, action.sheetName)
          : await this.spreadsheetService.readUsedRange(action.sheetName)
        return {
          resultText: `${attachment.title} 범위를 읽었습니다.`,
          generatedAttachment: attachment,
        }
      }
      case "write_range": {
        const range = requireString(action.range, "range")
        const rows = resolveRows(action, toolContext)
        return { resultText: await this.spreadsheetService.writeRange(range, rows, action.sheetName) }
      }
      default:
        throw new Error(`지원하지 않는 액션입니다: ${String(action.action)}`)
    }
  }
}

function requireString(value: string | undefined, fieldName: string) {
  if (!value?.trim()) {
    throw new Error(`${fieldName} 값이 필요합니다.`)
  }
  return value.trim()
}

function requireRows(rows: string[][] | undefined) {
  if (!rows || rows.length === 0) {
    throw new Error("rows 값이 필요합니다.")
  }
  return rows
}

function requireRowNumber(rowNumber: number | undefined) {
  if (!rowNumber || rowNumber < 1) {
    throw new Error("rowNumber 값이 필요합니다.")
  }
  return rowNumber
}

function resolveRows(action: SpreadsheetAgentAction, toolContext: ToolExecutionContext) {
  if (action.rows?.length) return action.rows
  if (action.source === "user_attachment") {
    if (!toolContext.userAttachment) {
      throw new Error("붙여넣은 표가 없습니다.")
    }
    return toolContext.userAttachment.rows
  }
  if (action.source === "last_web_search") {
    if (!toolContext.lastWebSearchTable) {
      throw new Error("최근 웹 검색 표가 없습니다.")
    }
    return toolContext.lastWebSearchTable.rows
  }
  throw new Error("write_table/write_range에는 rows 또는 source가 필요합니다.")
}

function shouldPreReadBeforeMutation(action: SpreadsheetAgentAction, toolContext: ToolExecutionContext) {
  // read_range, write_range만 있으므로 사전 읽기 불필요
  return false
}

function isRedundantReadAction(action: SpreadsheetAgentAction, lastReadRangeTable?: ChatAttachment) {
  if (action.action !== "read_range" || !lastReadRangeTable) {
    return false
  }

  const parsed = parseAttachmentTitle(lastReadRangeTable.title)
  if (!parsed) {
    return false
  }

  const requestedSheet = action.sheetName?.trim()
  const requestedRange = action.range?.trim()

  // range가 명시적으로 주어진 경우만 중복 체크
  // range 없이 전체 시트를 읽으려는 경우는 중복이 아니라 무조건 다시 읽어야 함
  if (!requestedRange) {
    return false
  }

  return normalizeRangeAddress(requestedRange) === normalizeRangeAddress(parsed.range) && (!requestedSheet || requestedSheet === parsed.sheetName)
}

function parseAttachmentTitle(title: string) {
  const separator = title.indexOf("!")
  if (separator === -1) {
    return undefined
  }

  return {
    sheetName: title.slice(0, separator).trim(),
    range: title.slice(separator + 1).trim(),
  }
}

function normalizeRangeAddress(range: string) {
  return range
    .trim()
    .replace(/[()\[\]]/g, "")
    .replace(/\$/g, "")
    .replace(/^.*!/g, "")
    .replace(/\s+/g, "")
    .toUpperCase()
}
