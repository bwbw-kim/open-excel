import type { ChatAttachment } from "@shared/types"
import type { AuthService } from "../auth/auth-service"
import type {
  ConversationTurn,
  SpreadsheetAgentAction,
  SpreadsheetAgentObservation,
  SpreadsheetAgentStepRecord,
} from "../agent/types"

const CODEX_API_ENDPOINT = "https://chatgpt.com/backend-api/codex/responses"
const MODEL = "gpt-5.4-mini"

export async function planNextSpreadsheetAction(input: {
  authService: AuthService
  userRequest: string
  observation: SpreadsheetAgentObservation
  history: SpreadsheetAgentStepRecord[]
  conversation: ConversationTurn[]
  userAttachment?: ChatAttachment
  lastReadRangeTable?: ChatAttachment
  lastWebSearchTable?: ChatAttachment
  lastError?: string
}) {
  const instructions = [
    "You are a spreadsheet agent planner.",
    "Return only one JSON object.",
    "Choose exactly one action at a time.",
    "Treat the connected workbook observation as the single source of truth for spreadsheet operations.",
    "Allowed actions: read_range, write_range, answer.",
    "Tool guide:",
    "- read_range: read a specific range or the entire used range. Provide optional range (e.g., A1:B10) and optional sheetName. If range is omitted, reads the current sheet used range.",
    "- write_range: write multi-cell data to a specific range. Provide range (e.g., A1:D5) plus rows array and optional sheetName.",
    "- answer: when the task is complete, return a Korean answer.",
    "If the task is done, return action=answer with a short Korean answer.",
    "Never wrap JSON in markdown fences.",
  ].join("\n")

  const prompt = [
    `User request: ${input.userRequest}`,
    `Conversation: ${formatConversation(input.conversation)}`,
    `Workbook observation: ${JSON.stringify(input.observation, null, 2)}`,
    `User attachment: ${formatAttachment(input.userAttachment)}`,
    `Last read range table: ${formatAttachment(input.lastReadRangeTable)}`,
    `Last web search table: ${formatAttachment(input.lastWebSearchTable)}`,
    `History: ${formatHistory(input.history)}`,
    `Last error: ${input.lastError ?? "none"}`,
    "Return the next action JSON now.",
  ].join("\n\n")

  const raw = await requestCodexText(input.authService, instructions, prompt)
  return {
    raw,
    action: parseAction(raw),
  }
}

export async function writeSpreadsheetFinalAnswer(input: {
  authService: AuthService
  userRequest: string
  observation: SpreadsheetAgentObservation
  history: SpreadsheetAgentStepRecord[]
  conversation: ConversationTurn[]
  lastReadRangeTable?: ChatAttachment
  lastWebSearchTable?: ChatAttachment
}) {
  const instructions = [
    "You are a spreadsheet assistant.",
    "Answer in Korean.",
    "Use only information present in the observation and action history.",
    "Be concise but useful.",
  ].join("\n")

  const prompt = [
    `User request: ${input.userRequest}`,
    `Conversation: ${formatConversation(input.conversation)}`,
    `Observation: ${JSON.stringify(input.observation, null, 2)}`,
    `Last read range table: ${formatAttachment(input.lastReadRangeTable)}`,
    `Last web search table: ${formatAttachment(input.lastWebSearchTable)}`,
    `Action history: ${JSON.stringify(input.history, null, 2)}`,
    "Write the final answer.",
  ].join("\n\n")

  return requestCodexText(input.authService, instructions, prompt)
}

function formatHistory(history: SpreadsheetAgentStepRecord[]) {
  if (history.length === 0) return "none"
  return history.map((item) => `${item.step}. ${JSON.stringify(item.action)} => ${item.result}`).join("\n")
}

function formatConversation(conversation: ConversationTurn[]) {
  if (conversation.length === 0) return "none"
  return conversation.map((turn) => `${turn.role}: ${turn.message}`).join("\n")
}

function formatAttachment(attachment?: ChatAttachment) {
  if (!attachment) return "none"
  return JSON.stringify({
    title: attachment.title,
    rows: attachment.rows,
  })
}

async function requestCodexText(authService: AuthService, instructions: string, prompt: string) {
  const auth = await authService.ensureAuth()
  const response = await fetch(CODEX_API_ENDPOINT, {
    method: "POST",
    headers: {
      accept: "application/json, text/event-stream",
      "Content-Type": "application/json",
      authorization: `Bearer ${auth.accessToken}`,
      ...(auth.accountId ? { "ChatGPT-Account-Id": auth.accountId } : {}),
      originator: "open-excel",
      "User-Agent": `open-excel (${process.platform}; ${process.arch})`,
    },
    body: JSON.stringify({
      model: MODEL,
      stream: true,
      store: false,
      instructions,
      input: [
        {
          role: "user",
          content: [{ type: "input_text", text: prompt }],
        },
      ],
    }),
  })

  if (!response.ok) {
    throw new Error(`모델 요청에 실패했습니다: ${response.status}\n${await response.text()}`)
  }

  return extractTextFromSse(await response.text())
}

function extractTextFromSse(payload: string) {
  const deltas: string[] = []
  for (const line of payload.split("\n")) {
    if (!line.startsWith("data: ")) continue
    const raw = line.slice(6).trim()
    if (!raw || raw === "[DONE]") continue

    let parsed: Record<string, unknown>
    try {
      parsed = JSON.parse(raw) as Record<string, unknown>
    } catch {
      continue
    }

    if (parsed.type === "response.output_text.delta" && typeof parsed.delta === "string") {
      deltas.push(parsed.delta)
      continue
    }

    if (parsed.type === "error") {
      throw new Error(`모델 스트림 오류: ${String(parsed.message ?? "unknown")}`)
    }
  }

  const result = deltas.join("").trim()
  if (!result) {
    throw new Error("모델 응답 텍스트를 찾지 못했습니다.")
  }
  return result
}

function parseAction(raw: string): SpreadsheetAgentAction {
  const json = extractJsonObject(raw)
  const parsed = JSON.parse(json) as SpreadsheetAgentAction
  if (!parsed.action) {
    throw new Error(`액션 응답에 action이 없습니다: ${raw}`)
  }
  return parsed
}

function extractJsonObject(raw: string) {
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i)
  const source = fenced?.[1]?.trim() ?? raw.trim()
  const extracted = extractFirstBalancedJsonObject(source)
  if (!extracted) {
    throw new Error(`JSON 액션을 해석하지 못했습니다: ${raw}`)
  }
  return extracted
}

function extractFirstBalancedJsonObject(source: string) {
  const start = source.indexOf("{")
  if (start === -1) return undefined

  let depth = 0
  let inString = false
  let escaped = false

  for (let index = start; index < source.length; index += 1) {
    const character = source[index]

    if (inString) {
      if (escaped) {
        escaped = false
        continue
      }
      if (character === "\\") {
        escaped = true
        continue
      }
      if (character === '"') {
        inString = false
      }
      continue
    }

    if (character === '"') {
      inString = true
      continue
    }

    if (character === "{") depth += 1
    if (character === "}") {
      depth -= 1
      if (depth === 0) return source.slice(start, index + 1)
    }
  }

  return undefined
}
