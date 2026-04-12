import type { ChatAttachment } from "@shared/types"
import type { SpreadsheetAgentAction, SpreadsheetAgentObservation } from "@main/agent/types"

const CODEX_API_ENDPOINT = "https://chatgpt.com/backend-api/codex/responses"
const MODEL = "gpt-5.4-mini"

export async function planNextSpreadsheetAction(input: {
  userRequest: string
  observation: SpreadsheetAgentObservation
  history: Array<{ step: number; action: unknown; result: string }>
  conversation: Array<{ role: "user" | "assistant"; message: string }>
  userAttachment?: ChatAttachment
  lastReadRangeTable?: ChatAttachment
  lastWebSearchTable?: ChatAttachment
  lastError?: string
}) {
  const raw = await requestCodexText(
    [
      "You are a spreadsheet agent planner.",
      "Return only one JSON object.",
      "Choose exactly one action at a time.",
      "Allowed actions: read_range, write_range, answer.",
      "Never wrap JSON in markdown fences.",
    ].join("\n"),
    [
      `User request: ${input.userRequest}`,
      `Conversation: ${formatConversation(input.conversation)}`,
      `Workbook observation: ${JSON.stringify(input.observation, null, 2)}`,
      `User attachment: ${formatAttachment(input.userAttachment)}`,
      `Last read range table: ${formatAttachment(input.lastReadRangeTable)}`,
      `Last web search table: ${formatAttachment(input.lastWebSearchTable)}`,
      `History: ${JSON.stringify(input.history, null, 2)}`,
      `Last error: ${input.lastError ?? "none"}`,
      "Return the next action JSON now.",
    ].join("\n\n"),
  )

  return { raw, action: parseAction(raw) }
}

export async function writeSpreadsheetFinalAnswer(input: {
  userRequest: string
  observation: SpreadsheetAgentObservation
  history: Array<{ step: number; action: unknown; result: string }>
  conversation: Array<{ role: "user" | "assistant"; message: string }>
  lastReadRangeTable?: ChatAttachment
  lastWebSearchTable?: ChatAttachment
}) {
  return requestCodexText(
    ["You are a spreadsheet assistant.", "Answer in Korean.", "Be concise but useful."].join("\n"),
    [
      `User request: ${input.userRequest}`,
      `Conversation: ${formatConversation(input.conversation)}`,
      `Observation: ${JSON.stringify(input.observation, null, 2)}`,
      `Last read range table: ${formatAttachment(input.lastReadRangeTable)}`,
      `Last web search table: ${formatAttachment(input.lastWebSearchTable)}`,
      `Action history: ${JSON.stringify(input.history, null, 2)}`,
      "Write the final answer.",
    ].join("\n\n"),
  )
}

async function requestCodexText(instructions: string, prompt: string) {
  const auth = await fetchJson<{ accessToken: string; accountId?: string }>("/api/auth/token")
  const response = await fetch(CODEX_API_ENDPOINT, {
    method: "POST",
    headers: {
      accept: "application/json, text/event-stream",
      "Content-Type": "application/json",
      authorization: `Bearer ${auth.accessToken}`,
      ...(auth.accountId ? { "ChatGPT-Account-Id": auth.accountId } : {}),
      originator: "open-excel",
      "User-Agent": `open-excel (${window.navigator.platform})`,
    },
    body: JSON.stringify({
      model: MODEL,
      stream: true,
      store: false,
      instructions,
      input: [{ role: "user", content: [{ type: "input_text", text: prompt }] }],
    }),
  })

  if (!response.ok) {
    throw new Error(`모델 요청에 실패했습니다: ${response.status}\n${await response.text()}`)
  }

  return extractTextFromSse(await response.text())
}

async function fetchJson<T>(input: string): Promise<T> {
  const response = await fetch(input)
  if (!response.ok) throw new Error(await response.text())
  return (await response.json()) as T
}

function formatHistory(history: Array<{ step: number; action: unknown; result: string }>) {
  if (history.length === 0) return "none"
  return history.map((item) => `${item.step}. ${JSON.stringify(item.action)} => ${item.result}`).join("\n")
}

function formatConversation(conversation: Array<{ role: "user" | "assistant"; message: string }>) {
  if (conversation.length === 0) return "none"
  return conversation.map((turn) => `${turn.role}: ${turn.message}`).join("\n")
}

function formatAttachment(attachment?: ChatAttachment) {
  if (!attachment) return "none"
  return JSON.stringify({ title: attachment.title, rows: attachment.rows })
}

function extractTextFromSse(payload: string) {
  const deltas: string[] = []
  for (const line of payload.split("\n")) {
    if (!line.startsWith("data: ")) continue
    const raw = line.slice(6).trim()
    if (!raw || raw === "[DONE]") continue
    try {
      const parsed = JSON.parse(raw) as Record<string, unknown>
      if (parsed.type === "response.output_text.delta" && typeof parsed.delta === "string") {
        deltas.push(parsed.delta)
      }
    } catch {
      continue
    }
  }

  const result = deltas.join("").trim()
  if (!result) throw new Error("모델 응답 텍스트를 찾지 못했습니다.")
  return result
}

function parseAction(raw: string): SpreadsheetAgentAction {
  const extracted = raw.match(/\{[\s\S]*\}/)?.[0]
  if (!extracted) throw new Error(`JSON 액션을 해석하지 못했습니다: ${raw}`)
  return JSON.parse(extracted) as SpreadsheetAgentAction
}
