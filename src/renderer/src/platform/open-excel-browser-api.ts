import type { AuthState, ChatSessionState, DevState, OpenWorkbookResult, SendMessageInput } from "@shared/types"
import { BrowserChatSessionService } from "./browser-chat-session-service"
import { BrowserSpreadsheetService } from "./browser-spreadsheet-service"

const spreadsheetService = new BrowserSpreadsheetService()
const chatSessionService = new BrowserChatSessionService(spreadsheetService)

const api = {
  getDevState: async () => ({ isDev: window.location.hostname === "localhost" } satisfies DevState),
  getAuthState: () => fetchJson<AuthState>("/api/auth/state"),
  startLogin: async () => {
    const result = await fetchJson<{ authUrl: string }>("/api/auth/start", { method: "POST" })
    window.open(result.authUrl, "_blank", "noopener,noreferrer")
    return pollAuthState()
  },
  connectLiveWorkbook: async () => {
    const workbook = await spreadsheetService.connectLiveWorkbook()
    chatSessionService.attachWorkbook(workbook)
    return workbook
  },
  pickWorkbook: async () => null,
  getChatSession: async () => chatSessionService.getState(),
  sendChatMessage: async (payload: SendMessageInput) => chatSessionService.sendMessage(payload),
  openExternalUrl: async (url: string) => {
    window.open(url, "_blank", "noopener,noreferrer")
  },
}

export function installBrowserOpenExcelApi() {
  if (typeof window === "undefined") return
  if (window.openExcel) return
  window.openExcel = api as typeof window.openExcel
}

async function pollAuthState() {
  for (let attempt = 0; attempt < 60; attempt += 1) {
    const state = await fetchJson<AuthState>("/api/auth/state")
    if (state.authenticated) {
      return state
    }
    await new Promise((resolve) => window.setTimeout(resolve, 1000))
  }

  return fetchJson<AuthState>("/api/auth/state")
}

async function fetchJson<T>(input: string, init?: RequestInit): Promise<T> {
  const response = await fetch(input, {
    headers: {
      "Content-Type": "application/json",
      ...(init?.headers ?? {}),
    },
    ...init,
  })

  if (!response.ok) {
    throw new Error(await response.text())
  }

  return (await response.json()) as T
}
