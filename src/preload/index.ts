import { contextBridge, ipcRenderer } from "electron"
import type { AuthState, ChatSessionState, DevState, OpenWorkbookResult, SendMessageInput } from "@shared/types"

const api = {
  getDevState: () => ipcRenderer.invoke("app:get-dev-state") as Promise<DevState>,
  getAuthState: () => ipcRenderer.invoke("auth:get-state") as Promise<AuthState>,
  startLogin: () => ipcRenderer.invoke("auth:start-login") as Promise<AuthState>,
  connectLiveWorkbook: () => ipcRenderer.invoke("excel:connect-live") as Promise<OpenWorkbookResult>,
  pickWorkbook: () => ipcRenderer.invoke("file:pick-workbook") as Promise<OpenWorkbookResult | null>,
  getChatSession: () => ipcRenderer.invoke("chat:get-session") as Promise<ChatSessionState>,
  sendChatMessage: (payload: SendMessageInput) => ipcRenderer.invoke("chat:send-message", payload) as Promise<ChatSessionState>,
  openExternalUrl: (url: string) => ipcRenderer.invoke("external:open-url", url) as Promise<void>,
}

contextBridge.exposeInMainWorld("openExcel", api)

declare global {
  interface Window {
    openExcel: typeof api
  }
}
