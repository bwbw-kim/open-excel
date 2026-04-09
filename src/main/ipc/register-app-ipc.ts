import { dialog, ipcMain, shell, type BrowserWindow } from "electron"
import { AuthService } from "../auth/auth-service"
import { ChatSessionService } from "../chat/chat-session-service"
import type { Logger } from "../logging/logger"
import { SpreadsheetService } from "../spreadsheet/spreadsheet-service"

let registered = false

export function registerAppIpc(window: BrowserWindow, logger: Logger) {
  if (registered) {
    return
  }

  registered = true
  const authService = new AuthService(logger)
  const spreadsheetService = new SpreadsheetService(logger)
  const chatSessionService = new ChatSessionService(logger, authService, spreadsheetService)

  ipcMain.handle("app:get-dev-state", async () => ({
    isDev: logger.enabled,
  }))

  ipcMain.handle("auth:get-state", async () => authService.getState())
  ipcMain.handle("auth:start-login", async () => authService.startLogin())

  ipcMain.handle("excel:connect-live", async () => {
    const workbook = await spreadsheetService.connectLiveWorkbook()
    chatSessionService.attachWorkbook(workbook)
    return workbook
  })

  ipcMain.handle("file:pick-workbook", async () => {
    const result = await dialog.showOpenDialog(window, {
      title: "Open spreadsheet",
      properties: ["openFile"],
      filters: [
        {
          name: "Spreadsheet files",
          extensions: ["xlsx", "csv", "numbers"],
        },
      ],
    })

    if (result.canceled || result.filePaths.length === 0) {
      return null
    }

    const openedWorkbook = await spreadsheetService.openWorkbook(result.filePaths[0])
    chatSessionService.attachWorkbook(openedWorkbook)
    return openedWorkbook
  })

  ipcMain.handle("chat:get-session", async () => chatSessionService.getState())
  ipcMain.handle("chat:send-message", async (_event, payload) => {
    await chatSessionService.sendMessage(payload)
    if (spreadsheetService.getActiveWorkbook()) {
      chatSessionService.syncWorkbookPreview(spreadsheetService.getActiveWorkbookData())
    }
    return chatSessionService.getState()
  })

  ipcMain.handle("external:open-url", async (_event, url: string) => {
    await shell.openExternal(url)
  })
}
