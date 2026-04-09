import { app, BrowserWindow } from "electron"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { registerAppIpc } from "./ipc/register-app-ipc"
import { createLogger } from "./logging/logger"

app.disableHardwareAcceleration()
app.commandLine.appendSwitch("disable-gpu")
app.commandLine.appendSwitch("disable-software-rasterizer")

const isDev = !app.isPackaged
const logger = createLogger(isDev)
const currentDir = path.dirname(fileURLToPath(import.meta.url))

function createMainWindow() {
  const window = new BrowserWindow({
    width: 1040,
    height: 920,
    minWidth: 760,
    minHeight: 760,
    backgroundColor: "#0f172a",
    title: "Open Excel",
    webPreferences: {
      preload: path.join(currentDir, "../preload/index.mjs"),
      contextIsolation: true,
      sandbox: false,
      nodeIntegration: false,
    },
  })

  registerAppIpc(window, logger)

  if (isDev && process.env["ELECTRON_RENDERER_URL"]) {
    window.loadURL(process.env["ELECTRON_RENDERER_URL"])
    return window
  }

  window.loadFile(path.join(currentDir, "../renderer/index.html"))
  return window
}

app.whenReady().then(() => {
  createMainWindow()

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow()
    }
  })
})

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit()
  }
})
