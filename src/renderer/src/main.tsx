import React from "react"
import ReactDOM from "react-dom/client"
import { App } from "./App"
import "./styles.css"
import { installBrowserOpenExcelApi } from "./platform/open-excel-browser-api"

installBrowserOpenExcelApi()


function renderApp() {
  ReactDOM.createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  )
}

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  void Office.onReady(() => renderApp())
} else {
  renderApp()
}
