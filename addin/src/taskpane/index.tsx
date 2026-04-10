import { StrictMode } from "react"
import { createRoot } from "react-dom/client"
import { FluentProvider, webLightTheme } from "@fluentui/react-components"
import { App } from "./App"
import "./styles.css"

Office.onReady(() => {
  const container = document.getElementById("root")
  if (!container) return

  const root = createRoot(container)
  root.render(
    <StrictMode>
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>
    </StrictMode>,
  )
})
