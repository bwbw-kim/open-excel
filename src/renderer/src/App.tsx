import { Fragment, useEffect, useMemo, useRef, useState, type ClipboardEvent, type KeyboardEvent, type ReactNode } from "react"
import type { AuthState, ChatAttachment, ChatMessage, ChatSessionState } from "@shared/types"
import { parseClipboardTable } from "./lib/clipboard"

const EMPTY_SESSION: ChatSessionState = {
  sessionId: "loading",
  messages: [],
}

export function App() {
  const [session, setSession] = useState<ChatSessionState>(EMPTY_SESSION)
  const [authState, setAuthState] = useState<AuthState>({ authenticated: false })
  const [composer, setComposer] = useState("")
  const [attachment, setAttachment] = useState<ChatAttachment | undefined>()
  const [isDev, setIsDev] = useState(false)
  const [isSending, setIsSending] = useState(false)
  const [workingDots, setWorkingDots] = useState(1)
  const pendingAssistantIdRef = useRef<string | undefined>(undefined)
  const messagesRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    void Promise.all([window.openExcel.getChatSession(), window.openExcel.getAuthState(), window.openExcel.getDevState()]).then(
      ([loadedSession, loadedAuthState, devState]) => {
        setSession(loadedSession)
        setAuthState(loadedAuthState)
        setIsDev(devState.isDev)
      },
    )
  }, [])

  useEffect(() => {
    if (!isSending) {
      setWorkingDots(1)
      return
    }

    const timer = window.setInterval(() => {
      setWorkingDots((current) => (current % 3) + 1)
    }, 450)

    return () => window.clearInterval(timer)
  }, [isSending])

  const canSend = useMemo(() => !isSending && (composer.trim().length > 0 || Boolean(attachment)), [attachment, composer, isSending])

  async function handleSend() {
    if (!canSend) return
    const message = composer.trim() || "이 요청을 처리해줘."
    const nextAttachment = attachment
    const userMessage = createClientMessage("user", message, nextAttachment)
    const pendingAssistantId = crypto.randomUUID()
    pendingAssistantIdRef.current = pendingAssistantId

    setSession((current) => ({
      ...current,
      messages: [
        ...current.messages,
        userMessage,
        createClientMessage("assistant", formatWorkingLabel(1), undefined, pendingAssistantId),
      ],
    }))
    setComposer("")
    setAttachment(undefined)

    setIsSending(true)

    try {
      const nextSession = await window.openExcel.sendChatMessage({
        message,
        attachment: nextAttachment,
      })
      setSession(nextSession)
    } finally {
      pendingAssistantIdRef.current = undefined
      setIsSending(false)
    }
  }

  useEffect(() => {
    const pendingAssistantId = pendingAssistantIdRef.current
    if (!pendingAssistantId) return

    setSession((current) => ({
      ...current,
      messages: current.messages.map((message) =>
        message.id === pendingAssistantId ? { ...message, content: formatWorkingLabel(workingDots) } : message,
      ),
    }))
  }, [workingDots])

  useEffect(() => {
    const container = messagesRef.current
    if (!container) return

    container.scrollTo({ top: container.scrollHeight, behavior: "smooth" })
  }, [session.messages, isSending])

  async function handleOpenWorkbook() {
    try {
      await window.openExcel.connectLiveWorkbook()
      const nextSession = await window.openExcel.getChatSession()
      setSession(nextSession)
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      setSession((current) => ({
        ...current,
        messages: [...current.messages, createClientMessage("assistant", message)],
      }))
    }
  }

  async function handleLogin() {
    const nextAuthState = await window.openExcel.startLogin()
    setAuthState(nextAuthState)
  }

  function handlePaste(event: ClipboardEvent<HTMLTextAreaElement>) {
    if (isSending) {
      event.preventDefault()
      return
    }

    const text = event.clipboardData.getData("text")
    const parsed = parseClipboardTable(text)
    if (!parsed) {
      return
    }

    event.preventDefault()
    setAttachment({
      kind: "table",
      title: `Pasted range (${parsed.rows.length}x${parsed.rows[0]?.length ?? 0})`,
      rows: parsed.rows,
    })
  }

  function handleComposerKeyDown(event: KeyboardEvent<HTMLTextAreaElement>) {
    if (event.key !== "Enter" || event.shiftKey) {
      return
    }

    event.preventDefault()
    void handleSend()
  }

  return (
    <div className="shell">
      <aside className="sidebar">
        <div className="sidebar-brand-block">
          <div className="brand">Open Excel</div>
          <p className="muted sidebar-subtitle">Live spreadsheet workspace</p>
        </div>

        <div className="panel">
          <h3>Account</h3>
          <p className="muted">{authState.authenticated ? "Connected" : "Not connected"}</p>
          <button className="button secondary" onClick={handleLogin}>
            {authState.authenticated ? "Reconnect" : "Login with OpenAI"}
          </button>
        </div>

        <div className="panel">
          <h3>Workbook</h3>
          <button className="button" onClick={handleOpenWorkbook}>
            Connect Excel
          </button>
          <p className="muted">{session.activeWorkbook ? "Excel connected" : "Excel not connected"}</p>
          {session.activeWorkbook?.activeSheetName ? (
            <p className="muted">Current sheet: {session.activeWorkbook.activeSheetName}</p>
          ) : null}
        </div>

        {isDev ? (
          <div className="panel dev-panel">
            <h3>Dev mode</h3>
            <p className="muted">Verbose logs are enabled only in dev.</p>
          </div>
        ) : null}
      </aside>

      <main className="workspace workspace-single">
        <section className="chat-column">
          <header className="chat-header">
            <div>
              <h1>Excel Live conversation</h1>
              <p className="muted">현재 열려 있는 Excel workbook을 기준으로 자연어로 작업합니다.</p>
            </div>
          </header>

          <div ref={messagesRef} className="messages">
            {session.messages.map((message) => (
              <article key={message.id} className={`message ${message.role}`}>
                <div className="message-role">{message.role === "assistant" ? "Open Excel" : "You"}</div>
                <div className="message-body">{renderMessageContent(message.content)}</div>
                {message.attachment ? <TableAttachment attachment={message.attachment} /> : null}
              </article>
            ))}
          </div>

          <div className="composer">
            {attachment ? <TableAttachment attachment={attachment} compact /> : null}
            <textarea
              value={composer}
              onChange={(event) => setComposer(event.target.value)}
              onPaste={handlePaste}
              onKeyDown={handleComposerKeyDown}
              placeholder="Ask what to change, analyze, or generate…"
              disabled={isSending}
            />
            <div className="composer-actions">
              <span className="muted">
                {isSending ? formatWorkingLabel(workingDots) : "Enter로 전송하고 Shift+Enter로 줄바꿈할 수 있습니다."}
              </span>
              <button className="button" onClick={handleSend} disabled={!canSend}>
                {isSending ? "Working" : "Send"}
              </button>
            </div>
          </div>
        </section>
      </main>
    </div>
  )
}

function createClientMessage(
  role: ChatMessage["role"],
  content: string,
  attachment?: ChatAttachment,
  id = crypto.randomUUID(),
): ChatMessage {
  return {
    id,
    role,
    content,
    createdAt: Date.now(),
    attachment,
  }
}

function formatWorkingLabel(workingDots: number) {
  return `working${".".repeat(workingDots)}`
}

function renderMessageContent(content: string) {
  const lines = content.split("\n")
  return lines.map((line, lineIndex) => (
    <Fragment key={`line-${lineIndex}`}>
      {renderInlineMarkdown(line)}
      {lineIndex < lines.length - 1 ? <br /> : null}
    </Fragment>
  ))
}

function renderInlineMarkdown(line: string): ReactNode[] {
  const segments = line.split(/(\*\*[^*]+\*\*)/g)
  return segments.filter(Boolean).map((segment, index) => {
    if (segment.startsWith("**") && segment.endsWith("**")) {
      return <strong key={`segment-${index}`}>{segment.slice(2, -2)}</strong>
    }

    return <Fragment key={`segment-${index}`}>{segment}</Fragment>
  })
}

function TableAttachment({ attachment, compact = false }: { attachment: ChatAttachment; compact?: boolean }) {
  const rows = normalizeAttachmentRows(attachment.rows)

  return (
    <div className={`attachment ${compact ? "compact" : ""}`}>
      <div className="attachment-title">{attachment.title}</div>
      <table>
        <tbody>
          {rows.slice(0, compact ? 3 : 6).map((row, rowIndex) => (
            <tr key={`attachment-${rowIndex}`}>
              {row.map((cell, cellIndex) => (
                <td key={`attachment-${rowIndex}-${cellIndex}`}>{cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function normalizeAttachmentRows(rows: unknown): string[][] {
  if (!Array.isArray(rows)) {
    return [[stringifyAttachmentCell(rows)]]
  }

  if (rows.length === 0) {
    return [[""]]
  }

  if (rows.every((row) => Array.isArray(row))) {
    return rows.map((row) => row.map((cell) => stringifyAttachmentCell(cell)))
  }

  return [rows.map((cell) => stringifyAttachmentCell(cell))]
}

function stringifyAttachmentCell(value: unknown) {
  if (value == null) return ""
  return String(value)
}
