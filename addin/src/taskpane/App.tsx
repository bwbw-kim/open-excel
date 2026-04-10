import { useEffect, useRef, useState, type KeyboardEvent, type ClipboardEvent } from "react"
import { Button, Spinner } from "@fluentui/react-components"
import { Send24Regular } from "@fluentui/react-icons"
import type { ChatMessage, ChatAttachment, WorkbookSummary } from "@/shared/types"
import { Header } from "./components/Header"
import { ChatMessages } from "./components/ChatMessages"
import { Composer } from "./components/Composer"
import { ExcelService } from "@/services/excel-service"

const excelService = new ExcelService()

export function App() {
  const [messages, setMessages] = useState<ChatMessage[]>([])
  const [workbook, setWorkbook] = useState<WorkbookSummary | null>(null)
  const [composer, setComposer] = useState("")
  const [attachment, setAttachment] = useState<ChatAttachment | undefined>()
  const [isSending, setIsSending] = useState(false)
  const messagesEndRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    loadWorkbookInfo()
  }, [])

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" })
  }, [messages])

  async function loadWorkbookInfo() {
    try {
      const info = await excelService.getWorkbookInfo()
      setWorkbook(info)
    } catch {
      setWorkbook(null)
    }
  }

  async function handleSend() {
    const trimmed = composer.trim()
    if (!trimmed && !attachment) return
    if (isSending) return

    const userMessage = createMessage("user", trimmed || "이 요청을 처리해줘.", attachment)
    const pendingId = crypto.randomUUID()

    setMessages((prev) => [
      ...prev,
      userMessage,
      createMessage("assistant", "처리 중...", undefined, pendingId),
    ])
    setComposer("")
    setAttachment(undefined)
    setIsSending(true)

    try {
      const response = await processUserRequest(trimmed, attachment)
      setMessages((prev) =>
        prev.map((msg) =>
          msg.id === pendingId ? { ...msg, content: response.reply, attachment: response.attachment } : msg,
        ),
      )
      await loadWorkbookInfo()
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "오류가 발생했습니다."
      setMessages((prev) =>
        prev.map((msg) => (msg.id === pendingId ? { ...msg, content: errorMessage } : msg)),
      )
    } finally {
      setIsSending(false)
    }
  }

  async function processUserRequest(
    message: string,
    userAttachment?: ChatAttachment,
  ): Promise<{ reply: string; attachment?: ChatAttachment }> {
    const lowerMessage = message.toLowerCase()

    if (lowerMessage.includes("읽") || lowerMessage.includes("보여") || lowerMessage.includes("read")) {
      const rangeMatch = message.match(/([A-Z]+\d+:[A-Z]+\d+)/i)
      if (rangeMatch) {
        const attachment = await excelService.readRange(rangeMatch[1])
        return { reply: `${rangeMatch[1]} 범위를 읽었습니다.`, attachment }
      }
      const attachment = await excelService.readUsedRange()
      return { reply: "현재 시트의 데이터를 읽었습니다.", attachment }
    }

    if (lowerMessage.includes("시트") && (lowerMessage.includes("만들") || lowerMessage.includes("생성") || lowerMessage.includes("create"))) {
      const nameMatch = message.match(/["']([^"']+)["']/) || message.match(/시트\s+(\S+)/)
      if (nameMatch) {
        const result = await excelService.createSheet(nameMatch[1])
        await loadWorkbookInfo()
        return { reply: result }
      }
      return { reply: "시트 이름을 지정해주세요. 예: '새시트' 시트 만들어줘" }
    }

    if (lowerMessage.includes("삭제") || lowerMessage.includes("지워") || lowerMessage.includes("delete")) {
      const cellMatch = message.match(/([A-Z]+\d+)/i)
      if (cellMatch) {
        const result = await excelService.deleteCell(cellMatch[1])
        return { reply: result }
      }
      const rowMatch = message.match(/(\d+)\s*행/)
      if (rowMatch) {
        const result = await excelService.deleteRow(parseInt(rowMatch[1], 10))
        return { reply: result }
      }
    }

    if (userAttachment && userAttachment.rows.length > 0) {
      const cellMatch = message.match(/([A-Z]+\d+)/i)
      const startCell = cellMatch ? cellMatch[1] : "A1"
      const rows = userAttachment.rows
      const endCol = String.fromCharCode(64 + rows[0].length)
      const endRow = parseInt(startCell.match(/\d+/)?.[0] || "1", 10) + rows.length - 1
      const range = `${startCell}:${endCol}${endRow}`

      const result = await excelService.writeRange(range, rows)
      return { reply: result }
    }

    const cellMatch = message.match(/([A-Z]+\d+)/i)
    if (cellMatch) {
      const valueMatch = message.match(/["']([^"']+)["']/) || message.match(/에\s+(.+?)(?:\s*(?:넣|쓰|입력|작성)|\s*$)/)
      if (valueMatch) {
        const result = await excelService.writeCell(cellMatch[1], valueMatch[1].trim())
        return { reply: result }
      }
    }

    return {
      reply: "무엇을 도와드릴까요? 예시:\n• A1 셀에 '제목' 넣어줘\n• B2:D10 범위 읽어줘\n• 3행 삭제해줘\n• '새시트' 시트 만들어줘",
    }
  }

  function handleKeyDown(event: KeyboardEvent<HTMLTextAreaElement>) {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault()
      handleSend()
    }
  }

  function handlePaste(event: ClipboardEvent<HTMLTextAreaElement>) {
    const text = event.clipboardData.getData("text")
    const parsed = parseClipboardTable(text)
    if (parsed) {
      event.preventDefault()
      setAttachment({
        kind: "table",
        title: `붙여넣기 (${parsed.rows.length}x${parsed.rows[0]?.length || 0})`,
        rows: parsed.rows,
      })
    }
  }

  return (
    <div className="taskpane">
      <Header workbook={workbook} onRefresh={loadWorkbookInfo} />

      {messages.length === 0 ? (
        <div className="empty-state">
          <div className="empty-state-icon">📊</div>
          <div className="empty-state-title">Excel Copilot</div>
          <div className="empty-state-description">
            자연어로 Excel 작업을 요청하세요. 셀 읽기/쓰기, 행 관리, 시트 생성 등을 지원합니다.
          </div>
        </div>
      ) : (
        <ChatMessages messages={messages} messagesEndRef={messagesEndRef} />
      )}

      <Composer
        value={composer}
        onChange={setComposer}
        onSend={handleSend}
        onKeyDown={handleKeyDown}
        onPaste={handlePaste}
        attachment={attachment}
        onClearAttachment={() => setAttachment(undefined)}
        disabled={isSending}
      />
    </div>
  )
}

function createMessage(
  role: ChatMessage["role"],
  content: string,
  attachment?: ChatAttachment,
  id = crypto.randomUUID(),
): ChatMessage {
  return { id, role, content, createdAt: Date.now(), attachment }
}

function parseClipboardTable(text: string): { rows: string[][] } | null {
  const lines = text.trim().split(/\r?\n/)
  if (lines.length < 1) return null

  const rows = lines.map((line) => line.split("\t"))
  const hasMultipleCells = rows.some((row) => row.length > 1) || rows.length > 1

  if (!hasMultipleCells) return null
  return { rows }
}
