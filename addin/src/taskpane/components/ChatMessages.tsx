import type { RefObject } from "react"
import type { ChatMessage, ChatAttachment } from "@/shared/types"

interface ChatMessagesProps {
  messages: ChatMessage[]
  messagesEndRef: RefObject<HTMLDivElement>
}

export function ChatMessages({ messages, messagesEndRef }: ChatMessagesProps) {
  return (
    <div className="messages">
      {messages.map((message) => (
        <article key={message.id} className={`message ${message.role}`}>
          <div className="message-role">{message.role === "assistant" ? "Copilot" : "나"}</div>
          <div className="message-content">{renderContent(message.content)}</div>
          {message.attachment && <TableAttachment attachment={message.attachment} />}
        </article>
      ))}
      <div ref={messagesEndRef} />
    </div>
  )
}

function renderContent(content: string) {
  return content.split("\n").map((line, index) => (
    <span key={index}>
      {line}
      {index < content.split("\n").length - 1 && <br />}
    </span>
  ))
}

function TableAttachment({ attachment }: { attachment: ChatAttachment }) {
  const displayRows = attachment.rows.slice(0, 5)
  const hasMore = attachment.rows.length > 5

  return (
    <div className="attachment">
      <div className="attachment-title">{attachment.title}</div>
      <table>
        <tbody>
          {displayRows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.slice(0, 6).map((cell, cellIndex) => (
                <td key={cellIndex}>{cell}</td>
              ))}
              {row.length > 6 && <td>…</td>}
            </tr>
          ))}
        </tbody>
      </table>
      {hasMore && (
        <div className="attachment-title" style={{ marginTop: 4 }}>
          +{attachment.rows.length - 5}행 더 있음
        </div>
      )}
    </div>
  )
}
