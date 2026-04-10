import type { KeyboardEvent, ClipboardEvent } from "react"
import { Button } from "@fluentui/react-components"
import { Send24Regular, Dismiss16Regular } from "@fluentui/react-icons"
import type { ChatAttachment } from "@/shared/types"

interface ComposerProps {
  value: string
  onChange: (value: string) => void
  onSend: () => void
  onKeyDown: (event: KeyboardEvent<HTMLTextAreaElement>) => void
  onPaste: (event: ClipboardEvent<HTMLTextAreaElement>) => void
  attachment?: ChatAttachment
  onClearAttachment: () => void
  disabled: boolean
}

export function Composer({
  value,
  onChange,
  onSend,
  onKeyDown,
  onPaste,
  attachment,
  onClearAttachment,
  disabled,
}: ComposerProps) {
  const canSend = !disabled && (value.trim().length > 0 || Boolean(attachment))

  return (
    <div className="composer">
      {attachment && (
        <div className="composer-attachment">
          <div className="attachment" style={{ position: "relative" }}>
            <Button
              appearance="subtle"
              size="small"
              icon={<Dismiss16Regular />}
              onClick={onClearAttachment}
              style={{ position: "absolute", top: 4, right: 4 }}
              title="첨부 삭제"
            />
            <div className="attachment-title">{attachment.title}</div>
            <table>
              <tbody>
                {attachment.rows.slice(0, 2).map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {row.slice(0, 4).map((cell, cellIndex) => (
                      <td key={cellIndex}>{cell}</td>
                    ))}
                    {row.length > 4 && <td>…</td>}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="composer-input-row">
        <textarea
          className="composer-textarea"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={onKeyDown}
          onPaste={onPaste}
          placeholder="무엇을 도와드릴까요?"
          disabled={disabled}
          rows={1}
        />
        <Button
          appearance="primary"
          icon={<Send24Regular />}
          onClick={onSend}
          disabled={!canSend}
          title="전송"
        />
      </div>

      <div className="composer-hint">
        {disabled ? "처리 중..." : "Enter로 전송, Shift+Enter로 줄바꿈"}
      </div>
    </div>
  )
}
