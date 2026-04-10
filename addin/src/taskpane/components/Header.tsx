import { Button } from "@fluentui/react-components"
import { ArrowSync24Regular } from "@fluentui/react-icons"
import type { WorkbookSummary } from "@/shared/types"

interface HeaderProps {
  workbook: WorkbookSummary | null
  onRefresh: () => void
}

export function Header({ workbook, onRefresh }: HeaderProps) {
  return (
    <header className="header">
      <div className="header-title">Excel Copilot</div>
      <div className="header-status">
        <span className={`header-status-dot ${workbook ? "connected" : ""}`} />
        {workbook ? (
          <span>
            {workbook.name} • {workbook.activeSheetName}
          </span>
        ) : (
          <span>Excel에 연결되지 않음</span>
        )}
        <Button
          appearance="subtle"
          size="small"
          icon={<ArrowSync24Regular />}
          onClick={onRefresh}
          title="새로고침"
        />
      </div>
    </header>
  )
}
