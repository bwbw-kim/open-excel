import type { OpenWorkbookResult } from "@shared/types"

export class NumbersAdapter {
  supports(format: "xlsx" | "csv" | "numbers") {
    return format === "numbers"
  }

  async openWorkbook(_filePath: string): Promise<OpenWorkbookResult> {
    throw new Error("Numbers support is not included in the MVP build yet.")
  }
}
