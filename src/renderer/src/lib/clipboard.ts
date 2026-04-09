import type { ParsedClipboardTable } from "@shared/types"

export function parseClipboardTable(rawText: string): ParsedClipboardTable | null {
  const plainText = rawText.trim()
  if (!plainText.includes("\t") || !plainText.includes("\n")) {
    return null
  }

  const rows = plainText
    .split(/\r?\n/)
    .map((line) => line.split("\t").map((cell) => cell.trim()))
    .filter((row) => row.some((cell) => cell.length > 0))

  if (rows.length === 0) {
    return null
  }

  return {
    rows,
    plainText,
  }
}
