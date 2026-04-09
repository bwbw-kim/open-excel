export interface Logger {
  enabled: boolean
  info(message: string, meta?: unknown): void
  error(message: string, meta?: unknown): void
}

export function createLogger(enabled: boolean): Logger {
  return {
    enabled,
    info(message, meta) {
      if (!enabled) return
      console.info(`[open-excel] ${message}`, meta ?? "")
    },
    error(message, meta) {
      if (!enabled) return
      console.error(`[open-excel] ${message}`, meta ?? "")
    },
  }
}
