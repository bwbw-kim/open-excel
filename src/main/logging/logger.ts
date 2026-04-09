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
      console.info(formatLogLine(message, meta))
    },
    error(message, meta) {
      if (!enabled) return
      console.error(formatLogLine(message, meta))
    },
  }
}

function formatLogLine(message: string, meta?: unknown) {
  const prefix = `[open-excel] ${message}`
  const formattedMeta = formatMeta(meta)
  return formattedMeta ? `${prefix} ${formattedMeta}` : prefix
}

function formatMeta(meta?: unknown) {
  if (meta == null || meta === "") {
    return ""
  }

  if (typeof meta === "string") {
    return meta
  }

  try {
    return JSON.stringify(meta, null, 0)
  } catch {
    return String(meta)
  }
}
