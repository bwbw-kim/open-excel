declare const Office: {
  HostType: { Excel: string }
  onReady: (callback: (info: { host: string }) => void) => Promise<{ host: string }>
}

declare const Excel: any
