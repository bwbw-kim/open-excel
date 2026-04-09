import { defineConfig, externalizeDepsPlugin } from "electron-vite"
import react from "@vitejs/plugin-react"

const alias = {
  "@renderer": new URL("./src/renderer", import.meta.url).pathname,
  "@shared": new URL("./src/shared", import.meta.url).pathname,
  "@main": new URL("./src/main", import.meta.url).pathname,
}

export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin()],
    resolve: {
      alias,
    },
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
    resolve: {
      alias,
    },
  },
  renderer: {
    plugins: [react()],
    resolve: {
      alias,
    },
  },
})
