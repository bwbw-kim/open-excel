import { createServer } from "node:http"
import { existsSync, readFileSync, statSync, createReadStream } from "node:fs"
import { mkdirSync, writeFileSync, readFileSync as readTextFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"

const SERVER_PORT = 3000
const OAUTH_PORT = 1455
const CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann"
const ISSUER = "https://auth.openai.com"
const AUTH_FILE = path.join(os.homedir(), ".open-excel", "chatgpt-auth.json")
const ROOT = path.resolve("out/renderer")

let pendingLogin = null

createServer(async (req, res) => {
  setCors(res)
  if (req.method === "OPTIONS") {
    res.writeHead(204)
    res.end()
    return
  }

  const url = new URL(req.url ?? "/", `http://localhost:${SERVER_PORT}`)

  if (url.pathname === "/api/health") {
    json(res, 200, { ok: true })
    return
  }

  if (url.pathname === "/api/auth/state") {
    const auth = await loadAuth()
    json(res, 200, auth ? { authenticated: auth.expiresAt > Date.now(), accountId: auth.accountId, expiresAt: auth.expiresAt } : { authenticated: false })
    return
  }

  if (url.pathname === "/api/auth/token") {
    const auth = await ensureAuth()
    json(res, 200, { accessToken: auth.accessToken, accountId: auth.accountId })
    return
  }

  if (url.pathname === "/api/auth/start" && req.method === "POST") {
    const result = await startLogin()
    json(res, 200, result)
    return
  }

  serveStatic(req, res, url.pathname)
}).listen(SERVER_PORT, () => {
  console.log(`open-excel add-in server running at http://localhost:${SERVER_PORT}`)
})

async function startLogin() {
  if (pendingLogin) return { authUrl: pendingLogin.authUrl }
  const { verifier, challenge } = await generatePkce()
  const state = generateState()
  const redirectUri = `http://localhost:${OAUTH_PORT}/auth/callback`
  const authUrl = buildAuthorizeUrl({ redirectUri, challenge, state })
  pendingLogin = { authUrl, verifier }
  void waitForAuthorizationCode(state, verifier).finally(() => {
    pendingLogin = null
  })
  return { authUrl }
}

async function ensureAuth() {
  const saved = await loadAuth()
  if (!saved) throw new Error("OpenAI 로그인이 필요합니다. 먼저 Login with OpenAI를 눌러 주세요.")
  if (saved.expiresAt > Date.now() + 30000) return saved
  return refreshAuth(saved.refreshToken, saved.accountId)
}

async function loadAuth() {
  try {
    return JSON.parse(readTextFileSync(AUTH_FILE, "utf8"))
  } catch {
    return undefined
  }
}

async function saveAuth(auth) {
  mkdirSync(path.dirname(AUTH_FILE), { recursive: true })
  writeFileSync(AUTH_FILE, JSON.stringify(auth, null, 2), { mode: 0o600 })
}

async function refreshAuth(refreshToken, accountId) {
  const response = await fetch(`${ISSUER}/oauth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "refresh_token", refresh_token: refreshToken, client_id: CLIENT_ID }).toString(),
  })
  if (!response.ok) throw new Error(`Failed token refresh: ${response.status}`)
  return saveAndReturnAuth(await response.json(), accountId)
}

async function saveAndReturnAuth(tokens, fallbackAccountId) {
  const auth = {
    accessToken: tokens.access_token,
    refreshToken: tokens.refresh_token,
    expiresAt: Date.now() + (tokens.expires_in ?? 3600) * 1000,
    accountId: extractAccountId(tokens) ?? fallbackAccountId,
  }
  await saveAuth(auth)
  return auth
}

function buildAuthorizeUrl(input) {
  const params = new URLSearchParams({
    response_type: "code",
    client_id: CLIENT_ID,
    redirect_uri: input.redirectUri,
    scope: "openid profile email offline_access",
    code_challenge: input.challenge,
    code_challenge_method: "S256",
    id_token_add_organizations: "true",
    codex_cli_simplified_flow: "true",
    state: input.state,
    originator: "open-excel",
  })
  return `${ISSUER}/oauth/authorize?${params.toString()}`
}

function waitForAuthorizationCode(expectedState, verifier) {
  return new Promise((resolve, reject) => {
    const server = createServer((req, res) => {
      const url = new URL(req.url ?? "/", `http://localhost:${OAUTH_PORT}`)
      if (url.pathname !== "/auth/callback") {
        res.writeHead(404)
        res.end("Not found")
        return
      }
      const error = url.searchParams.get("error")
      const code = url.searchParams.get("code")
      const state = url.searchParams.get("state")
      if (error) {
        server.close()
        res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" })
        res.end(`Authorization failed: ${error}`)
        reject(new Error(`OAuth failed: ${error}`))
        return
      }
      if (!code || state !== expectedState) {
        server.close()
        res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" })
        res.end("Invalid authorization callback")
        reject(new Error("Invalid OAuth callback."))
        return
      }
      server.close()
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" })
      res.end("<html><body><h1>Login complete</h1><p>You can return to Open Excel.</p></body></html>")
      exchangeCodeForTokens(code, `http://localhost:${OAUTH_PORT}/auth/callback`, verifier).then(resolve, reject)
    })
    server.listen(OAUTH_PORT)
    server.on("error", reject)
  })
}

async function exchangeCodeForTokens(code, redirectUri, verifier) {
  const response = await fetch(`${ISSUER}/oauth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "authorization_code", code, redirect_uri: redirectUri, client_id: CLIENT_ID, code_verifier: verifier }).toString(),
  })
  if (!response.ok) throw new Error(`Failed token exchange: ${response.status}`)
  return saveAndReturnAuth(await response.json())
}

function generatePkce() {
  const verifier = generateRandomString(43)
  return crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)).then((hash) => ({ verifier, challenge: base64UrlEncode(hash) }))
}

function generateRandomString(length) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
  const bytes = crypto.getRandomValues(new Uint8Array(length))
  return Array.from(bytes).map((value) => chars[value % chars.length]).join("")
}

function generateState() {
  return base64UrlEncode(crypto.getRandomValues(new Uint8Array(32)).buffer)
}

function base64UrlEncode(buffer) {
  return Buffer.from(buffer).toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "")
}

function extractAccountId(tokens) {
  return extractAccountIdFromJwt(tokens.id_token) ?? extractAccountIdFromJwt(tokens.access_token)
}

function extractAccountIdFromJwt(token) {
  const parts = token.split(".")
  if (parts.length !== 3) return undefined
  try {
    const claims = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"))
    const nested = claims["https://api.openai.com/auth"]
    if (typeof nested === "object" && nested && "chatgpt_account_id" in nested && typeof nested.chatgpt_account_id === "string") return nested.chatgpt_account_id
    return claims.chatgpt_account_id ?? claims.organizations?.[0]?.id
  } catch {
    return undefined
  }
}

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*")
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization")
}

function json(res, status, body) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" })
  res.end(JSON.stringify(body))
}

function serveStatic(req, res, pathname) {
  const filePath = pathname === "/" ? path.join(ROOT, "index.html") : path.join(ROOT, pathname)
  if (!filePath.startsWith(ROOT) || !existsSync(filePath)) {
    const fallback = path.join(ROOT, "index.html")
    if (existsSync(fallback)) {
      streamFile(fallback, res)
      return
    }
    res.writeHead(404)
    res.end("Build output not found. Run npm run build first.")
    return
  }
  streamFile(filePath, res)
}

function streamFile(filePath, res) {
  const ext = path.extname(filePath).toLowerCase()
  const contentType = ext === ".html" ? "text/html; charset=utf-8" : ext === ".js" ? "text/javascript; charset=utf-8" : ext === ".css" ? "text/css; charset=utf-8" : "application/octet-stream"
  res.writeHead(200, { "Content-Type": contentType })
  createReadStream(filePath).pipe(res)
}
