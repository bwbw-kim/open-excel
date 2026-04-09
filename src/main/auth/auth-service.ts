import { createServer } from "node:http"
import { mkdir, readFile, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { shell } from "electron"
import type { AuthState } from "@shared/types"
import type { Logger } from "../logging/logger"

const CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann"
const ISSUER = "https://auth.openai.com"
const OAUTH_PORT = 1455
const AUTH_FILE = path.join(os.homedir(), ".open-excel", "chatgpt-auth.json")

interface StoredAuth {
  accessToken: string
  refreshToken: string
  expiresAt: number
  accountId?: string
}

interface TokenResponse {
  id_token: string
  access_token: string
  refresh_token: string
  expires_in?: number
}

export class AuthService {
  constructor(private readonly logger: Logger) {}

  async ensureAuth() {
    const saved = await this.loadAuth()
    if (!saved) {
      throw new Error("OpenAI 로그인이 필요합니다. 먼저 Login with OpenAI를 눌러 주세요.")
    }

    if (saved.expiresAt > Date.now() + 30_000) {
      return saved
    }

    const refreshed = await refreshAuth(saved.refreshToken, saved.accountId)
    await this.saveAuth(refreshed)
    return refreshed
  }

  async getState(): Promise<AuthState> {
    const auth = await this.loadAuth()
    if (!auth) {
      return { authenticated: false }
    }

    return {
      authenticated: auth.expiresAt > Date.now(),
      accountId: auth.accountId,
      expiresAt: auth.expiresAt,
    }
  }

  async startLogin(): Promise<AuthState> {
    const { verifier, challenge } = await generatePkce()
    const state = generateState()
    const redirectUri = `http://localhost:${OAUTH_PORT}/auth/callback`
    const authUrl = buildAuthorizeUrl({ redirectUri, challenge, state })
    const codePromise = waitForAuthorizationCode(state)

    this.logger.info("opening oauth login", { authUrl })
    await shell.openExternal(authUrl)

    const code = await codePromise
    const tokens = await exchangeCodeForTokens(code, redirectUri, verifier)
    const auth = toStoredAuth(tokens)
    await this.saveAuth(auth)

    return {
      authenticated: true,
      accountId: auth.accountId,
      expiresAt: auth.expiresAt,
    }
  }

  private async loadAuth(): Promise<StoredAuth | undefined> {
    try {
      const raw = await readFile(AUTH_FILE, "utf8")
      return JSON.parse(raw) as StoredAuth
    } catch {
      return undefined
    }
  }

  private async saveAuth(auth: StoredAuth) {
    await mkdir(path.dirname(AUTH_FILE), { recursive: true })
    await writeFile(AUTH_FILE, JSON.stringify(auth, null, 2), { mode: 0o600 })
  }
}

async function exchangeCodeForTokens(code: string, redirectUri: string, verifier: string): Promise<TokenResponse> {
  const response = await fetch(`${ISSUER}/oauth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      code,
      redirect_uri: redirectUri,
      client_id: CLIENT_ID,
      code_verifier: verifier,
    }).toString(),
  })

  if (!response.ok) {
    throw new Error(`Failed token exchange: ${response.status}`)
  }

  return (await response.json()) as TokenResponse
}

async function refreshAuth(refreshToken: string, accountId?: string): Promise<StoredAuth> {
  const response = await fetch(`${ISSUER}/oauth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      client_id: CLIENT_ID,
    }).toString(),
  })

  if (!response.ok) {
    throw new Error(`Failed token refresh: ${response.status}`)
  }

  const tokens = (await response.json()) as TokenResponse
  return toStoredAuth(tokens, accountId)
}

function toStoredAuth(tokens: TokenResponse, fallbackAccountId?: string): StoredAuth {
  return {
    accessToken: tokens.access_token,
    refreshToken: tokens.refresh_token,
    expiresAt: Date.now() + (tokens.expires_in ?? 3600) * 1000,
    accountId: extractAccountId(tokens) ?? fallbackAccountId,
  }
}

function buildAuthorizeUrl(input: { redirectUri: string; challenge: string; state: string }) {
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

function waitForAuthorizationCode(expectedState: string): Promise<string> {
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
      resolve(code)
    })

    server.listen(OAUTH_PORT)
    server.on("error", reject)
  })
}

async function generatePkce() {
  const verifier = generateRandomString(43)
  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))

  return {
    verifier,
    challenge: base64UrlEncode(hash),
  }
}

function generateRandomString(length: number) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
  const bytes = crypto.getRandomValues(new Uint8Array(length))
  return Array.from(bytes)
    .map((value) => chars[value % chars.length])
    .join("")
}

function generateState() {
  return base64UrlEncode(crypto.getRandomValues(new Uint8Array(32)).buffer)
}

function base64UrlEncode(buffer: ArrayBuffer) {
  return Buffer.from(buffer)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "")
}

function extractAccountId(tokens: TokenResponse) {
  return extractAccountIdFromJwt(tokens.id_token) ?? extractAccountIdFromJwt(tokens.access_token)
}

function extractAccountIdFromJwt(token: string) {
  const parts = token.split(".")
  if (parts.length !== 3) return undefined

  try {
    const claims = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8")) as {
      chatgpt_account_id?: string
      organizations?: Array<{ id: string }>
      [key: string]: unknown
    }

    const nested = claims["https://api.openai.com/auth"]
    if (typeof nested === "object" && nested && "chatgpt_account_id" in nested) {
      const value = nested.chatgpt_account_id
      if (typeof value === "string") return value
    }

    return claims.chatgpt_account_id ?? claims.organizations?.[0]?.id
  } catch {
    return undefined
  }
}
