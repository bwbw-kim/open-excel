import type { AuthState } from "@/shared/types"

const CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann"
const AUTH_URL = "https://chatgpt.com/authorize"
const TOKEN_URL = "https://auth0.openai.com/oauth/token"
const REDIRECT_PORT = 1455
const REDIRECT_URI = `http://localhost:${REDIRECT_PORT}/auth/callback`
const STORAGE_KEY = "excel-copilot-auth"

export class AuthService {
  private authState: AuthState = { authenticated: false }
  private codeVerifier: string | null = null

  constructor() {
    this.loadStoredAuth()
  }

  getState(): AuthState {
    return this.authState
  }

  isAuthenticated(): boolean {
    if (!this.authState.authenticated || !this.authState.expiresAt) {
      return false
    }
    return Date.now() < this.authState.expiresAt
  }

  getAccessToken(): string | null {
    if (!this.isAuthenticated()) return null
    return this.authState.accessToken ?? null
  }

  getAccountId(): string | undefined {
    return this.authState.accountId
  }

  async startLogin(): Promise<string> {
    this.codeVerifier = generateCodeVerifier()
    const codeChallenge = await generateCodeChallenge(this.codeVerifier)
    const state = crypto.randomUUID()

    const params = new URLSearchParams({
      client_id: CLIENT_ID,
      redirect_uri: REDIRECT_URI,
      response_type: "code",
      scope: "openid profile email offline_access",
      code_challenge: codeChallenge,
      code_challenge_method: "S256",
      state,
    })

    return `${AUTH_URL}?${params.toString()}`
  }

  async handleCallback(code: string): Promise<AuthState> {
    if (!this.codeVerifier) {
      throw new Error("로그인 흐름이 시작되지 않았습니다.")
    }

    const response = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        grant_type: "authorization_code",
        code,
        redirect_uri: REDIRECT_URI,
        code_verifier: this.codeVerifier,
      }),
    })

    if (!response.ok) {
      throw new Error(`토큰 교환 실패: ${response.status}`)
    }

    const tokens = (await response.json()) as {
      access_token: string
      expires_in: number
      id_token?: string
    }

    this.authState = {
      authenticated: true,
      accessToken: tokens.access_token,
      expiresAt: Date.now() + tokens.expires_in * 1000,
    }

    this.saveAuth()
    this.codeVerifier = null

    return this.authState
  }

  logout(): void {
    this.authState = { authenticated: false }
    localStorage.removeItem(STORAGE_KEY)
  }

  private loadStoredAuth(): void {
    try {
      const stored = localStorage.getItem(STORAGE_KEY)
      if (stored) {
        const parsed = JSON.parse(stored) as AuthState
        if (parsed.expiresAt && Date.now() < parsed.expiresAt) {
          this.authState = parsed
        }
      }
    } catch {
      this.authState = { authenticated: false }
    }
  }

  private saveAuth(): void {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(this.authState))
    } catch {
      console.warn("Failed to save auth state")
    }
  }
}

function generateCodeVerifier(): string {
  const array = new Uint8Array(32)
  crypto.getRandomValues(array)
  return base64UrlEncode(array)
}

async function generateCodeChallenge(verifier: string): Promise<string> {
  const encoder = new TextEncoder()
  const data = encoder.encode(verifier)
  const hash = await crypto.subtle.digest("SHA-256", data)
  return base64UrlEncode(new Uint8Array(hash))
}

function base64UrlEncode(bytes: Uint8Array): string {
  const binary = String.fromCharCode(...bytes)
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "")
}
