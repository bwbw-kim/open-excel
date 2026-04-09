export interface WebSearchResult {
  title: string
  url: string
  snippet: string
}

export class WebSearchService {
  async search(query: string): Promise<WebSearchResult[]> {
    const response = await fetch("https://html.duckduckgo.com/html/", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({ q: query }).toString(),
    })

    if (!response.ok) {
      throw new Error(`웹 검색에 실패했습니다: ${response.status}`)
    }

    const html = await response.text()
    const results = extractDuckDuckGoResults(html)
    if (results.length === 0) {
      throw new Error("검색 결과를 찾지 못했습니다.")
    }

    return results.slice(0, 6)
  }
}

function extractDuckDuckGoResults(html: string): WebSearchResult[] {
  const blocks = html.split(/<div[^>]+class="(?:result|web-result)[^"]*"[^>]*>/i).slice(1)
  const results: WebSearchResult[] = []

  for (const block of blocks) {
    const titleMatch = block.match(/class="result__a"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i)
    if (!titleMatch) continue

    const snippetMatch = block.match(/class="result__snippet"[^>]*>([\s\S]*?)<\/a>|class="result__snippet"[^>]*>([\s\S]*?)<\/div>/i)
    const url = decodeHtmlEntities(titleMatch[1])
    const title = stripHtml(titleMatch[2])
    const snippet = stripHtml(snippetMatch?.[1] ?? snippetMatch?.[2] ?? "")

    results.push({ title, url, snippet })
  }

  return results
}

function stripHtml(value: string) {
  return decodeHtmlEntities(value.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim())
}

function decodeHtmlEntities(value: string) {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
}
