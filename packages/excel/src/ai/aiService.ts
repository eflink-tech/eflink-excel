// AI 服务：OpenAI 兼容接口（流式，浏览器直连，BYOK 模式与 eflink-pptx 一致）
import OpenAI from 'openai'
import { AIError, isAbortError, type AISettings } from './types'

export interface ChatMessage {
  role: 'user' | 'assistant' | 'system'
  content: string
}

export interface SendMessageParams {
  settings: AISettings
  messages: ChatMessage[]
  onContentChunk?: (chunk: string, fullSoFar: string) => void
  signal?: AbortSignal
}

export interface SendMessageResult {
  content: string
}

export async function sendMessage(params: SendMessageParams): Promise<SendMessageResult> {
  const { settings, messages, onContentChunk, signal } = params
  if (!settings.baseUrl || !settings.apiKey || !settings.model) {
    throw new AIError('请先在 AI 设置中配置接口地址、API Key 和模型')
  }

  const client = new OpenAI({
    baseURL: settings.baseUrl,
    apiKey: settings.apiKey,
    dangerouslyAllowBrowser: true,
  })

  try {
    const stream = await client.chat.completions.create(
      {
        model: settings.model,
        messages,
        stream: true,
      },
      { signal },
    )

    let full = ''
    for await (const chunk of stream) {
      const delta = chunk.choices[0]?.delta
      const piece = delta?.content ?? ''
      if (piece) {
        full += piece
        onContentChunk?.(piece, full)
      }
    }
    return { content: full }
  } catch (error) {
    if (isAbortError(error)) throw error
    throw new AIError(errorMessage(error), error)
  }
}

export function errorMessage(error: unknown): string {
  if (error instanceof AIError) return error.message
  if (error instanceof Error) {
    const anyErr = error as Error & { status?: number }
    if (anyErr.status === 401) return 'API Key 无效（401）'
    if (anyErr.status === 429) return '请求过于频繁或额度不足（429）'
    return error.message || '请求失败'
  }
  return '未知错误'
}
