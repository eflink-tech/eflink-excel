// AI 服务：OpenAI 兼容接口（流式，浏览器直连，BYOK 模式与 eflink-pptx 一致）
import OpenAI from 'openai'
import { excelT } from '../i18n'
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
    throw new AIError(excelT('excel.ai.notConfiguredError'))
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
    if (anyErr.status === 401) return excelT('excel.ai.error.unauthorized')
    if (anyErr.status === 429) return excelT('excel.ai.error.rateLimited')
    return error.message || excelT('excel.ai.error.requestFailed')
  }
  return excelT('excel.ai.error.unknown')
}
