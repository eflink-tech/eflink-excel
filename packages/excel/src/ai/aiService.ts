// AI 服务：OpenAI 兼容接口（流式，浏览器直连，BYOK 模式与 eflink-pptx 一致）
import type OpenAI from 'openai'
import { excelT } from '../i18n'
import { AIError, isAbortError, type AISettings } from './types'

/** 多模态消息片段：文本或图片（vision 接口格式） */
export type ChatMessagePart =
  | { type: 'text'; text: string }
  | { type: 'image_url'; image_url: { url: string } };

export interface ChatMessage {
  role: 'user' | 'assistant' | 'system'
  content: string | ChatMessagePart[]
}

export interface SendMessageParams {
  settings: AISettings
  messages: ChatMessage[]
  onContentChunk?: (chunk: string, fullSoFar: string) => void
  /** 思考过程增量（千问等模型通过 delta.reasoning_content 逐步吐出） */
  onReasoningChunk?: (chunk: string, fullSoFar: string) => void
  signal?: AbortSignal
}

export interface SendMessageResult {
  content: string
  reasoning?: string
}

export async function sendMessage(params: SendMessageParams): Promise<SendMessageResult> {
  const { settings, messages, onContentChunk, onReasoningChunk, signal } = params
  if (!settings.baseUrl || !settings.apiKey || !settings.model) {
    throw new AIError(excelT('excel.ai.notConfiguredError'))
  }

  const { default: OpenAIClient } = await import('openai')
  const client = new OpenAIClient({
    baseURL: settings.baseUrl,
    apiKey: settings.apiKey,
    dangerouslyAllowBrowser: true,
  })

  try {
    const stream = await client.chat.completions.create(
      {
        model: settings.model,
        messages: messages as OpenAI.Chat.ChatCompletionMessageParam[],
        stream: true,
      },
      { signal },
    )

    let full = ''
    let reasoning = ''
    for await (const chunk of stream) {
      const delta = chunk.choices[0]?.delta as
        | OpenAI.Chat.ChatCompletionChunk.Choice.Delta & { reasoning_content?: string }
        | undefined
      if (!delta) continue;

      // 思考过程与正文各自累积并回调，渲染层按 fullSoFar 整体替换占位消息
      const reasoningPiece = delta.reasoning_content ?? ''
      if (reasoningPiece) {
        reasoning += reasoningPiece
        onReasoningChunk?.(reasoningPiece, reasoning)
      }

      const piece = delta.content ?? ''
      if (piece) {
        full += piece
        onContentChunk?.(piece, full)
      }
    }
    return { content: full, reasoning: reasoning || undefined }
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
