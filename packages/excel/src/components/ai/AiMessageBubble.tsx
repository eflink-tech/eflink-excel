// AI 消息气泡：正文 + 思考过程（reasoning）折叠区 + 用户随消息发送的图片缩略图
// 五个办公编辑器（word / excel / mindmap / pptx / draw）共用同一结构，仅 accent 与文案 key 前缀不同。
import { useState } from 'react'
import { useI18n } from '@eflink-tech/office-kit'
import type { MessageImage } from '../../ai/types'

/** 气泡所需字段：与会话条目/持久化消息同构，不要求 time，便于面板内直接渲染 */
interface BubbleMessage {
  role: 'user' | 'assistant'
  content: string
  reasoning?: string
  images?: MessageImage[]
}

interface Props {
  message: BubbleMessage
  /** 当前应用品牌色，用于用户气泡底色 */
  accent: string
  /** 该条是否正在流式接收（用于「思考中」占位） */
  streaming?: boolean
}

export function AiMessageBubble({ message, accent, streaming = false }: Props) {
  const { t } = useI18n()
  const isUser = message.role === 'user'
  const [reasoningOpen, setReasoningOpen] = useState(true)

  const hasReasoning = !isUser && Boolean(message.reasoning)
  const hasContent = message.content.length > 0
  const images = message.images ?? []
  // 流式刚开始、思考过程与正文都还没吐字时的占位
  const isThinking = !isUser && streaming && !hasReasoning && !hasContent

  return (
    <div className={`flex flex-col ${isUser ? 'items-end' : 'items-start'}`}>
      {hasReasoning && (
        <div className="mr-2 mb-1.5 w-full max-w-[88%] rounded-md border border-slate-200 bg-slate-50">
          <button
            type="button"
            className="flex w-full items-center gap-1 rounded-md px-2 py-1.5 text-left text-[11px] text-slate-400 hover:bg-slate-100"
            onClick={() => setReasoningOpen((v) => !v)}
            data-testid="ai-reasoning-toggle"
          >
            <Chevron open={reasoningOpen} />
            {t('excel.ai.reasoning')}
          </button>
          {reasoningOpen && (
            <div className="max-h-[240px] overflow-y-auto border-t border-slate-200 px-2 py-1.5 text-[11px] leading-relaxed whitespace-pre-wrap text-slate-500">
              {message.reasoning}
            </div>
          )}
        </div>
      )}

      {images.length > 0 && (
        <div className="mb-1.5 flex max-w-[88%] flex-wrap justify-end gap-1.5">
          {images.map((img, i) => (
            <a key={i} href={img.dataUrl} target="_blank" rel="noreferrer">
              <img
                src={img.dataUrl}
                alt={t('excel.ai.imageAlt', { index: i + 1 })}
                className="max-h-[120px] max-w-[200px] rounded-md border border-slate-200"
              />
            </a>
          ))}
        </div>
      )}

      {isThinking && (
        <div className="mr-2 rounded-md bg-slate-100 px-2.5 py-2 text-xs text-slate-400 animate-pulse" role="status">
          💭 {t('excel.ai.thinking')}
        </div>
      )}

      {hasContent && (
        <div
          className={`rounded-md px-2.5 py-2 text-xs leading-relaxed whitespace-pre-wrap ${isUser ? 'ml-6 text-white' : 'mr-2 bg-slate-100 text-slate-700'}`}
          style={isUser ? { backgroundColor: accent } : undefined}
        >
          {message.content}
        </div>
      )}
    </div>
  )
}

function Chevron({ open }: { open: boolean }) {
  return (
    <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true" className={open ? '' : '-rotate-90'}>
      <path d="M2 3.5 5 6.5 8 3.5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
    </svg>
  )
}
