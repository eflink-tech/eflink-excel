// AI 模块类型定义（与 eflink-pptx / eflink-draw 的 AI 模块同构）
export interface AISettings {
  baseUrl: string
  apiKey: string
  model: string
}

export interface ConversationMessage {
  role: 'user' | 'assistant'
  content: string
  /** 关联产物（如已写入表格的结构化 JSON） */
  payload?: unknown
  time: number
}

export interface Conversation {
  id: string
  title: string
  messages: ConversationMessage[]
  createdAt: number
  updatedAt: number
}

export interface Template {
  id: string
  name: string
  builtin: boolean
  /** 提示词模板：{data} 占位符替换选区/已用区域数据 */
  content: string
}

export class AIError extends Error {
  readonly cause?: unknown

  constructor(message: string, cause?: unknown) {
    super(message)
    this.name = 'AIError'
    this.cause = cause
  }
}

/** AI 生成的表格结构（systemPrompt 约定，actionExecutor 写入工作表） */
export interface AIGenTable {
  /** 可选表格标题说明，仅用于会话展示 */
  name?: string
  /** 表头（首行写入并加粗） */
  columns: string[]
  /** 数据行；以 = 开头的字符串按公式写入 */
  rows: string[][]
}

/** 判断错误是否为用户取消 */
export function isAbortError(error: unknown): boolean {
  if (error instanceof DOMException && error.name === 'AbortError') return true
  const message = error instanceof Error ? error.message : ''
  return /aborted/i.test(message)
}
