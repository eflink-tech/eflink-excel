// AI 助手面板：设置 / 对话（携带表格上下文）/ 生成表格 / 选区快捷操作 / 会话历史
// 交互模式与 eflink-pptx / eflink-word 的 AIChatPanel 保持一致
import { useEffect, useRef, useState } from 'react'
import { ChevronLeft, History, Paperclip, Settings, Sparkles, Table2, Trash2, Wand2, X } from 'lucide-react'
import { useI18n } from '@eflink-tech/office-kit'
import { useUiStore } from '../../store/uiStore'
import { aiDb, aiSettingsStore, BUILTIN_TEMPLATES, listConversations } from '../../ai/db'
import { errorMessage, sendMessage, type ChatMessage, type ChatMessagePart } from '../../ai/aiService'
import { compressImageFile } from '../../ai/imageCompressor'
import { buildTablePrompt, parseAITable } from '../../ai/systemPrompt'
import { describeContext, getSheetContext, writeAITable } from '../../ai/actionExecutor'
import type { AIGenTable, AISettings, Conversation, MessageImage } from '../../ai/types'
import { isAbortError } from '../../ai/types'
import { AiMessageBubble } from './AiMessageBubble'

type PanelView = 'chat' | 'settings' | 'history'

interface SimpleMessage {
  role: 'user' | 'assistant'
  content: string
  reasoning?: string
  images?: MessageImage[]
}

const ACCENT = '#107c41'

/** 本地会话条目 → OpenAI 协议消息：带图时改用 vision 数组结构 */
function toApiMessages(entries: SimpleMessage[]): ChatMessage[] {
  return entries.map((m): ChatMessage => {
    if (!m.images?.length) return { role: m.role, content: m.content }
    const parts: ChatMessagePart[] = [
      { type: 'text', text: m.content },
      ...m.images.map((img) => ({ type: 'image_url' as const, image_url: { url: img.dataUrl } })),
    ]
    return { role: m.role, content: parts }
  })
}

/** 内置提示词模板 id → 文案 key（模板内容里的 {data} 由当前表格数据替换） */
const TEMPLATE_KEYS: Record<string, { name: string; content: string }> = {
  'tpl-analyze': { name: 'excel.ai.tpl.analyze.name', content: 'excel.ai.tpl.analyze.content' },
  'tpl-formula': { name: 'excel.ai.tpl.formula.name', content: 'excel.ai.tpl.formula.content' },
  'tpl-clean': { name: 'excel.ai.tpl.clean.name', content: 'excel.ai.tpl.clean.content' },
}

export function AIChatPanel() {
  const { t, locale } = useI18n()
  const tRef = useRef(t)
  tRef.current = t
  const toggleAIPanel = useUiStore((s) => s.toggleAIPanel)
  const showToast = useUiStore((s) => s.showToast)
  const [view, setView] = useState<PanelView>('chat')
  const [settings, setSettings] = useState<AISettings>({ baseUrl: '', apiKey: '', model: '' })
  const [configured, setConfigured] = useState(false)
  const [messages, setMessages] = useState<SimpleMessage[]>([])
  const [input, setInput] = useState('')
  const [pendingImages, setPendingImages] = useState<MessageImage[]>([])
  const [busy, setBusy] = useState(false)
  const [genView, setGenView] = useState(false)
  const [topic, setTopic] = useState('')
  const [rowCount, setRowCount] = useState(20)
  const [requirements, setRequirements] = useState('')
  const [pendingTable, setPendingTable] = useState<AIGenTable | null>(null)
  const [convId, setConvId] = useState<string | null>(null)
  const [historyList, setHistoryList] = useState<Conversation[]>([])
  const abortRef = useRef<AbortController | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const listRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    void aiSettingsStore.getSettings().then((s) => {
      setSettings(s)
      setConfigured(Boolean(s.baseUrl && s.apiKey && s.model))
    })
  }, [])

  useEffect(() => {
    // 用 scrollTop 赋值而非 scrollTo：面板常驻挂载后，jsdom 下 Element.scrollTo 不存在
    const el = listRef.current
    if (el) el.scrollTop = el.scrollHeight
  }, [messages, busy])

  // 会话持久化：消息变化时写入 Dexie
  useEffect(() => {
    if (!messages.length) return
    const id = convId ?? (() => { const nid = `conv-${Date.now()}`; setConvId(nid); return nid })()
    const now = Date.now()
    void aiDb.conversations.put({
      id,
      title: messages[0]?.content.replace(/\n/g, ' ').slice(0, 24) || tRef.current('excel.ai.newChat'),
      messages: messages.map((m) => ({ ...m, time: now })),
      createdAt: now,
      updatedAt: now,
    })
  }, [messages, convId])

  const loadHistory = () => {
    void listConversations().then(setHistoryList)
  }

  const openConversation = (conv: Conversation) => {
    setMessages(conv.messages.map((m) => ({ role: m.role, content: m.content, reasoning: m.reasoning, images: m.images })))
    setConvId(conv.id)
    setView('chat')
  }

  const deleteConversation = (id: string) => {
    void aiDb.conversations.delete(id).then(loadHistory)
  }

  const saveSettings = () => {
    void aiSettingsStore.saveSettings(settings).then(() => {
      setConfigured(Boolean(settings.baseUrl && settings.apiKey && settings.model))
      setView('chat')
    })
  }

  const buildSystemPrompt = (): string => {
    return [t('excel.ai.systemPrompt'), describeContext(getSheetContext())].join('\n')
  }

  /**
   * 发送一轮对话。images 为空时走纯文本，带图时拼成 vision 数组结构；
   * 流式期间用 content/reasoning 两个缓冲区整体替换尾部 assistant 气泡，
   * 结束后再以接口返回值覆盖一次，保证最终态与流式中间态一致。
   */
  const doSend = async (userText: string, onText?: (text: string) => string, images?: MessageImage[]) => {
    if (busy) return
    const s = await aiSettingsStore.getSettings()
    if (!s.baseUrl || !s.apiKey || !s.model) {
      setView('settings')
      return
    }
    const text = onText ? onText(userText) : userText
    const next: SimpleMessage[] = [...messages, { role: 'user', content: text, images: images?.length ? images : undefined }]
    setMessages(next)
    setInput('')
    setPendingImages([])
    setBusy(true)
    const controller = new AbortController()
    abortRef.current = controller

    const chat: ChatMessage[] = [
      { role: 'system', content: buildSystemPrompt() },
      ...toApiMessages(next),
    ]

    let streamedContent = ''
    let streamedReasoning = ''
    const flush = () => {
      setMessages([...next, {
        role: 'assistant',
        content: streamedContent,
        reasoning: streamedReasoning || undefined,
      }])
    }

    try {
      const result = await sendMessage({
        settings: s,
        messages: chat,
        signal: controller.signal,
        onContentChunk: (_chunk, full) => {
          streamedContent = full
          flush()
        },
        onReasoningChunk: (_chunk, full) => {
          streamedReasoning = full
          flush()
        },
      })
      setMessages([...next, {
        role: 'assistant',
        content: result.content,
        reasoning: result.reasoning,
      }])
    } catch (error) {
      // 用户主动停止：保留已生成的部分内容，不追加报错消息
      if (isAbortError(error)) return
      // 带图请求失败最常见的原因是模型不支持 vision，给出可操作提示
      const visionHint = images?.length ? ` ${t('excel.ai.visionHint')}` : ''
      setMessages([...next, { role: 'assistant', content: `⚠️ ${errorMessage(error)}${visionHint}` }])
    } finally {
      setBusy(false)
      abortRef.current = null
    }
  }

  /** 批量压缩选择的/粘贴的图片，单个失败逐条提示，不阻断其余 */
  const addImageFiles = async (files: File[]) => {
    for (const file of files) {
      try {
        const dataUrl = await compressImageFile(file)
        setPendingImages((prev) => [...prev, { dataUrl }])
      } catch (error) {
        showToast(t('excel.ai.imageFailed', { reason: errorMessage(error) }))
      }
    }
  }

  /** 粘贴截图走压缩管道；纯文本粘贴不拦截 */
  const handlePaste = (e: React.ClipboardEvent) => {
    const files = Array.from(e.clipboardData?.items ?? [])
      .filter((it) => it.kind === 'file' && it.type.startsWith('image/'))
      .map((it) => it.getAsFile())
      .filter((f): f is File => f !== null)
    if (!files.length) return
    e.preventDefault()
    void addImageFiles(files)
  }

  /** 生成表格（结构化 JSON），结果先展示摘要，由用户确认后写入工作表 */
  const generateTable = () => {
    if (busy) return
    if (!topic.trim()) return
    void (async () => {
      const s = await aiSettingsStore.getSettings()
      if (!s.baseUrl || !s.apiKey || !s.model) {
        setView('settings')
        return
      }
      setBusy(true)
      setPendingTable(null)
      const label = t('excel.ai.generatedLabel', { topic: topic.trim(), rows: rowCount })
      const next: SimpleMessage[] = [...messages, { role: 'user', content: label }]
      setMessages(next)
      try {
        const { content } = await sendMessage({
          settings: s,
          messages: [
            { role: 'system', content: t('excel.ai.tableSystemPrompt') },
            { role: 'user', content: buildTablePrompt(topic.trim(), rowCount, requirements.trim()) },
          ],
          onContentChunk: (_c, full) => {
            setMessages([
              ...next,
              { role: 'assistant', content: t('excel.ai.generatingChars', { count: full.length }) },
            ])
          },
        })
        const table = parseAITable(content)
        setPendingTable(table)
        setMessages([
          ...next,
          {
            role: 'assistant',
            content: t('excel.ai.generatedTable', {
              name: table.name ?? topic.trim(),
              cols: table.columns.length,
              rows: table.rows.length,
              columns: table.columns.join(' | '),
            }),
          },
        ])
      } catch (error) {
        setMessages([...next, { role: 'assistant', content: `⚠️ ${errorMessage(error)}` }])
      } finally {
        setBusy(false)
      }
    })()
  }

  const writePendingTable = () => {
    if (!pendingTable) return
    try {
      const count = writeAITable(pendingTable)
      setPendingTable(null)
      setGenView(false)
      showToast(t('excel.ai.writeSuccess', { count }))
      setMessages((prev) => [...prev, { role: 'assistant', content: t('excel.ai.writeSuccessMessage', { count }) }])
    } catch (error) {
      showToast(error instanceof Error ? error.message : t('excel.ai.writeFailed'))
      setMessages((prev) => [...prev, { role: 'assistant', content: `⚠️ ${error instanceof Error ? error.message : t('excel.ai.writeFailed')}` }])
    }
  }

  /** 选区/已用区域快捷操作：把表格数据发给 AI 分析 */
  const runTemplate = (templateId: string) => {
    const keys = TEMPLATE_KEYS[templateId]
    if (!keys) return
    const ctx = getSheetContext()
    const data = ctx.selectionData ?? ctx.usedData
    if (!data) {
      setMessages((prev) => [...prev, { role: 'assistant', content: `⚠️ ${t('excel.ai.emptyTable')}` }])
      return
    }
    void doSend(data, (d) => t(keys.content, { data: d }))
  }

  return (
    <aside
      className="flex h-full w-[320px] shrink-0 flex-col border-l border-slate-200 bg-white"
      data-testid="ai-panel"
      // Esc 只收起面板：绑定在面板根节点，焦点在面板内才生效，不会抢编辑器的 Esc 快捷键
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation()
          toggleAIPanel()
        }
      }}
    >
      {/* 头部 */}
      <div className="flex h-10 shrink-0 items-center justify-between border-b border-slate-100 px-3">
        <span className="flex items-center gap-1 text-sm font-medium text-slate-800">
          <Sparkles size={15} style={{ color: ACCENT }} />
          {t('excel.ai.title')}
        </span>
        <div className="flex items-center gap-0.5">
          <button
            type="button"
            className="rounded p-1 text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-600"
            title={t('excel.ai.history')}
            onClick={() => { loadHistory(); setView(view === 'history' ? 'chat' : 'history') }}
          >
            <History size={15} />
          </button>
          <button
            type="button"
            className="rounded p-1 text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-600"
            title={t('excel.ai.settings')}
            onClick={() => setView(view === 'settings' ? 'chat' : 'settings')}
          >
            <Settings size={15} />
          </button>
          <button
            type="button"
            className="rounded p-1 text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-600"
            title={t('excel.ai.newChat')}
            onClick={() => { setMessages([]); setPendingTable(null); setConvId(null); setView('chat') }}
          >
            <Trash2 size={15} />
          </button>
          <button
            type="button"
            className="rounded p-1 text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-600"
            title={t('excel.ai.close')}
            aria-label={t('excel.ai.close')}
            onClick={toggleAIPanel}
          >
            <X size={15} />
          </button>
        </div>
      </div>

      {view === 'history' ? (
        <div className="flex-1 overflow-y-auto p-3" data-testid="ai-history">
          <button
            type="button"
            className="mb-2 flex items-center gap-1 text-xs text-slate-500 hover:text-slate-700"
            onClick={() => setView('chat')}
          >
            <ChevronLeft size={14} />
            {t('excel.ai.backToChat')}
          </button>
          {historyList.length === 0 && <div className="py-6 text-center text-xs text-slate-400">{t('excel.ai.noHistory')}</div>}
          <div className="space-y-1.5">
            {historyList.map((conv) => (
              <div key={conv.id} className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2">
                <button
                  type="button"
                  className="flex-1 truncate text-left text-xs text-slate-700"
                  onClick={() => openConversation(conv)}
                >
                  {conv.title}
                </button>
                <span className="shrink-0 text-[10px] text-slate-400">
                  {new Date(conv.updatedAt).toLocaleString(locale, { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })}
                </span>
                <button
                  type="button"
                  className="shrink-0 rounded p-0.5 text-slate-400 hover:bg-red-50 hover:text-red-500"
                  title={t('excel.ai.delete')}
                  onClick={() => deleteConversation(conv.id)}
                >
                  🗑
                </button>
              </div>
            ))}
          </div>
        </div>
      ) : view === 'settings' ? (
        <div className="flex-1 overflow-y-auto p-3">
          <div className="mb-3 text-xs leading-relaxed text-slate-500">
            {t('excel.ai.settingsIntro')}
          </div>
          {([
            ['baseUrl', t('excel.ai.endpoint'), 'https://dashscope.aliyuncs.com/compatible-mode/v1'],
            ['apiKey', 'API Key', 'sk-…'],
            ['model', t('excel.ai.model'), 'qwen-plus'],
          ] as const).map(([key, label, placeholder]) => (
            <div key={key} className="mb-2">
              <div className="mb-1 text-xs text-slate-500">{label}</div>
              <input
                type={key === 'apiKey' ? 'password' : 'text'}
                name={`doculuma-ai-${key}`}
                autoComplete={key === 'apiKey' ? 'new-password' : 'off'}
                // 只读直到聚焦：避免 Chrome 把「接口地址 + 密码框」识别为登录表单而自动填充保存的账号密码
                readOnly
                onFocus={(e) => { e.currentTarget.readOnly = false }}
                value={settings[key]}
                placeholder={placeholder}
                onChange={(e) => setSettings({ ...settings, [key]: e.target.value })}
                className="w-full rounded border border-slate-200 px-2 py-1.5 text-xs outline-none focus:border-[#107c41]"
              />
            </div>
          ))}
          <button
            type="button"
            className="mt-2 w-full rounded-md py-1.5 text-xs text-white"
            style={{ backgroundColor: ACCENT }}
            onClick={saveSettings}
          >
            {t('excel.ai.saveSettings')}
          </button>
        </div>
      ) : (
        <>
          {/* 生成表格流程 */}
          <div className="shrink-0 border-b border-slate-100 p-2">
            {!genView ? (
              <button
                type="button"
                className="flex w-full items-center justify-center gap-1 rounded-md py-1.5 text-xs"
                style={{ backgroundColor: '#e7f4ec', color: ACCENT }}
                onClick={() => setGenView(true)}
                data-testid="ai-gen-toggle"
              >
                <Wand2 size={14} />
                {t('excel.ai.generateTableButton')}
              </button>
            ) : (
              <div className="rounded-md bg-slate-50 p-2">
                <div className="mb-1.5 flex items-center justify-between text-xs text-slate-500">
                  <span>{t('excel.ai.generateTable')}</span>
                  <button type="button" className="text-slate-400 hover:text-slate-600" onClick={() => setGenView(false)}>{t('excel.ai.collapse')}</button>
                </div>
                <input
                  value={topic}
                  onChange={(e) => setTopic(e.target.value)}
                  placeholder={t('excel.ai.topicPlaceholder')}
                  className="mb-1.5 w-full rounded border border-slate-200 px-2 py-1.5 text-xs outline-none focus:border-[#107c41]"
                  data-testid="ai-gen-topic"
                />
                <div className="mb-1.5 flex items-center gap-2 text-xs text-slate-500">
                  {t('excel.ai.rows')}
                  <input
                    type="number"
                    min={3}
                    max={200}
                    value={rowCount}
                    onChange={(e) => setRowCount(Math.min(200, Math.max(3, Number(e.target.value) || 20)))}
                    className="w-16 rounded border border-slate-200 px-1.5 py-1 text-xs"
                  />
                </div>
                <textarea
                  value={requirements}
                  onChange={(e) => setRequirements(e.target.value)}
                  placeholder={t('excel.ai.requirementsPlaceholder')}
                  rows={2}
                  className="mb-1.5 w-full resize-none rounded border border-slate-200 px-2 py-1.5 text-xs outline-none focus:border-[#107c41]"
                />
                <button
                  type="button"
                  className="w-full rounded-md py-1.5 text-xs text-white disabled:opacity-50"
                  style={{ backgroundColor: ACCENT }}
                  disabled={busy || !topic.trim()}
                  onClick={generateTable}
                  data-testid="ai-gen-table"
                >
                  {busy ? t('excel.ai.generating') : t('excel.ai.generate')}
                </button>
                {pendingTable && (
                  <button
                    type="button"
                    className="mt-1.5 flex w-full items-center justify-center gap-1 rounded-md border py-1.5 text-xs"
                    style={{ borderColor: ACCENT, color: ACCENT }}
                    onClick={writePendingTable}
                    data-testid="ai-insert-table"
                  >
                    <Table2 size={13} />
                    {t('excel.ai.writeToCell')}
                  </button>
                )}
              </div>
            )}
          </div>

          {/* 消息列表 */}
          <div ref={listRef} className="flex-1 space-y-2 overflow-y-auto p-3" data-testid="ai-messages">
            {!configured && (
              <div className="rounded-md bg-[#e7f4ec] p-2 text-xs text-slate-700">
                {t('excel.ai.notConfigured')}
              </div>
            )}
            {configured && messages.length === 0 && (
              <div className="rounded-md bg-slate-50 p-2.5 text-xs leading-relaxed text-slate-500">
                {t('excel.ai.intro')}
              </div>
            )}
            {messages.map((m, i) => (
              <AiMessageBubble
                key={i}
                message={m}
                accent={ACCENT}
                streaming={busy && i === messages.length - 1}
              />
            ))}
          </div>

          {/* 选区快捷操作 */}
          <div className="flex shrink-0 flex-wrap gap-1 border-t border-slate-100 px-2 pt-2">
            {BUILTIN_TEMPLATES.map((tpl) => (
              <button
                key={tpl.id}
                type="button"
                className="rounded-full border border-slate-200 px-2 py-0.5 text-[11px] text-slate-500 hover:border-[#107c41] hover:text-[#107c41] disabled:opacity-50"
                disabled={busy}
                onClick={() => runTemplate(tpl.id)}
              >
                {TEMPLATE_KEYS[tpl.id] ? t(TEMPLATE_KEYS[tpl.id]!.name) : tpl.name}
              </button>
            ))}
          </div>

          {/* 输入区：待发送图片 → 附件按钮 → 多行输入 → 发送/停止 */}
          <div className="shrink-0 border-t border-slate-100 p-2">
            {pendingImages.length > 0 && (
              <div className="mb-1.5 flex flex-wrap gap-1.5" data-testid="ai-pending-images">
                {pendingImages.map((img, i) => (
                  <div key={i} className="relative">
                    <img
                      src={img.dataUrl}
                      alt={t('excel.ai.imageAlt', { index: i + 1 })}
                      className="h-12 w-12 rounded-md border border-slate-200 object-cover"
                    />
                    <button
                      type="button"
                      className="absolute -top-1.5 -right-1.5 flex h-4 w-4 items-center justify-center rounded-full bg-slate-700 text-[10px] leading-none text-white"
                      title={t('excel.ai.removeImage')}
                      onClick={() => setPendingImages((prev) => prev.filter((_, idx) => idx !== i))}
                    >
                      ×
                    </button>
                  </div>
                ))}
              </div>
            )}
            <div className="flex items-end gap-1.5">
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                multiple
                className="sr-only"
                aria-hidden="true"
                tabIndex={-1}
                onChange={(e) => {
                  const files = e.target.files ? Array.from(e.target.files) : []
                  // 重置 value：连续选择同一张图片仍能触发 change
                  e.target.value = ''
                  void addImageFiles(files.filter((f) => f.type.startsWith('image/')))
                }}
              />
              <button
                type="button"
                className="rounded-md border border-slate-200 p-2 text-slate-500 hover:border-[#107c41] hover:text-[#107c41]"
                title={t('excel.ai.attach')}
                aria-label={t('excel.ai.attach')}
                data-testid="ai-attach"
                onClick={() => fileInputRef.current?.click()}
              >
                <Paperclip size={14} />
              </button>
              <textarea
                value={input}
                onChange={(e) => setInput(e.target.value)}
                onPaste={handlePaste}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault()
                    if (input.trim() || pendingImages.length) void doSend(input.trim(), undefined, pendingImages)
                  }
                }}
                rows={2}
                placeholder={t('excel.ai.inputPlaceholder')}
                className="flex-1 resize-none rounded-md border border-slate-200 px-2 py-1.5 text-xs outline-none focus:border-[#107c41]"
                data-testid="ai-input"
              />
              {busy ? (
                <button
                  type="button"
                  className="rounded-md bg-slate-200 px-3 py-2 text-xs text-slate-600"
                  onClick={() => abortRef.current?.abort()}
                >
                  {t('excel.ai.stop')}
                </button>
              ) : (
                <button
                  type="button"
                  className="rounded-md px-3 py-2 text-xs text-white disabled:opacity-50"
                  style={{ backgroundColor: ACCENT }}
                  disabled={!input.trim() && pendingImages.length === 0}
                  onClick={() => void doSend(input.trim(), undefined, pendingImages)}
                >
                  {t('excel.ai.send')}
                </button>
              )}
            </div>
          </div>
        </>
      )}
    </aside>
  )
}
