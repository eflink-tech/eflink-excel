// AI 数据持久化（Dexie）：设置 / 会话历史 / 内置提示词模板
import Dexie, { type Table } from 'dexie'
import { isLegacyImportDone, markLegacyImportDone, readLegacyIndexedDb } from '@eflink-tech/office-kit'
import { EXCEL_AI_DB_NAME } from '../storage/keys'
import type { AISettings, Conversation, Template } from './types'

export class AIDb extends Dexie {
  conversations!: Table<Conversation, string>
  settings!: Table<AISettings & { key: string }, string>
  templates!: Table<Template, string>

  constructor() {
    super(EXCEL_AI_DB_NAME)
    this.version(1).stores({
      conversations: 'id, updatedAt',
      settings: 'key',
      templates: 'id, builtin',
    })
  }
}

export const aiDb = new AIDb()

/** 旧库名（efLink 品牌时期）：只读搬运，绝不删除 */
const LEGACY_AI_DB_NAME = 'eflink-excel-ai'
const LEGACY_AI_STORES = ['conversations', 'settings', 'templates'] as const

let legacyImport: Promise<void> | null = null

/**
 * 一次性把旧 AI 库的数据搬进新库：仅在该表为空时写入；失败只告警，不阻断 AI 面板。
 * 模块级 memoize，一个会话只跑一次。
 */
export function ensureLegacyImport(): Promise<void> {
  if (!legacyImport) {
    legacyImport = importLegacyAiData().catch((err) => {
      console.warn('[ai/db] 旧 AI 数据搬运失败', err)
    })
  }
  return legacyImport
}

async function importLegacyAiData(): Promise<void> {
  const legacy = await readLegacyIndexedDb(LEGACY_AI_DB_NAME, LEGACY_AI_STORES)
  if (!legacy) return
  // 每表独立一次性标记（防「删光复活」）；bulkPut 单事务，失败不落标记，下个会话重试
  await importLegacyTable('conversations', aiDb.conversations, legacy.conversations as Conversation[])
  await importLegacyTable('settings', aiDb.settings, legacy.settings as (AISettings & { key: string })[])
  await importLegacyTable('templates', aiDb.templates, legacy.templates as Template[])
}

async function importLegacyTable<T>(scope: string, table: Table<T, string>, rows: T[]): Promise<void> {
  const markerKey = `doculuma:excel:legacy-ai-imported:${scope}`
  if (isLegacyImportDone(markerKey)) return
  if (!rows?.length) {
    markLegacyImportDone(markerKey)
    return
  }
  if ((await table.count()) > 0) {
    markLegacyImportDone(markerKey)
    return
  }
  await table.bulkPut(rows)
  markLegacyImportDone(markerKey)
}

const SETTINGS_KEY = 'main'
const DEFAULT_SETTINGS: AISettings = { baseUrl: '', apiKey: '', model: '' }

export const aiSettingsStore = {
  async getSettings(): Promise<AISettings> {
    await ensureLegacyImport()
    const row = await aiDb.settings.get(SETTINGS_KEY)
    if (!row) return { ...DEFAULT_SETTINGS }
    return { baseUrl: row.baseUrl, apiKey: row.apiKey, model: row.model }
  },

  async saveSettings(s: AISettings): Promise<void> {
    await aiDb.settings.put({ key: SETTINGS_KEY, ...s })
  },

  async isConfigured(): Promise<boolean> {
    const s = await this.getSettings()
    return Boolean(s.baseUrl && s.apiKey && s.model)
  },
}

/** 会话历史（按更新时间倒序）；读取前先搬运旧库数据 */
export async function listConversations(): Promise<Conversation[]> {
  await ensureLegacyImport()
  return aiDb.conversations.orderBy('updatedAt').reverse().toArray()
}

/** 选区/已用区域快捷操作模板：{data} 替换为当前表格数据（TSV 文本） */
export const BUILTIN_TEMPLATES: Template[] = [
  {
    id: 'tpl-analyze',
    name: '分析数据',
    builtin: true,
    content: '请分析以下表格数据（TSV 格式，首行为表头），给出数据洞察、趋势与异常点，用简洁的中文条目列出：\n{data}',
  },
  {
    id: 'tpl-formula',
    name: '建议公式',
    builtin: true,
    content:
      '以下是表格数据（TSV 格式，首行为表头）。请针对这份数据给出常用的汇总/统计公式建议，每条给出可直接使用的公式（用 A1 样式引用）并一句话说明用途：\n{data}',
  },
  {
    id: 'tpl-clean',
    name: '清洗建议',
    builtin: true,
    content:
      '以下是表格数据（TSV 格式，首行为表头）。请检查其中的数据质量问题（缺失、重复、格式不一致、异常值等），给出清洗建议：\n{data}',
  },
]
