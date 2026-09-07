// AI 提示词与结果解析：约定输出结构化 JSON，由 actionExecutor 写入工作表
import type { AIGenTable } from './types'

const MAX_ROWS = 200
const MAX_COLS = 30

export const TABLE_SCHEMA_PROMPT = `输出 JSON（不要输出 JSON 之外的任何文字），结构如下：
{
  "name": "表格名称",
  "columns": ["列一", "列二", "列三"],
  "rows": [
    ["文本", "123", "2026-01-01"],
    ["文本", "456", "2026-02-01"]
  ]
}
约定：columns 为表头；rows 每行列数与 columns 一致；数字不要加千分位或单位；需要计算的列可给出以 = 开头的公式（用 A1 样式引用写入位置的第一行数据）；不要输出 markdown 代码块。`

/** 生成表格的提示词 */
export function buildTablePrompt(topic: string, rowCount: number, requirements: string): string {
  return `请围绕主题「${topic}」生成一份表格，约 ${rowCount} 行数据（不含表头）。
${requirements ? `要求：${requirements}\n` : ''}${TABLE_SCHEMA_PROMPT}`
}

export function parseJSONFromText(text: string): unknown {
  // 去掉可能的 markdown 代码块包裹
  const cleaned = text.replace(/```(?:json)?/g, '').trim()
  try {
    return JSON.parse(cleaned)
  } catch {
    // 尝试截取第一个 { 到最后一个 }
    const start = cleaned.indexOf('{')
    const end = cleaned.lastIndexOf('}')
    if (start >= 0 && end > start) {
      return JSON.parse(cleaned.slice(start, end + 1))
    }
    throw new Error('AI 返回的内容不是有效的 JSON')
  }
}

/** 解析并校验 AI 表格：非法行列丢弃、超量截断、字段类型归一 */
export function parseAITable(text: string): AIGenTable {
  const data = parseJSONFromText(text) as Partial<AIGenTable>
  if (!data || !Array.isArray(data.columns)) throw new Error('表格 JSON 结构不正确')
  const columns = data.columns.map((c) => String(c ?? '')).slice(0, MAX_COLS)
  if (!columns.length) throw new Error('表格缺少表头')
  const rows: string[][] = []
  for (const raw of (data.rows ?? []).slice(0, MAX_ROWS)) {
    if (!Array.isArray(raw)) continue
    const row = raw.map((c) => (c == null ? '' : String(c)))
    while (row.length < columns.length) row.push('')
    rows.push(row.slice(0, columns.length))
  }
  if (!rows.length) throw new Error('表格没有数据行')
  return {
    name: typeof data.name === 'string' ? data.name : undefined,
    columns,
    rows,
  }
}
