// AI 提示词与结果解析：约定输出结构化 JSON，由 actionExecutor 写入工作表
import { excelT } from '../i18n'
import type { AIGenTable } from './types'

const MAX_ROWS = 200
const MAX_COLS = 30

/** 生成表格的提示词 */
export function buildTablePrompt(topic: string, rowCount: number, requirements: string): string {
  const requirementsLine = requirements
    ? excelT('excel.ai.requirementsLine', { requirements })
    : ''
  return excelT('excel.ai.tablePrompt', {
    topic,
    rows: rowCount,
    requirements: requirementsLine,
    schema: excelT('excel.ai.schemaPrompt'),
  })
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
    throw new Error(excelT('excel.ai.invalidJson'))
  }
}

/** 解析并校验 AI 表格：非法行列丢弃、超量截断、字段类型归一 */
export function parseAITable(text: string): AIGenTable {
  const data = parseJSONFromText(text) as Partial<AIGenTable>
  if (!data || !Array.isArray(data.columns)) throw new Error(excelT('excel.ai.invalidTable'))
  const columns = data.columns.map((c) => String(c ?? '')).slice(0, MAX_COLS)
  if (!columns.length) throw new Error(excelT('excel.ai.missingColumns'))
  const rows: string[][] = []
  for (const raw of (data.rows ?? []).slice(0, MAX_ROWS)) {
    if (!Array.isArray(raw)) continue
    const row = raw.map((c) => (c == null ? '' : String(c)))
    while (row.length < columns.length) row.push('')
    rows.push(row.slice(0, columns.length))
  }
  if (!rows.length) throw new Error(excelT('excel.ai.missingRows'))
  return {
    name: typeof data.name === 'string' ? data.name : undefined,
    columns,
    rows,
  }
}
