import { describe, expect, it } from 'vitest'
import { buildTablePrompt, parseAITable, parseJSONFromText } from './systemPrompt'

describe('parseJSONFromText', () => {
  it('解析裸 JSON', () => {
    expect(parseJSONFromText('{"a":1}')).toEqual({ a: 1 })
  })
  it('剥掉 markdown 代码块', () => {
    expect(parseJSONFromText('```json\n{"a":1}\n```')).toEqual({ a: 1 })
  })
  it('截取首尾大括号', () => {
    expect(parseJSONFromText('好的，以下是结果：{"a":1} 请查收')).toEqual({ a: 1 })
  })
  it('非法输入抛错', () => {
    expect(() => parseJSONFromText('不是JSON')).toThrow()
  })
})

describe('parseAITable', () => {
  it('解析表头与数据行', () => {
    const table = parseAITable('{"name":"销售","columns":["月份","金额"],"rows":[["1月","100"],["2月","200"]]}')
    expect(table.name).toBe('销售')
    expect(table.columns).toEqual(['月份', '金额'])
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0]).toEqual(['1月', '100'])
  })
  it('短行补齐、超长行截断、非数组行丢弃', () => {
    const table = parseAITable('{"columns":["A","B","C"],"rows":[["1"],["x","y","z","extra"],null,["p","q","r"]]}')
    expect(table.rows[0]).toEqual(['1', '', ''])
    expect(table.rows[1]).toEqual(['x', 'y', 'z'])
    expect(table.rows[2]).toEqual(['p', 'q', 'r'])
  })
  it('null 单元格转空串', () => {
    const table = parseAITable('{"columns":["A"],"rows":[[null]]}')
    expect(table.rows[0]).toEqual([''])
  })
  it('缺表头或无数据行抛错', () => {
    expect(() => parseAITable('{"rows":[["1"]]}')).toThrow()
    expect(() => parseAITable('{"columns":["A"],"rows":[]}')).toThrow()
    expect(() => parseAITable('{"a":1}')).toThrow()
  })
})

describe('buildTablePrompt', () => {
  it('包含主题、行数与要求', () => {
    const prompt = buildTablePrompt('季度销售', 20, '含合计列')
    expect(prompt).toContain('季度销售')
    expect(prompt).toContain('20')
    expect(prompt).toContain('含合计列')
    expect(prompt).toContain('"columns"')
  })
})
