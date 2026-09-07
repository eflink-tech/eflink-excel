// AI 产物 → 工作表数据（从当前活动单元格写入表格；读取选区/已用区域作为对话上下文）
import type { ICellData } from '@univerjs/core'
import { currentSnapshot } from '../core/saveService'
import { activeWorkbook } from '../core/univer/commands'
import { getUniverApi } from '../core/univer/controller'
import type { SnapshotCell, SnapshotSheet } from '../types/spreadsheet'
import type { AIGenTable } from './types'

/** 上下文最多携带的行/列数，防止大表撑爆 prompt */
const CONTEXT_MAX_ROWS = 40
const CONTEXT_MAX_COLS = 12

/** 0 起始列号 → A1 列名（25→Z，26→AA） */
export function colName(col: number): string {
  let n = col + 1
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/** 0 起始行列 → A1 地址（如 B3） */
function cellRef(row: number, col: number): string {
  return `${colName(col)}${row + 1}`
}

/**
 * 把 AI 表格写入当前工作表：表头写入活动单元格所在行列（加粗），数据随后。
 * 文本按「能否转数值」自动落数字；以 = 开头的字符串按公式写入。
 * @returns 写入的单元格数
 */
export function writeAITable(table: AIGenTable): number {
  const wb = activeWorkbook()
  const sheet = wb?.getActiveSheet()
  if (!wb || !sheet) throw new Error('工作表尚未就绪')
  const active = wb.getActiveRange()
  const startRow = active?.getRow() ?? 0
  const startCol = active?.getColumn() ?? 0

  const matrix: ICellData[][] = [table.columns.map((c) => toCell(c))]
  for (const row of table.rows) {
    matrix.push(row.map((c) => toCell(c)))
  }
  const rowCount = matrix.length
  const colCount = matrix[0].length
  // 容量不足时自动扩表（写入越界会失败）
  const needRows = startRow + rowCount - sheet.getMaxRows()
  if (needRows > 0) sheet.insertRowsAfter(sheet.getMaxRows() - 1, needRows)
  const needCols = startCol + colCount - sheet.getMaxColumns()
  if (needCols > 0) sheet.insertColumnsAfter(sheet.getMaxColumns() - 1, needCols)
  sheet.getRange(startRow, startCol, rowCount, colCount).setValues(matrix)
  // 表头加粗
  sheet.getRange(startRow, startCol, 1, colCount).setFontWeight('bold')
  return rowCount * colCount
}

function toCell(value: string): ICellData {
  const text = value.trim()
  if (text.startsWith('=')) return { f: text }
  if (text !== '' && Number.isFinite(Number(text))) return { v: Number(text) }
  return { v: value }
}

export interface SheetContext {
  sheetName: string
  /** 选区 A1 地址，如 B3:D10；无选区为 null */
  selection: string | null
  /** 选区内容（TSV），无选区为 null */
  selectionData: string | null
  /** 已用区域内容（TSV，含表头），空表为 null */
  usedData: string | null
  /** 已用区域 A1 地址 */
  usedAddress: string | null
}

/** 读取当前工作表上下文（选区 + 已用区域），供 AI 感知表格内容 */
export function getSheetContext(): SheetContext {
  const wb = activeWorkbook()
  const sheet = wb?.getActiveSheet()
  const empty: SheetContext = {
    sheetName: sheet?.getSheetName() ?? '',
    selection: null,
    selectionData: null,
    usedData: null,
    usedAddress: null,
  }
  if (!wb || !sheet) return empty

  // 选区
  const active = wb.getActiveRange()
  if (active && active.getWidth() > 0 && active.getHeight() > 0) {
    const r = active.getRow()
    const c = active.getColumn()
    empty.selection = `${cellRef(r, c)}:${cellRef(r + active.getHeight() - 1, c + active.getWidth() - 1)}`
    const values = active.getValues()
    empty.selectionData = matrixToTsv(values.map((row) => row.map((v) => (v == null ? '' : String(v)))))
  }

  // 已用区域：从实时快照的 cellData 求非空单元格包围盒
  const snap = currentSnapshot()
  const snapSheet: SnapshotSheet | undefined = snap ? findActiveSheet(snap.sheets, sheet.getSheetName()) : undefined
  if (snapSheet) {
    const bounds = usedBounds(snapSheet)
    if (bounds) {
      const [maxRow, maxCol] = bounds
      const rows: string[][] = []
      for (let r = 0; r <= Math.min(maxRow, CONTEXT_MAX_ROWS - 1); r += 1) {
        const line: string[] = []
        for (let c = 0; c <= Math.min(maxCol, CONTEXT_MAX_COLS - 1); c += 1) {
          line.push(cellText(snapSheet.cellData[r]?.[c]))
        }
        rows.push(line)
      }
      empty.usedAddress = `A1:${cellRef(maxRow, maxCol)}`
      empty.usedData = matrixToTsv(rows)
    }
  }
  return empty
}

function findActiveSheet(sheets: Record<string, SnapshotSheet>, name: string): SnapshotSheet | undefined {
  return Object.values(sheets).find((s) => s.name === name)
}

function usedBounds(sheet: SnapshotSheet): [number, number] | null {
  let maxRow = -1
  let maxCol = -1
  for (const rowKey of Object.keys(sheet.cellData)) {
    const r = Number(rowKey)
    const row = sheet.cellData[Number(rowKey)]
    if (!row) continue
    for (const colKey of Object.keys(row)) {
      const cell = row[Number(colKey)]
      if (cell && (cell.v !== undefined || cell.f !== undefined)) {
        if (r > maxRow) maxRow = r
        const c = Number(colKey)
        if (c > maxCol) maxCol = c
      }
    }
  }
  return maxRow >= 0 && maxCol >= 0 ? [maxRow, maxCol] : null
}

function cellText(cell: SnapshotCell | undefined): string {
  if (!cell) return ''
  if (cell.f) return cell.f
  if (cell.v !== undefined) return String(cell.v)
  return ''
}

function matrixToTsv(rows: string[][]): string {
  return rows.map((r) => r.join('\t')).join('\n')
}

/** 供 system prompt 使用的上下文描述 */
export function describeContext(ctx: SheetContext): string {
  if (!ctx.usedData && !ctx.selectionData) return '当前工作表为空。'
  const parts: string[] = [`当前工作表：${ctx.sheetName}`]
  if (ctx.usedData) {
    parts.push(`已用区域 ${ctx.usedAddress}（TSV，最多 ${CONTEXT_MAX_ROWS} 行 × ${CONTEXT_MAX_COLS} 列）：\n${ctx.usedData}`)
  }
  if (ctx.selectionData) {
    parts.push(`当前选区 ${ctx.selection}（TSV）：\n${ctx.selectionData}`)
  }
  return parts.join('\n')
}

/** Univer API 就绪检查（面板挂载时机用） */
export function isUniverReady(): boolean {
  return getUniverApi() != null
}
