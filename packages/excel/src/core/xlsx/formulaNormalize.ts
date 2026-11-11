import type { SnapshotCell, WorkbookSnapshot } from '../../types/spreadsheet';

/**
 * 修正历史导入留下的双等号公式（WPS 源文件 <f>=SUM(...)</f> 叠加导入补 '=' 的旧缺陷）。
 * 加载时自愈，已保存到云端/本地的受影响文档无需重新导入即可恢复引用公式。
 */
export function normalizeSnapshotFormulas(snapshot: WorkbookSnapshot): boolean {
  let changed = false;
  for (const sheet of Object.values(snapshot.sheets ?? {})) {
    for (const row of Object.values(sheet.cellData ?? {})) {
      for (const cell of Object.values(row) as SnapshotCell[]) {
        if (typeof cell.f === 'string' && cell.f.startsWith('==')) {
          cell.f = cell.f.slice(1);
          changed = true;
        }
      }
    }
  }
  return changed;
}

/**
 * 补全单元格值类型（Univer CellValueType）。
 * 导入/旧文档的 cellData 只有 v 没有 t，而条件格式的数字/平均值/排名规则会按
 * `cell.t === NUMBER` 过滤，缺 t 时规则虽已加载却匹配不到任何单元格；加载时统一补齐，
 * 新建与历史文档都能正常渲染条件格式。
 */
export function normalizeSnapshotCellTypes(snapshot: WorkbookSnapshot): boolean {
  let changed = false;
  for (const sheet of Object.values(snapshot.sheets ?? {})) {
    for (const row of Object.values(sheet.cellData ?? {})) {
      for (const cell of Object.values(row) as SnapshotCell[]) {
        if (cell.t !== undefined) continue;
        if (typeof cell.v === 'number') cell.t = 2;
        else if (typeof cell.v === 'boolean') cell.t = 3;
        else if (typeof cell.v === 'string') cell.t = 1;
        else continue;
        changed = true;
      }
    }
  }
  return changed;
}
