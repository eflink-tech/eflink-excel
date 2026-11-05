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
