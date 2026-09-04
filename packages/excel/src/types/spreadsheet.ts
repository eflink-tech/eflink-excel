// 表格领域类型：WorkbookSnapshot 是 Univer snapshot (IWorkbookData) 的结构子集，
// 仅覆盖本应用读写的字段；与 Univer 的类型对接只在 UniverContainer 处做一次 cast。

export interface CellStyle {
  bl?: 0 | 1;
  it?: 0 | 1;
  fs?: number;
  cl?: { rgb: string };
  bg?: { rgb: string };
  n?: { pattern: string };
}

export interface SnapshotCell {
  v?: string | number | boolean;
  f?: string;
  /** 内联样式对象（导入方向）或 Univer 归一化样式 id（运行时快照，经 workbook.styles 解析） */
  s?: CellStyle | string;
}

export interface MergeRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

export interface SnapshotSheet {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cellData: Record<number, Record<number, SnapshotCell>>;
  mergeData?: MergeRange[];
  rowData?: Record<number, { h: number }>;
  columnData?: Record<number, { w: number }>;
}

export interface WorkbookSnapshot {
  id: string;
  name: string;
  sheetOrder?: string[];
  sheets: Record<string, SnapshotSheet>;
  /** Univer 运行时快照把单元格样式归一化到这里（cell.s 存样式 id）；导入方向的快照用内联样式，无此字段 */
  styles?: Record<string, CellStyle>;
}

export interface DocumentMeta {
  id: string;
  title: string;
  updatedAt: number;
}

export interface SheetDocument {
  id: string;
  title: string;
  snapshot: WorkbookSnapshot;
  createdAt: number;
  updatedAt: number;
}

/** 新建空白工作簿快照（Univer 会为缺省字段补默认值） */
export function createEmptySnapshot(name = '未命名表格'): WorkbookSnapshot {
  return {
    id: `workbook-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    name,
    sheetOrder: ['sheet-01'],
    sheets: {
      'sheet-01': {
        id: 'sheet-01',
        name: 'Sheet1',
        rowCount: 100,
        columnCount: 20,
        cellData: {},
      },
    },
  };
}

export function createDocument(title = '未命名表格'): SheetDocument {
  const now = Date.now();
  return { id: newId(), title, snapshot: createEmptySnapshot(title), createdAt: now, updatedAt: now };
}

export function newId(): string {
  return `doc-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}
