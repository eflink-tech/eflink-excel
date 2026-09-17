# xlsx 导入导出增强设计（尽量全档 + 富文本）

日期：2026-09-17
状态：待评审
范围：`eflink-excel/packages/excel`（不影响宿主 eflink-frontend 代码，宿主无需改动）

## 背景与目标

包内已有基于 exceljs 的 xlsx 导入导出雏形（`src/core/xlsx/`，双向约 120 行映射 + 测试），
但未接入 UI，且样式兼容仅 6 个字段（粗体/斜体/字号/字色/背景/数字格式）。

本次目标：

1. UI 挂上「导入 Excel (.xlsx)」「导出 Excel (.xlsx)」菜单项
2. 导入行为：**替换当前文档**（二次确认）
3. 公式策略：**保留公式字符串，Univer 引擎重算**（Excel 独有函数显示 `#NAME?` 为已接受风险）
4. 兼容档位「尽量全」：边框、对齐、换行、下划线/删除线、字体族、冻结窗格、隐藏行列
5. 富文本分段样式一并支持

## 兼容能力清单（最终承诺）

### 完整支持

| 维度 | 说明 |
|---|---|
| 值 | 文本 / 数字 / 布尔 / 错误码 / 多行文本；日期→`yyyy-MM-dd` 字符串（保留数字格式 pattern） |
| 公式 | 保留公式字符串（`f`），Univer 重算 |
| 结构 | 多 sheet / 合并单元格 / 行高列宽 / 冻结窗格 / 隐藏行列 |
| 样式 | 粗体 / 斜体 / 下划线 / 删除线 / 字号 / 字色 / 字体族 / 背景色 / 四边边框 / 水平垂直对齐 / 自动换行 / 数字格式 |
| 富文本 | 段内粗/斜/下划线/删除线/字号/字色/字体族、多段落（`\n`） |

### 近似映射（接受视觉误差）

- 边框线型：exceljs 与 Univer 均沿用 OOXML 线型词表（thin/hair/dotted/dashed/dashDot/
  dashDotDot/double/medium/mediumDashed/mediumDashDot/mediumDashDotDot/slantDashDot/thick），
  经映射表**全量 1:1 直映**（Univer `BorderStyleTypes` 数值枚举 1–13），无视觉损失
- 列宽：字符宽↔px 经验换算维持现状（不同字体有像素级偏差）

### 明确不支持（导入静默丢弃，导出不生成）

图表、图片/绘图对象、条件格式、数据验证、透视表、批注、受保护工作表、宏、
段内图片/上下标/链接文本、`.xls` 老格式（exceljs 仅支持 `.xlsx/.xlsm`，选择器限扩展名）、
超链接（放二期：Univer 已装 hyper-link 插件、exceljs 可读 `{text, hyperlink}`，后续补映射即可）。

## 设计

### 1. 数据模型扩展（`types/spreadsheet.ts`）

`CellStyle` 向 Univer `IStyleData` 形状对齐扩展（现有键名本就沿用 Univer 惯例）：

```ts
interface BorderStyle { s: number; cl: { rgb: string } }   // s: Univer BorderStyleTypes 数值枚举（1=THIN … 13=THICK）
export interface CellStyle {
  bl?: 0 | 1; it?: 0 | 1; fs?: number;
  cl?: { rgb: string }; bg?: { rgb: string };
  n?: { pattern: string };
  ff?: string;                                    // 字体族
  ul?: { s: 0 | 1 }; st?: { s: 0 | 1 };           // 下划线 / 删除线
  ht?: number;                                    // 水平对齐（Univer HorizontalAlign：1=左 2=中 3=右）
  vt?: number;                                    // 垂直对齐（Univer VerticalAlign：1=上 2=中 3=下）
  tb?: number;                                    // 换行策略（Univer WrapStrategy：3=自动换行）
  bd?: { t?: BorderStyle; b?: BorderStyle; l?: BorderStyle; r?: BorderStyle };
}
```

`SnapshotSheet` 新增 `freeze?: { startRow: number; startColumn: number; xAxisSplit?: number; yAxisSplit?: number }`；
隐藏行列在现有 `rowData`/`columnData` 条目上加 `hd?: 0 | 1`。
另加 `SnapshotCell.p?: RichTextDoc`（Univer `IDocumentData` 的子集类型：`dataStream` + `textRuns: { st, ed, ts }[]`），
形状与 Univer 一致，快照仍 `as never` 直通 `createWorkbook`，持久化链路零改动。

依据：运行时快照是 Univer 原样持久化的，边框/对齐/富文本在「编辑→保存→打开」链路里已天然存活；
本次扩展打通的是 xlsx 映射层与 TS 类型面。

### 2. 导入映射（`workbookToSnapshot.ts`）

- `readStyle` 补：边框（exceljs `border.{top,bottom,left,right}` → `bd`，含线型映射表与 argb→`#rgb`）、
  `alignment.{horizontal,vertical,wrapText}` → `ht/vt/tb`、`font.{underline,strikethrough,name}` → `ul/st/ff`
- `readCell` 富文本分支改为构建 `cell.p`：`richText[]` 拼接 `dataStream`（段间 `\r\n`），
  按段切 `textRuns`，run 样式映射进 `ts`；单元格级样式 `s` 取 runs 公共样式（首个 run）
- `readSheet` 补：`ws.views[0].state === 'frozen'` 的 `xSplit/ySplit/topLeftCell` → `freeze`；
  `row.hidden` / `column.hidden` → `hd` 标记
- 公式分支维持现状（保留 `f` 字符串）

### 3. 导出映射（`snapshotToWorkbook.ts`）

- `applyStyle` 补双向字段：`bd` → exceljs `border`（线型反映射）、`ht/vt/tb` → `alignment`、`ul/st/ff` → `font`
- `cell.p` → `toCellValue` 返回 `richText[]`：遍历 `dataStream` + `textRuns`，run 样式
  `ts` 映射到 exceljs font；段落边界插入 `\n`
- `freeze` → `ws.views = [{ state: 'frozen', xSplit, ySplit, topLeftCell }]`；`hd` → `row/column.hidden`

### 4. UI 与交互（包内 `MenuDropdown` + `fileActions.ts`）

菜单新增（与现有 `.efexcel` 导出 / 导出图片并列）：

- **「导入 Excel (.xlsx)」**：隐藏 `<input type="file" accept=".xlsx,.xlsm">` 触发选择 →
  `requestConfirm`（提醒当前内容将被替换）→ `importXlsx` → **先解析成功才落盘**：
  经默认存储 `save` 替换当前 docId 内容、`rename` 标题为导入文件名 →
  `editorStore.reloadToken++` 触发 `SheetEditor` 重载（docId 不变，复用既有 load 流程）→ toast「导入成功」
- **「导出 Excel (.xlsx)」**：复用 `exportEfexcelAction` 流程骨架：`saveNow()` → `load(docId)` →
  `exportXlsx(snapshot, title)` → toast「已导出 xlsx」
- 失败路径：解析失败 toast「导入失败：无法解析该文件」，当前文档内容不受影响；
  `console.error` 记录原始错误

新增部件仅一个：`editorStore.reloadToken` 计数器 + `SheetEditor` 订阅重载。

### 5. 测试

- `xlsx.test.ts` 扩展：边框（逐线型枚举）/ 对齐 / 换行 / 下划线 / 删除线 / 字体族 /
  冻结 / 隐藏行列 / 富文本（多段多 run）的**双向 roundtrip 断言**
  （snapshot → `wb.xlsx.writeBuffer` → `wb.xlsx.load` → snapshot）
- `fileActions.test.ts`：导入确认取消不落盘、导入成功重载、解析失败当前文档不受影响、导出先保存
- 目标：`src/core/xlsx` 与 `fileActions` 覆盖率不低于现有水平

## 风险

1. Univer 对导入公式重算时，不支持函数显示 `#NAME?`（已确认接受；不做逐格回退）
2. 富文本 `IDocumentData` 需要构造最小合法文档骨架（`documentStyle` 等），细节以 Univer
   0.25 实际渲染要求为准，实现时以 roundtrip 测试验证
3. 列宽换算沿用经验值，字符级精度不承诺
