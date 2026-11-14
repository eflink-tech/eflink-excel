---
name: eflink-excel-gen
description: "生成易飞表格（eflink-excel / Univer）可编辑文件。从文字描述或截图（如 Excel/ProcessOn 表格模板）产出 .efx.json（编辑器『导入文件』直接打开）+ .snapshot.json（模板入库 content）+ HTML 多工作表预览，支持多工作表、合并单元格、列宽行高、冻结窗格、隐藏工作表/标签色、字体/底色/边框/对齐/自动换行/数字格式（金额、百分比、整数、日期），以及 SUM/AVERAGE/MIN/MAX/COUNT/COUNTIF/SUMIF/ROUND/IF 等公式。Use when: (1) 用户要求生成台账/对账单/预算表/排班表/统计表等 Excel 表格或模板, (2) 提到 .efx.json / efexcel / 易飞表格 / 易飞 Excel / efexcel 模板, (3) 把截图或 CSV/Markdown 表格转成易飞表格可编辑文件, (4) 批量生产 Excel 模板入库。"
---

# eflink-excel 表格生成

把表格描述转成易飞表格编辑器（Univer）原生的 **WorkbookSnapshot**，并包装成可导入文件。
核心是写 **SheetSpec JSON**（二维单元格数组，无坐标），用打包的转换器生成：

- `.efx.json` —— `SheetDocument + source`，编辑器顶栏「文件 → 导入文件」直接打开；
- `.snapshot.json` —— 纯 `WorkbookSnapshot`，模板入库时作为 `office_template_content.content` 原样写入；
- `.html` —— 多工作表预览（可切换工作表），用于自检和交付预览图。

## 工作流

1. **读约束**：先读 [references/constraints.md](references/constraints.md)（公式、合并、换行、样式继承的硬性要求）。
2. **写 spec**：字段与枚举见 [references/spec-format.md](references/spec-format.md)。
   完整示例参考 [assets/examples/sales-ledger.json](assets/examples/sales-ledger.json)
   （3 个工作表：销售台账 + 部门汇总 + 隐藏参数表，含合并、冻结、金额/百分比格式、SUM/COUNTIF/SUMIF/IF）。
3. **转换**：
   ```bash
   node <skill目录>/scripts/excel_convert.mjs my-spec.json -o ./out
   # 产出 my-spec.efx.json / my-spec.snapshot.json / my-spec.html
   # stderr 的 ⚠ 警告必须逐条核对（非法字段会被忽略而不是报错）
   ```
4. **自检**：转 PNG 检查版式（可 Read 图片）：
   ```bash
   "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu \
     --hide-scrollbars --window-size=1500,800 \
     --screenshot=preview.png "file://$(pwd)/out/my-spec.html?sheet=0"
   ```
   `?sheet=N` 切换工作表（0 起）。核对：标题/表头底色、边框、合并区、数字格式（千分位、百分比）、
   列宽是否截字、合计行是否与明细相加一致、冻结提示。预览公式为自算结果，用于核对数字；
   编辑器打开后会按 Univer 重算。
5. **交付**：给用户 `.efx.json` 路径（/office/excel 页面「文件 → 导入文件」）+ 预览图。
   模板入库用 `.snapshot.json`。

## 写作要点

- 一张表一个 `sheets[]` 元素；`data` 是行数组，行内是列数组，单元格可以是原始值或 `{value, formula, style}`。
- 样式用 `styles` 命名复用（`"header"`、`"money"` 等），单元格写 `style:"header"`；也可直接内联对象。
- 公式只写 `formula`，不要自己填计算结果 `value`（需要缓存值时才同时给 `value`，且必须与公式一致）。
- 合并用 A1 范围字符串 `"A1:I1"`，只在左上角单元格放值与样式，其余留 `null`。
- 列宽单位 px；正文数字列 70~120，名称列 120~160。不设时按 88px。
- 数字格式用预设 `money / integer / percent / date / datetime`，或直接给 Excel 格式串（如 `"0.00"`）。
- 冻结用 `freeze: {rows: 2, columns: 0}`（冻结前 N 行/列），不要手写 `startRow/xSplit`。
- 隐藏工作表 `hidden: true`（预览仍显示但带「隐藏」标记，便于核对）。

## 可选：入库为模板草稿

`.snapshot.json` 作为 `office_template_content.content` 写入本地库（localhost:3306/eflink-tech），
`office_template` 行 type=**12**（excel）、status=0 草稿、author_name 标来源。批量入库参考
`eflink-backend/scripts/seed-templates.mjs` 的 upsert 模式，以及 `scripts/excel-templates/gen.mjs`
的 30 套模板 DSL 与 SQL 转义经验（先转义反斜杠再转义单引号）。涉及数据库写入先和用户确认。

## 维护提示

格式升级时对照：
- 类型定义：`eflink-excel/packages/excel/src/types/spreadsheet.ts`
- 导入/导出：`eflink-excel/packages/excel/src/core/efexcel.ts`（`.efx.json`）、
  `core/xlsx/workbookToSnapshot.ts`（边框/对齐/换行枚举的唯一权威来源）
- 新建空工作簿：`createEmptySnapshot()`；模板内容只认 `snapshot`，不含 source/title
- 转换器已在预览侧实现常用公式求值；新增函数时同步 `scripts/excel_convert.mjs` 的 `callFn`
