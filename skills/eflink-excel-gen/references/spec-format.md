# SheetSpec 字段参考

```jsonc
{
  "title": "工作簿标题",            // 必填，.efx.json 的 title、预览标题
  "workbookName": "文件名",         // 可选，快照 name；缺省用 title
  "snapshotId": "workbook-xxx",     // 可选，固定快照 id（重跑不漂移）
  "docId": "doc-xxx",               // 可选，固定文档 id
  "sheets": [ /* 至少一个 */ ]
}
```

## 工作表

```jsonc
{
  "name": "月度销售",               // 必填，标签页名
  "id": "sheet-01",                // 可选，缺省 sheet-01 / sheet-02…
  "rowCount": 30,                  // 可选，自动不小于实际内容
  "columnCount": 10,
  "showGridlines": false,          // 可选，默认 true（显示网格线）
  "tabColor": "#2F5597",           // 可选，标签页颜色 #RRGGBB
  "hidden": true,                  // 可选，隐藏工作表
  "freeze": { "rows": 2, "columns": 0 },   // 冻结前 N 行 / 前 N 列
  "columns": [50, 100, 140],       // 可选，列宽 px（按列顺序）
  "rowHeights": { "0": 40, "1": 30 },      // 可选，行高 px（0 起）
  "defaultRowHeight": 24,          // 可选
  "defaultColumnWidth": 88,        // 可选
  "merges": ["A1:I1", "A11:F11"],  // 可选，A1 范围；只在左上角放值
  "styles": { "header": { ... }, "money": { ... } },
  "data": [
    [ /* 第 0 行 */ ],
    [ /* 第 1 行 */ ]
  ]
}
```

`rowCount / columnCount` 会被自动抬高到覆盖所有单元格与合并区，显式值只用于给空白区留空间。

## 单元格

| 写法 | 含义 |
| --- | --- |
| `"文本"` / `123` / `true` | 原始值 |
| `null` / `""` / 缺省 | 空单元格 |
| `{ "value": "华东", "style": "center" }` | 带样式 |
| `{ "formula": "=SUM(H3:H10)", "style": "money" }` | 公式 |
| `{ "formula": "=C3*D3", "value": 1200 }` | 公式 + 缓存值（可选） |
| `{ "value": 0.05, "style": "percent" }` | 数字 + 数字格式 |
| `{ "value": 1200, "format": "money" }` | 单元格级格式快捷写法 |

公式的单元格引用使用 **Excel 1-based A1**（第 0 行是第 1 行），与 `data` 下标差 1。

## 样式

```jsonc
{
  "bold": true,
  "italic": true,
  "underline": true,
  "strike": true,
  "fontSize": 11,                  // 6~96
  "font": "Microsoft YaHei",
  "color": "#1F3864",              // 字体色
  "background": "#DCE6F1",         // 填充色
  "align": "center",               // left | center | right
  "vAlign": "middle",              // top | middle | bottom
  "wrap": true,                    // 自动换行
  "numberFormat": "money",         // 预设或 Excel 格式串
  "rotation": 45,                  // 旋转角度
  "vertical": true,                // 竖排文字（等价 rotation -90 + v=1）
  "border": { "style": "thin", "color": "#C9D3E0" },
  "border": { "style": "thin", "color": "#C9D3E0", "sides": ["top", "bottom"] }
}
```

- 数字格式预设：`money` → `#,##0.00`、`integer` → `#,##0`、`percent` → `0.0%`、
  `date` → `yyyy-mm-dd`、`datetime` → `yyyy-mm-dd hh:mm`。也可写 `"0.00"` 等原始格式串。
- 边框样式：`thin hair dotted dashed dashDot dashDotDot double medium mediumDashed
  mediumDashDot mediumDashDotDot slantDashDot thick`，默认四边全加。
- 省略 `sides` 表示四边；只画下划线式分隔可写 `"sides": ["bottom"]`。

## 公式

转换器原样写入快照，编辑器打开后由 Univer 重算。预览 HTML 自带求值器，支持：

- 四则运算 `+ - * / ^ %`、比较 `= <> < > <= >=`、括号、数字/字符串；
- `SUM`、`AVERAGE`、`MIN`、`MAX`、`COUNT`；
- `COUNTIF(区域, 条件)`、`SUMIF(区域, 条件, 求和区域)`、`AVERAGEIF(区域, 条件, 求值区域)`；
- `ROUND(值, 位数)`、`ABS`、`IF(条件, 真值, 假值)`、`IFERROR`；
- A1 引用与 `$A$1` 绝对引用、区域 `H3:H10`。

不支持跨工作表引用（`Sheet1!A1`）、SUMPRODUCT、VLOOKUP 等；这些公式仍会写入快照，
只是预览显示为空，请用真实编辑器核对。

条件写法：`"华东"`、`">=60"`、`"<>已完成"`、`80`。

## 输出

| 文件 | 内容 | 用途 |
| --- | --- | --- |
| `<base>.efx.json` | `{source,id,title,snapshot,createdAt,updatedAt}` | 编辑器「文件 → 导入文件」 |
| `<base>.snapshot.json` | `WorkbookSnapshot` | 模板入库 `content` |
| `<base>.html` | 预览 | 自检/交付预览；`?sheet=N` 切页 |
