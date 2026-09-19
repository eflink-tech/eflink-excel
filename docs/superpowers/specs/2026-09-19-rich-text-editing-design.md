# eflink-excel 单元格内富文本编辑设计

日期：2026-09-19
状态：已确认（用户批准）

## 背景与现状

eflink-excel 数据层已支持单元格富文本（`p` 字段，形状对齐 Univer `IDocumentData` 子集）：

- **导入**：xlsx 富文本（exceljs `richText`）→ `p`（`workbookToSnapshot.ts` `applyRichText`）
- **导出**：`p` → exceljs `richText`（`snapshotToWorkbook.ts` `richTextToCellValue`）
- **渲染/持久化**：Univer 原生支持（快照直通 `createWorkbook`）

缺失的是**编辑能力**：

1. 双击进单元格编辑后，自定义工具栏样式操作走 `FRange` 整格 API，无法作用于编辑态中**选中的部分文字**（Univer 自带 ribbon 的编辑态命令被 `header: false` 关闭）
2. 公式栏提交用 `range.setValue(纯文本)`，会把富文本 `p` 打回纯文本

## 需求（已确认）

- **范围**：单元格内富文本编辑——编辑态下选中部分文字可设置样式，已有富文本导入后可继续编辑
- **样式能力**：加粗/斜体/下划线/删除线、字体颜色、字号/字体族、部分文字背景色（全选）
- **公式栏**：对齐 Excel 桌面版——公式栏提交后富文本退化为纯文本；需要改部分文字样式请双击进单元格

## 方案选型（已确认）

选定 **方案 A + C**：

- **A：工具栏编辑态路由到 doc 内联格式命令**（主体）
- **C：编辑态选区浮动格式条**（补充入口，同期实现）

否决方案：编辑态启用 Univer 自带 ribbon（样式与企微风格割裂、来回闪烁）。

## 总体架构

核心原则：样式命令永远作用于"当前焦点所在上下文"。工具栏 UI 一套，底层路由按编辑态分流。

```
SheetToolbar / FloatingFormatBar（UI 层，互为补充、共用路由）
        │ 调用
        ▼
commands.ts 路由层（新增 editMode 状态感知）
   ├─ 非编辑态 → FRange.setFontWeight 等（现状不变）
   └─ 编辑态   → univerAPI.executeCommand('doc.command.set-inline-format-*')
                 （doc 命令作用于编辑器内当前文本选区，
                   与 Univer 自家 ribbon 同路径，撤销/重走原生 mutation 体系）
        ▼
Univer 编辑器 ──提交──▶ 单元格 p 字段 ──▶ 快照/自动保存/xlsx 导入导出（现有链路零改动）
```

- **编辑态感知**：门面事件 `Event.SheetEditStarted / SheetEditEnded`（Univer 0.25.1 已确认存在）驱动状态模块
- **数据流不变**：本期不新增任何持久化结构

## 组件拆分

| 文件 | 动作 | 职责 |
|---|---|---|
| `core/univer/editMode.ts` | 新增 | 编辑态状态模块：订阅编辑态事件，暴露 `isEditing()` / `subscribeEditState()` |
| `core/univer/inlineFormat.ts` | 新增 | doc 内联格式命令封装与参数组装；toggle 判断（从编辑器文本选区读样式交集） |
| `core/univer/commands.ts` | 修改 | 样式入口（加粗/斜体/下划线/删除线、字色/背景色、字号、字体族）加编辑态分流 |
| `components/toolbar/FloatingFormatBar.tsx` | 新增 | 选区浮动格式条组件 |
| `components/toolbar/FloatingFormatBar.css`（或并入 myfsheet.css） | 新增 | 浮动条样式 |
| `core/univer/controller.ts` | 修改 | mount 时订阅编辑态事件；浮动条挂载 |
| `components/toolbar/SheetToolbar.tsx` | 小改 | 编辑态按钮激活态同步（可选） |

**toggle 状态读取（编辑态）**：通过 docs-ui `TextSelectionManagerService`（preset-sheets-core 传递依赖）读当前文本选区覆盖 run 的样式交集——选中文字**全部**加粗才显示激活，与现有 `currentStyle()` 语义对齐。

**命令执行单元**：doc 命令经 `univerAPI.executeCommand()` 派发，Univer 按焦点单元路由到编辑器文档，预期免 unitId（spike 验证；如需显式 unitId，在 `inlineFormat.ts` 收口）。

## 浮动格式条交互

- **触发**：编辑态下编辑器内文本选区非空且几何位置可获取时显示
- **定位**：以选区屏幕坐标为锚，浮于上方约 8px；贴近视口顶部翻转到下方；不遮输入光标
- **内容**：B / I / U / 删除线 ｜ 字色、背景色（复用 `components/color` 调色板）｜ 字号、字体族下拉。自身零样式逻辑，全部走 `commands.ts` 路由
- **消失**：选区清空、退出编辑态、点击选区外
- **与顶部工具栏**：仅入口不同，激活态由同一状态源驱动

## 边界情况与既定决策

1. **公式栏**：提交逻辑不改（对齐 Excel）；小改进——含富文本单元格显示纯文本内容（取 `dataStream`），避免误以为内容丢失
2. **整格样式覆盖**：非编辑态对含 `p` 单元格设整格样式，run 样式优先级由 Univer 原生语义决定，不额外处理（spike 验证）
3. **撤销/重做**：doc 命令走 Univer mutation 体系，Ctrl+Z 原生生效，不自研
4. **粘贴**：富文本粘贴由 Univer 原生管线处理，本期只验证不加工
5. **Spike（实现第一步，两个假设）**：
   - 双击含 `p` 单元格，编辑器内 run 样式原样呈现、可继续编辑
   - `doc.command.set-inline-format-*` 免 unitId 直达编辑器
   任一不成立，在 `editMode.ts` / `inlineFormat.ts` 内收口解决，不影响整体架构

## 测试策略

- **单元测试**：`editMode.ts` 状态机；`inlineFormat.ts` 参数组装与 toggle 交集判断（mock TextSelectionManagerService）；commands.ts 分流
- **E2E（Playwright）**：双击含富文本单元格 → 选中部分文字 → 浮动条加粗/变色 → 提交 → 重开验证 run 保留；导出 xlsx 后 exceljs 读回验证 richText 结构
- **手动验收**：企微表格对照

## 不做的事（YAGNI）

- 不做公式栏富文本编辑
- 不做部分文字背景色以外的文档级样式（对齐、段落等）
- 不做 Univer 原生 ribbon 的任何启用
- 不新增持久化结构
