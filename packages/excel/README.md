# @eflink-tech/excel

开箱即用的在线表格编辑器 React 组件（基于 [Univer](https://github.com/dream-num/univer)）。可独立运行，也可作为组件嵌入任意 React 应用。

## 安装

```bash
npm install @eflink-tech/excel
# react / react-dom >= 18 为 peer 依赖
```

## 使用

```tsx
import { SheetEditor } from '@eflink-tech/excel';
import '@eflink-tech/excel/styles.css';

<div style={{ height: '100vh' }}>
  <SheetEditor docId={docId} branding={{ logo: '/logo.png', name: '我的表格' }} />
</div>
```

存储可插拔（默认 IndexedDB，可注入 `StorageAdapter` 对接后端）；导出 `exportEfexcel / importEfexcel`（.efexcel JSON 备份）、`importXlsx / exportXlsx / exportPng` 等 headless API 与全部数据类型。

> **Tailwind 说明**：组件库内部布局用到少量 Tailwind 工具类（已随 `styles.css` 提供回退样式）。若宿主使用 Tailwind v4 且希望得到与 demo 一致的布局，请在入口 CSS 中显式扫描组件包：
>
> ```css
> @import "tailwindcss";
> @source "../node_modules/@eflink-tech/excel";
> ```

完整文档（props、存储适配器、本地开发、Univer 补丁说明）见仓库根 README：

https://github.com/eflink-tech/eflink-excel

## License

MIT
