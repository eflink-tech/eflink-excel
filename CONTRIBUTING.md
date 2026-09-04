# 贡献指南

感谢关注 EFLink Excel！欢迎 Issue 与 PR。

## 环境要求

- Node.js >= 20
- pnpm >= 10（`corepack enable` 或 `npm i -g pnpm`）

## 本地开发

```bash
pnpm install
pnpm dev          # 组件库 watch 构建 + demo dev server
pnpm lint         # oxlint
pnpm typecheck    # tsc -b
pnpm test         # vitest 单元测试
pnpm test:e2e     # Playwright 端到端（自动构建并启动 demo）
```

提交 PR 前，请确保 `lint / typecheck / test / test:e2e` 全部通过，并为新功能/修复补充测试。

## 提交规范

Commit message 使用 [Conventional Commits](https://www.conventionalcommits.org/zh-hans/)：

- `feat: ...` 新功能
- `fix: ...` 缺陷修复
- `docs: ...` 文档
- `refactor: ...` 重构
- `test: ...` 测试
- `chore: ...` 构建/工具链

## 发版流程（维护者）

使用 [Changesets](https://github.com/changesets/changesets) 管理：

1. 功能分支合入 `main` 前执行 `pnpm changeset` 生成变更记录并随 PR 提交
2. `main` 上执行 `pnpm version-packages` 消费变更、更新版本号与 CHANGELOG
3. `pnpm release` 构建并发布到 npm（`@eflink-tech/excel`，公开访问）

## 涉及 Univer 升级

`packages/excel/scripts/patch-univer.js` 依赖上游构建产物的字符串锚点。升级 Univer 版本后请运行 `pnpm --filter @eflink-tech/excel postinstall`，按输出中的告警人工核对失效锚点。
