/**
 * Univer 依赖补丁集合（幂等，可重复执行；postinstall 钩子自动运行）。
 * 每个补丁：锚点不存在时跳过并告警（上游升级后需人工核对），已打补丁时跳过。
 *
 * 1) sheets-ui mso-number-format：Excel/WPS/企微 导出的 HTML 用单引号包裹数字格式
 *    （mso-number-format:'yyyy"年"m"月"'），上游只剥离双引号，残留单引号使 numfmt
 *    抛 SyntaxError: Illegal character，导致此类内容粘贴失败。→ 增加剥离首尾单引号。
 * 2) sheets-filter-ui 面板定位：筛选面板 direction 由 horizontal（左右乱飘）改为
 *    vertical（固定在筛选列下方，空间不足自动上翻）。
 * 3) sheets-filter-ui 按钮 hover 与位置：hover 背景 gray.50 在白底上不可见 → gray.200，
 *    hover 时把光标设为 pointer；按钮垂直位置改为统一居中，避免表头单元格垂直对齐
 *    不一致时各列按钮高低错落。
 * 4) sheets-ui 冻结视口尺寸：冻结区合计宽高超过可视区域时，主视口被挤成负尺寸而无法滚动
 *    （xlsx 常见：冻结多列后在窄窗口打开）→ 冻结区尺寸钳制，为主视口保留最小尺寸。
 * 5) sheets-ui 编辑框跟随滚动：单元格内编辑时滚动网格，编辑框按打开时的视口坐标绘制，
 *    不随网格移动而悬浮 → 主视口滚动后重算编辑框位置（滚回时仍在原单元格编辑）。
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const packageDir = path.join(here, '..');
// 从本包 package.json 出发解析依赖真实路径：兼容 npm 扁平化与 pnpm 符号链接两种布局
const requireFromPackage = createRequire(path.join(packageDir, 'package.json'));

/** 直接依赖 / npm 扁平化布局：常规 require.resolve */
function resolveViaRequire(pkg) {
  try {
    return path.dirname(requireFromPackage.resolve(`@univerjs/${pkg}/package.json`));
  } catch {
    return null;
  }
}

/** pnpm 布局：传递依赖在 node_modules/.pnpm 提升区，沿目录树向上找最近的 store 取最高版本 */
function resolveViaPnpmStore(pkg) {
  const prefix = `@univerjs+${pkg}@`;
  let dir = packageDir;
  for (;;) {
    const store = path.join(dir, 'node_modules', '.pnpm');
    if (fs.existsSync(store)) {
      const best = fs
        .readdirSync(store)
        .filter((name) => name.startsWith(prefix))
        .map((name) => ({ name, ver: name.slice(prefix.length).split('_')[0] }))
        .sort((a, b) => {
          const pa = a.ver.split('.').map((n) => parseInt(n, 10) || 0);
          const pb = b.ver.split('.').map((n) => parseInt(n, 10) || 0);
          for (let i = 0; i < 3; i += 1) if ((pa[i] ?? 0) !== (pb[i] ?? 0)) return (pb[i] ?? 0) - (pa[i] ?? 0);
          return 0;
        })[0];
      return best ? path.join(store, best.name, 'node_modules', '@univerjs', pkg) : null;
    }
    const parent = path.dirname(dir);
    if (parent === dir) return null;
    dir = parent;
  }
}

/** 解析 @univerjs/<pkg> 的安装根目录；找不到返回 null */
function resolvePkgRoot(pkg) {
  return resolveViaRequire(pkg) ?? resolveViaPnpmStore(pkg);
}

/** 单个补丁定义 */
const PATCHES = [
  {
    pkg: 'sheets-ui',
    file: 'lib/es/index.js',
    name: 'mso-number-format 单引号剥离',
    anchor: 'if (value.startsWith("\\"") && value.endsWith("\\"")) value = value.slice(1, -1);',
    insert: 'if (value.startsWith("\'") && value.endsWith("\'")) value = value.slice(1, -1);',
  },
  {
    pkg: 'sheets-filter-ui',
    file: 'lib/es/index.js',
    name: '筛选面板定位到列下方',
    anchor: 'direction: "horizontal",',
    insert: 'direction: "vertical",',
  },
  {
    pkg: 'sheets-filter-ui',
    file: 'lib/es/index.js',
    name: '筛选按钮 hover 背景加强',
    anchor: 'this._hovered ? this._themeService.getColorFromTheme("gray.50") : "rgba(255, 255, 255, 1.0)"',
    insert: 'this._hovered ? this._themeService.getColorFromTheme("gray.200") : "rgba(255, 255, 255, 1.0)"',
  },
  {
    pkg: 'sheets-filter-ui',
    file: 'lib/es/index.js',
    name: '筛选按钮垂直居中（表头对齐不一致时错落）',
    anchor: 'top: computeIconTop(startY, endY, cellHeight, verticalAlign),',
    insert: 'top: startY + Math.max(0, (cellHeight - 16) / 2), // [patch-univer] 各列按钮统一垂直居中（表头对齐不一致时错落）',
  },
  {
    pkg: 'sheets-filter-ui',
    file: 'lib/es/index.js',
    name: '筛选按钮 hover 光标(pointer)',
    anchor: 'onPointerEnter() {\n\t\tthis._hovered = true;\n\t\tthis.makeDirty(true);\n\t}',
    insert: 'onPointerEnter() {\n\t\tthis._hovered = true;\n\t\tdocument.body.style.cursor = "pointer";\n\t\tthis.makeDirty(true);\n\t}',
  },
  {
    pkg: 'sheets-filter-ui',
    file: 'lib/es/index.js',
    name: '筛选按钮离开恢复光标',
    anchor: 'onPointerLeave() {\n\t\tthis._hovered = false;\n\t\tthis.makeDirty(true);\n\t}',
    insert: 'onPointerLeave() {\n\t\tthis._hovered = false;\n\t\tdocument.body.style.cursor = "";\n\t\tthis.makeDirty(true);\n\t}',
  },
  {
    pkg: 'drawing-ui',
    file: 'lib/es/index.js',
    name: '禁用图片双击预览浮层（双击图表改为打开属性面板）',
    anchor: '_addDialogForImage(o) {',
    insert: '_addDialogForImage(o) {\n\t\t// [patch-univer] 图片双击预览已禁用：双击图表改为打开图表属性面板\n\t\treturn;',
  },
  // 冻结区尺寸钳制：xlsx 里冻结列/行合计宽高可能超过可视区域（如冻结 A–E 列后在窄窗口打开），
  // 上游直接把主视口 left/top 设成冻结区尺寸，导致主视口宽高为负、判定 inactive，
  // 滚轮与滚动条全部失效（整表无法上下滚动）。这里给主视口保留至少 80px，冻结区超出部分裁切（与 Excel 一致）。
  {
    pkg: 'sheets-ui',
    file: 'lib/es/index.js',
    // 本包只发布 ESM（package.json 无 require 入口），cjs 产物无需同步
    files: ['lib/es/index.js'],
    name: '编辑中滚动网格时编辑框跟随单元格',
    anchor:
      '\t_init() {\n\t\tconst d = new DisposableCollection();\n\t\tthis._subscribeToCurrentCell(d);',
    insert:
      '\t_init() {\n\t\tconst d = new DisposableCollection();\n\t\t// [patch-univer] 编辑中滚动网格时重算编辑框位置，避免编辑框脱离单元格悬浮\n\t\tconst sheetUnit = this._univerInstanceService.getCurrentUnitOfType(UniverInstanceType.UNIVER_SHEET);\n\t\tconst sheetViewport = sheetUnit ? this._renderManagerService.getRenderById(sheetUnit.getUnitId())?.scene?.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN) : null;\n\t\tif (sheetViewport) d.add(sheetViewport.onScrollAfter$.subscribeEvent(() => {\n\t\t\tif (!this._editorBridgeService.isVisible().visible) return;\n\t\t\tthis._editorBridgeService.refreshEditCellPosition(false);\n\t\t\tthis._sheetCellEditorResizeService?.fitTextSize();\n\t\t}));\n\t\tthis._subscribeToCurrentCell(d);',
  },
  {
    pkg: 'sheets-ui',
    file: 'lib/es/index.js',
    name: '冻结列超出可视宽度时主视口塌陷（仅左右冻结）',
    anchor:
      '\t\t} else if (isTopView === false && isLeftView === true) {\n\t\t\tconst leftGap = endSheetView.startX - startSheetView.startX;',
    insert:
      '\t\t} else if (isTopView === false && isLeftView === true) {\n\t\t\t// [patch-univer] 冻结区超宽时保留主视口最小宽度，避免滚动失效\n\t\t\tconst leftGap = Math.max(0, Math.min(endSheetView.startX - startSheetView.startX, viewMain.scene.getParent().width / (Math.abs(viewMain.scene.scaleX) || 1) - rowHeaderWidthAndMarginLeft - 80));',
  },
  {
    pkg: 'sheets-ui',
    file: 'lib/es/index.js',
    name: '冻结行超出可视高度时主视口塌陷（仅上下冻结）',
    anchor:
      '\t\t} else if (isTopView === true && isLeftView === false) {\n\t\t\tconst topGap = endSheetView.startY - startSheetView.startY;',
    insert:
      '\t\t} else if (isTopView === true && isLeftView === false) {\n\t\t\t// [patch-univer] 冻结区超高时保留主视口最小高度，避免滚动失效\n\t\t\tconst topGap = Math.max(0, Math.min(endSheetView.startY - startSheetView.startY, viewMain.scene.getParent().height / (Math.abs(viewMain.scene.scaleY) || 1) - columnHeaderHeightAndMarginTop - 80));',
  },
  {
    pkg: 'sheets-ui',
    file: 'lib/es/index.js',
    name: '冻结行列同时超出可视尺寸时主视口塌陷',
    anchor:
      '\t\t\tconst leftGap = endSheetView.startX - startSheetView.startX;\n\t\t\tconst topGap = endSheetView.startY - startSheetView.startY;',
    insert:
      '\t\t\t// [patch-univer] 冻结区超宽/超高时保留主视口最小尺寸，避免滚动失效\n\t\t\tconst leftGap = Math.max(0, Math.min(endSheetView.startX - startSheetView.startX, viewMain.scene.getParent().width / (Math.abs(viewMain.scene.scaleX) || 1) - rowHeaderWidthAndMarginLeft - 80));\n\t\t\tconst topGap = Math.max(0, Math.min(endSheetView.startY - startSheetView.startY, viewMain.scene.getParent().height / (Math.abs(viewMain.scene.scaleY) || 1) - columnHeaderHeightAndMarginTop - 80));',
  },
];

// 同一文件多补丁只在读一次内存里链式应用；cjs / es / lib/index 内容一致，一并处理
const targets = new Map();
for (const patch of PATCHES) {
  const files = patch.files ?? [patch.file, patch.file.replace('/es/', '/cjs/')];
  // drawing-ui 另有 lib/index.js 入口（部分打包路径会走这里，补丁漏掉会导致双击仍弹出预览）
  if (patch.pkg === 'drawing-ui') files.push('lib/index.js');
  for (const file of files) {
    const key = `${patch.pkg}::${file}`;
    if (!targets.has(key)) targets.set(key, { pkg: patch.pkg, file, patches: [] });
    targets.get(key).patches.push(patch);
  }
}

let applied = 0;
let already = 0;
let skipped = 0;

for (const { pkg, file, patches } of targets.values()) {
  const pkgRoot = resolvePkgRoot(pkg);
  const filePath = pkgRoot ? path.join(pkgRoot, file) : null;
  if (!filePath || !fs.existsSync(filePath)) {
    console.warn(`[patch-univer] 跳过（文件不存在）: ${pkg}/${file}`);
    skipped += patches.length;
    continue;
  }
  let src = fs.readFileSync(filePath, 'utf8');
  let changed = false;
  for (const patch of patches) {
    if (src.includes(patch.insert)) {
      already += 1;
      continue;
    }
    if (!src.includes(patch.anchor)) {
      console.warn(`[patch-univer] 跳过（锚点不存在，上游代码已变化，请人工确认）: ${pkg}/${file} <- ${patch.name}`);
      skipped += 1;
      continue;
    }
    src = src.replace(patch.anchor, patch.insert);
    changed = true;
    applied += 1;
    console.log(`[patch-univer] 已修补: ${pkg}/${file} <- ${patch.name}`);
  }
  if (!changed) continue;
  // 先写临时文件再原子替换：pnpm 硬链接场景下避免直接改写全局 store 中的文件
  const tmpPath = `${filePath}.patched`;
  fs.writeFileSync(tmpPath, src);
  fs.renameSync(tmpPath, filePath);
}

console.log(`[patch-univer] 完成：新修补 ${applied}，已修补 ${already}，跳过 ${skipped}`);

// 补丁后清除 Vite deps 预构建目录（只删 Univer 相关文件会导致 504 Outdated Optimize Dep）。
// 覆盖本包与 monorepo 内 demo 两个常见位置；外部宿主工程由 Vite 自行检测依赖变化。
const viteDepsDirs = [
  path.join(here, '..', 'node_modules', '.vite', 'deps'),
  path.join(here, '..', '..', '..', 'apps', 'demo', 'node_modules', '.vite', 'deps'),
];
for (const viteDepsDir of viteDepsDirs) {
  if (fs.existsSync(viteDepsDir)) {
    fs.rmSync(viteDepsDir, { recursive: true, force: true });
    console.log(`[patch-univer] 已清除 Vite 预构建缓存: ${viteDepsDir}（请重启 dev server）`);
  }
}
