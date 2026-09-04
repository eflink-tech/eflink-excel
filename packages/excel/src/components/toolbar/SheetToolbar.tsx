import { Fragment, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { ChangeEvent, ReactNode } from 'react';
import { ChevronRight, PenLine } from 'lucide-react';
import * as cmd from '../../core/univer/commands';
import type { BorderStyleKey, BorderTypeKey } from '../../core/univer/commands';
import { getUniverApi } from '../../core/univer/controller';
import { importFileAction } from '../../core/fileActions';
import { useUiStore } from '../../store/uiStore';
import { ColorPicker } from '../color';
import { ChartPickerPanel } from './ChartPickerPanel';
import { MenuDropdown } from './MenuDropdown';
import { ToolbarAlignMenuItem, ToolbarDropdownPanel, ToolbarMenuDivider, ToolbarMenuItem } from './ToolbarMenuPanel';
import { closeAllToolbarDropdowns } from './toolbarDropdown';
import { UNIVER_READY_EVENT } from './FormulaBar';
import './myfsheet.css';

/** myf-sheet-demo 同款单行工具栏：按钮顺序 / 分组 / 图标 / 下拉面板完全按 demo，操作走 Univer 命令 */

interface SelState {
  ff?: string; // 字体
  fs?: number; // 字号
  bl: boolean; // 加粗
  it: boolean; // 倾斜
  u: boolean; // 下划线
  st: boolean; // 删除线
  ht?: string; // 水平对齐
  vt?: string; // 垂直对齐
  cl: string | null; // 字体颜色
  bg: string | null; // 填充颜色
  numfmt: string; // 数字格式 pattern
}

const INITIAL_SEL: SelState = { bl: false, it: false, u: false, st: false, cl: null, bg: null, numfmt: '' };

const FONT_GROUPS: string[][] = [
  ['默认字体', '宋体', '黑体', '楷体', '仿宋', '华文行楷', '华文隶书'],
  ['Arial', 'Tahoma', 'Verdana'],
];
const FONT_SIZES = [10, 11, 12, 14, 16, 18, 20, 24];

/** 字体色 / 填充色「最近使用」的 localStorage key（与 eflink-word 的 recentColors 约定一致） */
const RECENT_COLOR_KEYS = {
  fontColor: 'eflink-excel:recent-colors:fontColor',
  fillColor: 'eflink-excel:recent-colors:fillColor',
};

interface NumFormatItem {
  label: string;
  pattern: string;
  sample?: string;
}

/** 「常规」面板：分组顺序与示例文字完全按 demo；pattern 走 Univer numfmt */
const NUM_FORMAT_GROUPS: NumFormatItem[][] = [
  [
    { label: '常规', pattern: 'General' },
    { label: '文本', pattern: '@' },
  ],
  [
    { label: '数值', pattern: '#,##0.00', sample: '0.95' },
    { label: '百分比', pattern: '0.00%', sample: '95.00%' },
    { label: '分数', pattern: '# ?/?', sample: '1/2' },
    { label: '科学计数', pattern: '0.00E+00', sample: '9.5E-01' },
  ],
  [
    { label: '人民币', pattern: '"¥"#,##0.00', sample: '￥8.80' },
    { label: '港币', pattern: '"HK$"#,##0.00', sample: 'HK$8.80' },
    { label: '美元', pattern: '"$"#,##0.00', sample: '$8.80' },
  ],
  [
    { label: '日期', pattern: 'yyyy-MM-dd', sample: '2025-01-01' },
    { label: '年月', pattern: 'yyyy"年"M"月"', sample: '2025年1月' },
    { label: '月日', pattern: 'M"月"d"日"', sample: '1月15日' },
    { label: '时间', pattern: 'HH:mm', sample: '15:00' },
    { label: '日期时间', pattern: 'yyyy-MM-dd HH:mm', sample: '2025-01-01 15:00' },
  ],
];
const NUM_FORMATS = NUM_FORMAT_GROUPS.flat();

const H_ALIGN_ITEMS: [value: 'left' | 'center' | 'right', label: string, icon: string][] = [
  ['left', '左对齐', 'icon-paragraph-align-left'],
  ['center', '居中对齐', 'icon-paragraph-align-center'],
  ['right', '右对齐', 'myf-icon-paragraph-align-right'],
];
const V_ALIGN_ITEMS: [value: 'top' | 'middle' | 'bottom', label: string, icon: string][] = [
  ['top', '顶端对齐', 'myf-icon-paragraph-top'],
  ['middle', '居中对齐', 'myf-icon-paragraph-middle'],
  ['bottom', '底端对齐', 'myf-icon-paragraph-bottom'],
];

type PanelId =
  | 'insert' | 'convention' | 'fontFamily' | 'fontSize'
  | 'fontColor' | 'fillColor' | 'border'
  | 'hAlign' | 'vAlign' | 'freeze' | 'sum' | 'sort' | 'comment';

/** 「更多」按钮预留宽度（溢出收纳计算用） */
const MORE_RESERVE = 44;

export function SheetToolbar() {
  const showToast = useUiStore((s) => s.showToast);
  const [openPanel, setOpenPanel] = useState<PanelId | null>(null);
  const [moreOpen, setMoreOpen] = useState(false);
  const [chartPickerSource, setChartPickerSource] = useState<'insert' | 'chart' | null>(null);
  const [cut, setCut] = useState(Number.MAX_SAFE_INTEGER);
  const [sel, setSel] = useState<SelState>(INITIAL_SEL);
  const [borderStyle, setBorderStyle] = useState<BorderStyleKey>('thin');
  const [borderColor, setBorderColor] = useState('#000000');
  const importInputRef = useRef<HTMLInputElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  const wrapRef = useRef<HTMLDivElement>(null);
  const measureRef = useRef<HTMLUListElement>(null);
  const insertBtnRef = useRef<HTMLAnchorElement>(null);

  // 订阅选区变化：同步当前样式（字体/字号/加粗/颜色/对齐/数字格式）到按钮状态
  const subscribeDisposable = useRef<{ dispose(): void } | null>(null);
  const syncSelRef = useRef<() => void>(() => {});
  useEffect(() => {
    const syncFromSelection = () => {
      const range = getActiveRange();
      if (!range) return;
      const st = range.getCellStyleData();
      setSel({
        ff: typeof st?.ff === 'string' ? st.ff : undefined,
        fs: typeof st?.fs === 'number' ? st.fs : undefined,
        bl: st?.bl === 1,
        it: st?.it === 1,
        u: st?.ul?.s === 1,
        st: st?.st?.s === 1,
        ht: typeof st?.ht === 'string' ? st.ht : undefined,
        vt: typeof st?.vt === 'string' ? st.vt : undefined,
        cl: st?.cl?.rgb ?? null,
        bg: st?.bg?.rgb ?? null,
        numfmt: cmd.getActiveNumberFormat(),
      });
    };
    syncSelRef.current = syncFromSelection;
    const subscribe = () => {
      const api = getUniverApi();
      if (!api) return;
      subscribeDisposable.current?.dispose();
      subscribeDisposable.current = api.addEvent(api.Event.SelectionChanged, syncFromSelection);
      syncFromSelection();
    };
    subscribe();
    window.addEventListener(UNIVER_READY_EVENT, subscribe);
    return () => {
      window.removeEventListener(UNIVER_READY_EVENT, subscribe);
      subscribeDisposable.current?.dispose();
      subscribeDisposable.current = null;
    };
  }, []);

  // 点击工具栏外部时收起面板与「更多」浮层（capture：Univer 画布会拦截冒泡事件）
  useEffect(() => {
    if (openPanel == null && !moreOpen && chartPickerSource == null) return;
    const onDown = (e: PointerEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) {
        setOpenPanel(null);
        setMoreOpen(false);
        setChartPickerSource(null);
      }
    };
    document.addEventListener('pointerdown', onDown, true);
    return () => document.removeEventListener('pointerdown', onDown, true);
  }, [openPanel, moreOpen, chartPickerSource]);

  // 溢出收纳：以测量行宽度计算能放下多少个工具项，其余收进「更多」浮层（demo handlerToolsResize 同款行为）
  const computeCutRef = useRef(() => {});
  computeCutRef.current = () => {
    const wrap = wrapRef.current;
    const measure = measureRef.current;
    if (!wrap || !measure) return;
    const avail = wrap.clientWidth - MORE_RESERVE;
    let acc = 0;
    let next = measure.children.length;
    for (let i = 0; i < measure.children.length; i++) {
      acc += (measure.children[i] as HTMLElement).getBoundingClientRect().width;
      if (acc > avail) {
        next = i;
        break;
      }
    }
    setCut((prev) => (prev === next ? prev : next));
  };
  useLayoutEffect(() => {
    computeCutRef.current();
  });
  useEffect(() => {
    const wrap = wrapRef.current;
    if (!wrap) return;
    const ro = new ResizeObserver(() => computeCutRef.current());
    ro.observe(wrap);
    // 测量行也要监听：dev 下 CSS 异步注入，首帧量出的行宽偏大；
    // 样式就位后测量行尺寸突变会触发重算（字体加载后同理）
    if (measureRef.current) ro.observe(measureRef.current);
    return () => ro.disconnect();
  }, []);

  const toggle = (id: PanelId) => () => {
    closeAllToolbarDropdowns();
    setOpenPanel((p) => {
      const next = p === id ? null : id;
      if (next) {
        setMoreOpen(false);
        setChartPickerSource(null);
      }
      return next;
    });
  };
  // 设置格式（尤其数字格式）不一定会触发 SelectionChanged，命令执行后主动刷新一次按钮状态
  const refreshSoon = () => window.setTimeout(() => syncSelRef.current(), 60);
  const run = (fn: () => void) => () => {
    fn();
    refreshSoon();
    setOpenPanel(null);
  };
  const dev = () => showToast('功能开发中');

  const runComment = (fn: () => void) => () => {
    const range = getActiveRange();
    if (!range) {
      showToast('请先选择单元格');
      return;
    }
    fn();
    setOpenPanel(null);
  };

  const pickChart = (typeId: string) => {
    setChartPickerSource(null);
    void cmd
      .insertChart(typeId)
      .then((ok) => {
        if (!ok) showToast('插入图表失败');
      })
      .catch(() => showToast('插入图表失败'));
  };

  const openChartPicker = (source: 'insert' | 'chart') => {
    closeAllToolbarDropdowns();
    setOpenPanel(null);
    setMoreOpen(false);
    setChartPickerSource(source);
  };

  const closeToolbarPanels = () => {
    setOpenPanel(null);
    setMoreOpen(false);
    setChartPickerSource(null);
  };

  // 插入图片：模式写入 input 的 data-mode，选中文件后按该模式插入
  const openImagePicker = (mode: 'cell' | 'float') => () => {
    const input = imageInputRef.current;
    if (!input) return;
    input.dataset.mode = mode;
    setOpenPanel(null);
    input.click();
  };
  const onImagePicked = (e: ChangeEvent<HTMLInputElement>) => {
    const input = e.currentTarget;
    const file = input.files?.[0];
    if (file) {
      const mode = input.dataset.mode === 'float' ? 'float' : 'cell';
      const pick = mode === 'cell' ? cmd.insertCellImage(file) : cmd.insertFloatingImage(file);
      void pick
        .then((ok) => {
          if (!ok) showToast('插入图片失败');
        })
        .catch(() => showToast('插入图片失败'));
    }
    input.value = '';
  };

  // 构建：当前样式派生值
  const hCur = (sel.ht && H_ALIGN_ITEMS.some(([v]) => v === sel.ht) ? sel.ht : 'left') as 'left' | 'center' | 'right';
  const vCur = (sel.vt && V_ALIGN_ITEMS.some(([v]) => v === sel.vt) ? sel.vt : 'top') as 'top' | 'middle' | 'bottom';
  const hIcon = H_ALIGN_ITEMS.find(([v]) => v === hCur)?.[2] ?? 'icon-paragraph-align-left';
  const vIcon = V_ALIGN_ITEMS.find(([v]) => v === vCur)?.[2] ?? 'myf-icon-paragraph-top';
  const fontLabel = sel.ff || '默认字体';
  const fontSizeLabel = sel.fs != null ? String(sel.fs) : '12';
  const numfmtLabel =
    NUM_FORMATS.find((f) => f.pattern === sel.numfmt || (f.label === '常规' && (sel.numfmt === '' || sel.numfmt === 'General')))
      ?.label ?? '常规';

  // 构建工具项。measure=true 时用于隐藏测量行：不渲染面板内容（绝对定位面板不影响测宽），
  // 避免同一份面板 DOM 出现两次。
  const buildItems = (measure: boolean): ReactNode[] => {
    const open = (id: PanelId) => !measure && openPanel === id;
    const chartPicker = (source: 'insert' | 'chart') =>
      !measure && chartPickerSource === source ? (
        <ChartPickerPanel onPick={pickChart} onClose={() => setChartPickerSource(null)} />
      ) : null;
    const items: ReactNode[] = [];
    const pushSplit = (key: string) =>
      items.push(
        <li key={key}>
          <div className="myf-split">
            <span className="myf-split-line" />
          </div>
        </li>,
      );

    if (!measure) {
      items.push(
        <li key="menu">
          <MenuDropdown
            variant="sheet"
            importInputRef={importInputRef}
            onOpen={closeToolbarPanels}
            trigger={
              <a
                role="button"
                tabIndex={0}
                className="myf-tool-panel"
                aria-label="菜单"
                title="菜单"
              >
                <span className="myf-icon myf-icon-file-menu" />
                <span className="myf-icon-text">菜单</span>
                <div className="myf-icon myf-icon-arrow-down" />
              </a>
            }
          />
        </li>,
      );
    } else {
      items.push(
        <li key="menu">
          <a
            role="button"
            tabIndex={-1}
            className="myf-tool-panel"
            aria-hidden="true"
          >
            <span className="myf-icon myf-icon-file-menu" />
            <span className="myf-icon-text">菜单</span>
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </li>,
      );
    }
    pushSplit('sp-menu');

    items.push(
      <li key="undo">
        <a role="button" tabIndex={0} title="撤销" aria-label="撤销" onClick={run(cmd.undo)}>
          <span className="myf-icon myf-icon-undo" />
        </a>
      </li>,
      <li key="redo">
        <a role="button" tabIndex={0} title="重做" aria-label="重做" onClick={run(cmd.redo)}>
          <span className="myf-icon myf-icon-redo" />
        </a>
      </li>,
    );
    pushSplit('sp1');
    items.push(
      <li key="formatPainter">
        <a role="button" tabIndex={0} title="格式刷" aria-label="格式刷" onClick={run(cmd.formatPainter)}>
          <span className="myf-icon myf-icon-match-prop" />
        </a>
      </li>,
      <li key="clearFormat">
        <a role="button" tabIndex={0} title="清除格式" aria-label="清除格式" onClick={run(cmd.clearFormatOnly)}>
          <span className="myf-icon myf-icon-clear-format" />
        </a>
      </li>,
    );
    pushSplit('sp2');

    // 插入面板
    items.push(
      <li key="insert" className={chartPickerSource === 'insert' ? 'active' : undefined}>
        <a
          ref={insertBtnRef}
          role="button"
          tabIndex={0}
          title="插入"
          aria-label="插入"
          className={`myf-tool-panel${openPanel === 'insert' || chartPickerSource === 'insert' ? ' active' : ''}`}
          onClick={toggle('insert')}
        >
          <span className="myf-icon myf-icon-add" />
          <span className="myf-icon-text">插入</span>
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('insert') && (
          <ToolbarDropdownPanel style={{ width: 220 }}>
            <ToolbarMenuItem iconClass="icon-add-img-for-cell" label="单元格图片" onClick={openImagePicker('cell')} />
            <ToolbarMenuItem iconClass="icon-add-img-for-fix" label="浮动图片" onClick={openImagePicker('float')} />
            <ToolbarMenuItem iconClass="myf-icon-add-chart" label="图表" onClick={() => openChartPicker('insert')} />
            <ToolbarMenuItem iconClass="icon-add-link" label="链接" onClick={run(cmd.insertHyperLink)} />
            <ToolbarMenuDivider />
            <ToolbarMenuItem iconClass="myf-icon-insert-top-row" label="在上方插入行" onClick={run(cmd.insertRowAbove)} />
            <ToolbarMenuItem iconClass="myf-icon-insert-bottom-row" label="在下方插入行" onClick={run(cmd.insertRowBelow)} />
            <ToolbarMenuItem iconClass="myf-icon-insert-left-col" label="在左侧插入列" onClick={run(cmd.insertColLeft)} />
            <ToolbarMenuItem iconClass="myf-icon-insert-right-col" label="在右侧插入列" onClick={run(cmd.insertColRight)} />
            <ToolbarMenuDivider />
            <ToolbarMenuItem iconClass="icon-insert-doc-link-ent" label="文档" onClick={dev} />
            <ToolbarMenuItem iconClass="icon-insert-file-link" label="文件" onClick={dev} />
            <ToolbarMenuItem iconClass="icon-add-todo" label="代办" onClick={dev} />
          </ToolbarDropdownPanel>
        )}
        {chartPicker('insert')}
      </li>,
    );
    pushSplit('sp3');

    // 数字格式（常规）面板
    items.push(
      <li key="convention">
        <a
          role="button"
          tabIndex={0}
          title="常规"
          aria-label="常规"
          className={`myf-tool-panel${openPanel === 'convention' ? ' active' : ''}`}
          onClick={toggle('convention')}
        >
          <span className="myf-icon-text">{numfmtLabel}</span>
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('convention') && (
          <ToolbarDropdownPanel style={{ width: 220 }}>
            {NUM_FORMAT_GROUPS.map((group, gi) => (
              <Fragment key={gi}>
                {gi > 0 && <ToolbarMenuDivider />}
                {group.map((f) => (
                  <ToolbarMenuItem key={f.label} label={f.label} right={f.sample} onClick={run(() => cmd.setNumberFormatPattern(f.pattern))} />
                ))}
              </Fragment>
            ))}
            <ToolbarMenuDivider />
            <ToolbarMenuItem label="更多格式" onClick={dev} />
          </ToolbarDropdownPanel>
        )}
      </li>,
    );

    // 小数位数（.0 上下箭头）
    items.push(
      <li key="decimal">
        <div className="myf-arrows-container" title="小数位数">
          <span className="myf-icon-text">.0</span>
          <div className="myf-right-icons">
            <a role="button" tabIndex={0} className="myf-arrow-icon" title="增加小数位数" onClick={run(cmd.addDecimal)}>
              <span className="myf-icon myf-icons-arrow-up" />
            </a>
            <a role="button" tabIndex={0} className="myf-arrow-icon" title="减少小数位数" onClick={run(cmd.subtractDecimal)}>
              <span className="myf-icon myf-icons-arrow-down" />
            </a>
          </div>
        </div>
      </li>,
    );
    pushSplit('sp4');

    // 字体 / 字号面板
    items.push(
      <li key="fontFamily">
        <a
          role="button"
          tabIndex={0}
          title="字体"
          aria-label="字体"
          className={`myf-tool-panel${openPanel === 'fontFamily' ? ' active' : ''}`}
          onClick={toggle('fontFamily')}
        >
          <span className="myf-icon-text">{fontLabel}</span>
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('fontFamily') && (
          <ToolbarDropdownPanel style={{ width: 160 }}>
            {FONT_GROUPS.map((group, gi) => (
              <Fragment key={gi}>
                {gi > 0 && <ToolbarMenuDivider />}
                {group.map((f) => {
                  const selected = f === '默认字体' ? !sel.ff : sel.ff === f;
                  return (
                    <ToolbarMenuItem
                      key={f}
                      iconClass={selected ? 'myf-icon-selected' : undefined}
                      label={f}
                      onClick={run(() => cmd.setFontFamily(f === '默认字体' ? 'Arial' : f))}
                    />
                  );
                })}
              </Fragment>
            ))}
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="fontSize">
        <a
          role="button"
          tabIndex={0}
          title="字号"
          aria-label="字号"
          className={`myf-tool-panel${openPanel === 'fontSize' ? ' active' : ''}`}
          onClick={toggle('fontSize')}
        >
          <span className="myf-icon-text">{fontSizeLabel}</span>
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('fontSize') && (
          <ToolbarDropdownPanel style={{ width: 100 }}>
            {FONT_SIZES.map((s) => (
              <ToolbarMenuItem
                key={s}
                iconClass={(sel.fs ?? 12) === s ? 'myf-icon-selected' : undefined}
                label={String(s)}
                onClick={run(() => cmd.setFontSize(s))}
              />
            ))}
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="bold">
        <a role="button" tabIndex={0} title="加粗" aria-label="加粗" className={sel.bl ? 'active' : ''} onClick={run(cmd.toggleBold)}>
          <span className="myf-icon myf-icon-bold" />
        </a>
      </li>,
      <li key="italic">
        <a role="button" tabIndex={0} title="倾斜" aria-label="倾斜" className={sel.it ? 'active' : ''} onClick={run(cmd.toggleItalic)}>
          <span className="myf-icon myf-icon-italic" />
        </a>
      </li>,
      <li key="underline">
        <a role="button" tabIndex={0} title="下划线" aria-label="下划线" className={sel.u ? 'active' : ''} onClick={run(cmd.toggleUnderline)}>
          <span className="myf-icon myf-icon-underline" />
        </a>
      </li>,
      <li key="strike">
        <a role="button" tabIndex={0} title="删除线" aria-label="删除线" className={sel.st ? 'active' : ''} onClick={run(cmd.toggleStrikeThrough)}>
          <span className="myf-icon myf-icon-delete" />
        </a>
      </li>,
    );

    // 字体颜色 / 填充颜色 / 边框（分体式：图标直接生效，箭头开面板）
    items.push(
      <li className="right-icon" key="fontColor">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} className="arrows-left" title="字体颜色" onClick={run(() => cmd.setFontColor(sel.cl ?? '#000000'))}>
            <div className="myf-icon myf-icon-fontcolor" />
            {/* 未选色时底线显示白色（空色状态），选色后跟随所选颜色 */}
            <div className="myf-icon-bottom-color" style={{ backgroundColor: sel.cl ?? '#ffffff' }} />
          </a>
          <a role="button" tabIndex={0} className="myf-pr-2" title="选择字体颜色" onClick={toggle('fontColor')}>
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('fontColor') && (
          <div className="absolute top-full left-0 z-[99] mt-[2px] text-left">
            {/* eflink-word 同款颜色选择器：主题色/标准色/最近使用 + 「更多颜色」高级取色 */}
            <ColorPicker
              open
              onOpenChange={() => {}}
              value={sel.cl}
              recentStorageKey={RECENT_COLOR_KEYS.fontColor}
              onSelect={(color, options) => {
                cmd.setFontColor(color);
                refreshSoon();
                if (options?.close !== false) setOpenPanel(null);
              }}
              onDefault={() => {
                cmd.resetFontColor();
                refreshSoon();
                setOpenPanel(null);
              }}
            />
          </div>
        )}
      </li>,
      <li className="right-icon" key="fillColor">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} className="arrows-left" title="填充颜色" onClick={run(() => cmd.setBackgroundColor(sel.bg ?? '#ffffff'))}>
            <div className="myf-icon myf-icon-fill-color" />
            <div className="myf-icon-bottom-color" style={{ backgroundColor: sel.bg ?? '#ffffff' }} />
          </a>
          <a role="button" tabIndex={0} className="myf-pr-2" title="选择填充颜色" onClick={toggle('fillColor')}>
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('fillColor') && (
          <div className="absolute top-full left-0 z-[99] mt-[2px] text-left">
            <ColorPicker
              open
              onOpenChange={() => {}}
              value={sel.bg}
              recentStorageKey={RECENT_COLOR_KEYS.fillColor}
              onSelect={(color, options) => {
                cmd.setBackgroundColor(color);
                refreshSoon();
                if (options?.close !== false) setOpenPanel(null);
              }}
              onDefault={() => {
                cmd.resetBackgroundColor();
                refreshSoon();
                setOpenPanel(null);
              }}
            />
          </div>
        )}
      </li>,
      <li className="right-icon" key="border">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} title="边框" onClick={run(() => cmd.applyBorder('all', borderStyle, borderColor))}>
            <span className="myf-icon myf-icon-cell-border-all" />
          </a>
          <a role="button" tabIndex={0} className="myf-pr-2" title="边框样式" onClick={toggle('border')}>
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('border') && (
          <div className="myf-x-panel myf-x-visible" style={{ width: 232 }}>
            <div className="myf-border-grid">
              {BORDER_PRESETS.map((preset) => (
                <button key={preset.key} type="button" title={preset.title} onClick={() => cmd.applyBorder(preset.key, borderStyle, borderColor)}>
                  <BorderGlyph edges={preset.edges} diag={preset.diag} />
                </button>
              ))}
            </div>
            <div className="x-split" />
            <BorderSub label="边框颜色" icon={<PenLine size={13} />}>
              <div className="myf-palette" style={{ gridTemplateColumns: 'repeat(5, 20px)' }}>
                {BORDER_COLORS.map((c) => (
                  <button
                    key={c}
                    type="button"
                    title={c}
                    style={{ background: c, width: 20, height: 20 }}
                    onClick={() => setBorderColor(c)}
                  />
                ))}
              </div>
            </BorderSub>
            <BorderSub label="边框样式" icon={<DashLineGlyph />}>
              {BORDER_STYLES.map((s) => (
                <button
                  key={s.key}
                  type="button"
                  className={`myf-border-style-row${borderStyle === s.key ? ' on' : ''}`}
                  onClick={() => setBorderStyle(s.key)}
                >
                  <span>{s.label}</span>
                  <svg width="34" height="8" viewBox="0 0 34 8">
                    <path d="M1 4 H33" stroke={borderStyle === s.key ? '#1d4ed8' : '#475569'} strokeWidth={s.width} strokeDasharray={s.dasharray} />
                  </svg>
                </button>
              ))}
            </BorderSub>
            <button type="button" className="myf-x-plain" onClick={dev}>
              更多边框设置
            </button>
          </div>
        )}
      </li>,
    );
    pushSplit('sp5');

    // 水平 / 垂直对齐（图标显示当前对齐，箭头开面板）
    items.push(
      <li className="right-icon" key="hAlign">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} title="水平对齐" onClick={run(() => cmd.setHorizontalAlign(hCur))}>
            <span className={`myf-icon ${hIcon}`} />
          </a>
          <a
            role="button"
            tabIndex={0}
            className={`myf-tool-panel myf-pr-2${openPanel === 'hAlign' ? ' active' : ''}`}
            title="对齐方式"
            onClick={toggle('hAlign')}
          >
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('hAlign') && (
          <ToolbarDropdownPanel style={{ width: 180 }}>
            {H_ALIGN_ITEMS.map(([v, label, icon]) => (
              <ToolbarAlignMenuItem
                key={v}
                label={label}
                selected={hCur === v}
                alignIcon={icon}
                onClick={run(() => cmd.setHorizontalAlign(v))}
              />
            ))}
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li className="right-icon" key="vAlign">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} title="垂直对齐" onClick={run(() => cmd.setVerticalAlign(vCur))}>
            <span className={`myf-icon ${vIcon}`} />
          </a>
          <a
            role="button"
            tabIndex={0}
            className={`myf-tool-panel myf-pr-2${openPanel === 'vAlign' ? ' active' : ''}`}
            title="对齐方式"
            onClick={toggle('vAlign')}
          >
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('vAlign') && (
          <ToolbarDropdownPanel style={{ width: 180 }}>
            {V_ALIGN_ITEMS.map(([v, label, icon]) => (
              <ToolbarAlignMenuItem
                key={v}
                label={label}
                selected={vCur === v}
                alignIcon={icon}
                onClick={run(() => cmd.setVerticalAlign(v))}
              />
            ))}
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="merge">
        <a role="button" tabIndex={0} title="合并单元格" aria-label="合并单元格" onClick={run(cmd.toggleMergeCells)}>
          <span className="myf-icon myf-icon-merge-cells" />
        </a>
      </li>,
      <li key="freeze">
        <a
          role="button"
          tabIndex={0}
          title="冻结"
          aria-label="冻结"
          className={`myf-tool-panel${openPanel === 'freeze' ? ' active' : ''}`}
          onClick={toggle('freeze')}
        >
          <span className="myf-icon myf-icon-freeze" />
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('freeze') && (
          <ToolbarDropdownPanel style={{ width: 160 }}>
            <ToolbarMenuItem iconClass="myf-icon-freeze" label="冻结首行" onClick={run(cmd.freezeFirstRow)} />
            <ToolbarMenuItem iconClass="myf-icon-freeze" label="冻结首列" onClick={run(cmd.freezeFirstColumn)} />
            <ToolbarMenuItem iconClass="myf-icon-freeze" label="冻结至当前单元格" onClick={run(cmd.freezeToActiveCell)} />
            <ToolbarMenuItem iconClass="myf-icon-freeze" label="取消冻结" onClick={run(cmd.cancelFreeze)} />
          </ToolbarDropdownPanel>
        )}
      </li>,
    );
    pushSplit('sp6');

    // Σ 快速统计 / 图表 / 筛选 / 排序 / @人 / 评论 / 查找替换
    items.push(
      <li className="right-icon" key="sum">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} title="自动计算选框区域内的和" onClick={run(() => cmd.applyQuickStat('SUM'))}>
            <span className="myf-icon myf-icon-add-function" />
          </a>
          <a
            role="button"
            tabIndex={0}
            className={`myf-tool-panel myf-pr-2${openPanel === 'sum' ? ' active' : ''}`}
            title="统计"
            onClick={toggle('sum')}
          >
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('sum') && (
          <ToolbarDropdownPanel style={{ width: 140 }}>
            <ToolbarMenuItem label="求和" onClick={run(() => cmd.applyQuickStat('SUM'))} />
            <ToolbarMenuItem label="平均值" onClick={run(() => cmd.applyQuickStat('AVERAGE'))} />
            <ToolbarMenuItem label="计数" onClick={run(() => cmd.applyQuickStat('COUNT'))} />
            <ToolbarMenuItem label="最大值" onClick={run(() => cmd.applyQuickStat('MAX'))} />
            <ToolbarMenuItem label="最小值" onClick={run(() => cmd.applyQuickStat('MIN'))} />
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="chart" className={chartPickerSource === 'chart' ? 'active' : undefined}>
        <a
          role="button"
          tabIndex={0}
          title="图表"
          aria-label="图表"
          className={`myf-tool-panel${chartPickerSource === 'chart' ? ' active' : ''}`}
          onClick={() => openChartPicker('chart')}
        >
          <span className="myf-icon myf-icon-add-chart" />
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {chartPicker('chart')}
      </li>,
      <li className="right-icon" key="filter">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} title="筛选" onClick={run(cmd.toggleFilter)}>
            <span className="myf-icon myf-icon-filter" />
          </a>
          <a role="button" tabIndex={0} className="myf-pr-2" title="筛选" onClick={run(cmd.toggleFilter)}>
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
      </li>,
      <li key="sort">
        <a
          role="button"
          tabIndex={0}
          title="排序"
          aria-label="排序"
          className={`myf-tool-panel${openPanel === 'sort' ? ' active' : ''}`}
          onClick={toggle('sort')}
        >
          <span className="myf-icon myf-icon-sort" />
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('sort') && (
          <ToolbarDropdownPanel style={{ width: 120 }}>
            <ToolbarMenuItem iconClass="myf-icon-sort" label="升序" onClick={run(cmd.sortAscending)} />
            <ToolbarMenuItem iconClass="myf-icon-sort" label="降序" onClick={run(cmd.sortDescending)} />
            <ToolbarMenuItem iconClass="myf-icon-sort" label="自定义排序" onClick={run(cmd.sortCustom)} />
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="mention">
        <a role="button" tabIndex={0} title="@人" aria-label="@人" onClick={dev}>
          <span className="myf-icon myf-icon-mention-ent" />
        </a>
      </li>,
      <li className="right-icon" key="comment">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} title="添加评论" aria-label="添加评论" onClick={runComment(cmd.addComment)}>
            <span className="myf-icon myf-icon-comment-ent" />
          </a>
          <a
            role="button"
            tabIndex={0}
            className={`myf-tool-panel myf-pr-2${openPanel === 'comment' ? ' active' : ''}`}
            title="评论"
            aria-label="评论菜单"
            onClick={toggle('comment')}
          >
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('comment') && (
          <ToolbarDropdownPanel style={{ width: 140 }}>
            <ToolbarMenuItem iconClass="myf-icon-comment-ent" label="添加评论" onClick={runComment(cmd.addComment)} />
            <ToolbarMenuItem iconClass="myf-icon-comment-ent" label="评论管理" onClick={run(cmd.toggleCommentPanel)} />
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="findReplace">
        <a role="button" tabIndex={0} title="查找替换" aria-label="查找替换" onClick={run(cmd.openFindReplace)}>
          <span className="myf-icon myf-icon-find-replace" />
        </a>
      </li>,
    );

    return items;
  };

  const liveItems = buildItems(false);
  const measureItems = buildItems(true);
  const overflowed = cut < liveItems.length;

  return (
    <div ref={rootRef} className="relative z-30 min-w-0 flex-1 bg-white">
      <input
        ref={importInputRef}
        type="file"
        accept=".efexcel,application/json"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) void importFileAction(file);
          e.target.value = '';
        }}
      />
      <input ref={imageInputRef} type="file" accept="image/*" className="hidden" onChange={onImagePicked} />
      <div ref={wrapRef} className="myf-sheet-tools">
        <ul className="myf-toolbar">
          {liveItems.slice(0, overflowed ? cut : liveItems.length)}
          {overflowed && (
            <li className="myf-more-button">
              <a role="button" tabIndex={0} title="更多" aria-label="更多" className={moreOpen ? 'active' : ''} onClick={() => { closeAllToolbarDropdowns(); setMoreOpen((v) => !v); setOpenPanel(null); setChartPickerSource(null); }}>
                <span className="myf-icon myf-icon-more" />
              </a>
            </li>
          )}
        </ul>
        {overflowed && moreOpen && (
          <div className="myf-hidden-tools">
            <ul>{liveItems.slice(cut)}</ul>
          </div>
        )}
        {/* 测量行：始终完整渲染，用于计算溢出收纳位置 */}
        <ul ref={measureRef} className="myf-toolbar myf-measure" aria-hidden="true" inert>
          {measureItems}
        </ul>
      </div>
    </div>
  );
}

// ---------- 面板内部构件 ----------

function getActiveRange() {
  return getUniverApi()?.getActiveWorkbook()?.getActiveRange() ?? null;
}

/** 带 hover 二级面板的菜单行（边框颜色 / 边框样式） */
function BorderSub({ label, icon, children }: { label: string; icon: ReactNode; children: ReactNode }) {
  const [openSub, setOpenSub] = useState(false);
  return (
    <div className="myf-border-sub" onMouseEnter={() => setOpenSub(true)} onMouseLeave={() => setOpenSub(false)}>
      <button type="button">
        {icon}
        <span className="flex-1">{label}</span>
        <ChevronRight size={12} className="text-slate-400" />
      </button>
      {openSub && <div className="myf-border-sub-panel">{children}</div>}
    </div>
  );
}

// ---------- 边框数据与图标（沿用原边框面板） ----------

const BORDER_COLORS = ['#000000', '#595959', '#e33e33', '#e6821e', '#d8b100', '#41a62a', '#1e7ee6', '#7c31d6', '#9c1f5f', '#21a366'];

const BORDER_STYLES: { key: BorderStyleKey; label: string; width: number; dasharray?: string }[] = [
  { key: 'thin', label: '细实线', width: 1 },
  { key: 'medium', label: '中粗实线', width: 2.5 },
  { key: 'dashed', label: '虚线', width: 1.5, dasharray: '4 3' },
  { key: 'dotted', label: '点线', width: 1.5, dasharray: '1.5 2.5' },
  { key: 'double', label: '双实线', width: 3, dasharray: '3 1.5 0.8 1.5' },
];

const BORDER_PRESETS: { key: BorderTypeKey; title: string; edges: string[]; diag?: 'tlbr' | 'bltr' }[] = [
  { key: 'horizontal', title: '上下框线', edges: ['t', 'b', 'h'] },
  { key: 'all', title: '所有框线', edges: ['t', 'b', 'l', 'r', 'h', 'v'] },
  { key: 'outside', title: '外侧框线', edges: ['t', 'b', 'l', 'r'] },
  { key: 'none', title: '无框线', edges: [] },
  { key: 'tlbr', title: '斜下框线', edges: [], diag: 'tlbr' },
  { key: 'left', title: '左框线', edges: ['l'] },
  { key: 'top', title: '上框线', edges: ['t'] },
  { key: 'right', title: '右框线', edges: ['r'] },
  { key: 'bottom', title: '下框线', edges: ['b'] },
  { key: 'bltr', title: '斜上框线', edges: [], diag: 'bltr' },
];

/** 预设框线小图标：浅灰底格 + 深色高亮线 */
function BorderGlyph({ edges, diag }: { edges: string[]; diag?: 'tlbr' | 'bltr' }) {
  const dark = '#1f2937';
  const light = '#c9cfda';
  const seg = (d: string, color: string, w: number) => <path d={d} stroke={color} strokeWidth={w} fill="none" />;
  return (
    <svg width="20" height="20" viewBox="0 0 20 20">
      {seg('M2 2 H18 M2 10 H18 M2 18 H18', light, 1)}
      {seg('M2 2 V18 M10 2 V18 M18 2 V18', light, 1)}
      {edges.includes('t') && seg('M2 2 H18', dark, 1.6)}
      {edges.includes('b') && seg('M2 18 H18', dark, 1.6)}
      {edges.includes('l') && seg('M2 2 V18', dark, 1.6)}
      {edges.includes('r') && seg('M18 2 V18', dark, 1.6)}
      {edges.includes('h') && seg('M2 10 H18', dark, 1.6)}
      {edges.includes('v') && seg('M10 2 V18', dark, 1.6)}
      {diag === 'tlbr' && seg('M2 2 L18 18', dark, 1.6)}
      {diag === 'bltr' && seg('M18 2 L2 18', dark, 1.6)}
    </svg>
  );
}

/** 边框样式菜单图标（实线 + 虚线示意） */
function DashLineGlyph() {
  return (
    <svg width="18" height="14" viewBox="0 0 18 14">
      <path d="M2 4 H16" stroke="#1f2937" strokeWidth="1.4" />
      <path d="M2 10 H16" stroke="#1f2937" strokeWidth="1.4" strokeDasharray="2.5 2" />
    </svg>
  );
}
