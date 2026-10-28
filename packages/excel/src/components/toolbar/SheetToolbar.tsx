import { Fragment, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { ChangeEvent, ReactNode } from 'react';
import { ChevronRight, PenLine } from 'lucide-react';
import { useI18n } from '@eflink-tech/office-kit';
import * as cmd from '../../core/univer/commands';
import type { BorderStyleKey, BorderTypeKey } from '../../core/univer/commands';
import { getUniverApi } from '../../core/univer/controller';
import { useUiStore } from '../../store/uiStore';
import { RECENT_COLORS_KEYS as RECENT_COLOR_KEYS } from '../../storage/keys';
import { ColorPicker } from '../color';
import { ChartPickerPanel } from './ChartPickerPanel';
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
  ['__default__', '宋体', '黑体', '楷体', '仿宋', '华文行楷', '华文隶书'],
  ['Arial', 'Tahoma', 'Verdana'],
];
const FONT_SIZES = [10, 11, 12, 14, 16, 18, 20, 24];

interface NumFormatItem {
  label: string;
  pattern: string;
  sample?: string;
}

type PanelId =
  | 'insert' | 'convention' | 'fontFamily' | 'fontSize'
  | 'fontColor' | 'fillColor' | 'border'
  | 'hAlign' | 'vAlign' | 'freeze' | 'sum' | 'sort' | 'comment';

/** 「更多」按钮预留宽度（溢出收纳计算用） */
const MORE_RESERVE = 44;

export function SheetToolbar() {
  const { t, locale } = useI18n();
  const showToast = useUiStore((s) => s.showToast);
  const [openPanel, setOpenPanel] = useState<PanelId | null>(null);
  const [moreOpen, setMoreOpen] = useState(false);
  const [chartPickerSource, setChartPickerSource] = useState<'insert' | 'chart' | null>(null);
  const [cut, setCut] = useState(Number.MAX_SAFE_INTEGER);
  const [sel, setSel] = useState<SelState>(INITIAL_SEL);
  const [borderStyle, setBorderStyle] = useState<BorderStyleKey>('thin');
  const [borderColor, setBorderColor] = useState('#000000');
  const imageInputRef = useRef<HTMLInputElement>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  const wrapRef = useRef<HTMLDivElement>(null);
  const measureRef = useRef<HTMLUListElement>(null);
  const insertBtnRef = useRef<HTMLAnchorElement>(null);

  // 「常规」面板：分组顺序与示例文字按 demo；pattern 走 Univer numfmt（年月/月日按语言换格式）
  const isZh = locale === 'zh-CN';
  const numFormatGroups: NumFormatItem[][] = [
    [
      { label: t('excel.numfmt.general'), pattern: 'General' },
      { label: t('excel.numfmt.text'), pattern: '@' },
    ],
    [
      { label: t('excel.numfmt.number'), pattern: '#,##0.00', sample: '0.95' },
      { label: t('excel.numfmt.percent'), pattern: '0.00%', sample: '95.00%' },
      { label: t('excel.numfmt.fraction'), pattern: '# ?/?', sample: '1/2' },
      { label: t('excel.numfmt.scientific'), pattern: '0.00E+00', sample: '9.5E-01' },
    ],
    [
      { label: t('excel.numfmt.rmb'), pattern: '"¥"#,##0.00', sample: t('excel.numfmt.sample.rmb') },
      { label: t('excel.numfmt.hkd'), pattern: '"HK$"#,##0.00', sample: 'HK$8.80' },
      { label: t('excel.numfmt.usd'), pattern: '"$"#,##0.00', sample: '$8.80' },
    ],
    [
      { label: t('excel.numfmt.date'), pattern: 'yyyy-MM-dd', sample: '2025-01-01' },
      { label: t('excel.numfmt.yearMonth'), pattern: isZh ? 'yyyy"年"M"月"' : 'yyyy-MM', sample: t('excel.numfmt.sample.yearMonth') },
      { label: t('excel.numfmt.monthDay'), pattern: isZh ? 'M"月"d"日"' : locale === 'en' ? 'MM-dd' : 'dd-MM', sample: t('excel.numfmt.sample.monthDay') },
      { label: t('excel.numfmt.time'), pattern: 'HH:mm', sample: '15:00' },
      { label: t('excel.numfmt.dateTime'), pattern: 'yyyy-MM-dd HH:mm', sample: '2025-01-01 15:00' },
    ],
  ];
  const numFormats = numFormatGroups.flat();
  const hAlignItems: [value: 'left' | 'center' | 'right', label: string, icon: string][] = [
    ['left', t('excel.toolbar.alignLeft'), 'icon-paragraph-align-left'],
    ['center', t('excel.toolbar.alignCenter'), 'icon-paragraph-align-center'],
    ['right', t('excel.toolbar.alignRight'), 'myf-icon-paragraph-align-right'],
  ];
  const vAlignItems: [value: 'top' | 'middle' | 'bottom', label: string, icon: string][] = [
    ['top', t('excel.toolbar.alignTop'), 'myf-icon-paragraph-top'],
    ['middle', t('excel.toolbar.alignMiddle'), 'myf-icon-paragraph-middle'],
    ['bottom', t('excel.toolbar.alignBottom'), 'myf-icon-paragraph-bottom'],
  ];

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
  const dev = () => showToast(t('excel.toolbar.developing'));

  const runComment = (fn: () => void) => () => {
    const range = getActiveRange();
    if (!range) {
      showToast(t('excel.toolbar.selectCellFirst'));
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
        if (!ok) showToast(t('excel.toolbar.insertChartFailed'));
      })
      .catch(() => showToast(t('excel.toolbar.insertChartFailed')));
  };

  const openChartPicker = (source: 'insert' | 'chart') => {
    closeAllToolbarDropdowns();
    setOpenPanel(null);
    setMoreOpen(false);
    setChartPickerSource(source);
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
          if (!ok) showToast(t('excel.toolbar.insertImageFailed'));
        })
        .catch(() => showToast(t('excel.toolbar.insertImageFailed')));
    }
    input.value = '';
  };

  // 构建：当前样式派生值
  const hCur = (sel.ht && hAlignItems.some(([v]) => v === sel.ht) ? sel.ht : 'left') as 'left' | 'center' | 'right';
  const vCur = (sel.vt && vAlignItems.some(([v]) => v === sel.vt) ? sel.vt : 'top') as 'top' | 'middle' | 'bottom';
  const hIcon = hAlignItems.find(([v]) => v === hCur)?.[2] ?? 'icon-paragraph-align-left';
  const vIcon = vAlignItems.find(([v]) => v === vCur)?.[2] ?? 'myf-icon-paragraph-top';
  const fontLabel = sel.ff || t('excel.toolbar.defaultFont');
  const fontSizeLabel = sel.fs != null ? String(sel.fs) : '12';
  const numfmtLabel =
    numFormats.find(
      (f) =>
        f.pattern === sel.numfmt ||
        (f.pattern === 'General' && (sel.numfmt === '' || sel.numfmt === 'General')),
    )?.label ?? t('excel.toolbar.general');

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

    items.push(
      <li key="undo">
        <a role="button" tabIndex={0} title={t('excel.toolbar.undo')} aria-label={t('excel.toolbar.undo')} onClick={run(cmd.undo)}>
          <span className="myf-icon myf-icon-undo" />
        </a>
      </li>,
      <li key="redo">
        <a role="button" tabIndex={0} title={t('excel.toolbar.redo')} aria-label={t('excel.toolbar.redo')} onClick={run(cmd.redo)}>
          <span className="myf-icon myf-icon-redo" />
        </a>
      </li>,
    );
    pushSplit('sp1');
    items.push(
      <li key="formatPainter">
        <a role="button" tabIndex={0} title={t('excel.toolbar.formatPainter')} aria-label={t('excel.toolbar.formatPainter')} onClick={run(cmd.formatPainter)}>
          <span className="myf-icon myf-icon-match-prop" />
        </a>
      </li>,
      <li key="clearFormat">
        <a role="button" tabIndex={0} title={t('action.clearFormat')} aria-label={t('action.clearFormat')} onClick={run(cmd.clearFormatOnly)}>
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
          title={t('excel.toolbar.insert')}
          aria-label={t('excel.toolbar.insert')}
          className={`myf-tool-panel${openPanel === 'insert' || chartPickerSource === 'insert' ? ' active' : ''}`}
          onClick={toggle('insert')}
        >
          <span className="myf-icon myf-icon-add" />
          <span className="myf-icon-text">{t('excel.toolbar.insert')}</span>
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('insert') && (
          <ToolbarDropdownPanel style={{ width: 220 }}>
            <ToolbarMenuItem iconClass="icon-add-img-for-cell" label={t('excel.toolbar.cellImage')} onClick={openImagePicker('cell')} />
            <ToolbarMenuItem iconClass="icon-add-img-for-fix" label={t('excel.toolbar.floatingImage')} onClick={openImagePicker('float')} />
            <ToolbarMenuItem iconClass="myf-icon-add-chart" label={t('excel.toolbar.chart')} onClick={() => openChartPicker('insert')} />
            <ToolbarMenuItem iconClass="icon-add-link" label={t('excel.toolbar.link')} onClick={run(cmd.insertHyperLink)} />
            <ToolbarMenuDivider />
            <ToolbarMenuItem iconClass="myf-icon-insert-top-row" label={t('excel.toolbar.insertRowAbove')} onClick={run(cmd.insertRowAbove)} />
            <ToolbarMenuItem iconClass="myf-icon-insert-bottom-row" label={t('excel.toolbar.insertRowBelow')} onClick={run(cmd.insertRowBelow)} />
            <ToolbarMenuItem iconClass="myf-icon-insert-left-col" label={t('excel.toolbar.insertColLeft')} onClick={run(cmd.insertColLeft)} />
            <ToolbarMenuItem iconClass="myf-icon-insert-right-col" label={t('excel.toolbar.insertColRight')} onClick={run(cmd.insertColRight)} />
            <ToolbarMenuDivider />
            <ToolbarMenuItem iconClass="icon-insert-doc-link-ent" label={t('excel.toolbar.document')} onClick={dev} />
            <ToolbarMenuItem iconClass="icon-insert-file-link" label={t('excel.toolbar.file')} onClick={dev} />
            <ToolbarMenuItem iconClass="icon-add-todo" label={t('excel.toolbar.todo')} onClick={dev} />
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
          title={t('excel.toolbar.numberFormat')}
          aria-label={t('excel.toolbar.numberFormat')}
          className={`myf-tool-panel${openPanel === 'convention' ? ' active' : ''}`}
          onClick={toggle('convention')}
        >
          <span className="myf-icon-text">{numfmtLabel}</span>
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('convention') && (
          <ToolbarDropdownPanel style={{ width: 220 }}>
            {numFormatGroups.map((group, gi) => (
              <Fragment key={gi}>
                {gi > 0 && <ToolbarMenuDivider />}
                {group.map((f) => (
                  <ToolbarMenuItem key={f.label} label={f.label} right={f.sample} onClick={run(() => cmd.setNumberFormatPattern(f.pattern))} />
                ))}
              </Fragment>
            ))}
            <ToolbarMenuDivider />
            <ToolbarMenuItem label={t('excel.toolbar.moreFormats')} onClick={dev} />
          </ToolbarDropdownPanel>
        )}
      </li>,
    );

    // 小数位数（.0 上下箭头）
    items.push(
      <li key="decimal">
        <div className="myf-arrows-container" title={t('excel.toolbar.decimals')}>
          <span className="myf-icon-text">.0</span>
          <div className="myf-right-icons">
            <a role="button" tabIndex={0} className="myf-arrow-icon" title={t('excel.toolbar.increaseDecimals')} onClick={run(cmd.addDecimal)}>
              <span className="myf-icon myf-icons-arrow-up" />
            </a>
            <a role="button" tabIndex={0} className="myf-arrow-icon" title={t('excel.toolbar.decreaseDecimals')} onClick={run(cmd.subtractDecimal)}>
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
          title={t('excel.toolbar.fontFamily')}
          aria-label={t('excel.toolbar.fontFamily')}
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
                  const selected = f === '__default__' ? !sel.ff : sel.ff === f;
                  return (
                    <ToolbarMenuItem
                      key={f}
                      iconClass={selected ? 'myf-icon-selected' : undefined}
                      label={f === '__default__' ? t('excel.toolbar.defaultFont') : f}
                      onClick={run(() => cmd.setFontFamily(f === '__default__' ? 'Arial' : f))}
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
          title={t('excel.toolbar.fontSize')}
          aria-label={t('excel.toolbar.fontSize')}
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
        <a role="button" tabIndex={0} title={t('excel.toolbar.bold')} aria-label={t('excel.toolbar.bold')} className={sel.bl ? 'active' : ''} onClick={run(cmd.toggleBold)}>
          <span className="myf-icon myf-icon-bold" />
        </a>
      </li>,
      <li key="italic">
        <a role="button" tabIndex={0} title={t('excel.toolbar.italic')} aria-label={t('excel.toolbar.italic')} className={sel.it ? 'active' : ''} onClick={run(cmd.toggleItalic)}>
          <span className="myf-icon myf-icon-italic" />
        </a>
      </li>,
      <li key="underline">
        <a role="button" tabIndex={0} title={t('excel.toolbar.underline')} aria-label={t('excel.toolbar.underline')} className={sel.u ? 'active' : ''} onClick={run(cmd.toggleUnderline)}>
          <span className="myf-icon myf-icon-underline" />
        </a>
      </li>,
      <li key="strike">
        <a role="button" tabIndex={0} title={t('excel.toolbar.strikethrough')} aria-label={t('excel.toolbar.strikethrough')} className={sel.st ? 'active' : ''} onClick={run(cmd.toggleStrikeThrough)}>
          <span className="myf-icon myf-icon-delete" />
        </a>
      </li>,
    );

    // 字体颜色 / 填充颜色 / 边框（分体式：图标直接生效，箭头开面板）
    items.push(
      <li className="right-icon" key="fontColor">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} className="arrows-left" title={t('excel.toolbar.fontColor')} onClick={run(() => cmd.setFontColor(sel.cl ?? '#000000'))}>
            <div className="myf-icon myf-icon-fontcolor" />
            {/* 未选色时底线显示白色（空色状态），选色后跟随所选颜色 */}
            <div className="myf-icon-bottom-color" style={{ backgroundColor: sel.cl ?? '#ffffff' }} />
          </a>
          <a role="button" tabIndex={0} className="myf-pr-2" title={t('excel.toolbar.chooseFontColor')} onClick={toggle('fontColor')}>
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
          <a role="button" tabIndex={0} className="arrows-left" title={t('excel.toolbar.fillColor')} onClick={run(() => cmd.setBackgroundColor(sel.bg ?? '#ffffff'))}>
            <div className="myf-icon myf-icon-fill-color" />
            <div className="myf-icon-bottom-color" style={{ backgroundColor: sel.bg ?? '#ffffff' }} />
          </a>
          <a role="button" tabIndex={0} className="myf-pr-2" title={t('excel.toolbar.chooseFillColor')} onClick={toggle('fillColor')}>
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
          <a role="button" tabIndex={0} title={t('excel.toolbar.border')} onClick={run(() => cmd.applyBorder('all', borderStyle, borderColor))}>
            <span className="myf-icon myf-icon-cell-border-all" />
          </a>
          <a role="button" tabIndex={0} className="myf-pr-2" title={t('excel.toolbar.borderStyle')} onClick={toggle('border')}>
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('border') && (
          <div className="myf-x-panel myf-x-visible" style={{ width: 232 }}>
            <div className="myf-border-grid">
              {BORDER_PRESETS.map((preset) => (
                <button key={preset.key} type="button" title={t(preset.titleKey)} onClick={() => cmd.applyBorder(preset.key, borderStyle, borderColor)}>
                  <BorderGlyph edges={preset.edges} diag={preset.diag} />
                </button>
              ))}
            </div>
            <div className="x-split" />
            <BorderSub label={t('excel.toolbar.borderColor')} icon={<PenLine size={13} />}>
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
            <BorderSub label={t('excel.toolbar.borderStyle')} icon={<DashLineGlyph />}>
              {BORDER_STYLES.map((s) => (
                <button
                  key={s.key}
                  type="button"
                  className={`myf-border-style-row${borderStyle === s.key ? ' on' : ''}`}
                  onClick={() => setBorderStyle(s.key)}
                >
                  <span>{t(s.labelKey)}</span>
                  <svg width="34" height="8" viewBox="0 0 34 8">
                    <path d="M1 4 H33" stroke={borderStyle === s.key ? '#1d4ed8' : '#475569'} strokeWidth={s.width} strokeDasharray={s.dasharray} />
                  </svg>
                </button>
              ))}
            </BorderSub>
            <button type="button" className="myf-x-plain" onClick={dev}>
              {t('excel.toolbar.moreBorders')}
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
          <a role="button" tabIndex={0} title={t('excel.toolbar.horizontalAlign')} onClick={run(() => cmd.setHorizontalAlign(hCur))}>
            <span className={`myf-icon ${hIcon}`} />
          </a>
          <a
            role="button"
            tabIndex={0}
            className={`myf-tool-panel myf-pr-2${openPanel === 'hAlign' ? ' active' : ''}`}
            title={t('excel.toolbar.alignment')}
            onClick={toggle('hAlign')}
          >
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('hAlign') && (
          <ToolbarDropdownPanel style={{ width: 180 }}>
            {hAlignItems.map(([v, label, icon]) => (
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
          <a role="button" tabIndex={0} title={t('excel.toolbar.verticalAlign')} onClick={run(() => cmd.setVerticalAlign(vCur))}>
            <span className={`myf-icon ${vIcon}`} />
          </a>
          <a
            role="button"
            tabIndex={0}
            className={`myf-tool-panel myf-pr-2${openPanel === 'vAlign' ? ' active' : ''}`}
            title={t('excel.toolbar.alignment')}
            onClick={toggle('vAlign')}
          >
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('vAlign') && (
          <ToolbarDropdownPanel style={{ width: 180 }}>
            {vAlignItems.map(([v, label, icon]) => (
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
        <a role="button" tabIndex={0} title={t('excel.toolbar.merge')} aria-label={t('excel.toolbar.merge')} onClick={run(cmd.toggleMergeCells)}>
          <span className="myf-icon myf-icon-merge-cells" />
        </a>
      </li>,
      <li key="freeze">
        <a
          role="button"
          tabIndex={0}
          title={t('excel.toolbar.freeze')}
          aria-label={t('excel.toolbar.freeze')}
          className={`myf-tool-panel${openPanel === 'freeze' ? ' active' : ''}`}
          onClick={toggle('freeze')}
        >
          <span className="myf-icon myf-icon-freeze" />
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('freeze') && (
          <ToolbarDropdownPanel style={{ width: 160 }}>
            <ToolbarMenuItem iconClass="myf-icon-freeze" label={t('excel.view.freezeFirstRow')} onClick={run(cmd.freezeFirstRow)} />
            <ToolbarMenuItem iconClass="myf-icon-freeze" label={t('excel.view.freezeFirstColumn')} onClick={run(cmd.freezeFirstColumn)} />
            <ToolbarMenuItem iconClass="myf-icon-freeze" label={t('excel.view.freezeToCell')} onClick={run(cmd.freezeToActiveCell)} />
            <ToolbarMenuItem iconClass="myf-icon-freeze" label={t('excel.view.unfreeze')} onClick={run(cmd.cancelFreeze)} />
          </ToolbarDropdownPanel>
        )}
      </li>,
    );
    pushSplit('sp6');

    // Σ 快速统计 / 图表 / 筛选 / 排序 / @人 / 评论 / 查找替换
    items.push(
      <li className="right-icon" key="sum">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} title={t('excel.toolbar.sum')} onClick={run(() => cmd.applyQuickStat('SUM'))}>
            <span className="myf-icon myf-icon-add-function" />
          </a>
          <a
            role="button"
            tabIndex={0}
            className={`myf-tool-panel myf-pr-2${openPanel === 'sum' ? ' active' : ''}`}
            title={t('excel.toolbar.stats')}
            onClick={toggle('sum')}
          >
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('sum') && (
          <ToolbarDropdownPanel style={{ width: 140 }}>
            <ToolbarMenuItem label={t('excel.stats.sum')} onClick={run(() => cmd.applyQuickStat('SUM'))} />
            <ToolbarMenuItem label={t('excel.stats.average')} onClick={run(() => cmd.applyQuickStat('AVERAGE'))} />
            <ToolbarMenuItem label={t('excel.stats.count')} onClick={run(() => cmd.applyQuickStat('COUNT'))} />
            <ToolbarMenuItem label={t('excel.stats.max')} onClick={run(() => cmd.applyQuickStat('MAX'))} />
            <ToolbarMenuItem label={t('excel.stats.min')} onClick={run(() => cmd.applyQuickStat('MIN'))} />
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="chart" className={chartPickerSource === 'chart' ? 'active' : undefined}>
        <a
          role="button"
          tabIndex={0}
          title={t('excel.toolbar.chart')}
          aria-label={t('excel.toolbar.chart')}
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
          <a role="button" tabIndex={0} title={t('excel.toolbar.filter')} onClick={run(cmd.toggleFilter)}>
            <span className="myf-icon myf-icon-filter" />
          </a>
          <a role="button" tabIndex={0} className="myf-pr-2" title={t('excel.toolbar.filter')} onClick={run(cmd.toggleFilter)}>
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
      </li>,
      <li key="sort">
        <a
          role="button"
          tabIndex={0}
          title={t('excel.toolbar.sort')}
          aria-label={t('excel.toolbar.sort')}
          className={`myf-tool-panel${openPanel === 'sort' ? ' active' : ''}`}
          onClick={toggle('sort')}
        >
          <span className="myf-icon myf-icon-sort" />
          <div className="myf-icon myf-icon-arrow-down" />
        </a>
        {open('sort') && (
          <ToolbarDropdownPanel style={{ width: 120 }}>
            <ToolbarMenuItem iconClass="myf-icon-sort" label={t('excel.sort.ascending')} onClick={run(cmd.sortAscending)} />
            <ToolbarMenuItem iconClass="myf-icon-sort" label={t('excel.sort.descending')} onClick={run(cmd.sortDescending)} />
            <ToolbarMenuItem iconClass="myf-icon-sort" label={t('excel.sort.custom')} onClick={run(cmd.sortCustom)} />
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="mention">
        <a role="button" tabIndex={0} title={t('excel.toolbar.mention')} aria-label={t('excel.toolbar.mention')} onClick={dev}>
          <span className="myf-icon myf-icon-mention-ent" />
        </a>
      </li>,
      <li className="right-icon" key="comment">
        <div className="myf-arrows-container">
          <a role="button" tabIndex={0} title={t('excel.toolbar.addComment')} aria-label={t('excel.toolbar.addComment')} onClick={runComment(cmd.addComment)}>
            <span className="myf-icon myf-icon-comment-ent" />
          </a>
          <a
            role="button"
            tabIndex={0}
            className={`myf-tool-panel myf-pr-2${openPanel === 'comment' ? ' active' : ''}`}
            title={t('excel.toolbar.comment')}
            aria-label={t('excel.toolbar.commentMenu')}
            onClick={toggle('comment')}
          >
            <div className="myf-icon myf-icon-arrow-down" />
          </a>
        </div>
        {open('comment') && (
          <ToolbarDropdownPanel style={{ width: 140 }}>
            <ToolbarMenuItem iconClass="myf-icon-comment-ent" label={t('excel.toolbar.addComment')} onClick={runComment(cmd.addComment)} />
            <ToolbarMenuItem iconClass="myf-icon-comment-ent" label={t('excel.toolbar.manageComments')} onClick={run(cmd.toggleCommentPanel)} />
          </ToolbarDropdownPanel>
        )}
      </li>,
      <li key="findReplace">
        <a role="button" tabIndex={0} title={t('excel.toolbar.findReplace')} aria-label={t('excel.toolbar.findReplace')} onClick={run(cmd.openFindReplace)}>
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
      <input ref={imageInputRef} type="file" accept="image/*" className="hidden" onChange={onImagePicked} />
      <div ref={wrapRef} className="myf-sheet-tools">
        <ul className="myf-toolbar">
          {liveItems.slice(0, overflowed ? cut : liveItems.length)}
          {overflowed && (
            <li className="myf-more-button">
              <a role="button" tabIndex={0} title={t('excel.toolbar.more')} aria-label={t('excel.toolbar.more')} className={moreOpen ? 'active' : ''} onClick={() => { closeAllToolbarDropdowns(); setMoreOpen((v) => !v); setOpenPanel(null); setChartPickerSource(null); }}>
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

const BORDER_STYLES: { key: BorderStyleKey; labelKey: string; width: number; dasharray?: string }[] = [
  { key: 'thin', labelKey: 'excel.border.style.thin', width: 1 },
  { key: 'medium', labelKey: 'excel.border.style.medium', width: 2.5 },
  { key: 'dashed', labelKey: 'excel.border.style.dashed', width: 1.5, dasharray: '4 3' },
  { key: 'dotted', labelKey: 'excel.border.style.dotted', width: 1.5, dasharray: '1.5 2.5' },
  { key: 'double', labelKey: 'excel.border.style.double', width: 3, dasharray: '3 1.5 0.8 1.5' },
];

const BORDER_PRESETS: { key: BorderTypeKey; titleKey: string; edges: string[]; diag?: 'tlbr' | 'bltr' }[] = [
  { key: 'horizontal', titleKey: 'excel.border.preset.horizontal', edges: ['t', 'b', 'h'] },
  { key: 'all', titleKey: 'excel.border.preset.all', edges: ['t', 'b', 'l', 'r', 'h', 'v'] },
  { key: 'outside', titleKey: 'excel.border.preset.outside', edges: ['t', 'b', 'l', 'r'] },
  { key: 'none', titleKey: 'excel.border.preset.none', edges: [] },
  { key: 'tlbr', titleKey: 'excel.border.preset.tlbr', edges: [], diag: 'tlbr' },
  { key: 'left', titleKey: 'excel.border.preset.left', edges: ['l'] },
  { key: 'top', titleKey: 'excel.border.preset.top', edges: ['t'] },
  { key: 'right', titleKey: 'excel.border.preset.right', edges: ['r'] },
  { key: 'bottom', titleKey: 'excel.border.preset.bottom', edges: ['b'] },
  { key: 'bltr', titleKey: 'excel.border.preset.bltr', edges: [], diag: 'bltr' },
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
