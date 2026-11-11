import { useCallback, useEffect, useRef, useState } from 'react';
import { useI18n } from '@eflink-tech/office-kit';
import { IDescriptionService, IRefSelectionsService } from '@univerjs/preset-sheets-core';
// 注册表格专属文案，保证组件脱离 I18nProvider 单独渲染时也有中文兜底
import '../../i18n';
import { getUniverApi, getUniverInjector } from '../../core/univer/controller';
import {
  applyFunctionName,
  buildSignature,
  extractRefs,
  filterSuggestions,
  functionContextAtCaret,
  parseA1Body,
  readTokenAtCaret,
  REF_FILL_ALPHA,
  REF_STROKE,
  REF_STROKE_DASH,
  type SignatureParam,
  type SuggestItem,
} from './formulaAssist';

/** Univer 实例就绪事件（UniverContainer 创建完成后派发） */
export const UNIVER_READY_EVENT = 'doculuma-excel:univer-ready';

interface FunctionParamLike {
  name: string;
  detail?: string;
  require?: boolean | number;
  repeat?: boolean | number;
}

interface FunctionInfoLike {
  functionName?: string;
  abstract?: string;
  functionParameter?: FunctionParamLike[];
}

interface DescriptionServiceLike {
  getSearchListByNameFirstLetter(text: string): SuggestItem[];
  getFunctionInfo(text: string): FunctionInfoLike | undefined;
}

interface RefSelectionsServiceLike {
  setSelections(selections: unknown[]): void;
}

interface FormulaServices {
  desc: DescriptionServiceLike | null;
  refSel: RefSelectionsServiceLike | null;
}

/** 从 Univer DI 容器取函数描述与引用高亮服务 */
function resolveFormulaServices(): Promise<FormulaServices> {
  const getInjector = getUniverInjector as unknown;
  const injector = typeof getInjector === 'function'
    ? (getInjector as () => { get?(token: unknown): unknown } | null)()
    : null;
  if (!injector?.get) return Promise.resolve({ desc: null, refSel: null });
  const get = injector.get.bind(injector);
  const read = <T,>(token: unknown): T | null => {
    try {
      return (get(token) as T) ?? null;
    } catch {
      return null;
    }
  };
  return Promise.resolve({
    desc: read<DescriptionServiceLike>(IDescriptionService),
    refSel: read<RefSelectionsServiceLike>(IRefSelectionsService),
  });
}

function hexToRgba(hex: string, alpha: number): string {
  const n = Number.parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}

interface SuggestionState {
  items: SuggestItem[];
  index: number;
  token: string;
}

interface HintState {
  name: string;
  params: SignatureParam[];
  detail: string;
}

interface PickState {
  base: string;
  at: number;
  active: boolean;
}

interface EditRangeLike {
  getFormula(): string;
  getValue(): unknown;
  setValue(value: string | number): unknown;
  setFormula(formula: string): unknown;
  activate?(): unknown;
}

/** 企微式公式栏：左侧名称框（当前单元格引用）+ 右侧内容/公式编辑区 */
export function FormulaBar() {
  const { t } = useI18n();
  const [notation, setNotation] = useState('');
  const [value, setValue] = useState('');
  const [suggestions, setSuggestions] = useState<SuggestionState | null>(null);
  const [hint, setHint] = useState<HintState | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLUListElement>(null);
  const editingRef = useRef(false);
  const valueRef = useRef('');
  const caretRef = useRef(0);
  // 公式栏编辑的是按下 fx 时的单元格；拾取引用会移动选区，提交时必须写回原单元格
  const editTargetRef = useRef<EditRangeLike | null>(null);
  const servicesRef = useRef<FormulaServices>({ desc: null, refSel: null });
  const pickRef = useRef<PickState | null>(null);
  const gridInteractionRef = useRef(false);
  const refocusingRef = useRef(false);
  const disposableRef = useRef<{ dispose(): void } | null>(null);

  const clearHighlight = useCallback(() => {
    const refSel = servicesRef.current.refSel;
    if (!refSel) return;
    try {
      refSel.setSelections([]);
    } catch {
      // Univer 尚未就绪
    }
  }, []);

  const updateHighlight = useCallback((text: string) => {
    const refSel = servicesRef.current.refSel;
    const wb = getUniverApi()?.getActiveWorkbook();
    const sheet = wb?.getActiveSheet();
    if (!refSel || !wb || !sheet) return;
    const sheetId = sheet.getSheetId();
    const selections: unknown[] = [];
    if (text.startsWith('=')) {
      for (const ref of extractRefs(text)) {
        if (ref.sheetName) {
          const target = wb.getSheetByName(ref.sheetName);
          if (!target || target.getSheetId() !== sheetId) continue;
        }
        const parsed = parseA1Body(ref.text.slice(ref.text.lastIndexOf('!') + 1));
        if (!parsed) continue;
        selections.push({
          range: { ...parsed, unitId: wb.getId(), sheetId },
          primary: null,
          style: {
            id: `formula-ref-${selections.length}`,
            stroke: REF_STROKE,
            strokeWidth: 1,
            strokeDash: REF_STROKE_DASH,
            fill: hexToRgba(REF_STROKE, REF_FILL_ALPHA),
            widgets: { tl: true, tc: true, tr: true, ml: true, mr: true, bl: true, bc: true, br: true },
            widgetSize: 6,
            widgetStrokeWidth: 1,
            widgetStroke: REF_STROKE,
          },
        });
      }
    }
    try {
      refSel.setSelections(selections);
    } catch {
      // Univer 尚未就绪
    }
  }, []);

  const refreshSuggestions = useCallback((text: string, caret: number) => {
    const desc = servicesRef.current.desc;
    if (!desc || !text.startsWith('=')) {
      setSuggestions(null);
      return;
    }
    const { token, inString } = readTokenAtCaret(text, caret);
    if (inString || !token) {
      setSuggestions(null);
      return;
    }
    const items = filterSuggestions(desc.getSearchListByNameFirstLetter(token), token);
    setSuggestions(items.length > 0 ? { items, index: 0, token } : null);
  }, []);

  const refreshHint = useCallback((text: string, caret: number) => {
    const desc = servicesRef.current.desc;
    if (!desc || !text.startsWith('=')) {
      setHint(null);
      return;
    }
    const context = functionContextAtCaret(text, caret);
    if (!context) {
      setHint(null);
      return;
    }
    const info = desc.getFunctionInfo(context.name);
    if (!info) {
      setHint(null);
      return;
    }
    const params = info.functionParameter ?? [];
    const activeParam = params[Math.max(0, Math.min(context.argIndex, params.length - 1))];
    setHint({
      name: info.functionName || context.name,
      params: buildSignature(params.map((param) => ({
        name: param.name,
        require: param.require === undefined ? true : Boolean(param.require),
        repeat: Boolean(param.repeat),
      })), context.argIndex),
      detail: activeParam?.detail || info.abstract || '',
    });
  }, []);

  const updateAssist = useCallback((text: string, caret: number) => {
    refreshSuggestions(text, caret);
    refreshHint(text, caret);
    updateHighlight(text);
  }, [refreshHint, refreshSuggestions, updateHighlight]);

  // 服务随 Univer 实例重建而失效，重新解析
  useEffect(() => {
    let cancelled = false;
    const load = () => {
      void resolveFormulaServices().then((services) => {
        if (cancelled) return;
        servicesRef.current = services;
        if (editingRef.current) {
          refreshSuggestions(valueRef.current, caretRef.current);
          refreshHint(valueRef.current, caretRef.current);
          updateHighlight(valueRef.current);
        }
      });
    };
    load();
    const onReady = () => {
      servicesRef.current = { desc: null, refSel: null };
      load();
    };
    window.addEventListener(UNIVER_READY_EVENT, onReady);
    return () => {
      cancelled = true;
      window.removeEventListener(UNIVER_READY_EVENT, onReady);
    };
  }, [refreshHint, refreshSuggestions, updateHighlight]);

  const readRangeText = useCallback(() => {
    const range = getUniverApi()?.getActiveWorkbook()?.getActiveRange();
    if (!range) return;
    const formula = range.getFormula();
    if (formula) {
      setValue(formula.startsWith('=') ? formula : '=' + formula);
      return;
    }
    const v = range.getValue();
    setValue(typeof v === 'object' || v == null ? '' : String(v));
  }, []);

  const endSession = useCallback(() => {
    editingRef.current = false;
    pickRef.current = null;
    editTargetRef.current = null;
    setSuggestions(null);
    setHint(null);
    clearHighlight();
    readRangeText();
  }, [clearHighlight, readRangeText]);

  const syncFromSelection = useCallback(() => {
    const wb = getUniverApi()?.getActiveWorkbook();
    const range = wb?.getActiveRange();
    if (!range) return;
    setNotation(range.getA1Notation());
    if (editingRef.current) {
      const pick = pickRef.current;
      if (pick?.active) {
        const a1 = range.getA1Notation();
        const next = pick.base.slice(0, pick.at) + a1 + pick.base.slice(pick.at);
        valueRef.current = next;
        caretRef.current = pick.at + a1.length;
        setValue(next);
        setSuggestions(null);
        refreshHint(next, caretRef.current);
        updateHighlight(next);
      }
      return;
    }
    readRangeText();
  }, [readRangeText, refreshHint, updateHighlight]);

  useEffect(() => {
    // 每次收到 ready 事件都重新订阅当前实例（StrictMode 下实例会重建，旧订阅随实例销毁失效）
    const subscribe = () => {
      const api = getUniverApi();
      if (!api) return;
      disposableRef.current?.dispose();
      disposableRef.current = api.addEvent(api.Event.SelectionChanged, syncFromSelection);
      syncFromSelection();
    };
    subscribe();
    window.addEventListener(UNIVER_READY_EVENT, subscribe);
    return () => {
      window.removeEventListener(UNIVER_READY_EVENT, subscribe);
      disposableRef.current?.dispose();
      disposableRef.current = null;
    };
  }, [syncFromSelection]);

  // 公式栏编辑中点击表格：保留草稿，鼠标松开后把焦点交还输入框
  useEffect(() => {
    const onMouseDown = (event: MouseEvent) => {
      if (!editingRef.current) return;
      const target = event.target as Element | null;
      if (!target?.closest?.('#univer-container')) return;
      gridInteractionRef.current = true;
      if (!pickRef.current) {
        pickRef.current = { base: valueRef.current, at: caretRef.current, active: true };
      } else {
        pickRef.current.active = true;
      }
    };
    const onMouseUp = () => {
      window.setTimeout(() => {
        gridInteractionRef.current = false;
        // 点击可能未触发 SelectionChanged（焦点切换时序不同），松手时按当前选区补一次拾取
        if (editingRef.current && pickRef.current?.active) syncFromSelection();
        if (pickRef.current) pickRef.current.active = false;
        if (editingRef.current && document.activeElement !== inputRef.current) {
          const el = inputRef.current;
          if (el) {
            refocusingRef.current = true;
            el.focus({ preventScroll: true });
            el.setSelectionRange(caretRef.current, caretRef.current);
            refocusingRef.current = false;
          }
        }
      }, 0);
    };
    document.addEventListener('mousedown', onMouseDown, true);
    document.addEventListener('mouseup', onMouseUp, true);
    return () => {
      document.removeEventListener('mousedown', onMouseDown, true);
      document.removeEventListener('mouseup', onMouseUp, true);
    };
  }, [syncFromSelection]);

  useEffect(() => () => clearHighlight(), [clearHighlight]);

  useEffect(() => {
    if (!suggestions) return;
    const active = listRef.current?.querySelector<HTMLElement>('[data-active="true"]');
    active?.scrollIntoView?.({ block: 'nearest' });
  }, [suggestions]);

  const clearAssist = () => {
    setSuggestions(null);
    setHint(null);
  };

  const acceptSuggestion = (item: SuggestItem) => {
    const { text, caret } = applyFunctionName(valueRef.current, caretRef.current, item.name);
    valueRef.current = text;
    caretRef.current = caret;
    setValue(text);
    setSuggestions(null);
    pickRef.current = null;
    requestAnimationFrame(() => {
      const el = inputRef.current;
      if (!el) return;
      el.focus({ preventScroll: true });
      el.setSelectionRange(caret, caret);
    });
    refreshHint(text, caret);
    updateHighlight(text);
  };

  const commit = () => {
    const text = valueRef.current;
    const target = editTargetRef.current ?? getUniverApi()?.getActiveWorkbook()?.getActiveRange() ?? null;
    editingRef.current = false;
    pickRef.current = null;
    editTargetRef.current = null;
    clearAssist();
    clearHighlight();
    inputRef.current?.blur();
    if (!target) return;
    if (text === '') {
      target.setValue('');
    } else if (text.startsWith('=')) {
      target.setFormula(text);
    } else if (/^-?\d+(\.\d+)?$/.test(text.trim())) {
      target.setValue(Number(text.trim()));
    } else {
      target.setValue(text);
    }
    // 拾取引用后选区停在引用单元格，提交完回到被编辑单元格
    target.activate?.();
  };

  const cancel = () => {
    const target = editTargetRef.current;
    inputRef.current?.blur();
    endSession();
    target?.activate?.();
  };

  const updateCaret = () => {
    const el = inputRef.current;
    if (!el) return;
    caretRef.current = el.selectionStart ?? valueRef.current.length;
    if (!editingRef.current) return;
    refreshHint(valueRef.current, caretRef.current);
    if (suggestions) refreshSuggestions(valueRef.current, caretRef.current);
  };

  const changeValue = (text: string, caret: number) => {
    editingRef.current = true;
    pickRef.current = null;
    valueRef.current = text;
    caretRef.current = caret;
    setValue(text);
    updateAssist(text, caret);
  };

  const activeHint = suggestions ? null : hint;

  return (
    <div className="relative flex h-9 shrink-0 items-center border-b border-slate-200 bg-white">
      <input
        readOnly
        value={notation}
        placeholder="-"
        className="h-full w-24 shrink-0 border-r border-slate-200 bg-white px-3 text-[13px] text-slate-700 outline-none"
        title={t('excel.formula.currentCell')}
      />
      <span className="px-2 font-serif text-[13px] italic text-slate-400">fx</span>
      <input
        ref={inputRef}
        value={value}
        placeholder={t('excel.formula.placeholder')}
        className="h-full min-w-0 flex-1 px-1 text-[13px] text-slate-800 outline-none"
        autoComplete="off"
        spellCheck={false}
        onChange={(e) => {
          changeValue(e.target.value, e.target.selectionStart ?? e.target.value.length);
        }}
        onFocus={(e) => {
          // 只有重新进入编辑才记录目标单元格；拾取引用后的程序化聚焦不覆盖
          if (!editingRef.current) {
            editTargetRef.current = getUniverApi()?.getActiveWorkbook()?.getActiveRange() ?? null;
          }
          editingRef.current = true;
          if (!refocusingRef.current) pickRef.current = null;
          valueRef.current = e.target.value;
          caretRef.current = e.target.selectionStart ?? e.target.value.length;
          if (valueRef.current.startsWith('=')) {
            refreshHint(valueRef.current, caretRef.current);
            updateHighlight(valueRef.current);
          }
        }}
        onClick={updateCaret}
        onSelect={updateCaret}
        onKeyUp={(e) => {
          if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key)) updateCaret();
        }}
        onBlur={() => {
          window.setTimeout(() => {
            if (gridInteractionRef.current) return;
            endSession();
          }, 0);
        }}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
            if (!suggestions) return;
            e.preventDefault();
            const delta = e.key === 'ArrowDown' ? 1 : -1;
            const next = (suggestions.index + delta + suggestions.items.length) % suggestions.items.length;
            setSuggestions({ ...suggestions, index: next });
            return;
          }
          if (e.key === 'Tab' && suggestions) {
            e.preventDefault();
            acceptSuggestion(suggestions.items[suggestions.index]);
            return;
          }
          if (e.key === 'Enter') {
            e.preventDefault();
            if (suggestions) {
              acceptSuggestion(suggestions.items[suggestions.index]);
              return;
            }
            commit();
            return;
          }
          if (e.key === 'Escape') {
            e.preventDefault();
            if (suggestions) {
              clearAssist();
              return;
            }
            cancel();
          }
        }}
      />
      {suggestions ? (
        <ul
          ref={listRef}
          role="listbox"
          className="absolute top-full left-24 z-[60] max-h-64 w-[360px] overflow-y-auto rounded-b border border-t-0 border-slate-200 bg-white py-1 shadow-lg"
        >
          {suggestions.items.map((item, index) => (
            <li
              key={item.name}
              role="option"
              aria-selected={index === suggestions.index}
              data-active={index === suggestions.index}
              className={`flex cursor-pointer items-baseline gap-3 px-3 py-1.5 text-[12px] ${
                index === suggestions.index ? 'bg-[#e8f3ff] text-slate-800' : 'text-slate-700 hover:bg-slate-50'
              }`}
              onMouseDown={(e) => {
                e.preventDefault();
                acceptSuggestion(item);
              }}
              onMouseEnter={() => setSuggestions({ ...suggestions, index })}
            >
              <span className="w-24 shrink-0 font-mono font-medium">
                <span className="font-semibold text-[#1a73e8]">{item.name.slice(0, suggestions.token.length)}</span>
                {item.name.slice(suggestions.token.length)}
              </span>
              <span className="min-w-0 flex-1 truncate text-slate-500">{item.desc}</span>
            </li>
          ))}
        </ul>
      ) : activeHint ? (
        <div className="pointer-events-none absolute top-full left-24 z-[60] mt-1 max-w-[520px] rounded border border-slate-200 bg-white px-3 py-1.5 shadow-lg">
          <div className="whitespace-nowrap text-[12px] text-slate-700">
            {activeHint.name}({activeHint.params.map((param, index) => (
              <span key={`${param.label}-${index}`}>
                {index > 0 ? ', ' : ''}
                <span className={param.active ? 'font-semibold text-[#1a73e8]' : ''}>{param.label}</span>
              </span>
            ))})
          </div>
          {activeHint.detail ? (
            <div className="mt-0.5 truncate text-[11px] text-slate-400">{activeHint.detail}</div>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}
