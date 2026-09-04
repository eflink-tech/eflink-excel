import { useEffect, useRef, useState } from 'react';
import { getUniverApi } from '../../core/univer/controller';

/** Univer 实例就绪事件（UniverContainer 创建完成后派发） */
export const UNIVER_READY_EVENT = 'eflink-excel:univer-ready';

/** 企微式公式栏：左侧名称框（当前单元格引用）+ 右侧内容/公式编辑区 */
export function FormulaBar() {
  const [notation, setNotation] = useState('');
  const [value, setValue] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  const editingRef = useRef(false);
  const disposableRef = useRef<{ dispose(): void } | null>(null);

  useEffect(() => {
    const syncFromSelection = () => {
      const wb = getUniverApi()?.getActiveWorkbook();
      const range = wb?.getActiveRange();
      if (!range) return;
      setNotation(range.getA1Notation());
      if (editingRef.current) return;
      const formula = range.getFormula();
      if (formula) {
        setValue(formula.startsWith('=') ? formula : '=' + formula);
        return;
      }
      const v = range.getValue();
      setValue(typeof v === 'object' || v == null ? '' : String(v));
    };
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
  }, []);

  const commit = () => {
    editingRef.current = false;
    inputRef.current?.blur();
    const wb = getUniverApi()?.getActiveWorkbook();
    const range = wb?.getActiveRange();
    if (!range) return;
    const text = value;
    if (text === '') {
      range.setValue('');
    } else if (text.startsWith('=')) {
      range.setFormula(text);
    } else if (/^-?\d+(\.\d+)?$/.test(text.trim())) {
      range.setValue(Number(text.trim()));
    } else {
      range.setValue(text);
    }
  };

  return (
    <div className="flex h-9 shrink-0 items-center border-b border-slate-200 bg-white">
      <input
        readOnly
        value={notation}
        placeholder="-"
        className="h-full w-24 shrink-0 border-r border-slate-200 bg-white px-3 text-[13px] text-slate-700 outline-none"
        title="当前单元格"
      />
      <span className="px-2 font-serif text-[13px] italic text-slate-400">fx</span>
      <input
        ref={inputRef}
        value={value}
        placeholder="输入内容或公式，回车确认"
        className="h-full min-w-0 flex-1 px-1 text-[13px] text-slate-800 outline-none"
        onChange={(e) => {
          editingRef.current = true;
          setValue(e.target.value);
        }}
        onFocus={() => {
          editingRef.current = true;
        }}
        onBlur={() => {
          editingRef.current = false;
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') commit();
          if (e.key === 'Escape') inputRef.current?.blur();
        }}
      />
    </div>
  );
}
