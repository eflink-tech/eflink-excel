import { useEffect, useRef } from 'react';
import { bindEditor, notifyChanged, unbindEditor } from '../../core/saveService';
import { getUniverApi, mountUniver } from '../../core/univer/controller';
import { UNIVER_READY_EVENT } from '../toolbar/FormulaBar';
import { useEditorStore } from '../../store/editorStore';
import type { WorkbookSnapshot } from '../../types/spreadsheet';

export function UniverContainer({ snapshot }: { snapshot: WorkbookSnapshot }) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const ctrl = mountUniver(el, snapshot, { onChange: () => notifyChanged() });
    // Univer 初始没有 activeRange，公式栏首屏无从同步；显式激活 A1 触发 SelectionChanged
    getUniverApi()?.getActiveWorkbook()?.getActiveSheet()?.getRange(0, 0, 1, 1)?.activate();
    const docId = useEditorStore.getState().docId ?? '';
    bindEditor(docId, () => ctrl.getSnapshot());
    // 通知 FormulaBar 等：Univer API 已就绪
    window.dispatchEvent(new Event(UNIVER_READY_EVENT));
    return () => {
      unbindEditor();
      ctrl.destroy();
    };
  }, [snapshot]);

  return <div ref={ref} id="univer-container" />;
}
