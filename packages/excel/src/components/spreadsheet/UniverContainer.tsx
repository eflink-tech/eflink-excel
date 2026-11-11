import { useEffect, useRef } from 'react';
import { useI18n } from '@eflink-tech/office-kit';
import { bindEditor, notifyChanged, saveNow, unbindEditor } from '../../core/saveService';
import { getUniverApi, mountUniver, setUniverLocale } from '../../core/univer/controller';
import { insertQueuedImportCharts } from '../../core/univer/importedCharts';
import { insertQueuedImportAnnotations } from '../../core/univer/importedAnnotations';
import { UNIVER_READY_EVENT } from '../toolbar/FormulaBar';
import { useEditorStore } from '../../store/editorStore';
import type { WorkbookSnapshot } from '../../types/spreadsheet';

export function UniverContainer({ snapshot }: { snapshot: WorkbookSnapshot }) {
  const ref = useRef<HTMLDivElement>(null);
  const { locale } = useI18n();
  const localeRef = useRef(locale);
  localeRef.current = locale;

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const ctrl = mountUniver(el, snapshot, { onChange: () => notifyChanged() }, localeRef.current);
    // Univer 初始没有 activeRange，公式栏首屏无从同步；显式激活 A1 触发 SelectionChanged
    getUniverApi()?.getActiveWorkbook()?.getActiveSheet()?.getRange(0, 0, 1, 1)?.activate();
    const docId = useEditorStore.getState().docId ?? '';
    bindEditor(docId, () => ctrl.getSnapshot());
    // 历史双等号公式已在挂载时修正，需立即落库：否则下次加载/导出仍是错误公式
    if (ctrl.healedFormulas) {
      void saveNow().catch((err) => console.error('自愈公式保存失败', err));
    }
    // 通知 FormulaBar 等：Univer API 已就绪
    window.dispatchEvent(new Event(UNIVER_READY_EVENT));
    // xlsx 导入的图表在此落位（需 Univer 挂载后才能插浮动图片）
    insertQueuedImportCharts();
    // xlsx 导入的超链接/批注在此经 facade 写入单元格（无稳定的 resources 序列化器）
    insertQueuedImportAnnotations();
    return () => {
      unbindEditor();
      ctrl.destroy();
    };
  }, [snapshot]);

  // 界面语言切换：只换 Univer 内置语言包，不重建实例（不丢失当前编辑状态）
  useEffect(() => {
    setUniverLocale(locale);
  }, [locale]);

  // 单元格编辑中禁止滚轮滚动：Univer 编辑框按视口坐标定位，滚动后会脱离单元格悬浮
  useEffect(() => {
    const onWheel = (event: WheelEvent) => {
      const target = event.target as Element | null;
      if (!target?.closest?.('#univer-container')) return;
      if (!getUniverApi()?.getActiveWorkbook()?.isCellEditing()) return;
      event.preventDefault();
      event.stopPropagation();
    };
    window.addEventListener('wheel', onWheel, { capture: true, passive: false });
    return () => window.removeEventListener('wheel', onWheel, { capture: true });
  }, []);

  return <div ref={ref} id="univer-container" />;
}
