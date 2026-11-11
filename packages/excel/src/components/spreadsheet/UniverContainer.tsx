import { useEffect, useRef } from 'react';
import { useI18n } from '@eflink-tech/office-kit';
import { bindEditor, notifyChanged, saveNow, unbindEditor } from '../../core/saveService';
import { getUniverApi, mountUniver, setUniverLocale } from '../../core/univer/controller';
import { insertQueuedImportCharts } from '../../core/univer/importedCharts';
import { insertQueuedImportAnnotations } from '../../core/univer/importedAnnotations';
import { insertQueuedImportRetentions } from '../../core/univer/importedRetentions';
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
    // 历史双等号公式/缺失单元格类型已在挂载时修正，需立即落库：否则下次加载/导出仍是旧数据
    if (ctrl.healedSnapshot) {
      void saveNow().catch((err) => console.error('自愈快照保存失败', err));
    }
    // 通知 FormulaBar 等：Univer API 已就绪
    window.dispatchEvent(new Event(UNIVER_READY_EVENT));
    // xlsx 导入的图表在此落位（需 Univer 挂载后才能插浮动图片）
    insertQueuedImportCharts();
    // xlsx 导入的超链接/批注在此经 facade 写入单元格（无稳定的 resources 序列化器）
    insertQueuedImportAnnotations();
    // 条件格式/数据验证：同样在 Univer 挂载后经 facade/命令落位，保存时进入 resources。
    // 落位结果不在导入时的快照里，需立即再落库一次，否则首屏「标记已保存」后可能被丢弃。
    if (insertQueuedImportRetentions() > 0) {
      void saveNow().catch((err) => console.error('导入条件格式/数据验证保存失败', err));
    }
    return () => {
      unbindEditor();
      ctrl.destroy();
    };
  }, [snapshot]);

  // 界面语言切换：只换 Univer 内置语言包，不重建实例（不丢失当前编辑状态）
  useEffect(() => {
    setUniverLocale(locale);
  }, [locale]);

  return <div ref={ref} id="univer-container" />;
}
