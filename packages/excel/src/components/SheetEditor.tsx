import { Suspense, lazy, useEffect, useRef, useState } from 'react';
import { saveNow } from '../core/saveService';
import { getDefaultStorage, setDefaultStorage } from '../storage/registry';
import { ConfirmDialogHost } from './ConfirmDialog';
import { FormulaBar } from './toolbar/FormulaBar';
import { SheetToolbar } from './toolbar/SheetToolbar';
import { TitleBar } from './TitleBar';
import { ChartSettingsPanel } from './spreadsheet/ChartSettingsPanel';
import { GridEdgeControls } from './spreadsheet/GridEdgeControls';
import { useEditorStore } from '../store/editorStore';
import type { SheetDocument, WorkbookSnapshot } from '../types/spreadsheet';
import type { StorageAdapter } from '../storage/types';

// Univer 体积较大，编辑器进入时才加载
const UniverContainer = lazy(() => import('./spreadsheet/UniverContainer').then((m) => ({ default: m.UniverContainer })));

/** 品牌区配置：编辑器左上角的 logo 与产品名 */
export interface SheetEditorBrand {
  logo?: string;
  name: string;
}

export interface SheetEditorProps {
  /** 要打开的文档 id，快照经 storage 加载 */
  docId: string;
  /** 文档存储实现；传入后注册为全局默认。缺省用内置 IndexedDB（库名 eflink-excel） */
  storage?: StorageAdapter;
  /** 品牌区（logo + 产品名）；不传或传 false 隐藏 */
  branding?: SheetEditorBrand | false;
  /** 是否显示工具栏（默认显示） */
  showToolbar?: boolean;
  /** 是否显示公式栏（默认显示） */
  showFormulaBar?: boolean;
  /** 文档加载完成回调（宿主可在此记录「最近文档」等） */
  onDocLoaded?: (doc: SheetDocument) => void;
  /** 文档加载失败/不存在回调 */
  onDocError?: (err: unknown) => void;
}

export function SheetEditor({
  docId,
  storage,
  branding = false,
  showToolbar = true,
  showFormulaBar = true,
  onDocLoaded,
  onDocError,
}: SheetEditorProps) {
  const [snapshot, setSnapshot] = useState<WorkbookSnapshot | null>(null);
  // 回调走 ref：父组件传内联函数时避免触发重新加载
  const onLoadedRef = useRef(onDocLoaded);
  onLoadedRef.current = onDocLoaded;
  const onErrorRef = useRef(onDocError);
  onErrorRef.current = onDocError;

  useEffect(() => {
    if (storage) setDefaultStorage(storage);
  }, [storage]);

  useEffect(() => {
    let disposed = false;
    void getDefaultStorage()
      .load(docId)
      .then((doc) => {
        if (disposed) return;
        if (!doc) {
          onErrorRef.current?.(new Error(`文档不存在: ${docId}`));
          return;
        }
        useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
        setSnapshot(doc.snapshot);
        onLoadedRef.current?.(doc);
      })
      .catch((err) => {
        console.error('加载文档失败', err);
        onErrorRef.current?.(err);
      });
    return () => {
      disposed = true;
    };
  }, [docId, storage]);

  useEffect(() => {
    // Ctrl/Cmd+S 手动保存
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        void saveNow();
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);

  if (!snapshot) {
    return <div className="flex h-full items-center justify-center text-sm text-slate-400">加载中…</div>;
  }

  const hasTopRow = Boolean(branding) || showToolbar;
  return (
    <div className="flex h-full flex-col bg-white">
      {/* 顶行：品牌区与工具栏同行 */}
      {hasTopRow && (
        <div className="flex shrink-0 items-stretch border-b border-[#dadbdd] bg-white">
          {branding ? <TitleBar logo={branding.logo} name={branding.name} /> : null}
          {showToolbar ? <SheetToolbar /> : null}
        </div>
      )}
      {showFormulaBar ? <FormulaBar /> : null}
      <div className="relative flex-1">
        <Suspense fallback={<div className="absolute inset-0 flex items-center justify-center text-sm text-slate-400">编辑器加载中…</div>}>
          <UniverContainer snapshot={snapshot} />
        </Suspense>
        {/* 网格边缘追加行列（参考稿同款 ⊕）：紧挨最后一列右侧 / 最后一行下方，随滚动缩放跟随 */}
        <GridEdgeControls />
        {/* 图表设置面板：双击插入的图表后从右侧滑出 */}
        <ChartSettingsPanel />
      </div>
      {/* 确认弹窗宿主（新建文档等二次确认经 useUiStore.requestConfirm 发起） */}
      <ConfirmDialogHost />
    </div>
  );
}
