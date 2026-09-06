import { Suspense, lazy, useEffect, useRef, useState } from 'react';
import { Share2 } from 'lucide-react';
import { saveNow, currentSnapshot } from '../core/saveService';
import { getDefaultStorage, setDefaultStorage } from '../storage/registry';
import { getExcelShareHandler } from '../core/share/shareBridge';
import { ConfirmDialogHost } from './ConfirmDialog';
import { ShareDialog } from './ShareDialog';
import { FormulaBar } from './toolbar/FormulaBar';
import { SheetToolbar } from './toolbar/SheetToolbar';
import { TitleBar } from './TitleBar';
import { ChartSettingsPanel } from './spreadsheet/ChartSettingsPanel';
import { GridEdgeControls } from './spreadsheet/GridEdgeControls';
import { useEditorStore } from '../store/editorStore';
import type { SheetDocument, WorkbookSnapshot } from '../types/spreadsheet';
import type { ExcelShareDoc } from '../core/share/shareBridge';
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

  // 分享弹窗（doc 为点击"分享"时刻的快照，弹窗期间编辑不影响本次分享内容）
  const [shareOpen, setShareOpen] = useState(false);
  const [shareDoc, setShareDoc] = useState<ExcelShareDoc | null>(null);
  const closeShare = useRef(() => setShareOpen(false)).current;
  // 分享前先落库最新快照，再捕获当前文档
  const openShare = async () => {
    await saveNow();
    const { docId, title } = useEditorStore.getState();
    const snap = currentSnapshot();
    if (!docId || !snap) return;
    setShareDoc({ id: docId, title, snapshot: snap });
    setShareOpen(true);
  };

  if (!snapshot) {
    return <div className="flex h-full items-center justify-center text-sm text-slate-400">加载中…</div>;
  }

  const hasTopRow = Boolean(branding) || showToolbar;
  return (
    <div className="flex h-full flex-col bg-white">
      {/* 顶行：品牌区与工具栏同行，行尾分享入口 */}
      {hasTopRow && (
        <div className="flex shrink-0 items-stretch border-b border-[#dadbdd] bg-white">
          {branding ? <TitleBar logo={branding.logo} name={branding.name} /> : null}
          {showToolbar ? <SheetToolbar /> : null}
          {/* 分享入口仅在宿主注入分享实现后出现（纯组件独立运行时不显示） */}
          {getExcelShareHandler() !== null && (
            <div className="flex items-center px-2">
              <button
                type="button"
                onClick={() => void openShare()}
                title="生成分享链接"
                className="flex h-7 items-center gap-1 rounded-md px-2 text-xs text-[#5f6062] transition-colors hover:bg-[#f2f3f4] hover:text-[#26282b]"
              >
                <Share2 size={14} />
                分享
              </button>
            </div>
          )}
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
      <ShareDialog open={shareOpen} doc={shareDoc} onClose={closeShare} />
    </div>
  );
}
