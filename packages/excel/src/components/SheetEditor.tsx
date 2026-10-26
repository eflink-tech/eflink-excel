import { Suspense, lazy, useEffect, useRef, useState } from 'react';
import { Keyboard, MessageCircle, Save, Share2, Sparkles } from 'lucide-react';
import { I18nProvider, OfficeHeader, useI18n, type OfficeLocale, type OfficeMenu } from '@eflink-tech/office-kit';
import { saveNow, currentSnapshot } from '../core/saveService';
import { getDefaultStorage, setDefaultStorage } from '../storage/registry';
import { readDraft, clearDraft } from '../storage/draft';
import { getExcelShareHandler } from '../core/share/shareBridge';
import { getEditorBackHref } from '../core/chrome';
import {
  exportEfexcelAction,
  exportPngAction,
  exportXlsxAction,
  importFileAction,
  importXlsxMenuAction,
  newDocAction,
  saveAction,
} from '../core/fileActions';
import * as cmd from '../core/univer/commands';
import { EXCEL_MESSAGES, setExcelLocale } from '../i18n';
import { ConfirmDialogHost } from './ConfirmDialog';
import { ShareDialog } from './ShareDialog';
import { FormulaBar } from './toolbar/FormulaBar';
import { SheetToolbar } from './toolbar/SheetToolbar';
import { ShortcutsDialog } from './toolbar/ShortcutsDialog';
import { BottomBar } from './BottomBar';
import { ChartSettingsPanel } from './spreadsheet/ChartSettingsPanel';
import { GridEdgeControls } from './spreadsheet/GridEdgeControls';
import { AIChatPanel } from './ai/AIChatPanel';
import { useEditorStore } from '../store/editorStore';
import { useDocumentsStore } from '../store/documentsStore';
import { useUiStore } from '../store/uiStore';
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
  /** 界面语言；缺省 zh-CN */
  locale?: OfficeLocale | string;
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

/** 外层只负责语言上下文；内层使用 useI18n 取文案 */
export function SheetEditor(props: SheetEditorProps) {
  return (
    <I18nProvider locale={props.locale} messages={EXCEL_MESSAGES}>
      <SheetEditorInner {...props} />
    </I18nProvider>
  );
}

function SheetEditorInner({
  docId,
  locale,
  storage,
  branding = false,
  showToolbar = true,
  showFormulaBar = true,
  onDocLoaded,
  onDocError,
}: SheetEditorProps) {
  const { t } = useI18n();
  const tRef = useRef(t);
  tRef.current = t;
  const [snapshot, setSnapshot] = useState<WorkbookSnapshot | null>(null);
  const aiPanelOpen = useUiStore((s) => s.aiPanelOpen);
  const toggleAIPanel = useUiStore((s) => s.toggleAIPanel);
  // 导入替换等场景触发文档重载（docId 不变，仅此计数变化）
  const reloadToken = useEditorStore((s) => s.reloadToken);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const importInputRef = useRef<HTMLInputElement>(null);
  // 回调走 ref：父组件传内联函数时避免触发重新加载
  const onLoadedRef = useRef(onDocLoaded);
  onLoadedRef.current = onDocLoaded;
  const onErrorRef = useRef(onDocError);
  onErrorRef.current = onDocError;

  useEffect(() => {
    setExcelLocale(locale);
  }, [locale]);

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
          onErrorRef.current?.(new Error(tRef.current('excel.doc.notFound', { id: docId })));
          return;
        }
        useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
        setSnapshot(doc.snapshot);
        onLoadedRef.current?.(doc);
        // Univer 应用初始快照的命令流不属于用户编辑：待其落定后复位脏标记，
        // 避免刚打开文档就显示"● 未保存"并误触离页拦截
        window.setTimeout(() => {
          if (!disposed) useEditorStore.getState().markSaved();
        }, 1000);
      })
      .catch((err) => {
        if (disposed) return;
        console.error('加载文档失败', err);
        // 云端加载失败（网络异常等）：回退本地草稿恢复表格；云端成功时不读草稿
        const draft = readDraft(docId);
        if (draft) {
          const knownTitle = useDocumentsStore.getState().docs.find((d) => d.id === docId)?.title;
          useEditorStore.getState().openDoc({ id: docId, title: knownTitle ?? docId });
          useEditorStore.getState().markDirty(); // 草稿内容尚未云端落库，视为未保存
          setSnapshot(draft);
          return;
        }
        onErrorRef.current?.(err);
      });
    return () => {
      disposed = true;
    };
  }, [docId, storage, reloadToken]);

  useEffect(() => {
    // Ctrl/Cmd+S 手动保存（云端保存；成功后由 saveService 清 dirty 并删除本地草稿）
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        saveNow().catch((err) => {
          console.error('保存失败', err);
          useUiStore.getState().showToast(t('excel.toast.saveFailed'));
        });
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [t]);

  // 宿主 bridge：查询脏状态 / 触发云端保存 / 丢弃本地草稿（卸载时移除）
  useEffect(() => {
    (window as unknown as Record<string, unknown>).__eflinkEditorBridge = {
      isDirty: () => useEditorStore.getState().dirty,
      save: () => saveNow(),
      discard: () => clearDraft(docId),
    };
    return () => {
      delete (window as unknown as Record<string, unknown>).__eflinkEditorBridge;
    };
  }, [docId]);

  // 分享弹窗（doc 为点击"分享"时刻的快照，弹窗期间编辑不影响本次分享内容）
  const [shareOpen, setShareOpen] = useState(false);
  const [shareDoc, setShareDoc] = useState<ExcelShareDoc | null>(null);
  const closeShare = useRef(() => setShareOpen(false)).current;
  const canShare = getExcelShareHandler() !== null;
  // 分享前先落库最新快照，再捕获当前文档
  const openShare = async () => {
    await saveNow();
    const { docId, title } = useEditorStore.getState();
    const snap = currentSnapshot();
    if (!docId || !snap) return;
    setShareDoc({ id: docId, title, snapshot: snap });
    setShareOpen(true);
  };

  const menus: OfficeMenu[] = [
    {
      key: 'file',
      label: t('shell.menu.file'),
      items: [
        [
          { key: 'new', label: t('excel.file.new'), onSelect: () => void newDocAction() },
          { key: 'save', label: t('action.save'), icon: <Save size={16} />, shortcut: 'Ctrl S', onSelect: () => void saveAction() },
        ],
        [
          { key: 'importJson', label: t('excel.file.importJson'), onSelect: () => importInputRef.current?.click() },
          { key: 'importXlsx', label: t('excel.file.importXlsx'), onSelect: () => void importXlsxMenuAction() },
        ],
        [
          { key: 'exportJson', label: t('excel.file.exportJson'), onSelect: () => void exportEfexcelAction() },
          { key: 'exportXlsx', label: t('excel.file.exportXlsx'), onSelect: () => void exportXlsxAction() },
          { key: 'exportPng', label: t('excel.file.exportPng'), onSelect: () => void exportPngAction() },
        ],
      ],
    },
    {
      key: 'edit',
      label: t('shell.menu.edit'),
      items: [
        [
          { key: 'undo', label: t('action.undo'), shortcut: 'Ctrl Z', onSelect: cmd.undo },
          { key: 'redo', label: t('action.redo'), shortcut: 'Ctrl Y', onSelect: cmd.redo },
        ],
        [
          { key: 'clearContent', label: t('excel.edit.clearContent'), onSelect: cmd.clearContentOnly },
          { key: 'clearFormat', label: t('action.clearFormat'), onSelect: cmd.clearFormatOnly },
        ],
      ],
    },
    {
      key: 'view',
      label: t('shell.menu.view'),
      items: [
        [
          { key: 'freezeFirstRow', label: t('excel.view.freezeFirstRow'), onSelect: cmd.freezeFirstRow },
          { key: 'freezeFirstColumn', label: t('excel.view.freezeFirstColumn'), onSelect: cmd.freezeFirstColumn },
          { key: 'freezeToCell', label: t('excel.view.freezeToCell'), onSelect: cmd.freezeToActiveCell },
          { key: 'unfreeze', label: t('excel.view.unfreeze'), onSelect: cmd.cancelFreeze },
        ],
      ],
    },
    {
      key: 'help',
      label: t('shell.menu.help'),
      items: [
        [
          { key: 'shortcuts', label: t('action.shortcuts'), icon: <Keyboard size={16} />, onSelect: () => setShortcutsOpen(true) },
          { key: 'feedback', label: t('shell.feedback'), icon: <MessageCircle size={16} />, onSelect: () => window.open('/contact', '_blank') },
        ],
      ],
    },
  ];

  if (!snapshot) {
    return (
      <div className="flex h-full flex-col">
        <OfficeHeader
          brand={branding ? { logo: branding.logo, name: branding.name } : false}
          backHref={getEditorBackHref()}
        />
        <div className="flex flex-1 items-center justify-center text-sm text-slate-400">{t('excel.loading')}</div>
      </div>
    );
  }

  return (
    <div className="flex h-full flex-col bg-white">
      {/* 顶栏：返回 + 品牌 + 横向菜单 + 分享/AI */}
      <OfficeHeader
        brand={branding ? { logo: branding.logo, name: branding.name } : false}
        backHref={getEditorBackHref()}
        menus={menus}
        actions={
          <>
            {canShare && (
              <button
                type="button"
                onClick={() => void openShare()}
                title={t('shell.share')}
                aria-label={t('shell.share')}
                className="flex h-7 shrink-0 items-center gap-1 rounded-md px-2 text-xs text-[#5f6062] transition-colors hover:bg-[#f2f3f4] hover:text-[#26282b]"
              >
                <Share2 size={14} />
                {t('shell.share')}
              </button>
            )}
            <button
              type="button"
              onClick={toggleAIPanel}
              title={t('excel.toolbar.ai')}
              aria-label={t('excel.toolbar.ai')}
              data-testid="ai-toggle"
              className={`flex h-7 shrink-0 items-center gap-1 rounded-md px-2 text-xs transition-colors ${
                aiPanelOpen ? 'bg-[#e7f4ec] text-[#107c41]' : 'text-[#5f6062] hover:bg-[#f2f3f4] hover:text-[#26282b]'
              }`}
            >
              <Sparkles size={14} />
              {t('excel.toolbar.ai')}
            </button>
          </>
        }
      />
      <input
        ref={importInputRef}
        type="file"
        accept=".efx.json,.efexcel,application/json"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) void importFileAction(file);
          e.target.value = '';
        }}
      />
      {/* 第二行：表格工具栏 */}
      {showToolbar && (
        <div className="relative z-30 flex shrink-0 items-stretch border-b border-[#dadbdd] bg-white">
          <SheetToolbar />
        </div>
      )}
      {showFormulaBar ? <FormulaBar /> : null}
      <div className="relative flex flex-1">
        <div className="relative min-w-0 flex-1">
          <Suspense fallback={<div className="absolute inset-0 flex items-center justify-center text-sm text-slate-400">{t('excel.editor.loading')}</div>}>
            <UniverContainer snapshot={snapshot} />
          </Suspense>
          {/* 网格边缘追加行列（参考稿同款 ⊕）：紧挨最后一列右侧 / 最后一行下方，随滚动缩放跟随 */}
          <GridEdgeControls />
          {/* 图表设置面板：双击插入的图表后从右侧滑出 */}
          <ChartSettingsPanel />
        </div>
        {/* AI 助手面板（右侧独立列，参考 eflink-draw / eflink-pptx） */}
        {aiPanelOpen && <AIChatPanel />}
      </div>
      {/* 底部状态栏：文档名（点击改名）+ 保存状态；始终渲染，不随 branding 隐藏 */}
      <BottomBar />
      {/* 确认弹窗宿主（新建文档等二次确认经 useUiStore.requestConfirm 发起） */}
      <ConfirmDialogHost />
      <ShareDialog open={shareOpen} doc={shareDoc} onClose={closeShare} />
      <ShortcutsDialog open={shortcutsOpen} onClose={() => setShortcutsOpen(false)} />
    </div>
  );
}
