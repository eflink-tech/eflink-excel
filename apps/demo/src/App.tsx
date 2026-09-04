import { useEffect } from 'react';
import {
  SheetEditor,
  ToastHost,
  getDefaultStorage,
  useDocumentsStore,
  useEditorStore,
} from '@eflink-tech/excel';

export const LAST_DOC_KEY = 'eflink-excel:lastDoc';

/** 打开上次文档；没有则自动新建并直接进入编辑器（不经过列表页） */
async function openOrCreateEditor() {
  const last = localStorage.getItem(LAST_DOC_KEY);
  if (last) {
    try {
      const doc = await getDefaultStorage().load(last);
      if (doc) {
        useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
        return;
      }
    } catch (err) {
      console.error('加载上次文档失败', err);
    }
  }
  const doc = await useDocumentsStore.getState().createDoc();
  useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
}

export default function App() {
  const docId = useEditorStore((s) => s.docId);

  useEffect(() => {
    void openOrCreateEditor().catch((err) => console.error('启动编辑器失败', err));
  }, []);

  return (
    <>
      {docId ? (
        <SheetEditor
          key={docId}
          docId={docId}
          branding={{ logo: '/logo.png', name: '易飞表格' }}
          onDocLoaded={(doc) => localStorage.setItem(LAST_DOC_KEY, doc.id)}
        />
      ) : (
        <div className="flex h-full items-center justify-center text-sm text-slate-400">加载中…</div>
      )}
      <ToastHost />
    </>
  );
}
