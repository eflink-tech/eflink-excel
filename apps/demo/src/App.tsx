import { useEffect, useState } from 'react';
import {
  LAST_DOC_KEY,
  SheetEditor,
  ToastHost,
  builtinEditorBrand,
  getDefaultStorage,
  resolveEditorBrandId,
  useDocumentsStore,
  useEditorStore,
  type EditorBrandId,
  type OfficeLocale,
} from '@eflink-tech/excel';

export { LAST_DOC_KEY };

const LOCALES: { value: OfficeLocale; label: string }[] = [
  { value: 'zh-CN', label: '中文' },
  { value: 'en', label: 'English' },
  { value: 'es', label: 'Español' },
  { value: 'pt-BR', label: 'Português' },
];

const BRANDS: { value: EditorBrandId; label: string }[] = [
  { value: 'doculuma', label: 'DocuLuma' },
  { value: 'eflink', label: 'Eflink' },
];

const LOADING_TEXT: Record<OfficeLocale, string> = {
  'zh-CN': '加载中…',
  en: 'Loading…',
  es: 'Cargando…',
  'pt-BR': 'Carregando…',
};

const UNTITLED_NAME: Record<OfficeLocale, string> = {
  'zh-CN': '未命名表格',
  en: 'Untitled spreadsheet',
  es: 'Hoja de cálculo sin título',
  'pt-BR': 'Planilha sem título',
};

function readLocale(): OfficeLocale {
  const value = new URLSearchParams(window.location.search).get('locale');
  return LOCALES.find((item) => item.value === value)?.value ?? 'zh-CN';
}

/** ?brand=eflink 直达指定品牌，便于逐品牌验收 */
function readBrand(): EditorBrandId {
  return resolveEditorBrandId(new URLSearchParams(window.location.search).get('brand'));
}

function LanguageSwitcher({
  locale,
  onChange,
}: {
  locale: OfficeLocale;
  onChange: (locale: OfficeLocale) => void;
}) {
  return (
    <select
      value={locale}
      onChange={(e) => onChange(e.target.value as OfficeLocale)}
      aria-label="Language"
      className="fixed bottom-9 right-3 z-[100000] rounded-md border border-black/10 bg-white px-2 py-1 text-xs text-[#1f2329] shadow-md outline-none"
    >
      {LOCALES.map((item) => (
        <option key={item.value} value={item.value}>
          {item.label}
        </option>
      ))}
    </select>
  );
}

/** 打开上次文档；没有则自动新建并直接进入编辑器（不经过列表页） */
async function openOrCreateEditor(locale: OfficeLocale) {
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
  const doc = await useDocumentsStore.getState().createDoc(UNTITLED_NAME[locale]);
  useEditorStore.getState().openDoc({ id: doc.id, title: doc.title });
}

export default function App() {
  const docId = useEditorStore((s) => s.docId);
  const [locale, setLocale] = useState<OfficeLocale>(readLocale);
  const [brand, setBrand] = useState<EditorBrandId>(readBrand);

  useEffect(() => {
    document.title = builtinEditorBrand(brand).name;
  }, [brand]);

  useEffect(() => {
    void openOrCreateEditor(readLocale()).catch((err) => console.error('启动编辑器失败', err));
  }, []);

  const handleLocaleChange = (next: OfficeLocale) => {
    const url = new URL(window.location.href);
    url.searchParams.set('locale', next);
    window.history.replaceState(null, '', url);
    setLocale(next);
  };

  const handleBrandChange = (next: EditorBrandId) => {
    const url = new URL(window.location.href);
    url.searchParams.set('brand', next);
    window.history.replaceState(null, '', url);
    setBrand(next);
  };

  return (
    <>
      {docId ? (
        <SheetEditor
          key={docId}
          docId={docId}
          locale={locale}
          branding={brand}
          onDocLoaded={(doc) => localStorage.setItem(LAST_DOC_KEY, doc.id)}
        />
      ) : (
        <div className="flex h-full items-center justify-center text-sm text-slate-400">
          {LOADING_TEXT[locale]}
        </div>
      )}
      <select
        value={brand}
        aria-label="Brand"
        onChange={(e) => handleBrandChange(e.target.value as EditorBrandId)}
        className="fixed bottom-3 right-3 z-[100000] rounded-md border border-black/10 bg-white px-2 py-1 text-xs text-[#1f2329] shadow-md outline-none"
      >
        {BRANDS.map((item) => (
          <option key={item.value} value={item.value}>
            {item.label}
          </option>
        ))}
      </select>
      <LanguageSwitcher locale={locale} onChange={handleLocaleChange} />
      <ToastHost />
    </>
  );
}
