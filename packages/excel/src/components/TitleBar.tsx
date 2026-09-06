// 顶栏左段：返回按钮（宿主注入时）+ logo + 品牌名，与工具栏同行（品牌内容由 SheetEditor 的 branding 注入）
import { ArrowLeft } from 'lucide-react';
import { getEditorBackHref } from '../core/chrome';

export function TitleBar({ logo, name }: { logo?: string; name: string }) {
  const backHref = getEditorBackHref();
  return (
    <div className="flex h-[35px] shrink-0 items-center gap-2 pl-3 pr-1">
      {backHref ? (
        <a
          href={backHref}
          title="返回"
          className="flex size-7 shrink-0 items-center justify-center rounded-md text-slate-500 transition-colors hover:bg-slate-100 hover:text-slate-900"
        >
          <ArrowLeft size={18} />
        </a>
      ) : null}
      {logo ? <img src={logo} alt={name} className="h-7 w-7 rounded-full" /> : null}
      <span className="text-[15px] font-bold text-slate-900">{name}</span>
    </div>
  );
}
