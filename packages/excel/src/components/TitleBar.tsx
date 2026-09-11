// 顶栏左段：返回按钮（宿主注入时）+ logo + 品牌名 + 保存状态，与工具栏同行（品牌内容由 SheetEditor 的 branding 注入）
import { ArrowLeft } from 'lucide-react';
import { getEditorBackHref } from '../core/chrome';
import { useEditorStore } from '../store/editorStore';

/** 保存状态指示：有未保存改动时红点提示，云端保存成功后显示已保存 */
function SaveStatus() {
  const dirty = useEditorStore((s) => s.dirty);
  if (dirty) {
    return (
      <span className="flex shrink-0 items-center gap-1 text-xs text-[#e02e2e]" title="有未保存的修改，按 ⌘S / Ctrl+S 保存">
        <span className="text-[10px] leading-none">●</span>
        未保存
      </span>
    );
  }
  return (
    <span className="flex shrink-0 items-center gap-1 text-xs text-slate-400">
      <span className="text-[10px] leading-none">✓</span>
      已保存
    </span>
  );
}

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
      <SaveStatus />
    </div>
  );
}
