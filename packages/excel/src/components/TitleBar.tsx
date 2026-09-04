/** 顶栏左段：logo + 品牌名，与工具栏同行（品牌内容由 SheetEditor 的 branding 注入） */
export function TitleBar({ logo, name }: { logo?: string; name: string }) {
  return (
    <div className="flex h-[35px] shrink-0 items-center gap-2 pl-3 pr-1">
      {logo ? <img src={logo} alt={name} className="h-7 w-7 rounded-full" /> : null}
      <span className="text-[15px] font-bold text-slate-900">{name}</span>
    </div>
  );
}
