// 图表设置面板：双击插入的图表后在右侧显示，修改设置即重建图表（原位替换图片）
import { useState, useSyncExternalStore } from 'react';
import { X } from 'lucide-react';
import { CHART_GROUPS, CHART_PALETTES, type ChartElements } from '../../core/univer/chartOptions';
import { getChartMeta, getChartMetaVersion, getRangeMatrix, subscribeChartMeta, updateChartMeta } from '../../core/univer/chartSettings';
import { useUiStore } from '../../store/uiStore';

const ELEMENT_ITEMS: { key: keyof ChartElements; label: string }[] = [
  { key: 'legend', label: '图例' },
  { key: 'gridLine', label: '网格线' },
  { key: 'axis', label: '坐标轴' },
  { key: 'trendline', label: '趋势线' },
  { key: 'chartTitle', label: '图表标题' },
  { key: 'dataLabel', label: '数据标签' },
  { key: 'axisTitle', label: '轴标题' },
];

function ElementCheckbox({ label, checked, onChange }: { label: string; checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <label className="flex cursor-pointer items-center gap-2 text-[13px] text-[#1f2329]">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} className="h-4 w-4 accent-[#3370ff]" />
      {label}
    </label>
  );
}

export function ChartSettingsPanel() {
  const drawingId = useUiStore((s) => s.chartPanelDrawingId);
  const close = useUiStore((s) => s.closeChartPanel);
  const showToast = useUiStore((s) => s.showToast);
  const [tab, setTab] = useState<'global' | 'style'>('global');
  const [rangeDraft, setRangeDraft] = useState<string | null>(null);
  // 订阅元数据版本：任何设置变更后面板立即重渲染，勾选框等控件与元数据保持同步
  useSyncExternalStore(subscribeChartMeta, getChartMetaVersion);

  const meta = drawingId ? getChartMeta(drawingId) : null;
  if (!drawingId || !meta) return null;

  const update = (patch: Parameters<typeof updateChartMeta>[1]) => {
    void updateChartMeta(drawingId, patch).then((ok) => {
      if (!ok) showToast('图表更新失败');
    });
  };

  const applyRange = () => {
    const ref = (rangeDraft ?? meta.rangeRef).trim().toUpperCase();
    if (!/^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$/.test(ref)) {
      showToast('数据范围格式应为 H1:I4');
      setRangeDraft(null);
      return;
    }
    const matrix = getRangeMatrix(ref);
    if (!matrix) {
      showToast('数据范围无效');
      setRangeDraft(null);
      return;
    }
    setRangeDraft(null);
    update({ rangeRef: ref, matrix });
  };

  const transposeMatrix = () => {
    const m = meta.matrix;
    update({ matrix: m[0]?.map((_, c) => m.map((row) => row[c])) ?? m });
  };

  return (
    <div className="absolute bottom-0 right-0 top-0 z-[60] flex w-[320px] flex-col border-l border-[#dadbdd] bg-white shadow-[-4px_0_16px_rgba(0,0,0,0.08)]">
      {/* 标题 */}
      <div className="flex shrink-0 items-center justify-between px-4 pt-4">
        <div className="text-[17px] font-semibold text-[#1f2329]">图表设置</div>
        <button type="button" aria-label="关闭" onClick={close} className="rounded p-1 text-[#646a73] hover:bg-[#f2f3f4]">
          <X size={18} />
        </button>
      </div>
      {/* 页签 */}
      <div className="mt-3 flex shrink-0 gap-6 border-b border-[#e7e9eb] px-4">
        {([['global', '全局设置'], ['style', '元素样式']] as const).map(([key, label]) => (
          <button
            key={key}
            type="button"
            onClick={() => setTab(key)}
            className={`relative pb-2 text-[13px] ${tab === key ? 'font-medium text-[#1f2329]' : 'text-[#646a73] hover:text-[#1f2329]'}`}
          >
            {label}
            {tab === key && <span className="absolute inset-x-0 -bottom-px h-[2px] rounded bg-[#3370ff]" />}
          </button>
        ))}
      </div>

      <div className="flex-1 overflow-y-auto px-4 py-4">
        {tab === 'global' ? (
          <>
            {/* 图表类型 */}
            <div className="mb-1 text-[13px] font-medium text-[#1f2329]">图表类型</div>
            <select
              value={meta.typeId}
              onChange={(e) => update({ typeId: e.target.value })}
              className="mb-5 w-full rounded-md border border-[#e7e9eb] bg-white px-2 py-1.5 text-[13px] text-[#1f2329] outline-none focus:border-[#3370ff]"
            >
              {CHART_GROUPS.map(({ group, types }) => (
                <optgroup key={group} label={group}>
                  {types.map((t) => (
                    <option key={t.id} value={t.id}>
                      {t.label}
                    </option>
                  ))}
                </optgroup>
              ))}
            </select>

            {/* 图表元素 */}
            <div className="mb-2 text-[13px] font-medium text-[#1f2329]">图表元素</div>
            <div className="mb-5 grid grid-cols-2 gap-y-3">
              {ELEMENT_ITEMS.map(({ key, label }) => (
                <ElementCheckbox
                  key={key}
                  label={label}
                  checked={meta.elements[key]}
                  onChange={(v) => update({ elements: { ...meta.elements, [key]: v } })}
                />
              ))}
            </div>

            {/* 图表标题文字（勾选标题后可编辑） */}
            {meta.elements.chartTitle && (
              <>
                <div className="mb-1 text-[13px] font-medium text-[#1f2329]">标题文字</div>
                <input
                  value={meta.title}
                  onChange={(e) => update({ title: e.target.value })}
                  placeholder="图表标题"
                  className="mb-5 w-full rounded-md border border-[#e7e9eb] px-2 py-1.5 text-[13px] outline-none focus:border-[#3370ff]"
                />
              </>
            )}

            {/* 数据范围 */}
            <div className="mb-1 text-[13px] font-medium text-[#1f2329]">数据范围</div>
            <div className="mb-2 flex items-center gap-2">
              <input
                value={rangeDraft ?? meta.rangeRef}
                onChange={(e) => setRangeDraft(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') applyRange();
                }}
                placeholder="例如 Sheet1!H1:I4"
                className="min-w-0 flex-1 rounded-md border border-[#e7e9eb] px-2 py-1.5 text-[13px] outline-none focus:border-[#3370ff]"
              />
              <button
                type="button"
                onClick={applyRange}
                className="shrink-0 rounded-md border border-[#e7e9eb] px-2 py-1.5 text-[13px] text-[#1f2329] hover:bg-[#f2f3f4]"
                title="应用数据范围"
              >
                ✓
              </button>
            </div>
            <button
              type="button"
              onClick={transposeMatrix}
              className="mb-2 w-full rounded-md border border-[#e7e9eb] px-2 py-1.5 text-left text-[13px] text-[#1f2329] hover:bg-[#f2f3f4]"
            >
              切换行列
            </button>
            <p className="text-[12px] leading-5 text-[#8f959e]">修改范围或切换行列后图表立即按新数据重绘。</p>
          </>
        ) : (
          <>
            {/* 图表颜色 */}
            <div className="mb-2 text-[13px] font-medium text-[#1f2329]">图表颜色</div>
            <div className="mb-5 flex flex-col gap-2">
              {CHART_PALETTES.map((palette, index) => (
                <button
                  key={index}
                  type="button"
                  onClick={() => update({ paletteIndex: index })}
                  className={`flex items-center gap-1 rounded-md border px-2 py-2 ${
                    meta.paletteIndex === index ? 'border-[#3370ff] ring-1 ring-[#3370ff]' : 'border-[#e7e9eb] hover:border-[#c0c4cc]'
                  }`}
                >
                  {palette.map((color) => (
                    <span key={color} className="h-5 w-5 rounded-[3px]" style={{ backgroundColor: color }} />
                  ))}
                </button>
              ))}
            </div>
            <p className="text-[12px] leading-5 text-[#8f959e]">选择配色后图表立即按新配色重绘。</p>
          </>
        )}
      </div>
    </div>
  );
}
