// 图表选择面板：锚定在工具栏按钮下方，样式与其他下拉菜单一致
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import * as echarts from 'echarts';
import { CHART_GROUPS, SAMPLE_CHART_MATRIX, buildChartOption, type ChartTypeSpec } from '../../core/univer/chartOptions';

const THUMB_W = 88;
const THUMB_H = 58;

/** 单个类型缩略图：用示例数据渲染 mini 图表 */
function Thumb({ typeId }: { typeId: string }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!ref.current) return;
    const chart = echarts.init(ref.current, null, { width: THUMB_W, height: THUMB_H });
    chart.setOption(buildChartOption(SAMPLE_CHART_MATRIX, typeId, true));
    return () => chart.dispose();
  }, [typeId]);
  return <div ref={ref} className="myf-chart-picker__thumb" />;
}

export interface ChartPickerPanelProps {
  onPick: (typeId: string) => void;
  onClose: () => void;
}

export function ChartPickerPanel({ onPick, onClose }: ChartPickerPanelProps) {
  const [activeGroup, setActiveGroup] = useState(CHART_GROUPS[0]!.group);
  const [alignRight, setAlignRight] = useState(false);
  const groupRefs = useRef(new Map<string, HTMLDivElement>());
  const panelRef = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    const panel = panelRef.current;
    if (!panel) return;
    const rect = panel.getBoundingClientRect();
    setAlignRight(rect.right > window.innerWidth - 8);
  }, []);

  // Esc 关闭
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  const jumpTo = (group: string) => {
    setActiveGroup(group);
    groupRefs.current.get(group)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  const renderThumb = (spec: ChartTypeSpec) => (
    <button
      key={spec.id}
      type="button"
      title={spec.label}
      onClick={() => onPick(spec.id)}
      className="myf-chart-picker__item"
    >
      <Thumb typeId={spec.id} />
      <span className="myf-chart-picker__item-label">{spec.label}</span>
    </button>
  );

  return (
    <div
      ref={panelRef}
      data-testid="chart-picker-panel"
      className={`myf-chart-picker${alignRight ? ' myf-chart-picker--align-right' : ''}`}
    >
      <div className="myf-chart-picker__body">
        {CHART_GROUPS.map(({ group, types }) => (
          <div
            key={group}
            ref={(el) => {
              if (el) groupRefs.current.set(group, el);
              else groupRefs.current.delete(group);
            }}
            className="myf-chart-picker__group"
          >
            <div className="myf-chart-picker__group-title">{group}</div>
            <div className="myf-chart-picker__grid">{types.map(renderThumb)}</div>
          </div>
        ))}
      </div>
      <div className="myf-chart-picker__nav">
        {CHART_GROUPS.map(({ group }) => (
          <button
            key={group}
            type="button"
            onClick={() => jumpTo(group)}
            className={`myf-chart-picker__nav-item${activeGroup === group ? ' myf-chart-picker__nav-item--active' : ''}`}
          >
            {group}
          </button>
        ))}
      </div>
    </div>
  );
}
