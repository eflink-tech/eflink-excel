// 图表选择面板：锚定在工具栏按钮下方，样式与其他下拉菜单一致
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import * as echarts from 'echarts';
import { useI18n } from '@eflink-tech/office-kit';
import { CHART_GROUPS, buildChartOption, type ChartTypeSpec } from '../../core/univer/chartOptions';
import { CHART_GROUP_KEYS, CHART_TYPE_KEYS } from '../../i18n';

const THUMB_W = 88;
const THUMB_H = 58;

/** 单个类型缩略图：用示例数据渲染 mini 图表 */
function Thumb({ typeId, matrix }: { typeId: string; matrix: unknown[][] }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!ref.current) return;
    const chart = echarts.init(ref.current, null, { width: THUMB_W, height: THUMB_H });
    chart.setOption(buildChartOption(matrix, typeId, true));
    return () => chart.dispose();
  }, [typeId, matrix]);
  return <div ref={ref} className="myf-chart-picker__thumb" />;
}

export interface ChartPickerPanelProps {
  onPick: (typeId: string) => void;
  onClose: () => void;
}

export function ChartPickerPanel({ onPick, onClose }: ChartPickerPanelProps) {
  const { t, locale } = useI18n();
  const [activeGroup, setActiveGroup] = useState(CHART_GROUPS[0]!.group);
  const [alignRight, setAlignRight] = useState(false);
  const groupRefs = useRef(new Map<string, HTMLDivElement>());
  const panelRef = useRef<HTMLDivElement>(null);
  // 缩略图示例数据随语言切换（月份/系列名本地化）；依赖 locale 避免每帧重建
  const sampleMatrix = useMemo<unknown[][]>(
    () => [
      [
        t('excel.chart.sample.month'),
        t('excel.chart.sample.sales'),
        t('excel.chart.sample.profit'),
      ],
      [t('excel.chart.sample.month1'), 120, 90],
      [t('excel.chart.sample.month2'), 210, 150],
      [t('excel.chart.sample.month3'), 170, 160],
    ],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [locale],
  );

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

  const renderThumb = (spec: ChartTypeSpec) => {
    const label = t(CHART_TYPE_KEYS[spec.id] ?? spec.id);
    return (
      <button
        key={spec.id}
        type="button"
        title={label}
        onClick={() => onPick(spec.id)}
        className="myf-chart-picker__item"
      >
        <Thumb typeId={spec.id} matrix={sampleMatrix} />
        <span className="myf-chart-picker__item-label">{label}</span>
      </button>
    );
  };

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
            <div className="myf-chart-picker__group-title">{t(CHART_GROUP_KEYS[group] ?? group)}</div>
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
            {t(CHART_GROUP_KEYS[group] ?? group)}
          </button>
        ))}
      </div>
    </div>
  );
}
