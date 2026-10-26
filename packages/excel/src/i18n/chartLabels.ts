/** 图表注册表（chartOptions.ts）使用中文标签作为 id/分组键，这里映射到四语文案 */
export const CHART_GROUP_KEYS: Record<string, string> = {
  柱状图: 'excel.chart.group.bar',
  折线图: 'excel.chart.group.line',
  饼图: 'excel.chart.group.pie',
  条形图: 'excel.chart.group.barHorizontal',
  面积图: 'excel.chart.group.area',
  散点图: 'excel.chart.group.scatter',
  雷达图: 'excel.chart.group.radar',
};

export const CHART_TYPE_KEYS: Record<string, string> = {
  'bar-cluster': 'excel.chart.type.barCluster',
  'bar-stack': 'excel.chart.type.barStack',
  'bar-percent': 'excel.chart.type.barPercent',
  line: 'excel.chart.type.line',
  'line-stack': 'excel.chart.type.lineStack',
  'line-marker': 'excel.chart.type.lineMarker',
  pie: 'excel.chart.type.pie',
  'pie-doughnut': 'excel.chart.type.pieDoughnut',
  'bar-horizontal': 'excel.chart.type.barHorizontal',
  'bar-horizontal-stack': 'excel.chart.type.barHorizontalStack',
  area: 'excel.chart.type.area',
  'area-stack': 'excel.chart.type.areaStack',
  scatter: 'excel.chart.type.scatter',
  radar: 'excel.chart.type.radar',
};
