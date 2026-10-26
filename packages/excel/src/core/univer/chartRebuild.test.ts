// 图表联动重建纯逻辑测试：读值经注入函数模拟表格当前值，覆盖各图表类型与缺引用防御
import { describe, expect, it } from 'vitest';
import { cellToNum, cellToStr, flattenRange, rebuildExplicitFromRefs } from './chartRebuild';

/** 从 ref → 值列的静态映射构造读取函数 */
function readerOf(map: Record<string, unknown[]>): (ref: string) => unknown[] {
  return (ref) => map[ref] ?? [];
}

describe('rebuildExplicitFromRefs', () => {
  it('类别类图表：catRef + 每系列 valRef + nameRef 重建', () => {
    const explicit = rebuildExplicitFromRefs('bar-cluster', {
      catRef: 'A3:A5',
      series: [
        { nameRef: 'B1', valRef: 'B3:B5' },
        { nameRef: 'C1', valRef: 'C3:C5' },
      ],
    }, readerOf({ 'A3:A5': ['甲', '乙', '丙'], B1: ['频数'], 'B3:B5': [50, 30, 20], C1: ['占比'], 'C3:C5': [0.5, 0.3, 0.2] }));
    expect(explicit).toEqual({
      categories: ['甲', '乙', '丙'],
      series: [
        { name: '频数', data: [50, 30, 20] },
        { name: '占比', data: [0.5, 0.3, 0.2] },
      ],
    });
  });

  it('组合图：comboBars 分界柱/折线系列', () => {
    const explicit = rebuildExplicitFromRefs('combo-bar-line', {
      catRef: 'A3:A4',
      comboBars: 1,
      series: [
        { valRef: 'B3:B4' },
        { valRef: 'E3:E4' },
      ],
    }, readerOf({ 'A3:A4': ['甲', '乙'], 'B3:B4': [50, 30], 'E3:E4': [0.51, 0.82] }));
    expect(explicit).toEqual({
      categories: ['甲', '乙'],
      combo: { bars: [{ name: '系列1', data: [50, 30] }], lines: [{ name: '系列2', data: [0.51, 0.82] }] },
    });
  });

  it('散点：x/y 引用按位置配对', () => {
    const explicit = rebuildExplicitFromRefs('scatter', {
      series: [{ xRef: 'A3:A4', yRef: 'B3:B4' }],
    }, readerOf({ 'A3:A4': [2.1, 3.2], 'B3:B4': [1.1, 0.9] }));
    expect(explicit).toEqual({ scatterSeries: [{ name: '系列1', data: [[2.1, 1.1], [3.2, 0.9]] }] });
  });

  it('饼图：类别 + 首系列数值配对', () => {
    const explicit = rebuildExplicitFromRefs('pie', {
      catRef: 'A3:A4',
      series: [{ valRef: 'B3:B4' }],
    }, readerOf({ 'A3:A4': ['甲', '乙'], 'B3:B4': [50, 30] }));
    expect(explicit).toEqual({ pieData: [{ name: '甲', value: 50 }, { name: '乙', value: 30 }] });
  });

  it('公式单元格取计算结果、富文本取拼接、超链接取文本；非法值归零', () => {
    expect(cellToStr({ result: 42 })).toBe('42');
    expect(cellToStr({ p: { body: { dataStream: '不合格\r\n检查\r\n' } } })).toBe('不合格\r\n检查'); // Univer 富文本形状
    expect(cellToStr({ richText: [{ text: 'a' }] })).toBe(''); // 未知对象形状退空
    expect(cellToStr({ text: '官网', hyperlink: 'https://x' })).toBe('官网');
    expect(cellToStr(null)).toBe('');
    expect(cellToNum('12.5')).toBe(12.5);
    expect(cellToNum('abc')).toBe(0);
    expect(cellToNum(Number.NaN)).toBe(0);
  });

  it('缺必需引用返回 null（调用方保留旧图）', () => {
    const read = readerOf({});
    expect(rebuildExplicitFromRefs('bar-cluster', { catRef: 'A3:A5', series: [{ valRef: 'B3:B5' }] }, read)).toBeNull();
    expect(rebuildExplicitFromRefs('pie', { series: [{ valRef: 'B3:B5' }] }, readerOf({ 'B3:B5': [1] }))).toBeNull(); // 无 catRef
    expect(rebuildExplicitFromRefs('combo-bar-line', { catRef: 'A3:A4', series: [{ valRef: 'B3:B4' }] }, readerOf({ 'A3:A4': ['x'], 'B3:B4': [1] }))).toBeNull(); // 无折线系列
  });

  it('无引用系列沿用 fallback 旧数据（字面量/筛选残留系列保持完整）', () => {
    const explicit = rebuildExplicitFromRefs('bar-cluster', {
      catRef: 'A3:A4',
      series: [{ valRef: 'B3:B4' }, {}],
    }, readerOf({ 'A3:A4': ['甲', '乙'], 'B3:B4': [80, 30] }), {
      categories: ['甲', '乙'],
      series: [
        { name: '频数', data: [50, 30] },
        { name: '静态系列', data: [7, 8] },
      ],
    });
    expect(explicit).toEqual({
      categories: ['甲', '乙'],
      series: [
        { name: '系列1', data: [80, 30] },
        { name: '静态系列', data: [7, 8] },
      ],
    });
  });

  it('组合图折线全缺引用时沿用 fallback 折线', () => {
    const explicit = rebuildExplicitFromRefs('combo-bar-line', {
      catRef: 'A3:A4',
      comboBars: 1,
      series: [{ valRef: 'B3:B4' }, {}],
    }, readerOf({ 'A3:A4': ['甲', '乙'], 'B3:B4': [80, 30] }), {
      categories: ['甲', '乙'],
      combo: { bars: [{ name: '频数', data: [50, 30] }], lines: [{ name: '累计%', data: [0.5, 0.8] }] },
    });
    expect(explicit).toEqual({
      categories: ['甲', '乙'],
      combo: { bars: [{ name: '系列1', data: [80, 30] }], lines: [{ name: '累计%', data: [0.5, 0.8] }] },
    });
  });

  it('flattenRange 取每行首格', () => {
    expect(flattenRange([['a'], ['b']])).toEqual(['a', 'b']);
    expect(flattenRange(null)).toEqual([]);
    expect(flattenRange([['a', 'z'], ['b']])).toEqual(['a', 'b']);
  });
});
