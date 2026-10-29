import { describe, expect, it } from 'vitest';
import {
  applyTint,
  INDEXED_PALETTE,
  parseThemePalette,
  resolveColorHex,
  STANDARD_OFFICE_THEME,
} from './colorResolve';

// 最小 theme1.xml：clrScheme 覆盖 srgbClr / sysClr(lastClr) / sysClr(val) 三种写法
const THEME_XML = `<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F1F1F"/></a:dk2>
      <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`;

describe('colorResolve: applyTint', () => {
  it('tint 为 0/undefined 时原样返回', () => {
    expect(applyTint('4F81BD')).toBe('4F81BD');
    expect(applyTint('4F81BD', 0)).toBe('4F81BD');
  });

  it('正 tint 提亮：accent1 +0.4 → Excel 一致的 95B3D7', () => {
    expect(applyTint('4F81BD', 0.4)).toBe('95B3D7');
  });

  it('负 tint 压暗走亮度空间（非逐通道近似）', () => {
    expect(applyTint('4F81BD', -0.25)).toBe('376092');
  });

  it('接受 # 前缀与 8 位 ARGB，输出 6 位大写 RGB', () => {
    expect(applyTint('#ff0000', 0)).toBe('FF0000');
    expect(applyTint('80FF0000', 0)).toBe('FF0000');
  });

  it('灰阶（sat=0）压暗仍为灰阶', () => {
    expect(applyTint('808080', -0.5)).toBe('404040');
  });
});

describe('colorResolve: parseThemePalette', () => {
  it('解析 clrScheme 并按主题索引重排（lt/dk 互换）', () => {
    const palette = parseThemePalette(THEME_XML);
    expect(palette).toBeDefined();
    // 主题索引 0 = lt1(FFFFFF)，1 = dk1(000000)，2 = lt2(EEEEEE)，3 = dk2(1F1F1F)
    expect(palette![0]).toBe('FFFFFF');
    expect(palette![1]).toBe('000000');
    expect(palette![2]).toBe('EEEEEE');
    expect(palette![3]).toBe('1F1F1F');
    expect(palette![4]).toBe('4472C4'); // accent1
    expect(palette![10]).toBe('0563C1'); // hlink
    expect(palette).toHaveLength(12);
  });

  it('sysClr 无 lastClr 时回退 val（6 位十六进制）', () => {
    const xml = THEME_XML.replace(
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>',
      '<a:dk1><a:sysClr val="0A0A0A"/></a:dk1>',
    );
    const palette = parseThemePalette(xml);
    expect(palette![1]).toBe('0A0A0A'); // dk1 → 主题索引 1
  });

  it('无 clrScheme 或空串返回 undefined（交调用方兜底）', () => {
    expect(parseThemePalette('')).toBeUndefined();
    expect(parseThemePalette('<a:theme xmlns:a="x"><a:foo/></a:theme>')).toBeUndefined();
  });

  it('部分槽缺失时按索引回退标准 Office 主题', () => {
    const xml = `<?xml version="1.0"?><a:theme xmlns:a="x"><a:clrScheme>
      <a:accent1><a:srgbClr val="123456"/></a:accent1>
    </a:clrScheme></a:theme>`;
    const palette = parseThemePalette(xml);
    expect(palette![4]).toBe('123456'); // accent1 命中
    expect(palette![0]).toBe(STANDARD_OFFICE_THEME[0]); // 其余回退
  });
});

describe('colorResolve: resolveColorHex', () => {
  it('argb 直接归一化为小写 #rrggbb', () => {
    expect(resolveColorHex({ argb: 'FFFF0000' }, STANDARD_OFFICE_THEME)).toBe('#ff0000');
    expect(resolveColorHex({ argb: '#00FF00' }, STANDARD_OFFICE_THEME)).toBe('#00ff00');
  });

  it('theme 引用按主题板取基色并套 tint', () => {
    // 主题索引 4 = accent1 = 4F81BD（标准主题）
    expect(resolveColorHex({ theme: 4 }, STANDARD_OFFICE_THEME)).toBe('#4f81bd');
    expect(resolveColorHex({ theme: 4, tint: 0.4 }, STANDARD_OFFICE_THEME)).toBe('#95b3d7');
  });

  it('indexed 引用查索引色板', () => {
    expect(resolveColorHex({ indexed: 10 }, STANDARD_OFFICE_THEME)).toBe('#' + INDEXED_PALETTE[10].toLowerCase());
    expect(resolveColorHex({ indexed: 2 }, STANDARD_OFFICE_THEME)).toBe('#ff0000');
  });

  it('优先级 argb > theme > indexed', () => {
    expect(resolveColorHex({ argb: 'FF0000FF', theme: 4 }, STANDARD_OFFICE_THEME)).toBe('#0000ff');
    expect(resolveColorHex({ theme: 4, indexed: 2 }, STANDARD_OFFICE_THEME)).toBe('#4f81bd');
  });

  it('空引用或越界索引返回 undefined', () => {
    expect(resolveColorHex(undefined, STANDARD_OFFICE_THEME)).toBeUndefined();
    expect(resolveColorHex({}, STANDARD_OFFICE_THEME)).toBeUndefined();
    expect(resolveColorHex({ indexed: 999 }, STANDARD_OFFICE_THEME)).toBeUndefined();
  });
});
