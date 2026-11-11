import { describe, expect, it } from 'vitest';
import {
  applyFunctionName,
  buildSignature,
  extractRefs,
  filterSuggestions,
  functionContextAtCaret,
  parseA1Body,
  readTokenAtCaret,
} from './formulaAssist';

describe('readTokenAtCaret', () => {
  it('读取光标前的函数名前缀，忽略字符串内文本', () => {
    expect(readTokenAtCaret('=SU', 3)).toMatchObject({ token: 'SU', start: 1, inString: false });
    expect(readTokenAtCaret('="SU"', 4).inString).toBe(true);
    expect(readTokenAtCaret('=SUM(A1', 7).token).toBe('A1');
  });
});

describe('functionContextAtCaret', () => {
  it('定位最内层函数与当前参数序号', () => {
    expect(functionContextAtCaret('=SUM(A1', 7)).toEqual({ name: 'SUM', argIndex: 0 });
    expect(functionContextAtCaret('=SUM(A1,B2,', 12)).toEqual({ name: 'SUM', argIndex: 2 });
    expect(functionContextAtCaret('=IF(A1>0,SUM(B1:B2),', 20)).toEqual({ name: 'IF', argIndex: 2 });
    expect(functionContextAtCaret('="SUM("', 7)).toBeNull();
  });
});

describe('applyFunctionName', () => {
  it('补全函数名并补左括号，已存在括号时不重复', () => {
    expect(applyFunctionName('=SU', 3, 'SUM')).toEqual({ text: '=SUM(', caret: 5 });
    expect(applyFunctionName('=SU(A1)', 3, 'SUM')).toEqual({ text: '=SUM(A1)', caret: 4 });
  });
});

describe('filterSuggestions', () => {
  it('前缀匹配、完全匹配优先、最多 10 条', () => {
    const list = [
      { name: 'SUMIF', desc: '' },
      { name: 'SUM', desc: '' },
      { name: 'AVERAGE', desc: '' },
    ];
    expect(filterSuggestions(list, 'sum').map((item) => item.name)).toEqual(['SUM', 'SUMIF']);
  });
});

describe('buildSignature', () => {
  it('可选参数加方括号并高亮当前参数', () => {
    const params = buildSignature([
      { name: 'number1', require: true },
      { name: 'number2', require: false },
    ], 1);
    expect(params.map((param) => param.label)).toEqual(['number1', '[number2]']);
    expect(params[1].active).toBe(true);
  });
});

describe('parseA1Body / extractRefs', () => {
  it('解析单元格、区域、整列、整行', () => {
    expect(parseA1Body('A1')).toEqual({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 });
    expect(parseA1Body('$B$2:$C$4')).toEqual({ startRow: 1, startColumn: 1, endRow: 3, endColumn: 2 });
    expect(parseA1Body('A:C')).toMatchObject({ startRow: 0, startColumn: 0, endColumn: 2, endRow: 1048575 });
    expect(parseA1Body('1:3')).toMatchObject({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 16383 });
    expect(parseA1Body('XFE1')).toBeNull();
  });

  it('提取公式引用，跳过字符串、函数名与非法列', () => {
    expect(extractRefs('=SUM(A1:B2,Sheet2!C3)').map((ref) => ref.text)).toEqual(['A1:B2', 'Sheet2!C3']);
    expect(extractRefs('=IF(A1,"A1",LOG10(B1))').map((ref) => ref.text)).toEqual(['A1', 'B1']);
    expect(extractRefs('=BIN2DEC(101)')).toEqual([]);
    expect(extractRefs('=SUM(A:C,1:2)').map((ref) => ref.text)).toEqual(['A:C', '1:2']);
  });
});
