// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// 公式栏依赖 Univer 实例，用受控假 API 替代（单例保证测试与组件读到同一批 mock）
const state = {
  notation: 'A1',
  value: 123 as string | number | null,
  formula: null as string | null,
};
const rangeMock = {
  getA1Notation: () => state.notation,
  getValue: () => state.value,
  getFormula: () => state.formula,
  setValue: vi.fn((v: string | number) => {
    state.value = v;
    state.formula = null;
  }),
  setFormula: vi.fn((f: string) => {
    state.formula = f;
  }),
  activate: vi.fn(),
};
const addEventMock = vi.fn((_name: string, _fn: () => void) => ({ dispose: vi.fn() }));
const fakeApi = {
  Event: { SelectionChanged: 'SelectionChanged' },
  addEvent: addEventMock,
  getActiveWorkbook: () => ({
    getActiveRange: () => rangeMock,
    getActiveSheet: () => ({ getSheetId: () => 'sheet-1' }),
    getId: () => 'unit-1',
    getSheetByName: () => null,
  }),
};

const descMock = {
  getSearchListByNameFirstLetter: vi.fn((text: string) => [
    { name: 'SUM', desc: '求和' },
    { name: 'SUBTOTAL', desc: '分类汇总' },
  ].filter((item) => item.name.startsWith(text.toUpperCase()))),
  getFunctionInfo: vi.fn((name: string) => name === 'SUM' ? {
    functionName: 'SUM',
    abstract: '返回参数之和',
    functionParameter: [
      { name: 'number1', detail: '第一个数值', require: 1, repeat: 0 },
      { name: 'number2', detail: '第二个数值', require: 0, repeat: 1 },
    ],
  } : undefined),
};
const refSelMock = { setSelections: vi.fn() };

vi.mock('@univerjs/preset-sheets-core', () => ({
  IDescriptionService: 'desc-token',
  IRefSelectionsService: 'ref-token',
}));

vi.mock('../../core/univer/controller', () => ({
  getUniverApi: vi.fn(() => fakeApi),
  getUniverInjector: vi.fn(() => ({
    get: (token: unknown) => token === 'desc-token' ? descMock : token === 'ref-token' ? refSelMock : null,
  })),
}));

import { FormulaBar, UNIVER_READY_EVENT } from './FormulaBar';

function fireSelectionChanged() {
  const listener = addEventMock.mock.calls.at(-1)?.[1] as () => void;
  act(() => {
    listener();
  });
}

const nameBox = () => screen.getByTitle('当前单元格') as HTMLInputElement;
const contentInput = () => screen.getByPlaceholderText('输入内容或公式，回车确认') as HTMLInputElement;

beforeEach(() => {
  vi.clearAllMocks();
  state.notation = 'A1';
  state.value = 123;
  state.formula = null;
  descMock.getSearchListByNameFirstLetter.mockClear();
  refSelMock.setSelections.mockClear();
});

afterEach(() => {
  cleanup();
  document.getElementById('univer-container')?.remove();
});

describe('FormulaBar 公式栏', () => {
  it('挂载后同步名称框与单元格内容', () => {
    render(<FormulaBar />);
    expect(nameBox().value).toBe('A1');
    expect(contentInput().value).toBe('123');
    expect(addEventMock).toHaveBeenCalledWith('SelectionChanged', expect.any(Function));
  });

  it('选区变化事件刷新名称框与内容回显', () => {
    render(<FormulaBar />);
    state.notation = 'B2';
    state.value = 'hello';
    fireSelectionChanged();
    expect(nameBox().value).toBe('B2');
    expect(contentInput().value).toBe('hello');
  });

  it('公式单元格优先回显公式文本（自动补 =）', () => {
    state.formula = 'SUM(A1:A2)';
    render(<FormulaBar />);
    expect(contentInput().value).toBe('=SUM(A1:A2)');
  });

  it('Univer 实例重建（ready 事件）后重新订阅并同步', () => {
    render(<FormulaBar />);
    const callsBefore = addEventMock.mock.calls.length;
    window.dispatchEvent(new Event(UNIVER_READY_EVENT));
    expect(addEventMock.mock.calls.length).toBeGreaterThan(callsBefore);
    expect(nameBox().value).toBe('A1');
  });

  it('回车提交文本写入字符串', () => {
    render(<FormulaBar />);
    fireEvent.change(contentInput(), { target: { value: '销售数据' } });
    fireEvent.keyDown(contentInput(), { key: 'Enter' });
    expect(rangeMock.setValue).toHaveBeenCalledWith('销售数据');
    expect(rangeMock.setFormula).not.toHaveBeenCalled();
  });

  it('回车提交纯数字写入数值类型', () => {
    render(<FormulaBar />);
    fireEvent.change(contentInput(), { target: { value: ' 42 ' } });
    fireEvent.keyDown(contentInput(), { key: 'Enter' });
    expect(rangeMock.setValue).toHaveBeenCalledWith(42);
  });

  it('回车提交 = 开头内容写入公式', () => {
    render(<FormulaBar />);
    fireEvent.change(contentInput(), { target: { value: '=SUM(A1:A2)' } });
    fireEvent.keyDown(contentInput(), { key: 'Enter' });
    expect(rangeMock.setFormula).toHaveBeenCalledWith('=SUM(A1:A2)');
    expect(rangeMock.setValue).not.toHaveBeenCalled();
  });

  it('回车提交空内容清空单元格', () => {
    render(<FormulaBar />);
    fireEvent.change(contentInput(), { target: { value: '' } });
    fireEvent.keyDown(contentInput(), { key: 'Enter' });
    expect(rangeMock.setValue).toHaveBeenCalledWith('');
  });

  it('Escape 仅退出编辑不提交', () => {
    render(<FormulaBar />);
    fireEvent.change(contentInput(), { target: { value: '草稿' } });
    fireEvent.keyDown(contentInput(), { key: 'Escape' });
    expect(rangeMock.setValue).not.toHaveBeenCalled();
    expect(rangeMock.setFormula).not.toHaveBeenCalled();
  });

  it('输入函数前缀显示补全列表，回车补全并进入参数提示', async () => {
    render(<FormulaBar />);
    const input = contentInput();
    fireEvent.focus(input);
    fireEvent.change(input, { target: { value: '=SUM' } });
    await waitFor(() => expect(screen.getByRole('listbox')).toBeTruthy());
    expect(descMock.getSearchListByNameFirstLetter).toHaveBeenCalledWith('SUM');
    expect(screen.getByText('求和')).toBeTruthy();

    fireEvent.keyDown(input, { key: 'Enter' });
    expect(input.value).toBe('=SUM(');
    await waitFor(() => expect(screen.getByText('number1')).toBeTruthy());
  });

  it('上下键切换候选后回车补全', async () => {
    render(<FormulaBar />);
    const input = contentInput();
    fireEvent.focus(input);
    fireEvent.change(input, { target: { value: '=SU' } });
    await waitFor(() => expect(screen.getByRole('listbox')).toBeTruthy());
    // 字母序默认选中 SUBTOTAL，下键切到 SUM
    expect(screen.getByRole('option', { name: /SUBTOTAL/ }).getAttribute('aria-selected')).toBe('true');
    fireEvent.keyDown(input, { key: 'ArrowDown' });
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(input.value).toBe('=SUM(');
  });

  it('公式中的引用会请求高亮', async () => {
    render(<FormulaBar />);
    fireEvent.focus(contentInput());
    fireEvent.change(contentInput(), { target: { value: '=SUM(A1:B2)' } });
    await waitFor(() => expect(refSelMock.setSelections).toHaveBeenCalled());
    const selections = refSelMock.setSelections.mock.calls.at(-1)?.[0] as Array<{
      range: Record<string, number>;
      style: Record<string, unknown>;
    }>;
    expect(selections).toHaveLength(1);
    expect(selections[0].range).toMatchObject({ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });
    expect(selections[0].style).toMatchObject({ stroke: '#f97316', strokeDash: 8 });
  });

  it('编辑中点选单元格把引用插入草稿，提交写回原单元格', async () => {
    render(<FormulaBar />);
    const input = contentInput();
    fireEvent.focus(input);
    fireEvent.change(input, { target: { value: '=SUM(' } });

    const grid = document.createElement('div');
    grid.id = 'univer-container';
    document.body.appendChild(grid);
    fireEvent.mouseDown(grid);

    state.notation = 'A1:B1';
    fireSelectionChanged();
    expect(input.value).toBe('=SUM(A1:B1');

    fireEvent.keyDown(input, { key: 'Enter' });
    expect(rangeMock.setFormula).toHaveBeenCalledWith('=SUM(A1:B1');
    expect(rangeMock.activate).toHaveBeenCalled();
  });

  it('点击表格未触发 SelectionChanged 时，松手补拾取且 blur 不丢草稿', async () => {
    render(<FormulaBar />);
    const input = contentInput();
    fireEvent.focus(input);
    fireEvent.change(input, { target: { value: '=SUM(' } });

    const grid = document.createElement('div');
    grid.id = 'univer-container';
    document.body.appendChild(grid);
    fireEvent.mouseDown(grid);
    // 真实浏览器里点击画布会先触发输入框 blur，再触发 mouseup
    fireEvent.blur(input);
    state.notation = 'B2';
    fireEvent.mouseUp(document);

    await waitFor(() => expect(input.value).toBe('=SUM(B2'));
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(input.value).toBe('=SUM(B2');
  });
});
