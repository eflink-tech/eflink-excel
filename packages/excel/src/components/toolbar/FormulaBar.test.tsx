// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
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
};
const addEventMock = vi.fn((_name: string, _fn: () => void) => ({ dispose: vi.fn() }));
const fakeApi = {
  Event: { SelectionChanged: 'SelectionChanged' },
  addEvent: addEventMock,
  getActiveWorkbook: () => ({ getActiveRange: () => rangeMock }),
};

vi.mock('../../core/univer/controller', () => ({
  getUniverApi: vi.fn(() => fakeApi),
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
});

afterEach(() => {
  cleanup();
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
});
