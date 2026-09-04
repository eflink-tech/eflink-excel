// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('./xlsx/fileIO', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./xlsx/fileIO')>();
  return {
    ...actual,
    downloadBlob: vi.fn(),
  };
});

import { downloadBlob } from './xlsx/fileIO';
import { exportEfexcel, importEfexcel } from './efexcel';
import { createDocument } from '../types/spreadsheet';

const mockDownload = vi.mocked(downloadBlob);

beforeEach(() => {
  vi.clearAllMocks();
});

describe('efexcel 文档备份格式', () => {
  it('exportEfexcel 以标题命名下载 JSON 文件，协议含 source 来源信息', async () => {
    const doc = createDocument('季度报表');

    exportEfexcel(doc);

    expect(mockDownload).toHaveBeenCalledTimes(1);
    const [blob, filename] = mockDownload.mock.calls[0];
    expect(filename).toBe('季度报表.efexcel');
    expect(blob).toBeInstanceOf(Blob);
    const payload = JSON.parse(await blob!.text());
    expect(payload.source).toEqual({ app: '易飞表格', url: 'https://eflink.tech/office/excel' });
  });

  it('exportEfexcel 支持自定义来源信息', async () => {
    const doc = createDocument('外部应用表格');

    exportEfexcel(doc, { app: '自定义应用', url: 'https://example.com/sheets' });

    const [blob] = mockDownload.mock.calls[0];
    const payload = JSON.parse(await blob!.text());
    expect(payload.source).toEqual({ app: '自定义应用', url: 'https://example.com/sheets' });
  });

  it('importEfexcel 解析完整文档并生成全新 id', async () => {
    const source = createDocument('外部表格');
    source.snapshot.sheets['sheet-01'].cellData[0] = { 0: { v: '数据' } };
    const file = new File([JSON.stringify(source)], '外部表格.efexcel', { type: 'application/json' });

    const imported = await importEfexcel(file);

    expect(imported.id).not.toBe(source.id);
    expect(imported.title).toBe('外部表格');
    expect(imported.snapshot.sheets['sheet-01'].cellData[0]?.[0]?.v).toBe('数据');
  });

  it('importEfexcel 标题缺省时回退为导入的表格', async () => {
    const source = createDocument();
    delete (source as Partial<typeof source>).title;
    const file = new File([JSON.stringify(source)], '未命名.efexcel');

    const imported = await importEfexcel(file);

    expect(imported.title).toBe('导入的表格');
  });

  it('importEfexcel 非法 JSON 抛错', async () => {
    const file = new File(['这不是 JSON'], '坏文件.efexcel');

    await expect(importEfexcel(file)).rejects.toThrow('文件不是合法 JSON');
  });

  it('importEfexcel 缺少快照结构抛错', async () => {
    const file = new File([JSON.stringify({ title: '空文档' })], '空文档.efexcel');

    await expect(importEfexcel(file)).rejects.toThrow('缺少有效的表格快照');
  });
});
