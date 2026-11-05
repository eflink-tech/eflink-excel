// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('./efexcel', () => ({
  exportEfexcel: vi.fn(),
  importEfexcel: vi.fn(),
}));

vi.mock('./xlsx/fileIO', () => ({
  exportXlsx: vi.fn(),
  importXlsx: vi.fn(),
  trimExt: (name: string) => name.replace(/\.(xlsx|xlsm|xls)$/i, ''),
}));

vi.mock('./xlsx/sheetImage', () => ({
  exportSheetPng: vi.fn(),
}));

vi.mock('./univer/controller', () => ({
  getActiveSheetImageData: vi.fn(),
}));

vi.mock('./filePicker', () => ({
  pickFile: vi.fn(),
}));

vi.mock('./saveService', () => ({
  currentSnapshot: vi.fn(),
  saveNow: vi.fn(),
}));

import { exportEfexcel, importEfexcel } from './efexcel';
import { exportXlsx, importXlsx } from './xlsx/fileIO';
import { exportSheetPng } from './xlsx/sheetImage';
import { getActiveSheetImageData } from './univer/controller';
import { pickFile } from './filePicker';
import { currentSnapshot, saveNow } from './saveService';
import {
  exportEfexcelAction,
  exportPngAction,
  exportXlsxAction,
  importFileAction,
  importXlsxAction,
  importXlsxMenuAction,
  newDocAction,
  saveAction,
} from './fileActions';
import { getDefaultStorage, setDefaultStorage } from '../storage/registry';
import { memoryStorage } from '../storage/memory';
import { useDocumentsStore } from '../store/documentsStore';
import { useEditorStore } from '../store/editorStore';
import { useUiStore } from '../store/uiStore';
import { createDocument, createEmptySnapshot } from '../types/spreadsheet';

const mockExportEfexcel = vi.mocked(exportEfexcel);
const mockImportEfexcel = vi.mocked(importEfexcel);
const mockExportSheetPng = vi.mocked(exportSheetPng);
const mockGetActiveSheetImageData = vi.mocked(getActiveSheetImageData);
const mockExportXlsx = vi.mocked(exportXlsx);
const mockImportXlsx = vi.mocked(importXlsx);
const mockPickFile = vi.mocked(pickFile);
const mockSaveNow = vi.mocked(saveNow);
const mockCurrentSnapshot = vi.mocked(currentSnapshot);

beforeEach(async () => {
  vi.clearAllMocks();
  mockCurrentSnapshot.mockReset();
  mockGetActiveSheetImageData.mockReset();
  mockSaveNow.mockResolvedValue(undefined);
  setDefaultStorage(memoryStorage());
  useDocumentsStore.setState({ docs: [] });
  useEditorStore.getState().openDoc({ id: 'doc-base', title: '基础文档' });
  useUiStore.setState({ toast: null, confirm: null });
});

describe('fileActions 文件动作', () => {
  it('newDocAction 确认后新建文档并进入编辑器', async () => {
    const action = newDocAction();
    expect(useUiStore.getState().confirm?.title).toBe('新建文档');
    expect(useUiStore.getState().confirm?.message).toBe('新建文档将清空当前内容，请确保已经下载备份文档！');

    useUiStore.getState().resolveConfirm(true);
    await action;

    const editor = useEditorStore.getState();
    expect(editor.docId).not.toBe('doc-base');
    expect(useDocumentsStore.getState().docs).toHaveLength(1);
  });

  it('newDocAction 取消后不新建', async () => {
    const action = newDocAction();
    useUiStore.getState().resolveConfirm(false);
    await action;

    expect(useEditorStore.getState().docId).toBe('doc-base');
    expect(useDocumentsStore.getState().docs).toHaveLength(0);
  });

  it('saveAction 成功后提示已保存', async () => {
    await saveAction();

    expect(mockSaveNow).toHaveBeenCalledTimes(1);
    expect(useUiStore.getState().toast).toBe('已保存');
  });

  it('saveAction 失败后提示保存失败', async () => {
    mockSaveNow.mockRejectedValueOnce(new Error('写库失败'));

    await saveAction();

    expect(useUiStore.getState().toast).toBe('保存失败');
  });

  it('exportEfexcelAction 未绑定快照时不触发导出', async () => {
    mockCurrentSnapshot.mockReturnValueOnce(null);

    await exportEfexcelAction();

    expect(mockExportEfexcel).not.toHaveBeenCalled();
    expect(useUiStore.getState().toast).toBeNull();
  });

  it('exportEfexcelAction 先保存再导出当前文档，成功提示已导出', async () => {
    const doc = createDocument('基础文档');
    doc.id = 'doc-base';
    await getDefaultStorage().save(doc);
    mockCurrentSnapshot.mockReturnValueOnce(createEmptySnapshot());

    await exportEfexcelAction();

    expect(mockSaveNow).toHaveBeenCalledTimes(1);
    expect(mockExportEfexcel).toHaveBeenCalledWith(expect.objectContaining({ id: 'doc-base', title: '基础文档' }));
    expect(useUiStore.getState().toast).toBe('已导出 efexcel');
  });

  it('exportEfexcelAction 文档不存在时提示导出失败', async () => {
    mockCurrentSnapshot.mockReturnValueOnce(createEmptySnapshot());

    await exportEfexcelAction();

    expect(useUiStore.getState().toast).toBe('导出失败');
  });

  it('exportPngAction 没有当前 sheet 数据时不触发导出', async () => {
    mockGetActiveSheetImageData.mockReturnValueOnce(null);

    await exportPngAction();

    expect(mockExportSheetPng).not.toHaveBeenCalled();
  });

  it('exportPngAction 导出当前 sheet 整表并提示已导出图片', async () => {
    const snapshot = createEmptySnapshot();
    const data = {
      sheet: snapshot.sheets['sheet-01'],
      displayValues: [['1']],
      rawValues: [[1]],
      rowHeights: [24],
      columnWidths: [88],
    };
    mockCurrentSnapshot.mockReturnValueOnce(snapshot);
    mockGetActiveSheetImageData.mockReturnValueOnce(data);
    useEditorStore.setState({ dirty: true });

    await exportPngAction();

    expect(mockGetActiveSheetImageData).toHaveBeenCalledWith(snapshot);
    expect(mockExportSheetPng).toHaveBeenCalledWith(data, '基础文档');
    expect(useUiStore.getState().toast).toBe('已导出图片');
    // 纯读操作：不触发保存、不改变 dirty 状态
    expect(mockSaveNow).not.toHaveBeenCalled();
    expect(useEditorStore.getState().dirty).toBe(true);
  });

  it('importFileAction 成功后新建文档打开并提示导入成功', async () => {
    const imported = createDocument('外部报表');
    imported.snapshot.sheets['sheet-01'].cellData[0] = { 0: { v: '导入数据' } };
    mockImportEfexcel.mockResolvedValueOnce(imported);

    await importFileAction(new File([''], '外部报表.efexcel'));

    expect(mockImportEfexcel).toHaveBeenCalledTimes(1);
    const editor = useEditorStore.getState();
    expect(editor.docId).not.toBe('doc-base');
    expect(editor.title).toBe('外部报表');
    const docs = useDocumentsStore.getState().docs;
    expect(docs).toHaveLength(1);
    expect(docs[0]?.title).toBe('外部报表');
    expect(useUiStore.getState().toast).toBe('导入成功');
  });

  it('importFileAction 解析失败提示导入失败且不落库', async () => {
    mockImportEfexcel.mockRejectedValueOnce(new Error('解析失败'));

    await importFileAction(new File(['bad'], '坏文件.efexcel'));

    expect(useUiStore.getState().toast).toBe('导入失败：无法解析该文件');
    expect(useDocumentsStore.getState().docs).toHaveLength(0);
    expect(useEditorStore.getState().docId).toBe('doc-base');
  });

  it('importXlsxAction 确认后替换当前文档内容并触发重载', async () => {
    const base = createDocument('基础文档');
    base.id = 'doc-base';
    await getDefaultStorage().save(base);
    const reloadBefore = useEditorStore.getState().reloadToken;
    mockImportXlsx.mockResolvedValueOnce({ snapshot: createEmptySnapshot('导入表'), charts: [], annotations: { links: [], comments: [] } });

    const action = importXlsxAction(new File(['x'], '外部报表.xlsx'));
    expect(useUiStore.getState().confirm?.title).toBe('导入 Excel');
    expect(useUiStore.getState().confirm?.message).toContain('基础文档');
    // 500KB 本地存储规则须在导入前（确认弹窗）就告知，而不是只在导入后 toast
    expect(useUiStore.getState().confirm?.message).toContain('500KB');
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(mockImportXlsx).toHaveBeenCalledTimes(1);
    const doc = await getDefaultStorage().load('doc-base');
    expect(doc?.title).toBe('外部报表');
    expect(doc?.snapshot.name).toBe('导入表');
    expect(useEditorStore.getState().reloadToken).toBe(reloadBefore + 1);
    expect(useUiStore.getState().toast).toBe('导入成功');
  });

  it('importXlsxAction 导入超过 500KB 的文档提示仅保存在本地浏览器', async () => {
    const base = createDocument('基础文档');
    base.id = 'doc-base';
    await getDefaultStorage().save(base);
    const snapshot = createEmptySnapshot('大表');
    snapshot.sheets['sheet-01'].cellData[0] = { 0: { v: 'x'.repeat(500 * 1024 + 1024) } };
    mockImportXlsx.mockResolvedValueOnce({ snapshot, charts: [], annotations: { links: [], comments: [] } });

    const action = importXlsxAction(new File(['x'], '大表.xlsx'));
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(useUiStore.getState().toast).toBe('导入成功，文档超过 500KB，将仅保存在本地浏览器');
  });

  it('importXlsxAction 解析期间显示 loading，结束后关闭', async () => {
    const base = createDocument('基础文档');
    base.id = 'doc-base';
    await getDefaultStorage().save(base);
    let finish!: (value: Awaited<ReturnType<typeof importXlsx>>) => void;
    mockImportXlsx.mockReturnValueOnce(new Promise((resolve) => { finish = resolve; }));

    const action = importXlsxAction(new File(['x'], '外部报表.xlsx'));
    useUiStore.getState().resolveConfirm(true);
    await vi.waitFor(() => expect(useUiStore.getState().importing).toBe(true));
    finish({ snapshot: createEmptySnapshot('导入表'), charts: [], annotations: { links: [], comments: [] } });
    await action;

    expect(useUiStore.getState().importing).toBe(false);
  });

  it('importXlsxAction 取消确认不解析不落盘', async () => {
    const reloadBefore = useEditorStore.getState().reloadToken;
    const action = importXlsxAction(new File(['x'], '外部报表.xlsx'));
    useUiStore.getState().resolveConfirm(false);
    await action;

    expect(mockImportXlsx).not.toHaveBeenCalled();
    expect(mockSaveNow).not.toHaveBeenCalled(); // 取消路径不触发导入前保存
    expect(useEditorStore.getState().reloadToken).toBe(reloadBefore);
  });

  it('importXlsxAction 导入前保存失败提示保存文档失败且不解析', async () => {
    const reloadBefore = useEditorStore.getState().reloadToken;
    mockSaveNow.mockRejectedValueOnce(new Error('写库失败'));
    const action = importXlsxAction(new File(['x'], '外部报表.xlsx'));
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(mockImportXlsx).not.toHaveBeenCalled();
    expect(useUiStore.getState().toast).toBe('导入失败：保存文档失败');
    expect(useEditorStore.getState().reloadToken).toBe(reloadBefore);
  });

  it('importXlsxAction 落盘失败提示保存文档失败且不触发重载', async () => {
    const reloadBefore = useEditorStore.getState().reloadToken;
    setDefaultStorage({
      ...memoryStorage(),
      updateContent: vi.fn(async () => { throw new Error('写库失败'); }),
    });
    mockImportXlsx.mockResolvedValueOnce({ snapshot: createEmptySnapshot('导入表'), charts: [], annotations: { links: [], comments: [] } });
    const action = importXlsxAction(new File(['x'], '外部报表.xlsx'));
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(useUiStore.getState().toast).toBe('导入失败：保存文档失败');
    expect(useEditorStore.getState().reloadToken).toBe(reloadBefore);
  });

  it('importXlsxAction 解析失败提示导入失败且不触发重载', async () => {
    const reloadBefore = useEditorStore.getState().reloadToken;
    mockImportXlsx.mockRejectedValueOnce(new Error('解析失败'));
    const action = importXlsxAction(new File(['x'], '坏文件.xlsx'));
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(useUiStore.getState().toast).toBe('导入失败：无法解析该文件');
    expect(useUiStore.getState().importing).toBe(false);
    expect(useEditorStore.getState().reloadToken).toBe(reloadBefore);
  });

  it('exportXlsxAction 先保存再以当前文档快照触发导出', async () => {
    const doc = createDocument('基础文档');
    doc.id = 'doc-base';
    await getDefaultStorage().save(doc);
    mockCurrentSnapshot.mockReturnValueOnce(createEmptySnapshot());

    await exportXlsxAction();

    expect(mockSaveNow).toHaveBeenCalledTimes(1);
    expect(mockExportXlsx).toHaveBeenCalledTimes(1);
    expect(mockExportXlsx.mock.calls[0]?.[1]).toBe('基础文档');
    expect(useUiStore.getState().toast).toBe('已导出 xlsx');
  });

  it('importXlsxMenuAction 未选择文件时直接返回不弹确认', async () => {
    mockPickFile.mockResolvedValueOnce(null);
    await importXlsxMenuAction();
    expect(useUiStore.getState().confirm).toBeNull();
    expect(mockImportXlsx).not.toHaveBeenCalled();
  });

  it('importXlsxMenuAction 选中文件后走导入动作', async () => {
    const file = new File(['x'], '菜单报表.xlsx');
    mockPickFile.mockResolvedValueOnce(file);
    mockImportXlsx.mockResolvedValueOnce({ snapshot: createEmptySnapshot(), charts: [], annotations: { links: [], comments: [] } });

    const action = importXlsxMenuAction();
    // 等 pickFile 的 Promise 落地、确认弹窗打开后再点确定
    await vi.waitFor(() => expect(useUiStore.getState().confirm).not.toBeNull());
    useUiStore.getState().resolveConfirm(true);
    await action;

    expect(mockImportXlsx).toHaveBeenCalledWith(file);
  });
});
