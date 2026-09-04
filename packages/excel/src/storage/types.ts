import type { DocumentMeta, SheetDocument, WorkbookSnapshot } from '../types/spreadsheet';

/**
 * 文档存储适配器：编辑器所有落库行为都经由该接口。
 * 内置 IndexedDB / 内存实现；宿主可注入自定义实现对接后端 API。
 */
export interface StorageAdapter {
  /** 新建/整体覆盖保存（导入、新建文档时用） */
  save(doc: SheetDocument): Promise<void>;
  /** 按 id 读取完整文档 */
  load(id: string): Promise<SheetDocument | undefined>;
  /** 文档元信息列表（按更新时间倒序由调用方约定） */
  list(): Promise<DocumentMeta[]>;
  delete(id: string): Promise<void>;
  /** 仅改标题，不触碰快照内容 */
  rename(id: string, title: string): Promise<void>;
  /** 编辑器保存：标题 + 快照内容一起落库 */
  updateContent(id: string, title: string, snapshot: WorkbookSnapshot): Promise<void>;
  /** 测试辅助：清空全部文档（可选实现） */
  clear?(): Promise<void>;
}
