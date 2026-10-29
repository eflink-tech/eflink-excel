// xlsx 导入批注/超链接的落位：解析侧（annotationImport）只产出纯数据，真正写入单元格
// 依赖 Univer 挂载后的 facade（超链接/批注没有稳定的 resources 序列化器）。
// UniverContainer 每次挂载后调用 insertQueuedImportAnnotations 消费队列：
// 逐条应用，落位前校验实例仍是当前挂载的 Univer 实例（React StrictMode 双执行防护）。
import { UserManagerService, type Injector, type IUser } from '@univerjs/core';
import { getUniverApi, getUniverInjector } from './controller';
import {
  peekImportedAnnotation,
  shiftImportedAnnotation,
  unshiftImportedAnnotation,
  type AnnotationTask,
  type ImportedComment,
  type ImportedLink,
} from '../xlsx/annotationImport';

type UniverApi = NonNullable<ReturnType<typeof getUniverApi>>;
type WorkbookLike = NonNullable<ReturnType<UniverApi['getActiveWorkbook']>>;
type WorksheetLike = NonNullable<ReturnType<WorkbookLike['getSheetByName']>>;
type RangeLike = NonNullable<ReturnType<WorksheetLike['getRange']>>;

/** 批注纯文本 → Univer 文档 body（换行统一为段落标记 \r\n，末尾补一个） */
function commentBody(text: string): { dataStream: string } {
  const normalized = text.replace(/\r\n?/g, '\n').split('\n').join('\r\n');
  return { dataStream: normalized.endsWith('\r\n') ? normalized : `${normalized}\r\n` };
}

/**
 * 注册导入批注的原作者，返回可安全用于 setPersonId 的 personId。
 * facade 未开放 addUser，只能经底层 DI 容器操作 UserManagerService；
 * 拿不到容器/注册失败时返回 undefined —— 此时不设 personId，批注归到当前用户，
 * 避免出现「有 personId 但查无此人」的空白作者。
 */
function ensureAuthor(personId?: string, author?: string): string | undefined {
  if (!personId || !author) return undefined;
  const injector = getUniverInjector() as Injector | null;
  if (!injector) return undefined;
  try {
    const users = injector.get(UserManagerService);
    if (!users.getUser(personId)) users.addUser({ userID: personId, name: author, avatar: '' } as IUser);
    return personId;
  } catch {
    return undefined;
  }
}

/** 消费导入队列：把批注/超链接逐条落位。任何一条失败只告警，不影响其余 */
export function insertQueuedImportAnnotations(): void {
  const api = getUniverApi();
  const wb = api?.getActiveWorkbook();
  if (!api || !wb) return; // 实例未就绪：任务留在队列，等可用挂载再消费
  void consumeQueue(api, wb);
}

async function consumeQueue(api: UniverApi, wb: WorkbookLike): Promise<void> {
  for (;;) {
    if (getUniverApi() !== api) return;
    const task = peekImportedAnnotation();
    if (!task) return;
    shiftImportedAnnotation();
    try {
      await applyOne(api, wb, task);
    } catch (err) {
      if (getUniverApi() !== api) {
        unshiftImportedAnnotation(task); // 落位中途实例被销毁：塞回队首交给新挂载重试
        return;
      }
      console.warn('[annotationImport] 批注/超链接落位失败，已跳过', task.sheetName, err);
    }
  }
}

async function applyOne(api: UniverApi, wb: WorkbookLike, task: AnnotationTask): Promise<void> {
  const ws = wb.getSheetByName(task.sheetName);
  if (!ws) return;
  if (task.kind === 'link') await applyLink(wb, ws, task);
  else await applyComment(api, ws, task);
}

async function applyLink(wb: WorkbookLike, ws: WorksheetLike, link: ImportedLink): Promise<void> {
  const range = ws.getRange(link.row, link.col, 1, 1) as RangeLike;
  if (!range) return;
  let payload = link.url;
  if (!payload && link.internal) {
    const targetSheetName = link.internal.sheetName ?? link.sheetName;
    const gid = wb.getSheetByName(targetSheetName)?.getSheetId();
    if (!gid) return; // 目标 sheet 不在（被删/名不符）：跳过，不写成坏链接
    payload = `#gid=${gid}&range=${link.internal.ref}`;
  }
  if (!payload) return;
  const label = link.display ?? cellText(range);
  await range.setHyperLink(payload, label);
}

function cellText(range: RangeLike): string {
  const v = range.getValue?.();
  return v == null ? '' : String(v);
}

async function applyComment(api: UniverApi, ws: WorksheetLike, comment: ImportedComment): Promise<void> {
  const builder = api.newTheadComment().setContent(commentBody(comment.text) as never);
  if (comment.date) {
    const d = new Date(comment.date);
    if (!Number.isNaN(d.getTime())) builder.setDateTime(d);
  }
  const personId = ensureAuthor(comment.personId, comment.author);
  if (personId) builder.setPersonId(personId);

  // 回复：挂到已落位的根批注下（根在前、回复在后由解析侧排序保证）
  if (comment.parentId) {
    const root = ws.getCommentById?.(comment.parentId);
    if (root) {
      await root.replyAsync(builder);
      return;
    }
    // 根缺失（被跳过/孤儿回复）：退化为该单元格的独立批注
  }
  if (comment.id) builder.setId(comment.id);
  const range = ws.getRange(comment.row, comment.col, 1, 1) as RangeLike;
  if (!range) return;
  await range.addCommentAsync(builder);
}
