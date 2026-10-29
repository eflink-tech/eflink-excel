// xlsx 批注/超链接导入（解析侧）：exceljs 对两者的支持都不完整——
// 批注只在有值的单元格上回填 note、且完全不读现代 threadedComments；超链接只在
// {text,hyperlink} 对象值形态下才往返。这里像图表一样直接读 zip 里的原始 XML，
// 把每个 sheet 的超链接与批注（含旧版 comments 与现代 threadedComments）如实抽出。
// 落位需要 Univer 挂载后的 facade（无稳定的 resources 序列化器），见 univer/importedAnnotations.ts。
// 本模块保持纯逻辑（不依赖 DOM/Univer），供 node 环境单测。
import JSZip from 'jszip';
import { findAll, findFirst, parseXml, type XNode } from './xmlLite';
import { decodeCellRef, dirname, listSheets, readRels, resolvePath } from './zipNav';

/** 内部（工作簿内）超链接目标：跳到某 sheet 的某个区域 */
export interface InternalLinkTarget {
  /** 目标 sheet 名；缺省表示与链接所在 sheet 相同 */
  sheetName?: string;
  /** A1 记法区域（可带 $ 与 :范围），交给 Univer deserializeRangeWithSheet 解析 */
  ref: string;
}

export interface ImportedLink {
  sheetName: string;
  row: number;
  col: number;
  /** 外部 URL（http/mailto/相对路径等）；与 internal 二选一 */
  url?: string;
  /** 内部链接目标；与 url 二选一 */
  internal?: InternalLinkTarget;
  /** xlsx display 属性（缺省时落位用单元格现有文本） */
  display?: string;
  tooltip?: string;
}

export interface ImportedComment {
  sheetName: string;
  row: number;
  col: number;
  text: string;
  /** 现代 threadedComments 的 id / parentId（用于还原回复树）；旧版批注无 id */
  id?: string;
  parentId?: string;
  /** 作者稳定标识（用于注册 Univer 用户，使批注显示原作者名） */
  personId?: string;
  /** 作者显示名 */
  author?: string;
  /** ISO 日期时间（threadedComments 的 dT） */
  date?: string;
}

export interface ImportedAnnotations {
  links: ImportedLink[];
  comments: ImportedComment[];
}

// ---------- 超链接 ----------

/** location="Sheet2!A1" / "'My Sheet'!A1:B2" / "A1" → 内部目标；无法识别（如定义名）返回 null */
function parseLocation(location: string, selfSheetName: string): InternalLinkTarget | null {
  const bang = location.indexOf('!');
  if (bang === -1) {
    // 无 sheet 前缀：同 sheet 的 A1 引用才认；定义名无法可靠解析为区域，跳过
    return /^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(location) ? { sheetName: selfSheetName, ref: location } : null;
  }
  let sheetName = location.slice(0, bang);
  const ref = location.slice(bang + 1);
  if (sheetName.startsWith("'") && sheetName.endsWith("'")) sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
  if (!ref || !/^[A-Za-z$]+\d+(:[A-Za-z$]+\d+)?$/.test(ref)) return null;
  return { sheetName, ref };
}

function readHyperlinks(sheetXml: XNode | null, sheetRels: Record<string, string>, sheetName: string): ImportedLink[] {
  const out: ImportedLink[] = [];
  for (const el of findAll(sheetXml, 'hyperlink')) {
    const ref = el.attrs.ref;
    const cell = ref ? decodeCellRef(ref) : null;
    if (!cell) continue;
    const relId = el.attrs.id;
    const external = relId ? sheetRels[relId] : undefined;
    const location = el.attrs.location;
    let url: string | undefined;
    let internal: InternalLinkTarget | undefined;
    if (external) {
      // 外部关系 Target 即 URL（TargetMode=External），原样保留，不做路径归一化
      url = external;
    } else if (location) {
      internal = parseLocation(location, sheetName) ?? undefined;
    }
    if (!url && !internal) continue;
    out.push({
      sheetName,
      row: cell.row,
      col: cell.col,
      ...(url ? { url } : {}),
      ...(internal ? { internal } : {}),
      ...(el.attrs.display ? { display: el.attrs.display } : {}),
      ...(el.attrs.tooltip ? { tooltip: el.attrs.tooltip } : {}),
    });
  }
  return out;
}

// ---------- 批注 ----------

/** 旧版 commentsN.xml：<comment ref authorId><text><r><t>…</t></r></text></comment> */
function readLegacyComments(root: XNode | null, sheetName: string): ImportedComment[] {
  if (!root) return [];
  const authors = findAll(findFirst(root, 'authors'), 'author').map((a) => a.text);
  const out: ImportedComment[] = [];
  for (const c of findAll(root, 'comment')) {
    const cell = decodeCellRef(c.attrs.ref ?? '');
    if (!cell) continue;
    const text = legacyCommentText(c);
    if (!text) continue;
    const authorId = c.attrs.authorId;
    const author = authorId != null ? authors[Number(authorId)] : undefined;
    out.push({
      sheetName,
      row: cell.row,
      col: cell.col,
      text,
      ...(author ? { personId: `imported-author-${authorId}`, author } : {}),
    });
  }
  return out;
}

/** 取 <text> 下各 <r> 的 <t>（排除 rPh 拼音注解）；无 run 时取直接 <t> */
function legacyCommentText(comment: XNode): string {
  const textEl = findFirst(comment, 'text');
  if (!textEl) return '';
  const runs = textEl.children.filter((ch) => ch.name === 'r');
  if (runs.length) return runs.map((r) => findFirst(r, 't')?.text ?? '').join('');
  return findFirst(textEl, 't')?.text ?? '';
}

/** 现代 threadedCommentN.xml：<threadedComment ref id parentId personId dT text/> */
function readThreadedComments(root: XNode | null, persons: Record<string, string>, sheetName: string): ImportedComment[] {
  if (!root) return [];
  const out: ImportedComment[] = [];
  for (const c of findAll(root, 'threadedComment')) {
    const cell = decodeCellRef(c.attrs.ref ?? '');
    if (!cell) continue;
    const text = c.attrs.text ?? '';
    if (!text) continue;
    const personId = c.attrs.personId;
    out.push({
      sheetName,
      row: cell.row,
      col: cell.col,
      text,
      ...(c.attrs.id ? { id: c.attrs.id } : {}),
      ...(c.attrs.parentId ? { parentId: c.attrs.parentId } : {}),
      ...(personId ? { personId, author: persons[personId] } : {}),
      ...(c.attrs.dT ? { date: c.attrs.dT } : {}),
    });
  }
  return out;
}

/** xl/persons/person.xml：authorId → displayName */
async function readPersons(zip: JSZip): Promise<Record<string, string>> {
  const root = parseXml((await zip.file('xl/persons/person.xml')?.async('string')) ?? '');
  const map: Record<string, string> = {};
  for (const p of findAll(root, 'person')) {
    const id = p.attrs.authorId ?? p.attrs.id;
    if (id) map[id] = p.attrs.displayName ?? p.attrs.name ?? '';
  }
  return map;
}

/**
 * 从 xlsx 原始字节提取全部超链接与批注。任何一环缺失都静默跳过，绝不影响单元格导入。
 */
export async function extractAnnotations(buffer: ArrayBuffer): Promise<ImportedAnnotations> {
  const zip = await JSZip.loadAsync(buffer);
  const links: ImportedLink[] = [];
  const comments: ImportedComment[] = [];
  let persons: Record<string, string> | null = null;

  for (const { sheetName, sheetPath, sheetRelsPath } of await listSheets(zip)) {
    const sheetRels = await readRels(zip, sheetRelsPath);
    const sheetXml = parseXml((await zip.file(sheetPath)?.async('string')) ?? '');
    links.push(...readHyperlinks(sheetXml, sheetRels, sheetName));

    for (const target of Object.values(sheetRels)) {
      const partPath = resolvePath(dirname(sheetPath), target);
      const partFile = zip.file(partPath);
      if (!partFile) continue;
      if (/threadedComments\/threadedComment\d+\.xml$/.test(partPath)) {
        if (!persons) persons = await readPersons(zip);
        const root = parseXml(await partFile.async('string'));
        comments.push(...readThreadedComments(root, persons, sheetName));
      } else if (/(^|\/)comments\d*\.xml$/.test(partPath)) {
        const root = parseXml(await partFile.async('string'));
        comments.push(...readLegacyComments(root, sheetName));
      }
    }
  }
  return { links, comments };
}

// ---------- 待落位队列（编辑器挂载后由 univer/importedAnnotations.ts 消费） ----------

/** 单个落位任务：链接或批注（批注按根在前、回复在后排序，便于还原回复树） */
export type AnnotationTask = ({ kind: 'link' } & ImportedLink) | ({ kind: 'comment' } & ImportedComment);

/** 稳定排序批注：父在前子在后；孤儿回复（父不在集合内）当根处理 */
function orderComments(comments: ImportedComment[]): ImportedComment[] {
  const out: ImportedComment[] = [];
  const remaining = [...comments];
  const emitted = new Set<string>();
  const keyOf = (c: ImportedComment) => c.id ?? `${c.sheetName}:${c.row}:${c.col}`;
  while (remaining.length) {
    let progress = false;
    for (let i = 0; i < remaining.length; i++) {
      const c = remaining[i]!;
      if (!c.parentId || emitted.has(c.parentId)) {
        out.push(c);
        emitted.add(keyOf(c));
        remaining.splice(i, 1);
        i--;
        progress = true;
      }
    }
    if (!progress) {
      out.push(...remaining);
      break;
    }
  }
  return out;
}

let pendingAnnotations: AnnotationTask[] = [];

/** 解析出的批注/超链接入队（导入动作调用；清空旧队列防串档） */
export function queueImportedAnnotations(annotations: ImportedAnnotations): void {
  pendingAnnotations = [
    ...annotations.links.map((l) => ({ kind: 'link' as const, ...l })),
    ...orderComments(annotations.comments).map((c) => ({ kind: 'comment' as const, ...c })),
  ];
}

/** 窥视队首（不移除）；配合 shift/unshift 逐条消费，实例换代时剩余任务整体留在队列 */
export function peekImportedAnnotation(): AnnotationTask | null {
  return pendingAnnotations[0] ?? null;
}

/** 取走队首一条 */
export function shiftImportedAnnotation(): AnnotationTask | null {
  return pendingAnnotations.shift() ?? null;
}

/** 塞回队首：落位中途实例被销毁时退回，等下一次挂载的消费者接手 */
export function unshiftImportedAnnotation(task: AnnotationTask): void {
  pendingAnnotations.unshift(task);
}
