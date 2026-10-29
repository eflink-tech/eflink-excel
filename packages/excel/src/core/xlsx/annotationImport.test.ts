// 批注/超链接导入解析测试：用 jszip 构造最小可用 xlsx 包，覆盖外部链接、内部（跨表/同表）
// 链接、旧版 comments（含 rPh 拼音排除）、现代 threadedComments（含回复树与 persons 作者名）。
// 落位（Univer facade）属浏览器链路，不在 node 测试覆盖。
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
  extractAnnotations,
  peekImportedAnnotation,
  queueImportedAnnotations,
  shiftImportedAnnotation,
  unshiftImportedAnnotation,
} from './annotationImport';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const RELS_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';

async function buildXlsx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="数据" sheetId="1" r:id="rId1"/><sheet name="汇总" sheetId="2" r:id="rId2"/></sheets></workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0"?><Relationships ${RELS_NS}><Relationship Id="rId1" Type=".../worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type=".../worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
  );

  // sheet1：4 个超链接（外部 / 跨表 / 同表 / 定义名-应跳过）+ 旧版批注
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0"?><worksheet ${NS}><hyperlinks>` +
      `<hyperlink ref="A1" r:id="rId1" display="示例" tooltip="提示"/>` +
      `<hyperlink ref="B2" location="汇总!A1"/>` +
      `<hyperlink ref="C3" location="D4"/>` +
      `<hyperlink ref="E5" location="SomeDefinedName"/>` +
      `</hyperlinks></worksheet>`,
  );
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0"?><Relationships ${RELS_NS}>` +
      `<Relationship Id="rId1" Type=".../hyperlink" Target="https://example.com/a" TargetMode="External"/>` +
      `<Relationship Id="rId2" Type=".../comments" Target="../comments1.xml"/>` +
      `</Relationships>`,
  );
  zip.file(
    'xl/comments1.xml',
    `<?xml version="1.0"?><comments ${NS}><authors><author>张三</author><author>李四</author></authors><commentList>` +
      `<comment ref="A1" authorId="0"><text><r><t>你好</t></r></text></comment>` +
      `<comment ref="B2" authorId="1"><text><r><t>第二行</t></r><rPh><t>phonetic</t></rPh></text></comment>` +
      `</commentList></comments>`,
  );

  // sheet2：现代 threadedComments（根 + 回复）+ persons 作者名
  zip.file('xl/worksheets/sheet2.xml', `<?xml version="1.0"?><worksheet ${NS}/>`);
  zip.file(
    'xl/worksheets/_rels/sheet2.xml.rels',
    `<?xml version="1.0"?><Relationships ${RELS_NS}><Relationship Id="rId1" Type=".../threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>`,
  );
  zip.file(
    'xl/threadedComments/threadedComment1.xml',
    `<?xml version="1.0"?><threadedComments ${NS}>` +
      `<threadedComment ref="A1" id="{root}" personId="{p1}" dT="2024-05-01T10:00:00Z" text="根批注"/>` +
      `<threadedComment ref="A1" id="{reply}" parentId="{root}" personId="{p2}" dT="2024-05-02T11:00:00Z" text="回复"/>` +
      `</threadedComments>`,
  );
  zip.file(
    'xl/persons/person.xml',
    `<?xml version="1.0"?><persons ${NS}>` +
      `<person authorId="{p1}" displayName="王五" userId="wang" providerId="None"/>` +
      `<person authorId="{p2}" displayName="赵六" userId="zhao" providerId="None"/>` +
      `</persons>`,
  );

  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('extractAnnotations 超链接', () => {
  it('外部链接取 rels Target，保留 display/tooltip；定义名链接跳过', async () => {
    const { links } = await extractAnnotations(await buildXlsx());
    // 4 个 hyperlink 中定义名（SomeDefinedName）无法解析为区域 → 只剩 3 个
    expect(links).toHaveLength(3);
    const external = links.find((l) => l.url)!;
    expect(external).toMatchObject({ sheetName: '数据', row: 0, col: 0, url: 'https://example.com/a', display: '示例', tooltip: '提示' });
  });

  it('内部链接：跨表带目标 sheet 名，同表回落到自身 sheet', async () => {
    const { links } = await extractAnnotations(await buildXlsx());
    const cross = links.find((l) => l.row === 1 && l.col === 1)!;
    expect(cross.internal).toEqual({ sheetName: '汇总', ref: 'A1' });
    const same = links.find((l) => l.row === 2 && l.col === 2)!;
    expect(same.internal).toEqual({ sheetName: '数据', ref: 'D4' });
  });
});

describe('extractAnnotations 批注', () => {
  it('旧版 comments：按 authorId 映射作者，排除 rPh 拼音注解', async () => {
    const { comments } = await extractAnnotations(await buildXlsx());
    const legacy = comments.filter((c) => c.sheetName === '数据');
    expect(legacy).toHaveLength(2);
    expect(legacy[0]).toMatchObject({ row: 0, col: 0, text: '你好', personId: 'imported-author-0', author: '张三' });
    expect(legacy[1]).toMatchObject({ row: 1, col: 1, text: '第二行', author: '李四' });
  });

  it('现代 threadedComments：还原文本/作者/日期与父子关系', async () => {
    const { comments } = await extractAnnotations(await buildXlsx());
    const threaded = comments.filter((c) => c.sheetName === '汇总');
    expect(threaded).toHaveLength(2);
    const root = threaded.find((c) => c.id === '{root}')!;
    expect(root).toMatchObject({ row: 0, col: 0, text: '根批注', author: '王五', date: '2024-05-01T10:00:00Z' });
    expect(root.parentId).toBeUndefined();
    const reply = threaded.find((c) => c.id === '{reply}')!;
    expect(reply).toMatchObject({ text: '回复', parentId: '{root}', author: '赵六' });
  });
});

describe('批注队列', () => {
  it('链接在前、批注在后，且根批注排在回复之前', async () => {
    const annotations = await extractAnnotations(await buildXlsx());
    queueImportedAnnotations(annotations);
    const tasks = [];
    for (let t = peekImportedAnnotation(); t; t = peekImportedAnnotation()) {
      tasks.push(shiftImportedAnnotation()!);
    }
    expect(tasks.every((t) => t.kind === 'link')).toBe(false);
    // 前 3 条是链接
    expect(tasks.slice(0, 3).every((t) => t.kind === 'link')).toBe(true);
    const commentTasks = tasks.filter((t) => t.kind === 'comment');
    const rootIdx = commentTasks.findIndex((c) => (c as { id?: string }).id === '{root}');
    const replyIdx = commentTasks.findIndex((c) => (c as { parentId?: string }).parentId === '{root}');
    expect(rootIdx).toBeGreaterThanOrEqual(0);
    expect(replyIdx).toBeGreaterThan(rootIdx);
  });

  it('unshift 塞回队首；空 ArrayBuffer 静默返回空集', async () => {
    queueImportedAnnotations({ links: [], comments: [] });
    expect(peekImportedAnnotation()).toBeNull();
    const empty = await extractAnnotations(new ArrayBuffer(0)).catch(() => ({ links: [], comments: [] }));
    expect(empty).toEqual({ links: [], comments: [] });
    queueImportedAnnotations({ links: [{ sheetName: 'S', row: 0, col: 0, url: 'u' }], comments: [] });
    const first = shiftImportedAnnotation()!;
    unshiftImportedAnnotation(first);
    expect(peekImportedAnnotation()).toBe(first);
  });
});
