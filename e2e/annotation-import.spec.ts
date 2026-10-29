// 批注/超链接导入端到端：外部链接、内部（跨表）链接、旧版 comments、现代 threadedComments（根+回复）
// 经 facade 落位并随快照持久化。fixture 现场构造（exceljs 造数据 + 外部链接/旧版批注，jszip 注入
// 内部链接与 threadedComments/persons 部件），不落仓库文件。
import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { waitEditor, menuClick } from './helpers';

const TC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const PERSON_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/persons';

async function buildAnnotationFixture(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('e2e');
  ws.getCell('A1').value = '标题';
  ws.getCell('A1').note = '旧版批注'; // 旧版 comments（有值单元格）
  ws.getCell('D2').value = { text: '官网', hyperlink: 'https://example.com' }; // 外部超链接
  ws.getCell('E2').value = '跳转'; // 承载内部（跨表）超链接
  const hidden = wb.addWorksheet('隐藏页');
  hidden.getCell('A1').value = 'x';
  hidden.state = 'hidden';
  const buffer = await wb.xlsx.writeBuffer();

  const zip = await JSZip.loadAsync(buffer);

  // 内部超链接：向 exceljs 已写的 <hyperlinks> 追加一条 location 链接（指向隐藏页 A1）
  let sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const internalLink = '<hyperlink ref="E2" location="隐藏页!A1"/>';
  sheet1 = /<hyperlinks>/.test(sheet1)
    ? sheet1.replace('</hyperlinks>', `${internalLink}</hyperlinks>`)
    : sheet1.replace(/(<worksheet[^>]*>)/, `$1<hyperlinks>${internalLink}</hyperlinks>`);
  zip.file('xl/worksheets/sheet1.xml', sheet1);

  // 现代 threadedComments（根 + 回复）与 persons 作者名
  zip.file(
    'xl/threadedComments/threadedComment1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><threadedComments xmlns="${TC_NS}">` +
      '<threadedComment ref="B1" id="{tc-root}" personId="{p1}" dT="2024-05-01T10:00:00Z" text="根批注" done="0"/>' +
      '<threadedComment ref="B1" id="{tc-reply}" parentId="{tc-root}" personId="{p2}" dT="2024-05-02T11:00:00Z" text="回复内容" done="0"/>' +
      '</threadedComments>',
  );
  zip.file(
    'xl/persons/person.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><persons xmlns="${PERSON_NS}">` +
      '<person authorId="{p1}" displayName="王五" userId="wang" providerId="None"/>' +
      '<person authorId="{p2}" displayName="赵六" userId="zhao" providerId="None"/>' +
      '</persons>',
  );

  // sheet1 rels 追加 threadedComment 关系（解析侧据此发现批注部件）
  const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const tcRel = '<Relationship Id="rIdTC1" Type="http://schemas.microsoft.com/office/2017/09/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>';
  const existingRels = await zip.file(relsPath)?.async('string');
  zip.file(
    relsPath,
    existingRels
      ? existingRels.replace('</Relationships>', `${tcRel}</Relationships>`)
      : `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${tcRel}</Relationships>`,
  );

  // 内容类型补登记（保持 exceljs.load 容错）
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file(
    '[Content_Types].xml',
    ct.replace(
      '</Types>',
      '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.threadedComment+xml"/>' +
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.person+xml"/>' +
        '</Types>',
    ),
  );

  return zip.generateAsync({ type: 'nodebuffer' });
}

async function importAnnotationFixture(page: Page): Promise<void> {
  const fixture = await buildAnnotationFixture();
  await menuClick(page, '文件');
  const fileChooserPromise = page.waitForEvent('filechooser');
  await page.getByRole('menuitem', { name: '导入 Excel (.xlsx)' }).click();
  const chooser = await fileChooserPromise;
  await chooser.setFiles({
    name: '批注夹具.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: fixture,
  });
  const confirm = page.getByRole('button', { name: '确定' });
  await expect(confirm).toBeVisible({ timeout: 5000 });
  await confirm.click();
  await expect(page.getByText('导入成功')).toBeVisible({ timeout: 15000 });
}

async function exportRaw(page: Page): Promise<string> {
  const downloadPromise = page.waitForEvent('download', { timeout: 20000 });
  await menuClick(page, '文件', '导出数据 (.efx.json)');
  const download = await downloadPromise;
  return (await import('node:fs')).readFileSync(await download.path(), 'utf8');
}

test.describe('导入 Excel(.xlsx) 批注与超链接', () => {
  test('外部/内部超链接与旧版/现代批注落位并随快照持久化', async ({ page }) => {
    await waitEditor(page);
    await importAnnotationFixture(page);
    // 等 facade 落位（超链接写 customRange、批注写 thread-comment 模型）+ 自动保存余量
    await page.waitForTimeout(3500);

    const raw = await exportRaw(page);
    // 外部超链接 URL 进入单元格 customRange 并持久化
    expect(raw).toContain('https://example.com');
    // 内部链接解析为目标 sheet 的 gid 载荷
    expect(raw).toContain('#gid=');
    // 旧版批注与现代批注（根 + 回复）文本均落位
    expect(raw).toContain('旧版批注');
    expect(raw).toContain('根批注');
    expect(raw).toContain('回复内容');
    // 批注经 thread-comment 资源插件随快照持久化（刷新后仍在）
    const payload = JSON.parse(raw) as { snapshot?: { resources?: { name: string }[] } };
    expect(payload.snapshot?.resources?.some((r) => r.name === 'SHEET_UNIVER_THREAD_COMMENT_PLUGIN')).toBe(true);
  });
});
