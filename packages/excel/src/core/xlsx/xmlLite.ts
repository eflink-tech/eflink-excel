// 轻量 XML DOM（node/浏览器通用，无 DOMParser 依赖，用 saxes 建树）。
// 仅取 localName，命名空间前缀不参与匹配——OOXML 的前缀（a:/c:/xdr: 等）由各生产者自定，
// 按 localName 匹配才稳。图表解析（chartImport）与主题色解析（colorResolve）共用。
import { SaxesParser } from 'saxes';

export interface XNode {
  name: string;
  attrs: Record<string, string>;
  children: XNode[];
  text: string;
}

export function parseXml(xml: string): XNode | null {
  const parser = new SaxesParser({ fragment: false });
  let root: XNode | null = null;
  let failed = false;
  const stack: XNode[] = [];
  parser.on('error', () => {
    failed = true;
  });
  parser.on('opentag', (node) => {
    const el: XNode = { name: node.name.includes(':') ? node.name.split(':').pop()! : node.name, attrs: {}, children: [], text: '' };
    for (const [k, v] of Object.entries(node.attributes)) {
      el.attrs[k.includes(':') ? k.split(':').pop()! : k] = String(v);
    }
    const parent = stack[stack.length - 1];
    if (parent) parent.children.push(el);
    else if (!root) root = el;
    stack.push(el);
  });
  parser.on('text', (t) => {
    const top = stack[stack.length - 1];
    if (top) top.text += t;
  });
  parser.on('closetag', () => {
    stack.pop();
  });
  parser.write(xml).close();
  return failed ? null : root;
}

export function findFirst(node: XNode | null, name: string): XNode | null {
  if (!node) return null;
  if (node.name === name) return node;
  for (const child of node.children) {
    const hit = findFirst(child, name);
    if (hit) return hit;
  }
  return null;
}

export function findAll(node: XNode | null, name: string, out: XNode[] = []): XNode[] {
  if (!node) return out;
  if (node.name === name) out.push(node);
  for (const child of node.children) findAll(child, name, out);
  return out;
}

/** 取首个后代指定标签的文本 */
export function findText(node: XNode | null, name: string): string | undefined {
  return findFirst(node, name)?.text.trim() || undefined;
}

/** 取首个后代指定标签的 val 属性（OOXML 枚举值惯用属性而非元素文本，如 <c:grouping val="clustered"/>） */
export function findVal(node: XNode | null, name: string): string | undefined {
  return findFirst(node, name)?.attrs.val || undefined;
}
