/** 宿主编辑器外观注入：返回链接。设置后顶栏 logo 前渲染返回按钮。 */

let backHref: string | null = null;

/** 注入顶栏返回按钮的跳转地址；传 null 移除按钮 */
export function setEditorBackHref(href: string | null): void {
  backHref = href;
}

/** 读取当前返回链接（未设置返回 null） */
export function getEditorBackHref(): string | null {
  return backHref;
}
