import type { OfficeMessagesByLocale } from './types';

/**
 * 共用文案：菜单名、通用动作、状态栏基础词汇。
 * 编辑器专属文案由各包通过 I18nProvider.messages 覆盖/扩展。
 * 数组顺序固定为 [zh-CN, en, es, pt-BR]。
 */
type Row = readonly [string, string, string, string];

const ROWS = {
  'shell.back': ['返回', 'Back', 'Atrás', 'Voltar'],
  'shell.more': ['更多', 'More', 'Más', 'Mais'],
  'shell.menubar': ['主菜单', 'Main menu', 'Menú principal', 'Menu principal'],
  'shell.loading': ['正在加载…', 'Loading…', 'Cargando…', 'Carregando…'],
  'shell.share': ['分享', 'Share', 'Compartir', 'Compartilhar'],
  'shell.feedback': ['反馈', 'Feedback', 'Comentarios', 'Feedback'],
  'shell.saved': ['已保存', 'Saved', 'Guardado', 'Salvo'],
  'shell.unsaved': ['未保存', 'Unsaved', 'Sin guardar', 'Não salvo'],

  'shell.menu.file': ['文件', 'File', 'Archivo', 'Arquivo'],
  'shell.menu.edit': ['编辑', 'Edit', 'Editar', 'Editar'],
  'shell.menu.insert': ['插入', 'Insert', 'Insertar', 'Inserir'],
  'shell.menu.view': ['视图', 'View', 'Ver', 'Exibir'],
  'shell.menu.page': ['页面', 'Page', 'Página', 'Página'],
  'shell.menu.format': ['格式', 'Format', 'Formato', 'Formato'],
  'shell.menu.design': ['设计', 'Design', 'Diseño', 'Design'],
  'shell.menu.play': ['放映', 'Slideshow', 'Presentación', 'Apresentação'],
  'shell.menu.arrange': ['排列', 'Arrange', 'Organizar', 'Organizar'],
  'shell.menu.table': ['表格', 'Table', 'Tabla', 'Tabela'],
  'shell.menu.tools': ['工具', 'Tools', 'Herramientas', 'Ferramentas'],
  'shell.menu.help': ['帮助', 'Help', 'Ayuda', 'Ajuda'],

  'action.new': ['新建', 'New', 'Nuevo', 'Novo'],
  'action.open': ['打开', 'Open', 'Abrir', 'Abrir'],
  'action.save': ['保存', 'Save', 'Guardar', 'Salvar'],
  'action.saveToBrowser': ['保存到浏览器', 'Save to browser', 'Guardar en el navegador', 'Salvar no navegador'],
  'action.print': ['打印', 'Print', 'Imprimir', 'Imprimir'],
  'action.rename': ['重命名', 'Rename', 'Renombrar', 'Renomear'],
  'action.import': ['导入', 'Import', 'Importar', 'Importar'],
  'action.export': ['导出', 'Export', 'Exportar', 'Exportar'],
  'action.exportPdf': ['导出 PDF', 'Export PDF', 'Exportar PDF', 'Exportar PDF'],
  'action.exportImage': ['导出图片', 'Export image', 'Exportar imagen', 'Exportar imagem'],
  'action.undo': ['撤销', 'Undo', 'Deshacer', 'Desfazer'],
  'action.redo': ['重做', 'Redo', 'Rehacer', 'Refazer'],
  'action.cut': ['剪切', 'Cut', 'Cortar', 'Recortar'],
  'action.copy': ['复制', 'Copy', 'Copiar', 'Copiar'],
  'action.paste': ['粘贴', 'Paste', 'Pegar', 'Colar'],
  'action.selectAll': ['全选', 'Select all', 'Seleccionar todo', 'Selecionar tudo'],
  'action.delete': ['删除', 'Delete', 'Eliminar', 'Excluir'],
  'action.clearFormat': ['清除格式', 'Clear formatting', 'Borrar formato', 'Limpar formatação'],
  'action.find': ['查找', 'Find', 'Buscar', 'Localizar'],
  'action.zoomIn': ['放大', 'Zoom in', 'Acercar', 'Ampliar'],
  'action.zoomOut': ['缩小', 'Zoom out', 'Alejar', 'Reduzir'],
  'action.zoomReset': ['重置缩放', 'Reset zoom', 'Restablecer zoom', 'Redefinir zoom'],
  'action.fullscreen': ['全屏', 'Full screen', 'Pantalla completa', 'Tela cheia'],
  'action.settings': ['设置', 'Settings', 'Configuración', 'Configurações'],
  'action.shortcuts': ['快捷键', 'Keyboard shortcuts', 'Atajos de teclado', 'Atalhos de teclado'],
  'action.close': ['关闭', 'Close', 'Cerrar', 'Fechar'],
  'action.confirm': ['确定', 'OK', 'Aceptar', 'OK'],
  'action.cancel': ['取消', 'Cancel', 'Cancelar', 'Cancelar'],
  'action.apply': ['应用', 'Apply', 'Aplicar', 'Aplicar'],
  'action.reset': ['重置', 'Reset', 'Restablecer', 'Redefinir'],
  'action.copyLink': ['复制链接', 'Copy link', 'Copiar enlace', 'Copiar link'],
  'action.copied': ['已复制', 'Copied', 'Copiado', 'Copiado'],

  'status.pageOf': ['页面：{current}/{total}', 'Page {current} of {total}', 'Página {current} de {total}', 'Página {current} de {total}'],
  'status.words': ['字数：{count}', 'Words: {count}', 'Palabras: {count}', 'Palavras: {count}'],
} as const satisfies Record<string, Row>;

function build(): OfficeMessagesByLocale {
  const order = ['zh-CN', 'en', 'es', 'pt-BR'] as const;
  const out: OfficeMessagesByLocale = { 'zh-CN': {}, en: {}, es: {}, 'pt-BR': {} };
  for (const [key, values] of Object.entries(ROWS)) {
    order.forEach((locale, index) => {
      out[locale]![key] = values[index];
    });
  }
  return out;
}

export const COMMON_MESSAGES: OfficeMessagesByLocale = build();
