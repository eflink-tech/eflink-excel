import { useEffect, useRef, useState } from 'react';
import { CirclePlus } from 'lucide-react';
import * as cmd from '../../core/univer/commands';
import { getUniverApi } from '../../core/univer/controller';
import { UNIVER_READY_EVENT } from '../toolbar/FormulaBar';

/** Univer 主画布上的固定尺寸：行头列宽 / 列头行高（100% 缩放时取样） */
const ROW_HEADER_WIDTH = 46;
const COL_HEADER_HEIGHT = 24;
/** 最后一行/列之后的空白带宽度（参考稿 30px，⊕ 居中其中） */
const BLANK_STRIP = 30;
const BTN_SIZE = 22;
/** 超出网格内容区域的底色（空单元格透明，画布自身白底，因此超出区域需单独覆盖） */
const BEYOND_COLOR = '#f6f6f7';
const MAIN_CANVAS_SELECTOR = 'canvas[id^="univer-sheet-main-canvas"]';

/**
 * 网格边缘追加行列按钮（参考稿同款 ⊕）：
 * 紧挨最后一列右侧 / 最后一行下方的 30px 空白带内，随滚动与缩放跟随内容边缘。
 * 同时把超出网格内容的画布区域覆盖为参考稿的空白底色（不越过主画布边界，避免盖住底部工作表栏）。
 */
export function GridEdgeControls() {
  const wrapRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef({ x: 0, y: 0 });
  const disposeRef = useRef<{ dispose(): void } | null>(null);
  const [right, setRight] = useState({ left: 0, visible: false });
  const [bottom, setBottom] = useState({ top: 0, visible: false });
  const [rightBand, setRightBand] = useState({ left: 0, top: 0, width: 0, height: 0 });
  const [bottomBand, setBottomBand] = useState({ left: 0, top: 0, width: 0, height: 0 });

  useEffect(() => {
    const recompute = () => {
      try {
        const wrap = wrapRef.current;
        const ws = getUniverApi()?.getActiveWorkbook()?.getActiveSheet();
        if (!wrap || !ws) return;
        // 画布在兄弟节点 UniverContainer 内，需全局查找（id 前缀全局唯一）
        const canvas = document.querySelector<HTMLCanvasElement>(MAIN_CANVAS_SELECTOR);
        if (!canvas) return;
      // 画布相对覆盖层的位置与尺寸（覆盖带不能越过画布，避免盖住底部工作表栏）
      const wrapRect = wrap.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();
      const dx = canvasRect.left - wrapRect.left;
      const dy = canvasRect.top - wrapRect.top;
      const canvasW = canvasRect.width;
      const canvasH = canvasRect.height;
      // 累加行高/列宽得到内容总尺寸，再按滚动与缩放换算成画布上的屏幕位置
      let totalW = 0;
      for (let c = 0; c < ws.getMaxColumns(); c++) totalW += ws.getColumnWidth(c);
      let totalH = 0;
      for (let r = 0; r < ws.getMaxRows(); r++) totalH += ws.getRowHeight(r);
      const zoom = ws.getZoom();
      // 行头/列头尺寸同样参与缩放（实测 50% 时行头仅 23px）
      const xEnd = (ROW_HEADER_WIDTH + (totalW - scrollRef.current.x)) * zoom;
      const yEnd = (COL_HEADER_HEIGHT + (totalH - scrollRef.current.y)) * zoom;
      // 仅当内容末缘进入可视区（含滚动/缩放到末尾、末缘贴住视口边的钳制状态）才显示；
      // +6 容差吸收滚动钳制时的 1px 边界线与取整误差；末缘贴边时按钮钳到边缘内侧，保证可点。
      const rightVisible = xEnd >= ROW_HEADER_WIDTH && xEnd <= dx + canvasW + 6;
      const rightLeft = Math.min(xEnd + (BLANK_STRIP - BTN_SIZE) / 2, dx + canvasW - BTN_SIZE - 2);
      const bottomVisible = yEnd >= COL_HEADER_HEIGHT && yEnd <= dy + canvasH + 6;
      const bottomTop = Math.min(yEnd + (BLANK_STRIP - BTN_SIZE) / 2, dy + canvasH - BTN_SIZE - 2);
      setRight({ left: rightLeft, visible: rightVisible });
      setBottom({ top: bottomTop, visible: bottomVisible });
      // 超出区域覆盖带：+1 保留表格自身的边界线，且不越过画布边界
      setRightBand({
        left: xEnd + 1,
        top: dy,
        width: Math.max(0, dx + canvasW - xEnd - 1),
        height: canvasH,
      });
      setBottomBand({
        left: dx,
        top: yEnd + 1,
        width: canvasW,
        height: Math.max(0, dy + canvasH - yEnd - 1),
      });
           } catch {
        // 几何重算失败时保持现状（按钮/覆盖带维持上一次的位置）
      }
    };

    // 滚动状态同步：Univer 滚轮走 set-scroll-relative（增量），拖滚动条走 scroll-view（绝对值），
    // Scroll facade 事件在当前版本无派发点，故从 onCommandExecuted 的命令参数中还原滚动值。
    const syncScrollFromCommand = (id: string, params: unknown) => {
      const ws = getUniverApi()?.getActiveWorkbook()?.getActiveSheet();
      if (!ws) return;
      const p = (params ?? {}) as {
        sheetViewStartRow?: number;
        sheetViewStartColumn?: number;
        offsetX?: number;
        offsetY?: number;
      };
      if (id === 'sheet.command.scroll-view') {
        let x = 0;
        for (let c = 0; c < (p.sheetViewStartColumn ?? 0); c++) x += ws.getColumnWidth(c);
        let y = 0;
        for (let r = 0; r < (p.sheetViewStartRow ?? 0); r++) y += ws.getRowHeight(r);
        scrollRef.current = { x: x + (p.offsetX ?? 0), y: y + (p.offsetY ?? 0) };
      } else if (id === 'sheet.command.set-scroll-relative') {
        scrollRef.current = {
          x: scrollRef.current.x + (p.offsetX ?? 0),
          y: scrollRef.current.y + (p.offsetY ?? 0),
        };
      } else if (id === 'sheet.command.scroll-view-reset') {
        scrollRef.current = { x: 0, y: 0 };
      }
    };

    const bind = () => {
      const api = getUniverApi();
      if (!api) return false;
      disposeRef.current?.dispose();
      const disposables = [
        api.addEvent(api.Event.SheetZoomChanged, recompute),
        api.addEvent(api.Event.SelectionChanged, recompute),
        api.onCommandExecuted((command) => {
          const c = command as { id?: string; params?: unknown };
          const id = c.id ?? '';
          if (id.startsWith('sheet.command.scroll')) syncScrollFromCommand(id, c.params);
          recompute();
        }),
      ];
      disposeRef.current = {
        dispose: () => disposables.forEach((d) => d.dispose()),
      };
      recompute();
      return true;
    };
    bind();
    // Univer 异步挂载可能晚于本组件（ready 事件存在竞态），未就绪时定时重试绑定
    let retryTimer: ReturnType<typeof setTimeout> | null = null;
    const retryBind = () => {
      if (disposeRef.current) return;
      retryTimer = setTimeout(() => {
        if (!bind()) retryBind();
      }, 250);
    };
    retryBind();
    window.addEventListener(UNIVER_READY_EVENT, bind);
    window.addEventListener('resize', recompute);
    return () => {
      if (retryTimer) clearTimeout(retryTimer);
      window.removeEventListener(UNIVER_READY_EVENT, bind);
      window.removeEventListener('resize', recompute);
      disposeRef.current?.dispose();
      disposeRef.current = null;
    };
  }, []);

  const btnClass =
    'pointer-events-auto absolute flex items-center justify-center rounded-full text-slate-500 hover:bg-slate-200/70 hover:text-slate-700';

  return (
    <div ref={wrapRef} className="pointer-events-none absolute inset-0 z-20">
      {/* 超出网格内容的底色覆盖带（右 / 下），+1 保留表格自身的边界线 */}
      {rightBand.width > 0 && rightBand.height > 0 && (
        <div className="absolute" style={{ left: rightBand.left, top: rightBand.top, width: rightBand.width, height: rightBand.height, background: BEYOND_COLOR }} />
      )}
      {bottomBand.width > 0 && bottomBand.height > 0 && (
        <div className="absolute" style={{ left: bottomBand.left, top: bottomBand.top, width: bottomBand.width, height: bottomBand.height, background: BEYOND_COLOR }} />
      )}
      <button
        type="button"
        title="在右侧添加列"
        aria-label="在右侧添加列"
        className={`${btnClass} ${right.visible ? '' : 'hidden'}`}
        style={{ left: right.left, top: (COL_HEADER_HEIGHT - BTN_SIZE) / 2, width: BTN_SIZE, height: BTN_SIZE, background: BEYOND_COLOR }}
        onClick={() => cmd.addColumnRight()}
      >
        <CirclePlus size={17} strokeWidth={1.7} />
      </button>
      <button
        type="button"
        title="在底部添加行"
        aria-label="在底部添加行"
        className={`${btnClass} ${bottom.visible ? '' : 'hidden'}`}
        style={{ left: (ROW_HEADER_WIDTH - BTN_SIZE) / 2, top: bottom.top, width: BTN_SIZE, height: BTN_SIZE, background: BEYOND_COLOR }}
        onClick={() => cmd.addRowBottom()}
      >
        <CirclePlus size={17} strokeWidth={1.7} />
      </button>
    </div>
  );
}
