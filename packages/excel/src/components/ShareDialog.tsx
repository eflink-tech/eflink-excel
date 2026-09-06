// 分享弹窗：调用宿主注入的分享实现生成链接，展示并支持复制
// 分享内容以点击"分享"时刻的文档快照为准
import { useEffect, useRef, useState, type JSX } from 'react';
import { Check, Copy, Loader2, RefreshCw, Share2, TriangleAlert } from 'lucide-react';
import { getExcelShareHandler } from '../core/share/shareBridge';
import type { ExcelShareDoc } from '../core/share/shareBridge';
import type { ExcelShareResult } from '../core/share/shareBridge';

type ShareState =
  | { phase: 'loading' }
  | { phase: 'done' }
  | { phase: 'error'; message: string };

interface ShareDialogProps {
  open: boolean;
  /** 点击"分享"时刻的文档快照（由调用方捕获，弹窗打开期间不随编辑变化） */
  doc: ExcelShareDoc | null;
  onClose: () => void;
}

/** 复制到剪贴板：优先 Clipboard API，失败降级 execCommand */
async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    /* 继续尝试降级方案 */
  }
  const input = document.createElement('textarea');
  input.value = text;
  input.style.position = 'fixed';
  input.style.opacity = '0';
  document.body.appendChild(input);
  input.select();
  const ok = document.execCommand('copy');
  document.body.removeChild(input);
  return ok;
}

export function ShareDialog({ open, doc, onClose }: ShareDialogProps): JSX.Element | null {
  const [state, setState] = useState<ShareState>({ phase: 'loading' });
  const [copied, setCopied] = useState(false);
  // 分享结果（链接 + 宿主业务文案）与复制态
  const [result, setResult] = useState<ExcelShareResult | null>(null);
  // 仅在 open/attempt 变化时重新发起分享（onClose 每次渲染都是新引用，经 ref 读取避免重跑）
  const [attempt, setAttempt] = useState(0);
  const docRef = useRef(doc);
  docRef.current = doc;
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  const copiedTimerRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  useEffect(() => {
    if (!open) return;
    let cancelled = false;
    setState({ phase: 'loading' });
    setResult(null);
    setCopied(false);

    const handler = getExcelShareHandler();
    const target = docRef.current;
    if (!handler || !target) {
      setState({ phase: 'error', message: '当前环境不支持分享' });
      return;
    }
    void handler(target)
      .then((res) => {
        if (!cancelled) {
          setResult(res);
          setState({ phase: 'done' });
        }
      })
      .catch((err: unknown) => {
        if (!cancelled) {
          setState({ phase: 'error', message: err instanceof Error ? err.message : '分享失败，请稍后重试' });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [open, attempt]);

  // Esc 关闭（捕获阶段拦截，避免触发编辑器快捷键）
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        onCloseRef.current();
      }
    };
    document.addEventListener('keydown', onKey, true);
    return () => document.removeEventListener('keydown', onKey, true);
  }, [open]);

  useEffect(() => () => clearTimeout(copiedTimerRef.current), []);

  if (!open) return null;

  const handleCopy = async () => {
    if (state.phase !== 'done' || !result) return;
    if (await copyText(result.url)) {
      setCopied(true);
      clearTimeout(copiedTimerRef.current);
      copiedTimerRef.current = setTimeout(() => setCopied(false), 2000);
    }
  };

  return (
    <div
      className="fixed inset-0 z-[200] flex items-center justify-center bg-black/40 p-4"
      role="presentation"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCloseRef.current();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="excel-share-dialog-title"
        className="w-full max-w-md overflow-hidden rounded-xl border border-[#dadbdd] bg-white shadow-xl"
      >
        <div className="flex items-center justify-between border-b border-[#e8e9eb] px-5 py-3">
          <div className="flex items-center gap-2">
            <Share2 size={16} className="text-[#5f6062]" />
            <h2 id="excel-share-dialog-title" className="text-base font-semibold text-[#26282b]">
              分享表格
            </h2>
          </div>
        </div>

        <div className="px-5 py-4">
          {state.phase === 'loading' && (
            <div className="flex items-center justify-center gap-2 py-8 text-sm text-[#5f6062]">
              <Loader2 size={16} className="animate-spin" />
              正在生成分享链接...
            </div>
          )}

          {state.phase === 'error' && (
            <div className="py-4 text-center">
              <TriangleAlert size={24} className="mx-auto mb-2 text-[#ff8800]" />
              <p className="mb-4 text-sm text-[#26282b]">{state.message}</p>
              <button
                type="button"
                onClick={() => setAttempt((n) => n + 1)}
                className="inline-flex items-center gap-1.5 rounded-md border border-[#dadbdd] px-3 py-1.5 text-sm text-[#26282b] hover:bg-[#f5f6f7]"
              >
                <RefreshCw size={14} />
                重试
              </button>
            </div>
          )}

          {state.phase === 'done' && result && (
            <>
              <div className="flex gap-2">
                <input
                  type="text"
                  readOnly
                  value={result.url}
                  onFocus={(e) => e.currentTarget.select()}
                  aria-label="分享链接"
                  className="min-w-0 flex-1 rounded-md border border-[#dadbdd] bg-[#f5f6f7] px-3 py-2 text-sm text-[#26282b] focus:outline-none"
                />
                <button
                  type="button"
                  onClick={() => void handleCopy()}
                  className={`inline-flex shrink-0 items-center gap-1.5 rounded-md px-3 py-2 text-sm font-medium transition-colors ${
                    copied ? 'bg-[#e8f3ee] text-[#2ea66c]' : 'bg-[#11765c] text-white hover:bg-[#0d5f4a]'
                  }`}
                >
                  {copied ? <Check size={14} /> : <Copy size={14} />}
                  {copied ? '已复制' : '复制'}
                </button>
              </div>
              {result.tips && result.tips.length > 0 && (
                <ul className="mt-4 space-y-1.5 rounded-lg bg-[#f5f6f7] px-3.5 py-3 text-xs leading-relaxed text-[#5f6062]">
                  {result.tips.map((tip) => (
                    <li key={tip}>· {tip}</li>
                  ))}
                </ul>
              )}
            </>
          )}
        </div>

        <div className="flex justify-end border-t border-[#e8e9eb] px-5 py-3">
          <button
            type="button"
            onClick={onCloseRef.current}
            className="rounded-md bg-[#11765c] px-4 py-1.5 text-sm font-medium text-white hover:bg-[#0d5f4a]"
          >
            完成
          </button>
        </div>
      </div>
    </div>
  );
}
