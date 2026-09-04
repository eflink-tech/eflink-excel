#!/usr/bin/env bash
# CI 发布入口：失败时把关键输出转成 ::error:: 注解，直接显示在 commit 上便于排障
set -uo pipefail
LOG="$(mktemp)"
if ! { pnpm build && changeset publish; } 2>&1 | tee "$LOG"; then
  {
    grep -E "npm error|npm warn|code E[A-Z]+|40[13]|ENEEDAUTH|FORBIDDEN|OTP|two-factor|not allowed|permission" "$LOG" | head -6
    echo "---- 最后 6 行输出 ----"
    tail -6 "$LOG"
  } | head -14 | while IFS= read -r line; do echo "::error::${line:0:250}"; done
  exit 1
fi
