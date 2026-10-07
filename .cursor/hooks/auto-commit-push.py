#!/usr/bin/env python3
"""On agent stop: stage all changes, commit, and push to origin."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

COMMIT_MESSAGE = "수정 또는 생성"
REPO_ROOT = Path(__file__).resolve().parents[2]


def emit(payload: dict) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False))
    sys.stdout.flush()


def git(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", *args],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def main() -> int:
    raw = sys.stdin.read().strip()
    try:
        data = json.loads(raw) if raw else {}
    except json.JSONDecodeError:
        data = {}

    status = data.get("status") or "completed"
    if status != "completed":
        emit({})
        return 0

    porcelain = git("status", "--porcelain")
    if porcelain.returncode != 0:
        emit({})
        return 0
    if not porcelain.stdout.strip():
        emit({})
        return 0

    add = git("add", "-A")
    if add.returncode != 0:
        emit({})
        return 0

    commit = git("commit", "-m", COMMIT_MESSAGE)
    if commit.returncode != 0:
        # Nothing to commit (e.g. race) or hook blocked the commit.
        emit({})
        return 0

    push = git("push")
    if push.returncode != 0:
        # Fail open: do not block the agent turn; push can be retried later.
        emit({})
        return 0

    emit({})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
