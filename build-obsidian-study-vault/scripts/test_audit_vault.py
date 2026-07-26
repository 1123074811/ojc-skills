#!/usr/bin/env python3
"""Minimal self-check for audit_vault.py."""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path


def run(script: Path, vault: Path, sources: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            sys.executable,
            str(script),
            "--vault",
            str(vault),
            "--source-root",
            str(sources),
            "--root-note",
            "课程知识库",
            "--strict",
        ],
        capture_output=True,
        check=False,
        text=True,
        encoding="utf-8",
        env={**os.environ, "PYTHONUTF8": "1"},
    )


def main() -> None:
    script = Path(__file__).with_name("audit_vault.py")
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        vault = root / "vault"
        sources = root / "sources"
        vault.mkdir()
        sources.mkdir()
        (sources / "第一讲.pdf").write_bytes(b"example source")
        (vault / "课程知识库.md").write_text(
            '# 课程知识库\n\n- [[核心概念]]\n- 来源：[第一讲.pdf](../sources/第一讲.pdf)\n',
            encoding="utf-8",
        )
        (vault / "核心概念.md").write_text(
            '---\ntitle: "核心概念"\ntype: "concept"\n---\n\n'
            "# 核心概念\n\n上级：[[课程知识库]]\n\n"
            "这是用于验证知识点内容长度、双向链接、来源覆盖和 UTF-8 处理的正文。"
            "正文需要足够完整，因此补充定义、成立条件、核心结论、典型用法、常见错误、"
            "前置知识、后续知识以及原始资料位置，确保审计不会把合格知识点误判为空壳。"
            "这里继续给出推导思路、使用边界、相近概念差异和复习建议，模拟真实知识点笔记，"
            "并验证审计只检查实质正文，不把标题、元数据或链接文字错误计入正文长度。"
            "第一讲.pdf\n",
            encoding="utf-8",
        )

        clean = run(script, vault, sources)
        assert clean.returncode == 0, clean.stdout + clean.stderr

        with (vault / "核心概念.md").open("a", encoding="utf-8") as handle:
            handle.write("\n[[不存在的笔记]]\n")
        broken = run(script, vault, sources)
        assert broken.returncode == 1 and "BROKEN_LINK" in broken.stdout, broken.stdout + broken.stderr

    print("audit_vault self-check passed")


if __name__ == "__main__":
    main()
