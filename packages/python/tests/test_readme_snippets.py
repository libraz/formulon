"""Runs the ```python run blocks of the READMEs against this package.

A block is a program: its stdout must equal the expectations written in it,
either as a trailing ``# value`` comment on a ``print(...)`` line or as a run of
comment lines closing the block. A block with no expectation only has to exit
0. Each block runs in a scratch copy of ``fixtures/readme/``.

With ``FORMULON_DOCS_DIR`` set, every ``.md`` under it is walked as well and its
``python run`` fences run the same way. ``python no-run`` fences are not executed. A fence may
carry a VitePress code-group label (``python run [Python]``); a ``python`` fence that mentions
``run`` but does not match that grammar fails instead of being skipped.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import List, Tuple

HERE = Path(__file__).resolve().parent
PKG_ROOT = HERE.parent
REPO_ROOT = PKG_ROOT.parent.parent
FIXTURE_DIR = HERE / "fixtures" / "readme"
READMES = ("README.md", "README_ja.md", "packages/python/README.md")

# A ``python run`` / ``python no-run`` fence, optionally with a VitePress code-group label (``python run [Python]``).
_FENCE = re.compile(r"^```python (run|no-run)(?:\s+\[[^\]]*\])?\s*$")
_FIXTURE_FENCES = HERE / "fixtures" / "fences.md"
_DOCS_ENV = "FORMULON_DOCS_DIR"
_INLINE = re.compile(r"print\(.*\)\s*# (.+)$")


def run_blocks(markdown: str) -> List[Tuple[int, List[str]]]:
    """Return the (first line number, lines) of every ``python run`` block."""
    lines = markdown.split("\n")
    blocks: List[Tuple[int, List[str]]] = []
    i = 0
    while i < len(lines):
        fence = _FENCE.match(lines[i])
        if fence and fence.group(1) == "run":
            start = i + 2
            body: List[str] = []
            i += 1
            while i < len(lines) and not lines[i].startswith("```"):
                body.append(lines[i])
                i += 1
            blocks.append((start, body))
        i += 1
    return blocks


def near_misses(markdown: str) -> List[int]:
    """1-based lines of ``python`` fences that mention ``run`` but do not match the fence grammar."""
    return [
        n
        for n, line in enumerate(markdown.split("\n"), start=1)
        if line.startswith("```python") and re.search(r"\brun\b", line) and not _FENCE.match(line)
    ]


def expectations(lines: List[str]) -> List[str]:
    """Expected stdout lines: inline ``print(...)  # x`` comments, then the closing comment run."""
    expected = [m.group(1).rstrip() for m in map(_INLINE.search, lines) if m]
    end = len(lines)
    while end > 0 and not lines[end - 1].strip():
        end -= 1
    # The closing run may sit indented inside a ``with`` block.
    begin = end
    while begin > 0 and lines[begin - 1].lstrip().startswith("#"):
        begin -= 1
    expected.extend(re.sub(r"^\s*# ?", "", line).rstrip() for line in lines[begin:end])
    return expected


class ReadmeSnippetsTest(unittest.TestCase):
    def test_blocks_run_and_print_what_their_comments_say(self) -> None:
        sources = [(readme, REPO_ROOT / readme) for readme in READMES]
        docs = os.environ.get(_DOCS_ENV)
        if docs:
            root = Path(docs).resolve()
            sources.extend(
                (str(p.relative_to(root)), p) for p in sorted(root.rglob("*.md")) if "node_modules" not in p.parts
            )
        readme_blocks = 0
        for label, path in sources:
            text = path.read_text(encoding="utf-8")
            for line_no in near_misses(text):
                with self.subTest(fence=f"{label}:{line_no}"):
                    self.fail(f"{label}:{line_no}: fence mentions run but is not ```python run [label]")
            for line_no, lines in run_blocks(text):
                if label in READMES:
                    readme_blocks += 1
                with self.subTest(block=f"{label}:{line_no}"):
                    self._run(lines)
        self.assertGreaterEqual(readme_blocks, 3, f"found {readme_blocks} `python run` blocks")

    def test_labelled_fence_is_collected_and_malformed_fence_is_reported(self) -> None:
        text = _FIXTURE_FENCES.read_text(encoding="utf-8")
        self.assertEqual([start for start, _ in run_blocks(text)], [2, 18, 24])
        self.assertEqual(near_misses(text), [8, 11, 14])

    def test_indented_closing_run_is_the_expectation_so_a_mismatch_fails(self) -> None:
        matching, mismatching = run_blocks(_FIXTURE_FENCES.read_text(encoding="utf-8"))[-2:]
        for (_, lines), passes in ((matching, True), (mismatching, False)):
            expected = expectations(lines)
            self.assertEqual(expected, ["1" if passes else "3"])
            run = subprocess.run(  # noqa: S603 -- the interpreter running this test
                [sys.executable, "-c", "\n".join(lines)], capture_output=True, text=True, check=True
            )
            self.assertEqual(run.stdout.rstrip().split("\n") == expected, passes)

    def _run(self, lines: List[str]) -> None:
        with tempfile.TemporaryDirectory(prefix="formulon-readme-") as work:
            shutil.copytree(FIXTURE_DIR, work, dirs_exist_ok=True)
            script = Path(work) / "snippet.py"
            script.write_text("\n".join(lines) + "\n", encoding="utf-8")
            env = dict(
                os.environ, PYTHONPATH=os.pathsep.join(filter(None, [str(PKG_ROOT), os.environ.get("PYTHONPATH")]))
            )
            env["PYTHONIOENCODING"] = "utf-8"
            run = subprocess.run(  # noqa: S603 -- runs the interpreter that runs this test
                [sys.executable, str(script)],
                cwd=work,
                env=env,
                capture_output=True,
                text=True,
                encoding="utf-8",
                check=False,
            )
            self.assertEqual(run.returncode, 0, run.stderr)
            expected = expectations(lines)
            if expected:
                self.assertEqual(run.stdout.rstrip().split("\n"), expected)


if __name__ == "__main__":
    unittest.main()
