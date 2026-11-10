"""The JS and Python bindings word the same failure the same way, after the camelCase / snake_case mapping.

Each case fails one call per binding for one reason; the JS half (``fixtures/message_parity.mjs``) runs the
same case ids against the staged npm package (``packages/npm/dist``) under Node. The comparison maps every
camelCase identifier of the JS message to snake_case, so ``Workbook.setDefinedName definedName.name`` and
``Workbook.set_defined_name defined_name.name`` agree.
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import unittest
from pathlib import Path
from typing import Any, Callable, Dict

from formulon import ArrayFormula, DefinedNameInput, FormulonError, TableInput, Workbook

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parents[2]
NPM_PACKAGE = REPO_ROOT / "packages" / "npm"
NPM_ENTRY = (
    NPM_PACKAGE / json.loads((NPM_PACKAGE / "package.json").read_text(encoding="utf-8"))["exports"]["."]["import"]
)
JS_CASES = HERE / "fixtures" / "message_parity.mjs"

CASES: Dict[str, Callable[[Workbook, Any], Any]] = {
    "sheet past the count": lambda wb, ws: wb.sheet(5),
    "sheet name matching none": lambda wb, ws: wb.sheet("Nope"),
    "malformed defined name": lambda wb, ws: wb.set_defined_name(DefinedNameInput(name="A1", formula="1")),
    "malformed table name": lambda wb, ws: ws.add_table(
        TableInput(range="A1:B3", name="Has Space", columns=["a", "b"])
    ),
    "write into part of an array": lambda wb, ws: (
        ws.set_array_formula(ArrayFormula(range="E1:E3", formula="=A1:A3*2")),
        ws.get_value("E2"),
        ws.set_value("E2", 5),
    ),
}


def _snake(text: str) -> str:
    return re.sub(r"(?<=[a-z0-9])([A-Z])", lambda m: "_" + m.group(1).lower(), text)


def _code_name(error: BaseException) -> str:
    if isinstance(error, FormulonError):
        return "".join(part.capitalize() for part in error.code.name.split("_"))
    return type(error).__name__


class MessageParityTest(unittest.TestCase):
    def test_js_and_python_word_each_failure_alike(self) -> None:
        node = shutil.which("node")
        if node is None or not NPM_ENTRY.is_file():
            self.skipTest("needs node and a staged npm package (make npm-package)")
        run = subprocess.run(  # noqa: S603 -- node on the repo's own fixture
            [node, str(JS_CASES), NPM_ENTRY.as_uri()], capture_output=True, text=True, check=True
        )
        js: Dict[str, Any] = json.loads(run.stdout)
        self.assertEqual(sorted(js), sorted(CASES))
        for case, call in CASES.items():
            with self.subTest(case=case):
                self.assertIsNotNone(js[case], "the JS call did not fail")
                with Workbook.create_default() as wb:
                    with self.assertRaises(Exception) as caught:
                        call(wb, wb.sheet(0))
                error = caught.exception
                message = getattr(error, "message", None) or str(error)
                self.assertEqual((_code_name(error), message), (js[case][0], _snake(js[case][1])))


if __name__ == "__main__":
    unittest.main()
