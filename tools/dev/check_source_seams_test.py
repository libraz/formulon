#!/usr/bin/env python3
"""Fixture tests for the `limit-origin` check of check_source_seams."""

from __future__ import annotations

import unittest

from tools.dev.check_source_seams import limit_origin_violations


def run(text: str, name: str = "src/x.cpp") -> list[str]:
    return limit_origin_violations({name: text})


class LimitOriginTest(unittest.TestCase):
    def test_engine_with_resource_limit_passes(self) -> None:
        src = (
            "// @limit: engine\n"
            "constexpr unsigned kMaxThings = 10;\n"
            "int f(unsigned n) {\n"
            "  if (n > kMaxThings) {\n"
            '    return make_error(FormulonErrorCode::kSecurityResourceLimit, "x");\n'
            "  }\n"
            "}\n"
        )
        self.assertEqual(run(src), [])

    def test_trailing_tag_and_static_member_pass(self) -> None:
        src = (
            "constexpr int kMaxA = 1;  // @limit: internal\n"
            "struct S {\n  static constexpr int kMaxB = 2;  // @limit: format\n};\n"
        )
        self.assertEqual(run(src), [])

    def test_function_local_and_enum_not_definitions(self) -> None:
        src = "void f() {\n  constexpr int kMaxLocal = 3;\n}\nenum E { kMaxEnum = 4 };\n"
        self.assertEqual(run(src), [])

    def test_missing_tag(self) -> None:
        hits = run("constexpr int kMaxThings = 10;\n")
        self.assertEqual(len(hits), 1)
        self.assertIn("kMaxThings has no `@limit:` tag", hits[0])

    def test_bad_tag_value(self) -> None:
        hits = run("// @limit: soft\nconstexpr int kMaxThings = 10;\n")
        self.assertEqual(len(hits), 1)
        self.assertIn("unknown `@limit: soft`", hits[0])

    def test_engine_followed_by_invalid_argument(self) -> None:
        src = (
            "// @limit: engine\nconstexpr unsigned kMaxThings = 10;\n"
            "int f(unsigned n) {\n  if (n >= kMaxThings) {\n"
            '    return make_error(FormulonErrorCode::kInvalidArgument, "x");\n  }\n}\n'
        )
        hits = run(src)
        self.assertEqual(len(hits), 1)
        self.assertIn("refusal other than kSecurityResourceLimit", hits[0])

    def test_engine_followed_by_c_api_helper(self) -> None:
        src = (
            "// @limit: engine\nconstexpr unsigned kMaxThings = 10;\n"
            'int f(unsigned n) {\n  if (kMaxThings < n) {\n    return invalid_argument("x");\n  }\n}\n'
        )
        self.assertEqual(len(run(src)), 1)

    def test_engine_helper_naming_resource_limit_passes(self) -> None:
        src = (
            "// @limit: engine\nconstexpr unsigned kMaxThings = 10;\n"
            "int f(unsigned n) {\n  if (n > kMaxThings) {\n"
            '    return set_binding_error(FM_STATUS_SECURITY_RESOURCE_LIMIT, "x");\n  }\n}\n'
        )
        self.assertEqual(run(src), [])

    def test_engine_clamp(self) -> None:
        src = (
            "// @limit: engine\nconstexpr unsigned kMaxThings = 10;\n"
            "unsigned g(unsigned n) { return std::min(n, kMaxThings); }\n"
        )
        hits = run(src)
        self.assertEqual(len(hits), 1)
        self.assertIn("clamps", hits[0])

    def test_engine_resize_clamp(self) -> None:
        src = "// @limit: engine\nconstexpr unsigned kMaxThings = 10;\nvoid g(V& v) { v.resize(kMaxThings); }\n"
        self.assertEqual(len(run(src)), 1)

    def test_engine_budget_third_argument(self) -> None:
        head = "// @limit: engine\nconstexpr unsigned kMaxThings = 10;\n"
        bad = head + 'void g() { ResourceBudget b(kMaxThings, "things", code); }\n'
        good = head + 'void g() { ResourceBudget b(kMaxThings, "things"); }\n'
        hits = run(bad)
        self.assertEqual(len(hits), 1)
        self.assertIn("third argument", hits[0])
        self.assertEqual(run(good), [])

    def test_excel_followed_by_resource_limit(self) -> None:
        src = (
            "// @limit: excel\nconstexpr unsigned kMaxCols = 10;\n"
            "int f(unsigned n) {\n  if (n > kMaxCols) {\n"
            '    return make_error(FormulonErrorCode::kSecurityResourceLimit, "x");\n  }\n}\n'
        )
        hits = run(src)
        self.assertEqual(len(hits), 1)
        self.assertIn("(excel) is followed by kSecurityResourceLimit", hits[0])

    def test_excel_with_invalid_argument_passes(self) -> None:
        src = (
            "// @limit: excel\nconstexpr unsigned kMaxCols = 10;\n"
            "int f(unsigned n) {\n  if (n > kMaxCols) {\n"
            '    return make_error(FormulonErrorCode::kInvalidArgument, "x");\n  }\n}\n'
        )
        self.assertEqual(run(src), [])

    def test_use_in_comment_ignored(self) -> None:
        src = '// @limit: engine\nconstexpr unsigned kMaxThings = 10;\n// if (n > kMaxThings) invalid("x");\n'
        self.assertEqual(run(src), [])

    def test_tag_scoped_to_defining_file(self) -> None:
        sources = {
            "src/a.cpp": "// @limit: engine\nconstexpr int kMaxN = 1;\n",
            "src/b.cpp": "// @limit: excel\nconstexpr int kMaxN = 2;\nbool f(int n) { return n > kMaxN; }\n",
        }
        self.assertEqual(limit_origin_violations(sources), [])

    def test_js_facade(self) -> None:
        head = "// @limit: engine\nconst MAX_VALUES_CELLS = 4194304;\n"
        bad = head + "if (n > MAX_VALUES_CELLS) {\n  throw new FormulonError(StatusCode.InvalidArgument, 'x');\n}\n"
        good = (
            head + "if (n > MAX_VALUES_CELLS) {\n  throw new FormulonError(StatusCode.SecurityResourceLimit, 'x');\n}\n"
        )
        self.assertEqual(len(run(bad, "packages/npm/common.mjs")), 1)
        self.assertEqual(run(good, "packages/npm/common.mjs"), [])
        self.assertEqual(len(run("const MAX_VALUES_CELLS = 1;\n", "packages/npm/common.mjs")), 1)

    def test_python_facade(self) -> None:
        head = "# @limit: engine\n_GET_VALUES_CELL_CAP = 1 << 22\n"
        bad = head + "if n > _GET_VALUES_CELL_CAP:\n    raise FormulonError(StatusCode.INVALID_ARGUMENT, 'x')\n"
        good = head + "if n > _GET_VALUES_CELL_CAP:\n    raise FormulonError(StatusCode.SECURITY_RESOURCE_LIMIT, 'x')\n"
        self.assertEqual(len(run(bad, "packages/python/formulon/_sheet.py")), 1)
        self.assertEqual(run(good, "packages/python/formulon/_sheet.py"), [])
        self.assertEqual(len(run("_X_CAP = 1\n", "packages/python/formulon/_sheet.py")), 1)


if __name__ == "__main__":
    unittest.main()
