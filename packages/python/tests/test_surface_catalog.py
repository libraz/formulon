"""Function catalog surface tests."""

from __future__ import annotations

import json
import unittest
from pathlib import Path

import formulon
from formulon import (
    Workbook,
)


class FunctionCatalogTests(unittest.TestCase):
    def test_catalog_metadata(self) -> None:
        self.assertGreater(Workbook.function_count(), 0)
        meta = Workbook.function_metadata("SUM", 0)
        self.assertIsNotNone(meta)
        self.assertEqual(meta.name, "SUM")
        self.assertGreaterEqual(meta.min_arity, 1)
        # SUM is an unbounded variadic; the sentinel is normalized to None.
        self.assertIsNone(meta.max_arity)
        # Lazy-dispatch forms (not in the eager registry) still resolve.
        xlookup = Workbook.function_metadata("XLOOKUP", 0)
        self.assertIsNotNone(xlookup)
        self.assertEqual(xlookup.name, "XLOOKUP")
        names = {Workbook.function_name_at(i) for i in range(Workbook.function_count())}
        self.assertIn("XLOOKUP", names)
        self.assertIsNone(Workbook.function_metadata("NOT_A_REAL_FUNCTION", 0))

    def test_function_name_at(self) -> None:
        name = Workbook.function_name_at(0)
        self.assertIsInstance(name, str)
        self.assertGreater(len(name), 0)

    def test_merge_function_metadata(self) -> None:
        base = Workbook.function_metadata("XLOOKUP", 0)
        self.assertIsNotNone(base)
        # The engine leaves display metadata empty.
        self.assertIsNone(base.signature_template)
        self.assertIsNone(base.description)

        entry = {
            "signature": "XLOOKUP(lookup_value, lookup_array, return_array)",
            "description": "Searches a range or an array.",
            "aliases": {"fr-FR": "RECHERCHEX"},
            "localized": {"fr-FR": {"signature": "RECHERCHEX(...)", "description": "Recherche."}},
        }

        # Localized override wins for the matching locale.
        fr = formulon.merge_function_metadata(base, entry, "fr-FR")
        self.assertEqual(fr.signature_template, "RECHERCHEX(...)")
        self.assertEqual(fr.description, "Recherche.")
        self.assertEqual(fr.localized_name, "RECHERCHEX")
        # Structural fields survive the merge.
        self.assertEqual(fr.name, "XLOOKUP")

        # A locale with no localized/alias entry falls back to the default
        # signature/description and the canonical display name.
        de = formulon.merge_function_metadata(base, entry, "de-DE")
        self.assertEqual(
            de.signature_template,
            "XLOOKUP(lookup_value, lookup_array, return_array)",
        )
        self.assertEqual(de.description, "Searches a range or an array.")
        self.assertEqual(de.localized_name, "XLOOKUP")

        # No provider entry -> base returned verbatim; display metadata NULL.
        none = formulon.merge_function_metadata(base, None, "fr-FR")
        self.assertIs(none, base)
        self.assertIsNone(none.signature_template)
        self.assertIsNone(none.description)

    def test_shipped_example_document_conforms_to_the_engine(self) -> None:
        """The example provider document must stay usable against this engine.

        ``docs/examples/function-metadata.example.json`` is offered as a
        ready-to-use starting point, but nothing else loads it, so a
        function rename or a schema change would leave a broken example
        shipped in the docs. Checking it here also pins the two claims the
        schema doc makes about the keys -- canonical, uppercase, English --
        against the engine that decides what canonical means.
        """
        path = Path(__file__).resolve().parents[3] / "docs" / "examples" / "function-metadata.example.json"
        document = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(document["version"], 1)
        functions = document["functions"]
        self.assertTrue(functions, "the example document declares no functions")

        for name, entry in functions.items():
            self.assertEqual(name, name.upper(), f"{name} is not an uppercase key")
            base = Workbook.function_metadata(name, 0)
            self.assertIsNotNone(base, f"{name} is not a function this engine knows")
            self.assertEqual(base.name, name, f"{name} is not the canonical spelling")
            # Merging must actually reach the entry: an alias-only entry
            # (VLOOKUP here) still has to resolve its display name, which is
            # what distinguishes a real merge from returning `base` verbatim.
            for locale in sorted(set(entry.get("aliases", {})) | set(entry.get("localized", {}))):
                merged = formulon.merge_function_metadata(base, entry, locale)
                self.assertEqual(merged.name, name)
                alias = entry.get("aliases", {}).get(locale)
                self.assertEqual(merged.localized_name, alias if alias is not None else name)


if __name__ == "__main__":
    unittest.main()
