"""Regression tests for the signing prerequisites, independent of a certificate."""
from __future__ import annotations

import ast
import unittest

from write_version_info import render_version_info


class VersionResourceTests(unittest.TestCase):
    def test_metadata_matches_both_executables(self) -> None:
        for filename in ("Zadoo.exe", "Zadoo-x64.exe"):
            with self.subTest(filename=filename):
                text = render_version_info("1.1.4", filename)
                ast.parse(text)
                self.assertIn("filevers=(1, 1, 4, 0)", text)
                self.assertIn("prodvers=(1, 1, 4, 0)", text)
                self.assertIn("StringStruct('ProductName', 'Zadoo')", text)
                self.assertIn("StringStruct('ProductVersion', '1.1.4')", text)
                self.assertIn(f"StringStruct('OriginalFilename', '{filename}')", text)

    def test_resource_is_deterministic(self) -> None:
        self.assertEqual(render_version_info("1.1.4", "Zadoo.exe"), render_version_info("1.1.4", "Zadoo.exe"))

    def test_invalid_versions_fail_with_exact_error(self) -> None:
        for version in ("1.2", "v1.2.3", "1.2.3.4", "01.2.3", "1.2.3\n", "1.2.bad", "-1.2.3"):
            with self.subTest(version=version), self.assertRaisesRegex(ValueError, "^Version must contain exactly three non-negative integers without leading zeros\\.$"):
                render_version_info(version, "Zadoo.exe")
        with self.assertRaisesRegex(ValueError, "^Windows version components must be between 0 and 65535\\.$"):
            render_version_info("1.2.65536", "Zadoo.exe")

    def test_unexpected_filenames_are_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "^Original filename must be Zadoo.exe or Zadoo-x64.exe\\.$"):
            render_version_info("1.1.4", "other.exe")


if __name__ == "__main__":
    unittest.main()
