#!/usr/bin/env python3
"""Check repository ignore rules without creating files or changing the index."""
import subprocess
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent


class RepositorySafetyTests(unittest.TestCase):
    def test_private_files_are_ignored_by_repository_rules(self):
        paths = [
            "work/notice.doc",
            "work/notice.docx",
            "work/notice.DOCX",
            "work/template.dotx",
            "work/template.dotm",
            "work/notice.docm",
            "work/notice.rtf",
            "work/notice.odt",
            "work/data.xls",
            "work/data.xlsx",
            "work/data.xlsm",
            "work/data.csv",
            "work/report.pdf",
            "work/slides.pptx",
            "sample/private.docx",
            "sample/private.docx.bak-20261026-100000",
            "sample/private.docx.bak-20261026-100000-1",
            "sample/private.docx.bak",
            ".env",
            ".env.production",
            "config/.env.local",
            "config/private.pem",
            "config/private.key",
            "config/private.p12",
            "config/private.pfx",
        ]
        result = subprocess.run(
            ["git", "-c", "core.excludesFile=/dev/null", "check-ignore",
             "--no-index", "-v", "-z", "--stdin"],
            cwd=ROOT, input="\0".join(paths) + "\0",
            capture_output=True, text=True, check=False,
        )
        self.assertIn(result.returncode, (0, 1), result.stderr)
        fields = result.stdout.rstrip("\0").split("\0") if result.stdout else []
        ignored = {
            fields[i + 3]
            for i in range(0, len(fields), 4)
            if fields[i] == ".gitignore" and not fields[i + 2].startswith("!")
        }
        self.assertEqual(set(paths) - ignored, set())

    def test_source_files_remain_visible(self):
        paths = [
            "README.md", "SKILL.md", "requirements.txt", "LICENSE",
            "scripts/new_check.py", "references/new_note.md",
        ]
        result = subprocess.run(
            ["git", "-c", "core.excludesFile=/dev/null", "check-ignore",
             "--no-index", "-z", "--stdin"],
            cwd=ROOT, input="\0".join(paths) + "\0",
            capture_output=True, text=True, check=False,
        )
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)

    def test_reviewed_sample_stays_tracked(self):
        result = subprocess.run(
            ["git", "ls-files", "--error-unmatch", "--",
             "sample/\u6837\u4f8b\u516c\u6587.docx"],
            cwd=ROOT, capture_output=True, text=True, check=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == "__main__":
    unittest.main()
