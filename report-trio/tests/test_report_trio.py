import csv
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from report_trio import excel_text


class ReportPipelineTests(unittest.TestCase):
    def test_excel_text_escapes_formula_prefixes(self):
        for text in ("=1+1", "+SUM(A1:A2)", "-cmd", "@SUM(A1:A2)", "  =1+1"):
            with self.subTest(text=text):
                self.assertTrue(excel_text(text).startswith("'"))
        self.assertEqual(excel_text("ordinary label"), "ordinary label")

    def test_group_ties_share_competition_rank(self):
        script = Path(__file__).resolve().parents[1] / "scripts" / "report_trio.py"
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, output = root / "input.csv", root / "report.json"
            with source.open("w", encoding="utf-8", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(["Group", "Amount"])
                writer.writerows([["A", 10], ["B", 10], ["C", 5]])
            subprocess.run(
                [sys.executable, str(script), "compute", str(source), "--title", "Test",
                 "--measure", "Amount", "--dimension", "Group", "--out", str(output)],
                check=True, capture_output=True, text=True,
            )
            rows = json.loads(output.read_text(encoding="utf-8"))["breakdowns"][0]["rows"]
            self.assertEqual([(row["name"], row["rank"]) for row in rows], [("A", 1), ("B", 1), ("C", 3)])


if __name__ == "__main__":
    unittest.main()
