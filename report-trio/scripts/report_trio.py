#!/usr/bin/env python3
"""Compute and render a synchronized Excel, PDF, and PowerPoint report."""

from __future__ import annotations

import argparse
import csv
import json
import math
import re
import sys
import tempfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path


def fail(message: str) -> None:
    raise SystemExit(message)


def read_table(path: Path, sheet: str | int = 0, encoding: str | None = None):
    import pandas as pd

    if not path.is_file():
        fail(f"Input file does not exist: {path}")
    suffix = path.suffix.lower()
    if suffix == ".csv":
        encodings = [encoding] if encoding else ["utf-8-sig", "gb18030"]
        error = None
        for codec in encodings:
            try:
                frame = pd.read_csv(path, encoding=codec)
                break
            except UnicodeDecodeError as exc:
                error = exc
        else:
            fail(f"Could not decode CSV ({error}). Pass --encoding explicitly.")
    elif suffix in {".xlsx", ".xlsm"}:
        try:
            import openpyxl  # noqa: F401
        except ImportError:
            fail("Reading Excel requires openpyxl. Install this skill's requirements.txt.")
        try:
            frame = pd.read_excel(path, sheet_name=sheet)
        except Exception as exc:
            fail(f"Could not read workbook sheet {sheet!r}: {exc}")
    else:
        fail("Unsupported input format. Use a .csv or .xlsx file.")
    if isinstance(frame, dict):
        fail("Select one worksheet with --sheet.")
    if frame.empty or len(frame.columns) == 0:
        fail("The input has no data rows or columns.")
    if frame.columns.duplicated().any():
        fail("Column names must be unique; rename duplicate columns and retry.")
    frame.columns = [str(column).strip() for column in frame.columns]
    if any(not column for column in frame.columns):
        fail("Every column must have a non-empty header.")
    return frame


def infer_kind(frame) -> str:
    names = " ".join(str(c).lower() for c in frame.columns)
    patterns = {
        "scores": r"score|grade|成绩|分数|评分|得分",
        "sales": r"sales|revenue|amount|销售|金额|营收|订单额",
        "survey": r"feedback|satisfaction|满意度|评价|选项|问卷",
    }
    for kind, pattern in patterns.items():
        if re.search(pattern, names):
            return kind
    return "numeric"


def numeric_columns(frame) -> list[str]:
    import pandas as pd

    found = []
    for column in frame.columns:
        values = pd.to_numeric(frame[column], errors="coerce")
        if int(values.notna().sum()) > 0:
            found.append(str(column))
    return found


def inspect(args) -> None:
    import pandas as pd

    frame = read_table(Path(args.input), args.sheet, args.encoding)
    numeric = set(numeric_columns(frame))
    columns = []
    for column in frame.columns:
        values = frame[column]
        columns.append({
            "name": str(column),
            "type": "number" if str(column) in numeric else "text/date",
            "missing": int(values.isna().sum()),
            "sample": [None if pd.isna(value) else str(value) for value in values.head(3).tolist()],
        })
    print(json.dumps({"rows": len(frame), "columns": columns, "detected_type": infer_kind(frame)}, ensure_ascii=False, indent=2))


def compute(args) -> None:
    import pandas as pd

    source = Path(args.input)
    frame = read_table(source, args.sheet, args.encoding)
    if args.measure not in frame.columns:
        fail(f"Measure {args.measure!r} is not a column. Available: {', '.join(map(str, frame.columns))}")
    if args.dimension and args.dimension not in frame.columns:
        fail(f"Dimension {args.dimension!r} is not a column. Available: {', '.join(map(str, frame.columns))}")

    values = pd.to_numeric(frame[args.measure], errors="coerce")
    valid = values.notna() & values.map(lambda value: math.isfinite(float(value)))
    if not valid.any():
        fail(f"Measure {args.measure!r} has no usable numeric values.")
    kind = args.kind or infer_kind(frame)
    aggregation = "mean" if kind in {"scores", "survey"} else "sum"
    usable = values[valid].astype(float)
    unit = args.unit or "source units"

    def clean_number(value):
        value = float(value)
        return int(value) if value.is_integer() else round(value, 8)

    agg_value = usable.mean() if aggregation == "mean" else usable.sum()
    metrics = [
        {"label": f"{aggregation.title()} {args.measure}", "value": clean_number(agg_value), "unit": unit, "valid_count": int(valid.sum())},
        {"label": f"Average {args.measure}", "value": clean_number(usable.mean()), "unit": unit, "valid_count": int(valid.sum())},
        {"label": f"Maximum {args.measure}", "value": clean_number(usable.max()), "unit": unit, "valid_count": int(valid.sum())},
        {"label": f"Minimum {args.measure}", "value": clean_number(usable.min()), "unit": unit, "valid_count": int(valid.sum())},
    ]
    breakdowns = []
    charts = []
    if args.dimension:
        work = pd.DataFrame({"group": frame[args.dimension].fillna("(missing)"), "value": values})
        work = work[work["value"].notna()]
        grouped = work.groupby("group", dropna=False)["value"].agg(aggregation).sort_values(ascending=False)
        denominator = float(grouped.abs().sum())
        rows = []
        last_value = None
        last_rank = 0
        for position, (name, value) in enumerate(grouped.items(), start=1):
            rank = last_rank if last_value is not None and value == last_value else position
            last_value, last_rank = value, rank
            count = int((work["group"] == name).sum())
            rows.append({
                "name": str(name), "value": clean_number(value),
                "share": round(abs(float(value)) / denominator, 8) if denominator else None,
                "count": count, "rank": rank,
            })
        breakdowns.append({"dimension": args.dimension, "measure": args.measure, "aggregation": aggregation, "rows": rows})
        charts.append({"type": "bar", "title": f"{args.measure} by {args.dimension}", "x": [r["name"] for r in rows[:20]], "y": [r["value"] for r in rows[:20]]})

    report = {
        "meta": {
            "title": args.title, "data_type": kind, "source_file": source.name,
            "row_count": int(len(frame)), "generated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
            "measure": args.measure, "aggregation": aggregation, "unit": unit,
            "dimension": args.dimension, "excluded_count": int(len(frame) - int(valid.sum())),
        },
        "key_metrics": metrics, "breakdowns": breakdowns, "charts": charts,
        "insights": {"summary": "", "highlights": [], "risks": [], "recommendations": ""},
        "source_columns": [str(c) for c in frame.columns],
        "source_rows": frame.where(pd.notna(frame), None).to_dict(orient="records"),
    }
    output = Path(args.out)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(f"Wrote {output}; {valid.sum()} usable values, {len(frame) - int(valid.sum())} excluded.")


def display(value) -> str:
    if isinstance(value, (int, float)):
        return f"{value:,.2f}".rstrip("0").rstrip(".") if isinstance(value, float) and not value.is_integer() else f"{value:,.0f}"
    return str(value)


def excel_text(value) -> str:
    text = str(value)
    if text.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def make_chart(report, path: Path) -> Path | None:
    if not report.get("charts"):
        return None
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    chart = report["charts"][0]
    fig, ax = plt.subplots(figsize=(10, 4.8))
    ax.bar(chart["x"], chart["y"], color="#247A78")
    ax.set_title(chart["title"])
    ax.tick_params(axis="x", labelrotation=30)
    fig.tight_layout()
    fig.savefig(path, dpi=160)
    plt.close(fig)
    return path


def render_excel(report, path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    summary = wb.active
    summary.title = "摘要"
    summary.append([excel_text(report["meta"]["title"])])
    summary["A1"].font = Font(size=16, bold=True, color="17324D")
    summary.append(["指标", "数值", "单位", "有效记录"])
    for cell in summary[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="247A78")
    for metric in report["key_metrics"]:
        summary.append([excel_text(metric["label"]), metric["value"], excel_text(metric["unit"]), metric["valid_count"]])
    for breakdown in report.get("breakdowns", []):
        summary.append([])
        summary.append([excel_text(f"{breakdown['dimension']} 分组"), excel_text(report["meta"]["measure"]), "占比", "记录数", "排名"])
        start = summary.max_row
        for cell in summary[start]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="247A78")
        for row in breakdown["rows"]:
            summary.append([excel_text(row["name"]), row["value"], row["share"], row["count"], row["rank"]])
    detail = wb.create_sheet("明细")
    columns = report.get("source_columns", [])
    detail.append([excel_text(column) for column in columns])
    for cell in detail[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="247A78")
    for row in report.get("source_rows", []):
        detail.append([excel_text(row.get(column, "")) if isinstance(row.get(column), str) else row.get(column) for column in columns])
    charts = wb.create_sheet("图表")
    for item in report.get("charts", []):
        charts.append([excel_text(item["title"])])
        charts.append(["类别", "数值"])
        for label, value in zip(item["x"], item["y"]):
            charts.append([excel_text(label), value])
        chart = BarChart()
        chart.title = item["title"]
        chart.add_data(Reference(charts, min_col=2, min_row=2, max_row=2 + len(item["y"])), titles_from_data=True)
        chart.set_categories(Reference(charts, min_col=1, min_row=3, max_row=2 + len(item["x"])))
        charts.add_chart(chart, f"D{charts.max_row - len(item['x'])}")
    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        for col in sheet.columns:
            width = min(max(max(len(str(c.value or "")) for c in col) + 2, 12), 36)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def find_font() -> str | None:
    from matplotlib import font_manager

    for family in ("Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "Source Han Sans SC"):
        try:
            path = Path(font_manager.findfont(family, fallback_to_default=False))
            if path.is_file():
                return str(path)
        except (ValueError, RuntimeError):
            continue
    return None


def render_pdf(report, path: Path, chart_path: Path | None) -> None:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape
    from xml.sax.saxutils import escape

    font = find_font()
    face = "Helvetica"
    if font:
        try:
            pdfmetrics.registerFont(TTFont("ReportFont", font, subfontIndex=0))
            face = "ReportFont"
        except Exception:
            pass
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="ReportTitle", parent=styles["Title"], fontName=face, fontSize=22, leading=28, textColor=colors.HexColor("#17324D"), alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="ReportBody", parent=styles["BodyText"], fontName=face, leading=16))
    styles.add(ParagraphStyle(name="ReportHead", parent=styles["Heading2"], fontName=face, textColor=colors.HexColor("#247A78")))
    story = [Paragraph(report["meta"]["title"], styles["ReportTitle"]), Spacer(1, 10), Paragraph(f"Source: {report['meta']['source_file']} | Rows: {report['meta']['row_count']} | Generated: {report['meta']['generated_at']}", styles["ReportBody"]), Spacer(1, 18), Paragraph("Key metrics", styles["ReportHead"])]
    table_rows = [["Metric", "Value", "Unit", "Valid records"]]
    table_rows += [[m["label"], display(m["value"]), m["unit"], str(m["valid_count"])] for m in report["key_metrics"]]
    table = Table(table_rows, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#247A78")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5D8")), ("FONTNAME", (0, 0), (-1, -1), face), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F7F6")]), ("PADDING", (0, 0), (-1, -1), 7)]))
    story.extend([table, Spacer(1, 18)])
    insights = report.get("insights", {})
    for label, value in (("Summary", insights.get("summary")), ("Highlights", insights.get("highlights")), ("Risks", insights.get("risks")), ("Recommendations", insights.get("recommendations"))):
        if value:
            text = "<br/>".join(escape(str(v)) for v in value) if isinstance(value, list) else escape(str(value))
            story.extend([Paragraph(escape(str(label)), styles["ReportHead"]), Paragraph(text, styles["ReportBody"]), Spacer(1, 8)])
    if chart_path:
        story.extend([PageBreak(), Paragraph("Group comparison", styles["ReportHead"]), Image(str(chart_path), width=470, height=225)])
    for breakdown in report.get("breakdowns", []):
        story.extend([Spacer(1, 12), Paragraph(breakdown["dimension"], styles["ReportHead"])])
        rows = [["Group", "Value", "Share", "Count", "Rank"]]
        rows += [[r["name"], display(r["value"]), "" if r["share"] is None else f"{r['share']:.1%}", str(r["count"]), str(r["rank"])] for r in breakdown["rows"]]
        tbl = Table(rows, repeatRows=1)
        tbl.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#247A78")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.3, colors.grey), ("FONTNAME", (0, 0), (-1, -1), face), ("FONTSIZE", (0, 0), (-1, -1), 8)]))
        story.append(tbl)
    story.append(PageBreak())
    story.append(Paragraph("Appendix: source data", styles["ReportHead"]))
    columns = report.get("source_columns", [])
    source_rows = report.get("source_rows", [])
    if columns:
        data = [[str(c)[:24] for c in columns]] + [[str(row.get(c, ""))[:36] for c in columns] for row in source_rows]
        appendix = Table(data, repeatRows=1)
        appendix.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#247A78")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5D8")), ("FONTNAME", (0, 0), (-1, -1), face), ("FONTSIZE", (0, 0), (-1, -1), 7), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.append(appendix)
    path.parent.mkdir(parents=True, exist_ok=True)
    SimpleDocTemplate(str(path), pagesize=A4, rightMargin=38, leftMargin=38, topMargin=42, bottomMargin=42, title=report["meta"]["title"]).build(story)


def render_pptx(report, path: Path, chart_path: Path | None) -> None:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    teal = RGBColor(36, 122, 120)

    def new_slide(title: str):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(0.65), Inches(0.35), Inches(12), Inches(0.65))
        p = box.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(25)
        p.font.bold = True
        p.font.color.rgb = RGBColor(23, 50, 77)
        return slide

    slide = new_slide(report["meta"]["title"])
    sub = slide.shapes.add_textbox(Inches(0.7), Inches(1.2), Inches(11.8), Inches(0.45))
    sub.text_frame.text = f"{report['meta']['source_file']}  |  {report['meta']['row_count']} rows"
    summary = report.get("insights", {}).get("summary", "")
    if summary:
        box = slide.shapes.add_textbox(Inches(0.7), Inches(2.1), Inches(11.8), Inches(1.3))
        box.text_frame.text = summary
        box.text_frame.paragraphs[0].font.size = Pt(22)

    slide = new_slide("Key metrics")
    count = min(4, len(report["key_metrics"]))
    for index, metric in enumerate(report["key_metrics"][:count]):
        left = 0.7 + index * 3.1
        box = slide.shapes.add_textbox(Inches(left), Inches(2.2), Inches(2.8), Inches(1.7))
        tf = box.text_frame
        tf.text = display(metric["value"])
        tf.paragraphs[0].font.size = Pt(30)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = teal
        p = tf.add_paragraph()
        p.text = f"{metric['label']} ({metric['unit']})"
        p.font.size = Pt(13)
    if chart_path:
        slide = new_slide(report["charts"][0]["title"])
        slide.shapes.add_picture(str(chart_path), Inches(0.8), Inches(1.35), width=Inches(11.7))
    for breakdown in report.get("breakdowns", []):
        slide = new_slide(f"{breakdown['dimension']} comparison")
        table = slide.shapes.add_table(min(len(breakdown["rows"]) + 1, 13), 4, Inches(0.8), Inches(1.35), Inches(11.7), Inches(5.1)).table
        for col, heading in enumerate([breakdown["dimension"], report["meta"]["measure"], "Share", "Rank"]):
            table.cell(0, col).text = heading
        for row_index, row in enumerate(breakdown["rows"][:12], start=1):
            values = [row["name"], display(row["value"]), "" if row["share"] is None else f"{row['share']:.1%}", str(row["rank"])]
            for col, value in enumerate(values):
                table.cell(row_index, col).text = value
                if row_index == 1:
                    table.cell(row_index, col).fill.solid()
                    table.cell(row_index, col).fill.fore_color.rgb = RGBColor(242, 247, 246)
        for cell in table.rows[0].cells:
            cell.fill.solid()
            cell.fill.fore_color.rgb = teal
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.color.rgb = RGBColor(255, 255, 255)
                paragraph.font.bold = True
    slide = new_slide("Conclusions and recommendations")
    pieces = []
    insights = report.get("insights", {})
    pieces.extend(str(v) for v in insights.get("highlights", []))
    pieces.extend(str(v) for v in insights.get("risks", []))
    if insights.get("recommendations"):
        pieces.append(str(insights["recommendations"]))
    box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.6), Inches(5.2))
    box.text_frame.text = "\n".join(f"• {item}" for item in pieces) if pieces else "No interpretation provided."
    for paragraph in box.text_frame.paragraphs:
        paragraph.font.size = Pt(20)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(path)


def render(args) -> None:
    report_path = Path(args.report)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="report-trio-") as temp:
        chart = make_chart(report, Path(temp) / "chart.png")
        render_excel(report, out / "report.xlsx")
        render_pdf(report, out / "report.pdf", chart)
        render_pptx(report, out / "report.pptx", chart)
    print(f"Rendered Excel, PDF, and PPTX into {out}")


def validate(args) -> None:
    from openpyxl import load_workbook
    from pptx import Presentation
    from pypdf import PdfReader

    report = json.loads(Path(args.report).read_text(encoding="utf-8"))
    out = Path(args.out_dir)
    paths = [out / "report.xlsx", out / "report.pdf", out / "report.pptx"]
    missing = [str(path) for path in paths if not path.is_file()]
    if missing:
        fail("Validation incomplete; missing: " + ", ".join(missing))
    wb = load_workbook(paths[0], data_only=True, read_only=True)
    excel_values = {row[0]: row[1] for row in wb["摘要"].iter_rows(min_row=3, max_col=2, values_only=True) if row[0]}
    pdf_text = "\n".join(page.extract_text() or "" for page in PdfReader(str(paths[1])).pages)
    prs = Presentation(paths[2])
    ppt_text = "\n".join(shape.text for slide in prs.slides for shape in slide.shapes if shape.has_text_frame)
    checks = []
    for metric in report["key_metrics"]:
        label, value = metric["label"], metric["value"]
        excel = excel_values.get(label)
        excel_ok = excel is not None and math.isclose(float(excel), float(value), rel_tol=1e-8, abs_tol=1e-8)
        token = display(value)
        pdf_ok = token in pdf_text
        ppt_ok = token in ppt_text
        checks.append({"metric": label, "excel": excel_ok, "pdf": pdf_ok, "pptx": ppt_ok})
    failed = [item for item in checks if not all(item[key] for key in ("excel", "pdf", "pptx"))]
    result = {"passed": not failed, "checks": checks}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if failed:
        raise SystemExit(1)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    for name, handler in (("inspect", inspect), ("compute", compute)):
        cmd = sub.add_parser(name)
        cmd.add_argument("input")
        cmd.add_argument("--sheet", default=0, type=lambda value: int(value) if value.isdigit() else value)
        cmd.add_argument("--encoding")
        if name == "compute":
            cmd.add_argument("--title", required=True)
            cmd.add_argument("--measure", required=True)
            cmd.add_argument("--dimension")
            cmd.add_argument("--unit")
            cmd.add_argument("--kind", choices=["scores", "sales", "survey", "numeric"])
            cmd.add_argument("--out", default="report.json")
        cmd.set_defaults(func=handler)
    cmd = sub.add_parser("render")
    cmd.add_argument("report")
    cmd.add_argument("--out-dir", default=".")
    cmd.set_defaults(func=render)
    cmd = sub.add_parser("validate")
    cmd.add_argument("report")
    cmd.add_argument("--out-dir", default=".")
    cmd.set_defaults(func=validate)
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
