# Report Trio

Generate an Excel analysis workbook, a printable PDF, and a PowerPoint deck from one CSV or Excel table. Metrics are calculated by Python and shared through a single `report.json` intermediate representation.

## Requirements

- Python 3.10 or newer
- Install dependencies with `python -m pip install -r report-trio/requirements.txt`

## Quick start

Run these commands from the repository root in PowerShell:

```powershell
python report-trio/scripts/report_trio.py inspect .\data.xlsx
python report-trio/scripts/report_trio.py compute .\data.xlsx --title "Monthly sales" --measure Revenue --dimension Region --out .\report.json
python report-trio/scripts/report_trio.py render .\report.json --out-dir .\output
python report-trio/scripts/report_trio.py validate .\report.json --out-dir .\output
```

The `compute` command requires an explicit numeric measure. `--dimension` is optional. For workbooks with multiple sheets, choose one with `--sheet`. CSV encoding can be set with `--encoding`.

All generated reports and local report JSON files are excluded by `.gitignore` by default. Review the source data before sharing generated reports because they may contain its contents.

## Skill

The reusable Codex skill is in [`report-trio/`](report-trio/). See the [English instructions](report-trio/SKILL.md) and [中文说明](report-trio/SKILL.zh-CN.md).

## 许可证

本项目采用 MIT 许可证，详见 [`LICENSE`](LICENSE)。

## 使用方法

从仓库根目录在 PowerShell 中运行：

```powershell
python report-trio/scripts/report_trio.py inspect .\data.xlsx
python report-trio/scripts/report_trio.py compute .\data.xlsx --title "月度销售" --measure Revenue --dimension Region --out .\report.json
python report-trio/scripts/report_trio.py render .\report.json --out-dir .\output
python report-trio/scripts/report_trio.py validate .\report.json --out-dir .\output
```

安装依赖：

```powershell
python -m pip install -r report-trio/requirements.txt
```
