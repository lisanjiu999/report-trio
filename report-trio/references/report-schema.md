# Report IR

`report.json` is the sole numerical source for all rendered formats. `compute` creates numerical fields; the model may add text only to `insights`.

```json
{
  "meta": {
    "title": "Monthly sales report",
    "data_type": "sales",
    "source_file": "sales.csv",
    "row_count": 240,
    "generated_at": "ISO-8601 timestamp",
    "measure": "Revenue",
    "aggregation": "sum",
    "unit": "source values",
    "dimension": "Region"
  },
  "key_metrics": [
    {"label": "Total Revenue", "value": 125500, "unit": "source values", "valid_count": 240},
    {"label": "Average Revenue", "value": 523, "unit": "source values", "valid_count": 240},
    {"label": "Maximum Revenue", "value": 9000, "unit": "source values", "valid_count": 240}
  ],
  "breakdowns": [{"dimension": "Region", "measure": "Revenue", "rows": [
    {"name": "North", "value": 68000, "share": 0.5418, "count": 80, "rank": 1}
  ]}],
  "charts": [{"type": "bar", "title": "Revenue by Region", "x": ["North"], "y": [68000]}],
  "insights": {"summary": "", "highlights": [], "risks": [], "recommendations": ""}
}
```

Totals and group breakdown values are sums for sales/numeric data and averages for scores/survey data. Shares use absolute group values over the sum of absolute values; they are omitted when the denominator is zero. Missing, non-numeric, and non-finite measure values are excluded and reflected by `valid_count`. Ranks use descending competition ranking: ties share a rank and subsequent ranks may skip. A breakdown is omitted when no dimension is selected.

Do not add calculated metrics to the IR manually. Interpretation may quote values already present in `key_metrics` and `breakdowns`, but cannot modify them.
