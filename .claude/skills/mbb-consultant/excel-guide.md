# Excel Output Standards for MBB Deliverables

## Required Sheet Structure

Every Excel workbook MUST contain exactly these 5 sheets, in this order:

### Sheet 1: Methodology
- **Row 1**: Title "Methodology & Approach"
- **Analysis objective**: What question are we answering?
- **Analytical approach**: Which framework/method was used
- **Key assumptions**: Numbered, explicit, with rationale
- **Limitations and caveats**: What this analysis does NOT cover
- **Date of analysis**

### Sheet 2: Data Sources
| Column | Content |
|--------|---------|
| A | Source Name |
| B | Source Type (Primary / Secondary / Tertiary) |
| C | Date Accessed |
| D | Reliability Rating (High / Medium / Low) |
| E | Specific Reference (URL, page, exhibit) |
| F | Notes |

### Sheet 3: Calculations
- This is the analytical model itself
- **ALL values must use Excel formulas**, never hardcoded Python calculations
- Blue font (#0000FF) for input cells, black for formulas, green (#008000) for cross-sheet references
- Input cells get light yellow background (#FFF2CC)
- Assumptions in clearly labeled cells, referenced by formulas
- Headers must specify units: "Revenue ($M)", "Growth (%)", "Employees (#)"

### Sheet 4: Summary
- Executive summary of findings (Pyramid Principle: answer first)
- Key metrics table: Metric | Value | vs. Benchmark | Implication
- "So what?" for each finding
- Recommendations: numbered, actionable, with expected impact

### Sheet 5: Validation
- **Sensitivity analysis**: Vary key assumptions ±10%, ±20%, show impact on outputs
- **Cross-checks**: Do totals reconcile? Do ratios make sense vs. benchmarks?
- **Sanity checks**: Is the answer in a reasonable range?
- **Error log**: Any data quality issues encountered

## Formatting Standards

### Colors
| Element | Background | Font | Hex (bg) | Hex (font) |
|---------|------------|------|----------|------------|
| Header row | Dark blue | White bold | #1F4E79 | #FFFFFF |
| Sub-headers | Medium blue | White | #2E75B6 | #FFFFFF |
| Data rows (odd) | White | Black | #FFFFFF | #000000 |
| Data rows (even) | Light gray | Black | #F2F2F2 | #000000 |
| Input cells | Light yellow | Blue | #FFF2CC | #0000FF |
| Formula cells | White | Black | #FFFFFF | #000000 |
| Totals/summary | White | Black bold | #FFFFFF | #000000 |

### Number Formatting
- Currency: `$#,##0` or `$#,##0.0` (specify units in header — $K, $M, $B)
- Percentages: `0.0%`
- Multiples: `0.0x`
- Negative numbers: Parentheses `(1,234)` — use custom format `#,##0;(#,##0)`
- Large numbers: Use comma separators `#,##0`
- Zeros: Display as `"-"` — use custom format `#,##0;(#,##0);"-"`

### Layout
- Freeze panes on header rows and label columns
- Column widths: auto-fitted, minimum 12 characters
- Row heights: headers 25px, data 18px
- Print area set for each sheet
- No merged cells in data areas (headers only)

## Using the MBBWorkbookBuilder

The `scripts/generate_excel.py` script provides the `MBBWorkbookBuilder` class:

```python
import sys
sys.path.insert(0, '.claude/skills/mbb-consultant/scripts')
from generate_excel import MBBWorkbookBuilder

builder = MBBWorkbookBuilder("Market Sizing Analysis")

# Sheet 1: Methodology
builder.add_methodology(
    objective="Estimate TAM/SAM/SOM for AI consulting in North America",
    approach="Top-down market sizing with bottom-up triangulation",
    assumptions=["US and Canada only", "2024 data", "Excludes government"],
    limitations=["Limited primary research", "Relies on public data"]
)

# Sheet 2: Data Sources
builder.add_data_source("Gartner Market Report", "Secondary", "2024-01-15", "High", "gartner.com/report/123", "2023 data")
builder.add_data_source("Company 10-K Filing", "Primary", "2024-02-01", "High", "SEC EDGAR", "FY2023")

# Sheet 3: Calculations (headers + rows)
builder.add_calculation_headers(["Metric", "2022", "2023", "2024E"], ["", "$M", "$M", "$M"])
builder.add_calculation_row(["Total Market", 5000, 6200, None], row_type="input")
builder.add_calculation_row(["Growth Rate", None, "=C4/B4-1", "=D4/C4-1"], row_type="formula")

# Sheet 4: Summary
builder.add_summary(
    findings=[("TAM is $15B", "$15B", "Growing at 18% CAGR", "Attractive entry window")],
    recommendations=["Enter market via acquisition of niche player", "Target enterprise segment first"]
)

# Sheet 5: Validation
builder.add_sensitivity("Growth Rate", 0.18, [0.10, 0.14, 0.18, 0.22, 0.26])
builder.add_cross_check("Total market vs. sum of segments", "$15.0B", "$14.8B", "PASS")

builder.save("outputs/market_sizing.xlsx")
```

## Inline openpyxl Pattern

When writing custom models without the builder:

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers

wb = Workbook()

# Create all 5 required sheets
sheet_names = ["Methodology", "Data Sources", "Calculations", "Summary", "Validation"]
wb.active.title = sheet_names[0]
for name in sheet_names[1:]:
    wb.create_sheet(name)

# Apply MBB header formatting
HEADER_FILL = PatternFill(start_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
INPUT_FONT = Font(name="Calibri", size=11, color="0000FF")
INPUT_FILL = PatternFill(start_color="FFF2CC", fill_type="solid")

# ... populate sheets with formulas, not hardcoded values ...

wb.save("outputs/analysis.xlsx")
```
