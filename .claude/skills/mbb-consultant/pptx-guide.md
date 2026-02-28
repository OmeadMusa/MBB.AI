# PowerPoint Output Standards for MBB Deliverables

## Slide Deck Structure (Pyramid Principle)

Every deck follows this structure:

### Slide 1: Title Slide
- Engagement title (bold, centered)
- Client name or context
- Date
- "Confidential" if appropriate

### Slide 2: Executive Summary
- THE answer. One governing thought in the title.
- 3-5 key supporting points as bullets
- This slide alone should convey the full recommendation
- A senior executive who reads only this slide should understand the message

### Slide 3: Situation / Context
- Current state of the business/market
- Relevant background the audience needs
- Key data points that frame the problem

### Slides 4-N: Analysis Slides
- **ONE key message per slide**
- The slide title IS the message — an assertion, not a topic label
- Supporting data/charts/tables below the title
- Source citations on every data-bearing slide

### Slide N+1: Recommendations
- Numbered, actionable recommendations
- Each with: What | Why | Expected Impact | Timeline
- Prioritized by impact or urgency

### Slide N+2: Next Steps
- Specific actions with owners and deadlines
- Phased: Immediate (0-2 weeks), Short-term (1-3 months), Medium-term (3-12 months)

### Appendix (optional)
- Detailed data tables, methodology notes, additional analysis

## The #1 Rule: Action Titles

Every slide title MUST be an assertion — a complete sentence that states the key takeaway.

**BAD (topic titles — never do this):**
- "Revenue Analysis"
- "Market Overview"
- "Competitive Landscape"
- "Financial Summary"

**GOOD (action titles):**
- "Revenue declined 15% YoY driven by enterprise pricing erosion"
- "The North American AI consulting market represents a $15B opportunity growing at 18% CAGR"
- "Three competitors control 60% of market share, but none serve the mid-market effectively"
- "Acquiring TargetCo would add $200M revenue at 3.2x forward multiple"

**Test:** Cover everything except the title. Can someone understand the key message from the title alone? If yes, it's an action title.

## Slide Design Standards

### The Pyramid on Each Slide
```
┌──────────────────────────────────────────────────────┐
│  ACTION TITLE: "Revenue grew 23% driven by new       │
│  product launches in Q3 and Q4"                       │
├──────────────────────────────────────────────────────┤
│                                                       │
│  [Chart / Table / Visual Evidence]                    │
│                                                       │
│  • Supporting bullet 1                                │
│  • Supporting bullet 2                                │
│                                                       │
│  Source: Company filings, FY2024                       │
└──────────────────────────────────────────────────────┘
```

### Color Palette
| Element | Color | Hex |
|---------|-------|-----|
| Primary (headers, key shapes) | Dark Navy | #003A70 |
| Secondary (supporting elements) | Steel Blue | #4472C4 |
| Accent (highlights, callouts) | Teal | #00B0F0 |
| Positive / growth | Green | #00B050 |
| Negative / decline | Red | #FF0000 |
| Neutral / baseline | Gray | #808080 |
| Background | White | #FFFFFF |
| Body text | Dark Gray | #333333 |

### Typography
| Element | Font | Size | Weight |
|---------|------|------|--------|
| Slide title | Calibri | 20-24pt | Bold |
| Subtitle | Calibri | 16pt | Regular |
| Body text | Calibri | 12-14pt | Regular |
| Bullet points | Calibri | 12pt | Regular |
| Chart labels | Calibri | 10-12pt | Regular |
| Source line | Calibri | 8pt | Italic |
| Page number | Calibri | 8pt | Regular |

### Layout Grid
- Slide dimensions: 13.333" × 7.5" (16:9 widescreen)
- Margins: 0.5" all sides
- Title area: top 1.2" of slide
- Content area: below title to 0.5" from bottom
- Source line: bottom 0.3", left-aligned
- Page number: bottom 0.3", right-aligned

## Chart Standards

- Every chart must have a clear title (assertion, not topic)
- Axis labels with units
- Data labels on key data points (not all — avoid clutter)
- Source citation below the chart
- **Minimize chartjunk**: no 3D effects, no excessive gridlines, no decorative elements
- Use consistent colors across the deck
- Horizontal bar charts for comparisons (easier to read)
- Waterfall charts for bridges (revenue walk, cost breakdown)
- Line charts for time series

## Using the MBBDeckBuilder

The `scripts/generate_pptx.py` script provides the `MBBDeckBuilder` class:

```python
import sys
sys.path.insert(0, '.claude/skills/mbb-consultant/scripts')
from generate_pptx import MBBDeckBuilder

deck = MBBDeckBuilder("Market Entry Strategy", subtitle="Prepared for ClientCo", date="February 2026")

deck.add_title_slide()

deck.add_exec_summary(
    governing_thought="ClientCo should enter the North American AI consulting market via acquisition of a niche player, targeting $200M incremental revenue by 2028",
    key_points=[
        "The market represents a $15B opportunity growing at 18% CAGR",
        "Three incumbents dominate but leave the mid-market underserved",
        "Acquiring TargetCo provides immediate capabilities and client base at 3.2x forward revenue",
        "Expected payback period of 2.5 years with 25% IRR",
        "Key risk: integration of TargetCo's engineering talent (mitigated by retention packages)"
    ]
)

deck.add_content_slide(
    action_title="The North American AI consulting market represents a $15B opportunity growing at 18% CAGR",
    body_text="Market sizing analysis using top-down and bottom-up approaches...",
    source="Gartner, IDC, Company filings (2024)"
)

deck.add_table_slide(
    action_title="Three competitors control 60% share but none effectively serve mid-market clients",
    headers=["Company", "Revenue ($M)", "Share (%)", "Focus"],
    rows=[
        ["CompetitorA", "$3,200", "21%", "Enterprise"],
        ["CompetitorB", "$2,800", "19%", "Enterprise"],
        ["CompetitorC", "$3,000", "20%", "Enterprise + Gov"],
    ],
    source="Company filings, press releases (2024)"
)

deck.add_recommendation_slide([
    {"what": "Acquire TargetCo", "why": "Immediate market access + capabilities", "impact": "$200M revenue by 2028", "timeline": "Q2 2026"},
    {"what": "Invest in mid-market offering", "why": "Unserved segment with high margins", "impact": "15pp margin uplift", "timeline": "Q3 2026"},
    {"what": "Hire 50 senior consultants", "why": "Scale delivery capacity", "impact": "Support 3x client volume", "timeline": "Q1-Q4 2026"},
])

deck.add_next_steps_slide(
    immediate=["Sign LOI with TargetCo", "Engage integration planning team"],
    short_term=["Complete due diligence", "Develop mid-market product"],
    medium_term=["Close acquisition", "Launch combined offering"]
)

deck.save("outputs/market_entry_strategy.pptx")
```

## Inline python-pptx Pattern

When writing custom decks without the builder:

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Use blank layout for full control
blank_layout = prs.slide_layouts[6]

# Add slide with action title
slide = prs.slides.add_slide(blank_layout)

# Title text box
title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.9))
tf = title_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Revenue declined 15% YoY driven by enterprise pricing erosion"
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = RGBColor(0x00, 0x3A, 0x70)
p.font.name = "Calibri"

# ... add content, charts, source line ...

prs.save("outputs/deck.pptx")
```
