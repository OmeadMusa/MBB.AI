# MBB Consulting Frameworks Reference

## Problem Structuring Frameworks

### Issue Trees
An issue tree decomposes a hypothesis into testable sub-hypotheses.

**Rules:**
- Top node = the hypothesis (e.g., "Revenue decline is driven by pricing erosion in the enterprise segment")
- Each level breaks down MECE into sub-questions
- Leaves are testable with data
- Typically 2-4 levels deep, 2-4 branches per node

**Structure:**
```
Hypothesis
├── Branch 1 (e.g., Revenue side)
│   ├── Sub-branch 1a (Volume)
│   └── Sub-branch 1b (Price)
└── Branch 2 (e.g., Cost side)
    ├── Sub-branch 2a (Fixed costs)
    └── Sub-branch 2b (Variable costs)
```

**How to build one:**
1. State the hypothesis clearly
2. Ask "What would need to be true for this hypothesis to hold?"
3. For each answer, repeat: "What would need to be true for THIS to hold?"
4. Stop when you reach a level that can be directly tested with data
5. Verify MECE at every level

### MECE Principle
Every decomposition must be:
- **Mutually Exclusive**: No overlap. Test: Can any item belong to two categories? If yes, redefine boundaries.
- **Collectively Exhaustive**: No gaps. Test: Is there any scenario not covered? If yes, add a category or broaden.

**Common MECE splits:**
- Internal vs. External
- Revenue vs. Cost
- Organic vs. Inorganic
- Short-term vs. Long-term
- By geography / by segment / by product line / by customer type
- Supply-side vs. Demand-side
- Quantitative vs. Qualitative

### Pyramid Principle (Barbara Minto)
**Communication structure:**
1. **Governing thought** (the answer/recommendation) — always first
2. **Key line arguments** (3-5 supporting reasons) — MECE
3. **Supporting data/evidence** for each argument

**Rules:**
- Ideas at any level must summarize the ideas grouped below
- Ideas in each grouping must be the same kind of idea
- Ideas in each grouping must be logically ordered (time, structure, or degree)

**Application:**
- Written communication: Lead with the conclusion, then explain why
- Slide decks: Executive summary first, then supporting analysis
- Verbal updates: "We recommend X because of A, B, and C"

---

## Strategy Frameworks

### Porter's Five Forces
Assess industry attractiveness and profitability:
1. **Threat of new entrants** — Barriers to entry: capital requirements, economies of scale, brand loyalty, regulatory, switching costs
2. **Bargaining power of suppliers** — Supplier concentration, switching costs, forward integration threat
3. **Bargaining power of buyers** — Buyer concentration, price sensitivity, backward integration threat
4. **Threat of substitutes** — Availability, price-performance trade-off, switching costs
5. **Competitive rivalry** — Number of competitors, growth rate, differentiation, exit barriers

**When to use:** Industry analysis, market entry decisions, competitive strategy

### 3C's Framework
- **Company**: Core competencies, resources, cost structure, brand, capabilities
- **Customer**: Needs, segments, willingness to pay, buying behavior, decision criteria, pain points
- **Competitor**: Market share, strategy, strengths/weaknesses, cost position, likely responses

**When to use:** Competitive positioning, go-to-market strategy, product strategy

### Value Chain Analysis (Porter)
**Primary activities:** Inbound logistics → Operations → Outbound logistics → Marketing & Sales → Service
**Support activities:** Firm infrastructure, HR management, Technology development, Procurement

**When to use:** Identifying sources of competitive advantage, cost optimization, outsourcing decisions

### PESTEL Analysis
- **Political**: Government policy, trade regulations, political stability
- **Economic**: GDP growth, interest rates, inflation, exchange rates, unemployment
- **Social**: Demographics, cultural trends, health consciousness, education
- **Technological**: Innovation, R&D, automation, digital disruption
- **Environmental**: Climate, sustainability regulations, resource scarcity
- **Legal**: Employment law, consumer protection, antitrust, IP

**When to use:** Market entry, strategic planning, risk assessment, scenario planning

### McKinsey 7S
**Hard elements:** Strategy, Structure, Systems
**Soft elements:** Shared values, Skills, Style, Staff

**When to use:** Organizational assessment, post-merger integration, change management

### BCG Growth-Share Matrix
| | High Market Share | Low Market Share |
|---|---|---|
| **High Growth** | Stars (invest) | Question Marks (decide) |
| **Low Growth** | Cash Cows (harvest) | Dogs (divest) |

**When to use:** Portfolio strategy, resource allocation across business units

---

## Financial / Operational Frameworks

### Profitability Framework
```
Profit = Revenue - Costs

Revenue = Price × Volume
  Price: per unit, per transaction, subscription, tiered
  Volume: # customers × frequency × quantity per transaction

Costs = Fixed + Variable
  Fixed: rent, salaries, depreciation, insurance, G&A
  Variable: COGS, commissions, shipping, raw materials

Margin Analysis:
  Gross Margin = (Revenue - COGS) / Revenue
  EBITDA Margin = EBITDA / Revenue
  Net Margin = Net Income / Revenue
```

**Diagnostic questions:**
- Is the profit issue driven by revenue or costs?
- On the revenue side, is it price or volume?
- On the cost side, is it fixed or variable?
- How do margins compare to industry benchmarks?
- What has changed recently?

### Market Sizing
**Top-down approach:**
1. Start with total addressable market (TAM)
2. Apply segmentation filters (geography, customer type, use case)
3. Apply penetration assumptions
4. Result: Serviceable obtainable market (SOM)

**Bottom-up approach:**
1. Start with unit economics (price per unit, customers per location, etc.)
2. Estimate number of potential customers/transactions
3. Scale up with reasonable adoption assumptions

**Best practice:** Always use both methods and triangulate. If they diverge significantly, investigate why.

### DuPont Analysis
```
ROE = Net Margin × Asset Turnover × Equity Multiplier
    = (Net Income / Revenue) × (Revenue / Assets) × (Assets / Equity)
```

**When to use:** Diagnosing return on equity drivers, comparing company performance

### M&A / Due Diligence Framework
1. **Strategic rationale**: Revenue synergies, cost synergies, market access, capabilities acquisition, defensive move
2. **Financial evaluation**: DCF valuation, comparable transactions, comparable companies, accretion/dilution analysis
3. **Operational risks**: Integration complexity, technology compatibility, supply chain overlap
4. **Cultural fit**: Management alignment, organizational culture, retention risk
5. **Regulatory risk**: Antitrust, foreign investment review, industry-specific approval

### Pricing Strategy
- **Cost-plus**: Cost + markup (simple but ignores demand)
- **Value-based**: Willingness to pay based on perceived value
- **Competitive**: Benchmark against competitors
- **Dynamic**: Real-time adjustment based on demand

**When to use:** New product pricing, pricing optimization, margin improvement

---

## Reasoning Modes

### Deductive Reasoning
General premise + specific observation = certain conclusion

**Structure:** If all A are B, and C is A, then C is B.

**Example:** "All SaaS companies in this segment have 70%+ gross margins (industry data). Company X is a SaaS company in this segment. Therefore, Company X likely has 70%+ gross margins."

**When to use:** When you have reliable general rules and need to apply them to specific cases. Strong for validating assumptions against known industry patterns.

### Inductive Reasoning
Specific observations → probable general conclusion

**Structure:** A1 has property P, A2 has property P, A3 has property P → All A probably have property P.

**Example:** "Companies A, B, and C all saw margin compression after entering the SMB segment. Therefore, entering SMB likely compresses margins in this industry."

**When to use:** When building hypotheses from data patterns. Useful in market research, trend analysis, and benchmarking. Note: conclusions are probable, not certain.

### Abductive Reasoning
Observation + best available explanation = working hypothesis

**Structure:** We observe X. The best explanation for X is Y. Therefore, Y is our working hypothesis.

**Example:** "Revenue dropped 15% QoQ but volume was flat. The most likely explanation is pricing pressure, possibly from a new competitor's entry. We should investigate competitor pricing and contract renewal terms."

**When to use:** When diagnosing problems with incomplete data. This is the most common reasoning mode in consulting — you form hypotheses from limited data and then test them. Always flag abductive conclusions as hypotheses requiring validation.

---

## Framework Selection Guide

| Business Question | Primary Framework(s) |
|---|---|
| "Should we enter this market?" | Porter's Five Forces + Market Sizing + 3Cs |
| "Why are profits declining?" | Profitability Framework + Issue Tree |
| "Should we acquire this company?" | M&A Due Diligence + DCF + Synergy Analysis |
| "How should we organize our portfolio?" | BCG Matrix + Value Chain |
| "What's our competitive advantage?" | 3Cs + Value Chain + Porter's |
| "How do we price this product?" | Pricing Strategy + Willingness-to-Pay Analysis |
| "How big is the opportunity?" | Market Sizing (top-down + bottom-up) |
| "How should we restructure?" | McKinsey 7S + Profitability Framework |
| "What external risks exist?" | PESTEL + Scenario Planning |
| "How do we improve operations?" | Value Chain + Process Mapping + Benchmarking |
