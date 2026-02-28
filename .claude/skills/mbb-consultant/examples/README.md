# MBB Consultant Skill — Example Patterns

This directory describes example output patterns for reference. The skill generates outputs dynamically based on user requests; these patterns illustrate the expected quality and structure.

## Example 1: Market Sizing

**User request**: "How big is the AI consulting market in North America?"

**Expected workflow**:
1. State hypothesis: "The North American AI consulting market is $10-20B and growing at 15-25% CAGR"
2. Build issue tree: TAM (total spend on AI services) → SAM (consulting subset, North America) → SOM (achievable share)
3. Apply top-down sizing (Gartner/IDC data → filter by geography and service type)
4. Apply bottom-up sizing (average engagement size × number of buyers × frequency)
5. Triangulate and reconcile

**Excel output**: 5-sheet workbook with TAM/SAM/SOM calculations, data source tracking, sensitivity on growth assumptions

**PowerPoint output**: 6-8 slides leading with the market size answer, supported by methodology, competitive landscape, and entry recommendations

## Example 2: Profitability Diagnosis

**User request**: "Our margins have been declining for 3 quarters. Help me figure out why."

**Expected workflow**:
1. State hypothesis: "Margin decline is primarily driven by [revenue/cost] factors"
2. Decompose: Profit = Revenue - Cost, Revenue = Price × Volume, Cost = Fixed + Variable
3. Identify which branch explains the decline
4. Drill down into root causes
5. Propose improvement levers with quantified impact

**Excel output**: P&L decomposition with trend analysis, margin bridge, sensitivity on improvement levers

**PowerPoint output**: Diagnosis deck with waterfall chart showing margin walk, root cause analysis, and prioritized improvement recommendations

## Example 3: M&A Due Diligence

**User request**: "Should we acquire TargetCo? They're asking $500M."

**Expected workflow**:
1. State hypothesis: "TargetCo acquisition at $500M creates/destroys shareholder value"
2. Assess strategic fit (capability gaps, market access, synergies)
3. Financial evaluation (DCF, comparables, accretion/dilution)
4. Risk assessment (integration, cultural, regulatory)
5. Recommend go/no-go with conditions

**Excel output**: DCF model with revenue build, synergy quantification, sensitivity on WACC and growth assumptions

**PowerPoint output**: Due diligence deck with strategic rationale, financial assessment, risk matrix, and conditional recommendation
