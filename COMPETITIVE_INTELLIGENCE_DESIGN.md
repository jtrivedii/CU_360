# Competitive Intelligence: Recommendations and Data Mapping

## Executive recommendations

1. **Make the comparison universe explicit.** Offer state market, similar-size peers (the selected asset bucket plus adjacent tiers), and nationwide views. Show the universe count and definition next to every benchmark so executives know what each comparison means.
2. **Measure local scale and concentration.** Show the selected CU's rank and percentile by assets, in-state asset share, and membership share within the selected report universe. These are institution-level concentration indicators, not customer-level penetration estimates.
3. **Enable named institution comparisons.** Let users select up to three CUs and compare assets, members, efficiency, capital, profitability, delinquency, and loan-to-share. When no CUs are selected, provide an explicitly labeled closest-asset-size shortlist.
4. **Benchmark operating economics and resilience.** Compare efficiency ratio, ROA, net worth ratio, delinquency, and loan-to-share with peer medians; highlight lower efficiency, capital below 7%, and delinquency above the peer 75th percentile for management review.
5. **Compare balance-sheet structure without implying product mix.** Plot loans/assets, deposits/assets, and net worth/assets as independent ratios against peer medians. The raw Q1/Q2 CSVs include loan-category balances, but the dashboard's normalized CU model does not retain them and duplicate category headers require reconciliation before a product-mix chart is safe.
6. **Use geography only at the level supplied.** The CSVs support headquarters-state comparisons. They do not include city or branch addresses/service areas; do not claim geographic overlap or branch footprint.
7. **Use the two authoritative periods, but do not imply longer history.** Q1 and Q2 2026 support two-period balance growth and a simple comparison for common CUs. The app starts from a packaged Q1 workbook and requires both source CSVs to be imported in Connect Admin; imports are in-memory only and must be repeated after restart. Two quarters do not establish seasonality or a durable trend.

## Proposed section wireframe

```text
Competitive Intelligence                                      [scope: State | Similar-size | Nationwide]
Selected CU · period · comparison-universe definition and institution count

[Asset rank / percentile] [Asset share] [Member share] [Efficiency vs peer median]

[Market asset scale: largest institutions, selected CU highlighted]
[Balance-sheet structure ratios: selected CU vs peer median]

[Competitive scorecard: search/select up to 3 named CUs + peer median]
[Executive opportunity and risk signals]

[Data boundary: two imported periods only, headquarters-only geography, category fields require mapping]
```

The design prioritizes decision questions (market position, operating gap, resilience, and named competitors) over a large collection of disconnected charts. Selection changes the peer universe and all peer-based calculations together.

## Implemented visual-to-data mapping

| Visual / component | Data fields | Calculation / interpretation |
| --- | --- | --- |
| Comparison-universe selector | `state`, `bucket`, `assetsM`, `quarter` | State market = all records with the selected CU's `state`; similar-size = nationwide records within one `bucket` above/below; nationwide = all records. Universe count includes selected CU; peer benchmarks exclude it. |
| Asset rank and percentile KPI | `assetsM` | Competition rank = 1 + count of universe institutions with strictly greater assets (ties share rank). Percentile = count of universe assets less than or equal to selected CU assets / universe count x 100. |
| Asset share KPI | `assetsM` | Selected CU assets / sum of assets in the displayed universe x 100. Only the state-market version is labeled in-state asset share; broader scopes are labeled selected-universe share. |
| Member share KPI | `members` | Selected CU members / sum of reported members in the displayed universe x 100. A concentration indicator, not population penetration. |
| Efficiency vs peer KPI | `efficiencyRatio`, `grossIncomeM`, `operatingExpenseM` | CU efficiency ratio less peer median, in percentage points. Efficiency ratio = operating expense / gross income x 100; lower is generally better. |
| Market asset-scale bar chart | `name`, `assetsM`, `state`, `bucket` | Six largest institutions by assets in the displayed universe. Selected CU is included/highlighted if outside the top six. Values remain at source $M scale; display formatter abbreviates M/B. |
| Balance-sheet structure chart | `loansM`, `depositsM`, `assetsM`, `capRatio` | Selected CU and peer medians for loans/assets, deposits/assets, and net-worth/assets, each x 100. They are independent ratios and are not stacked or presented as a 100% composition. Net-worth ratio uses supplied `capRatio`. |
| Named-rival search and scorecard | `name`, `state`, `assetsM`, `members`, `efficiencyRatio`, `capRatio`, `roa`, `dr`, `ltd` | Up to three named institutions selected from the displayed universe. Shows direct current-period values beside the selected CU and peer median. Empty selection defaults to three closest asset-size peers using absolute asset difference. |
| Opportunity/risk indicators | `capRatio`, `efficiencyRatio`, `dr`; peer `capRatio`, `efficiencyRatio`, `dr` | Review signal for capital below 7%; efficiency above peer median; delinquency above peer 75th percentile. Indicators are screening prompts, not official regulatory determinations. |
| Data-boundary note | `quarter`, `city`, available field list | Documents why growth/movers, city/branch footprint, and loan product mix are not displayed. |

## Deferred measures and required data

| Requested measure | Current status | Data needed |
| --- | --- | --- |
| Asset, membership, loan, and deposits/share growth | Calculable for common Q1/Q2 CU IDs, not implemented as a KPI | Formula: `(Q2 balance - Q1 balance) / Q1 balance x 100`; handle zero/missing prior values as N/A. Both source CSVs must be imported. |
| Trend analysis and variance | Two observations are available after import; no growth-percent KPI is implemented | Use quarter-specific values and peer sets. Add more periods before interpreting seasonality or a durable trend. Net income/charge-off source fields are YTD and require quarterization/annualization rules. |
| Top gainers and decliners | Calculable for matched CUs, not implemented | Match on `CU_NUMBER`; rank percentage and dollar balance changes separately; disclose Q1-only and Q2-only institutions. |
| Loan portfolio composition | Present in raw CSVs, dropped by current CU normalization | Map category balances and resolve conflicting duplicate headers before aggregation. |
| Geographic competitive footprint / overlap | Headquarters-state comparisons only | Branch addresses/geocodes, service-area/field-of-membership definitions, and market population/household denominators are not supplied. |
| True market penetration | Not available | External addressable-market population/household/member-eligibility denominator; current member share is only the CU's share of reported membership within the selected CU universe. |

`KPI_Dictionary.xlsx` documents the earlier normalized 35-field packaged workbook. It does not cover all raw Q1/Q2 columns; see `LOGIC_AUDIT_REPORT.md` for the authoritative quarter-source audit and known mapping gaps.