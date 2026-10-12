# CU360 Q1/Q2 2026 Logic Audit

## Executive result

The reported Affinity NJ values do not all have the same cause. Q2 delinquency of 0.93% is correct after rounding. Net worth at 0% and loan-to-share at 677,820.7% were importer defects, now corrected. The authoritative raw values independently recalculate to:

| Metric | Q1 2026 | Q2 2026 | Assessment |
| --- | ---: | ---: | --- |
| Net worth / assets | 7.5762% | 7.6621% | Correct after equity-component and derivation fixes |
| Delinquent loans / total loans | 0.7214% | 0.9257% | Correct; Q2 display rounds to 0.93% |
| Total loans / total shares | 101.3878% | 102.6001% | Correct after mapping `Total_Shares_Amount` instead of share count |

Both supplied files were imported through the running dashboard after the fixes. Q1 loaded 4,336 CU rows; Q2 loaded 4,299. Affinity changed from $4.40B assets in Q1 to $4.35B in Q2, and the period selector switched the overview, peer calculations, Competitive Intelligence, demographic aggregates, risk view, and quarter series to the selected dataset. The frontend production build passes.

**Important startup limitation:** the backend API still starts from the packaged, single-quarter Q1 workbook. Q2 is not auto-loaded or persisted by the current application; the CSVs must be imported from Connect Admin after each fresh browser session. The report distinguishes “works after import” from “available immediately at startup.”

## A. Critical Logic Errors

| Severity | Issue and location | Root cause / business impact | Resolution |
| --- | --- | --- | --- |
| Critical, fixed | Loan-to-share denominator in `autoMapColumns` / `IMPORT_FIELD_DEFS` in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L100) | Fuzzy alias `totalshares` matched the earlier `Number_Total_Shares` column before `Total_Shares_Amount`. The count was then divided by 1,000,000 as if it were dollars. For Affinity Q2, loans / 577,237 shares x 100 = **677,820.7154%**. | Added exact-priority `totalsharesamount` and `sharesamount` aliases, plus an exact total-loan amount alias. Live Q1/Q2 import now shows 101.4% and 102.6%. |
| Critical, fixed | Capital derivation gate in `normalizeQuarterRows` in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L265) | The importer assembled equity from `Equity_*` components, but `capRatio` derivation required `has("equityM")`. With no recognized single equity field, the record retained the initialized zero ratio. The source `NET_WORTH_AMOUNT` is equal to total assets for Affinity and is not the usable numerator. | Added `networthamount` recognition and derive capital when parsed equity values exist. Affinity now shows 7.58% / 7.66%. |
| High, fixed | CSV parsing in `parseDelimitedText` in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L366) | The previous parser split the whole file on newline before honoring quoted fields. Q2 was reported as 977,952 rows with 973,653 “no name” skips even though it contains 4,299 logical CU records. Q1 similarly had 38 phantom fragments. This made import validation misleading and could drop/fragment records. | Replaced line-splitting parser with SheetJS CSV parsing and unique header names. Re-imports now report exactly 4,336 and 4,299 rows, with no name skips. |
| High, fixed | Period activation and identity lookup in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L72) and [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L2021) | Missing periods fell back to `ALL_CUS` while keeping the requested quarter label; unmatched selected CUs retained the prior quarter's record. Several comparison views keyed selections by `idx`, which can refer to a different CU in another period. | Missing periods now no-op; CU matching uses charter number, then case-insensitive name + state, and clears a missing CU. Compare extras/rivals reset on data changes; trend comparisons use stable CU identity. |
| High, open | Q2 source bootstrap/persistence in [main.py](backend/app/main.py) and the data-load effect in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L1990) | Startup fetches one `/api/institutions` dataset from `backend/data/cu360_data.xlsx` and initializes a one-quarter registry. The user-supplied Q2 CSV is not loaded automatically; browser imports are React memory only. After restart Q2 disappears and trends return to one period. | Configure the backend to load/publish both quarter files (or persist validated imports server-side) and return the available quarter registry on startup. Do not silently assume Q2 or rely on browser-only state. |
| High, open | Zero/missing ratio handling in `toNum`, `normalizeQuarterRows`, and `computeRiskFlags` in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L146), [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L290), [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L462) | Missing values become numeric zero; zero denominators yield ratio zero. A missing capital value can therefore look like a true 0% and trigger “critically undercapitalized”; zero loans/shares can yield misleading rates. | Preserve missing as `null`; return N/A for zero/missing denominators; suppress threshold alerts when the input or required peer distribution is unavailable. The supplied files have 0 missing values in the required asset/loan/share/member/delinquency/gross-income columns, but Q1 has 11 zero-loan and 2 zero-share rows; Q2 has 9 zero-loan and 2 zero-share rows. |
| High, open | ROA/ROE/NCO/yield formulas in `normalizeQuarterRows` and `METRIC_META` in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L300) and [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L420) | ROA and ROE use period net income divided by ending assets/equity without annualization or average balances. Q2 net income is roughly twice Q1, consistent with year-to-date values. NCO uses YTD charge-offs less recoveries divided by ending loans without annualization. `py` is gross income / assets, not yield on loans, but is labeled Portfolio Yield. | For regulatory-style ROA/ROE use annualized YTD income divided by average assets/equity, with explicit opening-balance requirements. Quarterize/annualize NCO and use average loans. Rename/remove `py` as an income/assets proxy unless true loan-yield inputs are defined. |
| Medium, fixed | Lower-is-better percentile bars in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L513) | Bar color was inverted for delinquency/NCO/loan-to-share while the width still used the unadjusted percentile. Empty peers could also display false “100th percentile” values. | Bar width, color, and label now use the same adjusted rank; unavailable rank displays as N/A. |
| Medium, open | Alert thresholds in `computeRiskFlags` in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx#L462) | Delinquency uses `max(peer P75, 1.5%)`, so it has no unconditional >3% rule when peer P75 is higher. Loan-to-share flags at >100%, earlier than the requested >120% aggressive-funding level. | Separate peer-outlier alerts from absolute guardrails; use a staged loan-to-share monitor/alert policy and disclose that thresholds are management controls, not universal regulatory limits. |

## B. Suspect KPIs and Recalculation

“Current” values below mean the value reported before this audit or the present implementation's formula. Source-derived values use the CSV raw dollars and component fields; confidence is high unless the definition/period convention is noted.

| KPI | Current formula/value | Correct formula and expected value | Variance / confidence |
| --- | --- | --- | --- |
| Net worth ratio, Q2 Affinity NJ | Imported result was 0% because `capRatio` was not derived | Sum mapped net-worth components / total assets x 100 = **333,434,982 / 4,351,759,697 x 100 = 7.6621%** | +7.6621 percentage points; high. Q1 = 7.5762%. |
| Delinquency, Q2 Affinity NJ | `Total_Amount_of_Delinquent_Loans_Leases / Amount_of_Total_Loans_Leases x 100`; displayed 0.93% | Same formula = **36,218,585 / 3,912,631,963 x 100 = 0.9257%** | Display rounding difference +0.0043 pp; high. 0.93% is not a defect. Q1 = 0.7214%. |
| Loan-to-share, Q2 Affinity NJ | Wrong alias produced `loans / Number_Total_Shares x 100` = **677,820.7154%** | `Amount_of_Total_Loans_Leases / Total_Shares_Amount x 100` = **3,912,631,963 / 3,813,476,548 x 100 = 102.6001%** | Wrong result exceeds corrected value by 677,718.1153 pp; high. Q1 = 101.3878%. |
| ROA, Q1 Affinity NJ | `Net_Income / Total_Assets x 100` = about **0.0477%**, shown 0.05% | If `Net_Income` is Q1 YTD, annualize by 4 and divide by average assets. Ending-assets proxy is about **0.1909%**; true value requires Q4 2025 opening/average assets. | +0.1432 pp vs end-asset annualized proxy; medium confidence because average balance is absent. |
| ROA, Q2 Affinity NJ | Current unannualized formula = about **0.1044%**, shown 0.10% | If Q2 net income is YTD, annualize by 2 and divide by average assets. Ending-assets proxy is about **0.2087%**; exact average-asset calculation requires proper period balances. | +0.1044 pp vs end-asset annualized proxy; medium confidence. Still slightly below the user's 0.25% typical lower bound. |
| Portfolio Yield (`py`) | Fallback is `Gross_Income / Total_Assets x 100`; Q2 proxy is about 2.93% | True loan yield should use loan interest income / average loans, or a validated weighted average of category balances and category rates. | Current metric is mislabeled, not a valid yield-on-loans KPI; high confidence. |
| Cost of Funds | No KPI/value in current dashboard | Requires interest/dividend expense on shares and an appropriate average funding balance. Neither is in the normalized fields/source CSV headers reviewed. | N/A; unavailable from current fields. |
| Asset growth, Q1->Q2 | Not shown | `(Q2 assets - Q1 assets) / Q1 assets x 100` = **-1.1745%** | Calculable for CU 857; high. |
| Loan growth, Q1->Q2 | Not shown | Same formula on total loans = **-1.0020%** | Calculable for CU 857; high. |
| Deposit/share growth, Q1->Q2 | Not shown | Same formula on `Total_Shares_Amount` = **-2.1718%** | Calculable for CU 857; high. |
| Membership growth, Q1->Q2 | Not shown | Same formula on `Total_Members` = **+1.3553%** | Calculable for CU 857; high. |
| Efficiency ratio, Q2 Affinity NJ | `Operating_Expenses / Gross_Income x 100` = **57.54%** | Mathematically consistent if both source fields use the same YTD basis. Validate that `Gross_Income` matches the institution's approved efficiency-ratio denominator before regulatory/peer publication. | No arithmetic variance identified; business-definition confidence medium. Q1 = 57.07%. |
| Net income margin, Q2 Affinity NJ | `Net_Income / Gross_Income x 100` = **3.56%** | Same-period net income / gross income x 100; confirm both are YTD and aligned. | No arithmetic defect found; period-definition confidence medium. Q1 = 3.27%. |
| Net charge-off rate (`nco`) | `(YTD charge-offs - YTD recoveries) / ending loans x 100` | For comparable annualized ratio, use net charge-offs annualized for the reporting period / average loans; or derive quarterly charge-offs by subtracting prior YTD values before comparing Q1/Q2. | Current formula is a YTD proxy and not period-comparable; no single expected value reported without validated annualization convention. |

## C. Credit Union Domain Findings

| KPI | Affinity Q1 / Q2 | Reference range | Status and commentary |
| --- | --- | --- | --- |
| Net worth ratio | 7.58% / 7.66% | User heuristic: 7-9.99% adequate; >=10% strong. NCUA PCA net-worth bands are conventionally 7% well-capitalized, 6-<7% adequately capitalized, 4-<6% undercapitalized, 2-<4% significantly undercapitalized, <2% critically undercapitalized. | Above 7%, consistent with well-capitalized PCA category. The user-provided “<5% critical” band is not the PCA classification and should not be labeled regulatory. |
| Delinquency rate | 0.72% / 0.93% | Typically <2%; >3% management alert per request | Within typical range; Q2 remains below 3%. Existing peer/1.5% alert is a separate relative-risk rule. |
| Loan-to-share | 101.39% / 102.60% | User reference 50-110% normal; >120% aggressive | Within the stated normal band. The prior 677,820.7% was a confirmed mapping defect. The current >100% medium flag is a conservative funding-monitor rule, not proof of distress. |
| ROA | 0.05% / 0.10% as currently displayed | User reference 0.25-1.50% | Current values are unannualized proxies. Annualized end-asset proxies are ~0.19% / ~0.21%, still weak-to-low; exact regulatory-style values require average assets and verified YTD semantics. |
| Efficiency ratio | 57.07% / 57.54% | Context-dependent; lower generally more efficient | Plausible operating range, but verify source income/expense definitions and benchmark methodology. |
| Delinquency alerts | No >3% Affinity condition | >3% should alert per request | Affinity should not trigger the absolute 3% rule. Add a distinct absolute threshold rather than relying only on peer P75/1.5%. |

Source-wide distributions use the component/loan/source formulas above; ROA here is deliberately **unannualized**. These are investigation queues, not automatic proof of bad data:

| Ratio | Q1 min / median / P95 / max | Q2 min / median / P95 / max | Audit interpretation |
| --- | --- | --- | --- |
| Net worth / assets | -60.79% / 11.99% / 26.66% / 100.00% | -11.66% / 12.11% / 26.63% / 111.03% | Negative and >100% values require record-level source/component review; Affinity itself is within expected range. |
| Delinquency / loans | 0% / 0.63% / 3.72% / 94.37% | 0% / 0.69% / 4.04% / 42.08% | Very high tails exceed normal operating levels; inspect denominators and underlying CU rows. |
| Loans / shares | 0% / 68.41% / 102.09% / 422.51% | 0% / 69.24% / 103.77% / 387.08% | Upper tails exceed the 120% aggressive-funding guide; they may be true funding pressure or source/denominator issues, not the Affinity mapping defect after correction. |
| Net income / ending assets (unannualized) | -128.56% / 0.16% / 0.50% / 100.00% | -169.24% / 0.35% / 0.98% / 100.00% | Extreme values demonstrate why annualization, average balances, and source outlier review are required before domain interpretation. |

## D. Data Quality Findings

- Q1 contains **4,336** unique CU numbers; Q2 contains **4,299** unique CU numbers. There are **4,295** common IDs, **41** Q1-only IDs, and **4** Q2-only IDs. There are no duplicate CU numbers within either period.
- Required source values checked for assets, total loans, total shares, members, delinquent balances, and gross income were nonblank for all reported CU rows.
- Zero denominators exist: Q1 has 11 zero-loan and 2 zero-share institutions; Q2 has 9 zero-loan and 2 zero-share institutions. Their ratios must be N/A, not a fabricated 0% or a risk flag.
- Raw amounts use Indian-style comma grouping but are dollar values. The importer correctly strips grouping commas and converts amounts to $ millions using the asset-median unit heuristic; ratios are multiplied by 100 once. Affinity's source recalculations confirm no remaining double percentage conversion.
- Q1 has 99 columns and two duplicated loan-category header names. Q2 has 102 columns, two blank headers, and the same duplicate names. Duplicate `Loans_by_Category_All_Other_Non_Commercial_Real_Estate_Loans` columns conflict on 4,336 Q1 rows and 773 Q2 rows; duplicate `Commercial_Loans_LoC_Real_Estate_Secured_Amount` columns conflict on 1,595 Q1 and 1,596 Q2 rows. In Affinity Q2, those duplicate cells include `$225` and `$27,34,12,952`; do not aggregate or map these names without identifying their actual meanings.
- For Affinity, `NET_WORTH_AMOUNT` equals `Total_Assets` in both source periods; the sum of equity component fields produces plausible net worth of $333.6M / $333.4M. The source field must not be trusted alone. Q1-wide reported net worth equals assets for 3,986 records, making component validation important beyond this one CU.
- Q1/Q2 sources have no city, branch address, service-area, average-asset, average-equity, or cost-of-funds fields. Raw loan-category balances and interest-rate fields exist, but the normalized CU model maps only 18 of its 36 standard fields and discards those category columns.
- The earlier `KPI_Dictionary.xlsx` is based on the packaged 35-field Q1 workbook. It is not a complete dictionary of all 99/102 raw CSV columns; this report supersedes its period and source-coverage notes.

## E. Code Changes Made

Changes are in [Dashboard.jsx](frontend/src/pages/Dashboard.jsx):

- Added exact-priority aliases for `Total_Shares_Amount`, `Amount_of_Total_Loans_Leases`, and `NET_WORTH_AMOUNT`; ambiguous fuzzy `totalshares` can no longer select `Number_Total_Shares` first.
- Capital ratio derivation now accepts assembled equity-component values, not only one mapped equity column.
- CSV parsing now uses SheetJS and preserves duplicate columns under unique names; quoted multiline cells no longer become phantom rows.
- Period activation no longer falls back to Q1 under a different selected-quarter label. It matches by charter, then case-insensitive name + state, and clears a CU not present in the target period.
- Comparison extras, selected rivals, and trend comparisons now use stable identity/reset behavior rather than assuming row indices persist across quarters.
- Lower-is-better percentile bars now use the adjusted percentile for width, color, and label; empty-peer ranks show N/A.

Reusable source verification is in [audit_quarter_sources.py](tools/audit_quarter_sources.py). Build: `npm.cmd run build` passes. No frontend or audit-script diagnostics are reported.

## F. Filter Validation

| Component | Filter tested / reviewed | Result | Remaining issue |
| --- | --- | --- | --- |
| CU search + Overview | Affinity NJ selected | Q1 and Q2 values match the corresponding source row; stable charter identity retained across quarter switch. | ROA/ROE annualization and average-balance rules remain unresolved. |
| Quarter selector | Q1 -> Q2 -> Q1 | Overview returned to $4.35B then $4.40B and quarter-specific ratios; no stale CU values after fixes. | Q2 is not auto-loaded/persisted on app restart. |
| Peer group / benchmarks | Active quarter, selected asset bucket, optional same-state and adjacent-tier options | Peer calculation excludes selected CU and recomputes from `ACTIVE_CUS`; medians/ranks changed with Q2 row count. | No global State/Asset Segment filters exist; filters are local to peer/competitive views. |
| Competitive Intelligence | State, similar-size, nationwide scopes; named-rival shortlist | Q1 NJ scope: 127 institutions; Q2 NJ scope: 126. Asset/member shares, ranks, efficiency median update with active period. | Share is CU-universe share, not share of all financial institutions or addressable market. |
| Demographic map | Selected CU and active period | State/member totals rebuild from the active dataset; Q2 Affinity membership is 248,891 vs Q1 245,563. | Map intentionally aggregates the active statewide/national CU universe; it does not apply peer-tier filters or show branch locations. |
| Risk Alerts | Q2 active period | Rules recalculate from Q2; Affinity shows thin-earnings and loan-to-share flags under current thresholds. | Add an explicit >3% delinquency rule and review >100% vs >120% loan-to-share severity. |
| Trend Comparison | Both quarters loaded | Shows Q1/Q2 values and period-specific peer medians; matching uses stable institution identity. | “Change over period” is an absolute delta, not `((current - previous) / previous) x 100`; no QoQ percentage KPI or gainers/decliners is implemented. |
| Comparison tabs | Period changes | Side-by-side extras reset on active dataset changes; trend rival uses charter/name-state identity. | CU selection, state and peer filters are not unified into one global filter model. |

## Required Follow-up

1. Load/persist Q1 and Q2 at backend startup; do not make users manually re-import Q2 after each restart.
2. Add a validated, quarter-aware growth KPI using matched `CU_NUMBER`, explicit prior-period denominator rules, and N/A handling for zero/missing prior values.
3. Define YTD annualization and average-balance inputs for ROA, ROE, and net charge-offs; rename `Portfolio Yield` until a true loan-yield formula is implemented.
4. Preserve null/missing values and suppress ratios/alerts when required inputs or denominators are absent.
5. Reconcile duplicated loan-category source headers with the data provider before adding a portfolio-composition visual. Resolve the `NET_WORTH_AMOUNT` source anomaly with the provider rather than relying indefinitely on the equity-component override.
6. Add automated unit tests for header mapping, unit scaling, equity fallback, ratio formulas, quarter switching, and comparison identity. The repository currently exposes a production build but no frontend test script.
