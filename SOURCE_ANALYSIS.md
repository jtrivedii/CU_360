# Source analysis — CU360_Intelligence_updated_v9.jsx

## File characteristics
- 2,566 lines; ~1.91 MB.
- Single default-exported React `App`; UI components, helper functions, data import logic, analytics, global mutable active-quarter state, and CSS are in one file.
- 4,336 institutions in the embedded 2026 Q1 dataset; 35 fields. The project generator extracted those exact values into `backend/data/cu360_data.xlsx`.
- Stack: React hooks, Recharts, lucide-react, SheetJS.

## Feature inventory
- Shell/navigation, home landing and credit-union search.
- Overview KPIs and operating analytics.
- Peer benchmarking with peer group controls and sortable peer table.
- Competitive Intel and closest/large rival comparisons.
- Compare Institutions.
- Trend Comparison across reporting quarters.
- Demographic Views and state-level rollups/map.
- Gap Analysis and Spend Analysis.
- Risk Alerts / rule-based flags.
- Print-oriented Board Report.
- AI Recommendations and floating chat assistant.
- Connect Admin: quarterly Excel/CSV/JSON import, field mapping, multi-quarter registry, Snowflake connection form/status, pipeline and refresh policy.
- Theme switch, reporting-quarter selector, login/admin visibility controls.

## Architectural risks observed
- A large inline dataset makes the frontend bundle heavy and couples UI to data.
- Data state uses module-level mutable `ACTIVE_CUS`, `ACTIVE_QUARTER`, and `DATA_VERSION`.
- AI chat/recommendations previously called an external LLM endpoint directly from React. These are routed to FastAPI in this migration.
- The admin spreadsheet preview/normalization path runs in the browser and needs a backend upload/preview/commit migration to satisfy a strict server-only file-processing policy.
- Responsive CSS has several fixed/min widths and multi-column grids; supplemental breakpoints have been added without changing theme tokens.
- The original dashboard is not fully decomposed into components/services; low-risk preservation was prioritized to avoid unverified business-logic changes.
