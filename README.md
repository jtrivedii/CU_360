# CU360 Intelligence — enterprise local architecture

This migration preserves the supplied dashboard's UI, CSS tokens, charts, tab labels, and analytics functions in `frontend/src/pages/Dashboard.jsx`, while moving the seed data and AI-provider calls behind a FastAPI API. The bundled Excel workbook is the current source of truth. Snowflake and real LLM providers are adapter placeholders/configurable integrations.

## Source analysis

The original JSX is a 2,566-line single-file React application with an embedded dataset, inline CSS, chart and analytics components, and administrative/import logic. Major user-facing areas: Home/search; Overview; Peer Benchmarking; Competitive Intel; Compare Institutions; Trend Comparison; Demographic Views; Gap Analysis; Spend Analysis; Risk Alerts; Board Report; AI Recommendations; Connect Admin; floating chat assistant; theme toggle; reporting-quarter selector; import/mapping and connection controls. Charts use Recharts; icons use lucide-react; spreadsheet handling uses SheetJS. Existing analytics include peer percentiles/medians, risk flags, comparative institution ranking, state-level summaries, multi-quarter trends, board reporting, and operational ratios.

## Architecture

- `frontend/`: Vite + React; dashboard source preserved for a low-risk first migration; API base URL is environment-configurable.
- `backend/app/api`: REST routes.
- `backend/app/services`: orchestration.
- `backend/app/repositories`: Excel implementation and Snowflake placeholder behind a repository interface.
- `backend/app/llm`: provider selection, prompts, mock, OpenAI and Azure OpenAI adapters.
- `backend/data/cu360_data.xlsx`: generated from the exact embedded 2026 Q1 rows in the supplied JSX.

## Run in VS Code (no Docker)

1. Install Python 3.10+ and Node.js 20+ (Node is required by Vite). Open this folder in VS Code.
2. Backend terminal:
   - `cd backend`
   - Windows: `py -m venv .venv` then `.venv\Scripts\activate`
   - macOS/Linux: `python3 -m venv .venv` then `source .venv/bin/activate`
   - `pip install -r requirements.txt`
   - Copy `.env.example` to `.env` if needed; defaults use Excel and Mock LLM.
   - `uvicorn app.main:app --reload --host 127.0.0.1 --port 8000`
3. Frontend in a second terminal:
   - `cd frontend`
   - `npm install`
   - copy `.env.example` to `.env` if using a non-default API URL
   - `npm run dev`
4. Open http://127.0.0.1:5173. API docs: http://127.0.0.1:8000/docs.

## API routes

- `GET /health` — health and configured adapters.
- `GET /api/institutions` — field metadata, row data, quarter, count.
- `GET /api/institutions/{name}` — exact institution lookup.
- `POST /api/llm/chat` — assistant chat, Anthropic-compatible response envelope for UI compatibility.
- `POST /api/llm/recommendations` — AI recommendation generation.

## Configuration

`DATA_SOURCE=excel` is the supported current implementation. `DATA_SOURCE=snowflake` selects the explicit placeholder, which raises a clear not-implemented error until the repository is implemented and credentials/query behavior are validated. `LLM_PROVIDER=mock` is the default; set `openai` or `azure_openai` only after configuring secrets in backend `.env`. Never place provider credentials in Vite environment variables.

## Notes and boundaries

- The original visual design and analytics calculations are kept together intentionally for the first safe migration; moving each component/helper to its own module should be done in tested increments rather than rewriting a 2,500-line app in one step.
- Responsive CSS is appended as layout-only overrides and uses the existing palette. Validate against screenshots at target display sizes and browser zoom levels before production sign-off.
- The original admin spreadsheet-preview/mapping flow still uses the existing SheetJS client-side parser. The main dashboard dataset and LLM requests use FastAPI. To meet a strict “no file bytes processed in the browser” policy, the admin upload/preview/mapping flow should be migrated to backend upload/preview/commit endpoints as a follow-on; it has been retained here to avoid silently removing existing import UX.
- No authentication/authorization hardening, rate limiting, audit trail, or production secrets management is implied by this local starter. Add these before deployment.
