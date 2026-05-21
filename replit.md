# Audtech Reconciliation

A financial reconciliation web application for auditors that compares Prior Year Signed Financial Statement balances with Current Year Opening Trial Balance balances.

## Run & Operate

- `pnpm --filter @workspace/api-server run dev` — run the API server (port 5000)
- `pnpm --filter @workspace/audtech run dev` — run the frontend (Vite)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from the OpenAPI spec

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5 + Multer (file upload) + xlsx (Excel parsing & export)
- Frontend: React 19, Vite 7, Tailwind CSS 4, Recharts, shadcn/ui
- API codegen: Orval (from OpenAPI spec)
- Build: esbuild (CJS bundle)
- No database required (stateless reconciliation engine)

## Where things live

- `lib/api-spec/openapi.yaml` — API contract source of truth
- `artifacts/api-server/src/routes/reconciliation.ts` — core reconciliation logic (parse, fuzzy match, reconcile, export)
- `artifacts/audtech/src/pages/Reconciliation.tsx` — full frontend UI
- `artifacts/audtech/src/index.css` — theme and CSS variables

## Architecture decisions

- Reconciliation is fully stateless on the backend; results are computed on upload and held in memory (no DB needed for MVP).
- File parsing supports both Debit/Credit columns (net = Dr − Cr) and single Balance/Amount columns.
- Fuzzy matching uses Jaccard similarity on normalized ledger names with a 60% threshold for "possible regroup" classification.
- Export produces a multi-sheet Excel file: Summary, full Detail, and one tab per reconciliation status.
- File upload uses the raw `/api/reconciliation/upload` endpoint (multipart) outside the OpenAPI codegen path to avoid Blob/File type issues in the Zod server validator.

## Product

- Upload two Excel/CSV files (Prior Year FS + Current Year Opening TB)
- Auto-detect column headers (code, name, debit/credit or balance)
- Classify each ledger as: Matched, Mismatched, Missing in CY, New in CY, or Possible Regroup
- Summary KPI cards, pie chart, and bar chart dashboard
- Filterable, sortable, searchable reconciliation detail table
- One-click Export to multi-sheet Excel report

## User preferences

- Project language: Python was requested but full-stack React/Node.js was used for the web app (richer UX). The reconciliation engine logic is implemented server-side in TypeScript/Node.js.

## Gotchas

- The `DATABASE_URL` env var is NOT required — this app has no DB.
- `multer` and `xlsx` must be in `dependencies` (not devDependencies) of `api-server` since they are used at runtime.
- `@types/multer` is a devDependency of `api-server`.

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
