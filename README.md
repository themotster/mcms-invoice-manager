# Invoice Master Dashboard

## Development Workflow

Use the live-reload setup when iterating on the Electron app:

1. Install dependencies with `npm install`.
2. Start the watcher-driven dev environment with `npm run dev`.
   - Webpack runs in watch mode for renderer changes.
   - Electron restarts automatically via `electronmon` when main or renderer bundles change.
3. The command performs an initial development build before launching, so the renderer bundle is always ready when Electron starts.

Use `npm start` for the production-style launch (it still performs a fresh production build first via `prestart`).

## Features

- Jobsheet list supports search, status filters, and sorting.

## Roadmap

- In review / follow-ups
  - Per-business template path settings UI. Current flow uses Templates Manager to set per-definition templates; decide if Settings needs separate defaults for invoice/quote/contract.
  - Import existing job files to auto-create jobsheets (partial: invoice filename importer exists; extend to create jobsheets when missing).

- Backlog
  - Update documentation for the hybrid document workflow and common flows.
  - Calendar of upcoming deadlines/reminders with background macOS notifications.
  - Outbound email to send generated documents to clients (templated, with attachments).
  - WhatsApp workflow for AhMen enquiries, gig sheets, and personnel follow-ups.
  - Drag & drop external documents into a jobsheet’s folder so the job folder becomes a complete container of all related files (future enhancement).

## One-off Migration: Import jobsheets from folders

Use the temporary script to create jobsheets from existing folders named like:

- `YYYY-MM-DD - Client Name - 14 June 2025`

Commands:

- Dry run (no writes):
  - `npm run migrate:import-jobsheets`
- Apply to a specific business id 2, link invoices from filenames:
  - `node scripts/import_jobsheets.js --business 2 --link-invoices`

Flags:

- `--business <id>` limit to one business
- `--dry-run` preview only
- `--max-depth <n>` directory depth (default 3)
- `--link-invoices` also import invoice PDFs (INV-###) using the existing indexer

This script is intended for one-time use and can be removed afterward.
