# MCMS Invoice Manager – Documented Behaviour (Benchmark)

This file is the **single source of truth** for current app behaviour. Any refactor or slim must preserve these behaviours. Tests encode the same; this doc is for humans and for ensuring nothing is lost.

---

## 1. Invoice generation (`createMCMSInvoice`)

- **Entry:** `documentService.createMCMSInvoice(options)` with `business_id`, `template_path` or definition, `save_path` or business `save_path`, `line_items`, `client_override`, dates, totals, optional `invoice_number`, `discount_amount`, `amount_received`, etc.
- **Output:** Excel workbook at staging path (`_staging.xlsx` in save dir) or a copy; PDF via Excel; DB row in `documents` with `doc_type = 'invoice'`, `file_path`, `invoice_snapshot` (JSON of form data). Optional: delete `.xlsx` after PDF (when not `_e2eKeepWorkbook`).
- **Numbering:** Next number from `business_settings.last_invoice_number`; can override with `options.invoice_number`. Phantom rows (same number, no file on disk) are purged before assign.

---

## 2. Placeholders and template fill

- **Tokens:** `{{token}}` in cells. Mid-cell: replace in place (e.g. `Invoice {{invoice_code}}` → `Invoice INV-123`). Value-only cells: replace whole cell with value; **no leading "0"** (e.g. `0 {{subtotal}}` or `{{subtotal}}` → numeric subtotal only).
- **Dates:** All date tokens formatted **dd/mm/yyyy** (`formatDateDDMMYYYY`).
- **Currency:** Line amounts and totals use currency formatting from template; no forced override of alignment (template-driven).
- **Line items (repeatable rows):** `{{item_date}}`, `{{item_description}}`, `{{item_amount}}` (and similar) filled **per row** in `writeRepeatableItemRows`. Row 1 gets item 1, row 2 gets item 2, etc. **Critical:** Row 2 description must be exactly the second item’s description (no "te:", no duplicate text, no corruption from merged cells). Only the primary description column is written per row; other columns in the same row are not cleared (to avoid merged-cell corruption).

---

## 3. Totals and maths

- **Subtotal:** Sum of line item amounts, or `options.total_amount` if provided (used as override).
- **Discount:** Stored and displayed as **negative** (`-Math.abs(discount_amount)`). Only applied when `discount_amount` is set and non-zero.
- **Received:** Stored and displayed as **negative** (`-Math.abs(amount_received)`). Label cell shows the word "Received"; value cell shows the negative amount.
- **Balance due:** `balance_due = max(0, subtotal - discount - received)`. Formula: subtotal minus (absolute discount + absolute received).
- **Labels:** Template may use `{{subtotal_label}}` / `{{received_label}}` for the words "Subtotal" / "Received"; values go in `{{subtotal}}` / `{{received}}` cells only.

---

## 4. Conditional rows (workbook)

- **Discount row:** If no discount (`discount_amount` not set or zero), the row containing `{{discount_amount}}` or `{{discount_description}}` is **removed** from the workbook so it does not appear on the PDF.
- **Received row:** If no amount received, the row containing `{{received}}` is **removed** so "0" or empty received does not show.

---

## 5. Staging and Excel

- **Staging file:** Fixed path `_staging.xlsx` in the save directory so Excel only ever opens one path per folder; minimizes "Grant Access" prompts on macOS.
- **Flow:** Template is copied to staging (or template path is used); workbook is filled and written to staging (or a temp copy for new invoices); PDF is generated from that file. On first run after a template update, Excel may ask for access once.
- **Watcher:** When the template file on disk changes, the app can update the staging copy so staging stays in sync with the template.

---

## 6. Database

- **Business:** Single business (id 1) for MCMS. `business_settings` holds `save_path`, `last_invoice_number`; no deletion of existing business rows on seed.
- **Documents:** `documents` table has `business_id`, `doc_type`, `number`, `file_path`, `invoice_snapshot` (TEXT JSON), `jobsheet_id` (nullable, unused for MCMS). Invoice log shows rows where `doc_type` (case-insensitive) is `invoice` and `business_id` is 1 **or** 2 (legacy).
- **dbReady:** All reads/writes that the UI depends on (e.g. `getDocuments`, `getDocumentById`, `businessSettings`, `updateBusinessSettings`, `getBusinessById`) are gated on `dbReady` so they run only after DB init and migrations.
- **Migration:** Existing DBs with `documents` table may be migrated to drop jobsheet FK; data is copied to new tables then renamed (no data loss). New DBs create tables without jobsheet FKs.

---

## 7. Contacts and UI

- **Contacts:** CRUD for clients; client picker for invoice modal; client details (emails, phones, addresses) stored and editable.
- **Invoice log:** List of invoices (business_id 1 or 2, doc_type invoice); optional filter by existing files (`includeMissing: true` so all DB rows show; `file_available` can indicate missing file).
- **Templates tab:** Template path, save folder, staging watcher status; DB path displayed for debugging.

---

## 8. Edit / Regenerate

- **Edit:** Load existing document; restore form from `invoice_snapshot` or from file (line items from Excel). Save updates snapshot and regenerates file (overwrites).
- **Regenerate:** Same as edit: overwrite existing PDF/Excel with new output; same invoice number.

---

## 9. Files and paths

- **Save path:** User-chosen directory for invoices (business_settings.save_path).
- **Template path:** Per definition (e.g. `invoice_balance`); set in Templates tab.
- **Reveal in Finder:** After PDF generation, app can reveal the output file in Finder (no Quick Look required).

---

## 10. What must not regress

- Line 2 description exactly as entered (no truncation, no "te:", no duplicate).
- Subtotal cell: numeric value only, no leading "0" or "0 £1,525.00".
- Balance due = subtotal − discount − received (all non-negative).
- Discount and received values negative in workbook.
- Received **label** shows the word "Received"; value cell shows the negative amount.
- Dates in dd/mm/yyyy.
- Mid-cell tokens (e.g. `Invoice {{invoice_code}}`) keep surrounding text.
- No data loss on migration or on app restart (same DB path for dev and packaged).
- Invoice log shows all invoices (business_id 1 or 2) and does not disappear when files are missing (includeMissing: true).
