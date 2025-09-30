#!/usr/bin/env node
/*
 One-off migration script: import jobsheets from existing folders.
 Pattern: "YYYY-MM-DD - Client Name - Human Date"

 Usage:
  node scripts/import_jobsheets.js [--business <id>] [--dry-run] [--max-depth <n>] [--link-invoices]

 Notes:
  - Only creates new jobsheets when none exist with same (business_id, event_date, client_name).
  - Optional: link invoices via existing filename indexer (PDFs with INV-###).
*/

const fs = require('fs');
const path = require('path');

const db = require('../db');
let docService = null;
try { docService = require('../documentService'); } catch (_) { docService = null; }

function parseArgs(argv) {
  const args = { business: null, dryRun: false, maxDepth: 3, linkInvoices: false };
  const list = argv.slice(2);
  for (let i = 0; i < list.length; i++) {
    const a = list[i];
    if (a === '--business' || a === '-b') { args.business = Number(list[++i]); }
    else if (a === '--dry-run' || a === '-n') { args.dryRun = true; }
    else if (a === '--max-depth' || a === '-d') { args.maxDepth = Number(list[++i]); }
    else if (a === '--link-invoices') { args.linkInvoices = true; }
  }
  if (!Number.isInteger(args.maxDepth) || args.maxDepth < 1) args.maxDepth = 3;
  if (args.business != null && !Number.isInteger(args.business)) args.business = null;
  return args;
}

function parseFolderName(name) {
  // Expected: YYYY-MM-DD - Client Name - Human date
  const re = /^(\d{4}-\d{2}-\d{2})\s*-\s*([^\-]+?)(?:\s*-\s*.+)?$/;
  const m = name.match(re);
  if (!m) return null;
  const eventDate = m[1];
  const clientName = m[2].trim();
  if (!clientName) return null;
  return { event_date: eventDate, client_name: clientName };
}

async function walkDirs(root, maxDepth = 3) {
  const out = [];
  async function walk(current, depth) {
    if (depth > maxDepth) return;
    let entries;
    try { entries = await fs.promises.readdir(current, { withFileTypes: true }); } catch (_) { return; }
    for (const e of entries) {
      if (e.name.startsWith('.')) continue;
      const full = path.join(current, e.name);
      if (e.isDirectory()) {
        out.push(full);
        await walk(full, depth + 1);
      }
    }
  }
  await walk(root, 1);
  return out;
}

function normalizeName(s) {
  return (s || '').toString().trim().toLowerCase();
}

async function main() {
  const args = parseArgs(process.argv);

  const businesses = await db.businessSettings();
  const targets = Array.isArray(businesses)
    ? businesses.filter(b => (args.business ? b.id === args.business : true))
    : [];
  if (!targets.length) {
    console.log('No businesses found (or filter excluded all).');
    process.exit(0);
  }

  let createdCount = 0;
  let skippedExisting = 0;
  const createdByBusiness = new Map();

  for (const biz of targets) {
    const root = (biz.save_path || '').trim();
    if (!root) { console.warn(`Business ${biz.id} has no save_path; skipping`); continue; }
    const exists = fs.existsSync(root);
    if (!exists) { console.warn(`Path does not exist: ${root}`); continue; }

    const existingSheets = await db.getAhmenJobsheets({ businessId: biz.id });
    const existingKey = new Set(
      existingSheets.map(s => `${normalizeName(s.client_name)}|${s.event_date || ''}`)
    );

    const dirs = await walkDirs(root, args.maxDepth);
    const candidates = [];
    for (const dir of dirs) {
      const base = path.basename(dir);
      const parsed = parseFolderName(base);
      if (!parsed) continue;
      const key = `${normalizeName(parsed.client_name)}|${parsed.event_date}`;
      if (existingKey.has(key)) {
        skippedExisting += 1;
        continue;
      }
      candidates.push({ dir, ...parsed });
    }

    if (!candidates.length) {
      console.log(`Business ${biz.id}: no new candidates.`);
      continue;
    }

    console.log(`Business ${biz.id}: ${candidates.length} new candidate jobsheets.`);

    if (!args.dryRun) {
      for (const c of candidates) {
        try {
          const id = await db.addAhmenJobsheet({
            business_id: biz.id,
            status: 'contracting',
            client_name: c.client_name,
            event_date: c.event_date
          });
          createdCount += 1;
          if (!createdByBusiness.has(biz.id)) createdByBusiness.set(biz.id, []);
          createdByBusiness.get(biz.id).push({ id, ...c });
        } catch (err) {
          console.warn('Failed to create jobsheet for', c.dir, err.message || err);
        }
      }
    } else {
      console.log('(dry-run) Would create:');
      candidates.forEach(c => console.log(`  - ${c.event_date} · ${c.client_name} (${c.dir})`));
    }

    if (!args.dryRun && args.linkInvoices && docService && typeof docService.indexInvoicesFromFilenames === 'function') {
      try {
        const res = await docService.indexInvoicesFromFilenames({ businessId: biz.id });
        console.log(`Linked invoices from filenames: imported ${res?.imported || 0}`);
      } catch (err) {
        console.warn('Invoice indexer failed', err.message || err);
      }
    }
  }

  console.log(`\nSummary: created ${createdCount} jobsheets; skipped ${skippedExisting} existing.`);
  if (args.dryRun) {
    console.log('Dry run complete. Re-run without --dry-run to apply changes.');
  }
}

main().then(() => process.exit(0)).catch(err => {
  console.error(err);
  process.exit(1);
});

