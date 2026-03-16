# Testing

## Run all tests

```bash
npm test
```

Run in band (serial) to avoid DB path conflicts:

```bash
npm test -- --runInBand
```

## Test suites

- **db-path.test.js** – DB path resolution (shared vs env); no app DB required.
- **db.test.js** – DB API (getBusinessById, businessSettings, getDocuments, getMergeFields, getMaxInvoiceNumber). Uses the **shared app DB**; run the app once so the DB exists and is seeded.
- **invoice-totals.test.js** – Invoice totals formula (balance_due, discount/received as negative); unit only, no DB.
- **e2e-mcms-invoice.test.js** – Full invoice creation (template fill, line items, totals). Two tests are **skipped** in Jest due to an ExcelJS/readable-stream `objectMode` bug in this environment. To verify E2E: run the app and create an invoice manually, or run the app once so the shared DB is seeded and re-enable the tests if your environment doesn’t hit the bug.

## Benchmark behaviour

See [docs/BEHAVIOUR.md](./BEHAVIOUR.md) for the documented app behaviour that tests and refactors must preserve.
