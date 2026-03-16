# MCMS Invoice Manager

Desktop app (Electron) for **Motti Cohen Music Services**: create and manage invoices from an Excel template, track invoice numbers, and store client details. PDFs are generated via Excel and saved to a configurable folder.

## Development

1. Install dependencies: `npm install`
2. Run in dev mode (live reload): `npm run dev` or use the double-click launcher in **Motti - run this for dev mode**
3. Build for production: `npm run build`
4. Pack macOS app: `npm run pack:mcms` (output in `release/`)

## Features

- **Invoices**: New invoice from template with placeholders filled (client, dates, line items, totals, discount, amount received). Edit and regenerate existing invoices. Staging file strategy keeps Excel file-access prompts to a minimum.
- **Templates**: Set Excel template and save folder per business. Template file is watched; changes are copied to a staging file automatically.
- **Clients**: Type-ahead client selection with optional new contact creation.
- **Invoice log**: List documents, view/open PDFs, reveal in Finder, mark paid/unpaid, delete.

## Tests

- Unit / E2E: `npm test`
- E2E invoice generation: `npm run test:e2e:mcms`
