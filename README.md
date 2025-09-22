# Invoice Master Dashboard

## Development Workflow

Use the live-reload setup when iterating on the Electron app:

1. Install dependencies with `npm install`.
2. Start the watcher-driven dev environment with `npm run dev`.
   - Webpack runs in watch mode for renderer changes.
   - Electron restarts automatically via `electronmon` when main or renderer bundles change.
3. The command performs an initial development build before launching, so the renderer bundle is always ready when Electron starts.

Use `npm start` for the production-style launch (it still performs a fresh production build first via `prestart`).

## TODO

- Add search and filtering controls to the jobsheet list so large account datasets stay manageable.
