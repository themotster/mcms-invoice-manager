# Run MCMS Invoice Manager in dev mode

Fully self-contained—no terminal commands needed. (Any other commands are run by the agent; you never need to run terminal commands yourself.)

**What it does:**
1. Loads your shell profile so `npm` is found (needed when double-clicking).
2. Stops any running MCMS Electron window from this project (avoids duplicate windows).
3. Runs `npm run mcms:dev` (build + webpack watch + Electron). The app window opens when ready; changes auto-reload.

**How to run:**
- **Double-click `run-dev.command`** in Finder — Terminal opens and starts the app. Use this.

**Important:** Keep the Terminal window open while using the app. If the app doesn’t open, check the Terminal for error messages; the window will stay open so you can read them.
