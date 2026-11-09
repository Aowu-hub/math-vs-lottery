# Math vs. Lottery

Thai Government Lottery statistics app — Google Apps Script + Google Sheets
web app (mobile-first). Private repo: source code + history. Users keep
opening the same Apps Script `/exec` link.

| Folder / file | What it is |
|---|---|
| `src/Code.gs` | Server: math engine, draw calendar, backtest, tuning, buzz, dreams, Fake vs Real |
| `src/MasterKey.gs` | 🔑 Master Key: locks the private links, key-protected data door (`doPost`) |
| `src/index.html` | The app page (roulette, results, stats, Wall of Fame) |
| `src/master.html` | Master Key page — `…/exec?page=master` |
| `src/appsscript.json` | Apps Script settings (Bangkok time, V8, web app "Anyone") |
| `secrets/master-link.json` | Private links, locked with the Master Key |
| `tools/` | `check` · `unlock` · `deploy` · `backup` · Master Key test |
| `SETUP.md` | Sheet setup steps (setupSheets, buildMathBacktest, tuneMath, …) |

## 🔑 Master Key (works like the STEAM family key)

Nothing that leads to the data is stored readable: the `/exec` link,
deployment ID, Sheet ID and project ID are locked in the browser with the
Master Key. Only the locked copy + a hash of the key are kept.

1. Apps Script editor → run **`startMasterKeySetup()`** → copy the setup code from the log (valid 30 min).
2. Open **`…/exec?page=master`** → type the code → **Unlock** → check the app link → type a new Master Key (3+ words) twice → **Lock**.
3. Drive now has **`Master Link.json`** next to the Sheet → it goes into `secrets/master-link.json` here.

Forgot the key? Run `startMasterKeySetup()` again (only the Google account owner can).

## Everyday commands (on your computer, once Node 18+ is installed)

| Command | Does |
|---|---|
| `npm install` | once: installs `clasp` |
| `npx clasp login` | once: sign in to Google |
| `npm run unlock` | type the Master Key → prepares `clasp` (files are git-ignored) |
| `npm run check` | checks: scripts parse, no one-line `//` scripts, no duplicate functions, no private links in code |
| `npm run deploy -- "what changed"` | check → push to Apps Script → update the **same** `/exec` deployment |
| `npm run backup` | Master Key → full copy of every Sheet tab into `backups/` (git-ignored) |

GitHub runs `check` + the Master Key test on every push.

## 🚀 Auto-deploy (every update to `src/` goes live by itself)

`.github/workflows/deploy.yml` runs on every push to `main` that changes `src/`:
checks + Master Key test → **safety check** (downloads the live project; stops,
changing nothing, if it has a file the repo doesn't or different
`appsscript.json` settings) → `clasp push` → updates the **same** `/exec` deployment.

It needs 2 repository secrets (Settings → Secrets and variables → Actions):

| Secret | Value |
|---|---|
| `CLASPRC_JSON` | the whole text of `.clasprc.json` made by `npx @google/clasp@2.4.2 login` |
| `MVL_MASTER_KEY` | the Master Key |

Without them the workflow just says "Auto-deploy is off". Run it by hand: Actions → deploy → Run workflow.
Rollback: revert the bad commit on GitHub → it deploys the old code by itself.
