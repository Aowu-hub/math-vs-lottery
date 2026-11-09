# Math vs. Lottery — Setup (2026-10-18 update)

## After pasting the new Code.gs + index.html

| Step | Run / do | What it does | Time |
|---|---|---|---|
| 1 | `setupSheets()` | Adds the `auto_tune` row to Config (nothing overwritten) | seconds |
| 2 | `buildMathBacktest()` | Fills the **Math Backtest** tab: for every past draw, what the Every-Draw and Same-Date columns would have picked using only earlier draws | ~10–60 s |
| 3 | `tuneMath()` | Tests 200 setting combos (incl. pure random) on every past draw → writes the **Math Tuning** tab; the wheel uses the winner | ~10–30 s |
| 4 | `backfillFakeVsRealComparisons()` | Fills any missing Lucky Draw vs Real rows | seconds |
| 5 | Deploy → Manage deployments → ✏️ → **New version** → Deploy | Publishes to the same /exec link | — |

After this, the daily trigger keeps everything up to date by itself: archive real draw → compare Lucky Draw → add backtest row → re-tune.

## New tabs

| Tab | Written by | Read by |
|---|---|---|
| Math Backtest | `buildMathBacktest()` + daily trigger | Aggregate Stats (🧮 / 📅) and Wall of Fame |
| Math Tuning | `tuneMath()` + daily trigger | Readable report; the winner drives the wheel |

## Config switches

| Key | Value | Meaning |
|---|---|---|
| auto_tune | 1 (default) | Wheel uses the best-tested method (today: pure random) |
| auto_tune | 0 | Wheel uses halflife_draws / prior_strength / sharpness from Config |
