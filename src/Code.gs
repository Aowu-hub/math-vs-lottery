/**
 * ==========================================================================
 * "Math vs. Lottery" — Thai Lottery Number Insight Generator
 * Backend: Apps Script (Code.gs)
 *
 * Math engine v2 (recency-weighted, Bayesian-smoothed digit + whole-number
 * model, chi-square evidence) and a DRAW-DAY LOCK: nothing new is generated
 * on the 1st, the 16th, or any moved draw date listed in the Draw Calendar
 * tab — the next period opens the following day.
 *
 * Fun extras: Flips (กลับเลข), Dreams (ทำนายฝัน — Dream Dictionary tab) and
 * Buzz numbers from Thai media (fetchBuzzNumbers(), needs GEMINI_API_KEY).
 *
 * After pasting this version: run setupSheets() once (adds Config math rows,
 * Draw Calendar, Dream Dictionary + Dream Log tabs). Nothing existing is
 * overwritten.
 *
 * This file is entertainment statistics, not a prediction system. Thai
 * lottery draws are independently random every time — the footer disclaimer
 * belongs in the frontend (Phase 5).
 * ==========================================================================
 */

// --------------------------------------------------------------------------
// CONFIG
// --------------------------------------------------------------------------

const TABS = {
  HISTORICAL: 'Historical Draws',
  SIGNAL_FEED: 'Signal Feed',
  GENERATED_LOG: 'Generated Log',
  RECOMMENDATION_NOTES: 'Recommendation Notes',
  DAILY_CACHE: 'Daily Cache',
  CONFIG: 'Config',
  FAKE_VS_REAL: 'Fake vs Real',
  DRAW_CALENDAR: 'Draw Calendar',
  DREAM_DICTIONARY: 'Dream Dictionary',
  DREAM_LOG: 'Dream Log',
  MATH_BACKTEST: 'Math Backtest',
  MATH_TUNING: 'Math Tuning',
};

// --------------------------------------------------------------------------
// MATH SETTINGS (v2) — editable from the Config tab (same two columns as the
// weights: key in col A, number in col B). Missing rows fall back to these.
//
//  halflife_draws  Recency weighting. A draw this many draws old counts half
//                  as much as the newest one. 48 ≈ 2 years. 0 = all equal.
//  prior_strength  Bayesian smoothing (Dirichlet prior) in pseudo-draws per
//                  digit position. Pulls noisy small-sample counts back
//                  toward the fair 10% baseline so 9 rows of "same date"
//                  data can't produce wild 40% digits. Higher = calmer.
//  sharpness       How hard the random pick leans toward favored numbers.
//                  1 = sample exactly by the math's probabilities,
//                  2 = squares them (leans harder), 0 = pure uniform random.
// --------------------------------------------------------------------------
const SIGNAL_MAX_SHARE = 0.2; // hand-typed signals can move the math by at most 20%

const MATH_DEFAULTS = {
  auto_tune: 1,          // 1 = the wheel uses the best settings found by tuneMath(); 0 = use the values below
  halflife_draws: 48,
  prior_strength: 30,
  sharpness: 1.5,
};

// Default weights. This is now a MATH draw, not a signal/buzz draw — the
// Signal Feed tab (manually-pasted news/viral/birthday/plate numbers) turned
// out to be impractical to keep populated with real data reliably, so out of
// the box every generation leans 100% on real historical draw frequency.
// The signal categories are kept at 0 by default (not removed) so the Config
// tab still has the option to turn them back on later if the Signal Feed
// ever does get populated — calculateWeights() normalizes whatever is in
// Config, so these don't need to sum to exactly 1.
const DEFAULT_WEIGHTS = {
  historical: 1.00,
  news: 0,
  viral: 0,
  birthday: 0,
  carplate: 0,
  events: 0,
};

// Maps each weight category to the Signal Feed "Source Type" value(s) it
// reads — a category can cover more than one Source Type string, which is
// how "events" bundles Event / Disaster / Pop-up Event / Crisis under one
// shared weight instead of needing a separate slider for each.
const SOURCE_TYPE_BY_CATEGORY = {
  news: ['News'],
  viral: ['Social Buzz'],
  birthday: ['Birthday'],
  carplate: ['Car Plate'],
  events: ['Event', 'Disaster', 'Pop-up Event', 'Crisis'],
  // "Anniversary" / "Other" rows are not weighted by default; add a row to
  // SOURCE_TYPE_BY_CATEGORY (and Config) if you want to score them too.
};

// Every valid Source Type value, in the order they'll appear in the Signal
// Feed "Source Type" dropdown (data validation) — includes both the
// weighted categories above and the two unweighted catch-alls.
const ALL_SOURCE_TYPES = [
  'News', 'Social Buzz', 'Car Plate', 'Birthday',
  'Event', 'Disaster', 'Pop-up Event', 'Crisis',
  'Anniversary', 'Other',
];

// Number-set definitions: how many digits, which Historical Draws column(s)
// feed the position-specific historical frequency table, and whether two
// distinct numbers are needed.
const SET_DEFS = {
  jackpot: { label: 'รางวัลที่ 1', labelEN: '1st Prize (Jackpot)', len: 6, historicalCols: ['Prize1 (6-digit)'], pairCount: 1 },
  front3: { label: 'เลขหน้า 3 ตัว', labelEN: 'Front 3-Digit', len: 3, historicalCols: ['Front3-A', 'Front3-B'], pairCount: 2 },
  back3: { label: 'เลขท้าย 3 ตัว', labelEN: 'Back 3-Digit', len: 3, historicalCols: ['Back3-A', 'Back3-B'], pairCount: 2 },
  back2: { label: 'เลขท้าย 2 ตัว', labelEN: 'Back 2-Digit', len: 2, historicalCols: ['Back2'], pairCount: 1 },
};

// --------------------------------------------------------------------------
// ONE-TIME SETUP — run setupSheets() once from the Script Editor (select the
// function in the dropdown, click Run). Safe to re-run: it only creates
// tabs/headers that don't already exist, never overwrites data.
// --------------------------------------------------------------------------

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // textCols = 1-based column indexes forced to Plain Text ('@') format so
  // digit strings like "017" or "2026-08-14" never get silently reinterpreted
  // as a number/date and lose leading zeros — this is what caused the
  // "0 disappears" bug and the cache lookup mismatch seen during testing.
  ensureTab_(ss, TABS.HISTORICAL,
    ['Draw Date', 'Prize1 (6-digit)', 'Front3-A', 'Front3-B', 'Back3-A', 'Back3-B', 'Back2'],
    [1, 2, 3, 4, 5, 6, 7]); // Draw Date (col 1) forced text too, same reason
                            // as Daily Cache below — the real-draw auto-fetch
                            // string-compares dates to avoid duplicate rows.
  ensureHistoricalDrawsExtraColumns_(ss.getSheetByName(TABS.HISTORICAL));

  const signalSheet = ensureTab_(ss, TABS.SIGNAL_FEED,
    ['Date', 'Source Type', 'Raw Text', 'Extracted Number(s)', 'Notes'],
    [4]);
  applySourceTypeDropdown_(signalSheet);

  ensureTab_(ss, TABS.GENERATED_LOG,
    ['Timestamp', 'Jackpot', 'Front3-A', 'Front3-B', 'Back3-A', 'Back3-B', 'Back2',
     'Confidence %', 'Reasoning Text', 'Signal Snapshot Ref'],
    [2, 3, 4, 5, 6, 7]);

  ensureTab_(ss, TABS.RECOMMENDATION_NOTES, ['Date', 'Note']);

  ensureTab_(ss, TABS.DAILY_CACHE,
    ['Date', 'Jackpot', 'Front3-A', 'Front3-B', 'Back3-A', 'Back3-B', 'Back2',
     'Confidence %', 'Reasoning', 'Signal Snapshot', 'Generated At'],
    [1, 2, 3, 4, 5, 6, 7]); // Date (col 1) is forced text too — removes any
                            // ambiguity between a real Date object and a
                            // string when generateSet() checks "is there
                            // already a cached row for today?"

  const configSheet = ensureTab_(ss, TABS.CONFIG, ['Signal Category', 'Weight (%)']);
  const configLastRow = configSheet.getLastRow();
  const existingRows = configLastRow > 1
    ? configSheet.getRange(2, 1, configLastRow - 1, 2).getValues()
    : [];
  const existingKeys = existingRows.map(([category]) => String(category).trim().toLowerCase());
  const missingKeys = Object.keys(DEFAULT_WEIGHTS).filter(k => !existingKeys.includes(k));

  if (missingKeys.length > 0) {
    const newRows = missingKeys.map(k => [k, DEFAULT_WEIGHTS[k] * 100]);
    configSheet.getRange(configSheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
    Logger.log('Config: added missing weight categories -> ' + missingKeys.join(', '));
  }
  if (existingRows.length === 0) {
    configSheet.getRange(1, 4).setValue('Edit the % column any time — generateSet() re-reads it on every run (except cache hits).');
  }
  // Math v2 settings live in the same tab (key / value), appended once.
  const allKeysNow = configSheet.getLastRow() > 1
    ? configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase())
    : [];
  const missingMath = Object.keys(MATH_DEFAULTS).filter(k => !allKeysNow.includes(k));
  if (missingMath.length > 0) {
    configSheet.getRange(configSheet.getLastRow() + 1, 1, missingMath.length, 2)
      .setValues(missingMath.map(k => [k, MATH_DEFAULTS[k]]));
    configSheet.getRange(2, 4).setValue('Math settings: halflife_draws (recency), prior_strength (smoothing), sharpness (how hard picks lean on the math).');
  }

  // Draw Calendar — list any OFFICIAL draw date that is NOT the 1st/16th
  // (GLO moves some draws: e.g. 17 Jan, 2 May, 30 Dec). Those days get the
  // same "no generation today" lock as the 1st/16th.
  const calSheet = ensureTab_(ss, TABS.DRAW_CALENDAR, ['Moved Draw Date (yyyy-MM-dd)', 'Note'], [1]);
  if (calSheet.getLastRow() === 1) {
    calSheet.getRange(1, 4).setValue('Add GLO draw dates that are NOT on the 1st/16th (e.g. 2027-01-17). The 1st and 16th are always locked automatically.');
  }

  ensureTab_(ss, TABS.FAKE_VS_REAL, [
    'Draw Date', 'Real Jackpot', 'Fake Jackpot', 'Jackpot Matches (/6)',
    'Real Front3-A', 'Real Front3-B', 'Fake Front3-A', 'Fake Front3-B', 'Front3 Best Match (/3)',
    'Real Back3-A', 'Real Back3-B', 'Fake Back3-A', 'Fake Back3-B', 'Back3 Best Match (/3)',
    'Real Back2', 'Fake Back2', 'Back2 Matches (/2)',
    'Tries This Period', 'Fake Generated At', 'Compared At',
  ], [1, 2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 15, 16]);

  seedDreamDictionary_(ss);
  ensureTab_(ss, TABS.DREAM_LOG, ['Timestamp', 'Period Key', 'Dream', 'Matched', 'Numbers'], [2, 5]);

  Logger.log('Setup complete. Tabs verified/created: ' + Object.values(TABS).join(', '));
}

// Historical Draws started with only the 7 spec columns; the real-draw
// auto-fetch below adds two more (Source, Fetched At) to every row it
// writes. This retroactively adds those headers if they're missing, without
// touching any existing data rows.
function ensureHistoricalDrawsExtraColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  ['Source', 'Fetched At'].forEach(label => {
    if (headerRow.indexOf(label) === -1) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(label).setFontWeight('bold');
    }
  });
}

// headers: column headers for row 1 (only written once, if the tab is new).
// textCols: optional array of 1-based column numbers to force to Plain Text
// format across many rows, so numeric-looking strings never get silently
// reinterpreted by Sheets. Safe to call repeatedly — re-applies formatting
// even on tabs that already exist, so re-running setupSheets() retroactively
// fixes format on a sheet you built before this fix existed.
function ensureTab_(ss, name, headers, textCols) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0 && headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  if (textCols && textCols.length > 0) {
    const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
    textCols.forEach(col => {
      sheet.getRange(2, col, maxRows, 1).setNumberFormat('@');
    });
  }
  return sheet;
}

// Appends one row while GUARANTEEING the given columns stay Plain Text —
// fixes leading zeros (e.g. back2 "06", front3 "019") silently turning into
// a number and losing the leading digit. ensureTab_()'s column-wide '@'
// format only reliably covers rows that existed when it last ran; forcing
// the format on the SPECIFIC row right before writing values is what
// actually guarantees it survives, regardless of when/how the tab was
// created. Use this instead of sheet.appendRow(...) wherever a row contains
// digit-string values (jackpot/front3/back3/back2 and similar).
function appendRowForcedText_(sheet, rowValues, textCols) {
  const rowIndex = sheet.getLastRow() + 1;
  (textCols || []).forEach(col => {
    sheet.getRange(rowIndex, col).setNumberFormat('@');
  });
  sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
}

// Adds a dropdown (data validation) to the Signal Feed "Source Type" column
// so entries always match one of ALL_SOURCE_TYPES exactly — a typo here
// (e.g. "car plate" instead of "Car Plate") would silently fall out of the
// weighting entirely, since buildSignalDigitFrequency_ matches by exact
// string.
function applySourceTypeDropdown_(sheet) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ALL_SOURCE_TYPES, true)
    .setAllowInvalid(false)
    .build();
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 2, maxRows, 1).setDataValidation(rule);
}

// --------------------------------------------------------------------------
// WEIGHTS — read from Config tab, fall back to DEFAULT_WEIGHTS for any
// category missing or blank.
// --------------------------------------------------------------------------

function calculateWeights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.CONFIG);
  const weights = Object.assign({}, DEFAULT_WEIGHTS);

  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    data.forEach(([category, pct]) => {
      const key = String(category).trim().toLowerCase();
      if (key in weights && pct !== '' && !isNaN(pct)) {
        weights[key] = Number(pct) / 100;
      }
    });
  }

  // Normalize so the categories always sum to 1, even if the user's Config
  // percentages don't add up to exactly 100.
  const sum = Object.values(weights).reduce((a, b) => a + b, 0) || 1;
  Object.keys(weights).forEach(k => weights[k] = weights[k] / sum);

  return weights;
}

function readMathSettings_() {
  const settings = Object.assign({}, MATH_DEFAULTS);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.CONFIG);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(([k, v]) => {
      const key = String(k).trim().toLowerCase();
      if (key in settings && v !== '' && !isNaN(v)) settings[key] = Math.max(0, Number(v));
    });
  }
  return settings;
}

// ==========================================================================
// MATH ENGINE v2
//
// What changed vs v1 (raw per-position digit counts, sampled as-is):
//  1. ONE read of Historical Draws per request, parsed into clean draw
//     objects (zero-padded digit strings, sorted newest-first).
//  2. Recency weighting — each draw's weight halves every `halflife_draws`
//     draws, so the model leans on the current era of draws instead of
//     treating 2017 the same as last month.
//  3. Bayesian smoothing (Dirichlet prior) — counts are blended with the
//     fair 10%-per-digit baseline, so a thin sample can't produce wild
//     digit probabilities.
//  4. "Same date" is now a hierarchical estimate: the same-slot draws
//     (e.g. every 16 Oct draw) are shrunk toward the ALL-TIME estimate,
//     not toward uniform — 9 rows nudge, they don't dominate.
//  5. Same-slot matching uses the NOMINAL slot, so moved draws (17 Jan,
//     2 May, 30 Dec, 31 Jul, 15 Jul…) count toward their real 1st/16th
//     slot instead of silently dropping out of the filter.
//  6. Whole-number model for 2- and 3-digit sets: actual repeat counts of
//     each full number (00-99 / 000-999) on top of the digit model, so
//     "this exact number keeps coming back" is captured, not just digits.
//  7. Exact top-5 ranking (all 100/1000 numbers scored; jackpot uses a
//     5-wide beam, which is exact for the top 5 of a product model).
//  8. Honest evidence: a chi-square uniformity test per set tells you
//     whether the real data deviates from pure random at all.
//  9. "Lift" (× vs pure random) replaces the meaningless relative %.
// 10. Math Score (0-100) replaces the old amplified "confidence": the
//     average percentile rank of the picked numbers inside the model —
//     50 = an average number, 100 = the math's absolute favorite.
// ==========================================================================

// --------------------------------------------------------------------------
// DATE HELPERS (all yyyy-MM-dd strings; built with numeric Date parts so
// there's no ISO-string timezone ambiguity)
// --------------------------------------------------------------------------

function tz_() {
  return Session.getScriptTimeZone() || 'Asia/Bangkok';
}

function toDateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return String(v || '').trim().slice(0, 10);
}

function parseYmd_(s) {
  const p = String(s).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function fmtYmd_(d) {
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

function addDays_(dateStr, n) {
  const d = parseYmd_(dateStr);
  return fmtYmd_(new Date(d.getFullYear(), d.getMonth(), d.getDate() + n));
}

// The official 1st/16th slot a real draw belongs to, even when GLO moved
// it a day or two (17 Jan -> 16 Jan, 2 May -> 1 May, 30 Dec -> 1 Jan).
function nominalSlot_(dateStr) {
  const d = parseYmd_(dateStr);
  const day = d.getDate();
  if (day <= 8) return fmtYmd_(new Date(d.getFullYear(), d.getMonth(), 1));
  if (day <= 23) return fmtYmd_(new Date(d.getFullYear(), d.getMonth(), 16));
  return fmtYmd_(new Date(d.getFullYear(), d.getMonth() + 1, 1));
}

// --------------------------------------------------------------------------
// LOAD DRAWS — one sheet read, parsed once, newest first.
// --------------------------------------------------------------------------

function fetchHistoricalSheetData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.HISTORICAL);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return { headerRow, data };
}

function loadDraws_() {
  const raw = fetchHistoricalSheetData_();
  if (!raw) return [];
  const { headerRow, data } = raw;
  const dateIdx = headerRow.indexOf('Draw Date');
  const colIdx = {};
  Object.keys(SET_DEFS).forEach(k => {
    colIdx[k] = SET_DEFS[k].historicalCols.map(n => headerRow.indexOf(n)).filter(i => i !== -1);
  });
  const draws = [];
  data.forEach(row => {
    const date = dateIdx === -1 ? '' : toDateStr_(row[dateIdx]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
    const values = {};
    let any = false;
    Object.keys(SET_DEFS).forEach(k => {
      const len = SET_DEFS[k].len;
      values[k] = colIdx[k]
        .map(i => String(row[i]).trim())
        .filter(v => /^\d+$/.test(v))
        .map(v => v.padStart(len, '0').slice(-len));
      if (values[k].length) any = true;
    });
    if (any) draws.push({ date, slot: nominalSlot_(date), values });
  });
  draws.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return draws;
}

function countHistoricalDraws_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.HISTORICAL);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return sheet.getLastRow() - 1;
}

// --------------------------------------------------------------------------
// SIGNAL FEED (optional, weight 0 by default) — read ONCE per request and
// only when a signal category actually has weight > 0.
// --------------------------------------------------------------------------

function loadSignalProbs_(weights) {
  const out = {};
  const active = Object.keys(SOURCE_TYPE_BY_CATEGORY).filter(c => (weights[c] || 0) > 0);
  if (active.length === 0) return out;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.SIGNAL_FEED);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const typeCol = headerRow.indexOf('Source Type');
  const numCol = headerRow.indexOf('Extracted Number(s)');
  if (typeCol === -1 || numCol === -1) return out;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  active.forEach(cat => {
    const counts = new Array(10).fill(0);
    let n = 0;
    const notesCol = headerRow.indexOf('Notes');
    data.forEach(row => {
      if (!SOURCE_TYPE_BY_CATEGORY[cat].includes(String(row[typeCol]).trim())) return;
      // Auto-collected buzz numbers have their own box — keep them out of the math.
      if (notesCol !== -1 && String(row[notesCol]).indexOf('[auto-buzz]') !== -1) return;
      (String(row[numCol]).match(/\d/g) || []).forEach(d => { counts[Number(d)]++; n++; });
    });
    if (n > 0) out[cat] = counts.map(c => (c + 1) / (n + 10)); // lightly smoothed
  });
  return out;
}

// --------------------------------------------------------------------------
// MODEL — one per set (jackpot/front3/back3/back2), optionally restricted
// to one nominal slot ('MM-DD', e.g. '10-16').
// --------------------------------------------------------------------------

function buildSetModel_(setKey, draws, settings, weights, signalProbs, slotMMDD) {
  const def = SET_DEFS[setKey];
  const len = def.len;
  const H = settings.halflife_draws;
  const A = settings.prior_strength;

  // rankStep converts a row's rank inside the subset into "draws of age":
  // all-time rows are 1 draw apart, same-slot rows are ~24 draws (1 year) apart.
  function counts(subset, rankStep) {
    const wc = Array.from({ length: len }, () => new Array(10).fill(0));
    const rc = Array.from({ length: len }, () => new Array(10).fill(0));
    let wTotal = 0, n = 0;
    subset.forEach((d, rank) => {
      const w = H > 0 ? Math.pow(0.5, (rank * rankStep) / H) : 1;
      d.values[setKey].forEach(num => {
        for (let pos = 0; pos < len; pos++) {
          const dg = Number(num[pos]);
          wc[pos][dg] += w;
          rc[pos][dg]++;
        }
        wTotal += w;
        n++;
      });
    });
    return { wc, rc, wTotal, n };
  }

  const all = counts(draws, 1);
  const allProbs = all.wc.map(c => c.map(x => (x + A * 0.1) / (all.wTotal + A)));

  let subset = draws, sub = all, histDigit = allProbs, rankStep = 1;
  if (slotMMDD) {
    subset = draws.filter(d => d.slot.slice(5) === slotMMDD);
    rankStep = 24;
    sub = counts(subset, rankStep);
    histDigit = sub.wc.map((c, pos) => c.map((x, dg) => (x + A * allProbs[pos][dg]) / (sub.wTotal + A)));
  }

  // Optional signal blend from HAND-TYPED Signal Feed rows. Signals are a
  // light nudge only: together they can never be more than SIGNAL_MAX_SHARE
  // of the math, whatever the Config weights say — a handful of typed digits
  // must not outweigh hundreds of real draws.
  const active = Object.keys(SOURCE_TYPE_BY_CATEGORY).filter(cat => (weights[cat] || 0) > 0 && signalProbs[cat]);
  const sigTotal = active.reduce((a, cat) => a + weights[cat], 0);
  const sigShare = sigTotal > 0 ? Math.min(SIGNAL_MAX_SHARE, sigTotal / ((weights.historical || 0) + sigTotal)) : 0;
  const digitProbs = histDigit.map(probs => {
    const mixed = probs.map(p => (1 - sigShare) * p);
    active.forEach(cat => {
      const share = sigShare * weights[cat] / sigTotal;
      for (let dg = 0; dg < 10; dg++) mixed[dg] += share * signalProbs[cat][dg];
    });
    const tot = mixed.reduce((a, b) => a + b, 0) || 1;
    return mixed.map(v => v / tot);
  });

  // Whole-number layer for 2/3-digit sets.
  let numberProbs = null, numberRaw = null;
  if (len <= 3) {
    const size = Math.pow(10, len);
    const wnum = new Array(size).fill(0);
    numberRaw = new Array(size).fill(0);
    let wSum = 0;
    subset.forEach((d, rank) => {
      const w = H > 0 ? Math.pow(0.5, (rank * rankStep) / H) : 1;
      d.values[setKey].forEach(num => { wnum[Number(num)] += w; numberRaw[Number(num)]++; wSum += w; });
    });
    const B = size * (A / 30); // ~1 pseudo-draw per number at default settings
    numberProbs = new Array(size);
    for (let n = 0; n < size; n++) {
      const s = String(n).padStart(len, '0');
      let prior = 1;
      for (let pos = 0; pos < len; pos++) prior *= digitProbs[pos][Number(s[pos])];
      numberProbs[n] = (wnum[n] + B * prior) / (wSum + B);
    }
  }

  // Chi-square uniformity test on RAW (unweighted) counts of this subset.
  let chiStat = 0;
  sub.rc.forEach(c => {
    const tot = c.reduce((a, b) => a + b, 0);
    if (tot === 0) return;
    const e = tot / 10;
    c.forEach(o => { chiStat += (o - e) * (o - e) / e; });
  });
  const df = 9 * len;
  const pValue = sub.n > 0 ? chiSquarePValue_(chiStat, df) : 1;

  return {
    setKey, len, digitProbs, numberProbs, numberRaw,
    nObs: sub.n, nDraws: subset.length,
    chi: { stat: +chiStat.toFixed(2), df, p: +pValue.toFixed(3) },
  };
}

function modelProb_(model, numStr) {
  if (model.numberProbs) return model.numberProbs[Number(numStr)];
  let p = 1;
  for (let pos = 0; pos < model.len; pos++) p *= model.digitProbs[pos][Number(numStr[pos])];
  return p;
}

function liftOf_(model, numStr) {
  return modelProb_(model, numStr) * Math.pow(10, model.len);
}

// --------------------------------------------------------------------------
// SAMPLING
// --------------------------------------------------------------------------

function sharpen_(probs, s) {
  if (s === 1) return probs.slice();
  if (s === 0) return probs.map(() => 1 / probs.length);
  const powed = probs.map(p => Math.pow(p, s));
  const tot = powed.reduce((a, b) => a + b, 0) || 1;
  return powed.map(p => p / tot);
}

function sampleIndex_(probs, exclude) {
  let tot = 0;
  for (let i = 0; i < probs.length; i++) if (!exclude || !exclude[i]) tot += probs[i];
  let r = Math.random() * tot;
  for (let i = 0; i < probs.length; i++) {
    if (exclude && exclude[i]) continue;
    r -= probs[i];
    if (r <= 0) return i;
  }
  for (let i = probs.length - 1; i >= 0; i--) if (!exclude || !exclude[i]) return i;
  return 0;
}

// Draws `count` DISTINCT numbers (front3/back3 A/B must differ).
function sampleNumbers_(model, count, sharpness) {
  const out = [];
  if (model.numberProbs) {
    const sp = sharpen_(model.numberProbs, sharpness);
    const used = {};
    while (out.length < count) {
      const idx = sampleIndex_(sp, used);
      const s = String(idx).padStart(model.len, '0');
      used[idx] = true;
      // Coverage: the second pick is never a flip (กลับเลข) of the first,
      // so a pair covers more different numbers.
      if (count > 1) flipsOf_(s).forEach(f => { used[Number(f)] = true; });
      out.push(s);
    }
  } else {
    const perPos = model.digitProbs.map(p => sharpen_(p, sharpness));
    let guard = 0;
    while (out.length < count && guard++ < 500) {
      const s = perPos.map(p => sampleIndex_(p)).join('');
      if (!out.includes(s)) out.push(s);
    }
  }
  return out.map(s => ({ number: s, prob: modelProb_(model, s), lift: +liftOf_(model, s).toFixed(2) }));
}

// Percentile rank (0-1) of a number inside its model: share of all
// possible numbers the math likes LESS (ties count half).
function percentileOf_(model, numStr) {
  if (model.numberProbs) {
    const p = model.numberProbs[Number(numStr)];
    let below = 0, ties = 0;
    model.numberProbs.forEach(q => { if (q < p) below++; else if (q === p) ties++; });
    return (below + (ties - 1) / 2) / (model.numberProbs.length - 1 || 1);
  }
  // Jackpot: average per-digit percentile (exact rank over 1M is overkill).
  let acc = 0;
  for (let pos = 0; pos < model.len; pos++) {
    const probs = model.digitProbs[pos];
    const p = probs[Number(numStr[pos])];
    let below = 0, ties = 0;
    probs.forEach(q => { if (q < p) below++; else if (q === p) ties++; });
    acc += (below + (ties - 1) / 2) / 9;
  }
  return acc / model.len;
}

function computeMathScore_(pairs) { // pairs: [{model, number}]
  if (!pairs.length) return 50;
  const avg = pairs.reduce((a, x) => a + percentileOf_(x.model, x.number), 0) / pairs.length;
  return Math.round(avg * 100);
}

// --------------------------------------------------------------------------
// TOP PICKS
// --------------------------------------------------------------------------

function computeTopPicks_(model, count) {
  let top;
  if (model.numberProbs) {
    top = model.numberProbs
      .map((p, n) => ({ digits: String(n).padStart(model.len, '0'), prob: p, seen: model.numberRaw[n] }))
      .sort((a, b) => b.prob - a.prob)
      .slice(0, count);
  } else {
    const K = Math.max(count, 5);
    const perPos = model.digitProbs.map(pv => pv
      .map((p, d) => ({ d, p }))
      .sort((a, b) => b.p - a.p)
      .slice(0, K));
    let combos = [{ digits: '', prob: 1 }];
    perPos.forEach(cands => {
      const next = [];
      combos.forEach(c => cands.forEach(cd => next.push({ digits: c.digits + cd.d, prob: c.prob * cd.p })));
      next.sort((a, b) => b.prob - a.prob);
      combos = next.slice(0, 200); // beam keeps it fast; top-5 stays exact
    });
    top = combos.slice(0, count).map(c => Object.assign(c, { seen: null }));
  }
  const sumProb = top.reduce((a, c) => a + c.prob, 0) || 1;
  const base = Math.pow(10, model.len);
  return top.map(c => ({
    number: c.digits,
    probabilityPct: +(c.prob / sumProb * 100).toFixed(1), // share among these top picks (bar width)
    lift: +(c.prob * base).toFixed(2),                     // × vs pure random
    seen: c.seen,                                          // exact past hits (2/3-digit only)
  }));
}

// --------------------------------------------------------------------------
// STATISTICS — chi-square p-value via the regularized upper incomplete
// gamma function Q(df/2, x/2) (Numerical Recipes gser/gcf).
// --------------------------------------------------------------------------

function gammaLn_(x) {
  const c = [76.18009172947146, -86.50532032941677, 24.01409824083091,
    -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  let y = x;
  const tmp = x + 5.5 - (x + 0.5) * Math.log(x + 5.5);
  let ser = 1.000000000190015;
  for (let j = 0; j < 6; j++) ser += c[j] / ++y;
  return -tmp + Math.log(2.5066282746310005 * ser / x);
}

function gammaQ_(a, x) {
  if (x <= 0) return 1;
  const gln = gammaLn_(a);
  if (x < a + 1) {
    let ap = a, sum = 1 / a, del = sum;
    for (let n = 0; n < 500; n++) {
      ap++; del *= x / ap; sum += del;
      if (Math.abs(del) < Math.abs(sum) * 1e-12) break;
    }
    return 1 - sum * Math.exp(-x + a * Math.log(x) - gln);
  }
  let b = x + 1 - a, c = 1e300, d = 1 / b, h = d;
  for (let i = 1; i < 500; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b; if (Math.abs(d) < 1e-300) d = 1e-300;
    c = b + an / c; if (Math.abs(c) < 1e-300) c = 1e-300;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-12) break;
  }
  return Math.exp(-x + a * Math.log(x) - gln) * h;
}

function chiSquarePValue_(stat, df) {
  return Math.min(1, Math.max(0, gammaQ_(df / 2, stat / 2)));
}

function evidenceVerdict_(p) {
  if (p < 0.01) return { th: 'เบี่ยงเบนจากการสุ่มอย่างชัดเจน', en: 'clearly deviates from pure random', level: 'strong' };
  if (p < 0.05) return { th: 'เบี่ยงเบนจากการสุ่มเล็กน้อย', en: 'slightly deviates from pure random', level: 'weak' };
  return { th: 'ไม่ต่างจากการสุ่มล้วน', en: 'indistinguishable from pure random', level: 'none' };
}

// --------------------------------------------------------------------------
// MATH BREAKDOWN endpoint (frontend "Math Probability Breakdown" boxes)
// --------------------------------------------------------------------------

function getMathBreakdown() {
  const weights = calculateWeights();
  const settings = readMathSettings_();
  const periodKey = currentPeriodKey_();
  const slotMMDD = periodKey.slice(5);
  const d = parseYmd_(periodKey);
  const monthDayLabelTH = d.getDate() + ' ' + THAI_MONTH_NAMES_[d.getMonth()];
  const monthDayLabelEN = EN_MONTH_NAMES_[d.getMonth()] + ' ' + d.getDate();

  const draws = loadDraws_();
  const signalProbs = loadSignalProbs_(weights);
  const drawCountFiltered = draws.filter(x => x.slot.slice(5) === slotMMDD).length;

  const sets = {};
  Object.keys(SET_DEFS).forEach(key => {
    const def = SET_DEFS[key];
    const mAll = buildSetModel_(key, draws, settings, weights, signalProbs);
    const mSlot = buildSetModel_(key, draws, settings, weights, signalProbs, slotMMDD);
    const verdict = evidenceVerdict_(mAll.chi.p);
    sets[key] = {
      label: def.label,
      labelEN: def.labelEN,
      len: def.len,
      topPicksAll: computeTopPicks_(mAll, 5),
      topPicksFiltered: computeTopPicks_(mSlot, 5),
      flipGroupsAll: computeFlipGroups_(mAll, 3),
      flipGroupsFiltered: computeFlipGroups_(mSlot, 3),
      evidence: { p: mAll.chi.p, chi: mAll.chi.stat, df: mAll.chi.df, th: verdict.th, en: verdict.en, level: verdict.level },
    };
  });

  const tuned = readMathTuning_();
  return {
    tuning: tuned ? {
      verdict: tuned.online.verdict, perDigitEdgePct: tuned.online.perDigitEdgePct, pValue: tuned.online.pValue,
      drawsTested: tuned.drawsTested, combosTested: tuned.combosTested, pureRandomRank: tuned.pureRandomRank,
      best: tuned.best, bestHistory: tuned.bestHistory, autoTune: !!settings.auto_tune, tunedAt: tuned.tunedAt,
    } : null,
    drawCountAll: draws.length,
    drawCountFiltered,
    monthDayLabelTH,
    monthDayLabelEN,
    settings,
    sets,
  };
}

// --------------------------------------------------------------------------
// REASONING — bilingual, one paragraph per set, built from the v2 model.
// --------------------------------------------------------------------------

function buildReasoning_(setKey, picks, model, settings, drawCount) {
  const def = SET_DEFS[setKey];
  const numbersText = picks.map(p => p.number).join(' / ');
  const disclaimerTH = ' หวยออกแบบสุ่มทุกงวด ใช้เพื่อความสนุกเท่านั้น';
  const disclaimerEN = ' Every draw is random — just for fun.';

  if (drawCount === 0) {
    return {
      th: `${def.label} (${numbersText}) ยังไม่มีข้อมูลย้อนหลัง จึงสุ่มเท่ากันทุกหลัก (10% ต่อหลัก)` + disclaimerTH,
      en: `${def.labelEN} (${numbersText}) has no historical draws yet, so every digit is a flat 10%.` + disclaimerEN,
    };
  }

  const top = computeTopPicks_(model, 1)[0];
  const picksTH = picks.map(p => p.number).join(', ');
  const picksEN = picksTH;
  const randomNow = model.chi.p >= 0.05;

  if (settings.sharpness === 0) {
    return {
      th: `${def.label}: ${picksTH} — สุ่มแบบทุกเลขมีโอกาสเท่ากัน เพราะผลทดสอบย้อนหลังกับผลจริง ${drawCount} งวด พบว่า "สุ่มล้วน" ดีที่สุด ` +
        `ไม่มีสูตรไหนเอาชนะดวงได้จริง (เลขมาแรงตอนนี้ ${top.number} ดูได้ด้านล่างเพื่อความสนุก).` + disclaimerTH,
      en: `${def.labelEN}: ${picksEN} — picked fully at random (every number equally likely), because honest testing on ${drawCount} real draws ` +
        `found pure random beats every history-based setting. The current hot number is ${top.number}, shown below for fun.` + disclaimerEN,
    };
  }
  const th = `${def.label}: ${picksTH} — สุ่มจากผลจริง ${drawCount} งวด โดยงวดหลังๆ นับมากกว่างวดเก่า ` +
    `เลขที่มาแรงที่สุดตอนนี้คือ ${top.number}. ` +
    (randomNow ? 'ผลจริงยังออกแบบสุ่มเหมือนเดิม' : 'ผลจริงเอียงไปบางเลขเล็กน้อย') + '.' + disclaimerTH;
  const en = `${def.labelEN}: ${picksEN} — picked from ${drawCount} real draws, with recent draws counting more. ` +
    `The hottest number right now is ${top.number}. ` +
    (randomNow ? 'Real results still look random.' : 'Real results lean slightly toward some numbers.') + disclaimerEN;
  return { th, en };
}

// --------------------------------------------------------------------------
// DRAW-DAY LOCK + PERIOD KEY
//
// Draw days (the 1st, the 16th, plus any moved date listed in the Draw
// Calendar tab) are LOCKED: nothing new is generated. If a set was already
// generated for that draw it's still shown; otherwise the wheel says
// "come back tomorrow". The next period opens the day after the draw.
//
//   day 1 (draw)  -> locked, period = this 1st
//   days 2-15     -> period = this month's 16th
//   day 16 (draw) -> locked, period = this 16th
//   days 17-end   -> period = next month's 1st
// --------------------------------------------------------------------------

function todayString_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}

let DRAW_OVERRIDES_CACHE_ = null;
function readDrawOverrides_() {
  if (DRAW_OVERRIDES_CACHE_) return DRAW_OVERRIDES_CACHE_;
  const out = [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.DRAW_CALENDAR);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(([v]) => {
      const s = toDateStr_(v);
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) out.push(s);
    });
  }
  DRAW_OVERRIDES_CACHE_ = out;
  return out;
}

// Returns { locked, slot, lockEnd } for a given day.
function drawDayInfo_(dateStr) {
  const day = Number(dateStr.slice(8, 10));
  let info = { locked: false, slot: null, lockEnd: null };
  if (day === 1 || day === 16) info = { locked: true, slot: dateStr, lockEnd: dateStr };
  readDrawOverrides_().forEach(o => {
    const s = nominalSlot_(o);
    const lo = o < s ? o : s;
    const hi = o > s ? o : s;
    if (dateStr >= lo && dateStr <= hi) {
      info = { locked: true, slot: s, lockEnd: (info.lockEnd && info.lockEnd > hi) ? info.lockEnd : hi };
    }
  });
  return info;
}

function nextDrawDate_(fromDateStr) {
  const d = parseYmd_(fromDateStr);
  if (d.getDate() === 1) return fromDateStr;
  if (d.getDate() <= 16) return fmtYmd_(new Date(d.getFullYear(), d.getMonth(), 16));
  return fmtYmd_(new Date(d.getFullYear(), d.getMonth() + 1, 1));
}

function currentPeriodKey_() {
  const today = todayString_();
  const info = drawDayInfo_(today);
  return info.locked ? info.slot : nextDrawDate_(today);
}

const THAI_MONTH_NAMES_ = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม',
];
const EN_MONTH_NAMES_ = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
function formatDrawDateLabels_(dateStr) {
  const d = parseYmd_(dateStr);
  return {
    th: 'งวดวันที่ ' + d.getDate() + ' ' + THAI_MONTH_NAMES_[d.getMonth()] + ' ' + (d.getFullYear() + 543),
    en: 'Draw of ' + EN_MONTH_NAMES_[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear(),
  };
}

function shortDayLabels_(dateStr) {
  const d = parseYmd_(dateStr);
  return {
    th: d.getDate() + ' ' + THAI_MONTH_NAMES_[d.getMonth()],
    en: EN_MONTH_NAMES_[d.getMonth()] + ' ' + d.getDate(),
  };
}

// Everything the frontend needs to show / explain the lock.
function lockStatus_() {
  const today = todayString_();
  const info = drawDayInfo_(today);
  if (!info.locked) return { drawDayLocked: false };
  const opens = addDays_(info.lockEnd, 1);
  const o = shortDayLabels_(opens);
  return {
    drawDayLocked: true,
    nextOpenDate: opens,
    lockMessageTH: 'วันนี้เป็นวันออกรางวัล — งดสร้างเลขใหม่ เปิดให้หมุนงวดถัดไปวันที่ ' + o.th,
    lockMessageEN: 'Today is draw day — no new numbers today. The wheel reopens for the next draw on ' + o.en + '.',
  };
}

function ensureTrailingColumn_(sheet, label) {
  const lastCol = sheet.getLastColumn();
  const headerRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (headerRow.indexOf(label) === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(label).setFontWeight('bold');
  }
}

function readTodaysCache_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.DAILY_CACHE);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const periodKey = currentPeriodKey_();
  const width = Math.max(12, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(width, sheet.getMaxColumns())).getValues();
  for (let i = data.length - 1; i >= 0; i--) { // newest first
    const row = data[i];
    if (toDateStr_(row[0]) !== periodKey) continue;
    const rawGeneratedAt = row[10];
    let reasoningBySet = null;
    try { reasoningBySet = row[11] ? JSON.parse(row[11]) : null; } catch (e) { reasoningBySet = null; }
    const periodLabels = formatDrawDateLabels_(periodKey);
    const pad = (v, n) => String(v).padStart(n, '0');
    return Object.assign({
      jackpot: pad(row[1], 6), front3A: pad(row[2], 3), front3B: pad(row[3], 3),
      back3A: pad(row[4], 3), back3B: pad(row[5], 3), back2: pad(row[6], 2),
      confidence: row[7], reasoning: row[8], signalSnapshot: row[9],
      generatedAt: rawGeneratedAt instanceof Date ? rawGeneratedAt.toISOString() : String(rawGeneratedAt),
      reasoningBySet, fromCache: true,
      periodTargetDate: periodKey, periodLabelTH: periodLabels.th, periodLabelEN: periodLabels.en,
      historicalDrawCount: countHistoricalDraws_(),
    }, lockStatus_());
  }
  return null;
}

function writeTodaysCache_(result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureTab_(ss, TABS.DAILY_CACHE,
    ['Date', 'Jackpot', 'Front3-A', 'Front3-B', 'Back3-A', 'Back3-B', 'Back2',
     'Confidence %', 'Reasoning', 'Signal Snapshot', 'Generated At', 'Reasoning Detail']);
  ensureTrailingColumn_(sheet, 'Reasoning Detail');
  appendRowForcedText_(sheet, [
    result.periodTargetDate, result.jackpot, result.front3A, result.front3B,
    result.back3A, result.back3B, result.back2,
    result.confidence, result.reasoning, result.signalSnapshot,
    new Date(), JSON.stringify(result.reasoningBySet || null),
  ], [1, 2, 3, 4, 5, 6, 7]);
}

// --------------------------------------------------------------------------
// MAIN ENTRY POINT — called from the frontend via google.script.run.
// --------------------------------------------------------------------------

// Wrapper adds how long the server took (shown in the debug panel).
function generateSet() {
  const t0 = Date.now();
  const res = generateSetInner_();
  if (res && typeof res === 'object') res.serverMs = Date.now() - t0;
  return res;
}

function generateSetInner_() {
  const cached = readTodaysCache_();
  if (cached) {
    Logger.log('CACHE HIT for period ' + cached.periodTargetDate);
    return cached;
  }

  // Draw day and nothing generated for this draw -> do NOT generate.
  const lock = lockStatus_();
  if (lock.drawDayLocked) {
    const periodKey = currentPeriodKey_();
    const labels = formatDrawDateLabels_(periodKey);
    return Object.assign({
      locked: true, periodTargetDate: periodKey,
      periodLabelTH: labels.th, periodLabelEN: labels.en,
    }, lock);
  }

  // Two people pressing at the same moment must not create two sets for
  // one period — serialize, then re-check the cache inside the lock.
  const scriptLock = LockService.getScriptLock();
  scriptLock.waitLock(20000);
  try {
    const again = readTodaysCache_();
    if (again) return again;
    return generateFreshSet_();
  } finally {
    scriptLock.releaseLock();
  }
}

function generateFreshSet_() {
  const weights = calculateWeights();
  let settings = readMathSettings_();
  // Best-math engine: use the settings that scored best on honest
  // walk-forward testing (see tuneMath). Pure random can win too.
  const tuned = settings.auto_tune ? readMathTuning_() : null;
  let slotMMDD;
  if (tuned && tuned.best) {
    settings = Object.assign({}, settings, {
      halflife_draws: tuned.best.halflife_draws,
      prior_strength: tuned.best.prior_strength,
      sharpness: tuned.best.sharpness,
    });
    if (tuned.best.column === 'slot') slotMMDD = currentPeriodKey_().slice(5);
  }
  const draws = loadDraws_();
  const signalProbs = loadSignalProbs_(weights);
  const sharp = settings.sharpness;

  const models = {};
  Object.keys(SET_DEFS).forEach(k => { models[k] = buildSetModel_(k, draws, settings, weights, signalProbs, slotMMDD); });

  const jackpot = sampleNumbers_(models.jackpot, 1, sharp);
  const front3 = sampleNumbers_(models.front3, 2, sharp);
  const back3 = sampleNumbers_(models.back3, 2, sharp);
  const back2 = sampleNumbers_(models.back2, 1, sharp);

  const confidence = computeMathScore_([
    ...jackpot.map(p => ({ model: models.jackpot, number: p.number })),
    ...front3.map(p => ({ model: models.front3, number: p.number })),
    ...back3.map(p => ({ model: models.back3, number: p.number })),
    ...back2.map(p => ({ model: models.back2, number: p.number })),
  ]);

  const parts = [
    { key: 'jackpot', picks: jackpot },
    { key: 'front3', picks: front3 },
    { key: 'back3', picks: back3 },
    { key: 'back2', picks: back2 },
  ].map(x => Object.assign(x, buildReasoning_(x.key, x.picks, models[x.key], settings, draws.length)));

  const reasoningBySet = parts.map(x => ({
    key: x.key, labelTH: SET_DEFS[x.key].label, labelEN: SET_DEFS[x.key].labelEN, th: x.th, en: x.en,
  }));
  const reasoning = parts.map(x => x.th).join(' ') + '\n\n' + parts.map(x => x.en).join(' ');

  const signalSnapshot = JSON.stringify({ engine: tuned ? 'math-v3-tuned' : 'math-v2', weights, settings,
    column: slotMMDD ? 'slot' : 'all', tuning: tuned ? tuned.online : null, generatedAt: new Date().toISOString() });

  const periodKey = currentPeriodKey_();
  const periodLabels = formatDrawDateLabels_(periodKey);

  const result = {
    jackpot: jackpot[0].number,
    front3A: front3[0].number, front3B: front3[1].number,
    back3A: back3[0].number, back3B: back3[1].number,
    back2: back2[0].number,
    confidence,
    reasoning,
    reasoningBySet,
    signalSnapshot,
    generatedAt: new Date().toISOString(),
    periodTargetDate: periodKey,
    periodLabelTH: periodLabels.th,
    periodLabelEN: periodLabels.en,
    historicalDrawCount: draws.length,
    fromCache: false,
    drawDayLocked: false,
  };

  writeTodaysCache_(result);
  logResult(result);
  Logger.log('Generated fresh set for period ' + periodKey + ': ' + JSON.stringify(result));
  return result;
}

// --------------------------------------------------------------------------
// LOGGING — one Generated Log row per fresh generation. Written BY HEADER
// NAME so old sheets with a different column order can't get misaligned.
// 'Period Key' lets Fake vs Real match a guess to its draw exactly.
// --------------------------------------------------------------------------

function logResult(result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureTab_(ss, TABS.GENERATED_LOG,
    ['Timestamp', 'Jackpot', 'Front3-A', 'Front3-B', 'Back3-A', 'Back3-B', 'Back2',
     'Confidence %', 'Reasoning Text', 'Signal Snapshot Ref', 'Reasoning Detail', 'Period Key']);
  ensureTrailingColumn_(sheet, 'Reasoning Detail');
  ensureTrailingColumn_(sheet, 'Period Key');
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const values = {
    'Timestamp': new Date(), 'Jackpot': result.jackpot,
    'Front3-A': result.front3A, 'Front3-B': result.front3B,
    'Back3-A': result.back3A, 'Back3-B': result.back3B, 'Back2': result.back2,
    'Confidence %': result.confidence, 'Reasoning Text': result.reasoning,
    'Signal Snapshot Ref': result.signalSnapshot,
    'Reasoning Detail': JSON.stringify(result.reasoningBySet || null),
    'Period Key': result.periodTargetDate,
  };
  const row = header.map(h => (h in values ? values[h] : ''));
  const textCols = ['Jackpot', 'Front3-A', 'Front3-B', 'Back3-A', 'Back3-B', 'Back2', 'Period Key']
    .map(h => header.indexOf(h) + 1).filter(c => c > 0);
  appendRowForcedText_(sheet, row, textCols);
}

// --------------------------------------------------------------------------
// DEPRECATED — do not use. This was the original Phase 7 stub, written
// before Historical Draws had a real-vs-fake distinction. It copied Daily
// Cache (FAKE generated guesses) into Historical Draws, which would corrupt
// the real-data tab with numbers that were never an actual GLO draw. It was
// never wired to a trigger, so it never ran automatically — left here only
// so any old reference to the function name doesn't hard-error, and it now
// intentionally does nothing. Real draw archiving is fetchAndArchiveRealDraw()
// below, which only ever writes officially-confirmed GLO results.
// --------------------------------------------------------------------------

function archiveDailyCacheIfDrawDay() {
  Logger.log('archiveDailyCacheIfDrawDay() is deprecated and does nothing — ' +
    'it would have copied FAKE generated numbers into Historical Draws. ' +
    'Use fetchAndArchiveRealDraw() / createRealDrawFetchTrigger() instead.');
}

// ==========================================================================
// REAL DRAW AUTO-FETCH — pulls the OFFICIAL GLO result and appends it to
// Historical Draws. Publish timing after a real draw (1st/16th) isn't fixed
// — sometimes it's up same-day, sometimes not until the day after or later
// — so rather than gating on one specific day (e.g. "only try on the 3rd"),
// the daily trigger just tries EVERY day: it always targets the most recent
// 1st-or-16th on/before today, checks whether Historical Draws already has
// that date, and if not, tries to fetch it. A day where it's not published
// yet just logs "could not confirm" and tries again tomorrow — once a date
// succeeds, historicalDrawsHasDate_() makes every later day's attempt an
// instant no-op (no wasted API calls). NEVER overwrites or deletes existing
// rows — this respects the "keep all real data intact, forever" decision:
// it only appends a new row, and only after checking that row doesn't
// already exist.
//
// Two sources, tried in order:
//   1. A free public Thai lottery JSON API (no key, structured data — no
//      hallucination risk). UNVERIFIED LIVE in this session (web tools were
//      unavailable here) — run testPrimaryLotteryApi() once and compare its
//      logged output against parsePrimaryApiResponse_() below; adjust the
//      field names there if the live shape differs.
//   2. Gemini API with Google Search grounding, as a fallback only. Requires
//      a free key from https://aistudio.google.com/apikey stored in
//      Project Settings > Script Properties as GEMINI_API_KEY.
// Either source returning anything that doesn't pass isValidRealDraw_() is
// discarded — Historical Draws only ever receives fully-validated results.
// ==========================================================================

// Gemini models used for every AI call (real-draw backup, buzz numbers,
// dream matching), tried in order — if the first is rate-limited the next
// one is used. Free tier: 3.6 Flash = 20/day, 3.5 Flash Lite = 500/day.
// (gemini-3.5-flash is retired — Google silently answers it with 3.6 Flash.)
// To override without editing code, add a Script Property GEMINI_MODEL,
// e.g. "gemini-3.6-flash" or a comma list "gemini-3.6-flash,gemini-3.5-flash-lite".
const GEMINI_DEFAULT_MODELS = ['gemini-3.6-flash', 'gemini-3.5-flash-lite'];

const REAL_DRAW_SOURCE = {
  PRIMARY_API: 'GitHub Archive (vicha-w/thai-lotto-archive)',
  GEMINI_FALLBACK: 'Gemini (Search-grounded)',
};

// The GitHub archive you already used by hand to fetch the real 227-row
// dataset — one plain-text file per draw date, e.g.
// lottonumbers/2026-08-01.txt, with lines like:
//   FIRST 123456
//   THREE_FIRST 123 456
//   THREE_LAST 789 012
//   TWO 34
// This is a real, validated source (it produced the data you already
// checked and imported), so it's the primary — no key, no hallucination
// risk, just a direct file fetch.
const PRIMARY_LOTTERY_ARCHIVE = {
  OWNER: 'vicha-w',
  REPO: 'thai-lotto-archive',
  PATH: 'lottonumbers',
};

function scriptProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v || fallback;
}

// Run this once by hand after setup to confirm your Script Properties are
// in place before relying on the automatic trigger.
function verifyRealDrawFetchSetup() {
  const key = scriptProp_('GEMINI_API_KEY', null);
  Logger.log(key
    ? 'GEMINI_API_KEY is set.'
    : 'GEMINI_API_KEY is MISSING — add it in Project Settings (gear icon) > ' +
      'Script Properties before relying on the Gemini fallback (only used if the ' +
      'GitHub archive is down or missing that date). Get a free key at ' +
      'https://aistudio.google.com/apikey.');
  Logger.log('Primary source: github.com/' + PRIMARY_LOTTERY_ARCHIVE.OWNER + '/' + PRIMARY_LOTTERY_ARCHIVE.REPO);
}

// Diagnostic — run by hand against a date you KNOW is in the archive (e.g.
// one from the 227-row dataset you already fetched, like '2026-08-01') and
// check the Execution log. If GitHub ever restructures this repo, this is
// what to re-run before touching parsePrimaryApiResponse_() below.
function testPrimaryLotteryApi() {
  const sample = fetchRealDrawFromPrimaryApi_('2026-08-01');
  Logger.log(JSON.stringify(sample));
}

// Fetches ONE draw's raw text file directly from raw.githubusercontent.com
// — a CDN-served static-file host, NOT the api.github.com REST API. This
// matters: api.github.com enforces a 60-requests/hour UNAUTHENTICATED rate
// limit per source IP, and Apps Script's outbound IP is shared across an
// unknown number of other scripts/users, so that quota can get exhausted by
// traffic that has nothing to do with this project (confirmed live: got a
// 403 "API rate limit exceeded" on the very first test run). Raw file
// delivery isn't subject to that same per-IP API quota, so this sidesteps
// the problem entirely — no auth token needed.
//
// Tries 'main' then 'master' (repo's default branch wasn't confirmed live,
// since web tools were unavailable this session) — whichever responds 200
// wins; a 404 on both means that date's file doesn't exist yet (not
// published, or the maintainer hasn't added it), treated as "not found" and
// falls through to the Gemini fallback, not as an error.
function fetchRealDrawFromPrimaryApi_(targetDateStr) {
  const branches = ['main', 'master'];
  for (let i = 0; i < branches.length; i++) {
    const url = 'https://raw.githubusercontent.com/' + PRIMARY_LOTTERY_ARCHIVE.OWNER + '/' +
      PRIMARY_LOTTERY_ARCHIVE.REPO + '/' + branches[i] + '/' + PRIMARY_LOTTERY_ARCHIVE.PATH + '/' +
      targetDateStr + '.txt';
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = resp.getResponseCode();
      if (code === 200) return parsePrimaryApiResponse_(resp.getContentText());
      if (code !== 404) {
        Logger.log('raw.githubusercontent.com (' + branches[i] + ') HTTP ' + code + ': ' + resp.getContentText().slice(0, 300));
      }
      // 404 -> try the next branch name, or fall through to null after the loop.
    } catch (e) {
      Logger.log('Primary lottery API (raw fetch, branch=' + branches[i] + ') failed: ' + e);
    }
  }
  return null;
}

// Parses the archive's plain-text format — same parsing logic as the Python
// script that already validated this source (FIRST / THREE_FIRST /
// THREE_LAST / TWO, space-separated).
function parsePrimaryApiResponse_(text) {
  const result = { jackpot: null, front3A: null, front3B: null, back3A: null, back3B: null, back2: null };
  text.split('\n').forEach(line => {
    const parts = line.trim().split(/\s+/);
    if (!parts.length || !parts[0]) return;
    const key = parts[0];
    const values = parts.slice(1);
    if (key === 'FIRST' && values[0]) result.jackpot = values[0];
    else if (key === 'THREE_FIRST') {
      if (values[0]) result.front3A = values[0];
      if (values[1]) result.front3B = values[1];
    } else if (key === 'THREE_LAST') {
      if (values[0]) result.back3A = values[0];
      if (values[1]) result.back3B = values[1];
    } else if (key === 'TWO' && values[0]) result.back2 = values[0];
  });
  const complete = result.jackpot && result.front3A && result.front3B &&
    result.back3A && result.back3B && result.back2;
  return complete ? result : null;
}

// Fallback when the public API is down or doesn't have the draw yet — asks
// Gemini, with Google Search grounding turned on, to confirm the officially
// published GLO result rather than estimate one. Returns null (not a guess)
// if Gemini itself reports it couldn't confirm an official result.
function fetchRealDrawFromGemini_(targetDateStr, candidateDates) {
  const dateList = (candidateDates && candidateDates.length) ? candidateDates : [targetDateStr];
  const res = callGemini_(
    'Search for the OFFICIAL Royal Thai Government Lottery (GLO) result for the draw held on one of ' +
    'these dates (yyyy-MM-dd): ' + dateList.join(', ') + ' (GLO sometimes moves a draw by a day or two). ' +
    'Reply ONLY with JSON: {"found":true,"drawDate":"yyyy-MM-dd","jackpot":"6 digits","front3A":"3 digits",' +
    '"front3B":"3 digits","back3A":"3 digits","back3B":"3 digits","back2":"2 digits"} — all numbers as zero-padded ' +
    'strings. If you cannot find a confirmed official result for these dates, reply {"found":false}. ' +
    'Never guess, estimate, or use a different draw\'s result.', true);
  const parsed = res && extractJson_(res.text);
  if (!parsed || !parsed.found) return null;
  const str = v => (v == null ? null : String(v));
  return {
    drawDate: parsed.drawDate || null,
    jackpot: str(parsed.jackpot), front3A: str(parsed.front3A), front3B: str(parsed.front3B),
    back3A: str(parsed.back3A), back3B: str(parsed.back3B), back2: str(parsed.back2),
  };
}

// Validates shape/digit-length before ever writing to Historical Draws — a
// malformed or partial result (from either source) is discarded rather than
// risking a corrupt row in the permanent real-data record.
function isValidRealDraw_(r) {
  if (!r) return false;
  const checks = [[r.jackpot, 6], [r.front3A, 3], [r.front3B, 3], [r.back3A, 3], [r.back3B, 3], [r.back2, 2]];
  return checks.every(([val, len]) => typeof val === 'string' && new RegExp('^\\d{' + len + '}$').test(val));
}

// The most recent NOMINAL draw slot (1st or 16th) on or before today.
function mostRecentDrawDate_() {
  const today = todayString_();
  const d = parseYmd_(today);
  let slot = fmtYmd_(new Date(d.getFullYear(), d.getMonth(), d.getDate() >= 16 ? 16 : 1));
  // A moved draw held BEFORE its slot (e.g. 30 Dec for the 1 Jan slot)
  // has already happened — target that slot as soon as the draw is done.
  readDrawOverrides_().forEach(o => {
    const s = nominalSlot_(o);
    if (o <= today && s > slot) slot = s;
  });
  return slot;
}

// Real dates the draw for `slot` may actually have been held on: any moved
// date listed in Draw Calendar first, then the slot itself, then the usual
// GLO shifts (+1 day e.g. 17 Jan / 2 May, -1/-2 days e.g. 31 Jul / 30 Dec).
// Never returns a date in the future.
function candidateDrawDates_(slot) {
  const today = todayString_();
  const list = readDrawOverrides_().filter(o => nominalSlot_(o) === slot);
  [slot, addDays_(slot, 1), addDays_(slot, -1), addDays_(slot, -2)].forEach(c => {
    if (!list.includes(c)) list.push(c);
  });
  return list.filter(c => c <= today);
}

function historicalDrawsHasDate_(sheet, dateStr) {
  if (sheet.getLastRow() < 2) return false;
  const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return dates.some(([d]) => toDateStr_(d) === dateStr);
}

// MAIN ENTRY for real-draw fetching. With no argument it targets the most
// recent 1st/16th slot and checks every real date that draw could have been
// held on (so moved draws like 17 Jan / 2 May / 30 Dec are found and are
// never fetched twice). Pass 'yyyy-MM-dd' to force one exact date.
function fetchAndArchiveRealDraw(forceDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureTab_(ss, TABS.HISTORICAL,
    ['Draw Date', 'Prize1 (6-digit)', 'Front3-A', 'Front3-B', 'Back3-A', 'Back3-B', 'Back2'],
    [1, 2, 3, 4, 5, 6, 7]);
  ensureHistoricalDrawsExtraColumns_(sheet);

  const slot = forceDateStr ? nominalSlot_(forceDateStr) : mostRecentDrawDate_();
  const candidates = forceDateStr ? [forceDateStr] : candidateDrawDates_(slot);

  const existing = sheet.getLastRow() < 2 ? [] :
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(([d]) => toDateStr_(d));
  const already = candidates.find(c => existing.includes(c));
  if (already) {
    Logger.log('Historical Draws already has the ' + slot + ' draw (' + already + ') — skipping.');
    return;
  }

  let result = null, source = null, actualDate = null;
  for (let i = 0; i < candidates.length && !isValidRealDraw_(result); i++) {
    result = fetchRealDrawFromPrimaryApi_(candidates[i]);
    if (isValidRealDraw_(result)) { source = REAL_DRAW_SOURCE.PRIMARY_API; actualDate = candidates[i]; }
  }

  if (!isValidRealDraw_(result)) {
    result = fetchRealDrawFromGemini_(candidates[0] || slot, candidates);
    if (isValidRealDraw_(result)) {
      source = REAL_DRAW_SOURCE.GEMINI_FALLBACK;
      actualDate = (result.drawDate && candidates.includes(result.drawDate)) ? result.drawDate : (candidates[0] || slot);
      if (existing.includes(actualDate)) return;
    }
  }

  if (!isValidRealDraw_(result)) {
    Logger.log('Could not confirm the official result for the ' + slot + ' draw (tried ' + candidates.join(', ') +
      ') — nothing written. Will retry on the next daily run.');
    return;
  }

  appendRowForcedText_(sheet, [
    actualDate, result.jackpot, result.front3A, result.front3B,
    result.back3A, result.back3B, result.back2, source, new Date(),
  ], [1, 2, 3, 4, 5, 6, 7]);
  Logger.log('Archived real draw ' + actualDate + ' (slot ' + slot + ') via ' + source + ': ' + JSON.stringify(result));

  recordFakeVsRealComparison_(actualDate, result);
  // Add this draw to the math-column backtest (fast: one new draw).
  try { updateMathBacktest_(60 * 1000); } catch (e) { Logger.log('Math backtest update failed: ' + e); }
  // Re-tune the best-math engine with the new draw included.
  try { tuneMath(); } catch (e) { Logger.log('Math tuning failed: ' + e); }
}

// One-time setup — installs a daily trigger that tries every day (publish
// timing after a draw isn't fixed, so this keeps retrying until it lands
// rather than gating on one specific day). Once a date's archived,
// historicalDrawsHasDate_() makes every later day's run an instant no-op.
// Running this will prompt an authorization dialog the first time (approve
// it), since it's the first thing in this project to use external HTTP
// calls plus a time-driven trigger.
function createRealDrawFetchTrigger() {
  const already = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'dailyRealDrawFetchCheck');
  if (already) {
    Logger.log('Trigger already exists — not creating a duplicate.');
    return;
  }
  ScriptApp.newTrigger('dailyRealDrawFetchCheck')
    .timeBased()
    .everyDays(1)
    .atHour(9) // ~9am Asia/Bangkok
    .create();
  Logger.log('Daily real-draw-fetch trigger installed (runs ~9am every day; ' +
    'no-ops instantly once the current draw is already archived).');
}

// The actual function the trigger calls every day — a thin wrapper so the
// registered trigger handler name never has to change even if
// fetchAndArchiveRealDraw()'s signature does.
function dailyRealDrawFetchCheck() {
  fetchAndArchiveRealDraw();
  try { autoBuzzRefresh_(); } catch (e) { Logger.log('Buzz refresh failed: ' + e); }
}

// ==========================================================================
// FAKE VS REAL COMPARISON — once a real draw is archived, this compares it
// against the most recent FAKE guess (from Generated Log) made before that
// draw happened, and appends one row to the "Fake vs Real" tab. This tab is
// the single source of truth both for a human browsing it directly and for
// getComparisonStats() below, which the frontend dashboard reads from.
// ==========================================================================

function digitMatches_(a, b) {
  if (!a || !b || String(a).length !== String(b).length) return 0;
  const sa = String(a), sb = String(b);
  let n = 0;
  for (let i = 0; i < sa.length; i++) if (sa[i] === sb[i]) n++;
  return n;
}

// A fake set has an A/B pair but there's no meaningful "correct order"
// between them, so score each fake number against whichever real number
// (A or B) it matches best.
function bestMatchOfTwo_(fakeNum, realA, realB) {
  return Math.max(digitMatches_(fakeNum, realA), digitMatches_(fakeNum, realB));
}

// Draws happen on the 1st and 16th. Given one of those dates, returns the
// PRIOR draw date — used to bound the "how many fake tries happened between
// the last real draw and this one" window (a period can be up to ~15 days,
// and the Daily Lock allows at most one fresh fake generation per day, so
// there can be several tries to compare against a single real draw).
function previousDrawDate_(slotDateStr) {
  const d = parseYmd_(nominalSlot_(slotDateStr));
  if (d.getDate() === 16) return fmtYmd_(new Date(d.getFullYear(), d.getMonth(), 1));
  return fmtYmd_(new Date(d.getFullYear(), d.getMonth() - 1, 16));
}

// Reads Daily Cache + Generated Log ONCE and groups every Lucky Draw set by
// the draw it was made for: { '2026-10-16': [set, set, ...], ... }.
//  - Daily Cache: its Date column IS that draw (most reliable).
//  - Generated Log: 'Period Key' when present; older rows fall back to the
//    calendar rule (spun on day X counts toward the next 1st/16th).
// One read of each tab — fast even with hundreds of past draws.
function buildGuessIndex_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pad = (v, n) => String(v).trim().padStart(n, '0');
  const index = {};
  const seen = {};
  const add = (slot, g) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(slot)) return;
    const k = slot + '|' + [g.jackpot, g.front3A, g.front3B, g.back3A, g.back3B, g.back2].join('|');
    if (seen[k]) return;
    seen[k] = true;
    (index[slot] = index[slot] || []).push(g);
  };
  const setFrom = (r, at) => ({
    timestamp: at && !isNaN(at.getTime()) ? at : new Date(),
    jackpot: pad(r[1], 6), front3A: pad(r[2], 3), front3B: pad(r[3], 3),
    back3A: pad(r[4], 3), back3B: pad(r[5], 3), back2: pad(r[6], 2),
  });

  const cache = ss.getSheetByName(TABS.DAILY_CACHE);
  if (cache && cache.getLastRow() > 1) {
    cache.getRange(2, 1, cache.getLastRow() - 1, 11).getValues().forEach(r => {
      const at = r[10] instanceof Date ? r[10] : new Date(r[10]);
      add(toDateStr_(r[0]), setFrom(r, at));
    });
  }

  const log = ss.getSheetByName(TABS.GENERATED_LOG);
  if (log && log.getLastRow() > 1) {
    const header = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0].map(String);
    const keyIdx = header.indexOf('Period Key');
    log.getRange(2, 1, log.getLastRow() - 1, log.getLastColumn()).getValues().forEach(r => {
      const ts = r[0] instanceof Date ? r[0] : new Date(r[0]);
      let slot = keyIdx !== -1 ? toDateStr_(r[keyIdx]) : '';
      if (!slot) {
        if (isNaN(ts.getTime())) return;
        slot = nextDrawDate_(fmtYmd_(ts));
      }
      add(slot, setFrom(r, ts));
    });
  }
  return index;
}

// Every Lucky Draw set made for the draw that `actualDrawDate` belongs to.
// Pass a prebuilt index (from buildGuessIndex_) when checking many draws.
function findAllFakeGuessesInPeriod_(actualDrawDate, index) {
  const idx = index || buildGuessIndex_();
  return idx[nominalSlot_(actualDrawDate)] || [];
}


// Scores every fake try that happened this period against the real result
// and keeps the BEST-matching one as the representative Fake vs Real row
// (jackpot weighted heaviest since it's the headline number). The full try
// count is recorded too, so the dashboard can show "best of N tries" rather
// than implying there was only ever one guess per draw.
function recordFakeVsRealComparison_(targetDate, real, index) {
  const guesses = findAllFakeGuessesInPeriod_(targetDate, index);
  if (guesses.length === 0) {
    Logger.log('No fake guesses exist for the period before ' + targetDate + ' — skipping Fake vs Real comparison for this draw.');
    return;
  }

  let best = null;
  let bestScore = -1;
  guesses.forEach(fake => {
    const jackpotMatches = digitMatches_(fake.jackpot, real.jackpot);
    const front3Best = Math.max(
      bestMatchOfTwo_(fake.front3A, real.front3A, real.front3B),
      bestMatchOfTwo_(fake.front3B, real.front3A, real.front3B)
    );
    const back3Best = Math.max(
      bestMatchOfTwo_(fake.back3A, real.back3A, real.back3B),
      bestMatchOfTwo_(fake.back3B, real.back3A, real.back3B)
    );
    const back2Matches = digitMatches_(fake.back2, real.back2);
    const score = jackpotMatches * 3 + front3Best + back3Best + back2Matches;
    if (score > bestScore) {
      bestScore = score;
      best = { fake, jackpotMatches, front3Best, back3Best, back2Matches };
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.FAKE_VS_REAL) || ensureTab_(ss, TABS.FAKE_VS_REAL, [
    'Draw Date', 'Real Jackpot', 'Fake Jackpot', 'Jackpot Matches (/6)',
    'Real Front3-A', 'Real Front3-B', 'Fake Front3-A', 'Fake Front3-B', 'Front3 Best Match (/3)',
    'Real Back3-A', 'Real Back3-B', 'Fake Back3-A', 'Fake Back3-B', 'Back3 Best Match (/3)',
    'Real Back2', 'Fake Back2', 'Back2 Matches (/2)',
    'Tries This Period', 'Fake Generated At', 'Compared At',
  ], [1, 2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 15, 16]);

  const fake = best.fake;
  appendRowForcedText_(sheet, [
    targetDate, real.jackpot, fake.jackpot, best.jackpotMatches,
    real.front3A, real.front3B, fake.front3A, fake.front3B, best.front3Best,
    real.back3A, real.back3B, fake.back3A, fake.back3B, best.back3Best,
    real.back2, fake.back2, best.back2Matches,
    guesses.length, fake.timestamp.toISOString(), new Date(),
  ], [1, 2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 15, 16]);
  Logger.log('Recorded Fake vs Real comparison for ' + targetDate + ' — best of ' + guesses.length +
    ' tries this period. jackpot ' + best.jackpotMatches + '/6, front3 ' + best.front3Best + '/3, ' +
    'back3 ' + best.back3Best + '/3, back2 ' + best.back2Matches + '/2.');
}

// Self-heal for real draws added BY HAND directly into Historical Draws
// (rather than via fetchAndArchiveRealDraw(), which normally triggers the
// comparison as a side-effect of archiving). Scans every Historical Draws
// row, and for any date that doesn't already have a Fake vs Real row, tries
// to record one now. If no Lucky Draw guess exists in that draw's period
// (Generated Log), recordFakeVsRealComparison_() just no-ops for that date —
// there's nothing to compare against, which is expected, not a bug.
//
// PERFORMANCE: this used to run its full scan (a complete read of Historical
// Draws, PLUS a complete read of Fake vs Real just to know which dates are
// already covered) on literally every single call — including every normal
// dashboard page load, even when nothing had changed since the last check.
// That was the main reason the "Lucky Draw vs. Real Draws" dashboard felt
// slow. Fix: Historical Draws only ever grows by appended rows (real draws
// are never edited in place — see the "never fabricate history" rule
// elsewhere in this file), so its row COUNT is a cheap, reliable signal for
// "could anything possibly need backfilling since last time?" We cache the
// row count we last scanned at (in Document Properties, so it survives
// across script executions) and skip the two full-table reads entirely
// when the count hasn't moved — checking getLastRow() alone doesn't cost a
// getValues() call, so the steady-state (nothing new) case is now nearly
// free instead of two full sheet reads.
const BACKFILL_LOGIC_VERSION = 'v3-fast-index';

function backfillFakeVsRealComparisons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const histSheet = ss.getSheetByName(TABS.HISTORICAL);
  if (!histSheet || histSheet.getLastRow() < 2) return 0;

  const props = PropertiesService.getDocumentProperties();
  const currentHistRowCount = histSheet.getLastRow();
  // The stored marker includes a logic version, so improving the matching
  // (like using Daily Cache) triggers one automatic re-scan of missed draws.
  const marker = currentHistRowCount + '|' + BACKFILL_LOGIC_VERSION;
  if (props.getProperty('backfillLastHistRowCount') === marker) {
    return 0; // nothing new since the last scan — skip both full reads below
  }

  const fakeVsRealSheet = ss.getSheetByName(TABS.FAKE_VS_REAL);
  const alreadyCompared = new Set();
  if (fakeVsRealSheet && fakeVsRealSheet.getLastRow() > 1) {
    fakeVsRealSheet.getRange(2, 1, fakeVsRealSheet.getLastRow() - 1, 1).getValues().forEach(row => {
      const d = row[0];
      const val = d instanceof Date
        ? Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd')
        : String(d).trim();
      alreadyCompared.add(val);
    });
  }

  const rows = histSheet.getRange(2, 1, currentHistRowCount - 1, 7).getValues();
  const guessIndex = buildGuessIndex_(); // ONE read of the guess tabs for the whole scan
  let backfilled = 0;
  rows.forEach(row => {
    const dateStr = row[0] instanceof Date
      ? Utilities.formatDate(row[0], Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd')
      : String(row[0]).trim();
    if (!dateStr || alreadyCompared.has(dateStr)) return;
    const p = (v, n) => String(v).trim().padStart(n, '0');
    const real = {
      jackpot: p(row[1], 6), front3A: p(row[2], 3), front3B: p(row[3], 3),
      back3A: p(row[4], 3), back3B: p(row[5], 3), back2: p(row[6], 2),
    };
    if (!guessIndex[nominalSlot_(dateStr)]) return; // no Lucky Draw set for that draw — nothing to compare
    recordFakeVsRealComparison_(dateStr, real, guessIndex);
    alreadyCompared.add(dateStr); // avoid double-processing if two rows ever share a date
    backfilled++;
  });

  props.setProperty('backfillLastHistRowCount', marker);

  if (backfilled > 0) {
    Logger.log('Backfill: checked ' + backfilled + ' Historical Draws row(s) not yet in Fake vs Real ' +
      '(some may have been skipped inside recordFakeVsRealComparison_ if no Lucky Draw guess existed for that period).');
  }
  return backfilled;
}

// Manual escape hatch — run this by hand from the Script Editor if you ever
// need to force a full backfill re-scan even though Historical Draws' row
// count hasn't changed (e.g. you logged a NEW Lucky Draw guess retroactively
// for a period that was previously skipped for having no guess to compare
// against). Clears the cached row count so the next getComparisonStats()
// call runs the full scan again, one time.
function forceBackfillRecheck() {
  PropertiesService.getDocumentProperties().deleteProperty('backfillLastHistRowCount');
  Logger.log('Backfill cache cleared — the next dashboard load will run a full re-scan.');
}

// Called from the frontend (google.script.run.getComparisonStats()) to feed
// the "fake vs real" stat-tile dashboard.
//
// IMPORTANT: this does NOT call fetchAndArchiveRealDraw() here anymore. That
// was tried in an earlier revision to make every dashboard load "fetch the
// latest real draw up front", but it made every single page load perform a
// live outbound UrlFetchApp call (to GitHub, and potentially the Gemini
// fallback) before the stats could render — on a slow network response,
// missing authorization, or a stalled connection, that blocks this ENTIRE
// function for a long time (Apps Script's web-app execution limit is several
// minutes), which reads to the user as "stuck on Loading comparison stats…
// forever" — exactly the symptom reported. Fetching the real draw now
// happens ONLY on the once-daily time-driven trigger (dailyRealDrawFetchCheck
// -> fetchAndArchiveRealDraw(), set up via createRealDrawFetchTrigger()),
// which runs in the background and never blocks a page load. This function
// stays fast and 100% local: it just runs the backfill (so any real draw
// already archived — by the trigger OR typed in by hand — gets matched
// against a Lucky Draw guess) and reads the Fake vs Real tab.
// Compares two digit strings position by position, returning a boolean
// array (true = that position matched). Used both to accumulate
// per-position hit rates for the aggregate stat tiles (answers "WHICH
// digits does the math actually get right, not just how many") and to
// figure out which specific fake/real pair produced an exact category
// match for the Hall of Fame.
function positionMatchArray_(a, b) {
  const sa = String(a), sb = String(b);
  const len = Math.min(sa.length, sb.length);
  const arr = [];
  for (let i = 0; i < len; i++) arr.push(sa[i] === sb[i]);
  return arr;
}

// For a 2-fake-number-vs-2-real-number category (front3/back3), finds
// whichever of the 4 possible pairings scores highest and returns its
// per-position match array plus which specific numbers made that pairing —
// this is what lets the Hall of Fame show the actual number that hit,
// not just a count.
function bestPairingDetail_(fakeA, fakeB, realA, realB) {
  const combos = [
    { fake: fakeA, real: realA }, { fake: fakeA, real: realB },
    { fake: fakeB, real: realA }, { fake: fakeB, real: realB },
  ];
  let best = null;
  combos.forEach(c => {
    const arr = positionMatchArray_(c.fake, c.real);
    const total = arr.filter(Boolean).length;
    if (!best || total > best.total) best = { fake: c.fake, real: c.real, arr, total };
  });
  return best;
}

// One Hall of Fame list for every source, newest first, each entry tagged
// with where it came from: 'lucky' (wheel), 'all' (Every Draw column),
// 'slot' (Same-Date column).
function mergeFame_(luckyFame, mathStats) {
  let all = luckyFame.slice();
  if (mathStats) MATH_COLUMNS_.forEach(col => { all = all.concat((mathStats[col] && mathStats[col].hallOfFame) || []); });
  all.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return all.slice(0, 40);
}

function getComparisonStats() {
  backfillFakeVsRealComparisons();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.FAKE_VS_REAL);
  let sourceHits = null;
  try { sourceHits = computeSourceHits_(); } catch (e) { Logger.log('Source hits failed: ' + e); }
  let mathStats = null;
  try { mathStats = readMathBacktestStats_(); } catch (e) { Logger.log('Math backtest read failed: ' + e); }
  if (!sheet || sheet.getLastRow() < 2) {
    return { totalCompared: 0, sourceHits, mathStats, hallOfFame: mergeFame_([], mathStats) };
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).getValues();
  let flipHitsFront3 = 0, flipHitsBack3 = 0, flipHitsBack2 = 0;
  const pad = (v, n) => String(v).trim().padStart(n, '0');
  const totalCompared = data.length;
  let jackpotSum = 0, front3Sum = 0, back3Sum = 0, back2Sum = 0;
  let exactHitsJackpot = 0, exactHitsFront3 = 0, exactHitsBack3 = 0, exactHitsBack2 = 0;
  let latest = null;
  let latestDateKey = '';

  // Per-position hit counters — jackpotPos[2] is how many of the compared
  // draws had position 3 (0-indexed 2) come out right, etc. Turned into a
  // % rate per position at the end, so the frontend can show "which
  // digits" instead of one confusing combined fraction.
  const jackpotPos = new Array(6).fill(0);
  const front3Pos = new Array(3).fill(0);
  const back3Pos = new Array(3).fill(0);
  const back2Pos = new Array(2).fill(0);

  // Every draw where ANY category landed a full exact match — shown as a
  // standalone "Hall of Fame" list (date + the actual number + which
  // category), instead of buried inside one "best match ever" stat tile.
  const hallOfFame = [];

  data.forEach(row => {
    const drawDate = row[0], realJackpot = row[1], fakeJackpot = row[2], jackpotMatches = Number(row[3]);
    const realFront3A = row[4], realFront3B = row[5], fakeFront3A = row[6], fakeFront3B = row[7];
    const front3Best = Number(row[8]);
    const realBack3A = row[9], realBack3B = row[10], fakeBack3A = row[11], fakeBack3B = row[12];
    const back3Best = Number(row[13]);
    const realBack2 = row[14], fakeBack2 = row[15], back2Matches = Number(row[16]);

    const dateKey = drawDate instanceof Date
      ? Utilities.formatDate(drawDate, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd')
      : String(drawDate).trim();

    jackpotSum += jackpotMatches;
    front3Sum += front3Best;
    back3Sum += back3Best;
    back2Sum += back2Matches;

    positionMatchArray_(fakeJackpot, realJackpot).forEach((hit, i) => { if (hit) jackpotPos[i]++; });
    positionMatchArray_(fakeBack2, realBack2).forEach((hit, i) => { if (hit) back2Pos[i]++; });
    const front3Detail = bestPairingDetail_(fakeFront3A, fakeFront3B, realFront3A, realFront3B);
    front3Detail.arr.forEach((hit, i) => { if (hit) front3Pos[i]++; });
    const back3Detail = bestPairingDetail_(fakeBack3A, fakeBack3B, realBack3A, realBack3B);
    back3Detail.arr.forEach((hit, i) => { if (hit) back3Pos[i]++; });

    if (jackpotMatches === 6) {
      exactHitsJackpot++;
      hallOfFame.push({ date: dateKey, category: 'jackpot', number: String(realJackpot) });
    }
    if (back2Matches === 2) {
      exactHitsBack2++;
      hallOfFame.push({ date: dateKey, category: 'back2', number: String(realBack2) });
    }
    if (front3Best === 3) {
      exactHitsFront3++;
      hallOfFame.push({ date: dateKey, category: 'front3', number: String(front3Detail.real) });
    }
    if (back3Best === 3) {
      exactHitsBack3++;
      hallOfFame.push({ date: dateKey, category: 'back3', number: String(back3Detail.real) });
    }

    // Flip hits (กลับเลข): a lucky number that is a reordering of a real one.
    const anyFlip = (fakes, reals) => fakes.some(f => reals.some(r => isFlipOf_(f, r)));
    const f3 = [pad(fakeFront3A, 3), pad(fakeFront3B, 3)], r3 = [pad(realFront3A, 3), pad(realFront3B, 3)];
    const b3 = [pad(fakeBack3A, 3), pad(fakeBack3B, 3)], rb3 = [pad(realBack3A, 3), pad(realBack3B, 3)];
    if (front3Best < 3 && anyFlip(f3, r3)) {
      flipHitsFront3++;
      hallOfFame.push({ date: dateKey, category: 'front3', number: r3.find(r => f3.some(f => isFlipOf_(f, r))), flip: true });
    }
    if (back3Best < 3 && anyFlip(b3, rb3)) {
      flipHitsBack3++;
      hallOfFame.push({ date: dateKey, category: 'back3', number: rb3.find(r => b3.some(f => isFlipOf_(f, r))), flip: true });
    }
    if (back2Matches < 2 && isFlipOf_(pad(fakeBack2, 2), pad(realBack2, 2))) {
      flipHitsBack2++;
      hallOfFame.push({ date: dateKey, category: 'back2', number: pad(realBack2, 2), flip: true });
    }

    // Track the chronologically most recent row — used to feed the
    // always-visible "Lucky Draw vs. Real Draw" side-by-side boxes on the
    // dashboard. Draw Date is stored as plain text ('yyyy-MM-dd'), so a
    // plain string comparison sorts correctly; fall back to Date handling
    // in case a row was ever entered as a real Date value.
    if (dateKey > latestDateKey) {
      latestDateKey = dateKey;
      latest = {
        date: dateKey,
        jackpot: { real: String(realJackpot), fake: String(fakeJackpot), matches: jackpotMatches, max: 6 },
        front3: { real: String(realFront3A), realB: String(realFront3B), fake: String(fakeFront3A), fakeB: String(fakeFront3B), matches: front3Best, max: 3 },
        back3: { real: String(realBack3A), realB: String(realBack3B), fake: String(fakeBack3A), fakeB: String(fakeBack3B), matches: back3Best, max: 3 },
        back2: { real: String(realBack2), fake: String(fakeBack2), matches: back2Matches, max: 2 },
      };
    }
  });

  // Newest first, capped at a reasonable display length so a long-running
  // sheet doesn't dump hundreds of rows onto the dashboard.
  hallOfFame.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));

  const toRatePct = (count) => Math.round((count / totalCompared) * 100);

  return {
    totalCompared,
    avgJackpotMatch: +(jackpotSum / totalCompared).toFixed(2),
    avgFront3Match: +(front3Sum / totalCompared).toFixed(2),
    avgBack3Match: +(back3Sum / totalCompared).toFixed(2),
    avgBack2Match: +(back2Sum / totalCompared).toFixed(2),
    posMatchRates: {
      jackpot: jackpotPos.map(toRatePct),
      front3: front3Pos.map(toRatePct),
      back3: back3Pos.map(toRatePct),
      back2: back2Pos.map(toRatePct),
    },
    exactHitsJackpot,
    exactHitsFront3,
    exactHitsBack3,
    exactHitsBack2,
    hallOfFame: mergeFame_(hallOfFame.map(h => Object.assign({ source: 'lucky' }, h)), mathStats),
    // What pure luck would give with the Lucky Draw's picks over the same draws.
    expectedExact: {
      jackpot: +(totalCompared / 1e6).toFixed(3),
      front3: +(totalCompared * (1 - Math.pow(1 - 2 / 1000, 2))).toFixed(2),
      back3: +(totalCompared * (1 - Math.pow(1 - 2 / 1000, 2))).toFixed(2),
      back2: +(totalCompared / 100).toFixed(2),
    },
    mathStats,
    flipHitsFront3,
    flipHitsBack3,
    flipHitsBack2,
    sourceHits,
    latest,
  };
}

// ==========================================================================
// SHARED HELPERS for Flips / Dreams / Buzz
// ==========================================================================

// Calls Gemini. useSearch=true turns on Google Search grounding. Returns
// { text, sources:[{title,url}] } or null. JSON is requested in the prompt
// and parsed from the text (Gemini does not allow a JSON response mode
// together with the search tool).
function geminiModelChain_() {
  const custom = scriptProp_('GEMINI_MODEL', '');
  const chain = custom ? custom.split(',').map(x => x.trim()).filter(Boolean) : GEMINI_DEFAULT_MODELS.slice();
  return chain;
}

// Pulls the useful part out of a Gemini error (which quota was hit).
function describeGeminiError_(body) {
  try {
    const err = JSON.parse(body).error || {};
    const quotas = [];
    (err.details || []).forEach(d => (d.violations || []).forEach(v => {
      quotas.push((v.quotaId || v.quotaMetric || '') + (v.quotaValue ? ' (limit ' + v.quotaValue + ')' : ''));
    }));
    const retry = (err.details || []).map(d => d.retryDelay).filter(Boolean)[0];
    return err.status + ': ' + (quotas.length ? 'quota hit -> ' + quotas.join('; ') : String(err.message || '').slice(0, 200)) +
      (retry ? ' | retry after ' + retry : '');
  } catch (e) {
    return String(body).slice(0, 300);
  }
}

// Calls Gemini. useSearch=true turns on Google Search grounding. Tries each
// model in the chain; if one is rate-limited / unavailable (429, 404, 503)
// it moves on to the next. Returns { text, sources:[{title,url}], model }
// or null. JSON is requested in the prompt and parsed from the text.
function callGemini_(prompt, useSearch, preferLite) {
  const apiKey = scriptProp_('GEMINI_API_KEY', null);
  if (!apiKey) {
    Logger.log('Gemini skipped — GEMINI_API_KEY not set in Script Properties.');
    return null;
  }
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  if (useSearch) payload.tools = [{ google_search: {} }];
  let chain = geminiModelChain_();
  if (preferLite) chain = chain.filter(m => /lite/.test(m)).concat(chain.filter(m => !/lite/.test(m)));
  for (let i = 0; i < chain.length; i++) {
    const model = chain[i];
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify(payload), muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      if (code !== 200) {
        Logger.log('Gemini ' + model + (useSearch ? ' (+Google Search)' : '') + ' HTTP ' + code + ' — ' + describeGeminiError_(resp.getContentText()));
        if (code === 429 || code === 404 || code === 503 || code === 500) continue; // try next model
        return null;
      }
      const json = JSON.parse(resp.getContentText());
      const cand = json.candidates && json.candidates[0];
      if (!cand || !cand.content || !cand.content.parts) return null;
      const text = cand.content.parts.filter(p => !p.thought).map(p => p.text || '').join('');
      const chunks = (cand.groundingMetadata && cand.groundingMetadata.groundingChunks) || [];
      const sources = chunks.filter(c => c.web).map(c => ({ title: c.web.title || '', url: c.web.uri || '' }));
      const used = json.modelVersion || model;
      Logger.log('Gemini answered with ' + used + (used !== model ? ' (requested ' + model + ')' : '') + '.');
      return { text, sources, model: used };
    } catch (e) {
      Logger.log('Gemini ' + model + ' call failed: ' + e);
    }
  }
  Logger.log('All Gemini models refused (' + chain.join(', ') + '). See the lines above for which quota was hit.');
  return null;
}

// Run by hand to check your key + quota: one tiny call WITHOUT search and
// one WITH search, so the log shows exactly which part is blocked.
function testGemini() {
  const plain = callGemini_('Reply with the single word OK.', false);
  Logger.log('Without Google Search: ' + (plain ? 'WORKS (' + plain.model + ')' : 'FAILED'));
  const search = callGemini_('Search the web: what is today\'s date in Bangkok? Reply in one short line.', true);
  Logger.log('With Google Search: ' + (search ? 'WORKS (' + search.model + ')' : 'FAILED'));
}

// Pulls the first {...} or [...] JSON block out of model text.
function extractJson_(text) {
  if (!text) return null;
  const cleaned = String(text).replace(/```json|```/g, '');
  const start = cleaned.search(/[\[{]/);
  if (start === -1) return null;
  const open = cleaned[start], close = open === '{' ? '}' : ']';
  const end = cleaned.lastIndexOf(close);
  if (end <= start) return null;
  try { return JSON.parse(cleaned.slice(start, end + 1)); } catch (e) { return null; }
}

function safeUrl_(u) {
  const s = String(u || '').trim();
  return /^https?:\/\//i.test(s) ? s : '';
}

// The draw (1st/16th) a calendar day counts toward — same rule as the
// wheel's period key (draw days count toward their own draw).
function periodOfDate_(dateStr) {
  return nextDrawDate_(dateStr);
}

// ==========================================================================
// PHASE 1 — FLIPS (กลับเลข)
// ==========================================================================

// All distinct digit orders of a number: '19' -> 19, 91; '123' -> 6; '112' -> 3.
function flipsOf_(numStr) {
  const s = String(numStr);
  const out = {};
  (function permute(prefix, rest) {
    if (!rest.length) { out[prefix] = true; return; }
    for (let i = 0; i < rest.length; i++) permute(prefix + rest[i], rest.slice(0, i) + rest.slice(i + 1));
  })('', s);
  return Object.keys(out).sort();
}

function flipKey_(numStr) {
  return String(numStr).split('').sort().join('');
}

function isFlipOf_(a, b) {
  return String(a).length === String(b).length && a !== b && flipKey_(a) === flipKey_(b);
}

// Hot flip SETS for a 2/3-digit model: groups every number with its flips,
// scores the group per ticket (so a 6-number set isn't favoured just for
// being bigger), and reports how many times any member really came out.
function computeFlipGroups_(model, count) {
  if (!model.numberProbs) return [];
  const groups = {};
  model.numberProbs.forEach((p, n) => {
    const s = String(n).padStart(model.len, '0');
    const k = flipKey_(s);
    if (!groups[k]) groups[k] = { key: k, members: [], prob: 0, seen: 0 };
    groups[k].members.push(s);
    groups[k].prob += p;
    groups[k].seen += model.numberRaw[n];
  });
  const base = Math.pow(10, model.len);
  return Object.keys(groups).map(k => groups[k])
    .filter(g => g.members.length >= 2)
    .map(g => Object.assign(g, { lift: g.prob / (g.members.length / base) }))
    .sort((a, b) => b.lift - a.lift)
    .slice(0, count)
    .map(g => ({
      digits: g.key.split('').join('·'),
      members: g.members,
      size: g.members.length,
      seen: g.seen,
      lift: +g.lift.toFixed(2),
    }));
}

// Frontend helper: flips for any list of numbers in one call.
function getFlips(numbers) {
  const out = {};
  (numbers || []).forEach(n => {
    const s = String(n).trim();
    if (/^\d{2,3}$/.test(s)) out[s] = flipsOf_(s);
  });
  return out;
}

// Checks a list of numbers against one real draw. Returns hits like
// { number, kind:'exact'|'flip', where:'back2'|'front3'|'back3'|'jackpot'|'jackpotLast3' }.
// Single digits are ignored (they'd "hit" almost every draw).
function numberHitsAgainstDraw_(numbers, real) {
  const hits = [];
  const seen = {};
  numbers.forEach(raw => {
    const n = String(raw).trim();
    if (seen[n] || !/^\d{2,6}$/.test(n)) return;
    seen[n] = true;
    if (n.length === 2) {
      if (n === real.back2) hits.push({ number: n, kind: 'exact', where: 'back2' });
      else if (isFlipOf_(n, real.back2)) hits.push({ number: n, kind: 'flip', where: 'back2' });
    } else if (n.length === 3) {
      const slots = [['front3', real.front3A], ['front3', real.front3B], ['back3', real.back3A], ['back3', real.back3B],
        ['jackpotLast3', String(real.jackpot).slice(-3)]];
      const exact = slots.find(s => s[1] === n);
      if (exact) hits.push({ number: n, kind: 'exact', where: exact[0] });
      else {
        const flip = slots.find(s => isFlipOf_(n, s[1]));
        if (flip) hits.push({ number: n, kind: 'flip', where: flip[0] });
      }
    } else if (n.length === 6 && n === real.jackpot) {
      hits.push({ number: n, kind: 'exact', where: 'jackpot' });
    }
  });
  return hits;
}

// ==========================================================================
// PHASE 2 — DREAMS (ทำนายฝัน)
// ==========================================================================

const DREAM_SOURCE_URLS_ = {
  general: 'https://shopee.co.th/blog/dream-prophecy/',
  cat: 'https://shopee.co.th/blog/cat-dreaming/',
  poop: 'https://shopee.co.th/blog/you-dream-interpreter-%e0%b8%9d%e0%b8%b1%e0%b8%99%e0%b9%80%e0%b8%ab%e0%b9%87%e0%b8%99%e0%b8%82%e0%b8%b5%e0%b9%89',
  gold: 'https://shopee.co.th/blog/dream-of-gold-interpretation/',
  monk: 'https://shopee.co.th/blog/dreaming-of-monk/',
};

// Starter dream dictionary: [keyword TH, aliases (comma), English, numbers, source key].
// Numbers copied from the public Thai dream-number articles above. Edit or
// extend freely in the Dream Dictionary tab — the app reads the tab, not this list.
const DREAM_SEED_ = [
  ['งู', '', 'snake', '56 66 599 55 568', 'general'],
  ['งูกัด', 'โดนงูกัด', 'snake bite', '95 864 71 178 94', 'general'],
  ['งูเลื้อย', '', 'snake slithering past', '16 61 166 561 665 666', 'general'],
  ['งูใหญ่', 'งูตัวใหญ่', 'big snake', '07 45 90 65 95', 'general'],
  ['งูเหลือม', '', 'python', '038 83 438 061 34', 'general'],
  ['งูเห่า', '', 'cobra', '804 779 86 25 817', 'general'],
  ['งูหลายตัว', '', 'many snakes', '683 490 35 94 82', 'general'],
  ['งูจงอาง', '', 'king cobra', '819 78 197 15', 'general'],
  ['งูเขียว', '', 'green snake', '26 62 726 35 642', 'general'],
  ['พญานาค', 'นาค', 'naga', '19 56 59 569 655 685', 'general'],
  ['ช้าง', '', 'elephant', '19 39 119 139 309', 'general'],
  ['เสือ', '', 'tiger', '34 30 349 309 344', 'general'],
  ['กบ', '', 'frog', '62 26 71 17 16 61 12 21', 'general'],
  ['ปู', '', 'crab', '4 5 9 8 45 56 98', 'general'],
  ['กุ้ง', '', 'shrimp', '89 79 78 788 889 709', 'general'],
  ['กุ้งใหญ่', 'กุ้งตัวใหญ่', 'big shrimp', '78 79 89 788 899', 'general'],
  ['หมาไล่กัดแมว', 'สุนัขไล่กัดแมว', 'dog chasing a cat', '73 78 10 976 170', 'general'],
  ['ท้อง', 'ตั้งครรภ์', 'pregnant', '10 19 59 101 195 910', 'general'],
  ['ตาย', '', 'dying', '07 04 407 714 447', 'general'],
  ['ศพ', '', 'corpse', '11 41 84 185 441 108', 'general'],
  ['งานศพ', '', 'funeral', '47 27 48 437 427 484', 'general'],
  ['ตัดผม', '', 'haircut', '19 18 16 196 918', 'general'],
  ['ฟันหลุด', 'ฟันหัก,ฟันร่วง', 'tooth falling out', '03 05 99 77 44 58', 'general'],
  ['ถอนฟัน', 'ดึงฟัน', 'pulling a tooth', '06 11 28 34 710', 'general'],
  ['แฟนเก่า', '', 'ex-partner', '06 67 661 675 766', 'general'],
  ['เปลือย', 'แก้ผ้า', 'naked', '10 594 504 533 423', 'general'],
  ['รถหาย', '', 'car stolen', '7 4 5 9 65 44 304', 'general'],
  ['เงินหาย', '', 'lost money', '79 74 92 13 59', 'general'],
  ['แมว', '', 'cat', '7 13 14 38 72', 'cat'],
  ['ลูกแมว', '', 'kitten', '1 13 34 42 59', 'cat'],
  ['แมวดำ', 'แมวสีดำ', 'black cat', '6 28 30 48 74', 'cat'],
  ['แมวขาว', 'แมวสีขาว', 'white cat', '81 40 31 75 71 14 33', 'cat'],
  ['แมวสามสี', '', 'calico cat', '83 54 49 34 88', 'cat'],
  ['แมวออกลูก', '', 'cat giving birth', '1 5 67 82', 'cat'],
  ['แมวตาย', '', 'dead cat', '4 43 86 42', 'cat'],
  ['ขี้', 'อุจจาระ', 'poop', '31 91 94 494 955', 'poop'],
  ['ขี้หมา', 'อุจจาระหมา,ขี้สุนัข', 'dog poop', '05 13 25 325 510', 'poop'],
  ['ขี้แมว', 'อุจจาระแมว', 'cat poop', '23 56 58 258', 'poop'],
  ['ขี้ไก่', '', 'chicken poop', '25 01 43 015 205 512', 'poop'],
  ['กินขี้', 'กินอุจจาระ', 'eating poop', '21 11 73', 'poop'],
  ['ทอง', '', 'gold', '4 7 15 39 47 79', 'gold'],
  ['ได้ทอง', '', 'receiving gold', '6 8 9 76 79 95', 'gold'],
  ['ทองคำแท่ง', 'ทองแท่ง', 'gold bar', '6 7 30 42 60 67', 'gold'],
  ['ซื้อทอง', '', 'buying gold', '3 9 36 48 63 69', 'gold'],
  ['พระสงฆ์', 'พระ', 'monk', '61 63 786 962', 'monk'],
  ['บิณฑบาต', 'ตักบาตร', 'monk on alms round', '15 25 527 761', 'monk'],
  ['พระพุทธรูป', '', 'Buddha image', '9 91 92 198 199 299', 'monk'],
  ['พระเครื่อง', '', 'amulet', '27 52 78 275 307', 'monk'],
  ['ไหว้พระ', '', 'praying to Buddha', '88 27 78 288 287 289', 'monk'],
  ['พระธุดงค์', '', 'wandering monk', '61 63 786 962', 'monk'],
  ['สังฆทาน', 'ถวายเทียน', 'offering to monks', '17 62 70 44 460', 'monk'],
];

// Extra dream entries (70) from more public Thai dream-number articles.
// Added to an existing sheet by running addMoreDreamWords() once.
const DREAM_SOURCE_MORE_URLS_ = {
  croc: 'https://shopee.co.th/blog/dream-crocodile-interpretation/',
  gecko: 'https://shopee.co.th/blog/dreaming-saw-gecko/',
  elephant: 'https://shopee.co.th/blog/dream-elephants-interpretation/',
  baby: 'https://shopee.co.th/blog/dream-fortune-teller-pregnancy/',
  partner: 'https://shopee.co.th/blog/dream-of-cheating-interpretation',
  naga: 'https://shopee.co.th/blog/dream-naga-interpretation/',
  blood: 'https://shopee.co.th/blog/dreaming-of-blood/',
  death: 'https://shopee.co.th/blog/dreaming-of-own-death/',
  snake: 'https://shopee.co.th/blog/dream-of-snake-interpretation/',
  hair: 'https://shopee.co.th/blog/dream-fortune-teller-short-haircut',
  turtle: 'https://shopee.co.th/blog/?p=127528',
};
const DREAM_SEED_MORE_ = [
  ['จระเข้', 'จระเข้ 1 ตัว', 'crocodile', '45 71 119 701', 'croc'],
  ['จระเข้ตัวเล็ก', 'ลูกจระเข้', 'small crocodile', '02 19 209 902', 'croc'],
  ['จระเข้ตัวใหญ่', 'จระเข้ใหญ่', 'big crocodile', '08 85 889 981', 'croc'],
  ['จระเข้เผือก', '', 'white crocodile', '27 59 708 995', 'croc'],
  ['จระเข้ 2 ตัว', '', 'two crocodiles', '22 28 730 884', 'croc'],
  ['จระเข้หลายตัว', '', 'many crocodiles', '26 77 941 977', 'croc'],
  ['จระเข้ไล่กัด', 'จระเข้กัด', 'crocodile attack', '09 71 359 940', 'croc'],
  ['จระเข้ในน้ำ', 'จระเข้อยู่ในน้ำ', 'crocodile in water', '26 59 352 864', 'croc'],
  ['จระเข้เข้าบ้าน', '', 'crocodile in the house', '44 55 448 909', 'croc'],
  ['จระเข้ตาย', '', 'dead crocodile', '79 85 533 819', 'croc'],
  ['ตุ๊กแก', 'ตุ๊กแก 1 ตัว', 'gecko', '65 71 77 771 707 761 982 837', 'gecko'],
  ['ตุ๊กแก 2 ตัว', '', 'two geckos', '22 29 30 44 62 90 716 143', 'gecko'],
  ['ตุ๊กแกหลายตัว', '', 'many geckos', '34 57 91 55 881 253', 'gecko'],
  ['ตุ๊กแกตัวใหญ่', '', 'big gecko', '41 83 86 521 564 837 982', 'gecko'],
  ['จับตุ๊กแก', '', 'catching a gecko', '47 62 58 982 521 837', 'gecko'],
  ['ตุ๊กแกกัด', '', 'gecko bite', '31 36 87 451 58 991', 'gecko'],
  ['ตุ๊กแกร้อง', 'ตุ๊กแกร้องทัก', 'gecko calling', '14 77 98 155 320 964', 'gecko'],
  ['ช้าง 2 ตัว', '', 'two elephants', '22 29 202', 'elephant'],
  ['ช้างตัวใหญ่', 'ช้างใหญ่', 'big elephant', '95 99 199', 'elephant'],
  ['ช้างหลายตัว', '', 'many elephants', '36 91 369', 'elephant'],
  ['ช้างเผือก', '', 'white elephant', '11 39 159', 'elephant'],
  ['ขี่ช้าง', '', 'riding an elephant', '19 59 309', 'elephant'],
  ['ช้างออกลูก', 'ลูกช้าง', 'elephant giving birth', '25 64 224', 'elephant'],
  ['คลอดลูก', '', 'giving birth', '39 41 66 148 339 669', 'baby'],
  ['ลูกแฝด', 'คลอดแฝด', 'twins', '5 33 46 81 92 494 903', 'baby'],
  ['ให้นมลูก', '', 'breastfeeding', '38 158 915 623', 'baby'],
  ['ทารก', 'เด็กแรกเกิด,เด็กทารก', 'newborn baby', '17 13 33 317 311 513', 'baby'],
  ['แฟนนอกใจ', '', 'partner cheating', '1 5 9 21 39 179', 'partner'],
  ['แฟนมีคนอื่น', 'แฟนมีชู้', 'partner has someone else', '1 4 6 36 64 612', 'partner'],
  ['แฟนแต่งงานกับคนอื่น', '', 'partner marries someone else', '2 9 26 66 96 451', 'partner'],
  ['แฟนตาย', '', 'partner dies', '0 4 7 59 78 329', 'partner'],
  ['พญานาค 2 ตัว', 'นาค 2 ตัว', 'two nagas', '20 56 69 560 695', 'naga'],
  ['พญานาคหลายตัว', '', 'many nagas', '64 76 98 188 789', 'naga'],
  ['พญานาคสีเขียว', 'นาคเขียว', 'green naga', '40 74 93 679 774', 'naga'],
  ['พญานาคสีทอง', 'นาคทอง', 'golden naga', '13 59 84 158 789', 'naga'],
  ['พญานาคสีขาว', 'นาคขาว', 'white naga', '36 58 69 109 508', 'naga'],
  ['พญานาคสีดำ', 'นาคดำ', 'black naga', '18 50 66 768 884', 'naga'],
  ['พญานาคเข้าบ้าน', '', 'naga in the house', '30 46 54 549 945', 'naga'],
  ['พญานาคบนท้องฟ้า', '', 'naga in the sky', '09 56 65 509 990', 'naga'],
  ['เลือด', 'เลือดออก', 'bleeding', '16 76 88 726', 'blood'],
  ['บริจาคเลือด', '', 'donating blood', '15 21 52 63', 'blood'],
  ['เลือดเต็ม', 'เลือดเต็มไปหมด', 'blood everywhere', '77 34 152', 'blood'],
  ['เลือดคนอื่น', '', "someone else's blood", '44 98 572 786', 'blood'],
  ['ตัวเองตาย', '', 'dying yourself', '34 41 47 84 047 471 711 469 994', 'death'],
  ['ตายแล้วฟื้น', '', 'dying then coming back', '15 67 575 725', 'death'],
  ['ขับรถชน', 'รถชน', 'car crash', '65 26 774 265', 'death'],
  ['ตกจากที่สูง', '', 'falling from a height', '59 95 599', 'death'],
  ['ญาติที่ตาย', 'ญาติที่ตายไปแล้ว,ญาติที่เสียชีวิต', 'seeing a dead relative', '47 74 79 97 799 749 904', 'death'],
  ['โลงศพ', '', 'coffin', '04 44 440 442 448', 'death'],
  ['คนตายชวน', '', 'a dead person inviting you', '89 98 189', 'death'],
  ['งูรัด', 'ถูกงูรัด', 'snake squeezing you', '55 56 66', 'snake'],
  ['ตีงู', 'ฆ่างู', 'killing a snake', '13 14 34 42 59', 'snake'],
  ['งูเข้าบ้าน', 'งูเลื้อยเข้าบ้าน', 'snake entering the house', '04 27 32 35 51', 'snake'],
  ['งูตาย', '', 'dead snake', '56 06 61 67 57', 'snake'],
  ['งูตัวเล็ก', 'งูเล็ก', 'small snake', '07 30 38 57 63', 'snake'],
  ['งูเผือก', 'งูขาว', 'white snake', '19 74 48 96 17', 'snake'],
  ['งูดำ', 'งูสีดำ', 'black snake', '14 78 87 59 95', 'snake'],
  ['งูไล่', '', 'snake chasing you', '13 979 630 806', 'snake'],
  ['จับงู', '', 'catching a snake', '06 416 202', 'snake'],
  ['งูกินคน', '', 'snake eating a person', '29 32 38 56 83', 'snake'],
  ['ตัดผมสั้น', '', 'short haircut', '05 07 15 75 922 243', 'hair'],
  ['เต่า', '', 'turtle', '41 14 04 401 114', 'turtle'],
  ['เต่าตัวใหญ่', 'เต่ายักษ์', 'big turtle', '28 48 045 439 945', 'turtle'],
  ['เต่าหลายตัว', '', 'many turtles', '12 48 465 745 848', 'turtle'],
  ['เต่าออกไข่', '', 'turtle laying eggs', '08 47 784 844 940', 'turtle'],
  ['ลูกเต่า', '', 'baby turtle', '00 30 104 403 334', 'turtle'],
  ['เต่าเข้าบ้าน', '', 'turtle in the house', '33 40 464 747 049', 'turtle'],
  ['ปล่อยเต่า', '', 'releasing a turtle', '79 54 944 465 905', 'turtle'],
  ['เต่าตาย', '', 'dead turtle', '10 13 050 120 024', 'turtle'],
  ['เต่าทอง', 'เต่าสีทอง', 'golden turtle', '78 97 949 402 932', 'turtle'],
];

// Run ONCE from the editor to add the extra entries to your Dream
// Dictionary tab. Skips any keyword you already have — never overwrites.
function addMoreDreamWords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = seedDreamDictionary_(ss);
  const have = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(r => { have[String(r[0]).trim()] = true; });
  }
  const rows = DREAM_SEED_MORE_
    .filter(r => !have[r[0]])
    .map(r => [r[0], r[1], r[2], r[3], DREAM_SOURCE_MORE_URLS_[r[4]] || '']);
  if (rows.length) {
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 4, rows.length, 1).setNumberFormat('@');
    sheet.getRange(start, 1, rows.length, 5).setValues(rows);
  }
  Logger.log('Dream Dictionary: added ' + rows.length + ' new entries (' + (DREAM_SEED_MORE_.length - rows.length) + ' already there).');
}

const DREAM_DICT_HEADERS_ = ['Keyword (TH)', 'Also matches (comma)', 'English', 'Numbers', 'Source'];

function seedDreamDictionary_(ss) {
  const sheet = ensureTab_(ss, TABS.DREAM_DICTIONARY, DREAM_DICT_HEADERS_, [4]);
  if (sheet.getLastRow() > 1) return sheet; // never overwrite your edits
  const rows = DREAM_SEED_.map(r => [r[0], r[1], r[2], r[3], DREAM_SOURCE_URLS_[r[4]] || ''])
    .concat(DREAM_SEED_MORE_.map(r => [r[0], r[1], r[2], r[3], DREAM_SOURCE_MORE_URLS_[r[4]] || '']));
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  sheet.getRange(1, 7).setValue('Add your own rows any time. Numbers: separate with spaces. "Also matches" = other words for the same thing.');
  return sheet;
}

function readDreamDictionary_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.DREAM_DICTIONARY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .map(r => ({
      keyword: String(r[0]).trim(),
      aliases: String(r[1]).split(',').map(s => s.trim()).filter(Boolean),
      english: String(r[2]).trim(),
      numbers: String(r[3]).match(/\d+/g) || [],
      source: safeUrl_(r[4]),
    }))
    .filter(e => e.keyword && e.numbers.length);
}

// For the Dreams page: keyword chips to tap.
function getDreamDictionary() {
  return readDreamDictionary_().map(e => ({ keyword: e.keyword, english: e.english }));
}

// Finds dictionary entries inside free dream text. Longer (more specific)
// matches win: "งูเห่า" beats plain "งู" for the same words.
function matchDreamEntries_(text, dict) {
  const t = String(text || '').toLowerCase();
  const found = [];
  dict.forEach(e => {
    const terms = [e.keyword].concat(e.aliases).filter(Boolean);
    let best = null;
    terms.forEach(term => {
      const low = term.toLowerCase();
      const idx = t.indexOf(low);
      if (idx !== -1 && (!best || low.length > best.len)) best = { idx, len: low.length };
    });
    if (!best && e.english) {
      const en = e.english.toLowerCase();
      const idx = t.indexOf(en);
      if (idx !== -1) best = { idx, len: en.length };
    }
    if (best) found.push({ entry: e, start: best.idx, end: best.idx + best.len });
  });
  found.sort((a, b) => (b.end - b.start) - (a.end - a.start));
  const kept = [];
  found.forEach(f => {
    const covered = kept.some(k => f.start >= k.start && f.end <= k.end);
    if (!covered) kept.push(f);
  });
  return kept.sort((a, b) => a.start - b.start).map(f => f.entry);
}

// Main dream endpoint. Returns matched entries + combined suggestion list.
function interpretDream(dreamText) {
  const text = String(dreamText || '').trim().slice(0, 500);
  if (!text) return { matches: [], suggestions: [], usedAI: false };
  const dict = readDreamDictionary_();
  let matches = matchDreamEntries_(text, dict);
  let usedAI = false;

  // Nothing matched word-for-word -> ask Gemini to pick the closest
  // dictionary keywords. Numbers still come ONLY from your dictionary.
  if (matches.length === 0 && dict.length) {
    const list = dict.map(e => e.keyword).join(', ');
    const res = callGemini_(
      'A Thai person describes their dream: "' + text + '". From this exact keyword list, choose up to 3 ' +
      'keywords that best match what appears in the dream: [' + list + ']. ' +
      'Reply ONLY with JSON like {"keywords":["งู"]}. If nothing fits, reply {"keywords":[]}.', false, true); // simple task -> Flash Lite first
    const parsed = res && extractJson_(res.text);
    const picked = (parsed && parsed.keywords) || [];
    matches = dict.filter(e => picked.includes(e.keyword));
    usedAI = matches.length > 0;
  }

  // Combine: numbers suggested by more matched entries rank first, then
  // 2-digit before 3-digit before running single digits.
  const tally = {};
  matches.forEach(e => e.numbers.forEach(n => { tally[n] = (tally[n] || 0) + 1; }));
  const lenRank = n => (n.length === 2 ? 0 : n.length === 3 ? 1 : 2);
  const suggestions = Object.keys(tally)
    .sort((a, b) => tally[b] - tally[a] || lenRank(a) - lenRank(b))
    .slice(0, 10);

  logDream_(text, matches, suggestions);

  return {
    matches: matches.map(e => ({ keyword: e.keyword, english: e.english, numbers: e.numbers, source: e.source })),
    suggestions,
    usedAI,
    periodTargetDate: currentPeriodKey_(),
  };
}

function logDream_(text, matches, suggestions) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureTab_(ss, TABS.DREAM_LOG, ['Timestamp', 'Period Key', 'Dream', 'Matched', 'Numbers'], [2, 5]);
    appendRowForcedText_(sheet, [
      new Date(), currentPeriodKey_(), text,
      matches.map(e => e.keyword).join(', '), suggestions.join(' '),
    ], [2, 5]);
  } catch (e) {
    Logger.log('Dream log failed: ' + e);
  }
}

// ==========================================================================
// PHASE 3 — BUZZ NUMBERS from Thai media (via Gemini + Google Search)
// ==========================================================================

const BUZZ_TAG_ = '[auto-buzz]';

// Reads Google News (free public RSS) headlines about เลขเด็ด for the
// coming draw, then asks Gemini — WITHOUT Google Search, which works on the
// free tier — to pick out the lucky numbers each headline mentions. Every
// number keeps the link to its article. Appends to Signal Feed; duplicates
// are skipped, so it's safe to run any time.
function fetchBuzzNumbers() {
  const periodKey = currentPeriodKey_();
  const labels = formatDrawDateLabels_(periodKey);
  const d = parseYmd_(periodKey);
  const thMonthsShort = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const beShort = String(d.getFullYear() + 543).slice(-2);
  const queries = [
    'เลขเด็ด ' + d.getDate() + ' ' + thMonthsShort[d.getMonth()] + ' ' + beShort,
    'เลขเด็ด งวด ' + d.getDate() + ' ' + THAI_MONTH_NAMES_[d.getMonth()] + ' ' + (d.getFullYear() + 543),
  ];

  // Only headlines published since the previous draw count for this draw.
  const since = parseYmd_(previousDrawDate_(periodKey));
  const items = [];
  const seenTitles = {};
  let nBing = 0, nGoogle = 0, nOld = 0;
  queries.forEach(q => {
    const bing = fetchBingNewsRss_(q);
    const goog = fetchNewsRss_(q);
    nBing += bing.length; nGoogle += goog.length;
    bing.concat(goog).forEach(it => {
      if (!it.title || seenTitles[it.title]) return;
      if (it.date && it.date < since) { nOld++; return; }
      seenTitles[it.title] = true;
      items.push(it);
    });
  });
  Logger.log('Buzz sources: Bing News ' + nBing + ' items, Google News ' + nGoogle + ' items, ' + nOld + ' skipped as older than the last draw.');
  // Direct-link (openable) articles first, so the 12 we open are real pages.
  items.sort((a, b) => (/news\.google\.com/.test(a.link) ? 1 : 0) - (/news\.google\.com/.test(b.link) ? 1 : 0));
  if (!items.length) {
    Logger.log('Buzz: no news headlines found yet for ' + periodKey + '.');
    return { added: 0, found: 0 };
  }
  const list = items.slice(0, 30);
  // The numbers are usually in the ARTICLE, not the headline — read the
  // number-bearing sentences from the first articles that have a direct link.
  let fetched = 0;
  list.forEach(it => {
    if (fetched >= 12 || !it.link || /news\.google\.com/.test(it.link)) return;
    it.excerpt = fetchArticleExcerpt_(it.link);
    fetched++;
  });

  const res = callGemini_(
    'These are Thai news headlines about lucky numbers (เลขเด็ด) for the Thai Government Lottery ' + labels.th + '.\n' +
    list.map((it, k) => k + '. ' + it.title + (it.source ? ' [' + it.source + ']' : '') +
      (it.desc ? '\n   snippet: ' + it.desc.slice(0, 300) : '') +
      (it.excerpt ? '\n   article: ' + it.excerpt : '')).join('\n') + '\n\n' +
    'Pick out ONLY lottery numbers (2 to 6 digits) that an item (headline, snippet or article text) explicitly presents as a lucky number for this draw. ' +
    'Ignore dates (e.g. 16/10/69), years, prices, ages, counts and phone numbers. Max 30 items. ' +
    'Reply ONLY with JSON: {"items":[{"i":0,"number":"59","type":"News|Social Buzz|Car Plate|Birthday|Event|Other",' +
    '"why":"short Thai reason, max 80 chars"}]} where i is the headline index.', false);
  if (!res) return { added: 0, error: 'Gemini not available (check GEMINI_API_KEY / quota).' };

  const parsed = extractJson_(res.text);
  const found = (parsed && parsed.items) || [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureTab_(ss, TABS.SIGNAL_FEED, ['Date', 'Source Type', 'Raw Text', 'Extracted Number(s)', 'Notes'], [4]);
  const existing = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(r => {
      if (toDateStr_(r[0]) !== periodKey) return;
      const link = (String(r[4]).match(/https?:\/\/\S+/) || [''])[0];
      existing[String(r[3]).trim() + '|' + link] = true;
    });
  }

  // Same number from DIFFERENT articles is kept (that's what "mentioned by
  // N sources" counts); the same number from the same article is skipped.
  let added = 0, notInText = 0, already = 0;
  found.forEach(it => {
    const num = String(it.number || '').replace(/\D/g, '');
    const src = list[Number(it.i)];
    if (!/^\d{2,6}$/.test(num) || !src) return;
    const key = num + '|' + src.link;
    if (existing[key]) { already++; return; }
    // Guard against made-up numbers: the digits must really appear in what we read.
    const readText = [src.title, src.desc || '', src.excerpt || ''].join(' ');
    if (readText.indexOf(num) === -1) { notInText++; return; }
    const type = ALL_SOURCE_TYPES.includes(it.type) ? it.type : 'News';
    const why = String(it.why || src.title).slice(0, 120);
    appendRowForcedText_(sheet, [periodKey, type, why + (src.source ? ' — ' + src.source : ''), num, BUZZ_TAG_ + ' ' + src.link], [1, 4]);
    existing[key] = true;
    added++;
  });
  PropertiesService.getDocumentProperties().setProperty('buzzLastRun', todayString_());
  Logger.log('Buzz numbers for ' + periodKey + ': ' + list.length + ' headlines read (' + fetched + ' articles opened), ' +
    found.length + ' numbers found, ' + added + ' new added, ' + already + ' already saved, ' +
    notInText + ' rejected (not written in the article).');
  return { added, found: found.length, headlines: list.length };
}

// Google News RSS search (Thai edition). Returns [{title, link, source, date}].
function fetchNewsRss_(query) {
  const url = 'https://news.google.com/rss/search?q=' + encodeURIComponent(query) + '&hl=th&gl=TH&ceid=TH:th';
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      Logger.log('News RSS HTTP ' + resp.getResponseCode() + ' for "' + query + '"');
      return [];
    }
    return parseRssItems_(resp.getContentText());
  } catch (e) {
    Logger.log('News RSS failed for "' + query + '": ' + e);
    return [];
  }
}

// Bing News RSS (Thai market) — gives direct publisher links + snippets,
// which lets us open the real article (Google News links are redirects).
function fetchBingNewsRss_(query) {
  // Two URL styles — Bing sometimes answers one with an empty/blocked page.
  const urls = [
    'https://www.bing.com/news/search?q=' + encodeURIComponent(query) + '&format=RSS', // the one that works from Apps Script
    'https://www.bing.com/news/search?q=' + encodeURIComponent(query) + '&format=rss&setlang=th&cc=TH&mkt=th-TH',
  ];
  for (let i = 0; i < urls.length; i++) {
    try {
      const resp = UrlFetchApp.fetch(urls[i], {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; LotteryBuzzBot/1.0)', 'Accept': 'application/rss+xml, text/xml' },
      });
      const code = resp.getResponseCode();
      const body = resp.getContentText();
      const items = code === 200 ? parseRssItems_(body) : [];
      if (items.length) {
        return items.map(it => {
          // Bing wraps links as .../apiclick.aspx?...&url=<real url>
          const m = it.link.match(/[?&]url=([^&]+)/);
          if (m) { try { it.link = safeUrl_(decodeURIComponent(m[1])); } catch (e) {} }
          return it;
        });
      }
      Logger.log('Bing News try ' + (i + 1) + ': HTTP ' + code + ', 0 items' + (code === 200 ? ' (got ' + body.slice(0, 60).replace(/\s+/g, ' ') + '…)' : ''));
    } catch (e) {
      Logger.log('Bing News try ' + (i + 1) + ' failed: ' + e);
    }
  }
  return [];
}

// Opens one article and keeps only short sentences that mention เลข/หวย
// together with a number — that's where the lucky numbers are. Max ~900 chars.
function fetchArticleExcerpt_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (resp.getResponseCode() !== 200) return '';
    const text = resp.getContentText()
      .replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>|<noscript[\s\S]*?<\/noscript>/gi, ' ')
      .replace(/<[^>]+>/g, '\n')
      .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#\d+;/g, ' ');
    const keep = [];
    let total = 0;
    text.split(/\n+/).map(x => x.replace(/\s+/g, ' ').trim()).forEach(line => {
      if (total > 900 || line.length < 4 || line.length > 300) return;
      if (/(เลข|หวย|ทะเบียน|ธูป|งวด)/.test(line) && /\d{2,6}/.test(line) && keep.indexOf(line) === -1) {
        keep.push(line);
        total += line.length;
      }
    });
    return keep.join(' | ');
  } catch (e) {
    return '';
  }
}

function parseRssItems_(xml) {
  const decode = s => String(s || '')
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
    .trim();
  const tag = (block, name) => {
    const m = block.match(new RegExp('<' + name + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + name + '>'));
    return m ? decode(m[1]) : '';
  };
  const out = [];
  (String(xml).match(/<item>[\s\S]*?<\/item>/g) || []).forEach(block => {
    const source = tag(block, 'source');
    let title = tag(block, 'title');
    if (source && title.endsWith(' - ' + source)) title = title.slice(0, -(' - ' + source).length);
    const pub = tag(block, 'pubDate');
    const date = pub ? new Date(pub) : null;
    const desc = tag(block, 'description').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    out.push({
      title, link: safeUrl_(tag(block, 'link')), source: source || tag(block, 'News:Source'),
      desc: desc === title ? '' : desc, date: date && !isNaN(date.getTime()) ? date : null,
    });
  });
  return out;
}

// Called from the daily trigger: refresh buzz once a day during the last
// 7 days before a draw (that's when the media posts เลขเด็ด), never on a
// locked draw day.
function autoBuzzRefresh_() {
  if (lockStatus_().drawDayLocked) return;
  const periodKey = currentPeriodKey_();
  const daysLeft = Math.round((parseYmd_(periodKey) - parseYmd_(todayString_())) / 86400000);
  if (daysLeft > 7) return;
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('buzzLastRun') === todayString_()) return;
  fetchBuzzNumbers();
}

// Reads Signal Feed rows (auto + typed by hand) and groups them by period.
function readSignalRowsByPeriod_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.SIGNAL_FEED);
  const byPeriod = {};
  if (!sheet || sheet.getLastRow() < 2) return byPeriod;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(r => {
    const d = toDateStr_(r[0]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    const notes = String(r[4] || '');
    const urlMatch = notes.match(/https?:\/\/\S+/);
    const p = periodOfDate_(d);
    (byPeriod[p] = byPeriod[p] || []).push({
      type: String(r[1]).trim(), text: String(r[2] || ''),
      numbers: String(r[3]).match(/\d{2,6}/g) || [],
      url: urlMatch ? safeUrl_(urlMatch[0]) : '',
      auto: notes.indexOf(BUZZ_TAG_) !== -1,
    });
  });
  return byPeriod;
}

// Buzz box on the Lucky Draw page: this draw's numbers, most-mentioned first.
function getBuzzNumbers() {
  const periodKey = currentPeriodKey_();
  const rows = readSignalRowsByPeriod_()[periodKey] || [];
  const groups = {};
  rows.forEach(r => r.numbers.forEach(n => {
    const g = groups[n] = groups[n] || { number: n, mentions: 0, reasons: [] };
    g.mentions++;
    if (g.reasons.length < 3) g.reasons.push({ text: r.text.slice(0, 120), url: r.url, type: r.type });
  }));
  const list = Object.keys(groups).map(k => groups[k])
    .sort((a, b) => b.mentions - a.mentions || a.number.length - b.number.length)
    .slice(0, 16);
  const labels = formatDrawDateLabels_(periodKey);
  return {
    periodTargetDate: periodKey, periodLabelTH: labels.th, periodLabelEN: labels.en,
    numbers: list,
    lastRun: PropertiesService.getDocumentProperties().getProperty('buzzLastRun') || '',
  };
}

// Dream Log numbers grouped by period (for the "did dream numbers hit?" score).
function readDreamNumbersByPeriod_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.DREAM_LOG);
  const byPeriod = {};
  if (!sheet || sheet.getLastRow() < 2) return byPeriod;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(r => {
    const p = toDateStr_(r[1]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(p)) return;
    (byPeriod[p] = byPeriod[p] || []).push.apply(byPeriod[p], String(r[4]).match(/\d{2,6}/g) || []);
  });
  return byPeriod;
}

// Scoreboard for Buzz + Dream numbers vs real draws.
function computeSourceHits_() {
  const draws = loadDraws_();
  const buzz = readSignalRowsByPeriod_();
  const dreams = readDreamNumbersByPeriod_();
  const result = {};
  [['buzz', p => (buzz[p] || []).reduce((a, r) => a.concat(r.numbers), [])],
   ['dream', p => dreams[p] || []]].forEach(([key, numbersFor]) => {
    let periods = 0, exact = 0, flip = 0;
    const recent = [];
    draws.forEach(d => {
      const nums = numbersFor(d.slot);
      if (!nums.length) return;
      periods++;
      const real = {
        jackpot: d.values.jackpot[0] || '', front3A: d.values.front3[0] || '', front3B: d.values.front3[1] || '',
        back3A: d.values.back3[0] || '', back3B: d.values.back3[1] || '', back2: d.values.back2[0] || '',
      };
      numberHitsAgainstDraw_(nums, real).forEach(h => {
        if (h.kind === 'exact') exact++; else flip++;
        if (recent.length < 10) recent.push(Object.assign({ date: d.date }, h));
      });
    });
    result[key] = { periods, exact, flip, recent };
  });
  return result;
}

// ==========================================================================
// MATH BACKTEST — fair "what would each math column have shown BEFORE that
// draw?" record, so the Every-Draw column and the Same-Date column get their
// own Aggregate Stats and Hall of Fame entries next to the Lucky Draw.
//
// For every real draw D: build the model from draws strictly OLDER than D
// only (no peeking), take that column's top-5 picks per prize, compare with
// D's real result. One row per draw × column × prize in the 'Math Backtest'
// tab (pre-calculated, so page loads only READ it — never recompute).
//
// Run buildMathBacktest() once from the editor; afterwards the daily
// trigger adds each new draw automatically right after it is archived.
// ==========================================================================

const MATH_BACKTEST_HEADERS_ = ['Draw Date', 'Column', 'Prize', 'Top 5 Picks', 'Real', '#1 Digit Matches',
  'Exact Hit', 'Exact Rank', 'Flip Hit', 'Logic'];
const MATH_BACKTEST_LOGIC_ = 'bt1';
const MATH_BACKTEST_MIN_HISTORY_ = 24; // need ~1 year of earlier draws before a backtest row is fair
const MATH_COLUMNS_ = ['all', 'slot'];

function realOfDraw_(d) {
  return {
    jackpot: d.values.jackpot[0] || '', front3A: d.values.front3[0] || '', front3B: d.values.front3[1] || '',
    back3A: d.values.back3[0] || '', back3B: d.values.back3[1] || '', back2: d.values.back2[0] || '',
  };
}

function realsForSet_(setKey, real) {
  if (setKey === 'jackpot') return [real.jackpot];
  if (setKey === 'front3') return [real.front3A, real.front3B];
  if (setKey === 'back3') return [real.back3A, real.back3B];
  return [real.back2];
}

// Compare one column's top-5 for one prize with the real numbers.
function scoreBacktestPicks_(setKey, picks, reals) {
  const top1 = picks[0] || '';
  let bestPos = [];
  reals.forEach(r => {
    const arr = positionMatchArray_(top1, r);
    if (arr.filter(Boolean).length > bestPos.filter(Boolean).length || !bestPos.length) bestPos = arr;
  });
  let exact = '', exactRank = '', flip = '';
  picks.forEach((p, i) => {
    if (!exact && reals.indexOf(p) !== -1) { exact = p; exactRank = i + 1; }
  });
  if (!exact && setKey !== 'jackpot') {
    picks.some(p => {
      const r = reals.find(x => isFlipOf_(p, x));
      if (r) { flip = p + '→' + r; return true; }
      return false;
    });
  }
  return { posMatches: bestPos.map(b => (b ? '1' : '0')).join(''), exact, exactRank, flip };
}

// Builds missing backtest rows. timeBudgetMs keeps one run inside Apps
// Script's limit; run again (or let the trigger run) to continue.
function updateMathBacktest_(timeBudgetMs) {
  const started = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TABS.MATH_BACKTEST);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.MATH_BACKTEST);
    sheet.getRange(1, 1, 1, MATH_BACKTEST_HEADERS_.length).setValues([MATH_BACKTEST_HEADERS_]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  const done = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues().forEach(r => {
      if (String(r[9]) === MATH_BACKTEST_LOGIC_) done[toDateStr_(r[0]) + '|' + r[1] + '|' + r[2]] = true;
    });
  }

  const draws = loadDraws_();              // newest first
  const settings = readMathSettings_();
  const weights = { historical: 1 };       // history only — old buzz is not available per draw
  const chrono = draws.slice().reverse();  // oldest first
  const rows = [];
  let processed = 0, stoppedEarly = false;

  for (let i = MATH_BACKTEST_MIN_HISTORY_; i < chrono.length; i++) {
    const d = chrono[i];
    const need = [];
    MATH_COLUMNS_.forEach(col => Object.keys(SET_DEFS).forEach(k => {
      if (!done[d.date + '|' + col + '|' + k]) need.push([col, k]);
    }));
    if (!need.length) continue;
    if (Date.now() - started > timeBudgetMs) { stoppedEarly = true; break; }

    const prior = chrono.slice(0, i).reverse(); // strictly older draws, newest first
    const real = realOfDraw_(d);
    need.forEach(([col, k]) => {
      const model = buildSetModel_(k, prior, settings, weights, {}, col === 'slot' ? d.slot.slice(5) : undefined);
      const picks = computeTopPicks_(model, 5).map(p => p.number);
      const reals = realsForSet_(k, real).filter(Boolean);
      const sc = scoreBacktestPicks_(k, picks, reals);
      rows.push([d.date, col, k, picks.join(' '), reals.join(' '), sc.posMatches, sc.exact, sc.exactRank, sc.flip, MATH_BACKTEST_LOGIC_]);
    });
    processed++;
  }

  if (rows.length) {
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, rows.length, 10).setNumberFormat('@');
    sheet.getRange(start, 1, rows.length, 10).setValues(rows);
  }
  Logger.log('Math Backtest: ' + processed + ' draw(s) added (' + rows.length + ' rows)' +
    (stoppedEarly ? ' — time budget reached, run buildMathBacktest() again to continue.' : ' — up to date.'));
  return { processed, rows: rows.length, stoppedEarly };
}

// Run ONCE from the editor after pasting this version (takes ~1-3 min).
// Safe to re-run: only missing draws are added.
function buildMathBacktest() {
  return updateMathBacktest_(4.5 * 60 * 1000);
}

// Aggregate stats + Hall of Fame entries per math column, read from the tab.
function readMathBacktestStats_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.MATH_BACKTEST);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const lens = { jackpot: 6, front3: 3, back3: 3, back2: 2 };
  const out = {};
  MATH_COLUMNS_.forEach(col => {
    out[col] = {
      draws: {}, exact: { jackpot: 0, front3: 0, back3: 0, back2: 0 }, flip: { front3: 0, back3: 0, back2: 0 },
      pos: { jackpot: new Array(6).fill(0), front3: new Array(3).fill(0), back3: new Array(3).fill(0), back2: new Array(2).fill(0) },
      posN: { jackpot: 0, front3: 0, back3: 0, back2: 0 }, fame: [], latest: null,
    };
  });
  data.forEach(r => {
    const col = String(r[1]), k = String(r[2]);
    const s = out[col];
    if (!s || !(k in lens)) return;
    const date = toDateStr_(r[0]);
    s.draws[date] = true;
    const pos = String(r[5]);
    if (pos.length === lens[k]) {
      s.posN[k]++;
      pos.split('').forEach((c, i) => { if (c === '1') s.pos[k][i]++; });
    }
    const exact = String(r[6] || '').trim();
    const flip = String(r[8] || '').trim();
    if (exact) {
      s.exact[k]++;
      s.fame.push({ date, category: k, number: exact.padStart(lens[k], '0'), rank: Number(r[7]) || null, source: col });
    } else if (flip && k in s.flip) {
      s.flip[k]++;
    }
    if (!s.latest || date > s.latest.date) s.latest = { date };
    if (s.latest.date === date) s.latest[k] = { picks: String(r[3]).split(' '), real: String(r[4]).split(' '), exact, flip };
  });
  const res = {};
  MATH_COLUMNS_.forEach(col => {
    const s = out[col];
    const n = Object.keys(s.draws).length;
    const rate = k => s.pos[k].map(c => (s.posN[k] ? Math.round(c / s.posN[k] * 100) : 0));
    res[col] = {
      totalCompared: n,
      exactHitsJackpot: s.exact.jackpot, exactHitsFront3: s.exact.front3, exactHitsBack3: s.exact.back3, exactHitsBack2: s.exact.back2,
      flipHitsFront3: s.flip.front3, flipHitsBack3: s.flip.back3, flipHitsBack2: s.flip.back2,
      posMatchRates: { jackpot: rate('jackpot'), front3: rate('front3'), back3: rate('back3'), back2: rate('back2') },
      // What pure luck would give with 5 picks per prize over the same draws.
      expectedExact: {
        jackpot: +(n * 5 / 1e6).toFixed(3), front3: +(n * (1 - Math.pow(1 - 2 / 1000, 5))).toFixed(1),
        back3: +(n * (1 - Math.pow(1 - 2 / 1000, 5))).toFixed(1), back2: +(n * 5 / 100).toFixed(1),
      },
      latest: s.latest,
      hallOfFame: s.fame,
    };
  });
  return res;
}

// ==========================================================================
// BEST-MATH ENGINE — picks the settings for the main numbers by HONEST
// testing on every past draw, and reports whether any real edge exists.
//
// 1. Candidates: recency half-life × smoothing × column (every draw / same
//    date) × sharpness — 200 combos, INCLUDING sharpness 0 = pure random.
// 2. Score = log score (a strictly proper scoring rule): for every real
//    draw, how much probability did the settings give to the digits that
//    actually came out, versus the fair 10%? Uses all 20 digits of every
//    draw (6 jackpot + 2×3 front + 2×3 back + 2 back), not rare exact hits.
// 3. Walk-forward: a draw is only ever predicted from EARLIER draws.
// 4. Online selection: before each draw, the engine picks whichever combo
//    had the best score SO FAR and is graded on that draw — this is the
//    true out-of-sample result of the whole method, with a significance
//    test against pure luck.
// 5. The overall winner drives the wheel's main numbers.
//
// Fast: running decayed counts are updated draw by draw (no model rebuild
// per draw), so all 200 combos are graded in one pass over the history.
// ==========================================================================

const TUNE_HALFLIVES_ = [0, 24, 48, 96, 192];   // draws (0 = all draws equal)
const TUNE_PRIORS_ = [10, 30, 100, 300];        // smoothing strength
const TUNE_COLUMNS_ = ['all', 'slot'];
const TUNE_SHARPNESS_ = [0, 0.5, 1, 1.5, 2];    // 0 = pure random
const TUNE_MIN_HISTORY_ = 24;
const TUNE_SETS_ = [['jackpot', 6], ['front3', 3], ['back3', 3], ['back2', 2]];

function tuneComboList_() {
  const list = [];
  TUNE_HALFLIVES_.forEach(h => TUNE_PRIORS_.forEach(a => TUNE_COLUMNS_.forEach(col => TUNE_SHARPNESS_.forEach(s => {
    list.push({ h, a, col, s });
  }))));
  return list;
}

// Digit distribution for one position from decayed counts, with the
// same hierarchical smoothing as the live model (slot -> all -> uniform).
function tuneDigitProbs_(allC, allW, slotC, slotW, a, col) {
  const pAll = new Array(10);
  for (let d = 0; d < 10; d++) pAll[d] = (allC[d] + a * 0.1) / (allW + a);
  if (col === 'all' || !slotC) return pAll;
  const out = new Array(10);
  for (let d = 0; d < 10; d++) out[d] = (slotC[d] + a * pAll[d]) / (slotW + a);
  return out;
}

function tuneSharpLogProb_(probs, digit, s) {
  if (s === 0) return Math.log(0.1);
  let z = 0;
  for (let d = 0; d < 10; d++) z += Math.pow(probs[d], s);
  return s * Math.log(probs[digit]) - Math.log(z);
}

function normalCdf_(x) { // Abramowitz-Stegun
  const t = 1 / (1 + 0.2316419 * Math.abs(x));
  const dd = 0.3989423 * Math.exp(-x * x / 2);
  const p = dd * t * (0.3193815 + t * (-0.3565638 + t * (1.781478 + t * (-1.821256 + t * 1.330274))));
  return x > 0 ? 1 - p : p;
}

function runMathTuning_() {
  const chrono = loadDraws_().slice().reverse(); // oldest first
  const combos = tuneComboList_();
  const ln10 = Math.log(0.1);

  // Running state per half-life: counts[set][pos][digit] for 'all', and per slot.
  const state = {};
  TUNE_HALFLIVES_.forEach(h => {
    const mk = () => TUNE_SETS_.map(([, len]) => Array.from({ length: len }, () => new Array(10).fill(0)));
    state[h] = { all: mk(), allW: TUNE_SETS_.map(() => 0), slot: {}, slotW: {} };
  });

  const perDraw = combos.map(() => []); // score vs uniform per evaluated draw
  const evalDates = [];
  let digitsPerDraw = 0;

  chrono.forEach((d, t) => {
    const slotKey = d.slot.slice(5);
    // ---- 1) grade this draw with counts from EARLIER draws only ----
    if (t >= TUNE_MIN_HISTORY_) {
      evalDates.push(d.date);
      const probCache = {}; // h|a|col|set|pos -> probs
      combos.forEach((c, ci) => {
        const st = state[c.h];
        let score = 0, n = 0;
        TUNE_SETS_.forEach(([key, len], si) => {
          (d.values[key] || []).forEach(num => {
            for (let pos = 0; pos < len; pos++) {
              const ck = c.h + '|' + c.a + '|' + c.col + '|' + si + '|' + pos;
              let probs = probCache[ck];
              if (!probs) {
                const sc = st.slot[slotKey], sw = st.slotW[slotKey];
                probs = tuneDigitProbs_(st.all[si][pos], st.allW[si], sc ? sc[si][pos] : null, sw ? sw[si] : 0, c.a, c.col);
                probCache[ck] = probs;
              }
              score += tuneSharpLogProb_(probs, Number(num[pos]), c.s) - ln10;
              n++;
            }
          });
        });
        perDraw[ci].push(score);
        digitsPerDraw = n;
      });
    }
    // ---- 2) then add this draw to the running counts ----
    TUNE_HALFLIVES_.forEach(h => {
      const st = state[h];
      const decayAll = h > 0 ? Math.pow(0.5, 1 / h) : 1;
      const decaySlot = h > 0 ? Math.pow(0.5, 24 / h) : 1; // same slot recurs ~ once a year
      if (!st.slot[slotKey]) {
        st.slot[slotKey] = TUNE_SETS_.map(([, len]) => Array.from({ length: len }, () => new Array(10).fill(0)));
        st.slotW[slotKey] = TUNE_SETS_.map(() => 0);
      }
      TUNE_SETS_.forEach(([key, len], si) => {
        st.allW[si] *= decayAll;
        st.slotW[slotKey][si] *= decaySlot;
        for (let pos = 0; pos < len; pos++) for (let dg = 0; dg < 10; dg++) {
          st.all[si][pos][dg] *= decayAll;
          st.slot[slotKey][si][pos][dg] *= decaySlot;
        }
        (d.values[key] || []).forEach(num => {
          for (let pos = 0; pos < len; pos++) {
            st.all[si][pos][Number(num[pos])] += 1;
            st.slot[slotKey][si][pos][Number(num[pos])] += 1;
          }
          st.allW[si] += 1;
          st.slotW[slotKey][si] += 1;
        });
      });
    });
  });

  const nEval = evalDates.length;
  if (!nEval) return null;

  // ---- Overall leaderboard (in-sample: best over the whole history) ----
  const totals = perDraw.map(arr => arr.reduce((a, b) => a + b, 0));
  // Ties (all pure-random combos score exactly 0) resolve to pure random.
  const order = totals.map((v, i) => i).sort((a, b) => (totals[b] - totals[a]) || (combos[a].s - combos[b].s));
  const best = combos[order[0]];
  const histIdx = order.find(i => combos[i].s > 0);

  // ---- Online selection: honest out-of-sample result of the method ----
  const defaultIdx = combos.findIndex(c => c.h === 48 && c.a === 30 && c.col === 'all' && c.s === 1);
  const cum = combos.map(() => 0);
  const online = [];
  for (let t = 0; t < nEval; t++) {
    let pick = defaultIdx;
    if (t > 0) { let bv = -Infinity; cum.forEach((v, i) => { if (v > bv) { bv = v; pick = i; } }); }
    online.push(perDraw[pick][t]);
    for (let i = 0; i < combos.length; i++) cum[i] += perDraw[i][t];
  }
  const mean = online.reduce((a, b) => a + b, 0) / nEval;
  const sd = Math.sqrt(online.reduce((a, b) => a + (b - mean) * (b - mean), 0) / Math.max(1, nEval - 1));
  const z = sd > 0 ? mean / (sd / Math.sqrt(nEval)) : 0;
  const pValue = 1 - normalCdf_(z);                       // one-sided: "better than luck"
  const perDigitEdgePct = (Math.exp(mean / digitsPerDraw) - 1) * 100; // avg likelihood ratio per digit vs 10%

  let verdict;
  if (pValue < 0.01 && mean > 0) verdict = 'edge';
  else if (pValue < 0.05 && mean > 0) verdict = 'weak';
  else verdict = 'none';

  return {
    drawsTested: nEval,
    firstDate: evalDates[0], lastDate: evalDates[nEval - 1],
    best: { halflife_draws: best.h, prior_strength: best.a, column: best.col, sharpness: best.s,
      pureRandom: best.s === 0, totalVsLuck: +totals[order[0]].toFixed(3) },
    bestHistory: { halflife_draws: combos[histIdx].h, prior_strength: combos[histIdx].a, column: combos[histIdx].col,
      sharpness: combos[histIdx].s, totalVsLuck: +totals[histIdx].toFixed(3),
      perDigitEdgePct: +((Math.exp(totals[histIdx] / (nEval * digitsPerDraw)) - 1) * 100).toFixed(2) },
    leaderboard: order.slice(0, 10).map(i => ({ h: combos[i].h, a: combos[i].a, col: combos[i].col, s: combos[i].s,
      totalVsLuck: +totals[i].toFixed(3) })),
    pureRandomRank: order.findIndex(i => combos[i].s === 0) + 1,
    online: { meanPerDraw: +mean.toFixed(4), z: +z.toFixed(2), pValue: +pValue.toFixed(3),
      perDigitEdgePct: +perDigitEdgePct.toFixed(2), verdict },
    combosTested: combos.length,
    tunedAt: new Date().toISOString(),
  };
}

// Run from the editor any time; the daily trigger also re-runs it after
// each new real draw. Saves the winner (used by the wheel) and writes a
// readable report to the 'Math Tuning' tab.
function tuneMath() {
  const t0 = Date.now();
  const res = runMathTuning_();
  if (!res) { Logger.log('Math tuning: not enough draws yet.'); return null; }
  PropertiesService.getDocumentProperties().setProperty('mathTuned', JSON.stringify(res));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TABS.MATH_TUNING);
  if (!sh) sh = ss.insertSheet(TABS.MATH_TUNING);
  sh.clear();
  const v = res.online.verdict;
  const verdictText = v === 'edge' ? 'The method beats pure luck (strong evidence).'
    : v === 'weak' ? 'Slight sign of an edge — could still be luck.'
      : 'No real edge — the draws behave like pure random (as a fair lottery should).';
  const rows = [
    ['Math Tuning report', ''],
    ['Tuned at', res.tunedAt],
    ['Draws tested (walk-forward)', res.drawsTested + ' (' + res.firstDate + ' → ' + res.lastDate + ')'],
    ['Setting combos tested', res.combosTested + ' (includes pure random)'],
    ['Winner used by the wheel', res.best.pureRandom ? 'PURE RANDOM (every number equally likely)' :
      'half-life ' + res.best.halflife_draws + ' · smoothing ' + res.best.prior_strength +
      ' · column ' + res.best.column + ' · sharpness ' + res.best.sharpness],
    ['Best history-based setting', 'half-life ' + res.bestHistory.halflife_draws + ' · smoothing ' + res.bestHistory.prior_strength +
      ' · column ' + res.bestHistory.column + ' · sharpness ' + res.bestHistory.sharpness + ' → ' + res.bestHistory.perDigitEdgePct + '% per digit vs luck'],
    ['Pure random ranked', '#' + res.pureRandomRank + ' of ' + res.combosTested],
    ['Honest out-of-sample edge per digit', res.online.perDigitEdgePct + '% vs luck'],
    ['Chance this edge is just luck (p)', String(res.online.pValue)],
    ['Verdict', verdictText],
    ['', ''],
    ['Top 10 combos', 'half-life · smoothing · column · sharpness → score vs luck (log, higher = better)'],
  ].concat(res.leaderboard.map((c, i) => ['#' + (i + 1), c.h + ' · ' + c.a + ' · ' + c.col + ' · ' + c.s + ' → ' + c.totalVsLuck]));
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1).setFontWeight('bold');
  Logger.log('Math tuning done in ' + (Date.now() - t0) + ' ms. Winner: ' + JSON.stringify(res.best) +
    ' | online edge ' + res.online.perDigitEdgePct + '% per digit, p=' + res.online.pValue + ' → ' + verdictText);
  return res;
}

function readMathTuning_() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty('mathTuned') || 'null'); } catch (e) { return null; }
}

// --------------------------------------------------------------------------
// FRONTEND SUPPORT — Phase 5. doGet() serves index.html as the web app;
// getHistoricalDraws() feeds the "Actual Data" nav tab.
// --------------------------------------------------------------------------

function doGet(e) {
  // …/exec?page=master → the Master Key page (owner only — needs the setup
  // code from the editor or the current Master Key before it shows anything).
  if (e && e.parameter && e.parameter.page === 'master') {
    return HtmlService.createHtmlOutputFromFile('master')
      .setTitle('Math vs. Lottery — Master Key')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Math vs. Lottery')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Returns every real draw in Historical Draws, newest first. Read-only —
// the Actual Data tab only displays this, never edits it, keeping with the
// "real data is permanent, only appended to" rule.
function getHistoricalDraws() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.HISTORICAL);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const sourceCol = headerRow.indexOf('Source');
  const fetchedAtCol = headerRow.indexOf('Fetched At');

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const rows = data.map(row => ({
    date: row[0] instanceof Date
      ? Utilities.formatDate(row[0], Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd')
      : String(row[0]),
    jackpot: String(row[1]).padStart(6, '0'), front3A: String(row[2]).padStart(3, '0'), front3B: String(row[3]).padStart(3, '0'),
    back3A: String(row[4]).padStart(3, '0'), back3B: String(row[5]).padStart(3, '0'), back2: String(row[6]).padStart(2, '0'),
    source: sourceCol !== -1 && row[sourceCol] ? String(row[sourceCol]) : '',
    fetchedAt: fetchedAtCol !== -1 && row[fetchedAtCol]
      ? (row[fetchedAtCol] instanceof Date ? row[fetchedAtCol].toISOString() : String(row[fetchedAtCol]))
      : '',
  }));

  rows.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0)); // newest first
  return rows;
}

// Returns every past day's generated Lucky Draw set from the Daily Cache
// tab (one row per calendar day, thanks to the Daily Lock), newest first —
// powers the "Past Draws" nav tab so a day that's already been generated
// can be browsed again without re-running the pipeline.
// Called from the frontend on page load (before any press) so the "this
// wheel is for [draw date]" caption can show immediately, without needing a
// generation first.
function getCurrentPeriodInfo() {
  const periodKey = currentPeriodKey_();
  const labels = formatDrawDateLabels_(periodKey);
  return Object.assign({ periodTargetDate: periodKey, periodLabelTH: labels.th, periodLabelEN: labels.en }, lockStatus_());
}

function getDailyCacheHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.DAILY_CACHE);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastCol = Math.min(sheet.getLastColumn(), 12);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  const rows = data.map(row => ({
    date: row[0] instanceof Date
      ? Utilities.formatDate(row[0], Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd')
      : String(row[0]),
    jackpot: String(row[1]).padStart(6, '0'), front3A: String(row[2]).padStart(3, '0'), front3B: String(row[3]).padStart(3, '0'),
    back3A: String(row[4]).padStart(3, '0'), back3B: String(row[5]).padStart(3, '0'), back2: String(row[6]).padStart(2, '0'),
    confidence: row[7],
    generatedAt: row[10] instanceof Date ? row[10].toISOString() : String(row[10] || ''),
  }));

  rows.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0)); // newest first
  return rows;
}