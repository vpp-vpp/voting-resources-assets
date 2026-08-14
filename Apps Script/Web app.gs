/**
 * Voting Resources — Apps Script data API
 *
 * Two GET actions:
 *   ?action=states        -> { states: ["OH", ...] }   (for the dropdown)
 *   ?state=OH              -> { state: "OH", graphics: [ {...} x12 ] }
 *
 * Deploy as a Web App (Execute as: Me, Who has access: Anyone) and call
 * with fetch() from the Squarespace Code Block. GET-only by design so no
 * CORS preflight is triggered.
 */

const STATES_DATA_SHEET_ID = "1-0RWNQKTHMx0lNfjBKCrFPDV5hjoIJU7kx_ioDbLxJI";
const SETUP_SHEET_NAME = "Setup";

// Image URLs are built as page-relative paths (no domain) rather than a
// hardcoded https://vpp-vpp.github.io/... base. voting-resources.html and
// the per-state image folders live as siblings in the same repo, so a
// relative path resolves correctly no matter which hostname served the
// page -- the plain github.io URL or the custom subdomain. A hardcoded
// cross-domain URL caused downloads to silently fail (opening the image
// instead of downloading it) for anyone loading the page from a
// different origin than the hardcoded image domain.

// CacheService.getScriptCache() max TTL is 6 hours (21600s). Every request
// currently opens the Spreadsheet and re-reads it from scratch, which is
// the main source of latency -- caching means only the first request after
// data changes (or after this TTL expires) pays that cost; everything else
// is served back near-instantly. Trade-off: sheet edits can take up to this
// long to show up live. Run clearCache() manually (from the editor's
// function dropdown) right after editing the Sheet if you want changes to
// appear immediately instead of waiting out the TTL.
const CACHE_SECONDS = 21600;

// Opt-in self-throttling against Google's "simultaneous executions per
// user" quota (30/user, shared across every visitor since this deploys
// "Execute as: Me"). Off by default -- caching already keeps real
// concurrency low in practice, so this is here as an available safety
// net rather than something needed from day one. Flip to true if traffic
// ever gets heavy enough to worry about hitting that ceiling.
const ENABLE_CONCURRENCY_LIMIT = false;
const MAX_CONCURRENT_EXECUTIONS = 20; // stay comfortably under Google's 30/user cap
const CONCURRENCY_KEY = "concurrent_executions";

function doGet(e) {
  var slotAcquired = false;
  if (ENABLE_CONCURRENCY_LIMIT) {
    slotAcquired = acquireSlot();
    if (!slotAcquired) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "High traffic right now — please try again in a moment." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  try {
    var params = (e && e.parameter) || {};
    var output;

    if ((params.action || "").toLowerCase() === "states") {
      output = getAvailableStates();
    } else {
      var state = String(params.state || "").trim().toUpperCase();
      output = state
        ? getStateGraphics(state)
        : { error: "Missing required 'state' parameter." };
    }

    return ContentService
      .createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (ENABLE_CONCURRENCY_LIMIT && slotAcquired) releaseSlot();
  }
}

// ---------------------------------------------------------------------
// Tiny LockService-backed counter, only exercised when
// ENABLE_CONCURRENCY_LIMIT is true. Each call costs a few ms uncontended
// -- negligible next to everything else doGet() does, so there's no
// performance reason to keep this disabled, only a complexity one. The
// short (30s) cache TTL on the counter is a safety net: if releaseSlot()
// is ever skipped (execution killed, times out, etc.) the counter
// self-heals in 30s instead of getting stuck rejecting requests forever.
// ---------------------------------------------------------------------
function acquireSlot() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) return true; // couldn't get the lock quickly -- fail open rather than block the request over it
  try {
    var cache = CacheService.getScriptCache();
    var current = parseInt(cache.get(CONCURRENCY_KEY) || "0", 10);
    if (current >= MAX_CONCURRENT_EXECUTIONS) return false;
    cache.put(CONCURRENCY_KEY, String(current + 1), 30);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function releaseSlot() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) return; // fail open on cleanup too -- don't hang the response over it
  try {
    var cache = CacheService.getScriptCache();
    var current = parseInt(cache.get(CONCURRENCY_KEY) || "0", 10);
    cache.put(CONCURRENCY_KEY, String(Math.max(0, current - 1)), 30);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------
// Setup tab -> list of states ready for the dropdown.
// Column A = full USPS state list (row 1 is a header, skipped).
// Column B = "TRUE" for states whose data/images are ready to go live.
// This is the *only* place that controls which states are shown — no
// separate list to maintain in code. As more states finish, just flip
// their Setup row to TRUE and they'll appear automatically.
// ---------------------------------------------------------------------
function getAvailableStates() {
  var cache = CacheService.getScriptCache();
  var cacheKey = "states";
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var ss = SpreadsheetApp.openById(STATES_DATA_SHEET_ID);
  var setup = ss.getSheetByName(SETUP_SHEET_NAME);
  var values = setup.getDataRange().getValues();
  var states = [];

  for (var i = 1; i < values.length; i++) { // row 0 is the header row
    var code = String(values[i][0] || "").trim().toUpperCase();
    var allowed = String(values[i][1] || "").trim().toUpperCase() === "TRUE";
    if (code && allowed) states.push(code);
  }

  var result = { states: states };
  cache.put(cacheKey, JSON.stringify(result), CACHE_SECONDS);
  return result;
}

// ---------------------------------------------------------------------
// One state's tab -> array of graphic objects.
// Each tab is transposed: Column A holds row labels, and each graphic
// occupies its own data column (blank spacer columns in between are
// skipped automatically, since their "name of graphic" cell is blank).
// Any row whose Column A label contains "ignore" (e.g. "Instructions
// for graphic designer (ignore)", "URL for graphic (ignore?)") is
// excluded from data extraction, per the sheet's own convention.
// ---------------------------------------------------------------------
var FIELD_MAP = [
  { match: /^name of graphic/i, field: "filenameKey" },
  { match: /^use dates/i, field: "useUntil" },
  { match: /^long text/i, field: "longText" },
  { match: /^short text/i, field: "shortText" },
  { match: /^mini-script/i, field: "miniScript" }
];

function getStateGraphics(state) {
  var cache = CacheService.getScriptCache();
  var cacheKey = "state_" + state;
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var ss = SpreadsheetApp.openById(STATES_DATA_SHEET_ID);
  var sheet = ss.getSheetByName(state);
  if (!sheet) {
    return { error: "No data found for state '" + state + "'." };
  }

  var values = sheet.getDataRange().getValues();
  var numCols = values[0].length;

  // rowFieldByIndex[rowIndex] = which JSON field that row supplies
  var rowFieldByIndex = {};
  for (var r = 0; r < values.length; r++) {
    var label = String(values[r][0] || "");
    if (/ignore/i.test(label)) continue; // explicitly excluded rows
    for (var m = 0; m < FIELD_MAP.length; m++) {
      if (FIELD_MAP[m].match.test(label)) {
        rowFieldByIndex[r] = FIELD_MAP[m].field;
        break;
      }
    }
  }

  var graphics = [];
  var id = 1;
  for (var col = 1; col < numCols; col++) { // col 0 is the label column
    var record = {};
    for (var rowIdx in rowFieldByIndex) {
      record[rowFieldByIndex[rowIdx]] = String(values[rowIdx][col] || "").trim();
    }
    if (!record.filenameKey) continue; // blank spacer column

    var filename = state + "-" + record.filenameKey + ".png";
    graphics.push({
      id: id++,
      imageUrl: state + "/" + filename, // page-relative -- see comment near the top of this file
      filename: filename,
      useUntil: record.useUntil || "",
      shortText: record.shortText || "",
      longText: record.longText || "",
      miniScript: record.miniScript || ""
    });
  }

  var result = { state: state, graphics: graphics };
  cache.put(cacheKey, JSON.stringify(result), CACHE_SECONDS);
  return result;
}

// ---------------------------------------------------------------------
// Manual cache-busting. Run this from the editor's function dropdown
// right after editing the Sheet if you want the change to show up on
// the live site immediately, instead of waiting up to CACHE_SECONDS.
// Reads the Setup tab the same way getAvailableStates() does, so every
// TRUE-flagged state gets cleared automatically -- nothing to maintain
// here as more states go live.
// ---------------------------------------------------------------------
function clearCache() {
  var cache = CacheService.getScriptCache();
  var ss = SpreadsheetApp.openById(STATES_DATA_SHEET_ID);
  var setup = ss.getSheetByName(SETUP_SHEET_NAME);
  var values = setup.getDataRange().getValues();

  var keys = ["states"];
  for (var i = 1; i < values.length; i++) { // row 0 is the header row
    var code = String(values[i][0] || "").trim().toUpperCase();
    var allowed = String(values[i][1] || "").trim().toUpperCase() === "TRUE";
    if (code && allowed) keys.push("state_" + code);
  }

  cache.removeAll(keys);
  Logger.log("Cleared: " + keys.join(", "));
}

// ---------------------------------------------------------------------
// Proactive nightly cache refresh. CacheService's max TTL is a hard
// 6-hour ceiling (Google-imposed, not configurable) -- so instead of
// relying on visitor traffic to naturally repopulate an expired cache
// during the day, this rebuilds every live state's cache entry directly,
// on a schedule, before it ever has the chance to go cold. Pair with a
// daily time-driven trigger (Triggers icon in the left sidebar -> Add
// trigger -> function: refreshCache -> Event source: Time-driven ->
// Day timer -> pick a low/no-traffic hour, e.g. 3am) and visitors should
// essentially never be the ones paying for a live Spreadsheet read.
// You can still call clearCache() by hand any time during the day if an
// urgent fix needs to go live immediately instead of waiting for the
// next scheduled refresh.
// ---------------------------------------------------------------------
function refreshCache() {
  var cache = CacheService.getScriptCache();
  cache.remove("states");
  var statesResult = getAvailableStates(); // repopulates "states"

  statesResult.states.forEach(function (code) {
    cache.remove("state_" + code);
    getStateGraphics(code); // repopulates state_<code>
  });

  Logger.log("Refreshed: states, " + statesResult.states.map(function (c) { return "state_" + c; }).join(", "));
}
