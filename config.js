/* ===========================================================================
 * Fleet Operations — connection settings  (EDIT THIS ONE LINE, ONCE)
 * ---------------------------------------------------------------------------
 * Paste your Google Apps Script Web App URL between the quotes below — the link
 * that ends with /exec  (see .tools/GoogleAppsScript.gs for how to get it).
 *
 *   >>>  Set it here ONCE and you are done.  <<<
 *
 * After that, EVERY computer that opens this folder connects to the same Google
 * Sheet AUTOMATICALLY on open — no pasting anything in the website. Just press
 * the  "↻ Refresh"  button to pull the latest data whenever you want, and it
 * also auto-refreshes on its own.
 *
 * Leave it empty ("") to run OFFLINE from the bundled data.js.
 *
 * Example:
 *   webAppUrl: "https://script.google.com/macros/s/AKfycbx....../exec",
 * =========================================================================== */
window.FLEET_CONFIG = {
  webAppUrl: "https://script.google.com/macros/s/AKfycbw3qIxcZym3ibGEYQHwrksMNI_T9Q4kkMVxMvbxYfXrR5lQOJSesthgE4rEWrmLL9_h/exec",            // <-- paste your /exec URL here (one time)

  autoRefreshSeconds: 180,  // auto-refresh the open section every 3 minutes (0 = off)
  bigSheetLimit: 500        // rows loaded per big tab (smaller = faster; pager shows them)
};
