/* ===== Fleet Operations Control Center ===== */
(function () {
  "use strict";

  var DATA = (window.FLEET_DATA && window.FLEET_DATA.sheets) ? window.FLEET_DATA.sheets : [];
  var GEN = (window.FLEET_DATA && window.FLEET_DATA.generated) || "";
  var LS_KEY = "fleet_ops_edits_v1";

  /* ---------- connection config (Google Sheets live mode) ---------- */
  var CFG = window.FLEET_CONFIG || {};
  var storedUrl = null;
  try { storedUrl = localStorage.getItem("fleet_webapp_url"); } catch (e) {}
  var CFG_URL = (CFG.webAppUrl || "").trim();      // URL set once in config.js
  var URL_FROM_CONFIG = !!CFG_URL;                 // config.js wins -> auto-connect everywhere
  var WEBAPP = CFG_URL || (storedUrl || "");
  var LIVE = !!WEBAPP;
  var BIG = CFG.bigSheetLimit || 250;
  var STALE_MS = 60000;   // skip re-fetching a sheet opened within the last 60s
  var autoOn = (function () {
    try { var a = localStorage.getItem("fleet_auto"); if (a !== null) return a === "1"; } catch (e) {}
    return (CFG.autoRefreshSeconds || 0) > 0;
  })();

  // icon per section name (keyword based)
  function iconFor(name) {
    var n = name.toLowerCase();
    if (n.indexOf("dashboard") >= 0) return "\u{1F4CA}";
    if (n.indexOf("attendance") >= 0) return "\u{1F465}";
    if (n.indexOf("call") >= 0) return "\u{1F4DE}";
    if (n.indexOf("dispatch") >= 0) return "\u{1F69A}";
    if (n.indexOf("repair") >= 0) return "\u{1F527}";
    if (n.indexOf("advance") >= 0 || n.indexOf("money") >= 0) return "\u{1F4B0}";
    if (n.indexOf("vehicle") >= 0) return "\u{1F69B}";
    if (n.indexOf("route") >= 0) return "\u{1F5FA}️";
    if (n.indexOf("profit") >= 0) return "\u{1F4C8}";
    if (n.indexOf("emergency") >= 0) return "\u{1F6A8}";
    if (n.indexOf("loading") >= 0 || n.indexOf("weight") >= 0) return "⚖️";
    if (n.indexOf("monitor") >= 0) return "\u{1F4F1}";
    if (n.indexOf("anomaly") >= 0 || n.indexOf("log") >= 0) return "⚠️";
    if (n.indexOf("history") >= 0) return "\u{1F4DC}";
    if (n.indexOf("driver") >= 0) return "\u{1F468}‍✈️";
    if (n.indexOf("reference") >= 0 || n.indexOf("quick") >= 0) return "\u{1F4D6}";
    if (n.indexOf("program") >= 0) return "\u{1F4C5}";
    return "\u{1F4C4}";
  }

  /* ---------- persistence (diff based) ---------- */
  var edits = loadEdits();
  function loadEdits() {
    try { return JSON.parse(localStorage.getItem(LS_KEY)) || {}; }
    catch (e) { return {}; }
  }
  function saveEdits() {
    try { localStorage.setItem(LS_KEY, JSON.stringify(edits)); }
    catch (e) { toast("⚠️ Could not save (storage full). Use Save File to keep changes."); }
  }
  function sheetEdit(si) {
    if (!edits[si]) edits[si] = { h: {}, cell: {}, del: {}, add: [] };
    var e = edits[si];
    if (!e.h) e.h = {}; if (!e.cell) e.cell = {}; if (!e.del) e.del = {}; if (!e.add) e.add = [];
    return e;
  }

  /* ---------- effective data builders ---------- */
  function gridLoaded(si) { return DATA[si] && DATA[si].grid && DATA[si].grid.length >= 0 && DATA[si].grid !== null; }
  function headers(si) {
    var grid = DATA[si].grid || [];
    var base = grid.length ? grid[0].slice() : [];
    if (LIVE) return base;
    var h = (edits[si] && edits[si].h) || {};
    for (var c in h) base[c] = h[c];
    return base;
  }
  // returns array of {id|gi, kind:'base'|'add'|'live', cells:[]}
  function effectiveRows(si) {
    var grid = DATA[si].grid || [];
    if (LIVE) {
      var out2 = [];
      for (var r = 1; r < grid.length; r++) out2.push({ gi: r, kind: "live", cells: grid[r].slice() });
      return out2;
    }
    var e = edits[si] || {};
    var del = e.del || {}, cell = e.cell || {}, add = e.add || [];
    var out = [];
    for (var r2 = 1; r2 < grid.length; r2++) {
      var id = r2 - 1;
      if (del[id]) continue;
      var cells = grid[r2].slice();
      var ov = cell[id];
      if (ov) for (var c2 in ov) cells[c2] = ov[c2];
      out.push({ id: id, kind: "base", cells: cells });
    }
    for (var a = 0; a < add.length; a++) {
      out.push({ id: "a" + a, kind: "add", addIdx: a, cells: add[a].slice() });
    }
    return out;
  }
  function colCount(si) {
    var n = DATA[si].cols || 0;
    if (LIVE) {
      var g = DATA[si].grid;
      if (g) g.forEach(function (r) { if (r.length > n) n = r.length; });
      return n;
    }
    var add = (edits[si] && edits[si].add) || [];
    add.forEach(function (r) { if (r.length > n) n = r.length; });
    return n;
  }
  // row count for nav/cards without needing the grid loaded
  function count(si) {
    if (LIVE) {
      var g = DATA[si].grid;
      if (g) return Math.max(0, g.length - 1);            // exact once the sheet is loaded
      return DATA[si].rows ? Math.max(0, DATA[si].rows - 1) : 0;
    }
    return effectiveRows(si).length;
  }

  /* ---------- DOM refs ---------- */
  var $ = function (id) { return document.getElementById(id); };
  var navList = $("navList"), cards = $("cards");
  var state = { si: -1, page: 1, pageSize: 50, q: "", filtered: null };

  /* ---------- build nav + cards ---------- */
  function totalRows() {
    var t = 0;
    for (var i = 0; i < DATA.length; i++) t += count(i);
    return t;
  }
  function buildNav() {
    navList.innerHTML = "";
    cards.innerHTML = "";
    DATA.forEach(function (s, i) {
      var rc = count(i);
      var item = document.createElement("div");
      item.className = "nav-item";
      item.dataset.si = i;
      item.dataset.name = s.name.toLowerCase();
      item.innerHTML = '<span class="ico">' + iconFor(s.name) + '</span><span class="nm"></span><span class="cnt">' + rc + '</span>';
      item.querySelector(".nm").textContent = s.name;
      item.addEventListener("click", function () { openSheet(i); closeSidebar(); });
      // prefetch on hover so the click feels instant
      item.addEventListener("mouseenter", function () { if (LIVE && !DATA[i].grid && !DATA[i]._loading) loadSheet(i, true); });
      navList.appendChild(item);

      var card = document.createElement("div");
      card.className = "card";
      card.innerHTML = '<div class="card-ico">' + iconFor(s.name) + '</div><h3></h3>' +
        '<div class="meta"><b>' + rc + '</b> rows · ' + colCount(i) + ' columns</div>';
      card.querySelector("h3").textContent = s.name;
      card.addEventListener("click", function () { openSheet(i); });
      cards.appendChild(card);
    });
    $("statSheets").textContent = DATA.length;
    $("statRows").textContent = totalRows().toLocaleString();
  }
  function refreshCounts() {
    var items = navList.querySelectorAll(".nav-item");
    items.forEach(function (it) {
      var i = +it.dataset.si;
      it.querySelector(".cnt").textContent = count(i);
    });
    $("statRows").textContent = totalRows().toLocaleString();
  }

  /* ---------- views ---------- */
  function showView(id) {
    document.querySelectorAll(".view").forEach(function (v) { v.classList.remove("active"); });
    $(id).classList.add("active");
    window.scrollTo(0, 0);
  }
  function goHome() {
    state.si = -1;
    showView("overview");
    navList.querySelectorAll(".nav-item").forEach(function (n) { n.classList.remove("active"); });
  }

  function openSheet(si) {
    state.si = si; state.page = 1; state.q = ""; state.filtered = null;
    $("tableSearch").value = "";
    navList.querySelectorAll(".nav-item").forEach(function (n) {
      n.classList.toggle("active", +n.dataset.si === si);
    });
    $("sheetTitle").textContent = DATA[si].name;
    $("crumbName").textContent = DATA[si].name;
    showView("sheetView");
    syncToggle();
    if (LIVE) {
      if (DATA[si].grid) {
        render();              // instant from memory/cache
        if (Date.now() - (DATA[si].fetchedAt || 0) > 10000) loadSheet(si, true);   // quick re-opens skip refetch; still fresh
      } else {
        loadSheet(si, false);  // first open -> fetch
      }
    } else {
      render();
    }
  }

  /* ---------- filtering ---------- */
  function getRows() {
    var rows = effectiveRows(state.si);
    if (state.q) {
      var q = state.q.toLowerCase();
      rows = rows.filter(function (r) {
        for (var c = 0; c < r.cells.length; c++) {
          if (r.cells[c] && String(r.cells[c]).toLowerCase().indexOf(q) >= 0) return true;
        }
        return false;
      });
    }
    return rows;
  }

  /* ---------- render dispatcher ---------- */
  var FMT_SHEETS = { "daily vehicle status": true };
  function isDash(si) { return DATA[si] && String(DATA[si].name).toLowerCase() === "dashboard"; }
  function isFmt(si) { return DATA[si] && FMT_SHEETS[String(DATA[si].name).toLowerCase()] === true; }
  var dashMode = true;   // dashboard sheets default to the dashboard view
  var fmtMode = true;    // formatted sheets default to the colored Sheet view
  function render() {
    var sv = $("sheetView");
    sv.classList.remove("dash-mode", "fmt-mode");
    if (isDash(state.si) && dashMode) { sv.classList.add("dash-mode"); renderDashboard(state.si); }
    else if (isFmt(state.si) && fmtMode && DATA[state.si].fmt) { sv.classList.add("fmt-mode"); renderFormatted(state.si); }
    else { renderTable(); }
    updateToggle();
  }
  function updateToggle() {
    var vt = $("viewToggle"), prim = $("vtPrimary");
    if (isDash(state.si)) { vt.classList.add("show"); prim.dataset.mode = "dash"; prim.innerHTML = "📊 Dashboard"; setToggleActive(dashMode ? "dash" : "table"); }
    else if (isFmt(state.si)) { vt.classList.add("show"); prim.dataset.mode = "fmt"; prim.innerHTML = "🎨 Sheet"; setToggleActive(fmtMode ? "fmt" : "table"); }
    else { vt.classList.remove("show"); }
  }
  function setToggleActive(mode) {
    $("viewToggle").querySelectorAll(".vt-btn").forEach(function (b) { b.classList.toggle("active", b.dataset.mode === mode); });
  }
  function syncToggle() { updateToggle(); }

  /* ---------- dashboard renderer (Excel-like) ---------- */
  function isNumCell(v) {
    if (v == null || v === "") return false;
    return /^-?\d[\d,]*(\.\d+)?%?$/.test(String(v).trim());
  }
  function neCols(row) {
    var out = []; for (var c = 0; c < row.length; c++) { if (row[c] != null && String(row[c]).trim() !== "" && String(row[c]).trim() !== "_") out.push(c); } return out;
  }
  function sameSet(a, b) { if (a.length !== b.length) return false; for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false; return true; }
  function allNum(cols, row) { for (var i = 0; i < cols.length; i++) if (!isNumCell(row[cols[i]])) return false; return cols.length > 0; }
  function allTxt(cols, row) { for (var i = 0; i < cols.length; i++) if (isNumCell(row[cols[i]])) return false; return cols.length > 0; }
  function sparse(cols) { return cols.length >= 2 && (cols[cols.length - 1] - cols[0]) > (cols.length - 1); }
  function fmtVal(v) {
    var s = String(v).trim();
    if (/^-?\d+\.\d+$/.test(s)) { var n = parseFloat(s); return (Math.round(n * 100) / 100).toString(); }
    if (/^-?\d{4,}$/.test(s)) return parseInt(s, 10).toLocaleString();
    return s;
  }
  function tileColor(label) {
    var n = String(label || "").toLowerCase();
    if (/pending|\blate\b|no driver|repair|emergency|halt|\bopen\b|not updated|expire/.test(n)) return "t-red";
    if (/done|complete|filled|closed|released|\bok\b|updated|on time/.test(n)) return "t-green";
    if (/loading|progress|follow|reported|parts|order|empty|sold|rent|action/.test(n)) return "t-amber";
    if (/morning|evening|total|running|unload|calls|fleet|trip|dispatch|docs|\blr\b|whatsapp|attendance|vehicle/.test(n)) return "t-blue";
    return "t-blue";
  }
  function esc2(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

  function renderDashboard(si) {
    var g = DATA[si].grid;
    var n = g.length;
    var html = [];
    var i = 0;
    while (i < n) {
      var cols = neCols(g[i]);
      if (cols.length === 0) { i++; continue; }
      // section title bar (single cell in col 0)
      if (cols.length === 1 && cols[0] === 0 && !isNumCell(g[i][0])) {
        html.push('<div class="dash-title">' + esc2(g[i][0]) + '</div>'); i++; continue;
      }
      var next = i + 1 < n ? g[i + 1] : null;
      var ncols = next ? neCols(next) : [];
      // KPI tiles: numeric row (sparse) + label row below, same columns
      if (allNum(cols, g[i]) && sparse(cols) && next && sameSet(cols, ncols) && allTxt(ncols, next)) {
        html.push('<div class="kpi-grid">');
        for (var k = 0; k < cols.length; k++) {
          var lbl = next[cols[k]], val = g[i][cols[k]];
          html.push('<div class="kpi-tile ' + tileColor(lbl) + '"><div class="kpi-val">' + esc2(fmtVal(val)) + '</div><div class="kpi-lbl">' + esc2(lbl) + '</div></div>');
        }
        html.push('</div>'); i += 2; continue;
      }
      // status strip: label row (text) + numeric row below, same columns
      if (allTxt(cols, g[i]) && next && sameSet(cols, ncols) && allNum(ncols, next)) {
        html.push('<div class="strip-grid">');
        for (var s = 0; s < cols.length; s++) {
          var sl = g[i][cols[s]], sv = next[cols[s]];
          html.push('<div class="strip-tile ' + tileColor(sl) + '"><div class="s-val">' + esc2(fmtVal(sv)) + '</div><div class="s-lbl">' + esc2(sl) + '</div></div>');
        }
        html.push('</div>'); i += 2; continue;
      }
      // table: header (text, starts at col0) + following data rows
      if (cols[0] === 0 && allTxt(cols, g[i])) {
        var header = g[i];
        var hcols = cols.slice();
        var maxc = hcols[hcols.length - 1];
        var dataRows = [];
        var j = i + 1;
        while (j < n) {
          var rc = neCols(g[j]);
          if (rc.length === 0) break;
          if (rc.length === 1 && rc[0] === 0 && !isNumCell(g[j][0])) break; // next title
          // stop if a KPI/strip pair begins
          var jn = j + 1 < n ? neCols(g[j + 1]) : [];
          if (allNum(rc, g[j]) && sparse(rc) && g[j + 1] && sameSet(rc, jn) && allTxt(jn, g[j + 1])) break;
          dataRows.push(g[j]); j++;
          if (rc[rc.length - 1] > maxc) maxc = rc[rc.length - 1];
        }
        html.push(buildDashTable(header, dataRows, maxc, i));
        i = j; continue;
      }
      // fallback: render the row as a simple strip of tiles
      html.push('<div class="strip-grid">');
      for (var f = 0; f < cols.length; f++) {
        html.push('<div class="strip-tile t-grey"><div class="s-val">' + esc2(fmtVal(g[i][cols[f]])) + '</div></div>');
      }
      html.push('</div>'); i++;
    }
    $("dashboard").innerHTML = html.join("");
    $("rowInfo").textContent = "";
  }

  function isDateStr(s) { return /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(String(s).trim()); }
  function dateToISO(s) { var m = String(s).trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/); return m ? (m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2)) : ""; }
  function isoToDisp(iso) { var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? (m[3] + "/" + m[2] + "/" + m[1]) : iso; }

  function buildDashTable(header, rows, maxc, hdrRow) {
    var pct = [];
    var out = ['<div class="dash-tablewrap"><table class="dash-table"><thead><tr>'];
    for (var c = 0; c <= maxc; c++) {
      var h = header[c] != null ? header[c] : "";
      if (String(h).indexOf("%") >= 0) pct[c] = true;
      if (hdrRow != null && isDateStr(h)) {
        out.push('<th class="dash-date-th"><input type="date" class="dash-date" data-r="' + hdrRow + '" data-c="' + c + '" value="' + dateToISO(h) + '" title="Change date"></th>');
      } else {
        out.push('<th>' + esc2(h) + '</th>');
      }
    }
    out.push('</tr></thead><tbody>');
    rows.forEach(function (r) {
      var isTotal = String(r[0] || "").trim().toUpperCase() === "TOTAL";
      out.push('<tr' + (isTotal ? ' class="total-row"' : '') + '>');
      for (var c = 0; c <= maxc; c++) {
        var v = r[c] != null ? r[c] : "";
        var num = isNumCell(v);
        if (pct[c] && num) { v = Math.round(parseFloat(v) * 100) + "%"; }
        else { v = fmtVal(v); }
        out.push('<td class="' + (num ? "num" : "") + '">' + esc2(v) + '</td>');
      }
      out.push('</tr>');
    });
    out.push('</tbody></table></div>');
    return out.join("");
  }

  // Dropdowns the sheet has as new-style "chips" that Apps Script getDataValidations() can't read.
  var CALLER_OPTS = ["All", "kishor", "Bhupesh", "Kiran", "Manshi", "Renuka", "Sahil", "Tanisha"];
  function forcedDropdowns(si, grid) {
    var map = {};
    if (String(DATA[si].name).toLowerCase() !== "daily vehicle status") return map;
    for (var r = 0; r < Math.min(grid.length, 14); r++) {
      var row = grid[r] || [];
      for (var c = 0; c < row.length; c++) {
        if (String(row[c]).toLowerCase().indexOf("caller") >= 0 && String(row[c]).indexOf(":") >= 0) {
          for (var cc = c + 1; cc < row.length && cc <= c + 4; cc++) {
            if (String(row[cc]).trim() !== "") { map[r + "_" + cc] = CALLER_OPTS; break; }
          }
        }
      }
    }
    return map;
  }

  /* ---------- formatted "Sheet" view: real colors + merges + dropdowns ---------- */
  function renderFormatted(si) {
    var grid = DATA[si].grid || [];
    var f = DATA[si].fmt || {};
    var bg = f.bg || [], fc = f.fc || [], fw = f.fw || [], merges = f.merges || [], dv = f.dv || {};
    var forced = forcedDropdowns(si, grid);   // dropdowns Apps Script can't read (new-style chips)
    var nrows = Math.min(grid.length, bg.length || grid.length);
    var ncol = colCount(si);
    // trim trailing fully-empty rows for a tidy view
    while (nrows > 0) {
      var any = false, rr0 = grid[nrows - 1] || [];
      for (var k = 0; k < ncol; k++) { if (rr0[k] != null && String(rr0[k]).trim() !== "") { any = true; break; } }
      if (any) break; nrows--;
    }
    var covered = {}, span = {};
    merges.forEach(function (m) {
      span[m.r + "_" + m.c] = { rs: m.nr, cs: m.nc };
      for (var a = 0; a < m.nr; a++) for (var b = 0; b < m.nc; b++) { if (a === 0 && b === 0) continue; covered[(m.r + a) + "_" + (m.c + b)] = true; }
    });
    var html = ['<div class="fmt-wrap"><table class="fmt-table">'];
    for (var r = 0; r < nrows; r++) {
      html.push("<tr>");
      for (var c = 0; c < ncol; c++) {
        var key = r + "_" + c;
        if (covered[key]) continue;
        var sp = span[key];
        var val = (grid[r] && grid[r][c] != null) ? grid[r][c] : "";
        var bgc = (bg[r] && bg[r][c]) || "", fcc = (fc[r] && fc[r][c]) || "", bold = (fw[r] && fw[r][c] === "bold");
        var style = "";
        if (bgc && bgc.toLowerCase() !== "#ffffff") style += "background:" + bgc + ";";
        if (fcc && fcc.toLowerCase() !== "#000000") style += "color:" + fcc + ";";
        if (bold) style += "font-weight:700;";
        var spanAttr = sp ? (" colspan='" + sp.cs + "' rowspan='" + sp.rs + "'") : "";
        var opts = dv[key] || forced[key];
        if (opts) {
          html.push("<td class='fmt-cell' data-r='" + r + "' data-c='" + c + "'" + spanAttr + " style=\"" + style + "\">" + buildSelect(c, val, opts) + "</td>");
        } else {
          html.push("<td class='fmt-cell' contenteditable='true' data-r='" + r + "' data-c='" + c + "'" + spanAttr + " style=\"" + style + "\">" + esc(val) + "</td>");
        }
      }
      html.push("</tr>");
    }
    html.push("</table></div>");
    $("formatted").innerHTML = html.join("");
    $("rowInfo").textContent = "";
  }

  function saveFmtCell(r, c, val, revert) {
    var si = state.si;
    var cur = (DATA[si].grid[r] && DATA[si].grid[r][c] != null) ? DATA[si].grid[r][c] : "";
    if (val === cur) return;
    if (!LIVE) { if (DATA[si].grid[r]) DATA[si].grid[r][c] = val; toast("Saved locally (offline)"); return; }
    if (DATA[si].locked) { if (revert) revert(cur); toast("🔒 " + DATA[si].name + " is read-only."); return; }
    setStatus("busy", "Saving…");
    API.update(DATA[si].name, r + 1, c + 1, val).then(function () {
      if (!DATA[si].grid[r]) DATA[si].grid[r] = []; DATA[si].grid[r][c] = val; setStatus("live", "Live"); toast("Saved to sheet");
    }).catch(function (err) { if (revert) revert(cur); setStatus("err", "Error"); toast("Save failed: " + err.message); });
  }

  // dashboard date-picker (e.g. E-Waybill date in a table heading) -> writes to the sheet
  $("dashboard").addEventListener("change", function (ev) {
    var inp = ev.target;
    if (!inp.classList || !inp.classList.contains("dash-date")) return;
    var r = +inp.dataset.r, c = +inp.dataset.c, iso = inp.value;
    if (!iso) return;
    var si = state.si, name = DATA[si].name;
    setStatus("busy", "Saving date…");
    API.update(name, r + 1, c + 1, iso).then(function () {
      if (DATA[si].grid[r]) DATA[si].grid[r][c] = isoToDisp(iso);
      setStatus("live", "Live"); toast("Date updated — refreshing…");
      DATA[si].fetchedAt = 0; loadSheet(si, true);   // recompute dashboard with new date
    }).catch(function (err) { setStatus("err", "Error"); toast("Couldn't update date: " + err.message); });
  });

  // formatted-view handlers
  $("formatted").addEventListener("change", function (ev) {
    var sel = ev.target; if (!sel.classList || !sel.classList.contains("cell-select")) return;
    var td = sel.closest("td"); var r = +td.dataset.r, c = +td.dataset.c; var val = sel.value;
    sel.className = "cell-select " + selClass(val);
    saveFmtCell(r, c, val, function (p) { sel.value = p; sel.className = "cell-select " + selClass(p); });
  });
  $("formatted").addEventListener("blur", function (ev) {
    var td = ev.target; if (td.tagName !== "TD" || !td.isContentEditable) return;
    var r = +td.dataset.r, c = +td.dataset.c;
    var val = td.innerText.replace(/ /g, " ").trim();
    saveFmtCell(r, c, val, function (p) { td.textContent = p; });
  }, true);
  $("formatted").addEventListener("keydown", function (ev) {
    if (ev.key === "Enter" && ev.target.isContentEditable) { ev.preventDefault(); ev.target.blur(); }
  });

  /* ---------- dropdown columns (data-validation, like the sheet) ---------- */
  var DROPDOWNS = {
    "calling sheet": {
      status: ["SOLD", "Loading", "Running", "Unloading", "Halted", "Emergency", "Empty", "RENT", "Load"],
      ontime: ["On Time", "Late", "NO Driver", "No Program", "Empty Running", "UL On Time", "Accident", "Repair", "Other"]
    }
  };
  function dropdownMap(si) {
    var cfg = DROPDOWNS[String(DATA[si].name).toLowerCase()];
    if (!cfg) return null;
    var grid = DATA[si].grid || [];
    var scan = Math.min(grid.length, 8);
    var map = {};
    var nc = colCount(si);
    for (var c = 0; c < nc; c++) {
      for (var r = 0; r < scan; r++) {
        var h = String((grid[r] && grid[r][c]) || "").trim().toLowerCase();
        if (h === "status") { map[c] = { opts: cfg.status, label: "status" }; break; }
        if (h === "on time / late" || h === "on time/late" || h === "on time / late ") { map[c] = { opts: cfg.ontime, label: "on time / late" }; break; }
      }
    }
    return map;
  }
  function selClass(val) {
    var n = String(val || "").toLowerCase().trim();
    if (n === "") return "sel-grey";
    if (n === "sold" || n === "rent") return "sel-purple";
    if (/on time|complete|running/.test(n)) return "sel-green";
    if (/late|no driver|halt|emergency|repair|accident/.test(n)) return "sel-red";
    if (/loading|empty running|empty|load|no program/.test(n)) return "sel-amber";
    if (/unload|ul /.test(n)) return "sel-blue";
    return "sel-blue";
  }
  function buildSelect(c, val, options) {
    var v = String(val == null ? "" : val);
    var html = "<select class='cell-select " + selClass(v) + "' data-c='" + c + "'>";
    html += "<option value=''" + (v === "" ? " selected" : "") + "></option>";
    var found = false;
    options.forEach(function (o) {
      var sel = v.toLowerCase() === String(o).toLowerCase();
      if (sel) found = true;
      html += "<option" + (sel ? " selected" : "") + ">" + esc(o) + "</option>";
    });
    if (!found && v !== "") html += "<option selected>" + esc(v) + "</option>";
    html += "</select>";
    return html;
  }

  /* ---------- shared cell commit (used by text edit + dropdown) ---------- */
  function commitCell(tr, c, val, revert) {
    var si = state.si;
    if (LIVE) {
      if (tr.dataset.kind !== "live") return;
      var gi = +tr.dataset.gi;
      var prev = (DATA[si].grid[gi] && DATA[si].grid[gi][c] != null) ? DATA[si].grid[gi][c] : "";
      if (val === prev) return;
      if (DATA[si].locked) { if (revert) revert(prev); toast("🔒 " + DATA[si].name + " is read-only."); return; }
      setStatus("busy", "Saving…");
      API.update(DATA[si].name, gi + 1, c + 1, val).then(function () {
        DATA[si].grid[gi][c] = val; setStatus("live", "Live"); toast("Saved to sheet");
      }).catch(function (err) { if (revert) revert(prev); setStatus("err", "Error"); toast("Save failed: " + err.message); });
      return;
    }
    var e = sheetEdit(si);
    if (tr.dataset.kind === "add") {
      var ai = +tr.dataset.add;
      while (e.add[ai].length <= c) e.add[ai].push("");
      e.add[ai][c] = val;
    } else {
      var id = +tr.dataset.id;
      var orig = (DATA[si].grid[id + 1] && DATA[si].grid[id + 1][c] != null) ? DATA[si].grid[id + 1][c] : "";
      if (val === orig) { if (e.cell[id]) { delete e.cell[id][c]; if (!Object.keys(e.cell[id]).length) delete e.cell[id]; } }
      else { if (!e.cell[id]) e.cell[id] = {}; e.cell[id][c] = val; }
    }
    saveEdits();
  }

  /* ---------- render table ---------- */
  function renderTable() {
    var si = state.si;
    var hd = headers(si);
    var ncol = colCount(si);
    var ddMap = dropdownMap(si);
    while (hd.length < ncol) hd.push("");
    var rows = getRows();
    var ps = state.pageSize;
    var pageCount = Math.max(1, Math.ceil(rows.length / ps));
    if (state.page > pageCount) state.page = pageCount;
    var start = (state.page - 1) * ps;
    var pageRows = rows.slice(start, start + ps);

    // head
    var thead = $("dataHead");
    var trh = "<tr><th class='col-act'>#</th>";
    for (var c = 0; c < ncol; c++) {
      var label = (hd[c] && String(hd[c]).trim()) ? hd[c] : ("Column " + (c + 1));
      trh += "<th data-c='" + c + "' contenteditable='true'>" + esc(label) + "</th>";
    }
    trh += "</tr>";
    thead.innerHTML = trh;

    // body
    var tbody = $("dataBody");
    if (!pageRows.length) {
      tbody.innerHTML = "<tr><td class='empty-state' colspan='" + (ncol + 1) + "'>No records found.</td></tr>";
    } else {
      var html = [];
      for (var i = 0; i < pageRows.length; i++) {
        var r = pageRows[i];
        var globalNo = start + i + 1;
        html.push("<tr data-id='" + r.id + "' data-kind='" + r.kind + "'" + (r.kind === "add" ? " data-add='" + r.addIdx + "'" : "") + (r.kind === "live" ? " data-gi='" + r.gi + "'" : "") + ">");
        html.push("<td class='col-act'><button class='del-btn' title='Delete row'>✖</button><div class='row-num'>" + globalNo + "</div></td>");
        for (var c2 = 0; c2 < ncol; c2++) {
          var val = r.cells[c2] != null ? r.cells[c2] : "";
          var dd = ddMap && ddMap[c2];
          var lv = String(val).toLowerCase();
          if (dd && lv !== "status" && lv !== "on time / late") {
            html.push("<td class='dd-cell' data-c='" + c2 + "'>" + buildSelect(c2, val, dd.opts) + "</td>");
          } else {
            html.push("<td contenteditable='true' data-c='" + c2 + "' title=" + JSON.stringify(String(val)) + ">" + esc(val) + "</td>");
          }
        }
        html.push("</tr>");
      }
      tbody.innerHTML = html.join("");
    }

    // pager + info
    $("pageCount").textContent = pageCount;
    $("pageInput").value = state.page;
    $("pageInput").max = pageCount;
    $("rowInfo").textContent = rows.length.toLocaleString() + " row" + (rows.length === 1 ? "" : "s") +
      (state.q ? " (filtered)" : "") + " · " + ncol + " cols" +
      (LIVE && DATA[si].truncated ? " · newest " + BIG + " shown (large sheet)" : "");
    $("firstPage").disabled = $("prevPage").disabled = state.page <= 1;
    $("lastPage").disabled = $("nextPage").disabled = state.page >= pageCount;
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  /* ---------- editing events (delegated) ---------- */
  $("dataBody").addEventListener("blur", function (ev) {
    var td = ev.target;
    if (td.tagName !== "TD" || !td.isContentEditable) return;
    var tr = td.parentNode;
    var c = +td.dataset.c;
    var val = td.innerText.replace(/ /g, " ").trim();
    var si = state.si;
    if (LIVE) {
      if (tr.dataset.kind !== "live") return;
      var gi = +tr.dataset.gi;
      var prev = (DATA[si].grid[gi] && DATA[si].grid[gi][c] != null) ? DATA[si].grid[gi][c] : "";
      if (val === prev) return;
      if (DATA[si].locked) { td.textContent = prev; toast("🔒 " + DATA[si].name + " is auto-calculated — edit the source sheets."); return; }
      setStatus("busy", "Saving…");
      API.update(DATA[si].name, gi + 1, c + 1, val).then(function () {
        DATA[si].grid[gi][c] = val; td.setAttribute("title", val); setStatus("live", "Live"); toast("Saved to sheet");
      }).catch(function (err) {
        td.textContent = prev; setStatus("err", "Error"); toast("Save failed: " + err.message);
      });
      return;
    }
    var e = sheetEdit(si);
    if (tr.dataset.kind === "add") {
      var ai = +tr.dataset.add;
      while (e.add[ai].length <= c) e.add[ai].push("");
      e.add[ai][c] = val;
    } else {
      var id = +tr.dataset.id;
      var orig = (DATA[si].grid[id + 1] && DATA[si].grid[id + 1][c] != null) ? DATA[si].grid[id + 1][c] : "";
      if (val === orig) { if (e.cell[id]) { delete e.cell[id][c]; if (!Object.keys(e.cell[id]).length) delete e.cell[id]; } }
      else { if (!e.cell[id]) e.cell[id] = {}; e.cell[id][c] = val; }
    }
    td.setAttribute("title", val);
    saveEdits();
  }, true);

  // header edit
  $("dataHead").addEventListener("blur", function (ev) {
    var th = ev.target;
    if (th.tagName !== "TH" || !th.isContentEditable) return;
    var c = +th.dataset.c;
    var val = th.innerText.replace(/ /g, " ").trim();
    var si = state.si;
    if (LIVE) {
      var prevH = (DATA[si].grid[0] && DATA[si].grid[0][c] != null) ? DATA[si].grid[0][c] : "";
      if (val === prevH) return;
      if (DATA[si].locked) { th.textContent = prevH || ("Column " + (c + 1)); toast("🔒 Read-only sheet."); return; }
      setStatus("busy", "Saving…");
      API.update(DATA[si].name, 1, c + 1, val).then(function () {
        if (!DATA[si].grid[0]) DATA[si].grid[0] = []; DATA[si].grid[0][c] = val; setStatus("live", "Live"); toast("Header saved");
      }).catch(function (err) { th.textContent = prevH; setStatus("err", "Error"); toast("Save failed: " + err.message); });
      return;
    }
    var e = sheetEdit(si);
    var orig = (DATA[si].grid[0] && DATA[si].grid[0][c] != null) ? DATA[si].grid[0][c] : "";
    if (val === orig) delete e.h[c]; else e.h[c] = val;
    saveEdits();
  }, true);

  // dropdown change -> save (live to sheet, offline to diff)
  $("dataBody").addEventListener("change", function (ev) {
    var sel = ev.target;
    if (!sel.classList || !sel.classList.contains("cell-select")) return;
    var tr = sel.closest("tr");
    var c = +sel.dataset.c;
    var val = sel.value;
    sel.className = "cell-select " + selClass(val);
    commitCell(tr, c, val, function (prev) { sel.value = prev; sel.className = "cell-select " + selClass(prev); });
  });

  // Enter key commits (prevents newline)
  function commitOnEnter(ev) {
    if (ev.key === "Enter") { ev.preventDefault(); ev.target.blur(); }
  }
  $("dataBody").addEventListener("keydown", commitOnEnter);
  $("dataHead").addEventListener("keydown", commitOnEnter);

  // delete row
  $("dataBody").addEventListener("click", function (ev) {
    var btn = ev.target.closest(".del-btn");
    if (!btn) return;
    var tr = btn.closest("tr");
    var si = state.si;
    if (LIVE) {
      if (DATA[si].locked) { toast("🔒 " + DATA[si].name + " is read-only."); return; }
      if (!confirm("Delete this row directly from the Google Sheet? This cannot be undone.")) return;
      var gi = +tr.dataset.gi;
      setStatus("busy", "Deleting…");
      API.del(DATA[si].name, gi + 1).then(function () {
        DATA[si].grid.splice(gi, 1);
        DATA[si].rows = Math.max(0, (DATA[si].rows || 1) - 1);
        setStatus("live", "Live"); refreshCounts(); render(); toast("Row deleted from sheet");
      }).catch(function (err) { setStatus("err", "Error"); toast("Delete failed: " + err.message); });
      return;
    }
    if (!confirm("Delete this row? You can restore via Reset, before saving the file.")) return;
    var e = sheetEdit(si);
    if (tr.dataset.kind === "add") {
      var ai = +tr.dataset.add;
      e.add.splice(ai, 1);
    } else {
      e.del[+tr.dataset.id] = 1;
    }
    saveEdits();
    refreshCounts();
    render();
    toast("Row deleted");
  });

  /* ---------- toolbar ---------- */
  $("addRowBtn").addEventListener("click", function () {
    var si = state.si;
    var ncol = colCount(si);
    if (LIVE) {
      if (DATA[si].locked) { toast("🔒 " + DATA[si].name + " is read-only."); return; }
      var blankL = []; for (var b = 0; b < ncol; b++) blankL.push("");
      setStatus("busy", "Adding…");
      API.add(DATA[si].name, blankL).then(function (res) {
        var rowNo = (res && res.row) || (DATA[si].grid.length + 1);
        while (DATA[si].grid.length < rowNo) DATA[si].grid.push([]);
        DATA[si].grid[rowNo - 1] = blankL.slice();
        DATA[si].rows = rowNo;
        setStatus("live", "Live"); refreshCounts();
        state.q = ""; $("tableSearch").value = "";
        state.page = Math.max(1, Math.ceil(count(si) / state.pageSize));
        render();
        var rs = $("dataBody").querySelectorAll("tr"); var last = rs[rs.length - 1];
        if (last) { var ft = last.querySelector("td[contenteditable]"); if (ft) ft.focus(); }
        toast("Row added to sheet — fill it in");
      }).catch(function (err) { setStatus("err", "Error"); toast("Add failed: " + err.message); });
      return;
    }
    var e = sheetEdit(si);
    var blank = []; for (var i = 0; i < ncol; i++) blank.push("");
    e.add.push(blank);
    saveEdits();
    refreshCounts();
    state.q = ""; $("tableSearch").value = "";
    var total = effectiveRows(si).length;
    state.page = Math.max(1, Math.ceil(total / state.pageSize));
    render();
    // focus first cell of new row
    var rows = $("dataBody").querySelectorAll("tr");
    var last = rows[rows.length - 1];
    if (last) { var firstTd = last.querySelector("td[contenteditable]"); if (firstTd) firstTd.focus(); }
    toast("Row added — fill it in");
  });

  var searchTimer;
  $("tableSearch").addEventListener("input", function () {
    clearTimeout(searchTimer);
    var v = this.value;
    searchTimer = setTimeout(function () {
      state.q = v.trim(); state.page = 1; render();
    }, 200);
  });

  $("pageSize").addEventListener("change", function () {
    state.pageSize = +this.value; state.page = 1; render();
  });
  $("firstPage").addEventListener("click", function () { state.page = 1; render(); });
  $("prevPage").addEventListener("click", function () { if (state.page > 1) { state.page--; render(); } });
  $("nextPage").addEventListener("click", function () { state.page++; render(); });
  $("lastPage").addEventListener("click", function () { state.page = 1e9; render(); });
  $("pageInput").addEventListener("change", function () {
    var p = parseInt(this.value, 10); if (!isNaN(p)) { state.page = Math.max(1, p); render(); }
  });

  $("exportCsvBtn").addEventListener("click", function () { exportCSV(state.si); });

  // view toggle (Dashboard / formatted Sheet vs Table)
  $("viewToggle").addEventListener("click", function (ev) {
    var b = ev.target.closest(".vt-btn"); if (!b) return;
    var m = b.dataset.mode;
    if (m === "dash") { dashMode = true; }
    else if (m === "fmt") { fmtMode = true; }
    else { if (isDash(state.si)) dashMode = false; else fmtMode = false; }
    render();
  });

  /* ---------- CSV export ---------- */
  function csvCell(v) {
    v = v == null ? "" : String(v);
    if (/[",\n\r]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
    return v;
  }
  function exportCSV(si) {
    var hd = headers(si), ncol = colCount(si);
    while (hd.length < ncol) hd.push("");
    var lines = [hd.slice(0, ncol).map(csvCell).join(",")];
    effectiveRows(si).forEach(function (r) {
      var row = [];
      for (var c = 0; c < ncol; c++) row.push(csvCell(r.cells[c]));
      lines.push(row.join(","));
    });
    var blob = new Blob(["﻿" + lines.join("\r\n")], { type: "text/csv;charset=utf-8;" });
    downloadBlob(blob, sanitize(DATA[si].name) + ".csv");
    toast("CSV exported");
  }

  /* ---------- Save File: rebuild data.js with edits baked in ---------- */
  $("saveBtn").addEventListener("click", function () {
    if (LIVE) {
      toast("You're connected live — changes already save into the Google Sheet automatically.");
      return;
    }
    if (!confirm("Download an updated 'data.js' with ALL your changes baked in?\n\nReplace the data.js in the 'operation' folder with this file to make changes permanent for everyone.")) return;
    var out = { generated: new Date().toLocaleString(), sheets: [] };
    for (var si = 0; si < DATA.length; si++) {
      var hd = headers(si);
      var ncol = colCount(si);
      while (hd.length < ncol) hd.push("");
      var grid = [hd.slice(0, ncol)];
      effectiveRows(si).forEach(function (r) {
        var row = [];
        for (var c = 0; c < ncol; c++) row.push(r.cells[c] != null ? r.cells[c] : "");
        grid.push(row);
      });
      out.sheets.push({ name: DATA[si].name, rows: grid.length - 1, cols: ncol, grid: grid });
    }
    var content = "window.FLEET_DATA = " + JSON.stringify(out) + ";";
    var blob = new Blob([content], { type: "application/javascript;charset=utf-8;" });
    downloadBlob(blob, "data.js");
    toast("✅ data.js downloaded — replace it in the operation folder");
  });

  function downloadBlob(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }
  function sanitize(s) { return String(s).replace(/[^a-z0-9]+/gi, "_"); }

  /* ---------- nav search ---------- */
  $("navSearch").addEventListener("input", function () {
    var q = this.value.toLowerCase();
    navList.querySelectorAll(".nav-item").forEach(function (it) {
      it.style.display = it.dataset.name.indexOf(q) >= 0 ? "" : "none";
    });
  });

  /* ---------- sidebar (mobile) ---------- */
  function closeSidebar() { $("sidebar").classList.remove("open"); $("overlay").classList.remove("show"); }
  $("menuBtn").addEventListener("click", function () {
    $("sidebar").classList.toggle("open"); $("overlay").classList.toggle("show");
  });
  $("overlay").addEventListener("click", closeSidebar);
  $("backHome").addEventListener("click", function (ev) { ev.preventDefault(); goHome(); });

  /* ---------- toast ---------- */
  var toastTimer;
  function toast(msg) {
    var t = $("toast"); t.textContent = msg; t.classList.add("show");
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.classList.remove("show"); }, 2600);
  }

  /* ========================================================================
   *  LIVE MODE — Google Sheets connector (JSONP, no CORS issues from file://)
   * ====================================================================== */
  var apiSeq = 0;
  function jsonp(params, baseOverride, timeoutMs) {
    var base = baseOverride || WEBAPP;
    return new Promise(function (resolve, reject) {
      if (!base) { reject(new Error("No Web App URL configured")); return; }
      var cb = "__fleetcb" + (++apiSeq);
      var s = document.createElement("script");
      var done = false;
      var timer = setTimeout(function () { finish(new Error("timeout — check internet / deployment access")); }, timeoutMs || 70000);
      function finish(err, data) {
        if (done) return; done = true;
        clearTimeout(timer);
        try { delete window[cb]; } catch (e) { window[cb] = undefined; }
        if (s.parentNode) s.parentNode.removeChild(s);
        if (err) reject(err); else resolve(data);
      }
      window[cb] = function (resp) {
        if (resp && resp.ok) finish(null, resp.data);
        else finish(new Error((resp && resp.error) || "server error"));
      };
      var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]); }).join("&");
      s.src = base + (base.indexOf("?") >= 0 ? "&" : "?") + qs + "&callback=" + cb;
      s.onerror = function () { finish(new Error("network / blocked")); };
      document.head.appendChild(s);
    });
  }
  var SRV_V = 1;        // detected Apps Script version (>=3 top-rows, >=4 colored view)
  var LATEST_V = 6;     // newest GoogleAppsScript.gs version
  var API = {
    ping: function () { return jsonp({ action: "ping" }, null, 30000); },
    list: function () { return jsonp({ action: "list" }, null, 30000); },
    meta: function () { return jsonp({ action: "meta" }); },
    sheet: function (name, limit, head, fmt) { var p = { action: "sheet", name: name }; if (limit) p.limit = limit; if (head) p.head = head; if (fmt) p.fmt = 1; return jsonp(p); },
    add: function (name, row) { return jsonp({ action: "add", sheet: name, row: JSON.stringify(row) }); },
    update: function (name, row, col, val) { return jsonp({ action: "update", sheet: name, row: row, col: col, val: val }); },
    del: function (name, row) { return jsonp({ action: "delete", sheet: name, row: row }); }
  };

  /* ---------- per-sheet cache (instant reopen + instant reload) ---------- */
  function cacheSet(name, p) {
    var s;
    try { s = JSON.stringify({ t: Date.now(), p: p }); } catch (e) { return; }
    if (s.length > 600000) return;   // too big to cache, skip
    try { localStorage.setItem("fleet_cache_" + name, s); }
    catch (e) {
      try { Object.keys(localStorage).forEach(function (k) { if (k.indexOf("fleet_cache_") === 0) localStorage.removeItem(k); }); localStorage.setItem("fleet_cache_" + name, s); } catch (e2) {}
    }
  }
  function cacheGet(name) {
    try { var s = localStorage.getItem("fleet_cache_" + name); return s ? JSON.parse(s) : null; } catch (e) { return null; }
  }

  function setStatus(kind, text) {
    var el = $("connStatus");
    el.className = "conn-chip " + (kind === "live" ? "conn-live" : kind === "busy" ? "conn-busy" : kind === "err" ? "conn-err" : "conn-off");
    el.innerHTML = "● " + (text || (kind === "live" ? "Live" : "Offline"));
  }
  function showLoader(t) { $("loaderText").textContent = t || "Loading…"; $("loader").classList.add("show"); }
  function hideLoader() { $("loader").classList.remove("show"); }

  function loadSheet(si, silent) {
    var name = DATA[si].name;
    var fmtWanted = isFmt(si) && SRV_V >= 4;   // colored view needs the fast v4 script
    var head = 0, limit = 0;
    if (SRV_V >= 3) {
      // updated script supports top-rows -> always fetch from the TOP (where data is)
      head = fmtWanted ? 80 : (DATA[si].locked ? 200 : BIG);
    } else {
      // old script: can only truncate from the bottom
      limit = BIG;
    }
    if (!silent) showLoader("Loading “" + name + "” from Google Sheets…");
    // de-dupe: if a fetch for this sheet is already in flight, reuse it
    if (DATA[si]._loading) return DATA[si]._loading;
    setStatus("busy", "Syncing…");
    var triedFmt = fmtWanted;
    function apply(d) {
      var v = (d && d.values) || [];
      DATA[si].grid = v;
      DATA[si].cols = v.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
      DATA[si].rows = d.totalRows || v.length;
      DATA[si].truncated = !!d.truncated;
      if (triedFmt && d.bg) { DATA[si].fmt = { bg: d.bg, fc: d.fc, fw: d.fw, merges: d.merges || [], dv: d.dv || {} }; }
      else { DATA[si].fmt = null; }   // no colors -> renders as a normal table
      DATA[si].fetchedAt = Date.now();
      cacheSet(name, { grid: v, cols: DATA[si].cols, rows: DATA[si].rows, truncated: DATA[si].truncated, fmt: DATA[si].fmt });
      $("genStamp").textContent = "Updated · " + new Date().toLocaleTimeString();
      hideLoader(); setStatus("live", "Live");
      refreshCounts();
      if (state.si === si) render();
    }
    var pr = API.sheet(name, limit, head, triedFmt ? 1 : 0).then(apply).catch(function (err) {
      if (triedFmt) {
        // the colored/merge/validation request failed — fall back to plain so data still shows
        triedFmt = false;
        return API.sheet(name, 0, head, 0).then(apply).catch(function (e2) {
          hideLoader(); setStatus("err", "Error"); toast("Could not load “" + name + "”: " + e2.message);
        });
      }
      hideLoader(); setStatus("err", "Error");
      toast("Could not load “" + name + "”: " + err.message);
    });
    DATA[si]._loading = pr;
    pr.then(function () { DATA[si]._loading = null; }, function () { DATA[si]._loading = null; });
    return pr;
  }

  function bannerOff() { try { return localStorage.getItem("fleet_banner_off") === "1"; } catch (e) { return false; } }
  function connectLive(cb, attempt) {
    attempt = attempt || 1;
    setStatus("busy", attempt > 1 ? "Retrying…" : "Connecting…");
    showLoader(attempt > 1 ? "Retrying connection…" : "Connecting to Google Sheets…");
    API.list().then(function (resp) {
      var list, ver = null;
      if (Array.isArray(resp)) { list = resp; }                       // old script: bare array
      else if (resp && resp.sheets) { list = resp.sheets; ver = resp.version || 1; }  // new: {version, sheets}
      else { list = resp || []; }
      function finish() {
        SRV_V = ver || 1;
        $("updateBanner").classList.toggle("show", SRV_V < 3 && !bannerOff());
        buildFromList(list, cb);
      }
      if (ver == null) { API.ping().then(function (p) { ver = (p && p.version) || 1; }, function () { ver = 1; }).then(finish); }
      else finish();
    }).catch(function (err) {
      if (attempt < 2) { setTimeout(function () { connectLive(cb, attempt + 1); }, 1500); return; }  // cold-start retry
      hideLoader(); setStatus("err", "Not connected");
      LIVE = false;
      showConnError(err && err.message);   // don't auto-download the 16MB offline file on mobile
      if (cb) cb(false);
    });
  }
  function showConnError(msg) {
    navList.innerHTML = "";
    $("statSheets").textContent = "–"; $("statRows").textContent = "–";
    cards.innerHTML =
      "<div class='conn-error'>" +
        "<div class='ce-ico'>📡</div>" +
        "<h3>Couldn't reach the Google Sheet</h3>" +
        "<p class='ce-msg'>" + esc(msg || "Network/blocked") + "</p>" +
        "<p class='ce-hint'>Check your internet, and make sure the Apps Script is deployed with <b>“Who has access: Anyone”</b>. Then tap Retry.</p>" +
        "<div class='ce-btns'><button id='retryConn' class='btn btn-accent'>↻ Retry</button>" +
        "<button id='offlineConn' class='btn btn-ghost'>Open offline snapshot</button></div>" +
      "</div>";
    var rb = document.getElementById("retryConn"); if (rb) rb.addEventListener("click", function () { showLoader("Connecting…"); connectLive(); });
    var ob = document.getElementById("offlineConn"); if (ob) ob.addEventListener("click", function () { ensureOfflineData(startOffline); });
    showView("overview");
  }
  function buildFromList(list, cb) {
    DATA = list.filter(function (s) { return !s.hidden; }).map(function (s) {
      var d = { name: s.name, rows: s.rows, cols: s.cols, locked: !!s.locked, grid: null, truncated: false };
      var c = cacheGet(s.name);   // hydrate from last session for instant display (only if recent)
      if (c && c.p && c.p.grid && (Date.now() - (c.t || 0)) < 6 * 3600000) {
        d.grid = c.p.grid; d.cols = c.p.cols || s.cols; d.truncated = c.p.truncated; d.fmt = c.p.fmt || null; d.fetchedAt = 0;
      }
      return d;
    });
    LIVE = true;
    hideLoader(); setStatus("live", "Live");
    $("genStamp").textContent = "Connected · " + new Date().toLocaleTimeString();
    buildNav(); goHome(); startPoll();
    if (cb) cb(true);
  }

  // Load the heavy offline data.js only when actually needed (keeps live mode fast).
  function ensureOfflineData(cb) {
    if (window.FLEET_DATA) { cb(); return; }
    showLoader("Loading offline data…");
    var s = document.createElement("script");
    s.src = "data.js";
    s.onload = function () { hideLoader(); cb(); };
    s.onerror = function () { hideLoader(); cb(); };
    document.head.appendChild(s);
  }

  function startOffline() {
    LIVE = false;
    DATA = (window.FLEET_DATA && window.FLEET_DATA.sheets) ? window.FLEET_DATA.sheets : [];
    GEN = (window.FLEET_DATA && window.FLEET_DATA.generated) || GEN || "";
    setStatus("off", "Offline");
    $("genStamp").textContent = GEN ? ("Data as of " + GEN) : "";
    if (!DATA.length) {
      cards.innerHTML = "<div class='empty-state'>No data. Connect to Google Sheets (gear) or add <b>data.js</b>.</div>";
    } else { buildNav(); }
    goHome();
  }

  /* ---------- auto-refresh polling ---------- */
  var pollTimer = null;
  function startPoll() {
    stopPoll();
    var secs = autoOn ? (CFG.autoRefreshSeconds || 30) : 0;
    if (!LIVE || !secs) return;
    pollTimer = setInterval(pollCheck, secs * 1000);
  }
  function stopPoll() { if (pollTimer) { clearInterval(pollTimer); pollTimer = null; } }
  // Auto-refresh: every interval, fully re-fetch the OPEN section so edits made on
  // other devices show up. Skips the tick if the tab is hidden or you're mid-edit
  // (so it never interrupts typing / an open dropdown / the search box).
  function pollCheck() {
    if (!LIVE || state.si < 0 || document.hidden) return;
    var ae = document.activeElement;
    if (ae && (ae.isContentEditable || ae.tagName === "SELECT" || ae.tagName === "INPUT")) return;
    loadSheet(state.si, true);
  }

  /* ---------- Refresh button ---------- */
  $("refreshBtn").addEventListener("click", function () {
    if (!LIVE) {
      if (WEBAPP) { connectLive(function (ok) { if (ok) toast("Connected & refreshed"); }); }   // retry using config URL
      else { openSettings("Set webAppUrl in config.js (one time) to enable live data."); }
      return;
    }
    if (state.si >= 0) { loadSheet(state.si, false).then(function () { toast("Refreshed from Google Sheets"); }); }
    else { connectLive(function () { toast("All sections refreshed"); }); }
  });

  $("bannerClose").addEventListener("click", function () {
    $("updateBanner").classList.remove("show");
    try { localStorage.setItem("fleet_banner_off", "1"); } catch (e) {}   // stay closed for good
  });

  /* ---------- Settings dialog ---------- */
  function openSettings(msg) {
    var cfgMode = URL_FROM_CONFIG;
    $("urlInput").value = WEBAPP || "";
    $("urlInput").readOnly = cfgMode;
    $("urlInput").style.display = cfgMode ? "none" : "";
    $("connectBtn").style.display = cfgMode ? "none" : "";
    $("disconnectBtn").style.display = cfgMode ? "none" : "";
    $("settingsDesc").innerHTML = cfgMode
      ? "Connected automatically from <code>config.js</code> — no URL pasting needed. Use <b>↻ Refresh</b> any time to fetch the latest data."
      : "Paste your Apps Script Web App URL (ends with <code>/exec</code>). See <code>.tools/GoogleAppsScript.gs</code> for the one-time setup steps.";
    $("autoChk").checked = autoOn;
    $("autoSecs").textContent = (CFG.autoRefreshSeconds || 30);
    var m = $("settingsMsg"); m.className = "modal-msg";
    if (msg) { m.textContent = msg; }
    else if (LIVE) {
      m.className = "modal-msg " + (SRV_V >= LATEST_V ? "ok" : "bad");
      m.innerHTML = SRV_V >= LATEST_V
        ? "Script version " + SRV_V + " — up to date. Colored views enabled. ✓"
        : "Script version " + SRV_V + " (latest is " + LATEST_V + "). <b>Redeploy</b> the latest GoogleAppsScript.gs to enable the colored Daily Vehicle Status view.";
    } else { m.textContent = ""; }
    $("settingsModal").classList.add("show");
  }
  // auto-refresh toggle works on its own (no Connect needed in config mode)
  $("autoChk").addEventListener("change", function () {
    autoOn = this.checked;
    try { localStorage.setItem("fleet_auto", autoOn ? "1" : "0"); } catch (e) {}
    startPoll();
    toast(autoOn ? "Auto-refresh on" : "Auto-refresh off");
  });
  function closeSettings() { $("settingsModal").classList.remove("show"); }
  $("settingsBtn").addEventListener("click", function () { openSettings(); });
  $("connStatus").addEventListener("click", function () { openSettings(); });
  $("closeSettings").addEventListener("click", closeSettings);
  $("settingsModal").addEventListener("click", function (ev) { if (ev.target === this) closeSettings(); });

  $("testBtn").addEventListener("click", function () {
    var url = $("urlInput").value.trim();
    var m = $("settingsMsg"); m.className = "modal-msg"; m.textContent = "Testing…";
    if (!url) { m.className = "modal-msg bad"; m.textContent = "Enter the /exec URL first."; return; }
    jsonp({ action: "list" }, url).then(function (list) {
      m.className = "modal-msg ok"; m.textContent = "✓ Connected — found " + list.length + " sheets.";
    }).catch(function (err) { m.className = "modal-msg bad"; m.textContent = "✗ " + err.message; });
  });

  $("connectBtn").addEventListener("click", function () {
    var url = $("urlInput").value.trim();
    if (!url) { var m = $("settingsMsg"); m.className = "modal-msg bad"; m.textContent = "Enter the /exec URL."; return; }
    autoOn = $("autoChk").checked;
    try { localStorage.setItem("fleet_webapp_url", url); localStorage.setItem("fleet_auto", autoOn ? "1" : "0"); } catch (e) {}
    WEBAPP = url; closeSettings();
    connectLive();
  });

  $("disconnectBtn").addEventListener("click", function () {
    try { localStorage.setItem("fleet_webapp_url", ""); } catch (e) {}
    WEBAPP = ""; stopPoll(); closeSettings(); startOffline();
    toast("Disconnected — offline mode");
  });

  /* ---------- init ---------- */
  setStatus("off", "Offline");
  if (LIVE) { connectLive(); }
  else { ensureOfflineData(startOffline); }
})();
