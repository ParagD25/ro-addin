/* =====================================================================================
   R&O Forecast add-in - all Excel + database logic (used by taskpane.html)
   Talks to Excel through Office.js and to Supabase through its REST API.
   ===================================================================================== */
(function (root) {
  "use strict";

  var MASTER_COLS = 25; // MasterTable columns A..Y
  var MONTHS = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];
  var RESET_FORMULAS = [
    ["D8", "=IF($D$6=\"Existing\",IF($P$2>0,INDEX(MasterTable!$C:$C,$P$2),\"\"),\"\")"],
    ["D11", "=IF($P$4,INDEX(MasterTable!$F:$F,$P$2),\"\")"],
    ["D12", "=IF($P$4,INDEX(MasterTable!$G:$G,$P$2),\"\")"],
    ["D17", "=E17"],
    ["D18", "=E18"],
    ["D19", "=E19"],
    ["D20", "=E20"],
    ["D21", "=E21"],
    ["D22", "=E22"],
    ["D23", "=E23"],
    ["D24", "=E24"],
    ["D25", "=E25"],
    ["D26", "=E26"],
    ["D27", "=E27"],
    ["D28", "=E28"],
    ["P2", "=IFERROR(MATCH($D$5&\"|\"&$D$7,MasterTable!$Y:$Y,0),0)"],
    ["P3", "=IFERROR(--$D$9,0)"],
    ["P4", "=AND($D$6=\"Existing\",$P$2>0)"],
    ["E5", "=IF($D$5=\"\",\"Choose a segment\",\"\")"],
    ["E6", "=IF($D$6=\"New\",\"You type the name and ID\",\"Values load automatically\")"],
    ["E7", "=IF($D$7=\"\",\"\",IF($D$6=\"Existing\",IF($P$2>0,\"✓ Found in master table\",\"✗ Not in \"&$D$5&\" - pick from the list\"),IF($P$2>0,\"✗ Already exists - switch to Existing\",\"✓ New project name\")))"],
    ["E8", "=IF($D$6=\"Existing\",\"Auto-filled from master table\",\"← Type the new Project ID\")"],
    ["E9", "=IF($P$3=0,\"← Pick 1-12\",IF($P$3=1,\"No months locked\",\"Jan - \"&INDEX(Lists!$E$2:$E$13,$P$3-1)&\" locked\"))"],
    ["D10", "=IF($P$3=0,\"\",EOMONTH(DATE(Settings!$C$5,$P$3,1),0))"],
    ["E11", "=IF($P$4,\"Previous: \"&TEXT(INDEX(MasterTable!$F:$F,$P$2),\"#,##0.00\"),\"Full project value\")"],
    ["E12", "=IF($P$4,\"Previous: \"&TEXT(INDEX(MasterTable!$G:$G,$P$2),\"0.0%\"),\"e.g. 32%\")"],
    ["I13", "=\"Data as of: \"&IF(ISNUMBER(Settings!$C$6),TEXT(Settings!$C$6,\"dd-mmm-yyyy hh:mm\"),Settings!$C$6)"],
    ["E14", "=IF($D$14=\"\",\"← Required: your name\",\"\")"],
    ["E17", "=IF($P$4,INDEX(MasterTable!$H:$H,$P$2),\"\")"],
    ["F17", "=IF(AND(ISNUMBER(D17),ISNUMBER(E17)),D17-E17,\"\")"],
    ["G17", "=IF($P$3=0,\"\",IF($B17<$P$3,\"Locked\",\"Open\"))"],
    ["E18", "=IF($P$4,INDEX(MasterTable!$I:$I,$P$2),\"\")"],
    ["F18", "=IF(AND(ISNUMBER(D18),ISNUMBER(E18)),D18-E18,\"\")"],
    ["G18", "=IF($P$3=0,\"\",IF($B18<$P$3,\"Locked\",\"Open\"))"],
    ["E19", "=IF($P$4,INDEX(MasterTable!$J:$J,$P$2),\"\")"],
    ["F19", "=IF(AND(ISNUMBER(D19),ISNUMBER(E19)),D19-E19,\"\")"],
    ["G19", "=IF($P$3=0,\"\",IF($B19<$P$3,\"Locked\",\"Open\"))"],
    ["E20", "=IF($P$4,INDEX(MasterTable!$K:$K,$P$2),\"\")"],
    ["F20", "=IF(AND(ISNUMBER(D20),ISNUMBER(E20)),D20-E20,\"\")"],
    ["G20", "=IF($P$3=0,\"\",IF($B20<$P$3,\"Locked\",\"Open\"))"],
    ["E21", "=IF($P$4,INDEX(MasterTable!$L:$L,$P$2),\"\")"],
    ["F21", "=IF(AND(ISNUMBER(D21),ISNUMBER(E21)),D21-E21,\"\")"],
    ["G21", "=IF($P$3=0,\"\",IF($B21<$P$3,\"Locked\",\"Open\"))"],
    ["E22", "=IF($P$4,INDEX(MasterTable!$M:$M,$P$2),\"\")"],
    ["F22", "=IF(AND(ISNUMBER(D22),ISNUMBER(E22)),D22-E22,\"\")"],
    ["G22", "=IF($P$3=0,\"\",IF($B22<$P$3,\"Locked\",\"Open\"))"],
    ["E23", "=IF($P$4,INDEX(MasterTable!$N:$N,$P$2),\"\")"],
    ["F23", "=IF(AND(ISNUMBER(D23),ISNUMBER(E23)),D23-E23,\"\")"],
    ["G23", "=IF($P$3=0,\"\",IF($B23<$P$3,\"Locked\",\"Open\"))"],
    ["E24", "=IF($P$4,INDEX(MasterTable!$O:$O,$P$2),\"\")"],
    ["F24", "=IF(AND(ISNUMBER(D24),ISNUMBER(E24)),D24-E24,\"\")"],
    ["G24", "=IF($P$3=0,\"\",IF($B24<$P$3,\"Locked\",\"Open\"))"],
    ["E25", "=IF($P$4,INDEX(MasterTable!$P:$P,$P$2),\"\")"],
    ["F25", "=IF(AND(ISNUMBER(D25),ISNUMBER(E25)),D25-E25,\"\")"],
    ["G25", "=IF($P$3=0,\"\",IF($B25<$P$3,\"Locked\",\"Open\"))"],
    ["E26", "=IF($P$4,INDEX(MasterTable!$Q:$Q,$P$2),\"\")"],
    ["F26", "=IF(AND(ISNUMBER(D26),ISNUMBER(E26)),D26-E26,\"\")"],
    ["G26", "=IF($P$3=0,\"\",IF($B26<$P$3,\"Locked\",\"Open\"))"],
    ["E27", "=IF($P$4,INDEX(MasterTable!$R:$R,$P$2),\"\")"],
    ["F27", "=IF(AND(ISNUMBER(D27),ISNUMBER(E27)),D27-E27,\"\")"],
    ["G27", "=IF($P$3=0,\"\",IF($B27<$P$3,\"Locked\",\"Open\"))"],
    ["E28", "=IF($P$4,INDEX(MasterTable!$S:$S,$P$2),\"\")"],
    ["F28", "=IF(AND(ISNUMBER(D28),ISNUMBER(E28)),D28-E28,\"\")"],
    ["G28", "=IF($P$3=0,\"\",IF($B28<$P$3,\"Locked\",\"Open\"))"],
    ["D29", "=SUM(D17:D28)"],
    ["E29", "=IF($P$4,SUM(E17:E28),\"\")"],
    ["F29", "=IF(ISNUMBER(E29),D29-E29,\"\")"]
  ];
  var UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  var SHEETS = ["Input", "MasterTable", "Settings", "Lists"];

  // ---------------------------------------------------------------- small helpers
  function cleanText(v) {
    return String(v === null || v === undefined ? "" : v).replace(/\s+/g, " ").trim();
  }
  function errorText(e) {
    if (!e) { return "Unknown error"; }
    if (e.message) { return e.message; }
    return String(e);
  }
  function nowText() {
    var d = new Date();
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return p(d.getHours()) + ":" + p(d.getMinutes());
  }
  function newToken() {
    if (root.crypto && typeof root.crypto.randomUUID === "function") { return root.crypto.randomUUID(); }
    var h = "0123456789abcdef", s = "";
    for (var i = 0; i < 32; i++) { s += h.charAt(Math.floor(Math.random() * 16)); }
    return s.substr(0, 8) + "-" + s.substr(8, 4) + "-4" + s.substr(13, 3) + "-" +
      "89ab".charAt(Math.floor(Math.random() * 4)) + s.substr(17, 3) + "-" + s.substr(20, 12);
  }
  function toExcelDate(iso) {
    var y = Number(iso.substr(0, 4)), m = Number(iso.substr(5, 2)), d = Number(iso.substr(8, 2));
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  function toRow(r) {
    var out = [];
    for (var c = 0; c < MASTER_COLS; c++) {
      var v = c < r.length ? r[c] : "";
      if (v === null || v === undefined) { v = ""; }
      if (c === 4 && typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v)) { v = toExcelDate(v); }
      out.push(v);
    }
    return out;
  }
  function result(kind, title, lines) {
    return { kind: kind, title: title, lines: lines || [] };
  }

  // ---------------------------------------------------------------- workbook access
  function getSheets(ctx) {
    var map = {};
    SHEETS.forEach(function (n) { map[n] = ctx.workbook.worksheets.getItemOrNullObject(n); });
    return map;
  }
  function assertSheets(sh) {
    SHEETS.forEach(function (n) {
      if (sh[n].isNullObject) {
        throw new Error('This workbook has no "' + n + '" sheet. Open RandO_Forecast_Input.xlsx and try again.');
      }
    });
  }
  function readCfg(values) {
    var url = cleanText(values[0][0]).replace(/\/+$/, "").replace(/\/rest\/v1$/, "");
    var key = cleanText(values[1][0]);
    if (!/^https:\/\/\S+$/.test(url) || url.indexOf("YOUR-PROJECT") >= 0) {
      throw new Error("The database address is missing. Open Connection settings below and paste your Supabase Project URL.");
    }
    if (key.length < 20 || key.indexOf("paste-your") >= 0) {
      throw new Error("The database key is missing. Open Connection settings below and paste your Supabase publishable key.");
    }
    return { url: url, key: key };
  }

  function statusColors(kind) {
    if (kind === "busy") { return ["#FFF4CE", "#5C4400"]; }
    if (kind === "ok") { return ["#E2F0D9", "#1E5631"]; }
    if (kind === "warn") { return ["#FFF2CC", "#7A4F00"]; }
    if (kind === "error") { return ["#FDE2E1", "#9B1C1C"]; }
    return ["#EEF2F7", "#1F2A1E"];
  }
  function statusText(res) {
    var icon = { ok: "✅ ", warn: "⚠ ", error: "❌ ", busy: "⏳ " }[res.kind] || "";
    return icon + res.title + (res.lines.length ? "\n• " + res.lines.join("\n• ") : "");
  }
  // mirror the panel's message into the status box on the Input sheet (queued, no sync)
  function queueStatus(inputSheet, res) {
    var c = statusColors(res.kind);
    var box = inputSheet.getRange("I7:L12");
    box.format.fill.color = c[0];
    box.format.font.color = c[1];
    inputSheet.getRange("I7").values = [[statusText(res)]];
  }
  async function writeStatusOnly(res) {
    try {
      await Excel.run(async function (ctx) {
        var ws = ctx.workbook.worksheets.getItemOrNullObject("Input");
        await ctx.sync();
        if (!ws.isNullObject) { queueStatus(ws, res); await ctx.sync(); }
      });
    } catch (e) { /* the panel still shows the message */ }
  }

  async function tryClearFilter(ctx, ws) {
    try { ws.autoFilter.clearCriteria(); await ctx.sync(); } catch (e) { /* no filter to clear */ }
  }

  // ---------------------------------------------------------------- database
  async function callRpc(cfg, fn, bodyObj) {
    var res, text;
    try {
      res = await fetch(cfg.url + "/rest/v1/rpc/" + fn, {
        method: "POST",
        headers: { "apikey": cfg.key, "Content-Type": "application/json", "Accept": "application/json" },
        body: JSON.stringify(bodyObj)
      });
      text = await res.text();
    } catch (e) {
      throw new Error("Could not reach the database. Check your internet connection and the Project URL in Connection settings.");
    }
    if (!res.ok) {
      var hint = "";
      if (res.status === 401 || res.status === 403) { hint = " The publishable key is wrong, or setup_database.sql has not been run."; }
      else if (res.status === 404) { hint = " The R&O functions were not found. Run setup_database.sql in the Supabase SQL Editor."; }
      else if (res.status >= 500) { hint = " The Supabase project may be paused. Open the Supabase dashboard and restore it."; }
      throw new Error("The database answered with error " + res.status + "." + hint);
    }
    try { return JSON.parse(text); } catch (e) { throw new Error("The database sent an unreadable answer."); }
  }

  // writes segments + master table + settings (queued; caller syncs). Returns project count.
  function queueWriteReference(sh, data, used) {
    var segs = data.segments || [];
    sh.Lists.getRange("A2:A200").clear("Contents");
    if (segs.length) {
      sh.Lists.getRangeByIndexes(1, 0, segs.length, 1).values = segs.map(function (s) { return [s]; });
    }
    var lastRow = (used && !used.isNullObject) ? used.rowIndex + used.rowCount : 1;
    if (lastRow > 1) {
      sh.MasterTable.getRangeByIndexes(1, 0, lastRow - 1, MASTER_COLS).clear("Contents");
    }
    var rows = (data.master || []).map(toRow);
    if (rows.length) {
      sh.MasterTable.getRangeByIndexes(1, 0, rows.length, MASTER_COLS).values = rows;
    }
    var fy = Number(data.config ? data.config.forecast_year : NaN);
    if (fy > 2000 && fy < 2100) { sh.Settings.getRange("C5").values = [[fy]]; }
    sh.Settings.getRange("C6").values = [[data.server_time || nowText()]];
    return rows.length;
  }

  function queueRestoreFormulas(inputSheet) {
    RESET_FORMULAS.forEach(function (pair) { inputSheet.getRange(pair[0]).formulas = [[pair[1]]]; });
  }

  // ---------------------------------------------------------------- 1. Refresh data
  async function refresh() {
    return Excel.run(async function (ctx) {
      var sh = getSheets(ctx);
      var cfgR = sh.Settings.getRange("C3:C4"); cfgR.load("values");
      var used = sh.MasterTable.getUsedRangeOrNullObject(true); used.load("rowIndex,rowCount");
      await ctx.sync();
      assertSheets(sh);
      var cfg = readCfg(cfgR.values);
      var data = await callRpc(cfg, "ro_get_reference_data", {});
      await tryClearFilter(ctx, sh.MasterTable);
      var n = queueWriteReference(sh, data, used);
      queueResetForm(sh);
      var res = result("ok", "Data refreshed at " + nowText(),
        [n + " projects across " + (data.segments || []).length + " segments loaded from the database.",
         "Form cleared. Pick a project to start."]);
      res.projects = n;
      queueStatus(sh.Input, res);
      await ctx.sync();
      return res;
    });
  }

  // ---------------------------------------------------------------- 2. Submit
  async function submit() {
    return Excel.run(async function (ctx) {
      var sh = getSheets(ctx);
      var top = sh.Input.getRange("D5:D14"); top.load("values");
      var grid = sh.Input.getRange("D17:E28"); grid.load("values");
      var look = sh.Input.getRange("P2"); look.load("values");
      var tok = sh.Lists.getRange("G2"); tok.load("values");
      var cfgR = sh.Settings.getRange("C3:C5"); cfgR.load("values");
      var used = sh.MasterTable.getUsedRangeOrNullObject(true); used.load("rowIndex,rowCount");
      await ctx.sync();
      assertSheets(sh);

      var t = top.values;
      var segment = cleanText(t[0][0]), ptype = cleanText(t[1][0]), name = cleanText(t[2][0]);
      var pid = cleanText(t[3][0]), roMonth = Number(t[4][0]);
      var totalValue = t[6][0], gm = t[7][0], comment = cleanText(t[8][0]), by = cleanText(t[9][0]);
      var months = grid.values.map(function (r) { return r[0]; });
      var previous = grid.values.map(function (r) { return r[1]; });
      var lookupRow = Number(look.values[0][0]);

      // quick checks in Excel: instant feedback, no network
      var errors = [];
      if (segment === "") { errors.push("Choose a Segment."); }
      if (ptype !== "Existing" && ptype !== "New") { errors.push("Choose a Project Type (Existing or New)."); }
      if (name === "") { errors.push("Enter or pick a Project Name."); }
      if (ptype === "Existing" && name !== "" && !(lookupRow > 0)) {
        errors.push('"' + name + '" is not a project in ' + segment + ". Pick it from the Project Name list, or switch to New.");
      }
      if (ptype === "New" && pid === "") { errors.push("Type a Project ID for the new project (or Not Booked / Sales Funnel)."); }
      if (!(roMonth >= 1 && roMonth <= 12)) { errors.push("Pick the R&O Month (1-12)."); }
      if (by === "") { errors.push("Type your name in Submitted By."); }
      if (gm === "") { errors.push("Enter the Gross Margin %."); }
      else if (typeof gm !== "number") { errors.push('Gross Margin "' + gm + '" is not a number.'); }
      else if (Math.abs(gm) > 1) { errors.push("Gross Margin looks like " + gm + ". Type it as a percentage, for example 32%."); }
      if (totalValue !== "" && typeof totalValue !== "number") { errors.push('Total Value "' + totalValue + '" is not a number.'); }
      for (var i = 0; i < 12; i++) {
        var v = months[i];
        if (v !== "" && typeof v !== "number") { errors.push(MONTHS[i] + ': "' + v + '" is not a number.'); continue; }
        var locked = roMonth >= 1 && i + 1 < roMonth;
        if (locked && ptype === "Existing" && typeof v === "number" && typeof previous[i] === "number" && Math.abs(v - previous[i]) > 0.005) {
          errors.push(MONTHS[i] + " is locked and must stay " + previous[i] + ". Press Clear form to reload it.");
        }
        if (locked && ptype === "New" && typeof v === "number" && v !== 0) {
          errors.push(MONTHS[i] + " is locked. A new project can only have values from " + MONTHS[roMonth - 1] + " onwards.");
        }
      }
      if (errors.length) {
        var bad = result("error", "Not submitted. Fix these first:", errors);
        queueStatus(sh.Input, bad);
        await ctx.sync();
        return bad;
      }

      var cfg = readCfg(cfgR.values);

      // one token per form, saved BEFORE sending: a retry after a dropped connection can never save twice
      var token = cleanText(tok.values[0][0]);
      if (!UUID_RE.test(token)) {
        token = newToken();
        sh.Lists.getRange("G2").values = [[token]];
        await ctx.sync();
      }

      var payload = {
        submission_id: token, segment: segment, project_type: ptype, project_name: name, project_id: pid,
        ro_year: Number(cfgR.values[2][0]), ro_month: roMonth, total_value: totalValue, gm_pct: gm,
        comment: comment, submitted_by: by,
        months: months.map(function (x) { return x === "" ? 0 : x; })
      };
      var res = await callRpc(cfg, "ro_submit_forecast", { payload: payload });
      if (!res.ok) {
        var rej = result("error", res.message || "Not submitted.", res.errors || []);
        queueStatus(sh.Input, rej);
        await ctx.sync();
        return rej; // form untouched so the user can correct it
      }

      // success: refresh master, reload this project as Existing
      if (res.reference) {
        await tryClearFilter(ctx, sh.MasterTable);
        queueWriteReference(sh, res.reference, used);
      }
      sh.Input.getRange("D5").values = [[res.segment || segment]];
      sh.Input.getRange("D6").values = [["Existing"]];
      sh.Input.getRange("D7").values = [[res.project_name || name]];
      sh.Input.getRange("D13").values = [[""]];
      queueRestoreFormulas(sh.Input);
      sh.Lists.getRange("G2").values = [[newToken()]];

      var warnings = res.warnings || [];
      var lines = [res.message || "Saved."];
      var out = result(warnings.length ? "warn" : "ok",
        res.status === "Duplicate" ? "Already saved at " + nowText() : "Submitted at " + nowText(),
        lines.concat(warnings.map(function (w) { return "Check: " + w; })));
      queueStatus(sh.Input, out);
      await ctx.sync();
      return out;
    });
  }

  // ---------------------------------------------------------------- 3. Clear form
  // keeps Segment, R&O Month and Submitted By; empties the project; puts EVERY formula on the form back
  // (so a pasted-over cell, e.g. in the Previous column, is repaired automatically)
  function queueResetForm(sh) {
    sh.Input.getRange("D6").values = [["Existing"]];
    sh.Input.getRange("D7").values = [[""]];
    sh.Input.getRange("D13").values = [[""]];
    queueRestoreFormulas(sh.Input);
    sh.Lists.getRange("G2").values = [[newToken()]];
    sh.Input.activate();
    sh.Input.getRange("D7").select();
  }

  async function clearForm() {
    return Excel.run(async function (ctx) {
      var sh = getSheets(ctx);
      await ctx.sync();
      assertSheets(sh);
      queueResetForm(sh);
      var res = result("info", "Form cleared", ["Pick a project name, or switch Project Type to New to add a project."]);
      queueStatus(sh.Input, res);
      await ctx.sync();
      return res;
    });
  }

  // ---------------------------------------------------------------- connection settings
  async function readSettings() {
    return Excel.run(async function (ctx) {
      var st = ctx.workbook.worksheets.getItemOrNullObject("Settings");
      await ctx.sync();
      if (st.isNullObject) { return { url: "", key: "", lastRefresh: "", missing: true }; }
      var r = st.getRange("C3:C6"); r.load("values,text");
      await ctx.sync();
      var url = cleanText(r.values[0][0]), key = cleanText(r.values[1][0]);
      return {
        url: url.indexOf("YOUR-PROJECT") >= 0 ? "" : url,
        key: key.indexOf("paste-your") >= 0 ? "" : key,
        lastRefresh: cleanText(r.text[3][0]),
        missing: false
      };
    });
  }
  async function saveSettings(url, key) {
    await Excel.run(async function (ctx) {
      var st = ctx.workbook.worksheets.getItemOrNullObject("Settings");
      await ctx.sync();
      if (st.isNullObject) { throw new Error('This workbook has no "Settings" sheet. Open RandO_Forecast_Input.xlsx.'); }
      st.getRange("C3:C4").values = [[cleanText(url)], [cleanText(key)]];
      await ctx.sync();
    });
    return refresh();
  }

  // ---------------------------------------------------------------- what is on the form right now (for the panel)
  async function readForm() {
    return Excel.run(async function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject("Input");
      await ctx.sync();
      if (ws.isNullObject) { return null; }
      var top = ws.getRange("D5:D9"); top.load("values");
      var help = ws.getRange("P2:P4"); help.load("values");
      var cells = ws.getRange("D17:D28"); cells.load("formulas");
      var head = ws.getRange("D8:D12"); head.load("formulas");
      await ctx.sync();
      var t = top.values, h = help.values, ro = Number(h[1][0]) || 0, loaded = h[2][0] === true;
      var isF = function (x) { return typeof x === "string" && x.charAt(0) === "="; };
      var typed = function (x) { return !isF(x) && x !== ""; };
      var hf = head.formulas; // D8 id, D9 month, D10 date, D11 total value, D12 GM
      var monthTyped = cells.formulas.some(function (r) { return typed(r[0]); });
      var ptype = cleanText(t[1][0]);
      var dirty = ptype === "New"
        ? (cleanText(t[2][0]) !== "" || typed(hf[0][0]) || typed(hf[3][0]) || typed(hf[4][0]) || monthTyped)
        : (loaded && (typed(hf[3][0]) || typed(hf[4][0]) || monthTyped));
      return {
        dirty: dirty,
        segment: cleanText(t[0][0]), type: cleanText(t[1][0]), name: cleanText(t[2][0]),
        id: cleanText(t[3][0]), roMonth: ro, found: Number(h[0][0]) > 0,
        months: cells.formulas.map(function (r, i) {
          var isFormula = typeof r[0] === "string" && r[0].charAt(0) === "=";
          return { locked: ro > 0 && i + 1 < ro, edited: loaded && !isFormula && !(ro > 0 && i + 1 < ro) };
        })
      };
    });
  }
  async function watchForm(callback) {
    try {
      await Excel.run(async function (ctx) {
        var ws = ctx.workbook.worksheets.getItemOrNullObject("Input");
        await ctx.sync();
        if (ws.isNullObject) { return; }
        var timer = null;
        ws.onChanged.add(function () {
          clearTimeout(timer);
          timer = setTimeout(callback, 250);
          return Promise.resolve();
        });
        await ctx.sync();
      });
    } catch (e) { /* live updates are a nice-to-have */ }
  }

  root.RO = {
    refresh: refresh, submit: submit, clearForm: clearForm, readSettings: readSettings,
    saveSettings: saveSettings, readForm: readForm, watchForm: watchForm,
    writeStatusOnly: writeStatusOnly, errorText: errorText, result: result, MONTHS: MONTHS
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = root.RO; }
})(typeof window !== "undefined" ? window : globalThis);
