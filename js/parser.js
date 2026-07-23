/* =====================================================================
 * parser.js — Reserve Advisors funding model workbook -> property object
 * ---------------------------------------------------------------------
 * Input: a SheetJS workbook (XLSX.read result). Output:
 *   { property, warnings: [...] }
 *
 * The parser anchors on labels rather than fixed addresses where it can,
 * and pushes a warning instead of guessing when an anchor is missing
 * (per build spec: "flag ambiguity rather than guessing").
 *
 * Expected sheets (v7.0 template): "Property Info", "Expenditures",
 * "Funding Plan". ("Use Instructions" and "Data" are ignored.)
 * ===================================================================== */

(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Parser = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---- SheetJS cell helpers (1-based row/col) --------------------------
  function cellAddr(r, c) {
    var col = "";
    while (c > 0) { var m = (c - 1) % 26; col = String.fromCharCode(65 + m) + col; c = (c - m - 1) / 26; }
    return col + r;
  }
  function getCell(sheet, r, c) {
    var cell = sheet[cellAddr(r, c)];
    return cell ? cell.v : null;
  }
  function getNum(sheet, r, c) {
    var v = getCell(sheet, r, c);
    return (typeof v === "number") ? v : null;
  }
  function sheetRange(sheet) {
    // "!ref" like "A1:EX2039"
    var ref = sheet["!ref"] || "A1:A1";
    var parts = ref.split(":");
    function dec(a) {
      var m = a.match(/([A-Z]+)(\d+)/);
      var c = 0;
      for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
      return { r: parseInt(m[2], 10), c: c };
    }
    return { start: dec(parts[0]), end: dec(parts[1] || parts[0]) };
  }

  // Find a cell whose string value starts with `label` (trimmed), within
  // given bounds; returns {r, c} or null.
  function findLabel(sheet, label, rMax, cMax) {
    var rng = sheetRange(sheet);
    var rEnd = Math.min(rng.end.r, rMax || rng.end.r);
    var cEnd = Math.min(rng.end.c, cMax || rng.end.c);
    var needle = label.toLowerCase();
    for (var r = 1; r <= rEnd; r++) {
      for (var c = 1; c <= cEnd; c++) {
        var v = getCell(sheet, r, c);
        if (typeof v === "string" && v.trim().toLowerCase().indexOf(needle) === 0) {
          return { r: r, c: c };
        }
      }
    }
    return null;
  }

  // Excel serial date -> ISO string (SheetJS may give number or Date)
  function toISODate(v) {
    if (v == null) return null;
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === "number") {
      var ms = Math.round((v - 25569) * 86400 * 1000);
      return new Date(ms).toISOString().slice(0, 10);
    }
    return String(v);
  }

  // ---- Property Info ---------------------------------------------------
  // Label (col A) -> field map; tolerant of row shuffling.
  var PI_FIELDS = {
    "association name": "name",
    "city": "city",
    "state": "state",
    "reference number": "refNumber",
    "length of study": "studyLength",
    "number of units": "numUnits",
    "current fiscal year": "currentFiscalYear",
    "first year of recommendation": "firstRecYear",
    "remaining budgeted intervals": "remainingPeriods",
    "remaining budgeted months of interest": "remainingInterestMonths",
    "near term inflation": "nearTermInflation",
    "last year of near term inflation": "nearTermEndYear",
    "remaining study inflation": "remainingInflation",
    "interest": "interest",
    "frequency of contributions": "totalPeriods",
    "rounded_by": "roundTo",
    "property_type": "propertyType",
    "version": "version"
  };

  function parsePropertyInfo(sheet, warnings) {
    var out = {};
    for (var r = 1; r <= 30; r++) {
      var label = getCell(sheet, r, 1);
      if (typeof label !== "string") continue;
      var key = label.trim().toLowerCase().replace(/:\s*$/, "").replace(/\s*\(years\)\s*$/, "");
      for (var k in PI_FIELDS) {
        if (key.indexOf(k) === 0) { out[PI_FIELDS[k]] = getCell(sheet, r, 2); break; }
      }
    }
    ["studyLength", "currentFiscalYear", "interest", "remainingInflation"].forEach(function (f) {
      if (out[f] == null) warnings.push("Property Info: could not find '" + f + "'.");
    });
    return out;
  }

  // ---- Expenditures ------------------------------------------------------
  function parseExpenditures(sheet, cfy, studyLength, warnings) {
    // Anchor: header row containing 'Reserve Component Inventory' (col N, row 9)
    var anchor = findLabel(sheet, "Reserve Component Inventory", 20, 30);
    if (!anchor) {
      warnings.push("Expenditures: 'Reserve Component Inventory' header not found; assuming template layout (row 9 / col N).");
      anchor = { r: 9, c: 14 };
    }
    var hdrRow = anchor.r;          // years row
    var nameCol = anchor.c;         // component name column (N)
    // First year column: first numeric == cfy in the header row right of name col
    var firstYearCol = null;
    for (var c = nameCol + 1; c <= nameCol + 30; c++) {
      if (getNum(sheet, hdrRow, c) === cfy) { firstYearCol = c; break; }
    }
    if (firstYearCol == null) {
      warnings.push("Expenditures: year columns not found; assuming col X.");
      firstYearCol = 24;
    }
    // Fixed parameter columns relative to template (B..M = 2..13, R=18, S=19)
    // Verified against v7.0; if name column moved, shift everything with it.
    var shift = nameCol - 14;
    var COL = {
      partialQty: 2 + shift, freq: 3 + shift, phaseLength: 4 + shift,
      eventsPerPhase: 5 + shift, nextRepl: 6 + shift, usefulLife: 7 + shift,
      roundPhase: 8 + shift, lineItem: 10 + shift, totalQty: 11 + shift,
      perPhaseQty: 12 + shift, units: 13 + shift, name: nameCol,
      unitCost: 18 + shift, ownership: 19 + shift
    };
    // Per-column inflation rates (2 rows above the years row)
    var colInflation = [];
    for (var k = 0; k <= studyLength; k++) {
      colInflation.push(getNum(sheet, hdrRow - 2, firstYearCol + k));
    }
    // Component rows: numeric line item in col J; category headers carry a
    // name but no line item.
    var rng = sheetRange(sheet);
    var components = [];
    var category = null;
    for (var r = hdrRow + 2; r <= rng.end.r; r++) {
      var li = getCell(sheet, r, COL.lineItem);
      var nm = getCell(sheet, r, COL.name);
      if (typeof li !== "number") {
        if (typeof nm === "string" && nm.trim() && nm.trim() !== "-") category = nm.trim();
        continue;
      }
      var comp = {
        id: "c" + r,
        row: r,
        lineItem: li,
        category: category,
        name: (typeof nm === "string") ? nm.trim() : String(nm),
        partialQty: getNum(sheet, r, COL.partialQty),
        frequencyOfEvents: getNum(sheet, r, COL.freq),
        phaseLength: getNum(sheet, r, COL.phaseLength),
        eventsPerPhase: getNum(sheet, r, COL.eventsPerPhase),
        nextReplYear: getNum(sheet, r, COL.nextRepl),
        usefulLife: getNum(sheet, r, COL.usefulLife),
        roundPhase: getNum(sheet, r, COL.roundPhase),
        totalQty: getNum(sheet, r, COL.totalQty),
        units: getCell(sheet, r, COL.units),
        unitCost: getNum(sheet, r, COL.unitCost),
        ownership: getNum(sheet, r, COL.ownership),
        deleted: false
      };
      if (comp.nextReplYear == null || comp.unitCost == null) {
        warnings.push("Expenditures row " + r + " ('" + comp.name + "'): missing timing or unit cost; skipped.");
        continue;
      }
      // Capture the workbook's cached year-cell values + formula flags so
      // Engine.deriveOverrides can detect hand-tuned cells (blanked
      // occurrences, scaled formulas, static dollar entries).
      comp.rawYearVals = [];
      comp.rawYearIsFormula = [];
      for (var kk = 0; kk <= studyLength; kk++) {
        var cellObj = sheet[cellAddr(r, firstYearCol + kk)];
        comp.rawYearVals.push(cellObj && typeof cellObj.v === "number" ? cellObj.v : null);
        comp.rawYearIsFormula.push(!!(cellObj && cellObj.f));
      }
      // Snapshot originals for reset (all fields the UI can edit)
      comp.original = {
        phaseLength: comp.phaseLength, eventsPerPhase: comp.eventsPerPhase,
        frequencyOfEvents: comp.frequencyOfEvents, nextReplYear: comp.nextReplYear,
        unitCost: comp.unitCost, usefulLife: comp.usefulLife,
        totalQty: comp.totalQty, units: comp.units, ownership: comp.ownership
      };
      components.push(comp);
    }
    if (!components.length) warnings.push("Expenditures: no component rows found.");
    return { components: components, colInflation: colInflation };
  }

  // ---- Funding Plan ------------------------------------------------------
  // Settings block labels (col V area) and the Recommended Reserve Funding
  // Table (3 column-pairs of Year/Contributions/Balances covering the
  // 30 projection years).
  function parseFundingPlan(sheet, cfy, studyLength, warnings) {
    var out = { settings: {}, recommended: [], expectedEndBalances: {} };
    var SETTING_LABELS = {
      "reserve balance ($)": "beginningBalance",
      "reserve balance date": "balanceDate",
      "reserve balance projected": "projected",
      "current year contributions ($)": "annualContribution",
      "next year contributions ($)": "nextYearContribution",
      "remaining budgeted periods": "remainingPeriods",
      "total budgeted periods": "totalPeriods",
      "round funding to ($)": "roundTo",
      "first non-budgeted year": "firstNonBudgetedYear"
    };
    var rng = sheetRange(sheet);
    for (var r = 40; r <= Math.min(rng.end.r, 120); r++) {
      for (var c = 15; c <= Math.min(rng.end.c, 30); c++) {
        var v = getCell(sheet, r, c);
        if (typeof v !== "string") continue;
        var key = v.trim().toLowerCase().replace(/:\s*$/, "");
        if (SETTING_LABELS[key] !== undefined) {
          // value sits 3 columns right (V -> Y) in the template; scan right
          for (var cc = c + 1; cc <= c + 5; cc++) {
            var val = getCell(sheet, r, cc);
            if (val != null && val !== "") { out.settings[SETTING_LABELS[key]] = val; break; }
          }
        }
      }
    }
    // Recommended funding table
    var tbl = findLabel(sheet, "Recommended Reserve Funding Table", 120, 30);
    if (!tbl) {
      warnings.push("Funding Plan: 'Recommended Reserve Funding Table' not found; recommended plan unavailable.");
      return out;
    }
    // Header row with 'Year' repeated; data starts below it.
    var hdr = null;
    for (var r2 = tbl.r + 1; r2 <= tbl.r + 6; r2++) {
      if (getCell(sheet, r2, tbl.c) === "Year") { hdr = r2; break; }
    }
    if (hdr == null) { warnings.push("Funding Plan: funding table header not found."); return out; }
    var byYear = {};
    // three column triplets: (K,L,M), (N,O,P), (Q,R,S) relative to tbl.c
    for (var t = 0; t < 3; t++) {
      var yc = tbl.c + t * 3, cc2 = yc + 1, bc = yc + 2;
      for (var r3 = hdr + 1; r3 <= hdr + 40; r3++) {
        var y = getNum(sheet, r3, yc);
        if (y == null) continue;
        byYear[y] = { contribution: getNum(sheet, r3, cc2), balance: getNum(sheet, r3, bc) };
      }
    }
    // Assemble arrays indexed by offset k
    var rec = new Array(studyLength + 1).fill(0);
    rec[0] = out.settings.annualContribution || 0; // annual current-year figure
    var missing = [];
    for (var k = 1; k <= studyLength; k++) {
      var yr = cfy + k;
      if (byYear[yr]) {
        rec[k] = byYear[yr].contribution || 0;
        out.expectedEndBalances[yr] = byYear[yr].balance;
      } else missing.push(yr);
    }
    if (missing.length) warnings.push("Funding Plan: no table entry for years " + missing.join(", ") + ".");
    out.recommended = rec;
    return out;
  }

  // ---- Top level ---------------------------------------------------------
  function parse(wb, opts) {
    opts = opts || {};
    var warnings = [];
    function sheetByName(fragment) {
      var name = wb.SheetNames.find(function (n) {
        return n.toLowerCase().indexOf(fragment.toLowerCase()) !== -1;
      });
      return name ? wb.Sheets[name] : null;
    }
    var pi = sheetByName("Property Info");
    var ex = sheetByName("Expenditures");
    var fp = sheetByName("Funding Plan");
    if (!pi || !ex || !fp) {
      throw new Error("Workbook is missing required sheet(s): " +
        [!pi && "Property Info", !ex && "Expenditures", !fp && "Funding Plan"]
          .filter(Boolean).join(", ") +
        ". Is this a Reserve Advisors funding model workbook?");
    }
    var info = parsePropertyInfo(pi, warnings);
    var cfy = info.currentFiscalYear;
    var n = info.studyLength;
    var exd = parseExpenditures(ex, cfy, n, warnings);
    var fpd = parseFundingPlan(fp, cfy, n, warnings);
    var s = fpd.settings;

    if (String(s.projected).toLowerCase() === "yes") {
      warnings.push("This workbook uses a PROJECTED reserve balance (Funding Plan setting " +
        "'Reserve Balance Projected' = Yes). The dashboard's current-year math assumes an " +
        "actual balance; current-year figures may not match the workbook. Please verify.");
    }

    var property = {
      schemaVersion: 1,
      id: opts.id || ("p_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8)),
      name: info.name || "Unnamed Association",
      city: info.city || "", state: info.state || "",
      refNumber: info.refNumber != null ? String(info.refNumber) : "",
      propertyType: info.propertyType || "",
      numUnits: info.numUnits,
      studyLength: n,
      currentFiscalYear: cfy,
      firstRecYear: info.firstRecYear || (s.firstNonBudgetedYear || cfy + 1),
      roundTo: s.roundTo || info.roundTo || 100,
      rates: {
        interest: info.interest,
        nearTermInflation: info.nearTermInflation,
        remainingInflation: info.remainingInflation,
        nearTermEndYear: info.nearTermEndYear || (cfy + n),
        inflationOverride: null
      },
      current: {
        beginningBalance: s.beginningBalance,
        balanceDate: toISODate(s.balanceDate),
        annualContribution: s.annualContribution || 0,
        remainingPeriods: s.remainingPeriods != null ? s.remainingPeriods : (info.remainingPeriods || 12),
        totalPeriods: s.totalPeriods != null ? s.totalPeriods : (info.totalPeriods || 12),
        remainingInterestMonths: info.remainingInterestMonths != null ? info.remainingInterestMonths : 12,
        projected: s.projected || "No"
      },
      colInflation: exd.colInflation,
      components: exd.components,
      fundingPlan: {
        mode: "recommended",              // recommended | manual | baseline | threshold
        thresholdValue: null,
        contributions: fpd.recommended.slice(),
        assessments: new Array(n + 1).fill(0), // manual Additional Assessments
        recommended: fpd.recommended
      },
      savedPlans: [],                      // [{id,name,contributions,assessments}]
      // For validation/tests only; harmless to persist
      _expectedEndBalances: fpd.expectedEndBalances,
      uploadedAt: new Date().toISOString()
    };
    if (property.current.beginningBalance == null) {
      warnings.push("Funding Plan: beginning reserve balance not found — defaulting to 0. Set it in the Current Year panel.");
      property.current.beginningBalance = 0;
    }
    // Study baseline snapshot for "Reset to study" (captured here so it always
    // includes the balance date and survives round-trips through storage).
    property._origAssume = {
      interest: property.rates.interest,
      balance: property.current.beginningBalance,
      balanceDate: property.current.balanceDate,
      contrib: property.current.annualContribution,
      remainingPeriods: property.current.remainingPeriods,
      remainingInterestMonths: property.current.remainingInterestMonths
    };
    return { property: property, warnings: warnings };
  }

  return { parse: parse, _cellAddr: cellAddr };
});
