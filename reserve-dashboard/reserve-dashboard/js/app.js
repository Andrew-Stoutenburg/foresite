/* =====================================================================
 * app.js — Multi-Property Reserve Study Dashboard (POC)
 * ---------------------------------------------------------------------
 * Views (hash routing):
 *   #/portfolio      — all properties, aggregate project explorer
 *   #/property/<id>  — single property detail + editing
 *
 * Dependencies (all bundled locally, no CDN):
 *   vendor/xlsx.full.min.js  (SheetJS)
 *   js/engine.js  js/parser.js  js/store.js  js/charts.js
 * ===================================================================== */
(function () {
  "use strict";

  var store = Store.create();
  var root = document.getElementById("app");

  // ---------- utilities -------------------------------------------------
  function $(sel, ctx) { return (ctx || document).querySelector(sel); }
  function $all(sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (ch) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch];
    });
  }
  function fmt$(v, dec) {
    if (v == null || isNaN(v)) return "—";
    var neg = v < 0;
    var s = "$" + Math.abs(v).toLocaleString(undefined, {
      minimumFractionDigits: dec || 0, maximumFractionDigits: dec || 0
    });
    return neg ? "−" + s : s;
  }
  function debounce(fn, ms) {
    var t; return function () {
      var args = arguments, self = this;
      clearTimeout(t); t = setTimeout(function () { fn.apply(self, args); }, ms);
    };
  }
  // ---- currency inputs ($##,###) --------------------------------------
  // Whole-dollar display for large figures (balances, contributions).
  function money(v) {
    if (v == null || isNaN(v)) return "";
    var neg = v < 0;
    return (neg ? "-$" : "$") + Math.abs(Math.round(v)).toLocaleString();
  }
  // Unit costs keep cents precision (many are like $3.50/sq yd) but still
  // group thousands; trailing ".00" is dropped for clean whole values.
  function moneyCents(v) {
    if (v == null || isNaN(v)) return "";
    var neg = v < 0;
    var s = Math.abs(v).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    s = s.replace(/\.00$/, "");
    return (neg ? "-$" : "$") + s;
  }
  function parseMoney(str) {
    if (typeof str === "number") return str;
    var n = parseFloat(String(str).replace(/[^0-9.\-]/g, ""));
    return isNaN(n) ? null : n;
  }
  // Wire a text input to behave as a currency field: reformats on blur,
  // selects-all on focus, calls onCommit(value) with the parsed number.
  function wireMoney(input, getValue, onCommit, opts) {
    opts = opts || {};
    var fmt = opts.cents ? moneyCents : money;
    input.value = fmt(getValue());
    input.addEventListener("focus", function () { input.select(); });
    input.addEventListener("change", function () {
      var v = parseMoney(input.value);
      if (v == null || v < (opts.min == null ? -Infinity : opts.min)) {
        input.value = fmt(getValue()); return;
      }
      onCommit(v);
    });
  }
  function toast(msg, isError) {
    var t = document.createElement("div");
    t.className = "toast" + (isError ? " error" : "");
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(function () { t.classList.add("show"); }, 10);
    setTimeout(function () { t.classList.remove("show"); setTimeout(function () { t.remove(); }, 300); }, isError ? 6500 : 3200);
  }
  // Normalized component key for cross-property grouping ("Asphalt Pavement, ..." -> "asphalt pavement")
  function componentKey(name) {
    return String(name || "").split(",")[0].trim().toLowerCase();
  }

  // ---------- upload ----------------------------------------------------
  function parseFiles(fileList, done) {
    var files = Array.prototype.slice.call(fileList);
    var results = [], pending = files.length;
    if (!pending) return done(results);
    files.forEach(function (file) {
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
          var res = Parser.parse(wb);
          var p = res.property;
          p.sourceFileName = file.name;
          p.overrideNotes = Engine.deriveOverrides(p);
          p.parseWarnings = res.warnings;
          results.push({ ok: true, property: p, file: file.name });
        } catch (err) {
          results.push({ ok: false, error: err.message, file: file.name });
        }
        if (--pending === 0) done(results);
      };
      reader.onerror = function () {
        results.push({ ok: false, error: "Could not read file.", file: file.name });
        if (--pending === 0) done(results);
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function handleUpload(fileList, opts) {
    parseFiles(fileList, function (results) {
      var saved = 0, failures = [];
      var chain = Promise.resolve();
      results.forEach(function (r) {
        if (!r.ok) { failures.push(r.file + ": " + r.error); return; }
        chain = chain.then(function () {
          return store.saveProperty(r.property).then(function () { saved++; });
        });
      });
      chain.then(function () {
        failures.forEach(function (f) { toast(f, true); });
        if (saved) toast(saved + (saved === 1 ? " property" : " properties") + " imported.");
        var okOnes = results.filter(function (r) { return r.ok; });
        if (opts && opts.goToProperty && okOnes.length === 1) {
          location.hash = "#/property/" + okOnes[0].property.id;
        } else {
          render();
        }
      }).catch(function (err) { toast(err.message, true); render(); });
    });
  }

  function uploadZoneHTML(id, label) {
    return '<div class="upload-zone" id="' + id + '">' +
      '<input type="file" accept=".xlsx,.xlsm" multiple hidden>' +
      '<div class="uz-icon">⬆</div><div class="uz-label">' + esc(label) + "</div>" +
      '<div class="uz-sub">Drop Reserve Advisors funding model workbooks here, or click to browse. ' +
      "Multiple files at once populate the portfolio.</div></div>";
  }
  function wireUploadZone(elId, opts) {
    var zone = document.getElementById(elId);
    if (!zone) return;
    var input = $("input[type=file]", zone);
    zone.addEventListener("click", function () { input.click(); });
    input.addEventListener("change", function () { handleUpload(input.files, opts); input.value = ""; });
    ["dragover", "dragenter"].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.add("drag"); });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.remove("drag"); });
    });
    zone.addEventListener("drop", function (e) { handleUpload(e.dataTransfer.files, opts); });
  }

  /* =====================================================================
   * PORTFOLIO VIEW
   * =================================================================== */
  var portfolioFilter = { window: "all", search: "", costMin: null, costMax: null, propertyId: "all",
    coordKey: null, coordFrom: null, coordTo: null };

  function occurrenceRows(props) {
    // Flatten: one row per component occurrence (property × component × year)
    var rows = [];
    props.forEach(function (p) {
      var em = Engine.computeExpenditures(p);
      em.perComponent.forEach(function (pc) {
        pc.amounts.forEach(function (a, k) {
          if (a > 0) rows.push({
            property: p, component: pc.component, k: k,
            year: p.currentFiscalYear + k, amount: a
          });
        });
      });
    });
    return rows;
  }

  function windowMatch(row, win) {
    var cfyK = row.year - row.property.currentFiscalYear; // == row.k
    switch (win) {
      case "this": return cfyK === 0;
      case "next": return cfyK === 1;
      case "next2": return cfyK >= 1 && cfyK <= 2;
      case "next5": return cfyK >= 1 && cfyK <= 5;
      case "pastdue": return row.component.nextReplYear < row.property.currentFiscalYear;
      default: return true;
    }
  }

  function applyPortfolioFilters(rows) {
    var f = portfolioFilter;
    var s = f.search.trim().toLowerCase();
    return rows.filter(function (r) {
      if (f.propertyId !== "all" && r.property.id !== f.propertyId) return false;
      // Coordination filter: exact component family AND within the card's year window
      if (f.coordKey) {
        if (componentKey(r.component.name) !== f.coordKey) return false;
        if (f.coordFrom != null && (r.year < f.coordFrom || r.year > f.coordTo)) return false;
      }
      if (!windowMatch(r, f.window)) return false;
      if (s && r.component.name.toLowerCase().indexOf(s) === -1 &&
          (r.component.category || "").toLowerCase().indexOf(s) === -1) return false;
      if (f.costMin != null && r.amount < f.costMin) return false;
      if (f.costMax != null && r.amount > f.costMax) return false;
      return true;
    });
  }

  function coordinationGroups(rows) {
    // Same component family occurring across >=2 properties within a 3-year span
    var byKey = {};
    rows.forEach(function (r) {
      var k = componentKey(r.component.name);
      (byKey[k] = byKey[k] || []).push(r);
    });
    var groups = [];
    Object.keys(byKey).forEach(function (key) {
      var list = byKey[key].slice().sort(function (a, b) { return a.year - b.year; });
      // sliding window: find clusters where >=2 distinct properties within 3 consecutive years
      var i = 0;
      while (i < list.length) {
        var j = i, props = {}, total = 0;
        while (j < list.length && list[j].year - list[i].year <= 2) {
          props[list[j].property.id] = list[j].property.name;
          total += list[j].amount; j++;
        }
        var ids = Object.keys(props);
        if (ids.length >= 2) {
          groups.push({
            key: key, label: list[i].component.name.split(",")[0],
            yearFrom: list[i].year, yearTo: list[j - 1].year,
            properties: ids.map(function (id) { return props[id]; }),
            total: total, count: j - i
          });
          i = j;
        } else i++;
      }
    });
    groups.sort(function (a, b) { return a.yearFrom - b.yearFrom || b.total - a.total; });
    return groups;
  }

  var WINDOW_BUTTONS = [
    ["this", "This year"], ["next", "Next year"], ["next2", "Next 2 years"],
    ["next5", "Next 5 years"], ["pastdue", "Past due"], ["all", "All years"]
  ];

  function renderPortfolio() {
    store.listProperties().then(function (props) {
      var cards = props.map(function (p) {
        var em = Engine.computeExpenditures(p);
        var rows = Engine.computeCashFlow(p, p.fundingPlan.contributions, em.totals);
        var minRow = rows.reduce(function (m, r) { return r.end < m.end ? r : m; }, rows[0]);
        var neg = minRow.end < 0;
        return '<div class="pcard" data-id="' + p.id + '">' +
          '<div class="pcard-head"><div class="pcard-name">' + esc(p.name) + "</div>" +
          '<button class="icon-btn pcard-del" title="Remove property" data-id="' + p.id + '">✕</button></div>' +
          '<div class="pcard-meta">' + esc(p.city) + (p.state ? ", " + esc(p.state) : "") +
          (p.numUnits ? " · " + esc(p.numUnits) + " units" : "") + "</div>" +
          '<div class="pcard-stats">' +
          '<div><span class="lbl">Reserves</span>' + fmt$(p.current.beginningBalance) + "</div>" +
          '<div><span class="lbl">Low point</span><span class="' + (neg ? "neg" : "pos") + '">' +
          fmt$(minRow.end) + " (" + minRow.year + ")</span></div>" +
          '<div><span class="lbl">Plan</span>' + esc(p.fundingPlan.mode) + "</div>" +
          "</div>" +
          (p.fundingPlan.mode !== "recommended" || hasEdits(p) ?
            '<div class="pcard-flag">modified from study</div>' : "") +
          "</div>";
      }).join("");

      var occRows = occurrenceRows(props);
      var filtered = applyPortfolioFilters(occRows)
        .sort(function (a, b) { return a.year - b.year || b.amount - a.amount; });
      var coord = coordinationGroups(applyPortfolioFilters(occRows));

      var totalReserves = props.reduce(function (s, p) { return s + (p.current.beginningBalance || 0); }, 0);

      root.innerHTML =
        '<div class="view portfolio">' +
        '<div class="page-head"><h1>Portfolio</h1>' +
        '<div class="page-stats">' + props.length + " propert" + (props.length === 1 ? "y" : "ies") +
        " · combined reserves " + fmt$(totalReserves) + "</div></div>" +
        uploadZoneHTML("uz-portfolio", props.length ? "Add properties" : "Upload reserve study workbooks") +
        (props.length ? '<div class="pcard-grid">' + cards + "</div>" : "") +
        (props.length ?
          '<section class="panel"><h2>Projects across the portfolio</h2>' +
          '<div class="filter-bar">' +
          '<div class="btn-group" id="pf-window">' + WINDOW_BUTTONS.map(function (b) {
            return '<button class="fbtn' + (portfolioFilter.window === b[0] ? " on" : "") +
              '" data-w="' + b[0] + '">' + b[1] + "</button>";
          }).join("") + "</div>" +
          '<input type="search" id="pf-search" placeholder="Search components… (e.g. roof)" value="' + esc(portfolioFilter.search) + '">' +
          '<input type="number" id="pf-costmin" placeholder="Min $" value="' + (portfolioFilter.costMin != null ? portfolioFilter.costMin : "") + '">' +
          '<input type="number" id="pf-costmax" placeholder="Max $" value="' + (portfolioFilter.costMax != null ? portfolioFilter.costMax : "") + '">' +
          '<select id="pf-prop"><option value="all">All properties</option>' +
          props.map(function (p) {
            return '<option value="' + p.id + '"' + (portfolioFilter.propertyId === p.id ? " selected" : "") + ">" + esc(p.name) + "</option>";
          }).join("") + "</select>" +
          "</div>" +
          (coord.length ?
            '<div class="coord"><h3>Coordination opportunities <span class="hint">(same component family, ≥2 properties, ≤3-year window — click a card to filter)</span></h3>' +
            '<div class="coord-cards">' + coord.slice(0, 12).map(function (g) {
              var active = portfolioFilter.coordKey === g.key &&
                portfolioFilter.coordFrom === g.yearFrom && portfolioFilter.coordTo === g.yearTo;
              return '<div class="coord-card' + (active ? " active" : "") + '" data-key="' + esc(g.key) +
                '" data-from="' + g.yearFrom + '" data-to="' + g.yearTo + '" data-label="' + esc(g.label) + '" role="button" tabindex="0">' +
                '<div class="cc-title">' + esc(g.label) + "</div>" +
                '<div class="cc-meta">' + (g.yearFrom === g.yearTo ? g.yearFrom : g.yearFrom + "–" + g.yearTo) +
                " · " + g.properties.length + " properties · " + fmt$(g.total) + " combined</div>" +
                '<div class="cc-props">' + g.properties.map(esc).join(" · ") + "</div></div>";
            }).join("") + "</div></div>" : "") +
          (portfolioFilter.coordKey ?
            '<div class="active-filter">Filtered to <b>' + esc(portfolioFilter.coordKey) + "</b> · " +
            (portfolioFilter.coordFrom === portfolioFilter.coordTo ? portfolioFilter.coordFrom
              : portfolioFilter.coordFrom + "–" + portfolioFilter.coordTo) +
            ' <button class="chip-clear" id="coord-clear">clear ✕</button></div>' : "") +
          '<div class="tbl-wrap"><table class="tbl"><thead><tr>' +
          "<th>Property</th><th>Project</th><th>Category</th><th>Year</th><th class=\"r\">Est. cost</th>" +
          "</tr></thead><tbody>" +
          (filtered.length ? filtered.slice(0, 400).map(function (r) {
            return "<tr><td><a href=\"#/property/" + r.property.id + '">' + esc(r.property.name) + "</a></td>" +
              "<td>" + esc(r.component.name) + "</td><td>" + esc(r.component.category || "") + "</td>" +
              "<td>" + r.year + '</td><td class="r">' + fmt$(r.amount) + "</td></tr>";
          }).join("") : '<tr><td colspan="5" class="empty">No projects match the current filters.</td></tr>') +
          "</tbody></table>" +
          (filtered.length > 400 ? '<div class="hint">Showing first 400 of ' + filtered.length + " rows — narrow the filters.</div>" : "") +
          "</div></section>" : "") +
        "</div>";

      wireUploadZone("uz-portfolio", {});
      $all(".pcard").forEach(function (card) {
        card.addEventListener("click", function (e) {
          if (e.target.classList.contains("pcard-del")) return;
          location.hash = "#/property/" + card.getAttribute("data-id");
        });
      });
      $all(".pcard-del").forEach(function (btn) {
        btn.addEventListener("click", function (e) {
          e.stopPropagation();
          if (confirm("Remove this property from the dashboard? (The Excel file is untouched.)")) {
            store.deleteProperty(btn.getAttribute("data-id")).then(render);
          }
        });
      });
      // Coordination card → filter the project list to that component family
      // AND to the card's year window (e.g. "Pool Finish 2027–2029").
      $all(".coord-card").forEach(function (card) {
        function apply() {
          var key = card.getAttribute("data-key");
          var from = +card.getAttribute("data-from"), to = +card.getAttribute("data-to");
          var isActive = portfolioFilter.coordKey === key &&
            portfolioFilter.coordFrom === from && portfolioFilter.coordTo === to;
          if (isActive) {
            portfolioFilter.coordKey = portfolioFilter.coordFrom = portfolioFilter.coordTo = null;
          } else {
            portfolioFilter.coordKey = key;
            portfolioFilter.coordFrom = from;
            portfolioFilter.coordTo = to;
            portfolioFilter.window = "all"; // year range comes from the card
          }
          renderPortfolio();
          var tbl = $(".portfolio .tbl-wrap");
          if (tbl && tbl.scrollIntoView) tbl.scrollIntoView({ behavior: "smooth", block: "center" });
        }
        card.addEventListener("click", apply);
        card.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); apply(); } });
      });
      var coordClear = $("#coord-clear");
      if (coordClear) coordClear.addEventListener("click", function () {
        portfolioFilter.coordKey = portfolioFilter.coordFrom = portfolioFilter.coordTo = null;
        renderPortfolio();
      });
      var w = $("#pf-window");
      if (w) w.addEventListener("click", function (e) {
        var b = e.target.closest("button[data-w]");
        if (b) { portfolioFilter.window = b.getAttribute("data-w"); renderPortfolio(); }
      });
      var se = $("#pf-search");
      if (se) se.addEventListener("input", debounce(function () {
        portfolioFilter.search = se.value; renderPortfolio();
        setTimeout(function () { var x = $("#pf-search"); if (x) { x.focus(); x.setSelectionRange(x.value.length, x.value.length); } }, 0);
      }, 350));
      var cmin = $("#pf-costmin"), cmax = $("#pf-costmax"), pp = $("#pf-prop");
      if (cmin) cmin.addEventListener("change", function () {
        portfolioFilter.costMin = cmin.value === "" ? null : +cmin.value; renderPortfolio();
      });
      if (cmax) cmax.addEventListener("change", function () {
        portfolioFilter.costMax = cmax.value === "" ? null : +cmax.value; renderPortfolio();
      });
      if (pp) pp.addEventListener("change", function () {
        portfolioFilter.propertyId = pp.value; renderPortfolio();
      });
    });
  }

  function hasEdits(p) {
    return p.components.some(function (c) {
      return c.deleted || (c.original && (
        c.unitCost !== c.original.unitCost ||
        c.nextReplYear !== c.original.nextReplYear ||
        c.phaseLength !== c.original.phaseLength ||
        c.eventsPerPhase !== c.original.eventsPerPhase));
    }) || p.rates.inflationOverride != null;
  }

  /* =====================================================================
   * PROPERTY VIEW
   * =================================================================== */
  var propFilter = { window: "all", search: "", category: "all" };
  var propSort = { key: null, dir: 1 }; // key: name|category|year|cost ; dir: 1 asc / -1 desc
  var hiddenSeries = {};                 // chart series id -> true when hidden
  var currentProp = null;
  var saveSoon = debounce(function () {
    if (currentProp) store.saveProperty(currentProp).catch(function (e) { toast(e.message, true); });
  }, 400);

  function recomputed(p) {
    var em = Engine.computeExpenditures(p);
    var rows = Engine.computeCashFlow(p, p.fundingPlan.contributions, em.totals);
    return { em: em, rows: rows };
  }

  // Backfill fields added in later versions onto properties saved earlier,
  // so old localStorage entries render correctly (e.g. component.original
  // gaining usefulLife/totalQty; fundingPlan.assessments; savedPlans).
  function migrateProperty(p) {
    var n = p.studyLength;
    p.savedPlans = p.savedPlans || [];
    // Backfill the study baseline; older saves may lack it or its date, which
    // caused "Reset to study" to blank the balance date.
    p._origAssume = p._origAssume || {};
    var oa = p._origAssume, cur = p.current;
    if (oa.interest == null) oa.interest = p.rates.interest;
    if (oa.balance == null) oa.balance = cur.beginningBalance;
    if (oa.balanceDate == null) oa.balanceDate = cur.balanceDate;
    if (oa.contrib == null) oa.contrib = cur.annualContribution;
    if (oa.remainingPeriods == null) oa.remainingPeriods = cur.remainingPeriods;
    if (oa.remainingInterestMonths == null) oa.remainingInterestMonths = cur.remainingInterestMonths;
    if (!p.fundingPlan.assessments || p.fundingPlan.assessments.length !== n + 1) {
      var a = new Array(n + 1).fill(0);
      (p.fundingPlan.assessments || []).forEach(function (v, i) { if (i <= n) a[i] = v || 0; });
      p.fundingPlan.assessments = a;
    }
    (p.components || []).forEach(function (c) {
      if (!c.original) return;
      ["unitCost", "nextReplYear", "usefulLife", "totalQty", "units", "ownership",
       "phaseLength", "eventsPerPhase", "frequencyOfEvents"].forEach(function (k) {
        if (c.original[k] == null && c[k] != null) c.original[k] = c[k];
      });
    });
    return p;
  }

  function renderProperty(id) {
    var task = currentProp && currentProp.id === id
      ? Promise.resolve(currentProp)
      : store.getProperty(id);
    task.then(function (p) {
      if (!p) { root.innerHTML = '<div class="view"><p>Property not found. <a href="#/portfolio">Back to portfolio</a></p></div>'; return; }
      currentProp = migrateProperty(p);
      drawProperty();
    });
  }

  function drawProperty() {
    var p = currentProp;
    var r = recomputed(p);
    var rows = r.rows, em = r.em;
    var minRow = rows.reduce(function (m, x) { return x.end < m.end ? x : m; }, rows[0]);
    var total30 = em.totals.reduce(function (s, v) { return s + v; }, 0);
    var notes = (p.parseWarnings || []).concat(p.overrideNotes || []);
    var fp = p.fundingPlan;
    var threshold = fp.mode === "threshold" ? fp.thresholdValue : null;

    root.innerHTML =
      '<div class="view property">' +
      '<a class="back" href="#/portfolio">← Portfolio</a>' +
      '<div class="page-head"><h1>' + esc(p.name) + "</h1>" +
      '<div class="page-stats">' + esc(p.city) + (p.state ? ", " + esc(p.state) : "") +
      (p.numUnits ? " · " + esc(p.numUnits) + " units" : "") +
      " · FY " + p.currentFiscalYear + " · " + p.studyLength + "-year study" +
      (p.refNumber ? " · #" + esc(p.refNumber) : "") + "</div></div>" +

      (notes.length ?
        '<details class="notes"><summary>' + notes.length + " note" + (notes.length > 1 ? "s" : "") +
        " from the imported study</summary><ul>" +
        notes.map(function (n) { return "<li>" + esc(n) + "</li>"; }).join("") + "</ul></details>" : "") +

      '<div class="kpis">' +
      kpiEdit("Starting reserves", fmt$(p.current.beginningBalance),
        (p.current.balanceDate ? "as of " + p.current.balanceDate : "") +
        " · contribution " + fmt$(p.current.annualContribution) + "/yr", "edit-current") +
      kpi("Lowest year-end balance", '<span class="' + (minRow.end < 0 ? "neg" : "pos") + '">' + fmt$(minRow.end) + "</span>", "in " + minRow.year) +
      kpiEdit("Interest rate", (p.rates.interest * 100).toFixed(2) + "%", "earned on reserve balance", "edit-interest") +
      kpiEdit("Inflation rate", ((p.rates.inflationOverride != null ? p.rates.inflationOverride : p.rates.remainingInflation) * 100).toFixed(2) + "%",
        p.rates.inflationOverride != null ? "override (study " + (p.rates.remainingInflation * 100).toFixed(2) + "%)" : "applied to future costs", "edit-inflation") +
      "</div>" +
      '<div class="assume-reset"><button class="mini" id="btn-reset-assume">↺ Reset to study assumptions</button>' +
      '<span class="hint">Restores balance, date, contribution, interest &amp; inflation to the imported study.</span></div>' +

      '<section class="panel"><h2>Reserve balance & expenditures</h2>' +
      '<div class="chartwrap"><div id="chart"></div>' +
      '<div id="chart-tip" class="chart-tip" style="display:none"></div></div>' +
      '<div class="legend" id="chart-legend"></div>' +
      "</section>" +

      '<section class="panel"><h2>Funding plan</h2>' +
      '<div class="scenario-bar">' +
      '<button class="btn' + (fp.mode === "recommended" ? " on" : "") + '" id="sc-rec">Study recommendation</button>' +
      '<button class="btn' + (fp.mode === "baseline" ? " on" : "") + '" id="sc-base">Baseline (minimum funding)</button>' +
      '<span class="thr-group"><button class="btn' + (fp.mode === "threshold" ? " on" : "") + '" id="sc-thr">Threshold</button>' +
      '<input type="text" inputmode="numeric" id="thr-value" placeholder="Floor $" value="' +
      (fp.thresholdValue != null ? money(fp.thresholdValue) : "") + '"></span>' +
      '<span class="mode-note" id="mode-note">' + esc(fp.modeNote || "") + "</span>" +
      "</div>" +
      planBarHTML(p) +
      '<div class="hint">Edit any year in the ledger to fine-tune — manual edits switch the plan to “manual”. ' +
      "Additional assessments are added in full to that year. Scenario-generated contributions round to " +
      fmt$(p.roundTo) + ". Balances update as you type.</div>" +
      fundingLedgerHTML(rows, fp) +
      "</section>" +

      '<section class="panel"><h2>Projects <span class="hint" id="proj-count"></span></h2>' +
      '<div class="filter-bar">' +
      '<div class="btn-group" id="pv-window">' + WINDOW_BUTTONS.map(function (b) {
        return '<button class="fbtn' + (propFilter.window === b[0] ? " on" : "") + '" data-w="' + b[0] + '">' + b[1] + "</button>";
      }).join("") + "</div>" +
      '<input type="search" id="pv-search" placeholder="Search components… (e.g. roof)" value="' + esc(propFilter.search) + '">' +
      categorySelectHTML(p) +
      '<button class="mini" id="pv-reset">Reset view</button>' +
      "</div>" +
      '<div class="tbl-wrap"><table class="tbl projects"><thead><tr>' +
      sortableTh("Project", "name") + sortableTh("Category", "category") +
      "<th class=\"r\">Qty</th><th class=\"r\">Unit cost</th>" +
      sortableTh("Next year", "year") +
      "<th class=\"r\">Useful life</th>" +
      sortableTh("Next cost", "cost", true) + "<th>Actions</th>" +
      '</tr></thead><tbody id="proj-body"></tbody></table></div></section>' +

      '<section class="panel"><details><summary><h2 class="inline">Cash flow table</h2></summary>' +
      '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Year</th><th class="r">Begin</th>' +
      '<th class="r">Contributions</th><th class="r">Add’l assessment</th><th class="r">Interest</th>' +
      '<th class="r">Expenditures</th><th class="r">Year end</th></tr></thead><tbody>' +
      rows.map(function (row) {
        return "<tr><td>" + row.year + '</td><td class="r">' + fmt$(row.begin) + '</td><td class="r">' + fmt$(row.contribution) +
          '</td><td class="r">' + (row.assessment ? fmt$(row.assessment) : "—") +
          '</td><td class="r">' + fmt$(row.interest) + '</td><td class="r">' + fmt$(row.expenditures) +
          '</td><td class="r ' + (row.end < 0 ? "neg" : "") + '">' + fmt$(row.end) + "</td></tr>";
      }).join("") + "</tbody></table></div></details></section>" +

      '<div id="modal-mount"></div>' +
      "</div>";

    drawProjectRows();
    wirePropertyEvents();
    renderChart(p);
  }

  function kpi(label, valueHTML, sub) {
    return '<div class="kpi"><div class="kpi-label">' + label + '</div><div class="kpi-value">' +
      valueHTML + '</div><div class="kpi-sub">' + esc(sub || "") + "</div></div>";
  }
  function kpiEdit(label, valueHTML, sub, btnId) {
    return '<div class="kpi"><button class="kpi-edit" id="' + btnId + '" title="Edit">✎</button>' +
      '<div class="kpi-label">' + label + '</div><div class="kpi-value">' + valueHTML +
      '</div><div class="kpi-sub">' + esc(sub || "") + "</div></div>";
  }
  function pctInput(id, label, value) {
    return '<label class="ni"><span>' + esc(label) + '</span><input type="number" id="' + id +
      '" value="' + value + '" step="0.05"></label>';
  }
  function moneyInputHTML(id, label) {
    return '<label class="ni"><span>' + esc(label) + '</span>' +
      '<input type="text" inputmode="numeric" class="money-in" id="' + id + '"></label>';
  }
  function categorySelectHTML(p) {
    var cats = {};
    p.components.forEach(function (c) { if (c.category) cats[c.category] = 1; });
    var opts = Object.keys(cats).sort().map(function (c) {
      return '<option value="' + esc(c) + '"' + (propFilter.category === c ? " selected" : "") + ">" + esc(c) + "</option>";
    }).join("");
    return '<select id="pv-cat"><option value="all"' + (propFilter.category === "all" ? " selected" : "") +
      ">All categories</option>" + opts + "</select>";
  }
  function sortableTh(label, key, right) {
    var on = propSort.key === key;
    var arrow = on ? (propSort.dir === 1 ? "▲" : "▼") : "↕";
    return '<th class="sortable' + (right ? " r" : "") + (on ? " sorted" : "") + '" data-sort="' + key + '">' +
      esc(label) + '<span class="arrow">' + arrow + "</span></th>";
  }

  // ----- plan selector bar ----------------------------------------------
  var PLAN_COLORS = ["#1F5132", "#B5552F", "#3C8C40", "#6A5ACD", "#0E7C86", "#9C27B0"];
  function planColor(i) { return PLAN_COLORS[i % PLAN_COLORS.length]; }
  function workingColor() { return "#A8852C"; } // brass

  function planBarHTML(p) {
    var chips = (p.savedPlans || []).map(function (pl, i) {
      return '<span class="plan-chip" data-pid="' + pl.id + '">' +
        '<span class="plan-swatch" style="background:' + planColor(i) + '"></span>' +
        esc(pl.name) + '<button class="plan-del" data-pid="' + pl.id + '" title="Delete plan">✕</button></span>';
    }).join("");
    return '<div class="plan-bar">' +
      '<button class="mini primary" id="plan-save">Save current plan…</button>' +
      (p.savedPlans && p.savedPlans.length
        ? '<select id="plan-load"><option value="">Load a saved plan…</option>' +
          p.savedPlans.map(function (pl) { return '<option value="' + pl.id + '">' + esc(pl.name) + "</option>"; }).join("") +
          "</select>"
        : '<span class="hint">Save a plan to compare it as a second line on the chart.</span>') +
      (chips ? '<span class="plan-chips">' + chips + "</span>" : "") +
      "</div>";
  }

  // ----- funding ledger -------------------------------------------------
  function fundingLedgerHTML(rows, fp) {
    var body = rows.slice(1).map(function (row, i) {
      var k = i + 1;
      var edited = fp.mode === "manual" && fp._editedYears && fp._editedYears[k];
      var assess = (fp.assessments && fp.assessments[k]) || 0;
      return '<div class="ledger-row' + (edited ? " edited" : "") + '" data-k="' + k + '">' +
        '<span class="yr">' + row.year + "</span>" +
        '<input class="ledger-in contrib" type="text" inputmode="numeric" data-k="' + k + '" value="' + money(fp.contributions[k] || 0) + '">' +
        '<input class="ledger-in assess" type="text" inputmode="numeric" data-ak="' + k + '" value="' + (assess ? money(assess) : "") + '" placeholder="—">' +
        '<span class="bal' + (row.end < 0 ? " neg" : "") + '" data-balk="' + k + '">' + money(row.end) + "</span>" +
        '<span class="dot">' + (edited ? "●" : "") + "</span>" +
        "</div>";
    }).join("");
    return '<div class="ledger" id="contrib-grid">' +
      '<div class="ledger-head"><span>Year</span><span class="r">Contribution</span>' +
      '<span class="r">Add’l assessment</span><span class="r">Projected balance</span><span></span></div>' +
      '<div class="ledger-body">' + body + "</div></div>";
  }
  function labelForMode(fp) {
    switch (fp.mode) {
      case "recommended": return "Study recommendation";
      case "baseline": return "Baseline (minimum)";
      case "threshold": return "Threshold ≥ " + fmt$(fp.thresholdValue);
      case "manual": return "Manual";
      default: return fp.mode;
    }
  }

  // ----- projects table -------------------------------------------------
  function componentEdited(c) {
    var o = c.original;
    if (!o) return false;
    // Only flag a field when the study baseline for it is known — guards
    // against older saved properties whose snapshot lacked some fields.
    function diff(key) { return o[key] != null && c[key] !== o[key]; }
    return diff("unitCost") || diff("nextReplYear") || diff("usefulLife") ||
      diff("phaseLength") || diff("eventsPerPhase");
  }

  function drawProjectRows() {
    var p = currentProp;
    var em = Engine.computeExpenditures(p);
    var byId = {};
    em.perComponent.forEach(function (pc) { byId[pc.component.id] = pc.amounts; });
    var s = propFilter.search.trim().toLowerCase();

    var visible = p.components.filter(function (c) {
      if (propFilter.category !== "all" && (c.category || "") !== propFilter.category) return false;
      if (s && c.name.toLowerCase().indexOf(s) === -1 &&
          (c.category || "").toLowerCase().indexOf(s) === -1) return false;
      if (propFilter.window === "all") return true;
      if (propFilter.window === "pastdue") return c.nextReplYear < p.currentFiscalYear;
      if (c.deleted) return false;
      var amounts = byId[c.id] || [];
      var lo, hi;
      if (propFilter.window === "this") { lo = 0; hi = 0; }
      else if (propFilter.window === "next") { lo = 1; hi = 1; }
      else if (propFilter.window === "next2") { lo = 1; hi = 2; }
      else { lo = 1; hi = 5; }
      for (var k = lo; k <= hi; k++) if (amounts[k] > 0) return true;
      return false;
    });

    // sort
    function nextInfo(c) {
      var amounts = byId[c.id] || [];
      for (var k = 0; k < amounts.length; k++) if (amounts[k] > 0) return { k: k, cost: amounts[k] };
      return { k: 999, cost: -1 };
    }
    if (propSort.key) {
      visible.sort(function (a, b) {
        var av, bv;
        if (propSort.key === "name") { av = a.name.toLowerCase(); bv = b.name.toLowerCase(); }
        else if (propSort.key === "category") { av = (a.category || "").toLowerCase(); bv = (b.category || "").toLowerCase(); }
        else if (propSort.key === "year") { av = a.nextReplYear; bv = b.nextReplYear; }
        else { av = nextInfo(a).cost; bv = nextInfo(b).cost; } // cost
        if (av < bv) return -1 * propSort.dir;
        if (av > bv) return 1 * propSort.dir;
        return a.name.localeCompare(b.name);
      });
    }

    var cnt = $("#proj-count");
    if (cnt) cnt.textContent = "(" + visible.length + " of " + p.components.length + ")";

    var body = $("#proj-body");
    if (!body) return;
    body.innerHTML = visible.length ? visible.map(function (c) {
      var amounts = byId[c.id] || [];
      var ni = nextInfo(c);
      var phased = c.eventsPerPhase > 1;
      var edited = componentEdited(c);
      var ulNote = phased
        ? c.eventsPerPhase + " phases · " + (c.frequencyOfEvents || 1) + " yr apart"
        : "";
      return '<tr class="' + (c.deleted ? "deleted" : "") + '" data-cid="' + c.id + '">' +
        '<td><span class="cname">' + esc(c.name) + "</span>" +
        (c.overrides ? ' <span class="badge tuned" title="Hand-tuned year cells from the study (kept as-is).">tuned</span>' : "") +
        (edited ? ' <span class="badge edit">edited</span>' : "") +
        (c.deleted ? ' <span class="badge del">deleted</span>' : "") + "</td>" +
        '<td><span class="ccat">' + esc(c.category || "") + "</span></td>" +
        '<td class="r qtycol">' + (c.totalQty != null ? c.totalQty.toLocaleString() : "") + " " + esc(c.units || "") + "</td>" +
        '<td class="r"><input type="text" inputmode="decimal" class="cell-in cost money-in" data-cid="' + c.id + '" value="' + moneyCents(c.unitCost) + '"' + (c.deleted ? " disabled" : "") + "></td>" +
        '<td class="r"><input type="number" class="cell-in year" data-cid="' + c.id + '" value="' + c.nextReplYear + '" step="1" min="' +
        (p.currentFiscalYear - 30) + '" max="' + (p.currentFiscalYear + p.studyLength + 30) + '"' + (c.deleted ? " disabled" : "") + "></td>" +
        '<td class="r"><input type="number" class="cell-in ul" data-cid="' + c.id + '" value="' + c.usefulLife + '" step="1" min="1" max="120"' + (c.deleted ? " disabled" : "") + ">" +
        (ulNote ? '<div class="sched-note">' + ulNote + "</div>" : "") + "</td>" +
        '<td class="r">' + (ni.k <= p.studyLength ? fmt$(ni.cost) + " <span class=\"hint\">(" + (p.currentFiscalYear + ni.k) + ")</span>" : "<span class=\"hint\">beyond study</span>") + "</td>" +
        '<td class="actions">' +
        '<button class="mini primary" data-act="edit" data-cid="' + c.id + '"' + (c.deleted ? " disabled" : "") + ">Edit</button>" +
        (c.deleted
          ? '<button class="mini" data-act="restore" data-cid="' + c.id + '">Restore</button>'
          : '<button class="mini warn" data-act="delete" data-cid="' + c.id + '">Delete</button>') +
        (edited ? '<button class="mini" data-act="reset" data-cid="' + c.id + '">Reset</button>' : "") +
        "</td></tr>";
    }).join("") : '<tr><td colspan="8" class="empty">No projects match the current filters.</td></tr>';
  }

  // ----- chart (multi-plan) + tooltip + legend --------------------------
  function planSeriesEnds(p, contribs, assess, expTotals) {
    return Engine.computeCashFlow(p, contribs, expTotals, assess).map(function (r) { return r.end; });
  }

  // previewWorking (optional) = {contributions, assessments} for live typing.
  function renderChart(p, previewWorking) {
    var chartEl = $("#chart");
    if (!chartEl) return;
    var em = Engine.computeExpenditures(p);
    var fp = p.fundingPlan;
    var wContribs = previewWorking ? previewWorking.contributions : fp.contributions;
    var wAssess = previewWorking ? previewWorking.assessments : fp.assessments;
    var workingRows = Engine.computeCashFlow(p, wContribs, em.totals, wAssess);

    var series = [{ id: "working", label: "Working plan", color: workingColor(),
                    end: workingRows.map(function (r) { return r.end; }) }];
    (p.savedPlans || []).forEach(function (pl, i) {
      series.push({ id: pl.id, label: pl.name, color: planColor(i),
        end: planSeriesEnds(p, pl.contributions, pl.assessments, em.totals) });
    });
    var threshold = fp.mode === "threshold" ? fp.thresholdValue : null;

    chartEl.innerHTML = Charts.planChart({
      startYear: p.currentFiscalYear, expenditures: em.totals,
      threshold: threshold, series: series, hidden: hiddenSeries
    });

    // legend
    var leg = $("#chart-legend");
    if (leg) {
      var items = '<span class="leg-static"><span class="sw bar"></span>Expenditures</span>';
      series.forEach(function (s) {
        items += '<span class="leg-item' + (hiddenSeries[s.id] ? " off" : "") + '" data-series="' + s.id + '">' +
          '<span class="sw line" style="background:' + s.color + '"></span>' + esc(s.label) + "</span>";
      });
      if (threshold != null) items += '<span class="leg-static"><span class="sw thr"></span>Threshold floor</span>';
      leg.innerHTML = items;
      $all(".leg-item", leg).forEach(function (it) {
        it.addEventListener("click", function () {
          var id = it.getAttribute("data-series");
          if (hiddenSeries[id]) delete hiddenSeries[id]; else hiddenSeries[id] = true;
          renderChart(p, previewWorking);
        });
      });
    }

    // per-year projects for the tooltip (based on the shown expenditures)
    var perYear = {};
    em.perComponent.forEach(function (pc) {
      pc.amounts.forEach(function (a, k) {
        if (a > 0) (perYear[k] = perYear[k] || []).push({ name: pc.component.name, cost: a });
      });
    });
    attachChartTooltip(p, workingRows, perYear);
  }

  function attachChartTooltip(p, rows, perYear) {
    var wrap = $(".chartwrap");
    var tip = $("#chart-tip");
    if (!wrap || !tip) return;
    $all(".hovercol", wrap).forEach(function (col) {
      col.addEventListener("mousemove", function (e) {
        var i = +col.getAttribute("data-i");
        var row = rows[i];
        if (!row) return;
        var projs = (perYear[i] || []).slice().sort(function (a, b) { return b.cost - a.cost; });
        var list = projs.length
          ? projs.slice(0, 8).map(function (pr) {
              return '<div class="tip-proj"><span>' + esc(pr.name) + '</span><span>' + fmt$(pr.cost) + "</span></div>";
            }).join("") + (projs.length > 8 ? '<div class="tip-more">+ ' + (projs.length - 8) + " more…</div>" : "")
          : '<div class="tip-none">No projects this year</div>';
        tip.innerHTML =
          '<div class="tip-year">' + row.year + "</div>" +
          '<div class="tip-line"><span>Contribution</span><b>' + fmt$(row.contribution) + "</b></div>" +
          (row.assessment ? '<div class="tip-line"><span>Add’l assessment</span><b>' + fmt$(row.assessment) + "</b></div>" : "") +
          '<div class="tip-line"><span>Year-end balance</span><b class="' + (row.end < 0 ? "neg" : "") + '">' + fmt$(row.end) + "</b></div>" +
          '<div class="tip-projhd">Projects (' + projs.length + ")</div>" + list;
        var wr = wrap.getBoundingClientRect();
        var x = e.clientX - wr.left, y = e.clientY - wr.top;
        tip.style.display = "block";
        var tw = tip.offsetWidth, th = tip.offsetHeight;
        tip.style.left = Math.max(4, Math.min(x + 14, wr.width - tw - 4)) + "px";
        tip.style.top = Math.max(4, Math.min(y + 14, wr.height - th - 4)) + "px";
      });
      col.addEventListener("mouseleave", function () { tip.style.display = "none"; });
    });
  }

  // ----- assumptions: editable KPI cards --------------------------------
  function ensureOrigAssume(p) {
    if (!p._origAssume) p._origAssume = {
      interest: p.rates.interest, balance: p.current.beginningBalance,
      balanceDate: p.current.balanceDate, contrib: p.current.annualContribution,
      remainingPeriods: p.current.remainingPeriods,
      remainingInterestMonths: p.current.remainingInterestMonths
    };
  }
  // Given a balance date, derive remaining monthly contributions in the
  // current fiscal year (contribution for the balance date's month is treated
  // as already made — matches the study's 6-of-12 at a mid-June date).
  function remainingFromDate(p, isoDate) {
    var d = isoDate ? new Date(isoDate) : null;
    if (!d || isNaN(d.getTime())) {
      return { periods: p.current.remainingPeriods, months: p.current.remainingInterestMonths };
    }
    var monthWithinFY = ((d.getUTCMonth() + 1 - 1 + 12) % 12) + 1; // FY starts Jan
    var remaining = Math.max(0, Math.min(12, 12 - monthWithinFY));
    return { periods: remaining, months: remaining };
  }

  // Small modal shell mounted into #modal-mount. Returns {close}.
  function mountModal(title, sub, bodyHTML, footHTML) {
    var mount = $("#modal-mount");
    mount.innerHTML =
      '<div class="modal-back" id="mm-back"><div class="modal" role="dialog" aria-modal="true">' +
      '<div class="modal-head"><div><h3>' + esc(title) + "</h3>" +
      (sub ? '<div class="sub">' + esc(sub) + "</div>" : "") + "</div>" +
      '<button class="modal-close" id="mm-x" aria-label="Close">×</button></div>' +
      '<div class="modal-body">' + bodyHTML + "</div>" +
      '<div class="modal-foot"><div class="left"></div><div>' + footHTML + "</div></div>" +
      "</div></div>";
    function close() { mount.innerHTML = ""; }
    $("#mm-x").addEventListener("click", close);
    $("#mm-back").addEventListener("mousedown", function (e) { if (e.target.id === "mm-back") close(); });
    return { close: close };
  }

  function openCurrentModal() {
    var p = currentProp;
    ensureOrigAssume(p);
    var body =
      '<div class="field-grid">' +
      '<div class="field"><label>Beginning reserve balance</label>' +
      '<input type="text" inputmode="numeric" class="money-in" id="cm-balance" value="' + money(p.current.beginningBalance) + '"></div>' +
      '<div class="field"><label>Balance date</label>' +
      '<input type="date" id="cm-date" value="' + esc(p.current.balanceDate || "") + '"></div>' +
      '<div class="field"><label>Current-year contribution (annual)</label>' +
      '<input type="text" inputmode="numeric" class="money-in" id="cm-contrib" value="' + money(p.current.annualContribution) + '"></div>' +
      "</div>" +
      '<div class="explain" id="cm-explain"></div>';
    var foot = '<button class="btn subtle" id="cm-cancel">Cancel</button> <button class="btn" id="cm-save">Save</button>';
    var m = mountModal("Current reserves & contributions", p.name, body, foot);

    var em = Engine.computeExpenditures(p);
    function readVals() {
      return {
        balance: parseMoney($("#cm-balance").value) || 0,
        date: $("#cm-date").value || null,
        annual: parseMoney($("#cm-contrib").value) || 0
      };
    }
    var origDate = p.current.balanceDate || "";
    function remForDate(dateStr) {
      // Recompute from the date only when the user actually changes it;
      // otherwise keep the study's own remaining periods (which may not be a
      // plain calendar count).
      if ((dateStr || "") !== origDate) return remainingFromDate(p, dateStr);
      return { periods: p.current.remainingPeriods, months: p.current.remainingInterestMonths };
    }
    function update() {
      var v = readVals();
      var rem = remForDate(v.date);
      var dateChanged = (v.date || "") !== origDate;
      var monthly = v.annual / 12;
      var effective = monthly * rem.periods;
      var draftProp = Object.assign({}, p, {
        current: Object.assign({}, p.current, {
          beginningBalance: v.balance, remainingPeriods: rem.periods,
          totalPeriods: 12, remainingInterestMonths: rem.months
        })
      });
      var contribs = p.fundingPlan.contributions.slice(); contribs[0] = v.annual;
      var rows = Engine.computeCashFlow(draftProp, contribs, em.totals, p.fundingPlan.assessments);
      var r0 = rows[0];
      $("#cm-explain").innerHTML =
        "<p>Assuming <b>monthly</b> contributions, <b>" + rem.periods + "</b> of 12 remain this fiscal year" +
        (dateChanged ? " (the balance date's month is treated as already contributed)" : " (per the study)") +
        ". That's " + fmt$(monthly) + "/month × " + rem.periods + " = <b>" + fmt$(effective) +
        "</b> still to come in " + p.currentFiscalYear + ".</p>" +
        '<div class="explain-calc">' +
        '<div><span>Beginning balance</span><b>' + fmt$(v.balance) + "</b></div>" +
        '<div><span>+ Remaining contributions</span><b>' + fmt$(effective) + "</b></div>" +
        '<div><span>+ Interest (partial year)</span><b>' + fmt$(r0.interest) + "</b></div>" +
        '<div><span>− ' + p.currentFiscalYear + " expenditures</span><b>" + fmt$(r0.expenditures) + "</b></div>" +
        '<div class="tot"><span>Projected ' + p.currentFiscalYear + " year-end balance</span><b class=\"" +
        (r0.end < 0 ? "neg" : "") + '">' + fmt$(r0.end) + "</b></div></div>";
    }
    var balI = $("#cm-balance"), conI = $("#cm-contrib");
    [balI, conI].forEach(function (inp) {
      inp.addEventListener("focus", function () { inp.select(); });
      inp.addEventListener("blur", function () { var x = parseMoney(inp.value); inp.value = x == null ? "" : money(x); });
      inp.addEventListener("input", update);
    });
    $("#cm-date").addEventListener("input", update);
    update();

    $("#cm-cancel").addEventListener("click", m.close);
    $("#cm-save").addEventListener("click", function () {
      var v = readVals();
      var rem = remForDate(v.date);
      p.current.beginningBalance = v.balance;
      p.current.balanceDate = v.date;
      p.current.annualContribution = v.annual;
      p.current.remainingPeriods = rem.periods;
      p.current.totalPeriods = 12;
      p.current.remainingInterestMonths = rem.months;
      p.fundingPlan.contributions[0] = v.annual;
      m.close();
      refreshAfterEdit();
    });
  }

  function openRateModal(kind) {
    var p = currentProp;
    ensureOrigAssume(p);
    var isInterest = kind === "interest";
    var cur = isInterest ? p.rates.interest
      : (p.rates.inflationOverride != null ? p.rates.inflationOverride : p.rates.remainingInflation);
    var title = isInterest ? "Interest rate" : "Inflation rate";
    var desc = isInterest
      ? "Annual return earned on the reserve balance. Higher interest raises the year-end balance each year."
      : "Annual rate applied to future replacement costs. Higher inflation raises every future expenditure, so the plan must fund more.";
    var body =
      '<div class="field" style="max-width:12rem"><label>' + title + " (%)</label>" +
      '<input type="number" id="rm-val" step="0.05" value="' + (cur * 100).toFixed(2) + '"></div>' +
      '<p class="hint" style="margin-top:0.6rem">' + desc + "</p>";
    var foot = '<button class="btn subtle" id="rm-cancel">Cancel</button> <button class="btn" id="rm-save">Save</button>';
    var m = mountModal(title, p.name, body, foot);
    $("#rm-cancel").addEventListener("click", m.close);
    $("#rm-save").addEventListener("click", function () {
      var v = parseFloat($("#rm-val").value);
      if (isNaN(v)) { toast("Enter a percentage.", true); return; }
      if (isInterest) p.rates.interest = v / 100;
      else {
        var f = v / 100;
        p.rates.inflationOverride = (Math.abs(f - p.rates.remainingInflation) < 1e-9) ? null : f;
      }
      m.close();
      refreshAfterEdit();
    });
  }

  // ----- component edit modal -------------------------------------------
  // Full-variable editor (modeled on the RA online portal modal). Works on a
  // draft clone; nothing is committed until Save. Read-only figures and the
  // forecast chart recompute live as fields change.
  function openComponentModal(c) {
    var p = currentProp;
    var mount = $("#modal-mount");
    if (!mount) return;

    // draft clone (only fields the modal touches, plus what the engine reads)
    var draft = JSON.parse(JSON.stringify(c));
    var phases = draft.eventsPerPhase > 1 ? draft.eventsPerPhase : 1;
    var gap = draft.frequencyOfEvents > 1 ? draft.frequencyOfEvents : 1;

    function previewAmounts() {
      var temp = {
        studyLength: p.studyLength, currentFiscalYear: p.currentFiscalYear,
        rates: p.rates, colInflation: p.colInflation, components: [draft]
      };
      return Engine.computeExpenditures(temp).perComponent[0].amounts;
    }

    function fieldNum(label, id, value, step, min) {
      return '<div class="field"><label>' + esc(label) + "</label>" +
        '<input type="number" id="' + id + '" value="' + value + '" step="' + (step || 1) + '"' +
        (min != null ? ' min="' + min + '"' : "") + "></div>";
    }
    function fieldText(label, id, value) {
      return '<div class="field"><label>' + esc(label) + "</label>" +
        '<input type="text" id="' + id + '" value="' + esc(value) + '"></div>';
    }
    function fieldMoney(label, id, value) {
      return '<div class="field"><label>' + esc(label) + "</label>" +
        '<input type="text" inputmode="decimal" class="money-in" id="' + id + '" value="' + moneyCents(value) + '"></div>';
    }
    function fieldRO(label, id, hi) {
      return '<div class="field readonly' + (hi ? " hi" : "") + '"><label>' + esc(label) +
        '</label><div class="val" id="' + id + '">—</div></div>';
    }

    mount.innerHTML =
      '<div class="modal-back" id="cmodal-back"><div class="modal" role="dialog" aria-modal="true">' +
      '<div class="modal-head"><div><h3>' + esc(c.name) + "</h3>" +
      '<div class="sub">' + esc(c.category || "") + (c.overrides ? " · has hand-tuned year cells" : "") + "</div></div>" +
      '<button class="modal-close" id="cmodal-x" aria-label="Close">×</button></div>' +
      '<div class="modal-body">' +
      '<div class="field-grid">' +
      // Row 1
      fieldNum("Total Quantity", "m-totalqty", draft.totalQty != null ? draft.totalQty : "", "any", 0) +
      fieldMoney("Unit Cost", "m-cost", draft.unitCost) +
      fieldNum("Useful Life, Years", "m-ul", draft.usefulLife, 1, 1) +
      // Row 2
      fieldNum("1st Year of Replacement", "m-year", draft.nextReplYear, 1) +
      fieldNum("Number of Phases", "m-phases", phases, 1, 1) +
      fieldNum("Years Between Phases", "m-gap", gap, 1, 1) +
      // Row 3
      fieldNum("Percentage (%)", "m-own", Math.round((draft.ownership == null ? 1 : draft.ownership) * 100), 1, 0) +
      fieldRO("Cost of Replacement per Phase", "m-costphase") +
      fieldRO("Total Future Cost of Replacement", "m-future", true) +
      "</div>" +
      '<div class="modal-sub-lbl">Forecasted expenditures over the study</div>' +
      '<div class="modal-chartwrap" id="m-chart"></div>' +
      "</div>" +
      '<div class="modal-foot"><div class="left">' +
      (c.original ? '<button class="mini" id="cmodal-reset">Reset to study</button>' : "") +
      "</div><div>" +
      '<button class="btn subtle" id="cmodal-cancel">Cancel</button> ' +
      '<button class="btn" id="cmodal-save">Save changes</button></div></div>' +
      "</div></div>";

    function readDraft() {
      var tq = parseFloat($("#m-totalqty").value);
      if (!isNaN(tq)) draft.totalQty = tq;
      phases = Math.max(1, parseInt($("#m-phases").value, 10) || 1);
      gap = Math.max(1, parseInt($("#m-gap").value, 10) || 1);
      var yr = parseInt($("#m-year").value, 10);
      if (!isNaN(yr)) draft.nextReplYear = yr;
      var uc = parseMoney($("#m-cost").value);
      if (uc != null) draft.unitCost = uc;
      var ul = parseInt($("#m-ul").value, 10);
      if (!isNaN(ul) && ul >= 1) draft.usefulLife = ul;
      var own = parseFloat($("#m-own").value);
      if (!isNaN(own)) draft.ownership = own / 100;
      Engine.phaseComponent(draft, phases, gap);
    }
    function updatePreview() {
      readDraft();
      var perPhase = Engine.perPhaseQuantity(draft);
      var own = draft.ownership == null ? 1 : draft.ownership;
      var costPhase = perPhase * draft.unitCost * own;
      var amounts = previewAmounts();
      var future = amounts.reduce(function (s, v) { return s + v; }, 0);
      $("#m-costphase").textContent = money(costPhase);
      $("#m-future").textContent = money(future);
      $("#m-chart").innerHTML = Charts.barChart(amounts, p.currentFiscalYear);
    }

    // money field UX
    var costEl = $("#m-cost");
    costEl.addEventListener("focus", function () { costEl.select(); });
    costEl.addEventListener("blur", function () {
      var v = parseMoney(costEl.value); if (v != null) costEl.value = moneyCents(v);
    });
    $all("#cmodal-back input").forEach(function (inp) {
      inp.addEventListener("input", updatePreview);
    });
    updatePreview();

    function close() { mount.innerHTML = ""; }
    $("#cmodal-x").addEventListener("click", close);
    $("#cmodal-cancel").addEventListener("click", close);
    $("#cmodal-back").addEventListener("mousedown", function (e) {
      if (e.target.id === "cmodal-back") close();
    });
    var resetBtn = $("#cmodal-reset");
    if (resetBtn) resetBtn.addEventListener("click", function () {
      Object.assign(draft, JSON.parse(JSON.stringify(c.original)));
      draft.deleted = false;
      $("#m-totalqty").value = draft.totalQty != null ? draft.totalQty : "";
      $("#m-year").value = draft.nextReplYear;
      $("#m-cost").value = moneyCents(draft.unitCost);
      $("#m-ul").value = draft.usefulLife;
      $("#m-phases").value = draft.eventsPerPhase > 1 ? draft.eventsPerPhase : 1;
      $("#m-gap").value = draft.frequencyOfEvents > 1 ? draft.frequencyOfEvents : 1;
      updatePreview();
    });
    $("#cmodal-save").addEventListener("click", function () {
      readDraft();
      var span = (phases - 1) * gap + 1;
      if (phases > 1 && span > draft.usefulLife) {
        toast("Phases span " + span + " years but useful life is " + draft.usefulLife +
          " years — phases would overlap the next cycle. Reduce phases or spacing.", true);
        return;
      }
      // commit editable fields to the real component
      c.totalQty = draft.totalQty;
      c.unitCost = draft.unitCost;
      c.units = draft.units;
      c.nextReplYear = draft.nextReplYear;
      c.usefulLife = draft.usefulLife;
      c.ownership = draft.ownership;
      Engine.phaseComponent(c, phases, gap);
      close();
      refreshAfterEdit();
      toast("Component updated.");
    });
  }

  // ----- recalc + partial refresh ----------------------------------------
  function refreshAfterEdit() {
    saveSoon();
    // full redraw keeps KPIs/cash-flow/notes/chart consistent; cheap at this scale
    drawProperty();
  }

  function setSchedule(mode, schedule, note, thresholdValue) {
    var fp = currentProp.fundingPlan;
    fp.mode = mode;
    fp.modeNote = note || "";
    if (thresholdValue !== undefined) fp.thresholdValue = thresholdValue;
    if (schedule) fp.contributions = schedule;
    refreshAfterEdit();
  }

  function wirePropertyEvents() {
    var p = currentProp;

    // Assumptions now live in editable KPI cards (modals) + a reset button.
    ensureOrigAssume(p);
    var eCur = $("#edit-current"); if (eCur) eCur.addEventListener("click", openCurrentModal);
    var eInt = $("#edit-interest"); if (eInt) eInt.addEventListener("click", function () { openRateModal("interest"); });
    var eInf = $("#edit-inflation"); if (eInf) eInf.addEventListener("click", function () { openRateModal("inflation"); });
    var resetAssume = $("#btn-reset-assume");
    if (resetAssume) resetAssume.addEventListener("click", function () {
      var o = p._origAssume;
      p.rates.interest = o.interest;
      p.rates.inflationOverride = null;
      p.current.beginningBalance = o.balance;
      p.current.balanceDate = o.balanceDate;
      p.current.annualContribution = o.contrib;
      p.current.remainingPeriods = o.remainingPeriods;
      p.current.remainingInterestMonths = o.remainingInterestMonths;
      p.fundingPlan.contributions[0] = o.contrib;
      refreshAfterEdit();
      toast("Assumptions reset to the imported study.");
    });

    // Scenario buttons
    var em = function () { return Engine.computeExpenditures(p).totals; };
    var scRec = $("#sc-rec");
    if (scRec) scRec.addEventListener("click", function () {
      setSchedule("recommended", p.fundingPlan.recommended.slice(), "");
    });
    var scBase = $("#sc-base");
    if (scBase) scBase.addEventListener("click", function () {
      var res = Engine.baselineSchedule(p, em());
      var note = res.note || ("scaled to " + Math.round(res.f * 1000) / 10 + "% of recommended");
      setSchedule("baseline", res.schedule, note);
      if (!res.feasible) toast(res.note, true);
    });
    var thrEl = $("#thr-value");
    if (thrEl) {
      thrEl.addEventListener("focus", function () { thrEl.select(); });
      thrEl.addEventListener("blur", function () {
        var v = parseMoney(thrEl.value);
        thrEl.value = v == null ? "" : money(v);
      });
    }
    var scThr = $("#sc-thr");
    if (scThr) scThr.addEventListener("click", function () {
      var v = parseMoney($("#thr-value").value);
      if (v == null) { toast("Enter a floor dollar amount first.", true); return; }
      var res = Engine.thresholdSchedule(p, em(), v);
      var note = res.feasible
        ? (res.note || ("scaled to " + Math.round(res.f * 1000) / 10 + "% of recommended"))
        : res.note;
      setSchedule("threshold", res.schedule, note, v);
      if (!res.feasible) toast(res.note, true);
    });

    // Manual contribution + assessment edits — live preview on input, commit on change
    var grid = $("#contrib-grid");
    if (grid) {
      // read current ledger DOM into {contributions, assessments} arrays
      function readLedger() {
        var contribs = p.fundingPlan.contributions.slice();
        var assess = (p.fundingPlan.assessments || []).slice();
        $all("input.contrib", grid).forEach(function (inp) {
          var k = +inp.getAttribute("data-k"); var v = parseMoney(inp.value);
          if (v != null && v >= 0) contribs[k] = v;
        });
        $all("input.assess", grid).forEach(function (inp) {
          var k = +inp.getAttribute("data-ak"); var v = parseMoney(inp.value);
          assess[k] = (v != null && v >= 0) ? v : 0;
        });
        return { contributions: contribs, assessments: assess };
      }
      grid.addEventListener("input", function (e) {
        if (!e.target.closest("input.ledger-in")) return;
        var draft = readLedger();
        var em2 = Engine.computeExpenditures(p);
        var rows2 = Engine.computeCashFlow(p, draft.contributions, em2.totals, draft.assessments);
        rows2.slice(1).forEach(function (row, i) {
          var span = grid.querySelector('.bal[data-balk="' + (i + 1) + '"]');
          if (span) { span.textContent = money(row.end); span.classList.toggle("neg", row.end < 0); }
        });
        renderChart(p, draft); // live line without stealing focus
      });
      grid.addEventListener("change", function (e) {
        var input = e.target.closest("input.ledger-in");
        if (!input) return;
        var fp = p.fundingPlan;
        if (input.classList.contains("contrib")) {
          var k = +input.getAttribute("data-k");
          var v = parseMoney(input.value);
          if (v == null || v < 0) { input.value = money(fp.contributions[k] || 0); return; }
          fp.contributions[k] = v;
          fp.mode = "manual"; fp.modeNote = "";
          fp._editedYears = fp._editedYears || {}; fp._editedYears[k] = true;
        } else if (input.classList.contains("assess")) {
          var ak = +input.getAttribute("data-ak");
          var av = parseMoney(input.value);
          fp.assessments = fp.assessments || [];
          fp.assessments[ak] = (av != null && av >= 0) ? av : 0;
        }
        refreshAfterEdit();
      });
    }

    // Save / load / delete funding plans
    var planSave = $("#plan-save");
    if (planSave) planSave.addEventListener("click", function () {
      var name = window.prompt("Name this funding plan:",
        p.fundingPlan.mode === "recommended" ? "Study recommendation" :
        p.fundingPlan.mode === "baseline" ? "Baseline" :
        p.fundingPlan.mode === "threshold" ? "Threshold " + money(p.fundingPlan.thresholdValue) : "Plan " + ((p.savedPlans || []).length + 1));
      if (!name) return;
      p.savedPlans = p.savedPlans || [];
      p.savedPlans.push({
        id: "pl_" + Date.now().toString(36),
        name: name,
        contributions: p.fundingPlan.contributions.slice(),
        assessments: (p.fundingPlan.assessments || []).slice()
      });
      refreshAfterEdit();
      toast('Saved "' + name + '" — shown as a line on the chart.');
    });
    var planLoad = $("#plan-load");
    if (planLoad) planLoad.addEventListener("change", function () {
      var pl = (p.savedPlans || []).find(function (x) { return x.id === planLoad.value; });
      if (!pl) return;
      p.fundingPlan.contributions = pl.contributions.slice();
      p.fundingPlan.assessments = pl.assessments.slice();
      p.fundingPlan.mode = "manual";
      p.fundingPlan.modeNote = "loaded from “" + pl.name + "”";
      p.fundingPlan._editedYears = {};
      refreshAfterEdit();
      toast('Loaded "' + pl.name + '" into the editor.');
    });
    var planBar = $(".plan-bar");
    if (planBar) planBar.addEventListener("click", function (e) {
      var del = e.target.closest(".plan-del");
      if (!del) return;
      var id = del.getAttribute("data-pid");
      p.savedPlans = (p.savedPlans || []).filter(function (x) { return x.id !== id; });
      delete hiddenSeries[id];
      refreshAfterEdit();
    });

    // Project filters
    var w = $("#pv-window");
    if (w) w.addEventListener("click", function (e) {
      var b = e.target.closest("button[data-w]");
      if (!b) return;
      propFilter.window = b.getAttribute("data-w");
      $all("#pv-window .fbtn").forEach(function (x) { x.classList.toggle("on", x === b); });
      drawProjectRows();
    });
    var se = $("#pv-search");
    if (se) se.addEventListener("input", debounce(function () {
      propFilter.search = se.value; drawProjectRows();
    }, 250));
    var cat = $("#pv-cat");
    if (cat) cat.addEventListener("change", function () {
      propFilter.category = cat.value; drawProjectRows();
    });
    var reset = $("#pv-reset");
    if (reset) reset.addEventListener("click", function () {
      propFilter = { window: "all", search: "", category: "all" };
      propSort = { key: null, dir: 1 };
      drawProperty();
    });

    // Sortable column headers
    var thead = $(".tbl.projects thead");
    if (thead) thead.addEventListener("click", function (e) {
      var th = e.target.closest("th.sortable");
      if (!th) return;
      var key = th.getAttribute("data-sort");
      if (propSort.key === key) propSort.dir *= -1;
      else { propSort.key = key; propSort.dir = 1; }
      // refresh header arrows + rows without full redraw
      $all(".tbl.projects th.sortable").forEach(function (h) {
        var k = h.getAttribute("data-sort");
        var on = k === propSort.key;
        h.classList.toggle("sorted", on);
        var arrow = h.querySelector(".arrow");
        if (arrow) arrow.textContent = on ? (propSort.dir === 1 ? "▲" : "▼") : "↕";
      });
      drawProjectRows();
    });

    // Project table inline edits (event delegation)
    var body = $("#proj-body");
    if (body) {
      body.addEventListener("change", function (e) {
        var input = e.target;
        if (input.tagName !== "INPUT") return;
        var cid = input.getAttribute("data-cid");
        var c = p.components.find(function (x) { return x.id === cid; });
        if (!c) return;
        if (input.classList.contains("cost")) {
          var v = parseMoney(input.value);
          if (v == null || v < 0) { input.value = moneyCents(c.unitCost); return; }
          c.unitCost = v;
        } else if (input.classList.contains("year")) {
          var y = parseInt(input.value, 10);
          if (isNaN(y)) { input.value = c.nextReplYear; return; }
          if (c.overrides) toast("Note: this item has hand-tuned year cells; those stay glued to their original years (same as Excel).");
          Engine.shiftTiming(c, y);
        } else if (input.classList.contains("ul")) {
          var ul = parseInt(input.value, 10);
          if (isNaN(ul) || ul < 1) { input.value = c.usefulLife; return; }
          c.usefulLife = ul;
        }
        refreshAfterEdit();
      });
      body.addEventListener("click", function (e) {
        var btn = e.target.closest("button[data-act]");
        if (!btn) return;
        var cid = btn.getAttribute("data-cid");
        var c = p.components.find(function (x) { return x.id === cid; });
        if (!c) return;
        var act = btn.getAttribute("data-act");
        if (act === "delete") { c.deleted = true; refreshAfterEdit(); }
        else if (act === "restore") { c.deleted = false; refreshAfterEdit(); }
        else if (act === "reset") { Object.assign(c, c.original); c.deleted = false; refreshAfterEdit(); }
        else if (act === "edit") openComponentModal(c);
      });
    }
  }

  /* =====================================================================
   * router
   * =================================================================== */
  function render() {
    var h = location.hash || "#/portfolio";
    var m = h.match(/^#\/property\/(.+)$/);
    if (m) { renderProperty(decodeURIComponent(m[1])); }
    else { currentProp = null; renderPortfolio(); }
  }
  window.addEventListener("hashchange", render);
  render();
})();
