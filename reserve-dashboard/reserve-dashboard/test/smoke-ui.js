/* jsdom UI smoke test — boots the app, imports the North Brook workbook via
 * the same code path as a browser upload (minus FileReader), and drives the
 * main interactions: portfolio render, property view, scenario buttons,
 * manual contribution edit, project edits, filters.
 * Run: node test/smoke-ui.js <path-to-xlsx>   (requires jsdom resolvable)
 */
"use strict";
var fs = require("fs");
var path = require("path");
var { JSDOM } = require("jsdom");

var ROOT = path.join(__dirname, "..");
var xlsxPath = process.argv[2];

var html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8")
  .replace(/<script[^>]*src="[^"]*"[^>]*><\/script>/g, ""); // load scripts manually

var dom = new JSDOM(html, { url: "http://localhost/#/portfolio", runScripts: "outside-only", pretendToBeVisual: true });
var w = dom.window;

// minimal shims
if (!w.localStorage) {
  var mem = {};
  w.localStorage = {
    getItem: function (k) { return k in mem ? mem[k] : null; },
    setItem: function (k, v) { mem[k] = String(v); },
    removeItem: function (k) { delete mem[k]; }
  };
}
w.HTMLDialogElement.prototype.showModal = w.HTMLDialogElement.prototype.showModal || function () { this.open = true; };
w.HTMLDialogElement.prototype.close = w.HTMLDialogElement.prototype.close || function (v) { this.returnValue = v; this.open = false; this.dispatchEvent(new w.Event("close")); };
w.confirm = function () { return true; };
w.prompt = function () { return "Test plan"; };

function load(file) {
  var code = fs.readFileSync(path.join(ROOT, file), "utf8");
  w.eval(code);
}
load("vendor/xlsx.full.min.js");
load("js/engine.js");
load("js/parser.js");
load("js/store.js");
load("js/charts.js");
load("js/app.js");

var fails = 0, checks = 0;
function ok(cond, msg) { checks++; if (!cond) { fails++; console.log("FAIL:", msg); } }
function $(sel) { return w.document.querySelector(sel); }
function $$(sel) { return Array.prototype.slice.call(w.document.querySelectorAll(sel)); }
function fire(el, type) { el.dispatchEvent(new w.Event(type, { bubbles: true })); }
function click(el) { el.dispatchEvent(new w.MouseEvent("click", { bubbles: true })); }
function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

(async function main() {
  await sleep(50);
  ok($("#uz-portfolio"), "portfolio upload zone renders on empty state");

  // Import workbook through the parser (same as upload handler, minus FileReader)
  var wb = w.XLSX.read(fs.readFileSync(xlsxPath), { type: "buffer" });
  var res = w.Parser.parse(wb);
  var p = res.property;
  p.overrideNotes = w.Engine.deriveOverrides(p);
  p.parseWarnings = res.warnings;
  await w.Store.create().saveProperty(p);

  // Portfolio view
  w.location.hash = "#/portfolio";
  fire(w, "hashchange");
  await sleep(80);
  ok($$(".pcard").length === 1, "property card renders");
  ok($(".pcard-name").textContent.indexOf("North Brook") !== -1, "card shows association name");
  ok($$("#pf-window .fbtn").length === 6, "portfolio window filter buttons");

  // Search filter in portfolio
  var se = $("#pf-search");
  se.value = "roof";
  fire(se, "input");
  await sleep(500);
  var rows = $$(".tbl tbody tr");
  ok(rows.length >= 1, "portfolio search returns rows");
  ok(rows.every(function (r) { return /roof/i.test(r.textContent) || /No projects/.test(r.textContent); }),
    "portfolio search rows all match 'roof'");
  w.eval("void 0"); // noop

  // Property view
  w.location.hash = "#/property/" + p.id;
  fire(w, "hashchange");
  await sleep(80);
  ok($(".property h1").textContent.indexOf("North Brook") !== -1, "property view renders");
  ok($$(".kpi").length === 4, "4 KPI cards");
  ok($("#chart svg"), "chart SVG renders");
  ok($$("#contrib-grid input.contrib").length === p.studyLength, "30 contribution inputs");
  ok($$("#contrib-grid input.assess").length === p.studyLength, "30 additional-assessment inputs");
  ok($$("#proj-body tr").length === p.components.length, "all components listed");
  var noteEl = $(".notes summary");
  ok(noteEl && noteEl.textContent.indexOf("6") !== -1, "6 study notes surfaced (overrides)");
  ok($$("#proj-body .badge.edit").length === 0, "no spurious 'edited' badges on a freshly loaded study");

  // Baseline scenario
  click($("#sc-base"));
  await sleep(80);
  ok($("#mode-note") && /scaled to/.test($("#mode-note").textContent), "baseline note shows scale factor");
  var store = w.Store.create();
  var saved = await store.getProperty(p.id);
  await sleep(600); // allow debounced save
  saved = await store.getProperty(p.id);
  ok(saved.fundingPlan.mode === "baseline", "baseline mode persisted");
  var em = w.Engine.computeExpenditures(saved);
  var cf = w.Engine.computeCashFlow(saved, saved.fundingPlan.contributions, em.totals);
  var minB = w.Engine.minEndBalance(cf);
  ok(minB >= 0 && minB < 20000, "baseline min balance small positive (" + Math.round(minB) + ")");

  // Threshold scenario
  $("#thr-value").value = "250000";
  click($("#sc-thr"));
  await sleep(700);
  saved = await store.getProperty(p.id);
  cf = w.Engine.computeCashFlow(saved, saved.fundingPlan.contributions,
    w.Engine.computeExpenditures(saved).totals);
  ok(w.Engine.minEndBalance(cf) >= 250000, "threshold floor respected");

  // Manual contribution edit
  var ci = $("#contrib-grid input[data-k='5']");
  ci.value = "123400";
  fire(ci, "change");
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.fundingPlan.mode === "manual", "manual mode after grid edit");
  ok(saved.fundingPlan.contributions[5] === 123400, "manual value stored");

  // Project edit: unit cost
  var costIn = $("#proj-body input.cost");
  var cid = costIn.getAttribute("data-cid");
  costIn.value = "9.99";
  fire(costIn, "change");
  await sleep(700);
  saved = await store.getProperty(p.id);
  var comp = saved.components.find(function (c) { return c.id === cid; });
  ok(comp.unitCost === 9.99, "unit cost edit persisted");

  // Project edit: timing shift
  var yearIn = $("#proj-body input.year");
  cid = yearIn.getAttribute("data-cid");
  var newYear = parseInt(yearIn.value, 10) + 2;
  yearIn.value = String(newYear);
  fire(yearIn, "change");
  await sleep(700);
  saved = await store.getProperty(p.id);
  comp = saved.components.find(function (c) { return c.id === cid; });
  ok(comp.nextReplYear === newYear, "timing shift persisted");

  // Delete + restore
  var delBtn = $("#proj-body button[data-act='delete']");
  cid = delBtn.getAttribute("data-cid");
  click(delBtn);
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.components.find(function (c) { return c.id === cid; }).deleted === true, "delete persisted");
  var restoreBtn = $("#proj-body button[data-act='restore']");
  click(restoreBtn);
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.components.find(function (c) { return c.id === cid; }).deleted === false, "restore persisted");

  // Filters: This year + search combinable
  click($("#pv-window button[data-w='next5']"));
  await sleep(50);
  var pvSearch = $("#pv-search");
  pvSearch.value = "pool";
  fire(pvSearch, "input");
  await sleep(400);
  var prows = $$("#proj-body tr").filter(function (r) { return !/No projects/.test(r.textContent); });
  ok(prows.length > 0, "next-5-years + 'pool' search returns rows");
  ok(prows.every(function (r) { return /pool/i.test(r.textContent); }), "filtered rows match search");

  // Assumptions via editable KPI cards: 4 cards, three with edit buttons
  ok($$(".kpi").length === 4, "four KPI cards (starting reserves, low point, interest, inflation)");
  ok($("#edit-current") && $("#edit-interest") && $("#edit-inflation"), "KPI edit buttons present");
  ok(!$(".assume"), "old Assumptions section removed");
  ok($("#btn-reset-assume"), "reset-to-study button present near KPIs");

  // Inflation rate modal → override lowers min balance
  saved = await store.getProperty(p.id);
  var beforeMin = w.Engine.minEndBalance(w.Engine.computeCashFlow(saved,
    saved.fundingPlan.contributions, w.Engine.computeExpenditures(saved).totals, saved.fundingPlan.assessments));
  click($("#edit-inflation"));
  await sleep(50);
  ok($("#mm-back"), "inflation modal opens");
  $("#rm-val").value = "5.0";
  click($("#rm-save"));
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.rates.inflationOverride === 0.05, "inflation override stored via modal");
  var afterMin = w.Engine.minEndBalance(w.Engine.computeCashFlow(saved,
    saved.fundingPlan.contributions, w.Engine.computeExpenditures(saved).totals, saved.fundingPlan.assessments));
  ok(afterMin < beforeMin, "higher inflation lowers min balance (" +
    Math.round(beforeMin) + " -> " + Math.round(afterMin) + ")");

  // Current-reserves modal: edit balance, see recomputed year-end explanation
  click($("#edit-current"));
  await sleep(50);
  ok($("#cm-balance") && $("#cm-date") && $("#cm-contrib"), "current modal has balance/date/contribution");
  ok(/monthly/i.test($("#cm-explain").innerHTML) && /year-end/i.test($("#cm-explain").innerHTML),
    "current modal explains the monthly-contribution year-end calc");
  var studyDate = p.current.balanceDate;
  $("#cm-balance").value = "500000"; fire($("#cm-balance"), "input");
  click($("#cm-save"));
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.current.beginningBalance === 500000, "current modal saved new beginning balance");
  ok(saved.current.balanceDate === studyDate, "balance date preserved after editing balance");

  // Reset to study restores balance AND keeps the balance date (regression)
  click($("#btn-reset-assume"));
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.current.balanceDate === studyDate, "reset keeps the study balance date (not blank)");
  ok(saved.current.beginningBalance === p._origAssume.balance, "reset restores study beginning balance");

  // ---- style-guide reskin + new features ----

  // Currency formatting on dollar inputs
  ok(/^\$[\d,]/.test($("#proj-body input.cost").value), "unit cost shows $ format (" + $("#proj-body input.cost").value + ")");

  // Funding ledger layout
  ok($("#contrib-grid").classList.contains("ledger"), "funding plan renders as ledger");
  ok($$("#contrib-grid .ledger-row").length === p.studyLength, "ledger has one row per projection year");
  ok($("#contrib-grid .bal"), "ledger shows projected balance column");

  // Reset filters to a clean table
  click($("#pv-window button[data-w='all']"));
  var pvs = $("#pv-search"); pvs.value = ""; fire(pvs, "input");
  await sleep(350);

  // Sortable Project header
  var nameTh = $(".tbl.projects th[data-sort='name']");
  ok(nameTh, "Project header is sortable");
  click(nameTh);
  await sleep(50);
  var names = $$("#proj-body .cname").map(function (n) { return n.textContent.toLowerCase(); });
  var sortedAsc = names.slice().sort();
  ok(JSON.stringify(names) === JSON.stringify(sortedAsc), "clicking Project sorts rows A→Z");

  // Category filter
  var catSel = $("#pv-cat");
  ok(catSel && catSel.options.length > 1, "category filter populated");
  var pick = catSel.options[1].value;
  catSel.value = pick; fire(catSel, "change");
  await sleep(50);
  var catRows = $$("#proj-body tr").filter(function (r) { return !/No projects/.test(r.textContent); });
  ok(catRows.length > 0 && catRows.every(function (r) { return r.children[1].textContent === pick; }),
    "category filter shows only '" + pick + "' rows");
  catSel.value = "all"; fire(catSel, "change");
  await sleep(50);

  // Inline useful-life edit
  var ulIn = $("#proj-body input.ul");
  var ulCid = ulIn.getAttribute("data-cid");
  ulIn.value = "17";
  fire(ulIn, "change");
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.components.find(function (c) { return c.id === ulCid; }).usefulLife === 17, "inline useful-life edit persisted");

  // Component edit modal
  var editBtn = $("#proj-body button[data-act='edit']");
  var editCid = editBtn.getAttribute("data-cid");
  click(editBtn);
  await sleep(60);
  ok($("#cmodal-back"), "edit modal opens");
  ok($("#m-future") && /^\$/.test($("#m-future").textContent), "modal shows Total Future Cost");
  ok($("#m-chart svg"), "modal shows forecast chart");
  var mUl = $("#m-ul");
  mUl.value = "23"; fire(mUl, "input");
  await sleep(30);
  click($("#cmodal-save"));
  await sleep(700);
  ok(!$("#cmodal-back"), "modal closes on save");
  saved = await store.getProperty(p.id);
  ok(saved.components.find(function (c) { return c.id === editCid; }).usefulLife === 23, "modal edit persisted useful life");

  // ---- second-round tweaks ----

  // Additional assessment column drives the balance
  var assessIn = $("#contrib-grid input.assess[data-ak='3']");
  ok(assessIn, "ledger has additional-assessment inputs");
  var balBefore = parseFloat($("#contrib-grid .bal[data-balk='3']").textContent.replace(/[^0-9.-]/g, ""));
  assessIn.value = "50000"; fire(assessIn, "change");
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.fundingPlan.assessments[3] === 50000, "assessment value persisted");
  var em3 = w.Engine.computeExpenditures(saved);
  var cf3 = w.Engine.computeCashFlow(saved, saved.fundingPlan.contributions, em3.totals, saved.fundingPlan.assessments);
  ok(Math.round(cf3[3].end - cf3[3].begin - cf3[3].contribution - cf3[3].assessment - cf3[3].interest - cf3[3].expenditures) === 0,
    "cash-flow identity holds with assessment (begin+contrib+assess+interest+exp = end)");
  ok(cf3[3].assessment === 50000, "assessment appears in cash-flow row");
  // clear it back
  assessIn = $("#contrib-grid input.assess[data-ak='3']"); assessIn.value = ""; fire(assessIn, "change");
  await sleep(500);

  // Save a funding plan → appears as a second chart line + legend item
  click($("#plan-save"));
  await sleep(700);
  saved = await store.getProperty(p.id);
  ok(saved.savedPlans && saved.savedPlans.length === 1, "funding plan saved");
  ok($$("#chart-legend .leg-item").length === 2, "chart legend shows working + saved plan");
  ok($("#chart svg").querySelectorAll("polyline").length === 2, "chart draws two plan lines");

  // Legend toggle hides a line
  var legItem = $$("#chart-legend .leg-item")[1];
  click(legItem);
  await sleep(40);
  ok($$("#chart-legend .leg-item")[1].classList.contains("off"), "legend item marked off after click");
  ok($("#chart svg").querySelectorAll("polyline").length === 1, "toggled line removed from chart");
  click($$("#chart-legend .leg-item")[1]); // toggle back on
  await sleep(40);

  // Chart hover tooltip
  var col = $(".hovercol");
  var mm = new w.MouseEvent("mousemove", { bubbles: true, clientX: 120, clientY: 120 });
  col.dispatchEvent(mm);
  await sleep(20);
  var tip = $("#chart-tip");
  ok(tip.style.display === "block", "hover tooltip shows");
  ok(/Contribution/.test(tip.innerHTML) && /Year-end balance/.test(tip.innerHTML) && /Projects/.test(tip.innerHTML),
    "tooltip lists contribution, balance, and projects");

  // Projects reset-view button
  click($("#pv-window button[data-w='next5']"));
  var s2 = $("#pv-search"); s2.value = "roof"; fire(s2, "input");
  await sleep(350);
  click($("#pv-reset"));
  await sleep(60);
  ok($("#pv-search").value === "" && $("#pv-window button[data-w='all']").classList.contains("on"),
    "Reset view clears search + window filter");

  // Edit modal field trim + order + phasing keeps total cost (equal split)
  var editBtn2 = $("#proj-body button[data-act='edit']");
  click(editBtn2);
  await sleep(60);
  ok(!$("#m-perphase") && !$("#m-units") && !$("#m-remain"), "modal drops qty-per-phase, units, remaining life");
  var firstLabel = $(".field-grid .field label").textContent;
  ok(firstLabel === "Total Quantity", "first modal field is Total Quantity");
  ok(/Percentage \(%\)/.test($(".field-grid").innerHTML) && !/Ownership/.test($(".field-grid").innerHTML),
    "percentage field renamed (no 'Ownership')");
  click($("#cmodal-cancel"));

  // Phasing bug fix: an allowance item split into N phases must divide the
  // cost (not repeat the full cost N times). Use a single-occurrence
  // allowance so window timing doesn't confound the comparison.
  saved = await store.getProperty(p.id);
  var studyEnd = saved.currentFiscalYear + saved.studyLength;
  var allow = saved.components.find(function (c) {
    return /allowance/i.test(c.units || "") && (c.eventsPerPhase || 1) === 1 &&
      c.nextReplYear + c.usefulLife > studyEnd &&           // no second cycle in window
      c.nextReplYear + 3 <= studyEnd;                        // room for 3 phases
  });
  ok(allow, "found a single-occurrence allowance to test phasing");
  if (allow) {
    var emA = w.Engine.computeExpenditures(saved);
    var pcA = emA.perComponent.find(function (x) { return x.component.id === allow.id; });
    var totalBefore = pcA.amounts.reduce(function (s, v) { return s + v; }, 0);
    var clone = JSON.parse(JSON.stringify(saved));
    var cc = clone.components.find(function (x) { return x.id === allow.id; });
    w.Engine.phaseComponent(cc, 3, 1);
    var emB = w.Engine.computeExpenditures(clone);
    var pcB = emB.perComponent.find(function (x) { return x.component.id === allow.id; });
    var phased = pcB.amounts.filter(function (a) { return a > 0; });
    var totalAfter = phased.reduce(function (s, v) { return s + v; }, 0);
    ok(phased.length === 3, "allowance splits into 3 phased events (got " + phased.length + ")");
    ok(Math.abs(totalAfter - totalBefore) / Math.max(1, totalBefore) < 0.05,
      "phased allowance total ≈ unchanged (equal split, not ×3): " + Math.round(totalBefore) + " → " + Math.round(totalAfter));
  }

  // Coordination card click filters the portfolio project list.
  // Add a second property (same workbook) so a coordination group exists.
  var wb2 = w.XLSX.read(fs.readFileSync(xlsxPath), { type: "buffer" });
  var res2 = w.Parser.parse(wb2, { id: "p2" });
  var p2 = res2.property; p2.name = "Second Property";
  p2.overrideNotes = w.Engine.deriveOverrides(p2);
  await w.Store.create().saveProperty(p2);
  w.location.hash = "#/portfolio";
  fire(w, "hashchange");
  await sleep(120);
  var coordCard = $(".coord-card");
  ok(coordCard, "coordination card appears with two properties");
  if (coordCard) {
    var ckey = coordCard.getAttribute("data-key");
    var cfrom = +coordCard.getAttribute("data-from"), cto = +coordCard.getAttribute("data-to");
    click(coordCard);
    await sleep(120);
    ok($(".active-filter"), "active coordination filter chip shows");
    var crows = $$(".portfolio .tbl tbody tr").filter(function (r) { return !/No projects/.test(r.textContent); });
    ok(crows.length > 0, "coordination filter returns rows");
    var allMatch = crows.every(function (r) {
      var proj = r.children[1].textContent.split(",")[0].trim().toLowerCase();
      var yr = +r.children[3].textContent;
      return proj === ckey && yr >= cfrom && yr <= cto;
    });
    ok(allMatch, "filtered rows all match component family '" + ckey + "' AND fall within " + cfrom + "–" + cto);
    // toggle off
    click($("#coord-clear"));
    await sleep(80);
    ok(!$(".active-filter"), "clear chip removes the coordination filter");
  }

  console.log("\n" + (checks - fails) + "/" + checks + " UI checks passed" + (fails ? " — " + fails + " FAILURES" : " OK"));
  process.exit(fails ? 1 : 0);
})().catch(function (e) { console.error(e); process.exit(1); });
