/* Node validation harness: parse the North Brook workbook with the vendored
 * SheetJS + parser, run the engine, and compare against the workbook's own
 * cached values (fixtures-northbrook.json extracted via openpyxl).
 * Run: node test/validate.js <path-to-xlsx>
 */
"use strict";
var fs = require("fs");
var path = require("path");
var XLSX = require(path.join(__dirname, "..", "vendor", "xlsx.full.min.js"));
var Parser = require(path.join(__dirname, "..", "js", "parser.js"));
var Engine = require(path.join(__dirname, "..", "js", "engine.js"));

var xlsxPath = process.argv[2];
var fixtures = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures-northbrook.json"), "utf8"));

var wb = XLSX.read(fs.readFileSync(xlsxPath), { type: "buffer", cellDates: false });
var res = Parser.parse(wb);
var p = res.property;
var overrideNotes = Engine.deriveOverrides(p);
console.log("Override notes:", overrideNotes.length ? overrideNotes : "(none)");

var fails = 0, checks = 0;
function ok(cond, msg) {
  checks++;
  if (!cond) { fails++; console.log("FAIL:", msg); }
}
function close(a, b, tol, msg) {
  ok(Math.abs(a - b) <= tol, msg + " (got " + a + ", want " + b + ")");
}

console.log("Warnings:", res.warnings.length ? res.warnings : "(none)");
console.log("Property:", p.name, "| CFY", p.currentFiscalYear, "| study", p.studyLength,
  "| components", p.components.length);

// --- 1. Parse-level checks ---
ok(p.components.length === fixtures.ex.components.length,
  "component count " + p.components.length + " vs " + fixtures.ex.components.length);
close(p.current.beginningBalance, fixtures.fp.settings.Y54, 0.01, "beginning balance");
close(p.rates.interest, fixtures.propertyInfo.B20, 1e-9, "interest rate");
close(p.rates.nearTermInflation, fixtures.propertyInfo.B17, 1e-9, "near-term inflation");
ok(p.roundTo === fixtures.fp.settings.Y62, "roundTo");
ok(p.current.remainingPeriods === fixtures.fp.settings.Y60, "remaining periods");

// Component params vs fixtures (match by row)
var fxByRow = {};
fixtures.ex.components.forEach(function (c) { fxByRow[c.row] = c; });
p.components.forEach(function (c) {
  var f = fxByRow[c.row];
  ok(!!f, "fixture for row " + c.row);
  if (!f) return;
  ok(c.name === f.name, "row " + c.row + " name");
  close(c.unitCost, f.R, 1e-9, "row " + c.row + " unitCost");
  close(c.nextReplYear, f.F, 0, "row " + c.row + " nextReplYear");
  close(c.usefulLife, f.G, 0, "row " + c.row + " usefulLife");
  close(Engine.perPhaseQuantity(c), f.L, 0.02, "row " + c.row + " perPhaseQty");
});

// --- 2. Expenditure matrix vs workbook year cells ---
var em = Engine.computeExpenditures(p);
var emByRow = {};
em.perComponent.forEach(function (pc) { emByRow[pc.component.row] = pc.amounts; });
var cellFails = 0, cellChecks = 0;
fixtures.ex.components.forEach(function (f) {
  var amounts = emByRow[f.row];
  if (!amounts) return;
  for (var k = 0; k <= p.studyLength; k++) {
    var want = f.yearVals[k];
    want = (typeof want === "number") ? want : 0;
    cellChecks++;
    if (Math.abs((amounts[k] || 0) - want) > 0.11) {
      cellFails++;
      if (cellFails <= 12) console.log("FAIL cell: row " + f.row + " '" + f.name +
        "' k=" + k + " got " + amounts[k] + " want " + want);
    }
  }
});
console.log("Expenditure cells: " + (cellChecks - cellFails) + "/" + cellChecks + " match");
ok(cellFails === 0, "all expenditure cells match");

// Yearly totals vs Funding Plan expenditure rows
var fpExp = fixtures.fp.exp1.slice(0, 16).concat(fixtures.fp.exp2.slice(0, 15)); // 2026..2056
for (var k = 0; k <= p.studyLength; k++) {
  var want = -(typeof fpExp[k] === "number" ? fpExp[k] : 0);
  close(em.totals[k], want, 0.51, "year total k=" + k);
}

// --- 3. Cash flow vs workbook ---
var rows = Engine.computeCashFlow(p, p.fundingPlan.contributions, em.totals);
var fpEnd = fixtures.fp.end1.slice(0, 16).concat(fixtures.fp.end2.slice(0, 15));
var fpContrib = fixtures.fp.contrib1.slice(0, 16).concat(fixtures.fp.contrib2.slice(0, 15));
var fpInt = fixtures.fp.interest1.slice(0, 16).concat(fixtures.fp.interest2.slice(0, 15));
for (var k2 = 0; k2 <= p.studyLength; k2++) {
  close(rows[k2].contribution, fpContrib[k2], 0.51, "contribution k=" + k2);
  close(rows[k2].interest, fpInt[k2], 1.01, "interest k=" + k2);
  close(rows[k2].end, fpEnd[k2], 2.0, "end balance k=" + k2 + " (year " + rows[k2].year + ")");
}

// Recommended funding table balances (parser's _expectedEndBalances)
Object.keys(p._expectedEndBalances).forEach(function (yr) {
  var row = rows.find(function (r) { return r.year === +yr; });
  close(row.end, p._expectedEndBalances[yr], 2.0, "funding-table balance " + yr);
});

// --- 4. Edit-recalc sanity checks ---
// (a) unit cost bump raises expenditures in the component's years only
var c0 = p.components.find(function (c) { return c.name.indexOf("Asphalt Pavement, Crack") === 0; });
var before = Engine.computeExpenditures(p).totals.slice();
var origCost = c0.unitCost;
c0.unitCost = origCost * 2;
var after = Engine.computeExpenditures(p).totals;
var changed = 0;
for (var k3 = 0; k3 <= p.studyLength; k3++) if (Math.abs(after[k3] - before[k3]) > 0.01) changed++;
ok(changed > 0, "unit-cost edit changes totals");
c0.unitCost = origCost;

// (b) timing shift moves cost
var f0 = c0.nextReplYear;
Engine.shiftTiming(c0, f0 + 1);
var shifted = Engine.computeExpenditures(p).totals;
ok(Math.abs(shifted[f0 - p.currentFiscalYear] - before[f0 - p.currentFiscalYear]) > 0.01,
  "timing shift removes cost from original year");
Engine.shiftTiming(c0, f0);

// (c) delete removes all of component's costs
c0.deleted = true;
var delTotals = Engine.computeExpenditures(p).totals;
var emc = emByRow[c0.row];
for (var k4 = 0; k4 <= p.studyLength; k4++) {
  if (emc[k4] > 0) { close(delTotals[k4], before[k4] - emc[k4], 0.2, "delete year k=" + k4); }
}
c0.deleted = false;

// (d) phasing: split a one-shot project into 3 phases 1yr apart => 3 events
var solar = p.components.find(function (c) { return c.name.indexOf("Solar Photovoltaic System, East") === 0; });
var snap = JSON.parse(JSON.stringify(solar));
Engine.phaseComponent(solar, 3, 1);
var pem = Engine.computeExpenditures(p);
var pAmounts = pem.perComponent.find(function (pc) { return pc.component.row === solar.row; }).amounts;
var events = pAmounts.filter(function (a, i) { return a > 0 && i <= 4; }).length;
ok(events === 3, "phasing produced 3 near-term events (got " + events + ")");
Object.assign(solar, snap);

// --- 5. Scenario solvers ---
var expT = Engine.computeExpenditures(p).totals;
var base = Engine.baselineSchedule(p, expT);
ok(base.feasible, "baseline feasible");
var baseRows = Engine.computeCashFlow(p, base.schedule, expT);
var mb = Engine.minEndBalance(baseRows);
ok(mb >= 0 && mb < 20000, "baseline min balance small positive (got " + Math.round(mb) + ", f=" + base.f.toFixed(4) + ")");

var thr = Engine.thresholdSchedule(p, expT, 250000);
var thrRows = Engine.computeCashFlow(p, thr.schedule, expT);
var mt = Engine.minEndBalance(thrRows);
ok(thr.feasible && mt >= 250000 && mt < 270000, "threshold(250k) min balance just above floor (got " + Math.round(mt) + ")");

console.log("\n" + (checks - fails) + "/" + checks + " checks passed" + (fails ? " — " + fails + " FAILURES" : " OK"));
process.exit(fails ? 1 : 0);
