/* =====================================================================
 * engine.js — Reserve funding recalculation engine
 * ---------------------------------------------------------------------
 * Pure functions. No DOM, no storage. Usable in browser and Node (tests).
 *
 * Formulas reverse-engineered from the Reserve Advisors funding model
 * workbook v7.0 ("The North Brook" reference copy):
 *
 * EXPENDITURE OCCURRENCE (Expenditures sheet, year columns X..):
 *   Let k    = year offset from Current Fiscal Year (0..studyLength)
 *       year = CFY + k
 *       F    = Next Full Replacement year (col F)
 *       G    = Useful Life / cycle years    (col G)
 *       D    = Length of Phase (years)      (col D)
 *       E    = Events Per Phase             (col E)
 *   A cost occurs at offset k when:
 *     (a) k == F - CFY                        (first full replacement), or
 *     (b) year > F and (year - F) mod G == 0  (subsequent full cycles), or
 *     (c) phased: D != 1 and k > F-CFY, with
 *           base = (F-CFY) + G*floor((k-(F-CFY))/G)
 *           step = ((D - E)/(E - 1)) + 1     (col C is display of this)
 *           rel  = k - base
 *         occurs when rel/step is an integer and rel < D
 *
 * EXPENDITURE AMOUNT:
 *   W = nearTermEndYear - CFY   (near-term inflation window, PropInfo B18)
 *   r = per-column rate (Expenditures row 7; = remaining-study inflation)
 *   if k <= W: amount = ROUND(L*R*S*(1+r)^k, 1)
 *   else:      amount = ROUND(L*R*S*(1+nearTerm)^W*(1+r)^(k-W), 1)
 *   where L = per-phase quantity, R = unit cost (today's $), S = %ownership
 *   L = (units == "Allowance") ? 1 : ROUND(K*B/E/H, 0)*H   (H = round phase)
 *
 * CASH FLOW (Funding Plan sheet):
 *   Current year (k=0):
 *     contrib0  = annualContribution * remainingPeriods/totalPeriods
 *     interest0 = ROUND((begin + exp0/2 + contrib0/2)
 *                        * rate * remainingInterestMonths/12, 0)
 *   Projection years (k>=1):
 *     interest  = ROUND(rate * (begin + exp/2 + contrib/2), 0)
 *   end = begin + contrib + interest + exp        (exp is negative)
 *
 * FUNDING SCENARIOS (ported from the validated reserve-stress-test skill):
 *   Baseline  = binary search on a uniform scale factor f applied to the
 *               recommended contributions (years >= firstRecYear), smallest
 *               f whose minimum year-end balance stays >= $0.
 *   Threshold = same search with floor = user's target dollar threshold
 *               (f may exceed 1). Monotonic in f, so 1-D search is valid.
 *   Contributions are rounded to property.roundTo (workbook Y62, e.g. $100).
 * ===================================================================== */

(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Engine = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // Excel ROUND (half away from zero; inputs here are non-negative in
  // practice, negative amounts handled for safety).
  function xlRound(x, digits) {
    var m = Math.pow(10, digits || 0);
    return (x < 0 ? -1 : 1) * Math.round(Math.abs(x) * m) / m;
  }

  function roundToNearest(x, step) {
    if (!step || step <= 0) return x;
    return Math.round(x / step) * step;
  }

  // Per-phase quantity, replicating Expenditures!L = ROUND(K*B/E/H,0)*H.
  // (The workbook wraps this in IF(M="Allowance",1,...); but for every
  // allowance in the reference study K*B/E/H rounds to exactly 1, so the
  // general formula reproduces all study values AND, unlike the hard "1",
  // splits the cost correctly when a user re-phases an allowance item —
  // e.g. a 1-unit allowance across 4 phases → 0.25 per phase. Verified
  // cell-for-cell by test/validate.js.)
  function perPhaseQuantity(c) {
    var H = c.roundPhase || 0.01;
    var E = c.eventsPerPhase || 1;
    return Math.round((c.totalQty * c.partialQty) / E / H) * H;
  }

  // Does component c incur a cost at offset k?
  function occursAt(c, k, cfy) {
    var fk = c.nextReplYear - cfy;
    if (k === fk) return true;
    var year = cfy + k;
    var G = c.usefulLife;
    if (G > 0 && year > c.nextReplYear && ((year - c.nextReplYear) % G) === 0) return true;
    // Phased events between cycle starts
    var D = c.phaseLength, E = c.eventsPerPhase;
    if (D !== 1 && E > 1 && k > fk && G > 0) {
      var base = fk + G * Math.floor((k - fk) / G);
      var step = ((D - E) / (E - 1)) + 1;
      var rel = k - base;
      if (rel < D && rel >= 0 && step > 0) {
        var q = rel / step;
        if (Math.abs(q - Math.round(q)) < 1e-9) return true;
      }
    }
    return false;
  }

  // Standard (formula) inflated cost of component c at offset k. ctx = property.
  function standardAmountAt(c, k, ctx) {
    var L = perPhaseQuantity(c);
    var baseCost = L * c.unitCost * (c.ownership == null ? 1 : c.ownership);
    var W = ctx.rates.nearTermEndYear - ctx.currentFiscalYear;
    var r = (ctx.colInflation && ctx.colInflation[k] != null)
      ? ctx.colInflation[k]
      : ctx.rates.remainingInflation;
    // Inflation override (UI rate control): replace both rates uniformly.
    if (ctx.rates.inflationOverride != null) {
      r = ctx.rates.inflationOverride;
      return xlRound(baseCost * Math.pow(1 + r, k), 1);
    }
    if (W >= k) return xlRound(baseCost * Math.pow(1 + r, k), 1);
    return xlRound(
      baseCost * Math.pow(1 + ctx.rates.nearTermInflation, W) * Math.pow(1 + r, k - W), 1);
  }

  // Cost at offset k, honoring per-cell manual overrides captured at parse
  // time (RA engineers hand-tune year cells: blank an occurrence, scale a
  // formula, or type a static number). Overrides are keyed by column offset
  // k — same semantics as Excel, where the edit is glued to the cell.
  //   c.overrides[k] = { t: "m", v: mult }  -> standard amount × mult
  //   c.overrides[k] = { t: "f", v: value } -> fixed dollar value
  function amountAt(c, k, ctx) {
    var ov = c.overrides && c.overrides[k];
    if (ov && ov.t === "f") return ov.v;
    var std = standardAmountAt(c, k, ctx);
    if (ov && ov.t === "m") return xlRound(std * ov.v, 1);
    return std;
  }

  /**
   * Derive per-cell overrides by comparing the workbook's cached year-cell
   * values (component.rawYearVals, captured by the parser) against the
   * engine's own computation at parse-time rates. Call once after parsing.
   * Returns a list of override descriptions for user-facing warnings.
   */
  function deriveOverrides(property) {
    var notes = [];
    var n = property.studyLength;
    property.components.forEach(function (c) {
      if (!c.rawYearVals) return;
      var ov = {};
      for (var k = 0; k <= n; k++) {
        var std = occursAt(c, k, property.currentFiscalYear)
          ? standardAmountAt(c, k, property) : 0;
        var raw = c.rawYearVals[k];
        raw = (typeof raw === "number") ? raw : 0;
        if (Math.abs(raw - std) <= 0.11) continue; // matches formula
        var isFormula = c.rawYearIsFormula ? !!c.rawYearIsFormula[k] : false;
        if (raw === 0 && std > 0) {
          ov[k] = { t: "m", v: 0 };
          notes.push(c.name + " (" + (property.currentFiscalYear + k) + "): occurrence manually removed in workbook.");
        } else if (std > 0 && isFormula) {
          ov[k] = { t: "m", v: raw / std };
          notes.push(c.name + " (" + (property.currentFiscalYear + k) + "): cost manually scaled to " +
            Math.round((raw / std) * 100) + "% of formula in workbook.");
        } else if (raw > 0) {
          ov[k] = { t: "f", v: raw };
          notes.push(c.name + " (" + (property.currentFiscalYear + k) + "): manual dollar entry in workbook ($" +
            raw + "); treated as fixed (will not scale with inflation edits).");
        }
      }
      if (Object.keys(ov).length) c.overrides = ov;
      // raw cells no longer needed at runtime; keep them out of storage
      delete c.rawYearVals;
      delete c.rawYearIsFormula;
    });
    return notes;
  }

  /**
   * Expenditure matrix for a property.
   * Returns { perComponent: [{component, amounts:[...len studyLength+1]}],
   *           totals: [...] } — amounts positive dollars, index = offset k.
   * Respects component.deleted and per-component field edits.
   */
  function computeExpenditures(property) {
    var n = property.studyLength; // offsets 0..n
    var totals = new Array(n + 1).fill(0);
    var perComponent = [];
    property.components.forEach(function (c) {
      if (c.deleted) return;
      var amounts = new Array(n + 1).fill(0);
      for (var k = 0; k <= n; k++) {
        var ov = c.overrides && c.overrides[k];
        var a = 0;
        if (ov && ov.t === "f") {
          // Manual dollar entry from the study — glued to this year even when
          // the recurrence formula places no event here.
          a = ov.v;
        } else if (occursAt(c, k, property.currentFiscalYear)) {
          a = amountAt(c, k, property);
        }
        if (a) { amounts[k] = a; totals[k] += a; }
      }
      perComponent.push({ component: c, amounts: amounts });
    });
    // Kill float dust from summation
    for (var k2 = 0; k2 <= n; k2++) totals[k2] = Math.round(totals[k2] * 100) / 100;
    return { perComponent: perComponent, totals: totals };
  }

  /**
   * Cash flow projection.
   * contributions: array length studyLength+1; index 0 = current-year ANNUAL
   * contribution (the engine applies remainingPeriods/totalPeriods to it),
   * indexes 1..n = full-year contributions.
   * expTotals: from computeExpenditures().totals (positive dollars).
   * assessments: optional array of manual one-off Additional Assessments per
   *   year (added in full, not prorated; included in the interest base like a
   *   contribution). Defaults to property.fundingPlan.assessments, else 0.
   * Returns rows with { contribution, assessment, interest, expenditures, end }
   * where `contribution` is the base (prorated) contribution and the year-end
   * balance already includes the assessment.
   */
  function computeCashFlow(property, contributions, expTotals, assessments) {
    var n = property.studyLength;
    var rate = property.rates.interest;
    var cur = property.current;
    var assess = assessments ||
      (property.fundingPlan && property.fundingPlan.assessments) || [];
    var rows = [];
    var begin = cur.beginningBalance;
    for (var k = 0; k <= n; k++) {
      var exp = -(expTotals[k] || 0);
      var addl = assess[k] || 0;
      var contrib, interest;
      if (k === 0) {
        contrib = (contributions[0] || 0) * (cur.remainingPeriods / cur.totalPeriods);
      } else {
        contrib = contributions[k] || 0;
      }
      var totalIn = contrib + addl;
      if (k === 0) {
        interest = xlRound(
          (begin + exp / 2 + totalIn / 2) * rate * (cur.remainingInterestMonths / 12), 0);
      } else {
        interest = xlRound(rate * (begin + exp / 2 + totalIn / 2), 0);
      }
      var end = begin + totalIn + interest + exp;
      rows.push({
        year: property.currentFiscalYear + k, k: k, begin: begin,
        contribution: contrib, assessment: addl, interest: interest,
        expenditures: exp, end: end
      });
      begin = end;
    }
    return rows;
  }

  function minEndBalance(rows) {
    return rows.reduce(function (m, r) { return Math.min(m, r.end); }, Infinity);
  }

  /**
   * Uniform-scale search (ported from reserve-stress-test skill Step 2).
   * Scales recommended contributions for years >= firstRecYear by f,
   * rounding each to property.roundTo; current-year contribution is fixed
   * (it comes from the current budget, not the levers — same rationale as
   * the workbook's I20:T23 block applying from the first non-budgeted year).
   *
   * Returns smallest f in [0, fMax] where min year-end balance >= floor,
   * plus the resulting schedule. Monotonic increasing in f => binary search.
   */
  function solveScaledSchedule(property, expTotals, floor, fMax) {
    var rec = property.fundingPlan.recommended;
    var firstK = property.firstRecYear - property.currentFiscalYear;

    function schedule(f) {
      var s = rec.slice();
      for (var k = Math.max(firstK, 1); k < s.length; k++) {
        s[k] = roundToNearest(rec[k] * f, property.roundTo);
      }
      return s;
    }
    function minBal(f) {
      return minEndBalance(computeCashFlow(property, schedule(f), expTotals));
    }

    // Feasibility at fMax (expand up to a hard cap for threshold searches)
    var hi = fMax || 1;
    var HARD_CAP = 50;
    while (minBal(hi) < floor && hi < HARD_CAP) hi *= 2;
    if (minBal(hi) < floor) {
      return { feasible: false, f: hi, schedule: schedule(hi),
               note: "No uniform scaling of the recommended plan reaches the target floor." };
    }
    // Guardrail (skill Step 2): if even f=0 satisfies the floor, the plan is
    // carried by its structural floor (starting balance + current-year path).
    if (minBal(0) >= floor) {
      return { feasible: true, f: 0, schedule: schedule(0), atFloorZero: true,
               note: "Balance stays above the floor even with contributions scaled to zero " +
                     "(from the first recommendation year onward). Current-year contributions " +
                     "and starting balance alone carry the plan." };
    }
    var lo = 0; // infeasible
    for (var i = 0; i < 45; i++) {
      var mid = (lo + hi) / 2;
      if (minBal(mid) >= floor) hi = mid; else lo = mid;
    }
    return { feasible: true, f: hi, schedule: schedule(hi) };
  }

  function baselineSchedule(property, expTotals) {
    return solveScaledSchedule(property, expTotals, 0, 1);
  }

  function thresholdSchedule(property, expTotals, floorDollars) {
    return solveScaledSchedule(property, expTotals, floorDollars, 1);
  }

  /* ------------------------------------------------------------------
   * Component editing helpers (used by the UI; kept here so edits and
   * recalc share one definition of the mechanics).
   * ---------------------------------------------------------------- */

  // Shift a component's series so its next occurrence lands on targetYear.
  function shiftTiming(component, targetYear) {
    component.nextReplYear = targetYear;
  }

  // Phase a component across `phases` events spaced `spacing` years apart.
  // Writes the workbook's D/E parameters; per-phase quantity recomputes
  // automatically (equal split via ROUND(K*B/E/H)*H).
  function phaseComponent(component, phases, spacing) {
    if (phases <= 1) { // un-phase
      component.phaseLength = 1;
      component.eventsPerPhase = 1;
      component.frequencyOfEvents = 1;
      return;
    }
    spacing = Math.max(1, spacing || 1);
    component.eventsPerPhase = phases;                    // E
    component.phaseLength = spacing * (phases - 1) + 1;   // D  => step == spacing
    component.frequencyOfEvents = spacing;                // C (display)
  }

  return {
    xlRound: xlRound,
    roundToNearest: roundToNearest,
    perPhaseQuantity: perPhaseQuantity,
    occursAt: occursAt,
    amountAt: amountAt,
    standardAmountAt: standardAmountAt,
    deriveOverrides: deriveOverrides,
    computeExpenditures: computeExpenditures,
    computeCashFlow: computeCashFlow,
    minEndBalance: minEndBalance,
    baselineSchedule: baselineSchedule,
    thresholdSchedule: thresholdSchedule,
    solveScaledSchedule: solveScaledSchedule,
    shiftTiming: shiftTiming,
    phaseComponent: phaseComponent
  };
});
