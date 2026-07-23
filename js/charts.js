/* =====================================================================
 * charts.js — tiny dependency-free SVG chart for balance + expenditures
 * (no CDN/runtime deps per build spec; adequate for POC)
 * ===================================================================== */

(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Charts = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  function fmtShort(v) {
    var a = Math.abs(v);
    if (a >= 1e6) return (v / 1e6).toFixed(1).replace(/\.0$/, "") + "M";
    if (a >= 1e3) return Math.round(v / 1e3) + "k";
    return String(Math.round(v));
  }

  /**
   * Render a combined chart: bars = annual expenditures, line = year-end
   * balance, optional dashed line = threshold floor.
   * rows: engine cash flow rows. Returns SVG markup string.
   */
  function cashFlowChart(rows, opts) {
    opts = opts || {};
    var W = opts.width || 920, H = opts.height || 300;
    var padL = 58, padR = 14, padT = 16, padB = 30;
    var iw = W - padL - padR, ih = H - padT - padB;
    var n = rows.length;
    var maxBal = Math.max.apply(null, rows.map(function (r) { return r.end; }).concat([0]));
    var minBal = Math.min.apply(null, rows.map(function (r) { return r.end; }).concat([0]));
    var maxExp = Math.max.apply(null, rows.map(function (r) { return -r.expenditures; }).concat([1]));
    if (opts.threshold != null) maxBal = Math.max(maxBal, opts.threshold);
    var lo = Math.min(minBal, 0), hi = Math.max(maxBal, maxExp) * 1.06;
    if (hi === lo) hi = lo + 1;
    function y(v) { return padT + ih - ((v - lo) / (hi - lo)) * ih; }
    function x(i) { return padL + (i + 0.5) * (iw / n); }
    var bw = Math.max(3, (iw / n) - 4);

    var s = '<svg viewBox="0 0 ' + W + " " + H + '" class="cfchart" role="img" aria-label="Reserve balance and expenditures by year">';
    // gridlines
    var ticks = 4;
    for (var t = 0; t <= ticks; t++) {
      var v = lo + (hi - lo) * (t / ticks);
      s += '<line x1="' + padL + '" y1="' + y(v) + '" x2="' + (W - padR) + '" y2="' + y(v) + '" class="grid"/>';
      s += '<text x="' + (padL - 6) + '" y="' + (y(v) + 4) + '" class="ylab">' + fmtShort(v) + "</text>";
    }
    // zero line emphasized
    if (lo < 0) s += '<line x1="' + padL + '" y1="' + y(0) + '" x2="' + (W - padR) + '" y2="' + y(0) + '" class="zero"/>';
    // bars
    rows.forEach(function (r, i) {
      var e = -r.expenditures;
      if (e > 0) {
        s += '<rect x="' + (x(i) - bw / 2) + '" y="' + y(e) + '" width="' + bw +
          '" height="' + Math.max(1, y(Math.max(lo, 0)) - y(e)) + '" class="bar"><title>' +
          r.year + " expenditures: $" + Math.round(e).toLocaleString() + "</title></rect>";
      }
    });
    // threshold
    if (opts.threshold != null) {
      s += '<line x1="' + padL + '" y1="' + y(opts.threshold) + '" x2="' + (W - padR) +
        '" y2="' + y(opts.threshold) + '" class="threshold"/>';
    }
    // balance line
    var pts = rows.map(function (r, i) { return x(i) + "," + y(r.end); }).join(" ");
    s += '<polyline points="' + pts + '" class="balline"/>';
    rows.forEach(function (r, i) {
      s += '<circle cx="' + x(i) + '" cy="' + y(r.end) + '" r="2.6" class="baldot' +
        (r.end < 0 ? " neg" : "") + '"><title>' + r.year + " year-end: $" +
        Math.round(r.end).toLocaleString() + "</title></circle>";
    });
    // x labels (every 5 years)
    rows.forEach(function (r, i) {
      if (i % 5 === 0 || i === n - 1) {
        s += '<text x="' + x(i) + '" y="' + (H - 8) + '" class="xlab">' + r.year + "</text>";
      }
    });
    s += "</svg>";
    return s;
  }

  /**
   * Simple bar chart for a single component's forecast expenditures.
   * values: array indexed by offset k (positive dollars). startYear: CFY.
   * Mirrors the "Forecasted expenditures" panel from the RA portal.
   */
  function barChart(values, startYear, opts) {
    opts = opts || {};
    var W = opts.width || 640, H = opts.height || 210;
    var padL = 52, padR = 10, padT = 12, padB = 26;
    var iw = W - padL - padR, ih = H - padT - padB;
    var n = values.length;
    var max = Math.max.apply(null, values.concat([1]));
    var hi = max * 1.08 || 1;
    function y(v) { return padT + ih - (v / hi) * ih; }
    function x(i) { return padL + (i + 0.5) * (iw / n); }
    var bw = Math.max(2, (iw / n) - 3);
    var s = '<svg viewBox="0 0 ' + W + " " + H + '" class="cfchart" role="img" aria-label="Forecasted expenditures over the study">';
    var ticks = 3;
    for (var t = 0; t <= ticks; t++) {
      var v = hi * (t / ticks);
      s += '<line x1="' + padL + '" y1="' + y(v) + '" x2="' + (W - padR) + '" y2="' + y(v) + '" class="grid"/>';
      s += '<text x="' + (padL - 6) + '" y="' + (y(v) + 4) + '" class="ylab">' + fmtShort(v) + "</text>";
    }
    values.forEach(function (val, i) {
      if (val > 0) {
        s += '<rect x="' + (x(i) - bw / 2) + '" y="' + y(val) + '" width="' + bw +
          '" height="' + Math.max(1, y(0) - y(val)) + '" class="bar"><title>' +
          (startYear + i) + ": $" + Math.round(val).toLocaleString() + "</title></rect>";
      }
      if (i % 5 === 0 || i === n - 1) {
        s += '<text x="' + x(i) + '" y="' + (H - 8) + '" class="xlab">' + (startYear + i) + "</text>";
      }
    });
    s += "</svg>";
    return s;
  }

  /**
   * Multi-series funding chart: expenditure bars + one line per funding plan,
   * threshold floor, and invisible full-height hover columns (class
   * "hovercol", data-i) the app wires a tooltip to.
   * config = { startYear, expenditures:[+$ per k], threshold,
   *            series:[{id,label,color,end:[per k]}], hidden:{id:true} }
   */
  function planChart(config) {
    var W = config.width || 920, H = config.height || 300;
    var padL = 58, padR = 14, padT = 16, padB = 30;
    var iw = W - padL - padR, ih = H - padT - padB;
    var exps = config.expenditures || [];
    var n = exps.length;
    var visible = (config.series || []).filter(function (s) { return !(config.hidden && config.hidden[s.id]); });

    var maxV = 0, minV = 0;
    exps.forEach(function (v) { if (v > maxV) maxV = v; });
    visible.forEach(function (s) {
      s.end.forEach(function (v) { if (v > maxV) maxV = v; if (v < minV) minV = v; });
    });
    if (config.threshold != null) maxV = Math.max(maxV, config.threshold);
    var lo = Math.min(minV, 0), hi = Math.max(maxV, 1) * 1.06;
    if (hi === lo) hi = lo + 1;
    function y(v) { return padT + ih - ((v - lo) / (hi - lo)) * ih; }
    function x(i) { return padL + (i + 0.5) * (iw / n); }
    var bw = Math.max(3, (iw / n) - 4);

    var s = '<svg viewBox="0 0 ' + W + " " + H + '" class="cfchart" role="img" aria-label="Reserve balance and expenditures by year">';
    var ticks = 4;
    for (var t = 0; t <= ticks; t++) {
      var gv = lo + (hi - lo) * (t / ticks);
      s += '<line x1="' + padL + '" y1="' + y(gv) + '" x2="' + (W - padR) + '" y2="' + y(gv) + '" class="grid"/>';
      s += '<text x="' + (padL - 6) + '" y="' + (y(gv) + 4) + '" class="ylab">' + fmtShort(gv) + "</text>";
    }
    if (lo < 0) s += '<line x1="' + padL + '" y1="' + y(0) + '" x2="' + (W - padR) + '" y2="' + y(0) + '" class="zero"/>';
    // expenditure bars
    exps.forEach(function (e, i) {
      if (e > 0) s += '<rect x="' + (x(i) - bw / 2) + '" y="' + y(e) + '" width="' + bw +
        '" height="' + Math.max(1, y(Math.max(lo, 0)) - y(e)) + '" class="bar"/>';
    });
    if (config.threshold != null) {
      s += '<line x1="' + padL + '" y1="' + y(config.threshold) + '" x2="' + (W - padR) +
        '" y2="' + y(config.threshold) + '" class="threshold"/>';
    }
    // series lines
    visible.forEach(function (ser) {
      var pts = ser.end.map(function (v, i) { return x(i) + "," + y(v); }).join(" ");
      s += '<polyline points="' + pts + '" fill="none" stroke="' + ser.color + '" stroke-width="2.2"/>';
      ser.end.forEach(function (v, i) {
        s += '<circle cx="' + x(i) + '" cy="' + y(v) + '" r="2.4" fill="' +
          (v < 0 ? "#B5552F" : ser.color) + '"/>';
      });
    });
    // x labels
    for (var i = 0; i < n; i++) {
      if (i % 5 === 0 || i === n - 1) s += '<text x="' + x(i) + '" y="' + (H - 8) + '" class="xlab">' + (config.startYear + i) + "</text>";
    }
    // invisible hover columns
    var cw = iw / n;
    for (var j = 0; j < n; j++) {
      s += '<rect class="hovercol" data-i="' + j + '" x="' + (padL + j * cw) + '" y="' + padT +
        '" width="' + cw + '" height="' + ih + '" fill="transparent"/>';
    }
    s += "</svg>";
    return s;
  }

  return { cashFlowChart: cashFlowChart, barChart: barChart, planChart: planChart, fmtShort: fmtShort };
});
