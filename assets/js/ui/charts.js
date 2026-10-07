/*
 * charts.js — SVG rendering for workbook charts (column, bar, line, pie, scatter).
 *
 * Input is { categories, series: [{ name, values, x? }] } from Workbook.chartData.
 * Colors follow a fixed categorical order (validated for color-vision deficiency);
 * series never get recolored by rank. Every mark carries a <title> tooltip and the
 * worksheet itself is the table view.
 */
(function (root) {
  'use strict';

  const PALETTE = ['#008300', '#2a78d6', '#eb6834', '#4a3aa7', '#eda100', '#e34948'];
  const W = 560, H = 320;
  const PAD = { top: 44, right: 20, bottom: 56, left: 64 };

  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const fmt = (n) => {
    if (n === null || n === undefined || Number.isNaN(n)) return '';
    const abs = Math.abs(n);
    if (abs >= 1e6) return (n / 1e6).toFixed(abs >= 1e7 ? 0 : 1) + 'M';
    if (abs >= 1e4) return (n / 1e3).toFixed(abs >= 1e5 ? 0 : 1) + 'K';
    return Number.isInteger(n) ? n.toLocaleString('en-US') : Number(n.toPrecision(4)).toLocaleString('en-US');
  };

  function niceTicks(min, max, count = 5) {
    if (min === max) { max = min + 1; }
    const span = max - min;
    const step0 = Math.pow(10, Math.floor(Math.log10(span / count)));
    const err = (span / count) / step0;
    const step = step0 * (err >= 7.5 ? 10 : err >= 3.5 ? 5 : err >= 1.5 ? 2 : 1);
    const lo = Math.floor(min / step) * step;
    const hi = Math.ceil(max / step) * step;
    const ticks = [];
    for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Number(v.toPrecision(12)));
    return ticks;
  }

  function legend(series, y) {
    if (series.length < 2) return '';
    let x = PAD.left;
    return series.map((s, i) => {
      const item = '<g transform="translate(' + x + ',' + y + ')"><rect width="10" height="10" rx="2" fill="' + PALETTE[i % PALETTE.length] + '"/>' +
        '<text x="15" y="9" class="ch-legend">' + esc(s.name) + '</text></g>';
      x += 26 + Math.min(18, String(s.name).length) * 7;
      return item;
    }).join('');
  }

  function frame(title, inner, note) {
    return '<svg class="chart-svg" viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="' + esc(title || 'Chart') + '" xmlns="http://www.w3.org/2000/svg">' +
      '<rect width="' + W + '" height="' + H + '" fill="#ffffff"/>' +
      (title ? '<text x="' + PAD.left + '" y="24" class="ch-title">' + esc(title) + '</text>' : '') +
      inner + (note ? '<text x="' + (W - PAD.right) + '" y="24" text-anchor="end" class="ch-note">' + esc(note) + '</text>' : '') +
      '</svg>';
  }

  function valueAxis(min, max, chart) {
    let lo = Math.min(0, min), hi = Math.max(0, max);
    const userMin = chart.axisMin !== null && chart.axisMin !== undefined && chart.axisMin !== '' ? Number(chart.axisMin) : null;
    if (userMin !== null && !Number.isNaN(userMin)) lo = userMin;
    const ticks = niceTicks(lo, hi);
    if (userMin !== null && !Number.isNaN(userMin)) ticks[0] = userMin;
    return { lo: userMin !== null && !Number.isNaN(userMin) ? userMin : ticks[0], hi: ticks[ticks.length - 1], ticks: ticks.filter((t) => t >= (userMin ?? -Infinity)) };
  }

  function renderBars(data, chart, horizontal) {
    const { categories, series } = data;
    const all = series.flatMap((s) => s.values).filter((v) => typeof v === 'number');
    if (!all.length) return empty(chart);
    const ax = valueAxis(Math.min(...all), Math.max(...all), chart);
    const plotW = W - PAD.left - PAD.right, plotH = H - PAD.top - PAD.bottom;
    const n = categories.length, k = series.length;
    let out = '';
    const scale = (v) => ((v - ax.lo) / (ax.hi - ax.lo || 1));
    if (!horizontal) {
      ax.ticks.forEach((t) => {
        const y = PAD.top + plotH - scale(t) * plotH;
        out += '<line x1="' + PAD.left + '" x2="' + (W - PAD.right) + '" y1="' + y + '" y2="' + y + '" class="ch-grid"/>' +
          '<text x="' + (PAD.left - 8) + '" y="' + (y + 4) + '" text-anchor="end" class="ch-axis">' + fmt(t) + '</text>';
      });
      const band = plotW / n, barW = Math.max(4, Math.min(48, (band * 0.72) / k - 2));
      const base = PAD.top + plotH - scale(Math.max(ax.lo, Math.min(0, ax.hi))) * plotH;
      categories.forEach((c, i) => {
        const cx = PAD.left + band * i + band / 2;
        series.forEach((s, j) => {
          const v = s.values[i];
          if (typeof v !== 'number') return;
          const x = cx - (k * (barW + 2)) / 2 + j * (barW + 2);
          const y = PAD.top + plotH - scale(Math.max(v, ax.lo)) * plotH;
          const top = Math.min(y, base), h = Math.max(1, Math.abs(base - y));
          out += '<rect x="' + x.toFixed(1) + '" y="' + top.toFixed(1) + '" width="' + barW.toFixed(1) + '" height="' + h.toFixed(1) + '" rx="3" fill="' + PALETTE[j % PALETTE.length] + '"><title>' + esc(c + ' · ' + s.name + ': ' + fmt(v)) + '</title></rect>';
        });
        const label = String(c).length > 12 ? String(c).slice(0, 11) + '…' : c;
        out += '<text x="' + cx + '" y="' + (PAD.top + plotH + 18) + '" text-anchor="middle" class="ch-axis">' + esc(label) + '</text>';
      });
      out += '<line x1="' + PAD.left + '" x2="' + (W - PAD.right) + '" y1="' + base + '" y2="' + base + '" class="ch-base"/>';
    } else {
      const left = 110;
      const pw = W - left - PAD.right;
      ax.ticks.forEach((t) => {
        const x = left + scale(t) * pw;
        out += '<line y1="' + PAD.top + '" y2="' + (PAD.top + plotH) + '" x1="' + x + '" x2="' + x + '" class="ch-grid"/>' +
          '<text y="' + (PAD.top + plotH + 18) + '" x="' + x + '" text-anchor="middle" class="ch-axis">' + fmt(t) + '</text>';
      });
      const band = plotH / n, barH = Math.max(4, Math.min(28, (band * 0.72) / k - 2));
      const base = left + scale(Math.max(ax.lo, 0)) * pw;
      categories.forEach((c, i) => {
        const cy = PAD.top + band * i + band / 2;
        series.forEach((s, j) => {
          const v = s.values[i];
          if (typeof v !== 'number') return;
          const y = cy - (k * (barH + 2)) / 2 + j * (barH + 2);
          const x = left + scale(Math.max(v, ax.lo)) * pw;
          out += '<rect y="' + y.toFixed(1) + '" x="' + Math.min(x, base).toFixed(1) + '" height="' + barH.toFixed(1) + '" width="' + Math.max(1, Math.abs(x - base)).toFixed(1) + '" rx="3" fill="' + PALETTE[j % PALETTE.length] + '"><title>' + esc(c + ' · ' + s.name + ': ' + fmt(v)) + '</title></rect>';
        });
        const label = String(c).length > 14 ? String(c).slice(0, 13) + '…' : c;
        out += '<text x="' + (left - 8) + '" y="' + (cy + 4) + '" text-anchor="end" class="ch-axis">' + esc(label) + '</text>';
      });
      out += '<line y1="' + PAD.top + '" y2="' + (PAD.top + plotH) + '" x1="' + base + '" x2="' + base + '" class="ch-base"/>';
    }
    const note = ax.lo > 0 && !horizontal ? 'Axis starts at ' + fmt(ax.lo) : ax.lo > 0 ? 'Axis starts at ' + fmt(ax.lo) : '';
    return frame(chart.title, out + legend(series, H - 18), note);
  }

  function renderLine(data, chart) {
    const { categories, series } = data;
    const all = series.flatMap((s) => s.values).filter((v) => typeof v === 'number');
    if (!all.length) return empty(chart);
    let lo = Math.min(...all), hi = Math.max(...all);
    const userMin = chart.axisMin !== null && chart.axisMin !== undefined && chart.axisMin !== '' ? Number(chart.axisMin) : null;
    if (userMin !== null && !Number.isNaN(userMin)) lo = userMin;
    else lo = Math.min(0, lo);
    const ticks = niceTicks(lo, hi);
    const tlo = userMin !== null && !Number.isNaN(userMin) ? userMin : ticks[0], thi = ticks[ticks.length - 1];
    const plotW = W - PAD.left - PAD.right, plotH = H - PAD.top - PAD.bottom;
    const n = categories.length;
    const xAt = (i) => PAD.left + (n === 1 ? plotW / 2 : (plotW * i) / (n - 1));
    const yAt = (v) => PAD.top + plotH - ((v - tlo) / (thi - tlo || 1)) * plotH;
    let out = '';
    ticks.filter((t) => t >= tlo).forEach((t) => {
      out += '<line x1="' + PAD.left + '" x2="' + (W - PAD.right) + '" y1="' + yAt(t) + '" y2="' + yAt(t) + '" class="ch-grid"/>' +
        '<text x="' + (PAD.left - 8) + '" y="' + (yAt(t) + 4) + '" text-anchor="end" class="ch-axis">' + fmt(t) + '</text>';
    });
    const every = Math.ceil(n / 10);
    categories.forEach((c, i) => {
      if (i % every) return;
      out += '<text x="' + xAt(i) + '" y="' + (PAD.top + plotH + 18) + '" text-anchor="middle" class="ch-axis">' + esc(String(c).length > 10 ? String(c).slice(0, 9) + '…' : c) + '</text>';
    });
    series.forEach((s, j) => {
      const color = PALETTE[j % PALETTE.length];
      const pts = s.values.map((v, i) => (typeof v === 'number' ? [xAt(i), yAt(Math.max(v, tlo))] : null));
      let d = '', pen = false;
      pts.forEach((p) => { if (!p) { pen = false; return; } d += (pen ? 'L' : 'M') + p[0].toFixed(1) + ',' + p[1].toFixed(1); pen = true; });
      out += '<path d="' + d + '" fill="none" stroke="' + color + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>';
      pts.forEach((p, i) => {
        if (!p) return;
        out += '<circle cx="' + p[0].toFixed(1) + '" cy="' + p[1].toFixed(1) + '" r="4" fill="' + color + '" stroke="#fff" stroke-width="2"><title>' + esc(categories[i] + ' · ' + s.name + ': ' + fmt(s.values[i])) + '</title></circle>';
      });
    });
    return frame(chart.title, out + legend(series, H - 18), tlo > 0 ? 'Axis starts at ' + fmt(tlo) : '');
  }

  function renderPie(data, chart) {
    const s = data.series[0];
    if (!s) return empty(chart);
    let items = data.categories.map((c, i) => ({ c, v: s.values[i] })).filter((x) => typeof x.v === 'number' && x.v > 0);
    if (!items.length) return empty(chart);
    if (items.length > PALETTE.length) {
      const sorted = [...items].sort((a, b) => b.v - a.v);
      const keep = sorted.slice(0, PALETTE.length - 1);
      const other = sorted.slice(PALETTE.length - 1).reduce((a, b) => a + b.v, 0);
      items = [...keep, { c: 'Other', v: other }];
    }
    const total = items.reduce((a, b) => a + b.v, 0);
    const cx = 170, cy = 180, r = 110;
    let a0 = -Math.PI / 2;
    let out = '';
    items.forEach((it, i) => {
      const frac = it.v / total;
      const a1 = a0 + frac * Math.PI * 2;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const p0 = [cx + r * Math.cos(a0), cy + r * Math.sin(a0)];
      const p1 = [cx + r * Math.cos(a1), cy + r * Math.sin(a1)];
      const d = items.length === 1
        ? 'M' + (cx - r) + ',' + cy + 'a' + r + ',' + r + ' 0 1,0 ' + 2 * r + ',0a' + r + ',' + r + ' 0 1,0 ' + -2 * r + ',0'
        : 'M' + cx + ',' + cy + 'L' + p0[0].toFixed(1) + ',' + p0[1].toFixed(1) + 'A' + r + ',' + r + ' 0 ' + large + ',1 ' + p1[0].toFixed(1) + ',' + p1[1].toFixed(1) + 'Z';
      out += '<path d="' + d + '" fill="' + PALETTE[i % PALETTE.length] + '" stroke="#fff" stroke-width="2"><title>' + esc(it.c + ': ' + fmt(it.v) + ' (' + (frac * 100).toFixed(1) + '%)') + '</title></path>';
      const ly = 70 + i * 26;
      out += '<rect x="320" y="' + (ly - 10) + '" width="12" height="12" rx="2" fill="' + PALETTE[i % PALETTE.length] + '"/>' +
        '<text x="340" y="' + ly + '" class="ch-legend">' + esc(String(it.c).slice(0, 22)) + ' — ' + (frac * 100).toFixed(1) + '%</text>';
      a0 = a1;
    });
    return frame(chart.title, out);
  }

  function renderScatter(data, chart) {
    const pts = [];
    data.series.forEach((s, j) => s.values.forEach((v, i) => {
      const x = s.x ? s.x[i] : i + 1;
      if (typeof v === 'number' && typeof x === 'number') pts.push({ x, y: v, j, label: data.categories[i], name: s.name });
    }));
    if (!pts.length) return empty(chart);
    const xs = niceTicks(Math.min(...pts.map((p) => p.x)), Math.max(...pts.map((p) => p.x)));
    const ys = niceTicks(Math.min(0, ...pts.map((p) => p.y)), Math.max(...pts.map((p) => p.y)));
    const plotW = W - PAD.left - PAD.right, plotH = H - PAD.top - PAD.bottom;
    const xAt = (x) => PAD.left + ((x - xs[0]) / (xs[xs.length - 1] - xs[0] || 1)) * plotW;
    const yAt = (y) => PAD.top + plotH - ((y - ys[0]) / (ys[ys.length - 1] - ys[0] || 1)) * plotH;
    let out = '';
    ys.forEach((t) => { out += '<line x1="' + PAD.left + '" x2="' + (W - PAD.right) + '" y1="' + yAt(t) + '" y2="' + yAt(t) + '" class="ch-grid"/><text x="' + (PAD.left - 8) + '" y="' + (yAt(t) + 4) + '" text-anchor="end" class="ch-axis">' + fmt(t) + '</text>'; });
    xs.forEach((t) => { out += '<text x="' + xAt(t) + '" y="' + (PAD.top + plotH + 18) + '" text-anchor="middle" class="ch-axis">' + fmt(t) + '</text>'; });
    pts.forEach((p) => {
      out += '<circle cx="' + xAt(p.x).toFixed(1) + '" cy="' + yAt(p.y).toFixed(1) + '" r="5" fill="' + PALETTE[p.j % PALETTE.length] + '" fill-opacity="0.85" stroke="#fff" stroke-width="2"><title>' + esc(p.name + ' (' + fmt(p.x) + ', ' + fmt(p.y) + ')') + '</title></circle>';
    });
    return frame(chart.title, out + legend(data.series, H - 18));
  }

  function empty(chart) {
    return frame(chart.title, '<text x="' + W / 2 + '" y="' + H / 2 + '" text-anchor="middle" class="ch-axis">No numbers to chart in ' + esc(chart.range) + '</text>');
  }

  function render(data, chart) {
    switch (chart.type) {
      case 'bar': return renderBars(data, chart, true);
      case 'line': return renderLine(data, chart);
      case 'pie': return renderPie(data, chart);
      case 'scatter': return renderScatter(data, chart);
      default: return renderBars(data, chart, false);
    }
  }

  const api = { render, PALETTE, niceTicks };
  root.DIM = root.DIM || {};
  root.DIM.Charts = api;
  if (typeof module === 'object' && module.exports) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
