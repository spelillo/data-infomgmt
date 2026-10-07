/*
 * dom.js — tiny DOM helpers shared by the UI: element builder, modal dialogs,
 * toasts and safe HTML escaping.
 */
(function (root) {
  'use strict';

  function h(tag, attrs, ...children) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'html') el.innerHTML = v;
      else if (k === 'text') el.textContent = v;
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k === 'dataset') Object.assign(el.dataset, v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    }
    for (const c of children.flat(Infinity)) {
      if (c === null || c === undefined || c === false) continue;
      el.append(c instanceof Node ? c : document.createTextNode(String(c)));
    }
    return el;
  }

  function esc(s) {
    return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  // Modal dialog. buttons: [{ label, primary, onClick(close) -> false keeps open }]
  function dialog(title, body, buttons = [{ label: 'Close' }], opts = {}) {
    const prevFocus = document.activeElement;
    const overlay = h('div', { class: 'modal-overlay', role: 'presentation' });
    const box = h('div', { class: 'modal' + (opts.wide ? ' modal-wide' : ''), role: 'dialog', 'aria-modal': 'true', 'aria-label': title });
    const close = () => {
      overlay.remove();
      document.removeEventListener('keydown', onKey, true);
      if (prevFocus && prevFocus.focus) prevFocus.focus();
      if (opts.onClose) opts.onClose();
    };
    const onKey = (e) => {
      if (e.key === 'Escape') { e.stopPropagation(); close(); }
      if (e.key === 'Tab') {
        const f = [...box.querySelectorAll('button, input, select, textarea, [tabindex]:not([tabindex="-1"])')].filter((x) => !x.disabled && x.offsetParent !== null);
        if (!f.length) return;
        if (e.shiftKey && document.activeElement === f[0]) { e.preventDefault(); f[f.length - 1].focus(); }
        else if (!e.shiftKey && document.activeElement === f[f.length - 1]) { e.preventDefault(); f[0].focus(); }
      }
    };
    const footer = h('div', { class: 'modal-actions' }, buttons.map((b) => h('button', {
      class: 'btn' + (b.primary ? ' btn-primary' : ''), type: 'button',
      onClick: () => { if (!b.onClick || b.onClick(close) !== false) close(); },
    }, b.label)));
    box.append(
      h('div', { class: 'modal-head' }, h('h3', {}, title), h('button', { class: 'icon-btn', 'aria-label': 'Close', type: 'button', onClick: close }, '×')),
      h('div', { class: 'modal-body' }, body),
      footer,
    );
    overlay.append(box);
    overlay.addEventListener('mousedown', (e) => { if (e.target === overlay) close(); });
    document.body.append(overlay);
    document.addEventListener('keydown', onKey, true);
    setTimeout(() => {
      const first = box.querySelector('[autofocus], input, select, textarea, .btn-primary');
      if (first) first.focus();
    }, 0);
    return { close, box };
  }

  function toast(message, kind = 'info') {
    let host = document.querySelector('.toasts');
    if (!host) { host = h('div', { class: 'toasts', 'aria-live': 'polite' }); document.body.append(host); }
    const t = h('div', { class: 'toast toast-' + kind }, message);
    host.append(t);
    setTimeout(() => t.classList.add('show'), 10);
    setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, 3200);
  }

  function field(label, control, help) {
    return h('label', { class: 'field' }, h('span', { class: 'field-label' }, label), control, help ? h('span', { class: 'field-help' }, help) : null);
  }

  function select(options, value, attrs = {}) {
    return h('select', attrs, options.map((o) => {
      const [v, l] = Array.isArray(o) ? o : [o, o];
      return h('option', { value: v, selected: String(v) === String(value) }, l);
    }));
  }

  function downloadText(text, filename, mime = 'text/plain') {
    const blob = new Blob([text], { type: mime + ';charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = h('a', { href: url, download: filename });
    document.body.append(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  function readFile(file, as = 'text') {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(r.error);
      if (as === 'buffer') r.readAsArrayBuffer(file); else r.readAsText(file);
    });
  }

  root.DIM = root.DIM || {};
  root.DIM.Dom = { h, esc, dialog, toast, field, select, downloadText, readFile };
})(typeof globalThis !== 'undefined' ? globalThis : this);
