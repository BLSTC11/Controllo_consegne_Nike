/* Interfaccia: caricamento file, tabelle dei risultati, export Excel e stampa. */
(function () {
  'use strict';

  const E = window.Engine;
  const $ = sel => document.querySelector(sel);

  // ---------------------------------------------------------------------------
  // Stato
  // ---------------------------------------------------------------------------
  const state = {
    order: null,        // { id, name, buffer, table, parsed, overrides }
    deliveries: [],     // [{ id, name, buffer, table, parsed, overrides }]
    result: null,
    season: '',         // filtro stagione ('' = tutte)
    tab: 'riepilogo',
    ui: { search: {}, sort: {}, status: '', onlyOpen: true },
  };

  // ---------------------------------------------------------------------------
  // Formattazione
  // ---------------------------------------------------------------------------
  const fInt = new Intl.NumberFormat('bg-BG', { maximumFractionDigits: 0, useGrouping: 'always' });
  const fEur = new Intl.NumberFormat('bg-BG', { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: 'always' });
  const fPct = new Intl.NumberFormat('bg-BG', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function fmt(v, type) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'string' && type !== 'text' && type !== 'status') return esc(v);
    switch (type) {
      case 'int': return fInt.format(v);
      case 'eur': return fEur.format(v);
      case 'price': return fEur.format(v);
      case 'diff': return (v > 0 ? '+' : '') + fEur.format(v);
      case 'pct': return fPct.format(v);
      case 'bool': return v ? 'Да' : '';
      default: return esc(v);
    }
  }

  const STATUS_CLASS = { 'Предстои доставка': 'b-da-consegnare', 'Частично': 'b-parziale', 'Изпълнено': 'b-completo', 'Излишък': 'b-eccedenza' };

  // ---------------------------------------------------------------------------
  // Memoria locale (IndexedDB): i file restano sul computer dell'utente
  // ---------------------------------------------------------------------------
  const DB_NAME = 'controllo-consegne-nike';
  let dbPromise = null;
  function db() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        try {
          const req = indexedDB.open(DB_NAME, 1);
          req.onupgradeneeded = () => req.result.createObjectStore('files', { keyPath: 'id' });
          req.onsuccess = () => resolve(req.result);
          req.onerror = () => reject(req.error);
        } catch (e) { reject(e); }
      });
    }
    return dbPromise;
  }
  async function dbDo(mode, fn) {
    try {
      const d = await db();
      return await new Promise((resolve, reject) => {
        const tx = d.transaction('files', mode);
        const out = fn(tx.objectStore('files'));
        tx.oncomplete = () => resolve(out && out.result !== undefined ? out.result : undefined);
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) { return undefined; }
  }
  const dbPut = rec => dbDo('readwrite', s => s.put(rec));
  const dbDel = id => dbDo('readwrite', s => s.delete(id));
  const dbAll = () => dbDo('readonly', s => s.getAll());
  function persist(f, role) {
    return dbPut({ id: f.id, role, name: f.name, buffer: f.buffer, overrides: f.overrides || {}, added: f.added || Date.now() });
  }

  // ---------------------------------------------------------------------------
  // Lettura file
  // ---------------------------------------------------------------------------
  let seq = 0;
  const newId = () => Date.now().toString(36) + '-' + (seq++).toString(36);

  function buildTable(buffer, fields, overrides) {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const table = E.readWorkbookTable(XLSX, wb, fields);
    if (!table) return null;
    if (overrides) for (const k in overrides) table.map[k] = overrides[k];
    return table;
  }

  function loadOrder(rec) {
    const table = buildTable(rec.buffer, E.ORDER_FIELDS, rec.overrides);
    const f = Object.assign({}, rec, { table });
    if (table) f.parsed = E.parseOrder(table);
    return f;
  }
  function loadDelivery(rec) {
    const table = buildTable(rec.buffer, E.DELIVERY_FIELDS, rec.overrides);
    const f = Object.assign({}, rec, { table });
    if (table) f.parsed = E.parseDelivery(table, rec.name);
    return f;
  }

  async function addOrderFile(file) {
    const buffer = await file.arrayBuffer();
    if (state.order) await dbDel(state.order.id);
    state.order = loadOrder({ id: newId(), name: file.name, buffer, overrides: {}, added: Date.now() });
    if (!state.order.table) message('error', `Във файла "${esc(file.name)}" не намирам колоните EAN и количество. Това ли е правилният файл с потвърждението?`);
    await persist(state.order, 'order');
    refresh();
  }

  async function addDeliveryFiles(files) {
    for (const file of files) {
      const buffer = await file.arrayBuffer();
      const existing = state.deliveries.find(d => d.name === file.name);
      if (existing) {
        await dbDel(existing.id);
        state.deliveries = state.deliveries.filter(d => d !== existing);
        message('info', `"${esc(file.name)}" вече беше зареден: заменен е с новата версия.`);
      }
      const f = loadDelivery({ id: newId(), name: file.name, buffer, overrides: {}, added: Date.now() });
      if (!f.table) message('error', `Във файла "${esc(file.name)}" не намирам колоните EAN и количество.`);
      state.deliveries.push(f);
      await persist(f, 'delivery');
    }
    state.deliveries.sort((a, b) => a.name.localeCompare(b.name, 'it', { numeric: true }));
    refresh();
  }

  async function restore() {
    const recs = await dbAll();
    if (!recs || !recs.length) return;
    for (const r of recs.sort((a, b) => a.added - b.added)) {
      try {
        if (r.role === 'order') state.order = loadOrder(r);
        else state.deliveries.push(loadDelivery(r));
      } catch (e) { /* file illeggibile: lo ignoro */ }
    }
    state.deliveries.sort((a, b) => a.name.localeCompare(b.name, 'it', { numeric: true }));
    refresh();
  }

  // ---------------------------------------------------------------------------
  // Messaggi
  // ---------------------------------------------------------------------------
  function message(kind, html) {
    const el = document.createElement('div');
    el.className = 'msg msg-' + kind;
    el.innerHTML = `<span>${html}</span><button class="link" aria-label="Затвори">×</button>`;
    el.querySelector('button').onclick = () => el.remove();
    $('#messages').appendChild(el);
  }

  // ---------------------------------------------------------------------------
  // Pannelli di caricamento
  // ---------------------------------------------------------------------------
  function mappingHtml(f, fields, role) {
    if (!f.table) return '';
    const t = f.table;
    const found = Object.keys(fields).filter(k => t.map[k] >= 0).length;
    const missingReq = Object.keys(fields).filter(k => fields[k].required && t.map[k] < 0);
    const opts = sel => ['<option value="-1">— липсва —</option>']
      .concat(t.headers.map((h, i) => `<option value="${i}"${i === sel ? ' selected' : ''}>${esc(h || '(колона ' + (i + 1) + ')')}</option>`)).join('');
    const rows = Object.keys(fields).map(k => `
      <label class="map-row${fields[k].required && t.map[k] < 0 ? ' map-missing' : ''}">
        <span>${esc(fields[k].label)}${fields[k].required ? ' *' : ''}</span>
        <select data-role="${role}" data-id="${f.id}" data-field="${k}">${opts(t.map[k])}</select>
      </label>`).join('');
    return `<details class="mapping"${missingReq.length ? ' open' : ''}>
      <summary>Разпознати колони: ${found} от ${Object.keys(fields).length} (лист "${esc(t.sheet)}", заглавия на ред ${t.headerRow + 1})</summary>
      <div class="map-grid">${rows}</div>
    </details>`;
  }

  function seasonsText(rawSeasons) {
    const parts = Object.keys(rawSeasons).map(r => rawSeasons[r] && rawSeasons[r] !== r ? `${esc(r)} → ${esc(rawSeasons[r])}` : esc(r));
    return parts.length ? parts.join(', ') : 'не е посочен';
  }

  function renderUploads() {
    const o = state.order;
    $('#clear-order').hidden = !o;
    $('#drop-order').classList.toggle('compact', !!o);
    if (!o) { $('#order-info').innerHTML = ''; }
    else if (!o.table) {
      $('#order-info').innerHTML = `<div class="file bad"><strong>${esc(o.name)}</strong><span>не са намерени колоните EAN и количество</span></div>`;
    } else {
      const L = o.parsed.lines;
      const qty = L.reduce((s, l) => s + l.qty, 0), val = L.reduce((s, l) => s + l.qty * l.price, 0);
      const pos = Array.from(new Set(L.map(l => l.po).filter(Boolean)));
      $('#order-info').innerHTML = `<div class="file">
        <div class="file-name"><strong>${esc(o.name)}</strong></div>
        <div class="file-stats">${fInt.format(L.length)} реда · ${fInt.format(qty)} бр. · ${fEur.format(val)} € · Сезон: ${seasonsText(o.parsed.rawSeasons)}${pos.length ? ' · Поръчки: ' + pos.map(esc).join(', ') : ''}</div>
        ${mappingHtml(o, E.ORDER_FIELDS, 'order')}
      </div>`;
    }

    $('#clear-deliveries').hidden = !state.deliveries.length;
    $('#delivery-list').innerHTML = state.deliveries.map(d => {
      if (!d.table) return `<div class="file bad"><div class="file-name"><strong>${esc(d.name)}</strong><button class="link" data-remove="${d.id}">Премахни</button></div><div class="file-stats">не са намерени колоните EAN и количество</div></div>`;
      const L = d.parsed.lines;
      const qty = L.reduce((s, l) => s + l.qty, 0);
      const trucks = new Set(L.map(l => l.truck).filter(Boolean)).size;
      return `<div class="file">
        <div class="file-name"><strong>${esc(d.name)}</strong><button class="link" data-remove="${d.id}">Премахни</button></div>
        <div class="file-stats">${fInt.format(L.length)} реда · ${fInt.format(qty)} бр.${trucks ? ' · ' + trucks + (trucks === 1 ? ' камион' : ' камиона') : ''} · Сезон: ${seasonsText(d.parsed.rawSeasons)}</div>
        ${mappingHtml(d, E.DELIVERY_FIELDS, 'delivery')}
      </div>`;
    }).join('');
  }

  // ---------------------------------------------------------------------------
  // Calcolo e filtro per stagione
  // ---------------------------------------------------------------------------
  function compute() {
    const o = state.order;
    const ds = state.deliveries.filter(d => d.table);
    if (!o || !o.table || !ds.length) { state.result = null; return; }
    const lines = [].concat(...ds.map(d => d.parsed.lines));
    state.result = E.compare(o.parsed.lines, lines);
    state.result.deliveryLines = lines;
  }

  const sumBy = (arr, f) => arr.reduce((s, x) => s + f(x), 0);

  function view() {
    const r = state.result;
    const s = state.season;
    const keep = x => !s || x.season === s;
    const v = {
      detail: r.detail.filter(keep),
      nonOrdered: r.nonOrdered.filter(keep),
      otherSeason: r.otherSeason.filter(x => !s || x.orderSeason === s || x.season === s),
      shipments: r.shipments.filter(keep),
    };
    v.excess = v.detail.filter(d => d.excess > 0);
    v.differences = v.detail.filter(d => (d.priceDiff !== null && Math.abs(d.priceDiff) >= 0.01) || d.poDiff);
    const d = v.detail;
    v.t = {
      ordered: { qty: sumBy(d, x => x.qty), value: sumBy(d, x => x.orderValue) },
      onOrder: { qty: sumBy(d, x => x.onOrder), value: sumBy(d, x => x.onOrderValue) },
      residual: { qty: sumBy(d, x => x.residual), value: sumBy(d, x => x.residualValue) },
      excess: { qty: sumBy(d, x => x.excess), value: sumBy(d, x => x.excessValue) },
      nonOrdered: { qty: sumBy(v.nonOrdered, x => x.qty), value: sumBy(v.nonOrdered, x => x.fpcValue) },
      otherSeason: { qty: sumBy(v.otherSeason, x => x.qty), value: sumBy(v.otherSeason, x => x.fpcValue) },
      all: { qty: sumBy(v.shipments, x => x.qty), value: sumBy(v.shipments, x => x.fpcValue), net: sumBy(v.shipments, x => x.netValue) },
    };
    v.hasNet = r.deliveryLines.some(l => l.netValue !== null);
    return v;
  }

  // ---------------------------------------------------------------------------
  // Definizioni delle tabelle (usate a video, in stampa e nell'export Excel)
  // ---------------------------------------------------------------------------
  const C = (key, label, type, extra) => Object.assign({ key, label, type: type || 'text' }, extra);

  const groupCols = first => first.concat([
    C('qty', 'Бр. поръчани', 'int'), C('onOrder', 'Бр. доставени', 'int'), C('residual', 'Бр. остатък', 'int'), C('excess', 'Бр. излишък', 'int'),
    C('orderValue', 'Стойност поръчано €', 'eur'), C('onOrderValue', 'Стойност доставено €', 'eur'), C('residualValue', 'Стойност остатък €', 'eur'),
    C('pct', '% доставено', 'pct', { bar: true }),
  ]);

  const COLS = {
    season: groupCols([C('season', 'Сезон')]),
    bu: groupCols([C('season', 'Сезон'), C('bu', 'Отдел')]),
    po: groupCols([C('season', 'Сезон'), C('po', '№ поръчка')]),
    gender: groupCols([C('season', 'Сезон'), C('gender', 'Пол')]),
    article: [
      C('season', 'Сезон'), C('bu', 'Отдел'), C('article', 'Артикул'), C('name', 'Описание'), C('color', 'Цвят'),
      C('price', 'Цена FPC €', 'price'), C('qty', 'Бр. поръчани', 'int'), C('onOrder', 'Бр. доставени', 'int'), C('residual', 'Бр. остатък', 'int'),
      C('excess', 'Бр. излишък', 'int'), C('orderValue', 'Стойност поръчано €', 'eur'), C('residualValue', 'Стойност остатък €', 'eur'),
      C('sizesTxt', 'Изпълнени размери'), C('pct', '% доставено', 'pct', { bar: true }),
    ],
    ean: [
      C('status', 'Статус', 'status'), C('season', 'Сезон'), C('po', '№ поръчка'), C('bu', 'Отдел'), C('article', 'Артикул'), C('name', 'Описание'),
      C('color', 'Цвят'), C('sizeUS', 'Размер US'), C('sizeEU', 'Размер EU'), C('ean', 'EAN'), C('price', 'Цена FPC €', 'price'),
      C('qty', 'Бр. поръчани', 'int'), C('delivered', 'Бр. доставени', 'int'), C('residual', 'Бр. остатък', 'int'), C('excess', 'Бр. излишък', 'int'),
      C('orderValue', 'Стойност поръчано €', 'eur'), C('onOrderValue', 'Стойност доставено €', 'eur'), C('residualValue', 'Стойност остатък €', 'eur'),
      C('deliveredNet', 'Нетна фактурирана стойност €', 'eur', { net: true }), C('files', 'Файл доставка'),
    ],
    excess: [
      C('season', 'Сезон'), C('po', '№ поръчка'), C('bu', 'Отдел'), C('article', 'Артикул'), C('name', 'Описание'), C('color', 'Цвят'),
      C('sizeUS', 'Размер US'), C('ean', 'EAN'), C('price', 'Цена FPC €', 'price'), C('qty', 'Бр. поръчани', 'int'), C('delivered', 'Бр. доставени', 'int'),
      C('excess', 'Бр. излишък', 'int'), C('excessValue', 'Стойност излишък €', 'eur'), C('files', 'Файл доставка'),
    ],
    nonOrdered: [
      C('season', 'Сезон доставка'), C('po', '№ поръчка доставка'), C('bu', 'Отдел'), C('article', 'Артикул'), C('name', 'Описание'), C('size', 'Размер'),
      C('ean', 'EAN'), C('prices', 'Цена FPC €', 'price'), C('qty', 'Бр. доставени', 'int'), C('fpcValue', 'Стойност FPC €', 'eur'),
      C('netValue', 'Нетна стойност €', 'eur', { net: true }), C('trucks', 'Камион'), C('files', 'Файл доставка'),
    ],
    otherSeason: [
      C('season', 'Сезон доставка'), C('orderSeason', 'Сезон поръчка'), C('po', '№ поръчка доставка'), C('bu', 'Отдел'), C('article', 'Артикул'),
      C('name', 'Описание'), C('color', 'Цвят'), C('size', 'Размер'), C('ean', 'EAN'), C('orderQty', 'Бр. поръчани', 'int'), C('qty', 'Бр. доставени', 'int'),
      C('fpcValue', 'Стойност FPC €', 'eur'), C('trucks', 'Камион'), C('files', 'Файл доставка'),
    ],
    differences: [
      C('season', 'Сезон'), C('po', '№ поръчка'), C('delPO', '№ поръчка доставка'), C('poDiff', 'Различна поръчка', 'bool'), C('bu', 'Отдел'),
      C('article', 'Артикул'), C('name', 'Описание'), C('sizeUS', 'Размер US'), C('ean', 'EAN'), C('price', 'Цена FPC поръчка €', 'price'),
      C('delPrice', 'Цена FPC доставка €', 'price'), C('priceDiff', 'Разлика €', 'diff'), C('delivered', 'Бр. доставени', 'int'), C('files', 'Файл доставка'),
    ],
    shipments: [
      C('file', 'Файл доставка'), C('season', 'Сезон'), C('packing', 'Опаковъчен лист'), C('truck', 'Камион'), C('po', '№ поръчка'), C('lines', 'Редове', 'int'),
      C('qty', 'Бройки', 'int'), C('fpcValue', 'Стойност FPC €', 'eur'), C('netValue', 'Нетна стойност €', 'eur', { net: true }),
      C('onOrder', 'Бр. по поръчката', 'int'), C('notOrdered', 'Бр. непоръчани', 'int'), C('otherSeason', 'Бр. друг сезон', 'int'),
    ],
  };

  function cols(name, v) {
    return COLS[name].filter(c => !c.net || v.hasNet);
  }

  function tableData(v) {
    const arts = E.byArticle(v.detail).map(a => Object.assign(a, { sizesTxt: `${a.sizesDone} di ${a.sizes}` }));
    return {
      season: E.groupDetail(v.detail, ['season']),
      bu: E.groupDetail(v.detail, ['season', 'bu']),
      po: E.groupDetail(v.detail, ['season', 'po']),
      gender: E.groupDetail(v.detail, ['season', 'gender']),
      article: arts,
      ean: v.detail,
      excess: v.excess,
      nonOrdered: v.nonOrdered,
      otherSeason: v.otherSeason,
      differences: v.differences,
      shipments: v.shipments,
    };
  }

  // Ordinamento predefinito di ogni tabella
  const DEFAULT_SORT = {
    season: ['season', 1], bu: ['bu', 1], po: ['po', 1], gender: ['gender', 1],
    article: ['residualValue', -1], ean: null, excess: ['excessValue', -1], nonOrdered: ['fpcValue', -1],
    otherSeason: ['fpcValue', -1], differences: null, shipments: null,
  };

  function sortRows(name, rows) {
    const s = state.ui.sort[name] || DEFAULT_SORT[name];
    if (!s) return rows;
    const [k, dir] = s;
    return rows.slice().sort((a, b) => {
      const x = a[k], y = b[k];
      if (typeof x === 'number' || typeof y === 'number') return ((x || 0) - (y || 0)) * dir;
      return String(x == null ? '' : x).localeCompare(String(y == null ? '' : y), 'it', { numeric: true }) * dir;
    });
  }

  function htmlTable(name, columns, rows, opts) {
    opts = opts || {};
    const s = state.ui.sort[name] || DEFAULT_SORT[name];
    const head = columns.map(c => {
      const num = ['int', 'eur', 'price', 'pct', 'diff'].includes(c.type);
      const arrow = !opts.printing && s && s[0] === c.key ? (s[1] > 0 ? ' ▲' : ' ▼') : '';
      return `<th class="${num ? 'num' : ''}" data-sort="${name}:${c.key}">${esc(c.label)}${arrow}</th>`;
    }).join('');
    const limit = opts.limit || Infinity;
    const shown = rows.slice(0, limit);
    const body = shown.map(r => '<tr>' + columns.map(c => {
      const v = r[c.key];
      const num = ['int', 'eur', 'price', 'pct', 'diff'].includes(c.type);
      if (c.type === 'status') return `<td><span class="badge ${STATUS_CLASS[v] || ''}">${esc(v)}</span></td>`;
      if (c.bar) return `<td class="num"><span class="pctbar"><span style="width:${Math.min(100, Math.max(0, (v || 0) * 100)).toFixed(1)}%"></span></span>${fmt(v, 'pct')}</td>`;
      const cls = (num ? 'num' : '') + (c.type === 'diff' && v ? (v > 0 ? ' neg' : ' pos') : '');
      return `<td class="${cls}">${fmt(v, c.type)}</td>`;
    }).join('') + '</tr>').join('');
    let foot = '';
    if (opts.totals && rows.length > 1) {
      foot = '<tfoot><tr>' + columns.map((c, i) => {
        if (i === 0) return '<td>Общо</td>';
        if (['int', 'eur'].includes(c.type)) return `<td class="num">${fmt(sumBy(rows, r => r[c.key] || 0), c.type)}</td>`;
        if (c.bar) {
          const ov = sumBy(rows, r => r.orderValue || 0), dv = sumBy(rows, r => r.onOrderValue || 0);
          return `<td class="num">${fmt(ov ? dv / ov : 0, 'pct')}</td>`;
        }
        return '<td></td>';
      }).join('') + '</tr></tfoot>';
    }
    const more = rows.length > shown.length ? `<p class="more">Показани са ${fInt.format(shown.length)} от ${fInt.format(rows.length)} реда: използвайте търсенето или експорта в Excel, за да видите всички.</p>` : '';
    return `<div class="table-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody>${foot}</table></div>${more}`;
  }

  function searchFilter(name, rows) {
    const q = (state.ui.search[name] || '').trim().toLowerCase();
    if (!q) return rows;
    const terms = q.split(/\s+/);
    return rows.filter(r => {
      const hay = [r.ean, r.article, r.name, r.color, r.sizeUS, r.size, r.po, r.bu, r.file, r.files, r.truck, r.packing].join(' ').toLowerCase();
      return terms.every(t => hay.includes(t));
    });
  }

  // ---------------------------------------------------------------------------
  // Risultati
  // ---------------------------------------------------------------------------
  const TABS = [
    { id: 'riepilogo', label: 'Обобщение', group: 'riepilogo' },
    { id: 'consegne', label: 'Доставки', group: 'consegne', count: v => v.shipments.length },
    { id: 'eccedenze', label: 'Излишъци', group: 'anomalie', count: v => v.excess.length, alert: true },
    { id: 'nonordinato', label: 'Непоръчано', group: 'anomalie', count: v => v.nonOrdered.length, alert: true },
    { id: 'altrastagione', label: 'Друг сезон', group: 'anomalie', count: v => v.otherSeason.length, alert: true },
    { id: 'differenze', label: 'Разлики цена/поръчка', group: 'anomalie', count: v => v.differences.length, alert: true },
    { id: 'articoli', label: 'По артикул', group: 'articoli' },
    { id: 'ean', label: 'Детайли по EAN', group: 'ean' },
  ];

  function kpi(label, q, val, pctOf, cls, note) {
    return `<div class="kpi ${cls || ''}">
      <div class="kpi-label">${label}</div>
      <div class="kpi-val">${fInt.format(q)} <small>бр.</small></div>
      <div class="kpi-sub">${fEur.format(val)} €${pctOf ? ' · ' + fPct.format(pctOf ? val / pctOf : 0) : ''}</div>
      ${note ? `<div class="kpi-note">${note}</div>` : ''}
    </div>`;
  }

  function renderKpis(v) {
    const t = v.t, ov = t.ordered.value;
    const pct = ov ? t.onOrder.value / ov : 0;
    $('#kpis').innerHTML = `
      <div class="progress-card">
        <div class="progress-head"><span>Изпълнение на доставките (стойност)</span><strong>${fPct.format(pct)}</strong></div>
        <div class="progress"><span style="width:${(pct * 100).toFixed(1)}%"></span></div>
        <div class="progress-foot">Доставено ${fEur.format(t.onOrder.value)} € от ${fEur.format(ov)} € поръчани · в бройки ${fPct.format(t.ordered.qty ? t.onOrder.qty / t.ordered.qty : 0)}</div>
      </div>
      <div class="kpi-grid">
        ${kpi('Поръчано', t.ordered.qty, t.ordered.value, 0, 'k-ord')}
        ${kpi('Доставено по поръчката', t.onOrder.qty, t.onOrder.value, ov, 'k-ok')}
        ${kpi('Остава за доставка', t.residual.qty, t.residual.value, ov, 'k-res')}
        ${kpi('Доставено в излишък', t.excess.qty, t.excess.value, 0, t.excess.qty ? 'k-bad' : 'k-zero')}
        ${kpi('Доставено непоръчано', t.nonOrdered.qty, t.nonOrdered.value, 0, t.nonOrdered.qty ? 'k-bad' : 'k-zero')}
        ${kpi('Доставено от друг сезон', t.otherSeason.qty, t.otherSeason.value, 0, t.otherSeason.qty ? 'k-bad' : 'k-zero')}
      </div>
      <p class="kpi-foot">Общо получено във всички файлове: ${fInt.format(t.all.qty)} бр., ${fEur.format(t.all.value)} € по цена FPC${v.hasNet ? `, ${fEur.format(t.all.net)} € нетно фактурирано` : ''}. Стойности по цена FPC без ДДС от поръчката; съпоставяне по EAN и сезон.</p>`;
  }

  function panel(id, group, title, inner, tools) {
    return `<section class="panel" data-panel="${id}" data-group="${group}">
      <h2 class="panel-title">${title}</h2>
      ${tools ? `<div class="panel-tools no-print">${tools}</div>` : ''}
      ${inner}
    </section>`;
  }

  const searchBox = (name, ph) => `<input type="search" class="search" data-search="${name}" placeholder="${ph || 'Търсене по артикул, EAN, описание…'}" value="${esc(state.ui.search[name] || '')}">`;

  function emptyOk(text) { return `<p class="ok-note">✓ ${text}</p>`; }

  function renderPanels(printing) {
    const v = view();
    const data = tableData(v);
    renderKpis(v);

    // Schede
    $('#tabs').innerHTML = TABS.map(t => {
      const n = t.count ? t.count(v) : null;
      const badge = n !== null ? `<span class="count${t.alert && n ? ' alert' : ''}">${fInt.format(n)}</span>` : '';
      return `<button role="tab" class="tab${state.tab === t.id ? ' active' : ''}" data-tab="${t.id}">${t.label}${badge}</button>`;
    }).join('');

    const T = (name, rows, opts) => htmlTable(name, cols(name, v), sortRows(name, printing ? rows : searchFilter(name, rows)), Object.assign({ printing }, opts));
    const lim = printing ? {} : { limit: 1500 };

    const multiSeason = data.season.length > 1;
    const riepilogo =
      (multiSeason ? '<h3>По сезон</h3>' + T('season', data.season, { totals: true }) : '') +
      '<h3>По отдел</h3>' + T('bu', data.bu, { totals: true }) +
      '<h3>По номер на поръчка</h3>' + T('po', data.po, { totals: true }) +
      (data.gender.some(g => g.gender) ? '<h3>По пол</h3>' + T('gender', data.gender, { totals: true }) : '');

    const anomalyNote = (rows, txt) => rows.length ? '' : emptyOk(txt);

    let articleRows = data.article;
    if (printing || state.ui.onlyOpen) articleRows = articleRows.filter(a => a.residual > 0 || a.excess > 0);
    let eanRows = data.ean;
    if (state.ui.status && !printing) eanRows = eanRows.filter(r => r.status === state.ui.status);

    const statuses = ['Предстои доставка', 'Частично', 'Изпълнено', 'Излишък'];
    const counts = Object.fromEntries(statuses.map(s => [s, data.ean.filter(r => r.status === s).length]));

    const html = [
      panel('riepilogo', 'riepilogo', 'Обобщение', riepilogo),
      panel('consegne', 'consegne', 'Получени доставки', T('shipments', data.shipments, { totals: true })),
      panel('eccedenze', 'anomalie', 'Доставено над поръчаното количество',
        anomalyNote(data.excess, 'Няма размер, доставен в количество над поръчаното.') + (data.excess.length ? T('excess', data.excess, { totals: true }) : ''),
        data.excess.length ? searchBox('excess') : ''),
      panel('nonordinato', 'anomalie', 'Доставено, но не е поръчано',
        anomalyNote(data.nonOrdered, 'Всички доставени EAN кодове присъстват в поръчката.') + (data.nonOrdered.length ? T('nonOrdered', data.nonOrdered, { totals: true }) : ''),
        data.nonOrdered.length ? searchBox('nonOrdered') : ''),
      panel('altrastagione', 'anomalie', 'Доставено с различен сезон от поръчания',
        '<p class="hint">Артикули от поръчката, пристигнали с друг сезон: не се отчитат като доставени по поръчката.</p>' +
        anomalyNote(data.otherSeason, 'Няма артикули, доставени с различен сезон.') + (data.otherSeason.length ? T('otherSeason', data.otherSeason, { totals: true }) : ''),
        data.otherSeason.length ? searchBox('otherSeason') : ''),
      panel('differenze', 'anomalie', 'Разлики в цената или в номера на поръчката',
        anomalyNote(data.differences, 'Цените FPC и номерата на поръчките в доставките съвпадат с поръчката.') + (data.differences.length ? T('differences', data.differences) : ''),
        data.differences.length ? searchBox('differences') : ''),
      panel('articoli', 'articoli', printing ? 'Остатък по артикул' : 'По артикул',
        T('article', articleRows, Object.assign({ totals: true }, lim)),
        searchBox('article') + `<label class="chk"><input type="checkbox" id="only-open"${state.ui.onlyOpen ? ' checked' : ''}> Само артикули с остатък или излишък</label>`),
      panel('ean', 'ean', 'Детайли по EAN',
        T('ean', eanRows, Object.assign({ totals: true }, lim)),
        searchBox('ean') + `<select id="status-filter"><option value="">Всички статуси (${fInt.format(data.ean.length)})</option>${statuses.map(s => `<option${state.ui.status === s ? ' selected' : ''} value="${s}">${s} (${fInt.format(counts[s])})</option>`).join('')}</select>`),
    ].join('');
    $('#panels').innerHTML = html;
    document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.dataset.panel === state.tab));
  }

  function renderSeasonFilter() {
    const seasons = Array.from(new Set(state.result.detail.map(d => d.season).filter(Boolean))).sort();
    if (state.season && !seasons.includes(state.season)) state.season = '';
    $('#season-filter').innerHTML = `<option value="">Всички</option>` + seasons.map(s => `<option${s === state.season ? ' selected' : ''}>${esc(s)}</option>`).join('');
    $('#season-filter').parentElement.hidden = seasons.length < 2;
  }

  function refresh() {
    renderUploads();
    compute();
    const has = !!state.result;
    $('#results').hidden = !has;
    $('#empty').hidden = has;
    $('#btn-excel').disabled = !has;
    $('#btn-print').disabled = !has;
    if (!has) return;
    renderSeasonFilter();
    renderPanels(false);
    const old = document.getElementById('dup-warning');
    if (old) old.remove();
    if (state.result.warnings.length) {
      const w = state.result.warnings;
      message('info', `Внимание: ${w.length} ${w.length === 1 ? 'дублиран ред' : 'дублирани реда'} в потвърждението. ${esc(w.slice(0, 3).join(' '))}${w.length > 3 ? '…' : ''}`);
      $('#messages').lastElementChild.id = 'dup-warning';
    }
  }

  // ---------------------------------------------------------------------------
  // Export Excel
  // ---------------------------------------------------------------------------
  const XL_FMT = { int: '#,##0', eur: '#,##0.00', price: '#,##0.00', diff: '+#,##0.00;-#,##0.00;0', pct: '0.0%' };

  function sheetFromTable(columns, rows, preface) {
    const aoa = (preface || []).concat([columns.map(c => c.label)]);
    const start = aoa.length;
    for (const r of rows) aoa.push(columns.map(c => {
      const v = r[c.key];
      if (c.type === 'bool') return v ? 'Да' : '';
      if (v === null || v === undefined) return '';
      if (c.key === 'ean' && /^\d+$/.test(String(v))) return String(v);
      if (typeof v === 'number' && ['eur', 'price', 'diff'].includes(c.type)) return E.round2(v);
      return v;
    }));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    columns.forEach((c, ci) => {
      const z = XL_FMT[c.type];
      if (!z) return;
      for (let ri = start; ri < aoa.length; ri++) {
        const cell = ws[XLSX.utils.encode_cell({ r: ri, c: ci })];
        if (cell && cell.t === 'n') cell.z = z;
      }
    });
    ws['!cols'] = columns.map(c => ({ wch: Math.max(10, Math.min(40, c.label.length + 2, ...rows.slice(0, 200).map(r => String(r[c.key] == null ? '' : r[c.key]).length + 2))) }));
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: start - 1, c: 0 }, e: { r: Math.max(start - 1, aoa.length - 1), c: columns.length - 1 } }) };
    return ws;
  }

  function summarySheet(v, data) {
    const t = v.t, ov = t.ordered.value || 1;
    const R = E.round2;
    const aoa = [
      ['Контрол на доставките Nike'],
      ['Поръчка', state.order.name],
      ['Доставки', state.deliveries.filter(d => d.table).map(d => d.name).join(', ')],
      ['Сезон', state.season || 'Всички'],
      ['Създаден на', new Date().toLocaleString('bg-BG')],
      [],
      ['Показател', 'Бройки', 'Стойност FPC €', '% от поръчаното (стойност)'],
      ['Поръчано (потвърдено)', t.ordered.qty, R(t.ordered.value), 1],
      ['Доставено по поръчката', t.onOrder.qty, R(t.onOrder.value), t.onOrder.value / ov],
      ['Остава за доставка', t.residual.qty, R(t.residual.value), t.residual.value / ov],
      ['Доставено в излишък', t.excess.qty, R(t.excess.value), t.excess.value / ov],
      ['Доставено НЕпоръчано', t.nonOrdered.qty, R(t.nonOrdered.value), t.nonOrdered.value / ov],
      ['Доставено от друг сезон', t.otherSeason.qty, R(t.otherSeason.value), t.otherSeason.value / ov],
      ['Общо получено (всички файлове)', t.all.qty, R(t.all.value), t.all.value / ov],
    ];
    if (v.hasNet) aoa.push(['Нетна фактурирана стойност (всички файлове)', '', R(t.all.net), '']);
    const blocks = [['По сезон', 'season'], ['По отдел', 'bu'], ['По номер на поръчка', 'po'], ['По пол', 'gender']];
    const fmts = [];
    for (let r = 7; r < aoa.length; r++) fmts.push([r, 1, 'int'], [r, 2, 'eur'], [r, 3, 'pct']);
    for (const [title, name] of blocks) {
      const rows = data[name];
      if (name === 'gender' && !rows.some(g => g.gender)) continue;
      const columns = cols(name, v);
      aoa.push([], [title], columns.map(c => c.label));
      for (const r of sortRows(name, rows)) {
        const ri = aoa.length;
        aoa.push(columns.map(c => (typeof r[c.key] === 'number' && c.type === 'eur' ? R(r[c.key]) : r[c.key])));
        columns.forEach((c, ci) => XL_FMT[c.type] && fmts.push([ri, ci, c.type]));
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (const [r, c, type] of fmts) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.t === 'n') cell.z = XL_FMT[type];
    }
    ws['!cols'] = [{ wch: 34 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 14 }];
    return ws;
  }

  function exportExcel() {
    const v = view();
    const data = tableData(v);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, summarySheet(v, data), 'Обобщение');
    const add = (title, name, rows) => XLSX.utils.book_append_sheet(wb, sheetFromTable(cols(name, v), sortRows(name, rows)), title);
    add('Доставки', 'shipments', data.shipments);
    add('Излишъци', 'excess', data.excess);
    add('Непоръчано', 'nonOrdered', data.nonOrdered);
    add('Друг сезон', 'otherSeason', data.otherSeason);
    add('Разлики цена-поръчка', 'differences', data.differences);
    add('Остатък по артикул', 'article', data.article.filter(a => a.residual > 0 || a.excess > 0));
    add('По артикул', 'article', data.article);
    add('Детайли по EAN', 'ean', data.ean);
    const seasons = state.season || Array.from(new Set(v.detail.map(d => d.season).filter(Boolean))).join('-');
    const d = new Date();
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    XLSX.writeFile(wb, `kontrol_dostavki${seasons ? '_' + seasons : ''}_${stamp}.xlsx`);
  }

  // ---------------------------------------------------------------------------
  // Stampa
  // ---------------------------------------------------------------------------
  const PRINT_GROUPS = ['riepilogo', 'consegne', 'anomalie', 'articoli', 'ean'];

  function printMeta() {
    const ds = state.deliveries.filter(d => d.table);
    $('#print-meta').innerHTML =
      `Поръчка: ${esc(state.order.name)}<br>` +
      `Анализирани доставки (${ds.length}): ${ds.map(d => esc(d.name)).join(', ')}<br>` +
      `${state.season ? 'Сезон: ' + esc(state.season) + ' · ' : ''}Създаден на ${new Date().toLocaleString('bg-BG')}. Стойности по цена FPC без ДДС; съпоставяне по EAN и сезон.`;
  }

  function startPrint(groups) {
    printMeta();
    PRINT_GROUPS.forEach(g => document.body.classList.toggle('print-' + g, groups.includes(g)));
    renderPanels(true);
    document.body.classList.add('printing');
    window.print();
  }
  function endPrint() {
    if (!document.body.classList.contains('printing')) return;
    document.body.classList.remove('printing');
    renderPanels(false);
  }
  window.addEventListener('afterprint', endPrint);
  // Stampa da menu del browser (Ctrl+P): stampa il riepilogo e le anomalie.
  window.addEventListener('beforeprint', () => {
    if (state.result && !document.body.classList.contains('printing')) {
      printMeta();
      PRINT_GROUPS.forEach(g => document.body.classList.toggle('print-' + g, ['riepilogo', 'consegne', 'anomalie', 'articoli'].includes(g)));
      renderPanels(true);
      document.body.classList.add('printing');
    }
  });

  // ---------------------------------------------------------------------------
  // Eventi
  // ---------------------------------------------------------------------------
  function setupDrop(labelSel, inputSel, handler) {
    const label = $(labelSel), input = $(inputSel);
    input.addEventListener('change', () => { if (input.files.length) handler(Array.from(input.files)); input.value = ''; });
    ['dragenter', 'dragover'].forEach(ev => label.addEventListener(ev, e => { e.preventDefault(); label.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => label.addEventListener(ev, e => { e.preventDefault(); label.classList.remove('over'); }));
    label.addEventListener('drop', e => { const f = Array.from(e.dataTransfer.files || []); if (f.length) handler(f); });
  }

  function wrap(fn) {
    return async (...args) => {
      try { await fn(...args); }
      catch (e) { console.error(e); message('error', 'Грешка при четене на файла: ' + esc(e.message || e)); }
    };
  }

  setupDrop('#drop-order', '#file-order', wrap(files => addOrderFile(files[0])));
  setupDrop('#drop-deliveries', '#file-deliveries', wrap(files => addDeliveryFiles(files)));

  $('#clear-order').onclick = wrap(async () => { if (state.order) await dbDel(state.order.id); state.order = null; refresh(); });
  $('#clear-deliveries').onclick = wrap(async () => {
    if (!confirm('Да се премахнат ли всички файлове с доставки?')) return;
    for (const d of state.deliveries) await dbDel(d.id);
    state.deliveries = []; refresh();
  });

  document.addEventListener('click', wrap(async e => {
    const rm = e.target.closest('[data-remove]');
    if (rm) {
      const id = rm.dataset.remove;
      await dbDel(id);
      state.deliveries = state.deliveries.filter(d => d.id !== id);
      refresh();
      return;
    }
    const tab = e.target.closest('[data-tab]');
    if (tab) {
      state.tab = tab.dataset.tab;
      document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === state.tab));
      document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.dataset.panel === state.tab));
      return;
    }
    const th = e.target.closest('th[data-sort]');
    if (th) {
      const [name, key] = th.dataset.sort.split(':');
      const cur = state.ui.sort[name] || DEFAULT_SORT[name];
      const numeric = ['int', 'eur', 'price', 'pct', 'diff'].includes((COLS[name].find(c => c.key === key) || {}).type);
      state.ui.sort[name] = [key, cur && cur[0] === key ? -cur[1] : (numeric ? -1 : 1)];
      renderPanels(false);
    }
  }));

  let searchTimer = null;
  document.addEventListener('input', e => {
    const s = e.target.closest('[data-search]');
    if (!s) return;
    state.ui.search[s.dataset.search] = s.value;
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      const name = s.dataset.search, pos = s.selectionStart;
      renderPanels(false);
      const again = document.querySelector(`[data-search="${name}"]`);
      if (again) { again.focus(); again.setSelectionRange(pos, pos); }
    }, 200);
  });

  document.addEventListener('change', wrap(async e => {
    const t = e.target;
    if (t.id === 'season-filter') { state.season = t.value; renderPanels(false); return; }
    if (t.id === 'status-filter') { state.ui.status = t.value; renderPanels(false); return; }
    if (t.id === 'only-open') { state.ui.onlyOpen = t.checked; renderPanels(false); return; }
    if (t.dataset.field) {
      const role = t.dataset.role, id = t.dataset.id;
      const f = role === 'order' ? state.order : state.deliveries.find(d => d.id === id);
      if (!f) return;
      f.overrides = Object.assign({}, f.overrides, { [t.dataset.field]: Number(t.value) });
      const reloaded = role === 'order' ? loadOrder(f) : loadDelivery(f);
      if (role === 'order') state.order = reloaded;
      else state.deliveries = state.deliveries.map(d => (d.id === id ? reloaded : d));
      await persist(reloaded, role);
      refresh();
    }
  }));

  $('#btn-excel').onclick = wrap(exportExcel);
  $('#btn-print').onclick = () => {
    const dlg = $('#print-dialog');
    if (typeof dlg.showModal === 'function') dlg.showModal();
    else startPrint(['riepilogo', 'consegne', 'anomalie', 'articoli']);
  };
  $('#print-dialog').addEventListener('close', () => {
    const dlg = $('#print-dialog');
    if (dlg.returnValue !== 'print') return;
    const groups = Array.from(dlg.querySelectorAll('input[name=sec]:checked')).map(i => i.value);
    setTimeout(() => startPrint(groups), 50);
  });

  restore();
})();
