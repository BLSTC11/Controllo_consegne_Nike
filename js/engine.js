/*
 * Motore di confronto: legge i fogli (tramite SheetJS), riconosce le colonne
 * e confronta l'ordine confermato con tutte le consegne (EAN + Stagione).
 * Nessun accesso alla rete: tutto avviene nel browser.
 */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------------------
  // Riconoscimento colonne
  // ---------------------------------------------------------------------------

  // Per ogni campo: elenco di regex (in ordine di priorità) da provare sulle
  // intestazioni normalizzate (minuscolo, solo lettere/numeri e spazi).
  const ORDER_FIELDS = {
    ean:     { label: 'EAN', required: true, rx: [/^ean upc cd$/, /^ean upc$/, /^ean$/, /\bean\b/, /\bupc\b/, /\bgtin\b/, /barcode/, /codice a barre/] },
    qty:     { label: 'Quantità confermata', required: true, rx: [/^qty status nnt$/, /\bnnt\b/, /confirm/, /conferm/, /^qty order entry$/, /order entry/, /^qty$/, /^quantity$/, /^quantita$/, /^pcs$/, /\bqty\b(?!.*reject)/] },
    price:   { label: 'Prezzo FPC', required: true, rx: [/^fpc price in eur w o vat$/, /\bfpc\b/, /wholesale/, /prezzo/, /\bprice\b/] },
    season:  { label: 'Stagione', rx: [/^season$/, /^stagione$/, /season/, /stagione/] },
    po:      { label: 'N. ordine (PO)', rx: [/^po name sap$/, /^po name$/, /\bpo\b/, /purchase order/, /ordine/] },
    bu:      { label: 'Reparto (BU)', rx: [/^bu$/, /business unit/, /division/, /reparto/] },
    article: { label: 'Articolo', rx: [/^article number$/, /^material$/, /^article$/, /style color/, /^articolo$/, /article n/] },
    name:    { label: 'Descrizione', rx: [/^article name$/, /^description$/, /^material description$/, /descri/, /name/] },
    color:   { label: 'Colore', rx: [/^color desc$/, /colou?r/, /colore/] },
    sizeUS:  { label: 'Taglia US', rx: [/^size us$/, /^size$/, /^us size$/, /^taglia$/] },
    sizeEU:  { label: 'Taglia EU', rx: [/^eu size$/, /^size eu$/] },
    gender:  { label: 'Genere', rx: [/^gender$/, /genere/] },
  };

  const DELIVERY_FIELDS = {
    ean:      { label: 'EAN', required: true, rx: [/^ean upc cd$/, /^ean upc$/, /^ean$/, /\bean\b/, /\bupc\b/, /\bgtin\b/, /barcode/, /codice a barre/] },
    qty:      { label: 'Quantità consegnata', required: true, rx: [/^dlv qty$/, /^delivery qty$/, /^qty$/, /^quantity$/, /^quantita$/, /shipped/, /delivered/, /invoiced/, /^pcs$/, /pieces/, /pezzi/, /\bqty\b(?!.*(order|confirm|reject))/, /quant/] },
    price:    { label: 'Prezzo FPC', rx: [/^fpc price in eur w o vat$/, /\bfpc\b/, /wholesale/, /^price$/, /unit price/, /prezzo/] },
    netValue: { label: 'Valore netto (totale riga)', rx: [/net value/, /net amount/, /valore netto/, /total net/, /net total/, /^amount$/, /^value$/] },
    netPrice: { label: 'Prezzo netto unitario', rx: [/net price/, /prezzo netto/] },
    season:   { label: 'Stagione', rx: [/^season$/, /^stagione$/, /season/, /stagione/] },
    po:       { label: 'N. ordine (PO)', rx: [/^po name sap$/, /^po name$/, /\bpo\b/, /purchase order/, /customer order/, /ordine/] },
    truck:    { label: 'Truck', rx: [/truck/, /camion/, /shipment/] },
    packing:  { label: 'Packing list', rx: [/packing/, /delivery n/, /^delivery$/, /dn number/] },
    invoice:  { label: 'Fattura', rx: [/invoice n/, /^invoice$/, /fattura/] },
    article:  { label: 'Articolo', rx: [/^article number$/, /^material$/, /^article$/, /style color/, /^articolo$/, /article n/, /material n/] },
    name:     { label: 'Descrizione', rx: [/^article name$/, /^description$/, /^material description$/, /descri/, /name/] },
    size:     { label: 'Taglia', rx: [/^size us$/, /^size$/, /^us size$/, /^taglia$/, /size/] },
    bu:       { label: 'Reparto (BU)', rx: [/^bu$/, /business unit/, /division/, /reparto/] },
  };

  function normHeader(h) {
    return String(h == null ? '' : h)
      .toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  // Assegna le colonne ai campi: per ogni campo prova le regex in ordine e
  // prende la prima colonna libera che corrisponde.
  function detectColumns(headers, fields) {
    const norm = headers.map(normHeader);
    const used = new Set();
    const map = {};
    // Prima i campi obbligatori, poi gli altri (nell'ordine dichiarato).
    const keys = Object.keys(fields).sort((a, b) => (fields[b].required ? 1 : 0) - (fields[a].required ? 1 : 0));
    for (const key of keys) {
      map[key] = -1;
      outer: for (const rx of fields[key].rx) {
        for (let c = 0; c < norm.length; c++) {
          if (!norm[c] || used.has(c)) continue;
          if (rx.test(norm[c])) { map[key] = c; used.add(c); break outer; }
        }
      }
    }
    return map;
  }

  // Trova la riga di intestazione: quella (tra le prime 30) con più campi
  // riconosciuti, purché contenga EAN e quantità.
  function findHeaderRow(rows, fields) {
    let best = { row: -1, score: 0, map: null };
    const limit = Math.min(rows.length, 30);
    for (let r = 0; r < limit; r++) {
      const cells = rows[r] || [];
      if (cells.filter(v => v !== null && v !== undefined && String(v).trim() !== '').length < 2) continue;
      // Una riga di intestazione è fatta di testo, non di numeri.
      const textCells = cells.filter(v => typeof v === 'string' && /[a-zA-Z]/.test(v)).length;
      if (textCells < 2) continue;
      const map = detectColumns(cells, fields);
      const score = Object.values(map).filter(c => c >= 0).length;
      const ok = Object.keys(fields).every(k => !fields[k].required || k === 'price' || map[k] >= 0);
      if (ok && score > best.score) best = { row: r, score, map };
    }
    return best;
  }

  // Sceglie il foglio migliore della cartella di lavoro.
  function readWorkbookTable(XLSX, wb, fields) {
    let best = null;
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
      const h = findHeaderRow(rows, fields);
      if (h.row >= 0 && (!best || h.score > best.score)) {
        best = { sheet: name, headerRow: h.row, headers: (rows[h.row] || []).map(v => (v == null ? '' : String(v).trim())), rows, map: h.map, score: h.score };
      }
    }
    return best;
  }

  // ---------------------------------------------------------------------------
  // Normalizzazioni
  // ---------------------------------------------------------------------------

  function eanKey(v) {
    if (v === null || v === undefined || v === '') return '';
    let s = typeof v === 'number' ? Math.round(v).toString() : String(v).trim();
    if (/e\+/i.test(s) && !isNaN(Number(s))) s = Math.round(Number(s)).toString();
    s = s.replace(/\D/g, '').replace(/^0+/, '');
    return s;
  }

  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let s = String(v).trim().replace(/[€\s]/g, '');
    if (!s) return null;
    // "1.234,56" -> 1234.56 ; "1,234.56" -> 1234.56 ; "12,5" -> 12.5
    if (s.includes(',') && s.includes('.')) {
      if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
      else s = s.replace(/,/g, '');
    } else if (s.includes(',')) {
      s = s.replace(',', '.');
    }
    const n = Number(s);
    return isFinite(n) ? n : null;
  }

  function str(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  const SEASON_DIGIT = { 1: 'SP', 2: 'SU', 3: 'FA', 4: 'HO' };

  // Riporta i codici stagione a una forma comune (es. FA26).
  // "263" (codice numerico Nike: anno 26 + stagione 3) -> FA26.
  function normSeason(v, overrides) {
    const raw = str(v).toUpperCase();
    if (!raw) return '';
    if (overrides && overrides[raw]) return overrides[raw];
    let m = raw.match(/^(SP|SU|FA|HO)\s*[-_/ ]?\s*(?:20)?(\d{2})$/);
    if (m) return m[1] + m[2];
    m = raw.match(/^(\d{2})(\d)$/);
    if (m && SEASON_DIGIT[m[2]]) return SEASON_DIGIT[m[2]] + m[1];
    m = raw.match(/^(?:20)?(\d{2})\s*[-_/ ]?\s*(SP|SU|FA|HO)$/);
    if (m) return m[2] + m[1];
    return raw;
  }

  function seasonFromTruck(truck) {
    const m = str(truck).toUpperCase().match(/(SP|SU|FA|HO)\s*(\d{2})\s*$/);
    return m ? m[1] + m[2] : '';
  }

  function normPO(v) { return str(v).toUpperCase().replace(/\s+/g, ' '); }

  function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

  // ---------------------------------------------------------------------------
  // Lettura file
  // ---------------------------------------------------------------------------

  function get(row, map, key) {
    const c = map[key];
    return c >= 0 ? row[c] : null;
  }

  function parseOrder(table, opts) {
    opts = opts || {};
    const map = table.map;
    const lines = [];
    const rawSeasons = {};
    for (let r = table.headerRow + 1; r < table.rows.length; r++) {
      const row = table.rows[r] || [];
      const key = eanKey(get(row, map, 'ean'));
      if (!key) continue;
      const qty = num(get(row, map, 'qty'));
      if (qty === null) continue;
      const rawSeason = str(get(row, map, 'season'));
      const season = normSeason(rawSeason, opts.seasonOverrides);
      if (rawSeason) rawSeasons[rawSeason.toUpperCase()] = season;
      lines.push({
        key,
        ean: str(get(row, map, 'ean')),
        qty,
        price: num(get(row, map, 'price')) || 0,
        season,
        po: str(get(row, map, 'po')),
        bu: str(get(row, map, 'bu')).toUpperCase(),
        article: str(get(row, map, 'article')),
        name: str(get(row, map, 'name')),
        color: str(get(row, map, 'color')),
        sizeUS: str(get(row, map, 'sizeUS')),
        sizeEU: str(get(row, map, 'sizeEU')),
        gender: str(get(row, map, 'gender')).toUpperCase(),
        row: r + 1,
      });
    }
    return { lines, rawSeasons };
  }

  function parseDelivery(table, fileName, opts) {
    opts = opts || {};
    const map = table.map;
    const lines = [];
    const rawSeasons = {};
    for (let r = table.headerRow + 1; r < table.rows.length; r++) {
      const row = table.rows[r] || [];
      const key = eanKey(get(row, map, 'ean'));
      if (!key) continue;
      const qty = num(get(row, map, 'qty'));
      if (qty === null || qty === 0) continue;
      const truck = str(get(row, map, 'truck'));
      const rawSeason = str(get(row, map, 'season'));
      let season = normSeason(rawSeason, opts.seasonOverrides);
      if (rawSeason) rawSeasons[rawSeason.toUpperCase()] = season;
      // Se il codice non è riconoscibile, prova dal nome del camion ("…/FA26").
      const fromTruck = seasonFromTruck(truck);
      if ((!season || !/^(SP|SU|FA|HO)\d{2}$/.test(season)) && fromTruck) {
        if (rawSeason) rawSeasons[rawSeason.toUpperCase()] = fromTruck;
        season = fromTruck;
      }
      const price = num(get(row, map, 'price'));
      let netValue = num(get(row, map, 'netValue'));
      const netPrice = num(get(row, map, 'netPrice'));
      if (netValue === null && netPrice !== null) netValue = netPrice * qty;
      lines.push({
        key,
        ean: str(get(row, map, 'ean')),
        qty,
        price,
        netValue,
        season,
        rawSeason,
        po: str(get(row, map, 'po')),
        truck,
        packing: str(get(row, map, 'packing')),
        invoice: str(get(row, map, 'invoice')),
        article: str(get(row, map, 'article')),
        name: str(get(row, map, 'name')),
        size: str(get(row, map, 'size')),
        bu: str(get(row, map, 'bu')).toUpperCase(),
        file: fileName,
        row: r + 1,
      });
    }
    return { lines, rawSeasons };
  }

  // ---------------------------------------------------------------------------
  // Confronto
  // ---------------------------------------------------------------------------

  function compare(orderLines, deliveryLines) {
    const warnings = [];
    const byKey = new Map();
    for (const o of orderLines) {
      const ex = byKey.get(o.key);
      if (ex) {
        // Stesso EAN su più righe dell'ordine: si sommano le quantità.
        ex.qty += o.qty;
        if (ex.po !== o.po && !ex.poList.includes(o.po)) ex.poList.push(o.po);
        warnings.push(`EAN ${o.ean} presente più volte nella conferma (righe ${ex.row} e ${o.row}): quantità sommate.`);
        continue;
      }
      byKey.set(o.key, Object.assign({}, o, {
        poList: [o.po],
        delivered: 0, deliveredFpc: 0, deliveredNet: 0,
        delPrices: new Set(), delPOs: new Set(), files: new Set(), names: new Set(), sizes: new Set(),
      }));
    }

    const nonOrdered = new Map();   // key|season -> aggregato
    const otherSeason = new Map();  // key|season -> aggregato

    function addAnomaly(map, d, o) {
      const k = d.key + '|' + d.season;
      let a = map.get(k);
      if (!a) {
        a = {
          key: d.key, ean: d.ean, season: d.season, orderSeason: o ? o.season : '',
          article: d.article || (o ? o.article : ''), name: d.name || (o ? o.name : ''),
          color: o ? o.color : '', size: d.size || (o ? o.sizeUS : ''), bu: d.bu || (o ? o.bu : ''),
          po: d.po, orderPO: o ? o.po : '', orderQty: o ? o.qty : 0, orderPrice: o ? o.price : null,
          qty: 0, fpcValue: 0, netValue: 0, prices: new Set(), files: new Set(), trucks: new Set(),
        };
        map.set(k, a);
      }
      a.qty += d.qty;
      const p = d.price !== null ? d.price : (o ? o.price : null);
      if (p !== null) a.fpcValue += p * d.qty;
      if (d.netValue !== null) a.netValue += d.netValue;
      if (d.price !== null) a.prices.add(d.price);
      a.files.add(d.file);
      if (d.truck) a.trucks.add(d.truck);
    }

    const shipments = new Map();
    for (const d of deliveryLines) {
      const o = byKey.get(d.key);
      // Riepilogo per consegna (file + packing list + truck + PO + stagione)
      const sk = [d.file, d.packing, d.truck, normPO(d.po), d.season].join('\u0001');
      let s = shipments.get(sk);
      if (!s) {
        s = { file: d.file, packing: d.packing, truck: d.truck, po: d.po, season: d.season, qty: 0, fpcValue: 0, netValue: 0, lines: 0, onOrder: 0, notOrdered: 0, otherSeason: 0 };
        shipments.set(sk, s);
      }
      s.qty += d.qty; s.lines++;
      const p = d.price !== null ? d.price : (o ? o.price : 0);
      s.fpcValue += p * d.qty;
      if (d.netValue !== null) s.netValue += d.netValue;

      if (!o) { addAnomaly(nonOrdered, d, null); s.notOrdered += d.qty; continue; }
      if (d.season && o.season && d.season !== o.season) { addAnomaly(otherSeason, d, o); s.otherSeason += d.qty; continue; }
      s.onOrder += d.qty;
      o.delivered += d.qty;
      o.deliveredFpc += p * d.qty;
      if (d.netValue !== null) o.deliveredNet += d.netValue;
      if (d.price !== null) o.delPrices.add(round2(d.price));
      if (d.po) o.delPOs.add(d.po);
      o.files.add(d.file);
      if (d.name) o.names.add(d.name);
      if (d.size) o.sizes.add(d.size);
    }

    // Righe di dettaglio per EAN
    const detail = [];
    for (const o of byKey.values()) {
      const onOrder = Math.min(Math.max(o.delivered, 0), o.qty);
      const residual = Math.max(o.qty - o.delivered, 0);
      const excess = Math.max(o.delivered - o.qty, 0);
      let status;
      if (o.delivered <= 0) status = 'Da consegnare';
      else if (excess > 0) status = 'Eccedenza';
      else if (residual > 0) status = 'Parziale';
      else status = 'Completo';
      const delPrices = Array.from(o.delPrices);
      const priceDiff = delPrices.length ? round2(delPrices.reduce((m, p) => (Math.abs(p - o.price) > Math.abs(m) ? p - o.price : m), 0)) : null;
      const delPOs = Array.from(o.delPOs);
      const poDiff = delPOs.some(p => !o.poList.map(normPO).includes(normPO(p)));
      detail.push({
        status, season: o.season, po: o.po, bu: o.bu, article: o.article, name: o.name, color: o.color,
        sizeUS: o.sizeUS, sizeEU: o.sizeEU, gender: o.gender, ean: o.ean, key: o.key, price: o.price,
        qty: o.qty, delivered: o.delivered, onOrder, residual, excess,
        orderValue: o.qty * o.price, onOrderValue: onOrder * o.price, residualValue: residual * o.price, excessValue: excess * o.price,
        deliveredFpc: o.deliveredFpc, deliveredNet: o.deliveredNet,
        delPrice: delPrices.length === 1 ? delPrices[0] : delPrices.join(' / '), priceDiff: priceDiff && Math.abs(priceDiff) >= 0.01 ? priceDiff : (delPrices.length ? 0 : null),
        delPO: delPOs.join(' / '), poDiff, files: Array.from(o.files).join(', '),
        delName: Array.from(o.names).join(' / '), delSize: Array.from(o.sizes).join(' / '),
      });
    }

    const finishAnomaly = a => Object.assign(a, {
      prices: a.prices.size === 1 ? round2(Array.from(a.prices)[0]) : Array.from(a.prices).map(round2).join(' / '),
      files: Array.from(a.files).join(', '),
      trucks: Array.from(a.trucks).join(', '),
    });

    const result = {
      detail,
      nonOrdered: Array.from(nonOrdered.values()).map(finishAnomaly),
      otherSeason: Array.from(otherSeason.values()).map(finishAnomaly),
      shipments: Array.from(shipments.values()),
      warnings,
    };
    result.excess = detail.filter(d => d.excess > 0);
    result.differences = detail.filter(d => (d.priceDiff !== null && Math.abs(d.priceDiff) >= 0.01) || d.poDiff);
    result.totals = totals(result, deliveryLines);
    return result;
  }

  function sum(arr, f) { let s = 0; for (const x of arr) s += f(x); return s; }

  function totals(res, deliveryLines) {
    const d = res.detail;
    const ordered = { qty: sum(d, x => x.qty), value: sum(d, x => x.orderValue) };
    const t = {
      ordered,
      onOrder: { qty: sum(d, x => x.onOrder), value: sum(d, x => x.onOrderValue) },
      residual: { qty: sum(d, x => x.residual), value: sum(d, x => x.residualValue) },
      excess: { qty: sum(d, x => x.excess), value: sum(d, x => x.excessValue) },
      nonOrdered: { qty: sum(res.nonOrdered, x => x.qty), value: sum(res.nonOrdered, x => x.fpcValue) },
      otherSeason: { qty: sum(res.otherSeason, x => x.qty), value: sum(res.otherSeason, x => x.fpcValue) },
      allDelivered: { qty: sum(res.shipments, x => x.qty), value: sum(res.shipments, x => x.fpcValue), net: sum(res.shipments, x => x.netValue) },
    };
    t.hasNet = deliveryLines.some(l => l.netValue !== null);
    return t;
  }

  // Raggruppa le righe di dettaglio per una o più chiavi.
  function groupDetail(detail, keys) {
    const m = new Map();
    for (const d of detail) {
      const k = keys.map(key => d[key]).join('\u0001');
      let g = m.get(k);
      if (!g) {
        g = { qty: 0, delivered: 0, onOrder: 0, residual: 0, excess: 0, orderValue: 0, onOrderValue: 0, residualValue: 0, excessValue: 0, lines: 0 };
        keys.forEach(key => { g[key] = d[key]; });
        m.set(k, g);
      }
      g.lines++;
      for (const f of ['qty', 'delivered', 'onOrder', 'residual', 'excess', 'orderValue', 'onOrderValue', 'residualValue', 'excessValue']) g[f] += d[f];
    }
    const out = Array.from(m.values());
    out.forEach(g => { g.pct = g.orderValue ? g.onOrderValue / g.orderValue : 0; });
    return out;
  }

  function byArticle(detail) {
    const m = new Map();
    for (const d of detail) {
      const k = d.season + '\u0001' + d.article;
      let g = m.get(k);
      if (!g) {
        g = { season: d.season, bu: d.bu, article: d.article, name: d.name, color: d.color, price: d.price, qty: 0, onOrder: 0, residual: 0, excess: 0, orderValue: 0, onOrderValue: 0, residualValue: 0, excessValue: 0, sizes: 0, sizesDone: 0 };
        m.set(k, g);
      }
      g.qty += d.qty; g.onOrder += d.onOrder; g.residual += d.residual; g.excess += d.excess;
      g.orderValue += d.orderValue; g.onOrderValue += d.onOrderValue; g.residualValue += d.residualValue; g.excessValue += d.excessValue;
      g.sizes++; if (d.residual === 0) g.sizesDone++;
    }
    const out = Array.from(m.values());
    out.forEach(g => { g.pct = g.orderValue ? g.onOrderValue / g.orderValue : 0; });
    return out;
  }

  const api = {
    ORDER_FIELDS, DELIVERY_FIELDS,
    normHeader, detectColumns, findHeaderRow, readWorkbookTable,
    eanKey, num, normSeason, seasonFromTruck, normPO, round2,
    parseOrder, parseDelivery, compare, groupDetail, byArticle,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Engine = api;
})(typeof window !== 'undefined' ? window : globalThis);
