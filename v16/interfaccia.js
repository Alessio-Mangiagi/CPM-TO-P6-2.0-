(function() {
'use strict';

// 🛡️ STATE MANAGER (Memory safe, validated, localStorage sync)
class StateManager {
  constructor() {
    this._data = { budget: null, sil: null, silI: null, p6: null, result: null };
    const _rawThr = localStorage.getItem('bridge_threshold');
    const _parsedThr = +_rawThr;
    this.threshold = (Number.isFinite(_parsedThr) && _parsedThr >= 100) ? Math.min(_parsedThr, 10_000_000) : 5000;

    const devInput = document.getElementById('devThreshold');
    if (devInput) {
      devInput.value = this.threshold;
      devInput.addEventListener('change', (e) => {
        const v = +e.target.value;
        this.threshold = (Number.isFinite(v) && v >= 100) ? Math.min(v, 10_000_000) : 5000;
        localStorage.setItem('bridge_threshold', this.threshold);
      });
    }
  }
  get(key) { return this._data[key]; }
  set(key, val) { this._data[key] = val; }
  reset() { Object.keys(this._data).forEach(k => this._data[k] = null); }
  cleanup() { 
    // ✅ FIXED: Rimossa assegnazione a null che rompeva istanza
    this._data = { budget: null, sil: null, silI: null, p6: null, result: null }; 
  }
}
const S = new StateManager();
let _log = [];

const SHEET_HINTS = {
  budget: ['BUDGET','Budget','budget', 'BUDGET CPM'],
  sil:    ['SIL diretti','SIL Diretti','SIL_DIRETTI', 'SIL DIRETTI'],
  silI:   ['SIL indiretti','SIL Indiretti','SIL_INDIRETTI', 'SIL INDIRETTI'],
  p6:     ['EXPORT_P6','Export_P6','export_p6','EXPORT P6']
};
const COLUMNS = {
  budget: ['cod. wbs','articolo','importo costo', 'costo', 'budget', 'wbs', 'codice wbs'],
  sil:    ['cod. s.i.l.','articolo','importo', 'sil', 'costo', 'codice sil'],
  silI:   ['cod. s.i.l.','articolo','importo', 'sil', 'costo', 'codice sil'],
  p6:     ['task_code','wbs_id','act_cost','calc_phys_complete_pct','status_code','act_name']
};

// ── UTILITY ──────────────────────────────────────────────────────────────────
function logMsg(m) { _log.push(`[${new Date().toLocaleTimeString()}] ${m}`); }
function clearLog() { _log = []; }
function sanitizeForLog(s) {
  return String(s).replace(/[\r\n\t\x00-\x1F\x7F]/g, ' ').substring(0, 200);
}
function fmt(n) { return Math.round(n || 0).toLocaleString('it-IT'); }
function fmtE(n) { return (+n || 0).toLocaleString('it-IT', {minimumFractionDigits:2, maximumFractionDigits:2}); }
function fmtDelta(n) {
  const s = Math.abs(n).toLocaleString('it-IT', {minimumFractionDigits:2, maximumFractionDigits:2});
  if (n > 0.01) return `+${s}€`;
  if (n < -0.01) return `(${s}€)`;
  return '—';
}
function el(tag, cls) { 
  const e = document.createElement(tag); 
  if (cls) e.className = cls; 
  return e; 
}

// ✅ FIXED: Robust alphanumeric WBS normalizer
function normWBS(w) {
  if (!w) return '';
  let s = String(w).trim().toUpperCase();
  if (s === '') return '';
  return s.split('.').map(part => {
    const isNum = /^\d+$/.test(part);
    return isNum ? part.replace(/^0+(\d)/, '$1') || '0' : part;
  }).join('.');
}
function normArt(a) { return String(a || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function isDataRow(row) {
  if (!row || !row.length) return false;
  const f = String(row[0]).trim();
  if (!f) return false;
  if (/^[\u{1F300}-\u{1FFFF}]/u.test(f)) return false;
  if (/^(istruzione|commento|\*|note|totale)/i.test(f)) return false;
  return true;
}
function excelDateToStr(raw) {
  const n = parseFloat(raw);
  if (!isNaN(n) && n > 40000) {
    const d = new Date(Math.round((n - 25569) * 86400 * 1000));
    return d.toISOString().substring(0, 10);
  }
  return raw || '';
}
function validateColumns(headers, type) {
  const required = COLUMNS[type] || [];
  const found = required.filter(req => headers.some(h => h.toLowerCase().includes(req)));
  return found.length > 0;
}

// ✅ SECURITY: Sanitizza output errori per prevenire DOM XSS
function showError(msg) {
  const cleanMsg = String(msg).replace(/[<>]/g, '');
  const box = document.getElementById('bnrError');
  const txt = document.getElementById('bnrErrTxt') || box;
  if (box) {
    box.style.display = 'flex';
    box.classList.add('show');
    txt.textContent = '⚠️ ' + cleanMsg;
    setTimeout(() => { box.style.display = 'none'; box.classList.remove('show'); }, 6000);
  } else {
    console.error(cleanMsg);
  }
}
function showToast(type, msg) {
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  t.textContent = msg;
  document.body.appendChild(t);
  requestAnimationFrame(() => { t.classList.add('show'); });
  setTimeout(() => {
    t.classList.remove('show');
    setTimeout(() => t.remove(), 300);
  }, 3500);
}

// ── FILE PARSER ───────────────────────────────────────────────────────────────
function parseFile(file, key) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        let rows;
        const isXlsx = /\.(xlsx|xls)$/i.test(file.name);
        if (isXlsx) {
          if (typeof XLSX === 'undefined') throw new Error('Libreria SheetJS (XLSX) non caricata.');
          const wb = XLSX.read(e.target.result, { type: 'array' });
          const hints = SHEET_HINTS[key] || [];
          let sheetName = wb.SheetNames[0];
          for (const h of hints) {
            const found = wb.SheetNames.find(n => n.toLowerCase().includes(h.toLowerCase()));
            if (found) { sheetName = found; break; }
          }
          logMsg(`Foglio selezionato per ${key}: "${sanitizeForLog(sheetName)}"`);
          const ws = wb.Sheets[sheetName];
          rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        } else {
          const text = new TextDecoder('utf-8').decode(e.target.result);
          const sep = text.split('\n')[0].includes('\t') ? '\t' : ',';
          rows = text.split(/\r?\n/).map(l => l.split(sep).map(c => c.replace(/^"|"$/g, '').trim()));
        }
        if (key === 'p6') { resolve({ data: null, rawRows: rows }); return; }

        const signals = COLUMNS[key] || [];
        let hIdx = -1;
        for (let i = 0; i < Math.min(12, rows.length); i++) {
          const j = rows[i].map(c => String(c).toLowerCase()).join('|');
          if (signals.filter(s => j.includes(s)).length >= 1) { hIdx = i; break; }
        }
        if (hIdx === -1) hIdx = 0;

        const headers = rows[hIdx].map(h => String(h).trim());
        if (!validateColumns(headers, key)) logMsg(`⚠ Colonne non standard rilevate per ${key}`);

        const data = [];
        for (let i = hIdx + 1; i < rows.length; i++) {
          if (!isDataRow(rows[i])) continue;
          if (rows[i].every(c => String(c).trim() === '')) continue;
          const obj = {};
          headers.forEach((h, j) => { obj[h] = String(rows[i][j] ?? '').trim(); });
          data.push(obj);
        }
        const MAX_ROWS = 50_000;
        if (data.length > MAX_ROWS) {
          logMsg(`⚠ File troncato a ${MAX_ROWS} righe per sicurezza`);
          data.splice(MAX_ROWS);
        }
        resolve({ headers, data, rawRows: rows });
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ── DROP ZONES ────────────────────────────────────────────────────────────────
document.querySelectorAll('.dz').forEach(dz => {
  const key = dz.dataset.key;
  const inp = dz.querySelector('input[type=file]');
  dz.addEventListener('click', () => inp.click());
  dz.addEventListener('keydown', (e) => { if(e.key==='Enter'||e.key===' ') { e.preventDefault(); inp.click(); }});
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('over'));
  dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('over'); if(e.dataTransfer.files[0]) handleFile(key, e.dataTransfer.files[0]); });
  inp.addEventListener('change', () => { if(inp.files[0]) handleFile(key, inp.files[0]); });
});

// ✅ STABILITY: Limite dimensione file (30MB) per evitare OOM del browser
async function handleFile(key, file) {
  if (file.size > 30 * 1024 * 1024) {
    showError('File troppo grande (>30MB). Il browser potrebbe crashare.');
    return;
  }
  showToast('info', `Caricamento ${file.name}...`);
  try {
    const parsed = await parseFile(file, key);
    S.set(key, parsed);
    document.getElementById(`dz-${key}`)?.classList.add('loaded');
    const rows = parsed.data ? parsed.data.length : (parsed.rawRows ? parsed.rawRows.length : 0);
    const infoEl = document.getElementById(`fi-${key}`);
    if (infoEl) infoEl.textContent = `${file.name} — ${rows} righe`;
    showPreview(key, parsed.rawRows, parsed.headers);
    updatePills();
    if (canCompute()) {
      const btnCalc = document.getElementById('btnCalc');
      const btnExp = document.getElementById('btnFinalExport');
      if(btnCalc) btnCalc.disabled = false;
      if(btnExp) btnExp.disabled = false;
      showToast('ok', 'Tutti i file necessari caricati. Premi "Calcola Bridge".');
    } else {
      showToast('ok', `${file.name}: ${rows} righe`);
    }
  } catch (err) {
    logMsg('Parse error dettaglio: ' + err.message);
    showError('Errore nel parsing del file. Controlla formato e codifica.');
  }
}

function showPreview(key, rawRows, headers) {
  const previewArea = document.getElementById(`preview-${key}`);
  if (!previewArea) return;
  previewArea.style.display = 'block';
  const headerElement = document.getElementById(`preview-${key}-header`);
  const bodyElement = document.getElementById(`preview-${key}-body`);
  if (headerElement) headerElement.innerHTML = '';
  if (bodyElement) bodyElement.innerHTML = '';
  
  if (headers && headers.length > 0 && headerElement) {
    const headerRow = document.createElement('tr');
    headers.forEach(h => { const th = document.createElement('th'); th.textContent = h; headerRow.appendChild(th); });
    headerElement.appendChild(headerRow);
  }
  if (rawRows && rawRows.length > 0 && bodyElement) {
    let headerIndex = 0;
    const sigs = COLUMNS[key] || [];
    for (let i = 0; i < Math.min(12, rawRows.length); i++) {
      const j = rawRows[i].map(c => String(c).toLowerCase()).join('|');
      if (sigs.filter(s => j.includes(s)).length >= 1) { headerIndex = i; break; }
    }
    for (let i = headerIndex + 1; i < Math.min(headerIndex + 6, rawRows.length); i++) {
      if (!isDataRow(rawRows[i]) || rawRows[i].every(c => String(c).trim() === '')) continue;
      const tr = document.createElement('tr');
      rawRows[i].forEach(cell => { const td = document.createElement('td'); td.textContent = cell; tr.appendChild(td); });
      bodyElement.appendChild(tr);
    }
  }
}

function canCompute() { return !!(S.get('budget') && S.get('sil') && S.get('p6')); }
function updatePills() {
  const specs = [
    { key: 'budget', pillId: 'pBudget', valId: 'pBudgetV', getCount: d => d.data?.length ?? 0, optional: false },
    { key: 'sil',    pillId: 'pSil',    valId: 'pSilV',    getCount: d => d.data?.length ?? 0, optional: false },
    { key: 'silI',   pillId: 'pSilI',   valId: 'pSilIV',   getCount: d => d.data?.length ?? 0, optional: true  },
    { key: 'p6',     pillId: 'pP6',     valId: 'pP6V',     getCount: d => d.rawRows?.length ?? 0, optional: false },
  ];
  let loaded = 0;
  specs.forEach(({ key, pillId, valId, getCount, optional }) => {
    const d = S.get(key);
    const pill = document.getElementById(pillId);
    const valEl = document.getElementById(valId);
    if (d) {
      loaded++;
      if (pill) pill.className = 'pill ok';
      if (valEl) valEl.textContent = getCount(d) + ' righe';
    } else {
      if (pill) pill.className = optional ? 'pill warn' : 'pill info';
      if (valEl) valEl.textContent = '—';
    }
  });
  const docsEl = document.getElementById('pDocsV');
  const docsPill = document.getElementById('pDocs');
  if (docsEl) docsEl.textContent = `${loaded}/4`;
  if (docsPill) docsPill.className = loaded >= 3 ? 'pill ok' : loaded > 0 ? 'pill warn' : 'pill';
}

// ── FIELD HELPERS ─────────────────────────────────────────────────────────────
function getF(row, ...cands) {
  for (const c of cands) { if (row[c] !== undefined && String(row[c]).trim() !== '') return String(row[c]).trim(); }
  const keys = Object.keys(row);
  for (const c of cands) {
    const f = keys.find(k => k.toLowerCase().replace(/\s+/g, '').includes(c.toLowerCase().replace(/\s+/g, '')));
    if (f && String(row[f]).trim() !== '') return String(row[f]).trim();
  }
  return '';
}

// ✅ CRITICAL FIX: Regex corretta per rimuovere migliaia (.) e convertire decimali (,)
function getN(row, ...cands) {
  const val = getF(row, ...cands);
  if (!val) return 0;
  const cleaned = val.toString().replace(/[^\d.,-]/g, '').replace(/\./g, '').replace(',', '.');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

// ── PARSERS ───────────────────────────────────────────────────────────────────
function parseBudget(data) {
  const byArt = new Map(), byWbs = new Map();
  data.forEach(r => {
    const wbs = normWBS(getF(r, 'Cod. WBS', 'Cod.WBS', 'CodWBS', 'WBS', 'Codice WBS'));
    const art = normArt(getF(r, 'Articolo', 'articolo', 'Descrizione Articolo'));
    const desWbs = getF(r, 'Des. WBS', 'Des.WBS', 'Descrizione WBS', 'Descrizione');
    const importo = getN(r, 'Importo Costo (€)', 'Importo Costo', 'Importo Costo (euro)', 'Costo', 'Importo', 'Budget');
    if (!wbs && !art) return;
    if (art) { if (!byArt.has(art)) byArt.set(art, []); byArt.get(art).push({ wbs, desWbs, importo }); }
    if (wbs) { if (!byWbs.has(wbs)) byWbs.set(wbs, { desWbs, total: 0 }); byWbs.get(wbs).total += importo; }
  });
  return { byArt, byWbs };
}

function parseSIL(data) {
  const items = []; let latestDate = '';
  data.forEach(r => {
    const silNum = getF(r, 'Cod. S.I.L.', 'Cod.S.I.L.', 'CodSIL', 'SIL', 'Codice SIL');
    const art = normArt(getF(r, 'Articolo', 'articolo', 'Descrizione Articolo'));
    const importo = getN(r, 'Importo', 'importo', 'Costo', 'Valore');
    const dataRaw = getF(r, 'Data', 'data', 'Data SIL');
    if (!silNum && !art) return;
    const dateFmt = excelDateToStr(dataRaw);
    if (dateFmt > latestDate) latestDate = dateFmt;
    items.push({ silNum, art, importo, dataSil: dateFmt });
  });
  return { items, latestDate };
}

function parseP6fromRawRows(rows) {
  const byWbs = new Map(), byAct = new Map();
  let hIdx = -1;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const j = rows[i].map(c => String(c).toLowerCase()).join('|');
    if (j.includes('task_code') && j.includes('wbs_id')) { hIdx = i; break; }
  }
  if (hIdx === -1) {
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const j = rows[i].map(c => String(c).toLowerCase()).join('|');
      if (j.includes('activity id') && j.includes('wbs code')) {
        hIdx = i;
        rows[i] = rows[i].map(c => {
          const lc = String(c).toLowerCase().trim();
          if (lc.includes('activity id')) return 'task_code';
          if (lc.includes('wbs code')) return 'wbs_id';
          if (lc.includes('actual total cost')) return 'act_cost';
          if (lc.includes('physical % complete')) return 'calc_phys_complete_pct';
          if (lc.includes('activity status')) return 'status_code';
          if (lc.includes('activity name')) return 'act_name';
          return c;
        });
        break;
      }
    }
  }
  if (hIdx === -1) return { byWbs, byAct };
  const headers = rows[hIdx].map(h => String(h).trim());
  const nextRow = rows[hIdx + 1] || [];
  const nextJ = nextRow.map(c => String(c).toLowerCase()).join('|');
  const startData = (nextJ.includes('activity id') || nextJ.includes('activity status')) ? hIdx + 2 : hIdx + 1;

  const iCode = headers.findIndex(h => /task_code|activity id/i.test(h));
  const iWbs = headers.findIndex(h => /wbs_id|wbs code/i.test(h));
  const iCost = headers.findIndex(h => /act_cost|actual total cost/i.test(h));
  const iPhys = headers.findIndex(h => /phys|% complete/i.test(h));
  const iStat = headers.findIndex(h => /status/i.test(h));
  const iName = headers.findIndex(h => /name|desc/i.test(h));

  let count = 0;
  for (let i = startData; i < rows.length; i++) {
    const row = rows[i];
    if (!isDataRow(row) || row.every(c => String(c).trim() === '')) continue;
    const actId = String(row[iCode] ?? '').trim();
    if (!actId) continue;
    const wbs = normWBS(String(row[iWbs] ?? '').trim());
    const cost = parseFloat(String(row[iCost] ?? '0').replace(/[^\d.-]/g, '')) || 0;
    const phys = parseFloat(String(row[iPhys] ?? '0').replace(/[^\d.]/g, '')) || 0;
    const status = String(row[iStat] ?? '').trim();
    const name = String(row[iName] ?? '').trim();

    byAct.set(actId, { actId, wbs, cost, phys, status, name });
    if (!byWbs.has(wbs)) byWbs.set(wbs, []);
    byWbs.get(wbs).push({ actId, cost, phys, status, name });
    count++;
  }
  logMsg(`P6 attività caricate: ${count}, WBS uniche: ${byWbs.size}`);
  return { byWbs, byAct };
}

// ── BRIDGE PRINCIPALE (Async/Chunked) ─────────────────────────────────────────
async function runBridge() {
  clearLog();
  logMsg('=== Bridge CPM→P6 v16 (Secure) ===');
  if (!S.get('budget') || !S.get('sil') || !S.get('p6')) {
    showToast('err', 'Carica Budget CPM, SIL Diretti e Export P6');
    return;
  }
  setProgress(10, 'Parsing budget...');
  const budget = parseBudget(S.get('budget').data);
  await yieldToMain();
  
  setProgress(30, 'Parsing SIL...');
  const silDir = parseSIL(S.get('sil').data);
  const silInd = (S.get('silI') && S.get('silI').data) ? parseSIL(S.get('silI').data) : { items: [], latestDate: '' };
  await yieldToMain();

  setProgress(50, 'Parsing P6...');
  const p6Data = S.get('p6');
  const p6 = p6Data ? parseP6fromRawRows(p6Data.rawRows) : { byWbs: new Map(), byAct: new Map() };
  await yieldToMain();

  logMsg(`Budget: ${budget.byWbs.size} WBS, ${budget.byArt.size} articoli`);
  logMsg(`SIL Diretti: ${silDir.items.length} | Indiretti: ${silInd.items.length}`);
  logMsg(`P6: ${p6.byAct.size} attività, ${p6.byWbs.size} WBS`);

  const latestSil = silDir.latestDate > silInd.latestDate ? silDir.latestDate : silInd.latestDate;
  if (latestSil) { const mSil = document.getElementById('mSil'); if(mSil) mSil.textContent = latestSil; }

  setProgress(60, 'Aggregazione SIL...');
  const allSil = [...silDir.items, ...silInd.items];
  const silByArt = new Map();
  allSil.forEach(item => {
    if (!item.art) return;
    if (!silByArt.has(item.art)) silByArt.set(item.art, { total: 0, rows: [] });
    silByArt.get(item.art).total += item.importo;
    silByArt.get(item.art).rows.push(item);
  });
  await yieldToMain();

  setProgress(75, 'Mapping Articolo → WBS...');
  const silByWbs = new Map();
  const unmappedArts = [];
  for (const [art, silGrp] of silByArt.entries()) {
    let budgetEntries = budget.byArt.get(art);
    if (!budgetEntries?.length) {
      const prefix = art.substring(0, 8);
      for (const [bArt, entries] of budget.byArt.entries()) {
        if (bArt.startsWith(prefix)) { budgetEntries = entries; break; }
      }
    }
    if (!budgetEntries?.length) {
      const lowerArt = art.toLowerCase();
      for (const [bArt, entries] of budget.byArt.entries()) {
        if (bArt.toLowerCase().includes(lowerArt.substring(0, 5))) { budgetEntries = entries; break; }
      }
    }
    if (!budgetEntries?.length) {
      unmappedArts.push({ art, importo: silGrp.total, reason: 'Non in Budget CPM' });
      continue;
    }
    const wbsMap = new Map();
    budgetEntries.forEach(e => {
      if (!e.wbs) return;
      if (!wbsMap.has(e.wbs)) wbsMap.set(e.wbs, { desWbs: e.desWbs, importo: 0 });
      wbsMap.get(e.wbs).importo += e.importo;
    });
    const totalBud = [...wbsMap.values()].reduce((s, v) => s + v.importo, 0);
    for (const [wbs, bEntry] of wbsMap.entries()) {
      const peso = totalBud > 0 ? bEntry.importo / totalBud : 1 / wbsMap.size;
      const importoAlloc = silGrp.total * peso;
      if (!silByWbs.has(wbs)) silByWbs.set(wbs, { total: 0, desWbs: bEntry.desWbs, arts: [] });
      silByWbs.get(wbs).total += importoAlloc;
      silByWbs.get(wbs).arts.push({ art, importo: importoAlloc, peso });
    }
  }
  await yieldToMain();

  setProgress(85, 'Distribuzione WBS → Activity P6...');
  const distrib = [];
  const unmappedWbs = [];
  for (const [wbs, silWbs] of silByWbs.entries()) {
    const acts = p6.byWbs.get(wbs) || [];
    if (!acts.length) {
      if (silWbs.total > 0) unmappedWbs.push({ wbs, importo: silWbs.total, desWbs: silWbs.desWbs });
      continue;
    }
    const totalCost = acts.reduce((s, a) => s + a.cost, 0);
    const totalPhys = acts.reduce((s, a) => s + a.phys, 0);
    acts.forEach(act => {
      let peso, method;
      if (totalCost > 0) { peso = act.cost / totalCost; method = 'COST'; }
      else if (totalPhys > 0) { peso = act.phys / totalPhys; method = 'PHY'; }
      else { peso = 1 / acts.length; method = 'EQ'; }
      distrib.push({
        actId: act.actId, wbs, desWbs: silWbs.desWbs, method,
        silImporto: silWbs.total * peso, p6Cost: act.cost, physPct: act.phys,
        status: act.status, actName: act.name, delta: (silWbs.total * peso) - act.cost
      });
    });
  }
  await yieldToMain();

  setProgress(95, 'Calcolo KPI e Alert...');
  const totalSil = [...silByWbs.values()].reduce((s, v) => s + v.total, 0);
  const totalSilRaw = allSil.reduce((s, i) => s + i.importo, 0);
  const totalP6 = distrib.reduce((s, d) => s + d.p6Cost, 0);
  const totalBudget = [...budget.byWbs.values()].reduce((s, v) => s + v.total, 0);
  const cpi = totalP6 > 0 ? totalSil / totalP6 : null;

  const summaryByWbs = new Map();
  distrib.forEach(d => {
    if (!summaryByWbs.has(d.wbs)) summaryByWbs.set(d.wbs, { wbs: d.wbs, desWbs: d.desWbs, silTot: 0, p6Tot: 0, acts: 0 });
    const s = summaryByWbs.get(d.wbs);
    s.silTot += d.silImporto; s.p6Tot += d.p6Cost; s.acts++;
  });

  const alerts = [];
  distrib.forEach(d => {
    if (d.silImporto > 0 && d.physPct === 0 && d.status && !d.status.toLowerCase().includes('not started'))
      alerts.push({ type: 'warn', actId: d.actId, wbs: d.wbs, msg: `SIL €${fmt(d.silImporto)} ma % fisica = 0`, method: d.method, actName: d.actName });
  });
  unmappedArts.forEach(u => alerts.push({ type: 'err', actId: '—', wbs: '—', msg: `Articolo non in Budget: "${u.art}" — €${fmt(u.importo)}`, method: '—', actName: '' }));
  unmappedWbs.forEach(u => alerts.push({ type: 'err', actId: '—', wbs: u.wbs, msg: `WBS "${u.wbs}" non in Export P6 — €${fmt(u.importo)}`, method: '—', actName: '' }));

  const threshold = S.threshold || 5000;
  const deviazioni = distrib
    .filter(d => Math.abs(d.delta) >= threshold)
    .map(d => ({ ...d, absDelta: Math.abs(d.delta), pct: d.p6Cost > 0 ? Math.abs(d.delta) / d.p6Cost : null }));

  S.set('result', { distrib, summaryByWbs, alerts, deviazioni, unmappedArts, unmappedWbs, totalSil, totalSilRaw, totalP6, totalBudget, cpi, budget, p6, silByWbs, allSil });
  setProgress(100, 'Completato');
  await yieldToMain();
  renderResults();
}

function yieldToMain() { return new Promise(r => setTimeout(r, 0)); }
function setProgress(pct, txt) {
  const bar = document.getElementById('progressBar');
  if (!bar) return;
  if (pct > 0) bar.style.display = 'block';
  const fill = bar.querySelector('.progress-fill');
  const text = bar.querySelector('.progress-text');
  if (fill) fill.style.width = `${pct}%`;
  if (text) text.textContent = txt || `${Math.round(pct)}%`;
  if (pct >= 100) setTimeout(() => bar.style.display = 'none', 1500);
}

// ── RENDER ────────────────────────────────────────────────────────────────────
function renderResults() {
  const R = S.get('result');
  if (!R) return;
  
  const cpiEl = document.getElementById('pCpiV');
  if(cpiEl) cpiEl.textContent = R.cpi ? R.cpi.toFixed(3) : '—';
  const cpiPill = document.getElementById('pCpi');
  if(cpiPill) cpiPill.className = `pill ${R.cpi === null ? 'info' : (R.cpi >= 0.95 && R.cpi <= 1.05 ? 'ok' : (R.cpi > 1.05 ? 'warn' : 'err'))}`;

  document.getElementById('bnrNoBudget')?.classList.toggle('show', !S.get('budget'));
  const totUnm = R.unmappedArts.length + R.unmappedWbs.length;
  document.getElementById('bnrUnmapped')?.classList.toggle('show', totUnm > 0);
  if (totUnm > 0) {
    const lostEur = R.unmappedArts.reduce((s, u) => s + u.importo, 0) + R.unmappedWbs.reduce((s, u) => s + u.importo, 0);
    const txtEl = document.getElementById('bnrUnmTxt');
    if(txtEl) txtEl.textContent = `${R.unmappedArts.length} art. non in Budget + ${R.unmappedWbs.length} WBS non in P6 — €${fmt(lostEur)}`;
  }
  document.getElementById('bnrOk')?.classList.toggle('show', totUnm === 0 && R.distrib.length > 0);
  const bnrError = document.getElementById('bnrError');
  if(bnrError) bnrError.style.display = 'none';
  document.getElementById('btnExpHdr')?.classList.toggle('show', R.distrib.length > 0);

  ['f1','f2','f3','f4','f5','f6'].forEach(id => document.getElementById(id)?.classList.remove('done','warn'));
  if (S.get('budget')) document.getElementById('f1')?.classList.add('done');
  if (S.get('sil')) document.getElementById('f2')?.classList.add('done');
  if (S.get('silI')) document.getElementById('f3')?.classList.add('done'); else document.getElementById('f3')?.classList.add('warn');
  if (S.get('p6')) document.getElementById('f4')?.classList.add('done');
  document.getElementById('f5')?.classList.add('done');
  document.getElementById('f6')?.classList.add(R.alerts.length > 0 ? 'warn' : 'done');

  const out = document.getElementById('outArea'); 
  if (!out) return;
  out.innerHTML = '';

  // KPI
  const kpiCard = makeCard('📊 KPI Bridge');
  const kg = el('div', 'kgrid');
  kg.appendChild(makeKpi(fmtE(R.totalSil), 'SIL Allocato (€)', 'ok'));
  kg.appendChild(makeKpi(fmtE(R.totalP6), 'P6 Costo (€)', R.totalP6 > 0 ? 'ok' : 'warn'));
  kg.appendChild(makeKpi(fmtE(R.totalBudget), 'Budget CPM (€)', 'ok'));
  kg.appendChild(makeKpi(R.cpi ? R.cpi.toFixed(3) : '—', 'CPI', R.cpi === null ? 'info' : (R.cpi >= 0.95 && R.cpi <= 1.05 ? 'ok' : 'err')));
  kg.appendChild(makeKpi(R.distrib.length, 'Attività P6', R.distrib.length > 0 ? 'ok' : 'warn'));
  kg.appendChild(makeKpi(R.alerts.length, 'Alert', R.alerts.length === 0 ? 'ok' : 'err'));
  kg.appendChild(makeKpi(R.deviazioni.length, 'Deviazioni', R.deviazioni.length === 0 ? 'ok' : 'warn'));
  const mappedPct = R.totalSilRaw > 0 ? Math.round(R.totalSil / R.totalSilRaw * 100) : 0;
  kg.appendChild(makeKpi(mappedPct + '%', 'SIL Mappato', mappedPct > 95 ? 'ok' : 'warn'));
  kpiCard.querySelector('.card-body').appendChild(kg);
  out.appendChild(kpiCard);

  // Tabs
  const tabCard = makeCard('');
  if(tabCard.querySelector('.card-head')) tabCard.querySelector('.card-head').style.display = 'none';
  const tabs = buildTabs([
    { label: 'Riepilogo WBS', id: 'riepilogo' },
    { label: 'Distribuzione P6 ', id: 'distrib', count: R.distrib.length, tnClass: 'tn-bri' },
    { label: 'Alert ', id: 'alert', count: R.alerts.length, tnClass: 'tn-alert' },
    { label: 'Deviazioni ', id: 'deviazioni', count: R.deviazioni.length, tnClass: 'tn-dev' },
    { label: '📄 Log', id: 'log' }
  ]);
  tabCard.appendChild(tabs);
  const tcWrap = el('div', '');

  // RIEPILOGO
  const tcR = el('div', 'tc card-body active'); tcR.id = 'tc-riepilogo';
  const wbsKeys = [...R.summaryByWbs.keys()].sort();
  if (!wbsKeys.length) { tcR.innerHTML = '<div class="empty"><div class="e-ico">📬</div><div class="e-title">Nessun dato</div></div>'; }
  else {
    const scr = el('div', 'card-scroll');
    const t = buildTable(['WBS', 'Des. WBS', 'SIL (€)', 'P6 Costo (€)', 'Delta (€)', 'N. Att.', 'Budget (€)', 'Status']);
    const tb = t.querySelector('tbody');
    let tS = 0, tP = 0;
    wbsKeys.forEach(wbs => {
      const s = R.summaryByWbs.get(wbs);
      const bud = R.budget.byWbs.get(wbs);
      const delta = s.silTot - s.p6Tot;
      tS += s.silTot; tP += s.p6Tot;
      const statusTag = delta === 0 ? {__html:'<span class="tag tag-ok">OK</span>'} : delta > 0 ? {__html:'<span class="tag tag-warn">SIL &gt;P6</span>'} : {__html:'<span class="tag tag-err">P6 &gt;SIL</span>'};
      tb.appendChild(buildRow([wbs, s.desWbs || '—', fmtE(s.silTot), fmtE(s.p6Tot), fmtDelta(delta), s.acts, bud ? fmtE(bud.total) : '—', statusTag], delta > 0 ? 'r-over' : delta < 0 ? 'r-miss' : 'r-ok'));
    });
    R.unmappedWbs.forEach(u => {
      tb.appendChild(buildRow([u.wbs, u.desWbs || '—', fmtE(u.importo), '—', '—', '—', '—', {__html:'<span class="tag tag-err">NO P6</span>'}], 'r-miss'));
    });
    tb.appendChild(buildRow(['TOTALE', '', fmtE(tS), fmtE(tP), fmtDelta(tS - tP), R.distrib.length, '', ''], 'r-total'));
    scr.appendChild(t); tcR.appendChild(scr);
  }
  tcWrap.appendChild(tcR);

  // DISTRIBUZIONE
  const tcD = el('div', 'tc card-body'); tcD.id = 'tc-distrib';
  if (!R.distrib.length) { tcD.innerHTML = '<div class="empty"><div class="e-ico">📬</div><div class="e-title">Nessuna attività</div></div>'; }
  else {
    const scr = el('div', 'card-scroll');
    const t = buildTable(['Activity ID', 'Nome Attività', 'WBS', 'Des. WBS', 'Metodo', 'SIL Alloc. (€)', 'P6 Costo (€)', 'Phys%', 'Status', 'Delta (€)']);
    const tb = t.querySelector('tbody');
    R.distrib.forEach(d => {
      const mtTag = d.method === 'COST' ? {__html:'<span class="tag tag-ok">COST</span>'} : d.method === 'PHY' ? {__html:'<span class="tag tag-phy">PHY</span>'} : {__html:'<span class="tag tag-eq">EQ</span>'};
      tb.appendChild(buildRow([d.actId, d.actName || '—', d.wbs, d.desWbs || '—', mtTag, fmtE(d.silImporto), fmtE(d.p6Cost), (d.physPct || 0).toFixed(1) + '%', d.status || '—', fmtDelta(d.delta)], d.method === 'COST' ? 'r-ok' : 'r-bri'));
    });
    scr.appendChild(t); tcD.appendChild(scr);
  }
  tcWrap.appendChild(tcD);

  // ALERT & DEVIAZIONI (Placeholder rendering sicuro)
  const tcA = el('div', 'tc card-body'); tcA.id = 'tc-alert';
  if (!R.alerts.length) { tcA.innerHTML = '<div class="empty">✅ Nessun alert</div>'; }
  else {
    const t = buildTable(['Tipo', 'Activity ID', 'WBS', 'Messaggio']);
    const tb = t.querySelector('tbody');
    R.alerts.forEach(a => tb.appendChild(buildRow([a.type === 'warn' ? '⚠️' : '🚨', a.actId, a.wbs, a.msg], a.type === 'warn' ? 'r-over' : 'r-miss')));
    tcA.appendChild(el('div','card-scroll')).appendChild(t);
  }
  tcWrap.appendChild(tcA);

  const tcDev = el('div', 'tc card-body'); tcDev.id = 'tc-deviazioni';
  if (!R.deviazioni.length) { tcDev.innerHTML = '<div class="empty">✅ Deviazioni nella soglia</div>'; }
  else {
    const t = buildTable(['WBS', 'Activity ID', 'Delta (€)', 'Abs (€)', '%']);
    const tb = t.querySelector('tbody');
    R.deviazioni.forEach(d => tb.appendChild(buildRow([d.wbs, d.actId, fmtDelta(d.delta), '€'+fmt(d.absDelta), d.pct ? (d.pct*100).toFixed(1)+'%' : '—'])));
    tcDev.appendChild(el('div','card-scroll')).appendChild(t);
  }
  tcWrap.appendChild(tcDev);

  // LOG
  const tcLog = el('div', 'tc card-body'); tcLog.id = 'tc-log';
  const logBox = el('div', 'log-box');
  logBox.textContent = _log.join('\n');
  tcLog.appendChild(logBox);
  tcWrap.appendChild(tcLog);

  out.appendChild(tabCard);
  out.appendChild(tcWrap);

  // Tab switching delegation
  tabCard.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      tabCard.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const target = btn.dataset.tab;
      tcWrap.querySelectorAll('.tc').forEach(tc => {
        tc.classList.toggle('active', tc.id === `tc-${target}`);
      });
    });
  });

  // Attach to DOM safely
  const container = document.getElementById('container');
  if (container) container.appendChild(out);
}

// ── SAFE DOM BUILDERS ─────────────────────────────────────────────────────────
function buildTabs(defs) {
  const wrap = el('div', 'tabs');
  defs.forEach(({ label, id, count, tnClass }, idx) => {
    const btn = el('button', idx === 0 ? 'tab active' : 'tab');
    btn.dataset.tab = id;
    btn.textContent = label;
    if (count !== undefined) {
      const sp = el('span', `tn ${tnClass}`);
      sp.textContent = String(count);
      btn.appendChild(sp);
    }
    wrap.appendChild(btn);
  });
  return wrap;
}
function makeCard(title) {
  const card = el('div', 'card');
  const head = el('div', 'card-head');
  const h3 = document.createElement('h3'); h3.textContent = title; head.appendChild(h3);
  card.appendChild(head);
  card.appendChild(el('div', 'card-body'));
  return card;
}
function makeKpi(val, label, status) {
  const kpi = el('div', `kpi ${status}`);
  const v = el('div', 'kv'); v.textContent = val;
  const l = el('div', 'kl'); l.textContent = label;
  kpi.appendChild(v); kpi.appendChild(l);
  return kpi;
}
function buildTable(headers) {
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tr = document.createElement('tr');
  headers.forEach(h => { const th = document.createElement('th'); th.textContent = h; tr.appendChild(th); });
  thead.appendChild(tr); table.appendChild(thead);
  table.appendChild(document.createElement('tbody'));
  return table;
}
function buildRow(cells, cls) {
  const tr = document.createElement('tr');
  if (cls) tr.className = cls;
  cells.forEach(c => {
    const td = document.createElement('td');
    if (c !== null && typeof c === 'object' && c.__html) {
      td.innerHTML = c.__html;
    } else {
      td.textContent = c ?? '';
    }
    tr.appendChild(td);
  });
  return tr;
}

// ── DOCS MODAL ────────────────────────────────────────────────────────────────
function mdToHtml(md) {
  const lines = md.split('\n');
  let out = '', inTable = false, inList = false, inCode = false, codeLines = [];

  function flush() {
    if (inList)  { out += '</ul>'; inList  = false; }
    if (inTable) { out += '</tbody></table>'; inTable = false; }
  }
  function inl(s) {
    return String(s)
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/\*\*(.+?)\*\*/g,'<strong>$1</strong>')
      .replace(/`([^`]+)`/g,'<code>$1</code>');
  }

  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];
    const line = raw.trim();

    if (/^```/.test(line)) {
      if (!inCode) { flush(); inCode = true; codeLines = []; }
      else {
        const esc = codeLines.map(l => l.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')).join('\n');
        out += `<pre><code>${esc}</code></pre>`;
        inCode = false;
      }
      continue;
    }
    if (inCode) { codeLines.push(raw); continue; }

    if (/^-{3,}$/.test(line))  { flush(); out += '<hr>'; continue; }
    if (/^# /.test(line))      { flush(); out += `<h1>${inl(line.slice(2))}</h1>`; continue; }
    if (/^## /.test(line))     { flush(); out += `<h2>${inl(line.slice(3))}</h2>`; continue; }
    if (/^### /.test(line))    { flush(); out += `<h3>${inl(line.slice(4))}</h3>`; continue; }
    if (/^#### /.test(line))   { flush(); out += `<h4>${inl(line.slice(5))}</h4>`; continue; }

    if (/^\|/.test(line)) {
      if (/^\|[-| :]+\|$/.test(line)) continue;
      const cells = line.split('|').filter((_,j,a) => j > 0 && j < a.length - 1).map(c => c.trim());
      if (!inTable) {
        if (inList) { out += '</ul>'; inList = false; }
        out += '<table><thead><tr>' + cells.map(c => `<th>${inl(c)}</th>`).join('') + '</tr></thead><tbody>';
        inTable = true;
        if (i + 1 < lines.length && /^\|[-| :]+\|$/.test(lines[i+1].trim())) i++;
      } else {
        out += '<tr>' + cells.map(c => `<td>${inl(c)}</td>`).join('') + '</tr>';
      }
      continue;
    } else if (inTable) { out += '</tbody></table>'; inTable = false; }

    if (/^- /.test(line)) {
      if (!inTable && !inList) { out += '<ul>'; inList = true; }
      out += `<li>${inl(line.slice(2))}</li>`;
      continue;
    } else if (inList) { out += '</ul>'; inList = false; }

    if (line === '') { continue; }

    out += `<p>${inl(line)}</p>`;
  }
  flush();
  return out;
}

function openDoc(key) {
  const rawEl = document.getElementById(key === 'manuale' ? 'raw-manuale' : 'raw-config');
  const modal  = document.getElementById('docModal');
  const titleEl = document.getElementById('docModalTitle');
  const bodyEl  = document.getElementById('docModalBody');
  if (!rawEl || !modal || !bodyEl) return;
  titleEl.textContent = key === 'manuale' ? '📖 Manuale Utente' : '⚙️ Configurazione Tecnica';
  bodyEl.innerHTML = mdToHtml(rawEl.textContent);
  modal.style.display = 'flex';
  document.body.style.overflow = 'hidden';
  modal.focus();
}

function closeDoc() {
  const modal = document.getElementById('docModal');
  if (modal) { modal.style.display = 'none'; document.body.style.overflow = ''; }
}

// ── INIT ──────────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  const btnCalc = document.getElementById('btnCalc');
  const btnReset = document.getElementById('btnReset');
  if(btnCalc) btnCalc.addEventListener('click', runBridge);
  if(btnReset) btnReset.addEventListener('click', () => {
    S.reset();
    S.cleanup();
    document.querySelectorAll('.dz').forEach(d => d.classList.remove('loaded'));
    document.querySelectorAll('input[type=file]').forEach(i => i.value = '');
    document.querySelectorAll('.dz-info').forEach(i => i.textContent = '');
    document.querySelectorAll('.preview-area').forEach(p => p.style.display = 'none');
    const outArea = document.getElementById('outArea');
    if (outArea) outArea.innerHTML = '<div class="empty"><div class="e-ico">🔗</div><div class="e-title">Stato resettato</div><div class="e-desc">Carica i file per avviare una nuova elaborazione.</div></div>';
    document.getElementById('btnExpHdr')?.classList.remove('show');
    const btnCalc = document.getElementById('btnCalc');
    if (btnCalc) btnCalc.disabled = true;
    ['bnrUnmapped','bnrOk'].forEach(id => document.getElementById(id)?.classList.remove('show'));
    document.getElementById('bnrNoBudget')?.classList.add('show');
    ['f1','f2','f3','f4','f5','f6'].forEach(id => document.getElementById(id)?.classList.remove('done','warn','out'));
    const mComm = document.getElementById('mComm'); if(mComm) mComm.textContent = '—';
    const mSil = document.getElementById('mSil'); if(mSil) mSil.textContent = '—';
    clearLog();
    updatePills();
    showToast('info', 'Stato resettato');
  });
  document.getElementById('btnDocManuale')?.addEventListener('click', () => openDoc('manuale'));
  document.getElementById('btnDocConfig')?.addEventListener('click',  () => openDoc('config'));
  document.getElementById('docModalClose')?.addEventListener('click', closeDoc);
  document.getElementById('docModal')?.addEventListener('click', e => { if (e.target.id === 'docModal') closeDoc(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && document.getElementById('docModal')?.style.display !== 'none') closeDoc(); });

  logMsg('App inizializzata. Modalità client-side sicura attiva.');
});

})();