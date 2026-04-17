/* 
 * Oggetto principale dello stato dell'applicazione
 * Contiene i dati caricati dai file e i risultati del bridge
 */
const S = { 
  budget: null,  // Dati del budget CPM
  sil: null,     // Dati dei SIL diretti
  silI: null,    // Dati dei SIL indiretti (opzionali)
  p6: null,      // Dati dell'export P6
  result: null   // Risultati del bridge
};

/* 
 * Array per memorizzare i messaggi di log
 * Usato per tenere traccia delle operazioni eseguite
 */
let _log = [];

/* 
 * Oggetto che definisce i nomi possibili per ogni foglio Excel
 * Aiuta a trovare il foglio giusto anche se ha nomi diversi
 */
const SHEET_HINTS = {
  budget: ['BUDGET','Budget','budget', 'BUDGET CPM'],
  sil:    ['SIL diretti','SIL Diretti','SIL_DIRETTI', 'SIL DIRETTI'],
  silI:   ['SIL indiretti','SIL Indiretti','SIL_INDIRETTI', 'SIL INDIRETTI'],
  p6:     ['EXPORT_P6','Export_P6','export_p6','EXPORT P6', 'EXPORT_P6.XLSX']
};

// ── FUNZIONI DI UTILITÀ ──────────────────────────────────────────────────────────────────

/**
 * Aggiunge un messaggio al log
 * @param {string} m - Messaggio da aggiungere al log
 */
function logMsg(m){ 
  _log.push(m); 
}

/**
 * Azzera il log
 */
function clearLog(){ 
  _log = []; 
}

/**
 * Formatta un numero come stringa localizzata
 * @param {number} n - Numero da formattare
 * @returns {string} - Numero formattato
 */
function fmt(n){ 
  return Math.round(n).toLocaleString('it-IT'); 
}

/**
 * Formatta un numero come valore monetario
 * @param {number} n - Numero da formattare
 * @returns {string} - Numero formattato come valore monetario
 */
function fmtE(n){ 
  return (+n).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2}); 
}

/**
 * Formatta una differenza numerica con indicatori visivi
 * @param {number} n - Valore della differenza
 * @returns {string} - HTML rappresentante la differenza
 */
function fmtDelta(n){
  const s = Math.abs(n).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2});
  if(n > 0.01) return '<span class="ora-t">+' + s + '</span>';
  if(n < -0.01) return '<span class="err-t">(' + s + ')</span>';
  return '<span class="ok-t">&mdash;</span>';
}

/**
 * Crea un elemento DOM con classe opzionale
 * @param {string} tag - Nome del tag HTML
 * @param {string} cls - Classe CSS da applicare
 * @returns {HTMLElement} - Elemento creato
 */
function el(tag, cls){ 
  const e = document.createElement(tag); 
  if(cls) e.className = cls; 
  return e; 
}

/**
 * Normalizza un codice WBS rimuovendo zeri iniziali
 * @param {string|number} w - Codice WBS da normalizzare
 * @returns {string} - Codice WBS normalizzato
 */
function normWBS(w){
  if(!w) return '';
  const s = String(w).trim();
  
  // Gestione casi speciali
  if(s === '' || isNaN(parseFloat(s.replace(/[^0-9.]/g, '')))) return '';
  
  // Rimuove zeri iniziali da parti numeriche separate dal punto
  const parts = s.split('.');
  const normalizedParts = [];
  
  for(let i = 0; i < parts.length; i++) {
    let part = parts[i].replace(/^0+(\d)/, '$1'); // Rimuove zeri iniziali eccetto per lo 0 puro
    if(part === '') part = '0'; // Se dopo aver rimosso zeri rimane vuoto, mette 0
    normalizedParts.push(part);
  }
  
  return normalizedParts.join('.');
}

/**
 * Normalizza un articolo (case insensitive)
 * @param {string} a - Articolo da normalizzare
 * @returns {string} - Articolo normalizzato
 */
function normArt(a){ 
  return String(a).trim().toLowerCase().replace(/\s+/g,' '); 
}

/**
 * Controlla se una riga contiene dati significativi
 * @param {Array} row - Rigua da controllare
 * @returns {boolean} - True se la riga contiene dati
 */
function isDataRow(row){
  if(!row || !row.length) return false;
  const f = String(row[0]).trim();
  if(!f) return false;
  if(/^[\u{1F300}-\u{1FFFF}]/u.test(f)) return false; // Emojis
  if(/^istruzione/i.test(f)) return false;
  if(/^\(\*\)/i.test(f)) return false;
  if(/^comment/i.test(f)) return false;
  return true;
}

/**
 * Converte una data Excel in formato stringa
 * @param {string|number} raw - Valore grezzo della data
 * @returns {string} - Data in formato stringa
 */
function excelDateToStr(raw){
  const n = parseFloat(raw);
  if(!isNaN(n) && n > 40000){
    const d = new Date(Math.round((n-25569)*86400*1000));
    return d.toISOString().substring(0,10);
  }
  return raw || '';
}

// ── PARSER DEI FILE ───────────────────────────────────────────────────────────────

/**
 * Analizza un file e ne estrae i dati
 * @param {File} file - File da analizzare
 * @param {string} key - Chiave che identifica il tipo di file
 * @returns {Promise<Object>} - Promessa che risolve con i dati estratti
 */
function parseFile(file, key){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try{
        let rows;
        const isXlsx = /\.(xlsx|xls)$/i.test(file.name);
        if(isXlsx){
          // Legge file Excel usando la libreria XLSX
          const wb = XLSX.read(e.target.result, {type: 'array'});
          const hints = SHEET_HINTS[key] || [];
          let sheetName = wb.SheetNames[0];
          
          // Cerca il foglio appropriato usando le indicazioni
          for(const h of hints){
            const found = wb.SheetNames.find(n => n.toLowerCase().includes(h.toLowerCase()));
            if(found){
              sheetName = found;
              break;
            }
          }
          
          logMsg('Foglio selezionato per ' + key + ': "' + sheetName + '"');
          const ws = wb.Sheets[sheetName];
          rows = XLSX.utils.sheet_to_json(ws, {header: 1, defval: ''});
        } else {
          // Legge file CSV o TXT
          const text = new TextDecoder('utf-8').decode(e.target.result);
          const sep = text.split('\n')[0].includes('\t') ? '\t' : ',';
          rows = text.split(/\r?\n/).map(l => l.split(sep).map(c => c.replace(/^"|"$/g,'').trim()));
        }

        // Per P6 usiamo parser dedicato
        if(key === 'p6'){
          resolve({data: null, rawRows: rows});
          return;
        }

        // Trova header per Budget e SIL
        const signals = {
          budget: ['cod. wbs','articolo','importo costo', 'costo', 'budget', 'wbs', 'codice wbs'],
          sil: ['cod. s.i.l.','articolo','importo', 'sil', 'costo', 'codice sil'],
          silI: ['cod. s.i.l.','articolo','importo', 'sil', 'costo', 'codice sil']
        };
        const sigs = signals[key] || [];
        let hIdx = -1;
        for(let i = 0; i < Math.min(12, rows.length); i++){
          const j = rows[i].map(c => String(c).toLowerCase()).join('|');
          // Controllo più flessibile: almeno 1 termine deve corrispondere
          if(sigs.filter(s => j.includes(s)).length >= 1){hIdx = i; break;}
        }
        if(hIdx === -1) hIdx = 0;

        const headers = rows[hIdx].map(h => String(h).trim());
        const data = [];
        for(let i = hIdx + 1; i < rows.length; i++){
          if(!isDataRow(rows[i])) continue;
          if(rows[i].every(c => String(c).trim() === '')) continue;
          const obj = {};
          headers.forEach((h, j) => {obj[h] = String(rows[i][j] ?? '').trim();});
          data.push(obj);
        }
        resolve({headers, data});
      } catch(err){
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ── GESTIONE DELLE AREE DI DROP ────────────────────────────────────────────────────────

// Aggiunge eventi per tutte le aree di drop
document.querySelectorAll('.dz').forEach(dz => {
  const key = dz.dataset.key;
  const inp = dz.querySelector('input[type=file]');
  
  // Clic sull'area apre il dialogo per selezionare file
  dz.addEventListener('click', () => inp.click());
  
  // Eventi per il drag and drop
  dz.addEventListener('dragover', e => {
    e.preventDefault();
    dz.classList.add('over');
  });
  dz.addEventListener('dragleave', () => dz.classList.remove('over'));
  dz.addEventListener('drop', e => {
    e.preventDefault();
    dz.classList.remove('over');
    if(e.dataTransfer.files[0]) handleFile(key, e.dataTransfer.files[0]);
  });
  
  // Quando viene selezionato un file
  inp.addEventListener('change', () => {
    if(inp.files[0]) handleFile(key, inp.files[0]);
  });
});

/**
 * Gestisce un file caricato
 * @param {string} key - Chiave del tipo di file
 * @param {File} file - File caricato
 */
async function handleFile(key, file){
  showToast('info', 'Caricamento ' + file.name + '...');
  try{
    const parsed = await parseFile(file, key);
    S[key] = parsed;
    document.getElementById('dz-' + key).classList.add('loaded');
    const rows = parsed.data ? parsed.data.length : (parsed.rawRows ? parsed.rawRows.length : 0);
    document.getElementById('fi-' + key).textContent = file.name + ' \u2014 ' + rows + ' righe';
    updatePills();
    if(canCompute()){
      document.getElementById('btnCalc').disabled = false;
      // Mostra messaggio informativo quando tutti i file sono caricati
      showToast('ok', 'Tutti i file necessari sono stati caricati. Premi "Calcola Bridge" per elaborare i dati.');
    } else {
      showToast('ok', file.name + ': ' + rows + ' righe');
    }
  } catch(err){
    showToast('err', 'Errore: ' + err.message);
    console.error(err);
  }
}

/**
 * Controlla se tutti i file necessari sono stati caricati
 * @returns {boolean} - True se tutti i file necessari sono caricati
 */
function canCompute(){ 
  return !!(S.budget && S.sil && S.p6); 
}

// ── FUNZIONI DI SUPPORTO PER L'ACCESSO AI DATI ─────────────────────────────────────────────────────

/**
 * Ottiene un valore da una riga cercando per nomi di colonna
 * @param {Object} row - Riga da cui ottenere il valore
 * @param {...string} cands - Nomi possibili della colonna
 * @returns {string} - Valore trovato o stringa vuota
 */
function getF(row, ...cands){
  for(const c of cands){
    if(row[c] !== undefined && String(row[c]).trim() !== '') return String(row[c]).trim();
  }
  // Cerca case-insensitive
  const keys = Object.keys(row);
  for(const c of cands){
    const f = keys.find(k => k.toLowerCase().replace(/\s+/g, '').includes(c.toLowerCase().replace(/\s+/g, '')));
    if(f && String(row[f]).trim() !== '') return String(row[f]).trim();
  }
  return '';
}

/**
 * Ottiene un valore numerico da una riga
 * @param {Object} row - Riga da cui ottenere il valore
 * @param {...string} cands - Nomi possibili della colonna
 * @returns {number} - Valore numerico o 0
 */
function getN(row, ...cands){
  const val = getF(row, ...cands);
  if (!val) return 0;
  // Pulizia e conversione più robusta
  const cleaned = val.toString()
    .replace(/[^\d.,\-]/g, '')  // Rimuove tutto ciò che non è numero, punto, virgola o meno
    .replace(/\./g, '')         // Rimuove punti (migliaia in formato IT)
    .replace(',', '.');         // Sostituisce virgola con punto (decimale in formato IT)
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

// ── PARSER SPECIFICI PER TIPOLOGIA DI DATI ───────────────────────────────────────────────────

/**
 * Parsea i dati del budget
 * @param {Array} data - Dati grezzi del budget
 * @returns {Object} - Dati del budget organizzati
 */
function parseBudget(data){
  const byArt = new Map(), byWbs = new Map();
  data.forEach(r => {
    const wbs = normWBS(getF(r, 'Cod. WBS', 'Cod.WBS', 'CodWBS', 'WBS', 'Codice WBS'));
    const art = normArt(getF(r, 'Articolo', 'articolo', 'Descrizione Articolo'));
    const desWbs = getF(r, 'Des. WBS', 'Des.WBS', 'Descrizione WBS', 'Descrizione');
    const importo = getN(r, 'Importo Costo (€)', 'Importo Costo', 'Importo Costo (euro)', 'Costo', 'Importo', 'Budget');
    
    if(!wbs && !art) return;
    if(art){
      if(!byArt.has(art)) byArt.set(art, []);
      byArt.get(art).push({wbs, desWbs, importo});
    }
    if(wbs){
      if(!byWbs.has(wbs)) byWbs.set(wbs, {desWbs, total: 0});
      byWbs.get(wbs).total += importo;
    }
  });
  return {byArt, byWbs};
}

/**
 * Parsea i dati dei SIL
 * @param {Array} data - Dati grezzi dei SIL
 * @returns {Object} - Dati dei SIL organizzati
 */
function parseSIL(data){
  const items = [];
  let latestDate = '';
  data.forEach(r => {
    const silNum = getF(r, 'Cod. S.I.L.', 'Cod.S.I.L.', 'CodSIL', 'SIL', 'Codice SIL');
    const art = normArt(getF(r, 'Articolo', 'articolo', 'Descrizione Articolo'));
    const importo = getN(r, 'Importo', 'importo', 'Costo', 'Valore');
    const dataRaw = getF(r, 'Data', 'data', 'Data SIL');
    if(!silNum && !art) return;
    const dateFmt = excelDateToStr(dataRaw);
    if(dateFmt > latestDate) latestDate = dateFmt;
    items.push({silNum, art, importo, dataSil: dateFmt});
  });
  return {items, latestDate};
}

/**
 * Parsea i dati P6 dai dati grezzi
 * @param {Array} rows - Righe grezze dei dati P6
 * @returns {Object} - Dati P6 organizzati
 */
function parseP6fromRawRows(rows){
  const byWbs = new Map(), byAct = new Map();

  // Cerca la riga header TECNICA (contiene task_code e wbs_id)
  let hIdx = -1;
  for(let i = 0; i < Math.min(15, rows.length); i++){
    const j = rows[i].map(c => String(c).toLowerCase()).join('|');
    if(j.includes('task_code') && j.includes('wbs_id')){hIdx = i; break;}
  }
  if(hIdx === -1){
    // fallback: cerca riga con "activity id" e "wbs code"
    for(let i = 0; i < Math.min(15, rows.length); i++){
      const j = rows[i].map(c => String(c).toLowerCase()).join('|');
      if(j.includes('activity id') && j.includes('wbs code')){
        hIdx = i;
        // rimappa i nomi colonna ai nomi tecnici
        rows[i] = rows[i].map(c => {
          const lc = String(c).toLowerCase().trim();
          if(lc.includes('activity id')) return 'task_code';
          if(lc.includes('wbs code')) return 'wbs_id';
          if(lc.includes('actual total cost')) return 'act_cost';
          if(lc.includes('physical % complete')) return 'calc_phys_complete_pct';
          if(lc.includes('activity status')) return 'status_code';
          if(lc.includes('activity name')) return 'act_name';
          return c;
        });
        break;
      }
    }
  }
  if(hIdx === -1){
    logMsg('ERRORE: header P6 non trovato, provando altri pattern comuni...');
    // Ultimo tentativo: cerca pattern comuni
    for(let i = 0; i < Math.min(15, rows.length); i++){
      const j = rows[i].map(c => String(c).toLowerCase()).join('|');
      if(j.includes('task') && j.includes('wbs')){
        hIdx = i;
        // Tenta di identificare le colonne
        const headers = rows[i].map(h => String(h).toLowerCase().trim());
        // Qui non possiamo rinominare perché non sappiamo quali sono esattamente
        logMsg(`Header trovato alla riga ${i+1}: ${headers.join(', ')}`);
        break;
      }
    }
  }
  if(hIdx === -1){
    logMsg('ERRORE: header P6 non trovato');
    return {byWbs, byAct};
  }

  const headers = rows[hIdx].map(h => String(h).trim());
  logMsg('P6 header (riga ' + (hIdx + 1) + '): ' + headers.slice(0, 6).join(' | '));

  // Salta la riga successiva se è la riga etichetta leggibile
  const nextRow = rows[hIdx + 1] || [];
  const nextJ = nextRow.map(c => String(c).toLowerCase()).join('|');
  const startData = (nextJ.includes('activity id') || nextJ.includes('activity status')) ? hIdx + 2 : hIdx + 1;
  logMsg('P6 dati da riga ' + (startData + 1) + ', totale righe: ' + rows.length);

  // Indici colonne chiave
  const iCode = headers.findIndex(h => h.toLowerCase().includes('task_code') || h.toLowerCase().includes('activity id'));
  const iWbs = headers.findIndex(h => h.toLowerCase().includes('wbs_id') || h.toLowerCase().includes('wbs code'));
  const iCost = headers.findIndex(h => h.toLowerCase().includes('act_cost') || h.toLowerCase().includes('actual total cost'));
  const iPhys = headers.findIndex(h => h.toLowerCase().includes('phys') || h.toLowerCase().includes('% complete'));
  const iStat = headers.findIndex(h => h.toLowerCase().includes('status'));
  const iName = headers.findIndex(h => h.toLowerCase().includes('name') || h.toLowerCase().includes('desc'));
  logMsg('Indici colonne P6 — task_code:' + iCode + ' wbs_id:' + iWbs + ' act_cost:' + iCost + ' phys:' + iPhys);

  let count = 0;
  for(let i = startData; i < rows.length; i++){
    const row = rows[i];
    if(!isDataRow(row)) continue;
    if(row.every(c => String(c).trim() === '')) continue;

    const actId = String(row[iCode] ?? '').trim();
    if(!actId) continue;
    const wbs = normWBS(String(row[iWbs] ?? '').trim());
    const cost = parseFloat(String(row[iCost] ?? '0').replace(/[^\d.-]/g, '')) || 0;
    const phys = parseFloat(String(row[iPhys] ?? '0').replace(/[^\d.]/g, '')) || 0;
    const status = String(row[iStat] ?? '').trim();
    const name = String(row[iName] ?? '').trim();

    byAct.set(actId, {actId, wbs, cost, phys, status, name});
    if(!byWbs.has(wbs)) byWbs.set(wbs, []);
    byWbs.get(wbs).push({actId, cost, phys, status, name});
    count++;
  }
  logMsg('P6 attività caricate: ' + count + ', WBS uniche: ' + byWbs.size);
  return {byWbs, byAct};
}

// ── FUNZIONE PRINCIPALE DEL BRIDGE ─────────────────────────────────────────────────────────

/**
 * Esegue il processo di bridge tra dati CPM e P6
 */
function runBridge(){
  clearLog();
  logMsg('=== Bridge CPM→P6 v13.1 ===');
  if(!S.budget || !S.sil || !S.p6){
    showToast('err', 'Carica Budget CPM, SIL Diretti e Export P6');
    return;
  }

  const budget = parseBudget(S.budget.data);
  const silDir = parseSIL(S.sil.data);
  const silInd = S.silI && S.silI.data ? parseSIL(S.silI.data) : {items: [], latestDate: ''};
  const p6 = parseP6fromRawRows(S.p6.rawRows);

  logMsg('Budget: ' + budget.byWbs.size + ' WBS, ' + budget.byArt.size + ' articoli unici');
  logMsg('SIL Diretti: ' + silDir.items.length + ' righe');
  logMsg('SIL Indiretti: ' + silInd.items.length + ' righe');
  logMsg('Export P6: ' + p6.byAct.size + ' attività, ' + p6.byWbs.size + ' WBS');

  const latestSil = silDir.latestDate > silInd.latestDate ? silDir.latestDate : silInd.latestDate;
  if(latestSil) document.getElementById('mSil').textContent = latestSil;

  const allSil = [...silDir.items, ...silInd.items];

  // STEP 1: aggrega SIL per Articolo
  const silByArt = new Map();
  allSil.forEach(item => {
    if(!item.art) return;
    if(!silByArt.has(item.art)) silByArt.set(item.art, {total: 0, rows: []});
    silByArt.get(item.art).total += item.importo;
    silByArt.get(item.art).rows.push(item);
  });

  // STEP 2: Articolo → WBS tramite Budget
  const silByWbs = new Map();
  const unmappedArts = [];

  silByArt.forEach((silGrp, art) => {
    let budgetEntries = budget.byArt.get(art);

    // Fallback: match parziale sui primi 8 caratteri
    if(!budgetEntries || !budgetEntries.length){
      const prefix = art.substring(0, 8);
      budget.byArt.forEach((entries, bArt) => {
        if(!budgetEntries && bArt.startsWith(prefix)) budgetEntries = entries;
      });
    }

    // Fallback: match parziale più flessibile
    if(!budgetEntries || !budgetEntries.length){
      const lowerArt = art.toLowerCase();
      budget.byArt.forEach((entries, bArt) => {
        if(!budgetEntries && bArt.toLowerCase().includes(lowerArt.substring(0, 5))) {
          budgetEntries = entries;
        }
      });
    }

    if(!budgetEntries || !budgetEntries.length){
      unmappedArts.push({art, importo: silGrp.total, reason: 'Non in Budget CPM'});
      logMsg('⚠ Articolo non trovato in Budget: "' + art + '" (€' + fmt(silGrp.total) + ')');
      return;
    }

    const wbsMap = new Map();
    budgetEntries.forEach(e => {
      if(!e.wbs) return;
      if(!wbsMap.has(e.wbs)) wbsMap.set(e.wbs, {desWbs: e.desWbs, importo: 0});
      wbsMap.get(e.wbs).importo += e.importo;
    });

    const totalBud = [...wbsMap.values()].reduce((s, v) => s + v.importo, 0);
    wbsMap.forEach((bEntry, wbs) => {
      const peso = totalBud > 0 ? bEntry.importo/totalBud : 1/wbsMap.size;
      const importoAlloc = silGrp.total * peso;
      if(!silByWbs.has(wbs)) silByWbs.set(wbs, {total: 0, desWbs: bEntry.desWbs, arts: []});
      silByWbs.get(wbs).total += importoAlloc;
      silByWbs.get(wbs).arts.push({art, importo: importoAlloc, peso});
    });
  });

  logMsg('Mapping Art→WBS: ' + silByWbs.size + ' WBS con SIL, ' + unmappedArts.length + ' articoli non mappati');

  // STEP 3: WBS → Activity P6
  const distrib = [];
  const unmappedWbs = [];

  silByWbs.forEach((silWbs, wbs) => {
    const acts = p6.byWbs.get(wbs) || [];
    if(!acts.length){
      if(silWbs.total > 0) logMsg('⚠ WBS non trovata in P6: ' + wbs + ' (€' + fmt(silWbs.total) + ')');
      if(silWbs.total > 0) unmappedWbs.push({wbs, importo: silWbs.total, desWbs: silWbs.desWbs});
      return;
    }

    const totalCost = acts.reduce((s, a) => s + a.cost, 0);
    const totalPhys = acts.reduce((s, a) => s + a.phys, 0);

    acts.forEach(act => {
      let peso, method;
      if(totalCost > 0){
        peso = act.cost/totalCost;
        method = 'COST';
      } else if(totalPhys > 0){
        peso = act.phys/totalPhys;
        method = 'PHY';
      } else{
        peso = 1/acts.length;
        method = 'EQ';
      }

      distrib.push({
        actId: act.actId, wbs, desWbs: silWbs.desWbs,
        method, silImporto: silWbs.total * peso,
        p6Cost: act.cost, physPct: act.phys,
        status: act.status, actName: act.name,
        delta: silWbs.total * peso - act.cost
      });
    });
  });

  logMsg('Distribuzione: ' + distrib.length + ' attività P6');

  // KPI
  const totalSil = [...silByWbs.values()].reduce((s, v) => s + v.total, 0);
  const totalSilRaw = allSil.reduce((s, i) => s + i.importo, 0);
  const totalP6 = distrib.reduce((s, d) => s + d.p6Cost, 0);
  const totalBudget = [...budget.byWbs.values()].reduce((s, v) => s + v.total, 0);
  const cpi = totalP6 > 0 ? totalSil/totalP6 : null;

  // Riepilogo per WBS
  const summaryByWbs = new Map();
  distrib.forEach(d => {
    if(!summaryByWbs.has(d.wbs)) summaryByWbs.set(d.wbs, {wbs: d.wbs, desWbs: d.desWbs, silTot: 0, p6Tot: 0, acts: 0});
    const s = summaryByWbs.get(d.wbs);
    s.silTot += d.silImporto; s.p6Tot += d.p6Cost; s.acts++;
  });

  // Alert
  const alerts = [];
  distrib.forEach(d => {
    if(d.silImporto > 0 && d.physPct === 0 && d.status && !d.status.toLowerCase().includes('not started'))
      alerts.push({type: 'warn', actId: d.actId, wbs: d.wbs, msg: 'SIL €' + fmt(d.silImporto) + ' ma % fisica = 0', method: d.method, actName: d.actName});
  });
  unmappedArts.forEach(u => alerts.push({type: 'err', actId: '—', wbs: '—', msg: 'Articolo non in Budget: "' + u.art + '" — €' + fmt(u.importo), method: '—', actName: ''}));
  unmappedWbs.forEach(u => alerts.push({type: 'err', actId: '—', wbs: u.wbs, msg: 'WBS "' + u.wbs + '" non in Export P6 — €' + fmt(u.importo), method: '—', actName: ''}));

  // Deviazioni
  const deviazioni = distrib
    .filter(d => Math.abs(d.delta) >= 5000)
    .map(d => ({...d, absDelta: Math.abs(d.delta), pct: d.p6Cost > 0 ? Math.abs(d.delta)/d.p6Cost : null}));

  S.result = {distrib, summaryByWbs, alerts, deviazioni, unmappedArts, unmappedWbs,
            totalSil, totalSilRaw, totalP6, totalBudget, cpi, budget, p6, silByWbs, allSil};
  renderResults();
}

// ── RENDER DEI RISULTATI ────────────────────────────────────────────────────────────────────

/**
 * Renderizza i risultati del bridge nell'interfaccia
 */
function renderResults(){
  const R = S.result; if(!R) return;

  document.getElementById('pCpiV').textContent = R.cpi ? R.cpi.toFixed(3) : '—';
  document.getElementById('pCpi').className = 'pill ' + (R.cpi === null ? 'info' : R.cpi >= 0.95 && R.cpi <= 1.05 ? 'ok' : R.cpi > 1.05 ? 'warn' : 'err');

  document.getElementById('bnrNoBudget').classList.toggle('show', !S.budget);
  const totUnm = R.unmappedArts.length + R.unmappedWbs.length;
  document.getElementById('bnrUnmapped').classList.toggle('show', totUnm > 0);
  if(totUnm > 0){
    const lostEur = R.unmappedArts.reduce((s, u) => s + u.importo, 0) + R.unmappedWbs.reduce((s, u) => s + u.importo, 0);
    document.getElementById('bnrUnmTxt').textContent =
      R.unmappedArts.length + ' articoli non in Budget + ' + R.unmappedWbs.length + ' WBS non in P6 — €' + fmt(lostEur);
  }
  document.getElementById('bnrOk').classList.toggle('show', totUnm === 0 && R.distrib.length > 0);
  document.getElementById('btnExpHdr').classList.toggle('show', R.distrib.length > 0);

  ['f1','f2','f3','f4','f5','f6'].forEach(id => document.getElementById(id).classList.remove('done','warn'));
  if(S.budget) document.getElementById('f1').classList.add('done');
  if(S.sil)    document.getElementById('f2').classList.add('done');
  if(S.silI)   document.getElementById('f3').classList.add('done');
  else         document.getElementById('f3').classList.add('warn');
  if(S.p6)     document.getElementById('f4').classList.add('done');
  document.getElementById('f5').classList.add('done');
  document.getElementById('f6').classList.add(R.alerts.length > 0 ? 'warn' : 'done');

  const out = document.getElementById('outArea'); out.innerHTML = '';

  // KPI card
  const kpiCard = makeCard('&#128202; KPI Bridge');
  const kg = el('div', 'kgrid');
  kg.appendChild(makeKpi(fmtE(R.totalSil), 'SIL Allocato (€)', 'ok'));
  kg.appendChild(makeKpi(fmtE(R.totalP6), 'P6 Costo (€)', R.totalP6 > 0 ? 'ok' : 'warn'));
  kg.appendChild(makeKpi(fmtE(R.totalBudget), 'Budget CPM (€)', 'ok'));
  kg.appendChild(makeKpi(R.cpi ? R.cpi.toFixed(3) : '—', 'CPI', R.cpi === null ? 'info' : R.cpi >= 0.95 && R.cpi <= 1.05 ? 'ok' : 'err'));
  kg.appendChild(makeKpi(R.distrib.length, 'Attività P6', R.distrib.length > 0 ? 'ok' : 'warn'));
  kg.appendChild(makeKpi(R.alerts.length, 'Alert', R.alerts.length === 0 ? 'ok' : 'err'));
  kg.appendChild(makeKpi(R.deviazioni.length, 'Deviazioni', R.deviazioni.length === 0 ? 'ok' : 'warn'));
  const mappedPct = R.totalSilRaw > 0 ? Math.round(R.totalSil/R.totalSilRaw*100) : 0;
  kg.appendChild(makeKpi(mappedPct + '%', 'SIL Mappato', mappedPct > 95 ? 'ok' : 'warn'));
  kpiCard.querySelector('.card-body').appendChild(kg);
  out.appendChild(kpiCard);

  // Tab card
  const tabCard = makeCard(''); tabCard.querySelector('.card-head').style.display = 'none';
  const tabs = el('div', 'tabs');
  tabs.innerHTML =
    '<div class="tab active" onclick="showTab(\'riepilogo\',this)">Riepilogo WBS</div>'+
    '<div class="tab" onclick="showTab(\'distrib\',this)">Distribuzione P6 <span class="tn tn-bri">' + R.distrib.length + '</span></div>'+
    '<div class="tab" onclick="showTab(\'alert\',this)">Alert <span class="tn tn-alert">' + R.alerts.length + '</span></div>'+
    '<div class="tab" onclick="showTab(\'deviazioni\',this)">Deviazioni <span class="tn tn-dev">' + R.deviazioni.length + '</span></div>'+
    '<div class="tab" onclick="showTab(\'export\',this)">Export P6 <span class="tn tn-exp">' + R.distrib.length + '</span></div>'+
    '<div class="tab" onclick="showTab(\'log\',this)">&#128196; Log</div>';
  tabCard.appendChild(tabs);
  const tcWrap = el('div', '');

  // TAB RIEPILOGO
  const tcR = el('div', 'tc card-body active'); tcR.id = 'tc-riepilogo';
  const wbsKeys = [...R.summaryByWbs.keys()].sort();
  if(!wbsKeys.length){tcR.innerHTML = '<div class="empty"><div class="e-ico">&#128236;</div><div class="e-title">Nessun dato</div></div>';}
  else{
    const scr = el('div', 'card-scroll');
    const t = buildTable(['WBS', 'Des. WBS', 'SIL (€)', 'P6 Costo (€)', 'Delta (€)', 'N. Att.', 'Budget (€)', 'Status']);
    const tb = t.querySelector('tbody');
    let tS = 0, tP = 0;
    wbsKeys.forEach(wbs => {
      const s = R.summaryByWbs.get(wbs);
      const bud = R.budget.byWbs.get(wbs);
      const delta = s.silTot - s.p6Tot;
      tS += s.silTot; tP += s.p6Tot;
      tb.appendChild(buildRow([wbs, s.desWbs || '—', fmtE(s.silTot), fmtE(s.p6Tot), fmtDelta(delta), s.acts,
        bud ? fmtE(bud.total) : '—',
        delta === 0 ? '<span class="tag tag-ok">OK</span>' : delta > 0 ? '<span class="tag tag-warn">SIL&gt;P6</span>' : '<span class="tag tag-err">P6&gt;SIL</span>'
      ], delta > 0 ? 'r-over' : delta < 0 ? 'r-miss' : 'r-ok'));
    });
    R.unmappedWbs.forEach(u => {
      tb.appendChild(buildRow([u.wbs, u.desWbs || '—', fmtE(u.importo), '—', '—', '—', '—', '<span class="tag tag-err">NO P6</span>'], 'r-miss'));
    });
    tb.appendChild(buildRow(['TOTALE', '', fmtE(tS), fmtE(tP), fmtDelta(tS-tP), R.distrib.length, '', ''], 'r-total'));
    scr.appendChild(t); tcR.appendChild(scr);
  }
  tcWrap.appendChild(tcR);

  // TAB DISTRIBUZIONE
  const tcD = el('div', 'tc card-body'); tcD.id = 'tc-distrib';
  if(!R.distrib.length){tcD.innerHTML = '<div class="empty"><div class="e-ico">&#128236;</div><div class="e-title">Nessuna attività</div></div>';}
  else{
    const scr = el('div', 'card-scroll');
    const t = buildTable(['Activity ID', 'Nome Attività', 'WBS', 'Des. WBS', 'Metodo', 'SIL Alloc. (€)', 'P6 Costo (€)', 'Phys%', 'Status', 'Delta (€)']);
    const tb = t.querySelector('tbody');
    R.distrib.forEach(d => {
      const mt = d.method === 'COST' ? '<span class="tag tag-phy">COST</span>' : d.method === 'PHY' ? '<span class="tag tag-bri">PHY</span>' : '<span class="tag tag-eq">EQ</span>';
      tb.appendChild(buildRow([d.actId, d.actName || '—', d.wbs, d.desWbs || '—', mt, fmtE(d.silImporto), fmtE(d.p6Cost), (d.physPct || 0).toFixed(1) + '%', d.status || '—', fmtDelta(d.delta)], d.method === 'COST' ? 'r-ok' : 'r-bri'));
    });
  }
  tcWrap.appendChild(tcD);

  // TAB ALERT
  const tcA = el('div', 'tc card-body'); tcA.id = 'tc-alert';
  if(!R.alerts.length){tcA.innerHTML = '<div class="empty"><div class="e-ico">&#10003;</div><div class="e-title">Nessun alert</div></div>';}
  else{
    const scr = el('div', 'card-scroll');
    const t = buildTable(['Tipo', 'Activity ID', 'Nome Attività', 'WBS', 'Messaggio', 'Metodo']);
    const tb = t.querySelector('tbody');
    R.alerts.forEach(a => {
      const tt = a.type === 'err' ? '<span class="tag tag-err">ERR</span>' : '<span class="tag tag-warn">WARN</span>';
      tb.appendChild(buildRow([tt, a.actId, a.actName || '—', a.wbs, a.msg, a.method]));
    });
  }
  tcWrap.appendChild(tcA);

  // TAB DEVIAZIONI
  const tcDv = el('div', 'tc card-body'); tcDv.id = 'tc-deviazioni';
  if(!R.deviazioni.length){tcDv.innerHTML = '<div class="empty"><div class="e-ico">&#10003;</div><div class="e-title">Nessuna deviazione</div><div class="e-desc">(soglia |delta| &ge; &euro;5.000)</div></div>';}
  else{
    const scr = el('div', 'card-scroll');
    const t = buildTable(['Activity ID', 'Nome Attività', 'WBS', 'SIL (€)', 'P6 (€)', 'Delta (€)', '&Delta;%', 'Severità']);
    const tb = t.querySelector('tbody');
    R.deviazioni.sort((a, b) => b.absDelta - a.absDelta).forEach(d => {
      const ps = d.pct !== null ? (d.pct*100).toFixed(1) + '%' : '&#8734;';
      const sv = d.absDelta > 50000 ? '<span class="tag tag-err">ALTA</span>' : d.absDelta > 10000 ? '<span class="tag tag-warn">MEDIA</span>' : '<span class="tag tag-ok">BASSA</span>';
      tb.appendChild(buildRow([d.actId, d.actName || '—', d.wbs, fmtE(d.silImporto), fmtE(d.p6Cost), fmtDelta(d.delta), ps, sv], d.absDelta > 50000 ? 'r-miss' : d.absDelta > 10000 ? 'r-over' : ''));
    });
  }
  tcWrap.appendChild(tcDv);

  // TAB EXPORT
  const tcE = el('div', 'tc card-body'); tcE.id = 'tc-export';
  const note = el('div', ''); note.style.cssText = 'font-size:.7rem;color:var(--tx2);margin-bottom:10px';
  note.innerHTML = '<strong>' + R.distrib.length + '</strong> attività &nbsp;|&nbsp; COST: <strong>' + R.distrib.filter(d => d.method === 'COST').length + '</strong> &nbsp;|&nbsp; PHY: <strong>' + R.distrib.filter(d => d.method === 'PHY').length + '</strong> &nbsp;|&nbsp; EQ: <strong>' + R.distrib.filter(d => d.method === 'EQ').length + '</strong>';
  tcE.appendChild(note);
  const br = el('div', 'btn-row'); br.style.marginBottom = '10px';
  br.innerHTML = '<button class="btn btn-green btn-sm" onclick="exportCSV()">&#128190; CSV per P6</button><button class="btn btn-purple btn-sm" onclick="exportXER()">&#128190; XER per P6</button>';
  tcE.appendChild(br);
  const scr2 = el('div', 'card-scroll');
  const t2 = buildTable(['Activity ID', 'Nome Attività', 'WBS', 'SIL Alloc. (€)', 'P6 Costo (€)', 'Delta (€)', 'Metodo', 'Status']);
  const tb2 = t2.querySelector('tbody');
  R.distrib.slice(0, 150).forEach(d => {
    const mt = d.method === 'COST' ? '<span class="tag tag-phy">COST</span>' : d.method === 'PHY' ? '<span class="tag tag-bri">PHY</span>' : '<span class="tag tag-eq">EQ</span>';
    tb2.appendChild(buildRow([d.actId, d.actName || '—', d.wbs, fmtE(d.silImporto), fmtE(d.p6Cost), fmtDelta(d.delta), mt, d.status || '—']));
  });
  if(R.distrib.length > 150){
    const tr = document.createElement('tr');const td = document.createElement('td');
    td.colSpan = 8;td.textContent = '... e altre ' + (R.distrib.length - 150) + ' righe';
    td.style.cssText = 'text-align:center;color:var(--tx3)';tr.appendChild(td);tb2.appendChild(tr);
  }
  scr2.appendChild(t2); tcE.appendChild(scr2);
  tcWrap.appendChild(tcE);

  // TAB LOG
  const tcL = el('div', 'tc card-body'); tcL.id = 'tc-log';
  const lb = el('div', 'log-box'); lb.id = 'logBox';
  tcL.appendChild(lb);
  tcWrap.appendChild(tcL);

  tabCard.appendChild(tcWrap);
  out.appendChild(tabCard);
  setTimeout(() => {const lb = document.getElementById('logBox');if(lb) lb.textContent = _log.join('\n');}, 30);
}

// ── FUNZIONI DI ESPORTAZIONE ────────────────────────────────────────────────────────────────────

/**
 * Esporta i dati in formato CSV
 */
function exportCSV(){
  const R = S.result; if(!R || !R.distrib.length){showToast('err', 'Nessun dato');return;}
  const lines = ['Activity ID,Nome Attività,WBS,Des. WBS,Metodo,SIL Allocato (EUR),P6 Costo Attuale (EUR),Delta (EUR),Phys%,Status'];
  R.distrib.forEach(d => {
    lines.push('"' + d.actId + '","' + (d.actName || '') + '","' + d.wbs + '","' + (d.desWbs || '') + '",' + d.method + ',' +
      d.silImporto.toFixed(2) + ',' + d.p6Cost.toFixed(2) + ',' + d.delta.toFixed(2) + ',' +
      (d.physPct || 0).toFixed(2) + ',"' + (d.status || '') + '"');
  });
  download('bridge_export_p6.csv', lines.join('\r\n'), 'text/csv');
  showToast('ok', 'CSV esportato (' + R.distrib.length + ' righe)');
}

/**
 * Esporta i dati in formato XER
 */
function exportXER(){
  const R = S.result; if(!R || !R.distrib.length){showToast('err', 'Nessun dato');return;}
  const lines = [
    'ERMHDR\t4.1\t2006-05-24\tCPM-P6-Bridge\tExport\tcosedil_pmo\tMinutes',
    '%T\tTASK',
    '%F\ttask_id\tproj_id\twbs_id\ttask_code\tact_this_per_work_qty\tact_cost\ttarget_cost\tphys_complete_pct'
  ];
  R.distrib.forEach(d => {
    const p6a = R.p6.byAct.get(d.actId) || {};
    lines.push('%R\t\t\t' + d.wbs + '\t' + d.actId + '\t' + d.silImporto.toFixed(2) + '\t' + d.silImporto.toFixed(2) + '\t' + d.silImporto.toFixed(2) + '\t' + (d.physPct || 0).toFixed(2));
  });
  lines.push('%E');
  download('bridge_update.xer', lines.join('\r\n'), 'text/plain');
  showToast('ok', 'XER esportato');
}

/**
 * Funzione wrapper per esportazione
 */
function doExport(){ exportCSV(); }

// ── FUNZIONI DI SUPPORTO PER L'INTERFACCIA ────────────────────────────────────────────────────────────────

/**
 * Aggiorna gli indicatori di stato
 */
function updatePills(){
  const map = {budget: 'pBudget', sil: 'pSil', silI: 'pSilI', p6: 'pP6'};
  const vmap = {budget: 'pBudgetV', sil: 'pSilV', silI: 'pSilIV', p6: 'pP6V'};
  let loaded = 0;
  Object.keys(map).forEach(k => {
    const pill = document.getElementById(map[k]);
    const vEl = document.getElementById(vmap[k]);
    if(S[k]){
      loaded++;
      pill.className = 'pill ok';
      const n = S[k].data ? S[k].data.length : S[k].rawRows ? S[k].rawRows.length : 0;
      vEl.textContent = n + ' righe';
    } else {pill.className = 'pill'; vEl.textContent = '—';}
  });
  document.getElementById('pDocsV').textContent = loaded + '/4';
  document.getElementById('pDocs').className = 'pill ' + (loaded >= 3 ? 'ok' : 'info');
}

/**
 * Azzera tutti i dati e reimposta l'interfaccia
 */
function resetAll(){
  Object.keys(S).forEach(k => S[k] = null);
  document.querySelectorAll('.dz').forEach(dz => dz.classList.remove('loaded'));
  document.querySelectorAll('.dz-info').forEach(e => e.textContent = '');
  document.querySelectorAll('input[type=file]').forEach(e => e.value = '');
  document.getElementById('outArea').innerHTML = '<div class="empty"><div class="e-ico">&#128279;</div><div class="e-title">Reset completato</div></div>';
  document.getElementById('btnCalc').disabled = true;
  document.getElementById('btnExpHdr').classList.remove('show');
  ['bnrNoBudget', 'bnrUnmapped', 'bnrOk'].forEach(id => document.getElementById(id).classList.remove('show'));
  updatePills(); _log = [];
}

/**
 * Crea una tabella HTML
 * @param {Array} headers - Intestazioni della tabella
 * @returns {HTMLTableElement} - Elemento tabella creato
 */
function buildTable(headers){
  const t = document.createElement('table');
  const thead = document.createElement('thead');
  const tr = document.createElement('tr');
  headers.forEach(h => {const th = document.createElement('th');th.innerHTML = h;tr.appendChild(th);});
  thead.appendChild(tr);t.appendChild(thead);
  t.appendChild(document.createElement('tbody'));
  return t;
}

/**
 * Crea una riga di tabella
 * @param {Array} cells - Contenuto delle celle
 * @param {string} rowClass - Classe CSS per la riga
 * @returns {HTMLTableRowElement} - Elemento riga creato
 */
function buildRow(cells, rowClass){
  const tr = document.createElement('tr');
  if(rowClass) tr.className = rowClass;
  cells.forEach(c => {const td = document.createElement('td');td.innerHTML = String(c ?? '');tr.appendChild(td);});
  return tr;
}

/**
 * Crea una card
 * @param {string} title - Titolo della card
 * @returns {HTMLDivElement} - Elemento card creato
 */
function makeCard(title){
  const card = el('div', 'card');
  const head = el('div', 'card-head');
  const h3 = el('h3', '');h3.innerHTML = title;head.appendChild(h3);
  card.appendChild(head);card.appendChild(el('div', 'card-body'));
  return card;
}

/**
 * Crea un elemento KPI
 * @param {string|number} val - Valore del KPI
 * @param {string} label - Etichetta del KPI
 * @param {string} cls - Classe per lo stato del KPI
 * @returns {HTMLDivElement} - Elemento KPI creato
 */
function makeKpi(val, label, cls){
  const k = el('div', 'kpi ' + (cls || ''));
  const kv = el('div', 'kv');kv.textContent = val;
  const kl = el('div', 'kl');kl.textContent = label;
  k.appendChild(kv);k.appendChild(kl);
  return k;
}

/**
 * Mostra una scheda specifica
 * @param {string} id - ID della scheda da mostrare
 * @param {HTMLElement} btn - Bottone che ha attivato la funzione
 */
function showTab(id, btn){
  document.querySelectorAll('.tc').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tc-' + id).classList.add('active');
  btn.classList.add('active');
  if(id === 'log'){const lb = document.getElementById('logBox');if(lb) lb.textContent = _log.join('\n');}
}

/**
 * Scarica un file
 * @param {string} name - Nome del file
 * @param {string} content - Contenuto del file
 * @param {string} mime - Tipo MIME del file
 */
function download(name, content, mime){
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([content], {type: mime}));
  a.download = name;a.click();
}

// Timer per gestire la visualizzazione dei toast
let toastTimer;

/**
 * Mostra un messaggio temporaneo
 * @param {string} type - Tipo di messaggio (ok, err, warn, info)
 * @param {string} msg - Testo del messaggio
 */
function showToast(type, msg){
  const t = document.getElementById('toast');
  t.textContent = msg;t.className = 'toast show ' + type;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3000);
}

/**
 * Mostra la finestra di aiuto
 */
function showHelp() {
  const helpContent = `
    <div style="padding: 15px;">
      <h3 style="margin-top: 0; color: var(--a2);">Guida Rapida - CPM → P6 Bridge v13.1</h3>
      
      <h4>Prerequisiti</h4>
      <ul style="margin-left: 20px; margin-bottom: 15px;">
        <li>File Excel (.xlsx, .xls) o CSV (.csv) contenente i dati</li>
        <li>Fogli denominati correttamente: "BUDGET", "SIL diretti", "SIL indiretti", "EXPORT_P6"</li>
      </ul>
      
      <h4>Passaggi</h4>
      <ol style="margin-left: 20px; margin-bottom: 15px;">
        <li>Carica il file "Budget CPM" (foglio BUDGET)</li>
        <li>Carica il file "SIL Diretti" (foglio SIL diretti)</li>
        <li>Opzionale: Carica il file "SIL Indiretti" (foglio SIL indiretti)</li>
        <li>Carica il file "Export P6" (foglio EXPORT_P6)</li>
        <li>Clicca su "Calcola Bridge"</li>
        <li>Analizza i risultati e scarica l'export</li>
      </ol>
      
      <h4>Mappatura Colonne Comuni</h4>
      <p>Il sistema riconosce automaticamente queste colonne:</p>
      <ul style="margin-left: 20px; margin-bottom: 15px;">
        <li><strong>Budget CPM:</strong> "Cod. WBS", "Articolo", "Importo Costo (€)"</li>
        <li><strong>SIL Diretti:</strong> "Cod. S.I.L.", "Articolo", "Importo"</li>
        <li><strong>Export P6:</strong> "task_code", "wbs_id", "act_cost", "calc_phys_complete_pct"</li>
      </ul>
      
      <h4>Risultati</h4>
      <ul style="margin-left: 20px; margin-bottom: 15px;">
        <li><strong>Riepilogo WBS:</strong> Visualizza la corrispondenza tra WBS e costi</li>
        <li><strong>Distribuzione P6:</strong> Come i costi SIL sono distribuiti alle attività P6</li>
        <li><strong>Alert:</strong> Problemi identificati durante il mapping</li>
        <li><strong>Deviazioni:</strong> Grandi differenze tra costi allocati e costi P6</li>
      </ul>
      
      <p style="font-style: italic;">Per maggiori informazioni, contattare il team PMO Cosedil.</p>
    </div>`;
  
  // Mostra la guida in un elemento temporaneo
  const tempDiv = document.createElement('div');
  tempDiv.className = 'card';
  tempDiv.innerHTML = `<div class="card-head"><h3>Guida Rapida</h3></div><div class="card-body">${helpContent}</div>`;
  
  const outArea = document.getElementById('outArea');
  outArea.innerHTML = '';
  outArea.appendChild(tempDiv);
  outArea.insertAdjacentHTML('beforeend', '<div class="btn-row" style="margin-top: 15px;"><button class="btn btn-ghost" onclick="resetOutputToInfo()">Torna indietro</button></div>');
}

/**
 * Ripristina il contenuto iniziale della sezione output
 */
function resetOutputToInfo() {
  // Ricrea il contenuto iniziale
  document.getElementById('outArea').innerHTML = `
    <div class="empty">
      <div class="e-ico">&#128279;</div>
      <div class="e-title">Come funziona il Bridge v13.1</div>
      <div class="e-desc">
        <strong>Mapping automatico a 3 livelli:</strong><br><br>
        1&#65039;&#8419; <strong>SIL Articolo</strong> &#8594; <strong>Budget Articolo</strong> &#8594; ottieni <strong>Cod. WBS</strong><br>
        2&#65039;&#8419; <strong>Cod. WBS</strong> &#8594; <strong>Export P6 wbs_id</strong> &#8594; lista attivit&agrave; P6<br>
        3&#65039;&#8419; Distribuisci SIL tra attivit&agrave; proporzionalmente al <strong>act_cost</strong><br><br>
        <strong>Carica lo stesso file Excel</strong> su tutti i campi &mdash;<br>
        il tool seleziona automaticamente il foglio giusto.<br><br>
        
        <strong>Formati supportati:</strong> .xlsx, .xls, .csv, .txt<br>
        <strong>Nomi fogli riconosciuti:</strong> BUDGET, SIL diretti/indiretti, EXPORT_P6<br><br>
        
        <button class="btn btn-primary" onclick="showHelp()">Visualizza guida completa</button>
      </div>
    </div>`;
}