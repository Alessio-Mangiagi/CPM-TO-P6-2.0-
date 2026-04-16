// Stato globale dell'applicazione
const state = {
    fileData: null,
    sheetsLoaded: { budget: false, sildir: false, silind: false, p6: false }
};

// Utility: Formatta valuta
const fmtMoney = (n) => new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' }).format(n);

// Utility: Mostra alert
function showAlert(msg, type = 'error') {
    const c = document.getElementById('alerts');
    const icons = { error: '🔴', warning: '⚠️', success: '✅' };
    c.innerHTML += `<div class="alert alert-${type}"><span class="alert-icon">${icons[type]}</span><span>${msg}</span></div>`;
    if (type === 'success') setTimeout(() => c.innerHTML = '', 4000);
}

// Gestione Drag & Drop + Click sulle card
document.querySelectorAll('.upload-area').forEach(area => {
    const type = area.dataset.target;
    
    area.addEventListener('dragover', e => { e.preventDefault(); area.classList.add('drag-over'); });
    area.addEventListener('dragleave', () => area.classList.remove('drag-over'));
    area.addEventListener('drop', e => {
        e.preventDefault(); area.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        if (file) processFile(file, type);
    });
    area.addEventListener('click', () => {
        const input = document.createElement('input');
        input.type = 'file'; input.accept = '.xlsx,.xls';
        input.onchange = () => processFile(input.files[0], type);
        input.click();
    });
});

// Lettura e parsing file Excel
function processFile(file, type) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showAlert('Formato non supportato. Usa .xlsx o .xls', 'warning');
        return;
    }

    document.getElementById(`info-${type}`).classList.add('show');
    document.querySelector(`#info-${type} .file-name`).textContent = file.name;
    document.querySelector(`[data-target="${type}"]`).classList.add('has-file');
    document.querySelector(`[data-target="${type}"]`).nextElementSibling.className = 'status-badge ready';
    document.querySelector(`[data-target="${type}"]`).nextElementSibling.innerHTML = '<span class="status-dot"></span> Caricato';

    const reader = new FileReader();
    reader.onload = e => {
        try {
            const wb = XLSX.read(e.target.result, { type: 'array' });
            state.sheetsLoaded[type] = true;
            state.fileData = { wb, fileName: file.name };
            checkReady();
        } catch (err) {
            showAlert('Errore nel parsing del file: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Verifica se tutti i fogli necessari sono caricati
function checkReady() {
    const required = ['budget', 'sildir', 'p6'];
    const allReady = required.every(k => state.sheetsLoaded[k]);
    document.getElementById('btn-calculate').disabled = !allReady;
}

// Logica principale del Bridge
document.getElementById('btn-calculate').addEventListener('click', () => {
    const btn = document.getElementById('btn-calculate');
    btn.classList.add('loading');
    btn.disabled = true;
    document.getElementById('alerts').innerHTML = '';

    setTimeout(() => {
        try {
            const { wb } = state.fileData;
            const getSheet = name => wb.SheetNames.find(s => s.toLowerCase() === name.toLowerCase());
            
            if (!getSheet('BUDGET')) throw new Error('Foglio "BUDGET" mancante');
            if (!getSheet('SIL diretti')) throw new Error('Foglio "SIL diretti" mancante');
            if (!getSheet('EXPORT_P6')) throw new Error('Foglio "EXPORT_P6" mancante');

            const budget = XLSX.utils.sheet_to_json(wb.Sheets[getSheet('BUDGET')]);
            const silDir = XLSX.utils.sheet_to_json(wb.Sheets[getSheet('SIL diretti')]);
            const silInd = getSheet('SIL indiretti') ? XLSX.utils.sheet_to_json(wb.Sheets[getSheet('SIL indiretti')]) : [];
            const p6 = XLSX.utils.sheet_to_json(wb.Sheets[getSheet('EXPORT_P6')]);

            // Step 1: SIL → WBS
            const budgetMap = {};
            budget.forEach(r => { if (r['Articolo'] && r['Cod. WBS']) budgetMap[r['Articolo'].trim()] = r['Cod. WBS'].trim(); });

            const silWBS = {};
            const processSIL = (arr) => arr.forEach(r => {
                const art = r['Articolo']?.trim();
                if (!art || !budgetMap[art]) return;
                const wbs = budgetMap[art];
                if (!silWBS[wbs]) silWBS[wbs] = 0;
                silWBS[wbs] += Number(r['Importo']) || 0;
            });

            processSIL(silDir);
            processSIL(silInd);

            // Step 2 & 3: WBS → P6 → Distribuzione
            const p6ByWBS = {};
            p6.forEach(r => {
                const wbs = r['wbs_id']?.trim() || r['WBS']?.trim();
                if (!wbs || !silWBS[wbs]) return;
                if (!p6ByWBS[wbs]) p6ByWBS[wbs] = { totalCost: 0, activities: [] };
                const cost = Number(r['act_cost']) || 0;
                p6ByWBS[wbs].activities.push({
                    id: r['activity_id'] || r['ID'],
                    name: r['activity_name'] || r['Nome'],
                    cost
                });
                p6ByWBS[wbs].totalCost += cost;
            });

            // Costruzione risultati
            const results = [];
            let totalSIL = 0, totalAct = 0, wbsCount = 0;

            Object.entries(silWBS).forEach(([wbs, silTot]) => {
                if (!p6ByWBS[wbs]) return;
                wbsCount++;
                const { activities, totalCost } = p6ByWBS[wbs];
                totalSIL += silTot;
                totalAct += activities.length;

                activities.forEach(act => {
                    const ratio = totalCost > 0 ? act.cost / totalCost : 0;
                    results.push({
                        wbs, actId: act.id, actName: act.name,
                        p6Cost: act.cost, silTot, ratio,
                        allocated: silTot * ratio
                    });
                });
            });

            if (results.length === 0) throw new Error('Nessun collegamento trovato tra SIL e P6. Verifica i nomi WBS/Articolo.');

            // Render tabella
            const tbody = document.getElementById('results-body');
            tbody.innerHTML = results.map(r => `
                <tr>
                    <td><strong>${r.wbs}</strong></td>
                    <td>${r.actId}</td>
                    <td>${r.actName}</td>
                    <td class="money">${fmtMoney(r.p6Cost)}</td>
                    <td class="money">${fmtMoney(r.silTot)}</td>
                    <td class="ratio">${(r.ratio * 100).toFixed(2)}%</td>
                    <td class="money" style="color: var(--primary); font-weight: 700;">${fmtMoney(r.allocated)}</td>
                </tr>
            `).join('');

            // Aggiorna statistiche
            document.getElementById('stat-wbs').textContent = wbsCount;
            document.getElementById('stat-act').textContent = totalAct;
            document.getElementById('stat-sil').textContent = fmtMoney(totalSIL);

            document.getElementById('results-container').style.display = 'block';
            showAlert('Bridge completo — Tutti gli articoli SIL mappati su attività P6.', 'success');

        } catch (err) {
            showAlert(err.message, 'error');
        } finally {
            btn.classList.remove('loading');
            btn.disabled = false;
        }
    }, 800);
});

// Reset completo
function resetAll() {
    state.sheetsLoaded = { budget: false, sildir: false, silind: false, p6: false };
    state.fileData = null;
    document.querySelectorAll('.upload-area').forEach(a => {
        a.classList.remove('has-file', 'drag-over');
        a.nextElementSibling.className = 'status-badge pending';
        a.nextElementSibling.innerHTML = '<span class="status-dot"></span> In attesa';
    });
    document.querySelectorAll('.file-info').forEach(f => f.classList.remove('show'));
    document.getElementById('alerts').innerHTML = '';
    document.getElementById('results-container').style.display = 'none';
    document.getElementById('results-body').innerHTML = '';
    document.getElementById('btn-calculate').disabled = true;
}