# CPM → P6 Bridge — Feature Reference
**COSEDIL S.p.A. · PMO** | Aggiornato: 2026-04-30

---

## Architettura file

```
Tk1_v1.0/
├── cpm_to_p6.py            # core: logica pura, usabile anche da CLI
├── cpm_to_p6_controller.py # controller: do_analizza/rigenera/export
├── cpm_to_p6_gui.py        # view: CustomTkinter UI
└── file/
    └── CPM to P6_rev.0.13live.xlsx   # file Bridge di default
```

**Pattern**: View (`App`) → Controller (`AppController`) → Core (`cpm_to_p6`)  
Il controller riceve `app` in `__init__` e usa `app.after(0, …)` per aggiornamenti UI thread-safe.

---

## File Excel Bridge — fogli richiesti

| Foglio | Contenuto |
|--------|-----------|
| `MAPPING` | WBS → Activity ID P6, descrizione WBS, tipo nodo (DI/IN) |
| `BUDGET da CPM` | Articolo → costo unitario (header riga 4, dati da riga 5) |
| `INPUT da P6` | Dati XER: sezione `%T TASK` con `task_code`, `task_name`, `phys_complete_pct`, `target_drtn_hr_cnt` |
| `SIL diretti` | Record SIL con WBS, SIL#, data, articolo, qta, importo |
| `SIL indiretti` | Stessa struttura, colonne diverse (SI_WBS=13, SI_IMP=9) |
| `BRIDGE_SIL` | Output: generato/sovrascritto da "Rigenera" |
| `P6_IMPORT_PULITO` | Output: generato/sovrascritto da "Rigenera" |

---

## Core — `cpm_to_p6.py`

### Caricamento dati (`_load_all`)
- Apre Excel in `read_only + data_only`
- Verifica presenza dei 5 fogli richiesti (esce con errore se mancano)
- Restituisce: `wbs_to_acts, wbs_to_des, wbs_to_tipo, art_to_cost, act_info, sil_records, sil_max`

### MAPPING (`load_mapping`)
- `wbs_to_acts`: `{wbs: [act_id, …]}` — Activity ID separati da virgola; ignora "—" e "?"
- `wbs_to_des`: `{wbs: descrizione}`
- `wbs_to_tipo`: `{wbs: "DI" | "IN"}`

### BUDGET (`load_budget`)
- `{articolo: costo_unitario}` — usato per calcolare `Costo Distribuito`

### INPUT da P6 / XER (`load_p6_tasks`)
- Trova l'ultima sezione `%T TASK`
- Legge header `%F` dinamicamente (fallback su indici hardcoded)
- Restituisce `{task_code: {name, pct, dur}}`

### SIL records
- `load_sil_diretti`: `tipo_base = "Diretto"`, skip se WBS o SIL mancanti
- `load_sil_indiretti`: `tipo_base = "Indiretto"`, skip se qta=0 e imp=0

### Generazione BRIDGE_SIL (`generate_bridge`)
**Distribuzione importo per attività P6:**
1. Peso = `phys_complete_pct` proporzionale (somma pct > 0)
2. Fallback: `target_drtn_hr_cnt` proporzionale (se tutti pct = 0)
3. Fallback: peso uguale (se anche tutte dur = 0)

**Tipo finale:**
- WBS mappata → `tipo_base` (Diretto/Indiretto)
- WBS non mappata + tipo "IN" → `"Indiretto"`
- WBS = "SIC" → `"Sicurezza"`
- WBS non mappata altrimenti → `"MAPPING MANCANTE"` (Activity ID = "N/A")

**Colonne BRIDGE_SIL (0-based):**
| # | Campo |
|---|-------|
| 0 | SIL # |
| 1 | Data SIL |
| 2 | Cod. WBS CPM |
| 3 | Des. WBS CPM |
| 4 | Articolo CPM |
| 5 | Importo SIL Ricavo (€) |
| 6 | P6 Activity ID |
| 7 | P6 Activity Name |
| 8 | Peso Proporzionale |
| 9 | Importo Distribuito Ricavo (€) |
| 10 | Importo Cumulativo Ricavo (€) — calcolato post-sort |
| 11 | Tipo |
| 12 | Costo Distribuito (€) |

**Cumulativo**: calcolato dopo sort per `(act_id, data)`, per activity.

### Generazione P6_IMPORT_PULITO (`generate_p6_import`)
- Aggrega bridge per `act_id`: somma `imp_dist` totale e solo per `sil_corrente`
- Include **tutte** le attività XER (anche quelle senza costo → 0)
- Ordina per `(len(act_id), act_id)`

**Colonne output CSV / foglio:**
| # | Campo |
|---|-------|
| 0 | Activity ID |
| 1 | Activity Name |
| 2 | Actual This Period Cost |
| 3 | Actual Total Cost |

### Report RIEPILOGO (`show_riepilogo`)
- Quadratura per WBS: `Σ SIL` vs `Σ Bridge`
- Check: `abs(delta) ≤ 1.0 €` → "✓ OK" altrimenti "⚠ DELTA"
- Riga totale: "✓ QUADRA" / "⚠ NON QUADRA"

### Report ALERT (`show_alert`)
- Attività con `Importo Distribuito > 0` nel SIL corrente ma `phys_complete_pct = 0`
- Mostra metodo usato: `DUR_FALLBACK` (dur > 0) o `NESSUN_PESO`

### Scrittura Excel
- `_write_bridge`: cancella righe da 2 in poi, riscrive header + dati
- `_write_p6_import`:
  - `B4` = sil_corrente, `C4` = sil_max
  - `A6` = stringa riepilogo periodo + cumulativo
  - Riga 8 = header, righe da 9 = dati (max 4 colonne)

### Backup automatico
- Prima di `rigenera`: copia `{stem}_BACKUP.xlsx` nella stessa cartella

### CLI
```bash
python cpm_to_p6.py analizza [--file PATH] [--sil N]
python cpm_to_p6.py rigenera [--file PATH] [--sil N]
python cpm_to_p6.py export   [--file PATH] [--sil N] [--output PATH]
```
Default file: `file/CPM to P6_rev.0.13live.xlsx`  
Default SIL: max trovato nei dati

---

## Controller — `cpm_to_p6_controller.py`

`AppController(app)` — riceve ref alla finestra, espone:

| Metodo | Fa |
|--------|----|
| `do_analizza(wb_path, sil)` | Carica, genera bridge, mostra riepilogo+alert, aggiorna KPI |
| `do_rigenera(wb_path, sil)` | Analizza + scrive Excel + backup |
| `do_export(wb_path, sil, out_path)` | Analizza + scrive CSV |

Tutti usano `app.after(0, …)` per aggiornare UI dal thread worker.  
`stdout` già reindirizzato dal `TextboxWriter` prima della chiamata.

---

## GUI — `cpm_to_p6_gui.py`

### Pannello sinistro
| Widget | Funzione |
|--------|----------|
| Entry + "..." | Selettore file Excel (dialog in `file/`) |
| Spinner −/+ + Entry | SIL corrente (default 15, min 1) |
| Label "max disponibile" | Aggiornata dopo ogni operazione |
| Btn "🔍 Analizza" | `_run_analizza` → `ctrl.do_analizza` |
| Btn "⚡ Rigenera Bridge" | Conferma dialog → `ctrl.do_rigenera` |
| Btn "💾 Export CSV per P6" | asksaveasfilename → `ctrl.do_export` |
| KPI: check | "✓ QUADRA" / "⚠ NON QUADRA" / "✓ Rigenera OK" / "✓ CSV SIL N" |
| KPI: periodo | `Periodo SIL N: € …` |
| KPI: cumulativo | `Cumulativo 1→N: € …` |
| Switch Dark/Light | `ctk.set_appearance_mode` |

### Pannello destro
| Widget | Funzione |
|--------|----------|
| CTkTextbox Consolas 12pt | Log output (stdout redirect via `TextboxWriter`) |
| "Pulisci" | Svuota log |
| "Apri cartella output" | `os.startfile` sulla parent del file (attivo dopo rigenera/export) |
| "Apri CSV" | `os.startfile` sull'ultimo CSV esportato (attivo dopo export) |
| Status bar | Messaggio + colore (OK verde, WARN arancio, ERR rosso) |
| ProgressBar indeterminate | Visibile solo durante operazioni |

### Threading
- Tutte le operazioni girano in `daemon thread`
- `_busy` flag blocca doppio-click
- `_set_busy(True/False)` disabilita/riabilita i 3 pulsanti e gestisce la progress bar

### Colori
```python
ACCENT    = "#1f6aa5"
OK_COLOR  = "#2fa84f"
ERR_COLOR = "#c0392b"
WARN_COLOR = "#e67e22"
```

---

## Indici colonne Excel (costanti in `cpm_to_p6.py`)

```python
# SIL diretti
SD_WBS=0  SD_SIL=3  SD_DATA=4  SD_ART=6  SD_QTA=10  SD_IMP=11

# SIL indiretti
SI_SIL=1  SI_DATA=2  SI_ART=4  SI_QTA=8  SI_IMP=9  SI_WBS=13

# MAPPING
MP_WBS=7  MP_DES=8  MP_TIPO=9  MP_ACTS=11

# BUDGET da CPM
BD_ART=4  BD_COST=12

# XER TASK (fallback se header dinamico fallisce)
XER_ID=14  XER_NM=15  XER_PCT=5  XER_DUR=23
```

---

## Dipendenze

```
customtkinter
openpyxl
```
Standard library: `io, os, sys, csv, shutil, threading, argparse, datetime, pathlib, collections`
