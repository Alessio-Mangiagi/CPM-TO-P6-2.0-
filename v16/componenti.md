# Componenti — CPM → P6 v16

## File

| File | Ruolo |
|------|-------|
| `interfaccia.html` | Struttura UI: header, pillole stato, barre progresso, drop zone, area risultati, modal documentazione |
| `interfaccia.css` | Stili: tema dark/light, layout a griglia, pillole, tabelle, toast, modal |
| `interfaccia.js` | Tutta la logica applicativa: parsing, calcolo bridge, rendering, export |
| `xlsx.full.min.js` | Libreria SheetJS — legge/scrive file `.xlsx`/`.xls` nel browser senza server |

---

## Struttura di `interfaccia.js`

### `StateManager`
Gestisce lo stato globale dell'applicazione. Tiene i 4 dataset caricati (`budget`, `sil`, `silI`, `p6`), il risultato del calcolo (`result`) e la soglia di deviazione. Sincronizza la soglia con `localStorage` così persiste tra sessioni.

### `SHEET_HINTS`
Dizionario che mappa ogni tipo di file agli alias del foglio Excel cercato. Se nessun alias corrisponde, viene usato il primo foglio del workbook.

### `COLUMNS`
Dizionario delle intestazioni attese per ogni tipo di file. Usato per trovare la riga header e validare che il file sia del formato corretto.

---

### Funzioni di utilità

| Funzione | Cosa fa |
|----------|---------|
| `logMsg` | Aggiunge una riga timestampata al log interno (visibile nella scheda Log) |
| `fmt` / `fmtE` | Formatta numeri in stile italiano (separatore migliaia / due decimali) |
| `fmtDelta` | Formatta delta con segno `+` o parentesi per negativi |
| `el` | Crea un elemento DOM con className — shortcut per `createElement` |
| `sanitizeForLog` | Rimuove caratteri di controllo dai messaggi di log |

### Normalizzatori

| Funzione | Cosa fa |
|----------|---------|
| `normWBS` | Porta il codice WBS in uppercase, rimuove zeri iniziali nei segmenti numerici (`01.002.A` → `1.2.A`) |
| `normArt` | Porta l'articolo in lowercase e collassa gli spazi multipli |
| `isDataRow` | Scarta righe vuote, emoji, totali, commenti |
| `excelDateToStr` | Converte i numeri seriali Excel in data ISO (`yyyy-mm-dd`) |
| `validateColumns` | Controlla che almeno una colonna attesa sia presente negli header |
| `extractCommessa` | Cerca il codice commessa nella prima riga del Budget CPM o nel nome file |

---

### Parsing

#### `parseFile(file, key)`
Legge il file con `FileReader` come `ArrayBuffer`. Per `.xlsx`/`.xls` usa SheetJS; per `.csv`/`.txt` divide per tabulazione o virgola. Per il file P6 restituisce solo le righe raw (il parsing strutturato avviene dopo). Per gli altri file trova la riga header nelle prime 12 righe e costruisce un array di oggetti chiave→valore.

#### `parseBudget(data)`
Costruisce due Map:
- `byArt` — articolo normalizzato → lista di `{ wbs, desWbs, importo }`
- `byWbs` — codice WBS normalizzato → `{ desWbs, total }`

#### `parseSIL(data)`
Restituisce la lista di righe SIL (`silNum`, `art`, `importo`, `dataSil`) e la data più recente trovata (usata per aggiornare "Ultimo SIL" nell'header).

#### `parseP6fromRawRows(rows)`
Supporta due formati Primavera (`task_code`/`wbs_id` e `Activity ID`/`WBS Code` — converte automaticamente). Costruisce:
- `byAct` — Activity ID → `{ actId, wbs, cost, phys, status, name }`
- `byWbs` — codice WBS → lista di attività

---

### Gestione file

#### Drop zone handlers
Ogni `.dz` intercetta click, tastiera (Enter/Spazio), drag-over e drop. All'evento chiama `handleFile`.

#### `handleFile(key, file)`
Controlla il limite 30 MB, chiama `parseFile`, salva il risultato nello state, aggiorna la drop zone e l'info file, mostra l'anteprima e aggiorna le pillole. Se il file è `budget` prova a estrarre il codice commessa. Se tutti i file obbligatori sono presenti abilita il bottone Calcola.

#### `showPreview(key, rawRows, headers)`
Mostra le prime 5 righe di dati in una mini-tabella sotto la drop zone, per conferma visiva del file caricato.

#### `canCompute` / `updatePills`
`canCompute` restituisce `true` se Budget, SIL Diretti e Export P6 sono caricati. `updatePills` aggiorna colore e valore di ogni pillola in base allo stato dei dati.

---

### Calcolo — `runBridge()`

Funzione asincrona principale. Cede il controllo al browser (`yieldToMain`) tra ogni fase per tenere la UI reattiva.

| Fase | Operazione |
|------|-----------|
| 10% | Parsing Budget CPM |
| 30% | Parsing SIL Diretti e Indiretti |
| 50% | Parsing Export P6 |
| 60% | Aggregazione SIL per articolo |
| 75% | Mapping Articolo → WBS (con tre livelli di fallback: esatto, prefisso 8 char, fuzzy 5 char) |
| 85% | Distribuzione WBS → Attività P6 (proporzionale a `act_cost`, poi a `phys%`, poi quota uguale) |
| 95% | Calcolo KPI, alert, deviazioni |
| 100% | `renderResults()` |

Il risultato viene salvato in `S.set('result', {...})` e contiene: `distrib`, `summaryByWbs`, `alerts`, `deviazioni`, KPI totali e riferimenti alle Map sorgente.

---

### Export — `exportXLSX()`

Usa SheetJS per creare un workbook con 4 fogli e lo scarica come file `.xlsx`:

| Foglio | Contenuto |
|--------|-----------|
| Distribuzione P6 | Riga per ogni attività P6 con SIL allocato, costo P6, metodo, delta |
| Riepilogo WBS | Riga per ogni WBS con totali SIL, P6, delta, budget |
| Alert | Tutti gli alert WARN ed ERR |
| Deviazioni | Attività con scarto assoluto oltre la soglia configurata |

Il nome file include il codice commessa e la data (`CPM-P6-Bridge_12345_2026-04-27.xlsx`).

---

### Rendering — `renderResults()`

Costruisce l'intera area risultati via DOM (nessun `innerHTML` con dati utente — solo `textContent`). Struttura:
1. Card KPI con griglia di 8 indicatori
2. Card con tab switcher (Riepilogo WBS / Distribuzione P6 / Alert / Deviazioni / Log)
3. Aggiornamento banner (bridge completo / articoli non mappati / WBS non in P6)
4. Aggiornamento pillole e step flow

#### Builder DOM sicuri
`buildTabs`, `makeCard`, `makeKpi`, `buildTable`, `buildRow` — costruiscono elementi senza interpolazione HTML. `buildRow` gestisce l'unica eccezione: celle con `{ __html: '...' }` contengono solo tag `<span>` hardcoded, mai dati utente.

---

### Modal documentazione

#### `mdToHtml(md)`
Renderer Markdown minimale (headings, tabelle, liste, code block, inline bold/code). Usato per mostrare Manuale e Configurazione senza caricare librerie esterne.

#### `openDoc` / `closeDoc`
Aprono/chiudono il modal overlay leggendo il contenuto dai tag `<pre hidden>` nell'HTML (non serve fetch di rete).

---

### Init — `DOMContentLoaded`

Collega tutti i listener agli elementi DOM: Calcola Bridge, Reset, Esporta Finale, Manuale, Config, chiusura modal (click overlay e tasto Escape).
