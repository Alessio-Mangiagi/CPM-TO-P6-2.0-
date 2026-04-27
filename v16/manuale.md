# Manuale Utente — CPM → P6 v16

**Cosedil PMO · Bridge Budget vs SIL vs P6**

---

## Panoramica

Lo strumento **CPM → P6 v16** è un'applicazione web client-side che esegue il mapping automatico tra i dati di avanzamento economico CPM (SIL) e le attività di Oracle Primavera P6. Tutto il calcolo avviene nel browser: nessun dato viene inviato a server esterni.

**Flusso logico in tre passi:**

```
SIL Articolo → Budget Articolo → Cod. WBS
Cod. WBS     → Export P6 wbs_id → Lista Attività P6
SIL Allocato → distribuito tra attività (proporzionale a act_cost o % fisica)
```

---

## File richiesti

| File          | Etichetta        | Obbligatorio | Foglio atteso     |
| ------------- | ---------------- | :-----------: | ----------------- |
| Budget CPM    | 💰 Budget CPM    | **Sì** | `BUDGET`        |
| SIL Diretti   | 📊 SIL Diretti   | **Sì** | `SIL diretti`   |
| Export P6     | 📅 Export P6     | **Sì** | `EXPORT_P6`     |
| SIL Indiretti | 📋 SIL Indiretti |      No      | `SIL indiretti` |

Formati accettati: `.xlsx`, `.xls`, `.csv`, `.txt`
Limite dimensione: **30 MB** per file
Limite righe: **50.000 righe** per file (troncamento automatico oltre)

### Struttura minima colonne attese

**Budget CPM**

- `Cod. WBS` — codice WBS
- `Articolo` — descrizione articolo
- `Importo Costo (€)` — importo budget

**SIL Diretti / Indiretti**

- `Cod. S.I.L.` — numero SIL
- `Articolo` — descrizione articolo
- `Importo` — importo SIL

**Export P6**

- `task_code` (o `Activity ID`) — codice attività
- `wbs_id` (o `WBS Code`) — codice WBS Primavera
- `act_cost` (o `Actual Total Cost`) — costo attuale
- `calc_phys_complete_pct` (o `Physical % Complete`) — avanzamento fisico
- `status_code` (o `Activity Status`) — stato attività
- `act_name` (o `Activity Name`) — nome attività

---

## Procedura di utilizzo

### 1. Caricamento file

Trascina i file nelle rispettive aree (drag & drop) oppure fai clic sull'area per aprire il selettore file.

Dopo il caricamento corretto, ogni area mostra:

- La barra in alto aggiorna il contatore **Caricati X/4**
- Il nome del file e il numero di righe rilevate
- Un'anteprima delle prime righe del file
- La pillola corrispondente diventa **verde**

Il pulsante **▶ Calcola Bridge** si attiva automaticamente quando sono caricati i tre file obbligatori.

### 2. Calcolo

Premi **▶ Calcola Bridge**. La barra di avanzamento mostra le fasi:

| %    | Fase                             |
| ---- | -------------------------------- |
| 10%  | Parsing Budget CPM               |
| 30%  | Parsing SIL Diretti/Indiretti    |
| 50%  | Parsing Export P6                |
| 60%  | Aggregazione SIL per articolo    |
| 75%  | Mapping Articolo → WBS          |
| 85%  | Distribuzione WBS → Activity P6 |
| 95%  | Calcolo KPI e Alert              |
| 100% | Completato                       |

### 3. Lettura risultati

I risultati si articolano in quattro schede:

#### Scheda Riepilogo WBS

Tabella aggregata per codice WBS con colonne:

- **WBS** — codice WBS
- **Des. WBS** — descrizione
- **SIL (€)** — SIL allocato su questa WBS
- **P6 Costo (€)** — costo attuale da Primavera
- **Delta (€)** — differenza SIL − P6
- **N. Att.** — numero attività P6 sotto la WBS
- **Budget (€)** — budget CPM per la WBS
- **Status** — `SIL>P6` / `P6>SIL` / `OK`

Riga **TOTALE** in fondo alla tabella.

#### Scheda Distribuzione P6

Dettaglio per singola attività P6 con metodo di distribuzione usato:

| Tag      | Metodo                       | Quando applicato                  |
| -------- | ---------------------------- | --------------------------------- |
| `COST` | Proporzionale a `act_cost` | `act_cost` totale WBS > 0       |
| `PHY`  | Proporzionale a `% fisica` | `act_cost` = 0 ma `phys%` > 0 |
| `EQ`   | Quota uguale tra attività   | Entrambi i valori = 0             |

#### Scheda Alert

Avvisi automatici generati:

- **⚠️ WARN** — attività con SIL assegnato ma `% fisica = 0` (e non in stato "Not Started")
- **🚨 ERR** — articoli SIL non trovati nel Budget CPM
- **🚨 ERR** — WBS presenti nel SIL ma assenti dall'Export P6

#### Scheda Deviazioni

Elenco delle attività il cui delta `|SIL − P6 Costo|` supera la soglia configurata.
Colonne: WBS, Activity ID, Delta, Valore assoluto, % sul P6 Costo.

#### Scheda Log

Registro cronologico dell'elaborazione con timestamp.

---

## KPI principali

| KPI                         | Descrizione                                              |
| --------------------------- | -------------------------------------------------------- |
| **SIL Allocato (€)** | Totale SIL correttamente mappato su attività P6         |
| **P6 Costo (€)**     | Somma `act_cost` di tutte le attività mappate         |
| **Budget CPM (€)**   | Totale budget da file CPM                                |
| **CPI**               | Cost Performance Index = SIL Allocato / P6 Costo         |
| **SIL Mappato %**     | Quota del SIL totale che ha trovato corrispondenza in P6 |

Il **CPI** è semaforo:

- Verde: 0,95 ≤ CPI ≤ 1,05
- Arancio: CPI > 1,05
- Rosso: CPI < 0,95

---

## Esportazione

### Pulsante "⬇ Esporta Finale" (header)

Disponibile dopo il calcolo. Scarica il risultato finale in formato Excel.

### Reset

Il pulsante **↻ Reset** cancella tutti i file caricati e i risultati, riportando l'applicazione allo stato iniziale. La soglia di deviazione impostata viene mantenuta (salvata in localStorage).

---

## Messaggi di stato

| Banner                                    | Significato                                           |
| ----------------------------------------- | ----------------------------------------------------- |
| ⚠ Budget CPM mancante                    | Nessun Budget caricato — il mapping non è possibile |
| 🚨 X art. non in Budget + Y WBS non in P6 | Articoli o WBS senza corrispondenza                   |
| ✓ Bridge completo                        | Tutti gli articoli SIL mappati correttamente          |

---

## Limiti e note operative

- Il file Export P6 accetta sia il formato di export standard Primavera (`task_code`, `wbs_id`) sia il formato report con intestazioni estese (`Activity ID`, `WBS Code`, ecc.): la conversione avviene automaticamente.
- Il matching Articolo → WBS è case-insensitive e normalizza gli spazi.
- Il codice WBS è normalizzato: zeri iniziali rimossi dai segmenti numerici, segmenti alfanumerici mantenuti invariati.
- Il matching fallisce con fuzzy search fino a 8 caratteri iniziali se il match esatto non viene trovato.
- I dati non vengono mai inviati fuori dal browser (CSP `connect-src 'none'`).

Creato da Alessio Mangiagi ><(((º> sabusabu <º)))><
