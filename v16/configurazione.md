# Configurazione — CPM → P6 v16

**Cosedil PMO · Parametri tecnici e personalizzazione**

---

## Soglia Deviazione

**Campo:** `Soglia Deviazione (€)` — pannello laterale sinistro, sezione *Configurazione*

Controlla quale scarto minimo tra SIL allocato e P6 Costo viene segnalato nella scheda **Deviazioni**.

| Parametro          | Valore            |
| ------------------ | ----------------- |
| Valore predefinito | `5.000 €`      |
| Minimo             | `100 €`        |
| Passo              | `500 €`        |
| Massimo accettato  | `10.000.000 €` |

La soglia viene salvata in `localStorage` con chiave `bridge_threshold` e ripristinata automaticamente alla riapertura dell'applicazione.

**Effetto:** un'attività compare nella scheda Deviazioni se `|SIL allocato − P6 act_cost| ≥ soglia`.

---

## Nomi foglio attesi (SHEET_HINTS)

L'applicazione cerca il foglio corretto all'interno del file Excel usando un elenco di alias (case-insensitive, corrispondenza parziale):

| File          | Alias cercati nell'ordine                                                  |
| ------------- | -------------------------------------------------------------------------- |
| Budget CPM    | `BUDGET`, `Budget`, `budget`, `BUDGET CPM`                         |
| SIL Diretti   | `SIL diretti`, `SIL Diretti`, `SIL_DIRETTI`, `SIL DIRETTI`         |
| SIL Indiretti | `SIL indiretti`, `SIL Indiretti`, `SIL_INDIRETTI`, `SIL INDIRETTI` |
| Export P6     | `EXPORT_P6`, `Export_P6`, `export_p6`, `EXPORT P6`                 |

Se nessun alias corrisponde viene usato il **primo foglio** del workbook.

---

## Colonne riconosciute (COLUMNS)

Le intestazioni vengono cercate nelle prime 12 righe del file; il match è su sottostringa case-insensitive.

**Budget CPM**

```
cod. wbs | articolo | importo costo | costo | budget | wbs | codice wbs
```

**SIL Diretti / Indiretti**

```
cod. s.i.l. | articolo | importo | sil | costo | codice sil
```

**Export P6**

```
task_code | wbs_id | act_cost | calc_phys_complete_pct | status_code | act_name
```

Per l'Export P6 è supportata anche la variante con intestazioni estese (`Activity ID`, `WBS Code`, `Actual Total Cost`, `Physical % Complete`, `Activity Status`, `Activity Name`): la conversione avviene automaticamente prima del parsing.

---

## Normalizzazione WBS (normWBS)

I codici WBS vengono normalizzati prima di qualsiasi confronto:

1. Conversione in uppercase, rimozione spazi iniziali/finali
2. Split per carattere `.`
3. Segmenti **numerici puri**: rimossi gli zeri iniziali (`01` → `1`, `001.02` → `1.2`)
4. Segmenti **alfanumerici**: mantenuti invariati

Esempi:

| Input        | Normalizzato |
| ------------ | ------------ |
| `01.002.A` | `1.2.A`    |
| `WBS.010`  | `WBS.10`   |
| `A.B.C`    | `A.B.C`    |

---

## Normalizzazione Articolo (normArt)

- Conversione in lowercase
- Rimozione spazi multipli (collasso a singolo spazio)
- Trim iniziale e finale

Il matching Articolo → WBS usa tre livelli di fallback in ordine:

1. **Match esatto** sull'articolo normalizzato
2. **Match prefisso** — confronto sui primi 8 caratteri
3. **Match fuzzy** — sottostringa di 5 caratteri

---

## Limiti tecnici

| Parametro                          | Valore                          |
| ---------------------------------- | ------------------------------- |
| Dimensione massima file            | 30 MB                           |
| Righe massime per file             | 50.000 (troncamento automatico) |
| Righe esaminate per trovare header | 12                              |
| Righe anteprima mostrate in UI     | 5                               |
| Timeout alert errore UI            | 6.000 ms                        |
| Timeout toast notifica             | 3.500 ms                        |

---

## Metodi di distribuzione SIL → Attività P6

Quando una WBS ha più attività P6 associate, il SIL viene ripartito con questa priorità:

| Priorità | Metodo                       | Condizione                             | Tag      |
| --------- | ---------------------------- | -------------------------------------- | -------- |
| 1         | Proporzionale a `act_cost` | Somma `act_cost` della WBS > 0       | `COST` |
| 2         | Proporzionale a `% fisica` | `act_cost` = 0 e somma `phys%` > 0 | `PHY`  |
| 3         | Quota equa                   | Tutti i valori = 0                     | `EQ`   |

---

## Sicurezza e privacy

- **Content Security Policy** attiva: `default-src 'none'; script-src 'self'; style-src 'self' 'unsafe-inline'; connect-src 'none'`
- Nessuna chiamata di rete: tutti i calcoli avvengono interamente nel browser
- Output DOM costruito esclusivamente via `textContent` / `createElement` — nessuna interpolazione HTML diretta da dati utente
- Messaggi di errore sanitizzati (rimossi `<` e `>`) prima di essere visualizzati

---

## LocalStorage

| Chiave               | Tipo       | Contenuto                                    |
| -------------------- | ---------- | -------------------------------------------- |
| `bridge_threshold` | `number` | Soglia deviazione in euro (default `5000`) |

Nessun altro dato viene persistito tra sessioni.

---

## Requisiti browser

- Browser moderno con supporto ES2020+ (Chrome 90+, Firefox 88+, Edge 90+, Safari 14+)
- JavaScript abilitato
- Almeno 512 MB RAM disponibile per file da 30 MB / 50.000 righe

Creato da Alessio Mangiagi ><(((º> sabusabu <º)))><
