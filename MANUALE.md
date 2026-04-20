# Ponte CPM → P6 v13.1

## Descrizione del Progetto

Il **Ponte CPM → P6** è uno strumento web sviluppato per facilitare la conversione e il mapping dei dati tra sistemi di gestione progetti CPM (Cost Planning Management) ed il sistema Primavera P6. Questo strumento permette di effettuare un bridge automatico tra i dati finanziari e pianificazionali del CPM e le attività operative nel sistema P6.

## Funzionalità Principali

- **Mapping automatico Articolo → WBS → Activity**: Collegamento automatico tra articoli del budget CPM, WBS e attività P6
- **Caricamento multiplo di file**: Supporto per file Excel (.xlsx, .xls) e CSV (.csv, .txt)
- **Anteprima dei dati caricati**: Visualizzazione immediata delle prime righe dei file caricati
- **Calcolo del CPI (Cost Performance Index)**: Analisi della performance finanziaria del progetto
- **Identificazione deviazioni**: Rilevamento automatico di scostamenti significativi
- **Gestione di SIL (Scheda Impegno Lavori)**: Integrazione tra dati SIL e sistema P6
- **Esportazione multi-formato**: Generazione di file CSV, XER e XLSX per l'integrazione con P6

## Architettura del Sistema

Il progetto è composto dalle seguenti componenti:

### Frontend
- [interfaccia.html](file:///c:/Users/aless/Desktop/Cosedil/CPM-TO-P6/CPM-TO-P6-2.0-/interfaccia/interfaccia.html): Pagina principale dell'applicazione
- [interfaccia.js](file:///c:/Users/aless/Desktop/Cosedil/CPM-TO-P6/CPM-TO-P6-2.0-/interfaccia/interfaccia.js): Logica di business e gestione dei dati
- [interfaccia.css](file:///c:/Users/aless/Desktop/Cosedil/CPM-TO-P6/CPM-TO-P6-2.0-/interfaccia/interfaccia.css): Stile e presentazione dell'interfaccia

### Struttura delle cartelle
```
CPM-TO-P6/
├── interfaccia/           # File principali dell'applicazione
├── v14/                 # Versione precedente (v14.0)
├── MANUALE.md            # Documentazione del progetto
├── CMPTOP6.PY           # Script Python per alcune operazioni ausiliarie
└── vari file di esempio # File di test e dimostrazione
```

## Requisiti di Sistema

- Browser web moderno (supporto per JavaScript ES6+)
- Connessione Internet per caricare la libreria XLSX da CDN
- File di dati in formato Excel (.xlsx, .xls) o CSV (.csv, .txt)

## Installazione e Avvio

1. Clonare o scaricare il repository
2. Aprire il file [interfaccia/interfaccia.html](file:///c:/Users/aless/Desktop/Cosedil/CPM-TO-P6/CPM-TO-P6-2.0-/interfaccia/interfaccia.html) in un browser web
3. Non è richiesta alcuna installazione aggiuntiva

## Utilizzo dello Strumento

### Passaggi Principali

1. **Caricamento dati CPM**:
   - Budget CPM (foglio "BUDGET"): Contiene le colonne "Cod. WBS", "Articolo", "Importo Costo (€)"
   - SIL Diretti (foglio "SIL diretti"): Contiene colonne "Cod. S.I.L.", "Articolo", "Importo"
   - SIL Indiretti (foglio "SIL indiretti", opzionale): Formato simile ai SIL diretti

2. **Caricamento dati P6**:
   - Export P6 (foglio "EXPORT_P6"): Contiene colonne "task_code", "wbs_id", "act_cost", "calc_phys_complete_pct"

3. **Elaborazione**:
   - Fare clic su "Calcola Bridge" per avviare l'elaborazione
   - Visualizzare i risultati nelle diverse schede (Riepilogo WBS, Distribuzione P6, ecc.)

4. **Esportazione**:
   - Utilizzare i pulsanti di esportazione per generare file CSV, XER o il file finale completo

### Tipologie di File Supportate

- **Formati accettati**: .xlsx, .xls, .csv, .txt
- **Nomi fogli riconosciuti**:
  - Budget: BUDGET, Budget, budget, BUDGET CPM
  - SIL Diretti: SIL diretti, SIL Diretti, SIL_DIRETTI, SIL DIRETTI
  - SIL Indiretti: SIL indiretti, SIL Indiretti, SIL_INDIRETTI, SIL INDIRETTI
  - P6: EXPORT_P6, Export_P6, export_p6, EXPORT P6

## Funzionalità Avanzate

### Anteprima Dati
Dopo il caricamento di ciascun file, viene mostrata un'anteprima delle prime righe per verificare che i dati siano stati interpretati correttamente.

### Esportazione File Finale
Il pulsante "Genera File Finale" crea un file Excel con 5 fogli contenenti:
1. Riepilogo WBS
2. Distribuzione P6
3. Alert
4. Deviazioni
5. KPI (Indicatori chiave di prestazione)

### Identificazione Scostamenti
Lo strumento evidenzia automaticamente eventuali discrepanze tra:
- Importi SIL e costi P6
- WBS non trovate in entrambi i sistemi
- Attività con impegno finanziario ma stato fisico = 0%

## Mappatura Automatica

Il processo di bridge segue questi passaggi:

1. **Mapping Articolo → WBS** dal budget CPM
2. **Mapping WBS → Attività P6** dall'export P6
3. **Distribuzione SIL** proporzionalmente ai costi attuali delle attività P6
4. **Calcolo scostamenti** tra dati previsti e dati effettivi

## Risoluzione dei Problemi Comuni

### File non riconosciuti
- Verificare che i fogli abbiano i nomi corretti
- Controllare che le colonne abbiano intestazioni simili a quelle indicate
- Assicurarsi che i formati numerici siano corretti (virgole come separatore decimale)

### Mapping incompleto
- Controllare che le WBS siano presenti in entrambi i sistemi
- Verificare che la formattazione delle WBS sia coerente
- Confrontare i codici articolo tra budget e SIL

## Sicurezza e Privacy

I dati rimangono nel browser e non vengono mai inviati a server esterni. Tutto l'elaborazione avviene localmente sul computer dell'utente.

## Versione Corrente

- Versione: 13.1
- Sviluppato da: Cosedil PMO
- Caratteristiche principali: Mapping automatico Articolo→WBS→Activity, supporto multiplo formati, anteprima dati, esportazione completa

## Supporto e Contatti

Per supporto tecnico o richieste di funzionalità, contattare il team PMO Cosedil.