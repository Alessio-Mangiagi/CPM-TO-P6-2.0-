import pandas as pd
import re
from pathlib import Path
import os # Importiamo 'os' per gestire i path in modo robusto

def parse_xer(filepath: Path) -> dict[str, pd.DataFrame]:
    """
    Legge un file .xer da un percorso specificato e restituisce un dizionario 
    di DataFrame Pandas per ogni tabella identificata nel file.

    Args:
        filepath: Il percorso assoluto o relativo al file .xer da analizzare.

    Returns:
        dict[str, pd.DataFrame]: Dizionario contenente i dati strutturati.
    """
    # Verifica che il file esista prima di tentare l'apertura
    if not filepath.exists():
        raise FileNotFoundError(f"Errore: Il file non è stato trovato al percorso specificato: {filepath}")

    print(f"[*] Inizio parsing del file: {filepath}...")
    
    # Lettura del contenuto (encoding 'latin-1' mantenuto per compatibilità con i dati sorgente)
    try:
        with open(filepath, encoding='latin-1') as f:
            content = f.read()
    except Exception as e:
        print(f"Errore durante la lettura del file: {e}")
        return {} # Ritorna un dizionario vuoto in caso di fallimento

    tables = {}
    # Regex per dividere il contenuto nei blocchi basati sull'intestazione di una nuova tabella
    blocks = re.split(r'(?=^%T\t)', content, flags=re.MULTILINE)
    
    for i, block in enumerate(blocks):
        block = block.strip()
        if not block or not block.startswith('%T\t'):
            continue

        lines = block.split('\n')
        
        # 1. Identificazione del nome della tabella
        try:
            table_name = lines[0].split('\t')[1].strip()
        except IndexError:
            print(f"Attenzione: Blocco {i+1} non contiene un'intestazione valida e viene saltato.")
            continue

        fields = None
        rows = []
        
        # 2. Estrazione dei dati (Campi e Righe)
        for line in lines[1:]:
            line = line.strip()
            if not line:
                continue
                
            if line.startswith('%F\t'):
                fields = line[3:].rstrip().split('\t')
            elif line.startswith('%R\t'):
                vals = line[3:].rstrip().split('\t')
                rows.append(vals)
        
        # 3. Creazione del DataFrame solo se sono stati trovati campi e dati
        if fields and rows:
            print(f"   -> Tabella rilevata: '{table_name}' (Righe: {len(rows)}, Colonne: {len(fields)})")

            # Normalizzazione lunghezza righe (Assicura che ogni riga abbia tutte le colonne)
            num_fields = len(fields)
            
            processed_rows = []
            for r in rows:
                # Riempie o trunca la riga per adattarla esattamente ai campi attesi
                normalized_row = (r + [''] * max(0, num_fields - len(r)))[:num_fields]
                processed_rows.append(normalized_row)

            tables[table_name] = pd.DataFrame(processed_rows, columns=fields)
        else:
            # print(f"   -> Blocco '{table_name}' saltato per mancanza di dati o campi.")
            pass # Silenzioso se non ci sono dati validi da estrarre

    return tables


# ========================================================
# 📂 USO DEL PROGRAMMA (DEVI SOLO MODIFICARE IL PERCORSO QUI)
# ========================================================

# *** AGGIUNGI O MODIFICA QUI SOLO LA VARIABILE FILEPATH ***
FILEPATH_PER_ANALISI = Path(r'3123 CEMEX 2026 -15.04.2026.xer') 


try:
    # Esecuzione della funzione passando il percorso definito sopra
    xer = parse_xer(filepath=FILEPATH_PER_ANALISI)

    if xer:
        print("\n✅ Parsing completato con successo.")
        print("Le tabelle caricate nel dizionario 'xer' sono:")
        # Stampa i nomi delle tabelle per confermare l'operatività
        print(list(xer.keys())) 
        
        # Esempio di utilizzo: recuperare la tabella TASK e vederne le prime righe
        task = xer.get('TASK')
        if task is not None:
            print("\n--- Preview della tabella 'TASK' (le 5 prime righe) ---")
            print(task.head())
        else:
             print("\nAttenzione: La tabella 'TASK' non è stata trovata nel file.")


except FileNotFoundError as e:
    print(f"\n❌ ERRORE FATALE: {e}")
except Exception as e:
    print(f"\n❌ Si è verificato un errore imprevisto durante l'esecuzione: {e}")

