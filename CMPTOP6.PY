import pandas as pd
from typing import Dict, Any, Tuple, Optional

def find_and_map_p6_headers(df: pd.DataFrame) -> Tuple[Optional[pd.DataFrame], Dict[str, int]]:
    """
    Analizza un DataFrame (presumibilmente il contenuto grezzo di P6) 
    per trovare la riga header corretta e mappare i nomi delle colonne.

    Args:
        df: Il DataFrame che contiene le prime righe del file P6 (massimo ~15 righe).

    Returns:
        Una tupla contenente il DataFrame con gli header correttamente rinominati 
        e un dizionario degli indici delle colonne chiave.
    """

    rows = df.values.tolist()
    h_idx = -1  # Indice della riga che contiene l'header effettivo
    
    # I nomi tecnici standardizzati richiesti dall'analisi
    REQUIRED_MAP = {
        'task_code': ['task code', 'activity id'], 
        'wbs_id': ['wbs code', 'wbs id'], 
        'act_cost': ['actual total cost', 'costo totale effettivo']
        # Aggiungere altri mapping se necessario
    }

    # --- Fase 1: Ricerca della riga header (Simulazione del for loop JS) ---
    for i in range(min(15, len(rows))):
        row_str = ' | '.join([str(c).lower() for c in rows[i]])
        
        # Test 1: Cerca la combinazione ideale (task_code e wbs_id)
        if 'task_code' in row_str and 'wbs_id' in row_str:
            h_idx = i
            break

    # Fallback 2: Cerca le keywords meno precise ("activity id" e "wbs code")
    if h_idx == -1:
        for i in range(min(15, len(rows))):
            row_str = ' | '.join([str(c).lower() for c in rows[i]])
            # Questa è la logica di fallback usata nello snippet JS
            if 'activity id' in row_str and 'wbs code' in row_str: 
                h_idx = i
                break

    if h_idx == -1:
        print("ERRORE: Header P6 non trovato.")
        return None, {}

    # --- Fase 2: Mappatura e Ridenominazione delle Colonne ---
    header_row = rows[h_idx]
    headers = [str(h).strip() for h in header_row]
    
    new_columns = [] # Lista per i nuovi nomi delle colonne
    col_mapping = {} # Dizionario che memorizza la corrispondenza: {nuovo nome: vecchio nome}

    print(f"P6 Header trovato alla riga (indice {h_idx}): {headers[:6]}...")


    # Iteriamo su ogni colonna per vedere se è un header conosciuto
    for j, old_header in enumerate(headers):
        old_header_lower = str(old_header).lower()
        mapped = False

        # Tentativo di mappare l'header corrente a una delle chiavi standard
        if 'task code' in old_header_lower or 'activity id' in old_header_lower:
            new_col_name = 'task_code'
            mapped = True
        elif 'wbs code' in old_header_lower or 'wbs id' in old_header_lower:
            new_col_name = 'wbs_id'
            mapped = True
        elif 'actual total cost' in old_header_lower or 'costo totale effettivo' in old_header_lower:
            new_col_name = 'act_cost'
            mapped = True
        else:
            # Se non è un campo chiave, lo manteniamo con il nome originale (pulito)
            new_col_name = str(old_header).strip().replace(" ", "_").lower()

        new_columns.append(new_col_name)
        col_mapping[j] = new_col_name

    # Creare un DataFrame temporaneo con gli header corretti (questa è la riga da usare per i dati)
    renamed_df = pd.DataFrame(rows[h_idx], columns=new_columns)


    # --- Fase 3: Estrazione degli Indici Chiave ---
    key_indices = {}
    try:
        key_indices['task_code'] = new_columns.index('task_code')
    except ValueError:
        print("ATTENZIONE: task_code non trovato tra gli header mappati.")
