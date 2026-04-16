import pandas as pd
import re
from pathlib import Path

filepath = Path(r'3123 CEMEX 2026 -15.04.2026.xer')
def parse_xer(filepath: Path) -> dict[str, pd.DataFrame]:
    """Legge un file .xer e restituisce un dict di DataFrame per tabella."""
    with open(filepath, encoding='latin-1') as f:
        content = f.read()
    
    tables = {}
    blocks = re.split(r'(?=^%T\t)', content, flags=re.MULTILINE)
    
    for block in blocks:
        lines = block.strip().split('\n')
        if not lines or not lines[0].startswith('%T\t'):
            continue
        table_name = lines[0].split('\t')[1].strip()
        
        fields = None
        rows = []
        for line in lines[1:]:
            if line.startswith('%F\t'):
                fields = line[3:].rstrip().split('\t')
            elif line.startswith('%R\t'):
                vals = line[3:].rstrip().split('\t')
                rows.append(vals)
        
        if fields and rows:
            # Normalizza lunghezza righe
            rows = [r + [''] * (len(fields) - len(r)) for r in rows]
            rows = [r[:len(fields)] for r in rows]
            tables[table_name] = pd.DataFrame(rows, columns=fields)
    
    return tables

# Uso
xer = parse_xer('filepath')
task = xer.get('TASK')
pred = xer.get('TASKPRED')
wbs = xer.get('PROJWBS')