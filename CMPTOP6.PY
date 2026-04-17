#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CPM → P6 Bridge v13.1 (Python Edition)
Replica la logica del tool Cosedil PMO:
Mapping a 3 livelli: SIL Articolo → Budget WBS → P6 Activities → Distribuzione proporzionale su act_cost
"""

import pandas as pd
import logging
import sys
from pathlib import Path
import numpy as np

# Configurazione base dei log per tracciare il flusso di esecuzione
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

class CPMtoP6Bridge:
    def __init__(self, excel_path: str):
        """
        Inizializza l'istanza del Bridge.
        - Memorizza il percorso del file Excel sorgente
        - Prepara i DataFrame (inizialmente vuoti) per i 4 fogli richiesti
        - Definisce un dizionario di mapping per standardizzare i nomi delle colonne
        - Inizializza una lista per raccogliere alert/warning durante l'elaborazione
        """
        self.excel_path = Path(excel_path)
        self.df_budget = None
        self.df_sil_dir = None
        self.df_sil_ind = None
        self.df_p6 = None
        self.df_bridge = None
        self.alerts = []

        # ⚙️ CONFIGURAZIONE NOMI COLONNE
        # Adatta le chiavi ai nomi effettivi presenti nel tuo file Excel
        self.cols = {
            "budget": ["Cod. WBS", "Cod.WBS", "WBS", "Codice WBS", "articolo", "Articolo", "Descrizione Articolo", "Importo Costo (€)", "Importo Costo", "Importo Costo (euro)", "Costo", "Importo", "Budget"],
            "sil_dir": ["Cod. S.I.L.", "Cod.S.I.L.", "CodSIL", "SIL", "Codice SIL", "articolo", "Articolo", "Descrizione Articolo", "Importo", "Costo", "Valore"],
            "sil_ind": ["Cod. S.I.L.", "Cod.S.I.L.", "CodSIL", "SIL", "Codice SIL", "articolo", "Articolo", "Descrizione Articolo", "Importo", "Costo", "Valore"],
            "p6": ["task_code", "activity_id", "wbs_id", "act_cost", "act_name", "activity_name", "calc_phys_complete_pct", "activity_status", "status_code"]
        }

    def load_sheets(self, sheet_mapping=None):
        """
        Legge i 4 fogli Excel richiesti dal flusso CPM→P6.
        Verifica l'esistenza del file, carica i fogli con i nomi standard e gestisce
        errori di nomi mancanti o formattazione errata.
        """
        logging.info("📥 Lettura fogli Excel...")
        if not self.excel_path.exists():
            raise FileNotFoundError(f"File non trovato: {self.excel_path}")

        # Mapping dei nomi standard ai nomi effettivi nei file
        default_sheet_mapping = {
            "budget": ["BUDGET", "Budget", "budget", "BUDGET CPM", "Budget CPM"],
            "sil_dir": ["SIL diretti", "SIL_diretti", "SIL DIRETTI", "sil_diretti", "SIL diretti", "SIL Diretti"],
            "sil_ind": ["SIL indiretti", "SIL_indiretti", "SIL INDIRETTI", "sil_indiretti", "SIL indiretti", "SIL Indiretti"],
            "p6": ["EXPORT_P6", "Export_P6", "export_p6", "EXPORT P6", "export_p6.xlsx", "P6 Export"]
        }
        sheet_mapping = sheet_mapping or default_sheet_mapping

        try:
            # Trova i nomi reali dei fogli
            xl_file = pd.ExcelFile(self.excel_path)
            available_sheets = xl_file.sheet_names
            
            # Associa i fogli richiesti ai nomi disponibili
            sheets_to_load = {}
            for key, possible_names in sheet_mapping.items():
                found_sheet = None
                for name in possible_names:
                    # Cerca corrispondenze parziali
                    matching = [s for s in available_sheets if name.lower() in s.lower()]
                    if matching:
                        found_sheet = matching[0]  # Usa la prima corrispondenza
                        break
                if found_sheet:
                    sheets_to_load[key] = found_sheet
                    logging.info(f"Foglio trovato per {key}: {found_sheet}")
                else:
                    logging.warning(f"Nessun foglio trovato per {key} tra i nomi: {possible_names}")
                    # Potrebbe essere opzionale
                    if key != 'sil_ind':  # sil_ind è opzionale
                        sheets_to_load[key] = None

            # Carica i fogli trovati
            self.df_budget = pd.read_excel(self.excel_path, sheet_name=sheets_to_load["budget"]) if sheets_to_load["budget"] else None
            self.df_sil_dir = pd.read_excel(self.excel_path, sheet_name=sheets_to_load["sil_dir"]) if sheets_to_load["sil_dir"] else None
            self.df_sil_ind = pd.read_excel(self.excel_path, sheet_name=sheets_to_load["sil_ind"]) if sheets_to_load["sil_ind"] else None
            self.df_p6 = pd.read_excel(self.excel_path, sheet_name=sheets_to_load["p6"]) if sheets_to_load["p6"] else None

            # Verifica che i fogli richiesti esistano
            if self.df_budget is None or self.df_sil_dir is None or self.df_p6 is None:
                missing = []
                if self.df_budget is None: missing.append("Budget CPM")
                if self.df_sil_dir is None: missing.append("SIL Diretti")
                if self.df_p6 is None: missing.append("Export P6")
                raise ValueError(f"Mancano fogli essenziali: {', '.join(missing)}")

        except Exception as e:
            logging.error(f"❌ Errore durante la lettura del file Excel: {e}")
            sys.exit(1)

    def normalize_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Normalizza i nomi delle colonne rimuovendo spazi e rendendo minuscoli i nomi
        """
        df = df.copy()
        new_columns = {}
        for col in df.columns:
            normalized = str(col).lower().strip().replace(' ', '_').replace('.', '_').replace('__', '_')
            new_columns[col] = normalized
        df.rename(columns=new_columns, inplace=True)
        return df

    def find_column(self, df: pd.DataFrame, possible_names: list) -> str:
        """
        Trova una colonna in un DataFrame cercando tra una lista di possibili nomi
        """
        # Normalizza i nomi cercati
        normalized_possible = [name.lower().strip().replace(' ', '_').replace('.', '_').replace('__', '_') for name in possible_names]
        normalized_cols = [str(col).lower().strip().replace(' ', '_').replace('.', '_').replace('__', '_') for col in df.columns]
        
        for name in normalized_possible:
            # Cerca corrispondenza esatta
            if name in normalized_cols:
                original_col = df.columns[normalized_cols.index(name)]
                return original_col
            
            # Cerca corrispondenza parziale
            for idx, col in enumerate(normalized_cols):
                if name in col or col in name:
                    original_col = df.columns[idx]
                    return original_col
                    
        # Se non trovato, ritorna None
        return None

    def build_bridge(self):
        """
        Cuore del processo. Implementa il mapping a 3 livelli descritto nel tool:
        1️⃣ SIL Articolo → Budget Articolo → ottieni Cod. WBS
        2️⃣ Cod. WBS → Export P6 wbs_id → lista attività P6
        3️⃣ Distribuisce il SIL tra le attività proporzionalmente al loro act_cost
        Gestisce anche SIL indiretti, casi di mancata corrispondenza e divisioni per zero.
        """
        if self.df_budget is None or self.df_sil_dir is None or self.df_p6 is None:
            self.alerts.append("⚠️ Budget CPM, SIL Diretti o Export P6 mancanti.")
            return

        # Normalizza i nomi delle colonne
        self.df_budget = self.normalize_column_names(self.df_budget)
        self.df_sil_dir = self.normalize_column_names(self.df_sil_dir)
        if self.df_sil_ind is not None:
            self.df_sil_ind = self.normalize_column_names(self.df_sil_ind)
        self.df_p6 = self.normalize_column_names(self.df_p6)

        # Trova le colonne necessarie
        budget_wbs_col = self.find_column(self.df_budget, ["Cod._wbs", "wbs", "codice_wbs", "cod_wbs"])
        budget_art_col = self.find_column(self.df_budget, ["articolo", "descrizione_articolo", "descrizione"])
        budget_cost_col = self.find_column(self.df_budget, ["importo_costo_€", "importo_costo", "importo_costo_euro", "costo", "importo", "budget"])

        sil_dir_art_col = self.find_column(self.df_sil_dir, ["articolo", "descrizione_articolo", "descrizione"])
        sil_dir_val_col = self.find_column(self.df_sil_dir, ["importo", "costo", "valore"])
        sil_dir_wbs_col = self.find_column(self.df_sil_dir, ["cod._s_i_l_", "sil", "codice_sil"])

        sil_ind_art_col = self.find_column(self.df_sil_ind, ["articolo", "descrizione_articolo", "descrizione"]) if self.df_sil_ind is not None else None
        sil_ind_val_col = self.find_column(self.df_sil_ind, ["importo", "costo", "valore"]) if self.df_sil_ind is not None else None

        p6_wbs_col = self.find_column(self.df_p6, ["wbs_id", "wbs_code", "wbs"])
        p6_task_col = self.find_column(self.df_p6, ["task_code", "activity_id", "activity_id"])
        p6_cost_col = self.find_column(self.df_p6, ["act_cost", "actual_total_cost", "cost"])
        p6_name_col = self.find_column(self.df_p6, ["act_name", "activity_name", "name", "descrizione"])
        p6_phys_col = self.find_column(self.df_p6, ["calc_phys_complete_pct", "physical_%_complete", "phys_complete"])

        if not all([budget_wbs_col, budget_art_col, budget_cost_col, sil_dir_art_col, sil_dir_val_col, p6_wbs_col, p6_task_col, p6_cost_col]):
            missing_cols = []
            if not budget_wbs_col: missing_cols.append("Colonna WBS in Budget")
            if not budget_art_col: missing_cols.append("Colonna Articolo in Budget")
            if not budget_cost_col: missing_cols.append("Colonna Costo in Budget")
            if not sil_dir_art_col: missing_cols.append("Colonna Articolo in SIL Diretti")
            if not sil_dir_val_col: missing_cols.append("Colonna Valore in SIL Diretti")
            if not p6_wbs_col: missing_cols.append("Colonna WBS in P6")
            if not p6_task_col: missing_cols.append("Colonna Task in P6")
            if not p6_cost_col: missing_cols.append("Colonna Costo in P6")
            self.alerts.append(f"⚠️ Colonne mancanti: {', '.join(missing_cols)}")
            return

        logging.info("🔗 Step 1: Mapping Articolo → WBS")

        # Pulizia dei dati
        self.df_budget[budget_wbs_col] = self.df_budget[budget_wbs_col].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x)
        self.df_budget[budget_art_col] = self.df_budget[budget_art_col].astype(str).apply(lambda x: x.lower().strip())

        # Unisce SIL diretti e Budget tramite la colonna comune "Articolo"
        merged_sil = pd.merge(
            self.df_sil_dir[[sil_dir_art_col, sil_dir_val_col]].rename(columns={sil_dir_art_col: 'articolo', sil_dir_val_col: 'sil_valore'}),
            self.df_budget[[budget_art_col, budget_wbs_col]].rename(columns={budget_art_col: 'articolo', budget_wbs_col: 'wbs'}),
            on='articolo',
            how='left'
        )

        # Controlla articoli SIL senza corrispondenza nel Budget
        unmapped = merged_sil[merged_sil['wbs'].isna() | (merged_sil['wbs'] == 'nan')]
        if not unmapped.empty:
            self.alerts.append(f"🔴 {len(unmapped)} articoli SIL diretti non mappati sul Budget.")
            logging.warning(f"Articoli non mappati: {unmapped['articolo'].tolist()}")
        merged_sil = merged_sil[(~merged_sil['wbs'].isna()) & (merged_sil['wbs'] != 'nan')]

        # Gestione SIL Indiretti (opzionali ma sommati ai Diretti se presenti)
        if self.df_sil_ind is not None and sil_ind_art_col and sil_ind_val_col:
            sil_ind_data = self.df_sil_ind[[sil_ind_art_col, sil_ind_val_col]].rename(columns={
                sil_ind_art_col: 'articolo', 
                sil_ind_val_col: 'sil_valore'
            })
            merged_ind = pd.merge(
                sil_ind_data,
                self.df_budget[[budget_art_col, budget_wbs_col]].rename(columns={budget_art_col: 'articolo', budget_wbs_col: 'wbs'}),
                on='articolo',
                how='left'
            )
            merged_ind = merged_ind[(~merged_ind['wbs'].isna()) & (merged_ind['wbs'] != 'nan')]

            # Concatena Diretti e Indiretti e somma i valori monetari per ogni WBS
            combined_sil = pd.concat([
                merged_sil[['wbs', 'sil_valore']].groupby('wbs').sum().reset_index(),
                merged_ind[['wbs', 'sil_valore']].groupby('wbs').sum().reset_index()
            ], ignore_index=True).groupby('wbs').sum().reset_index()
        else:
            # Se non ci sono indiretti, usa solo i diretti
            combined_sil = merged_sil[['wbs', 'sil_valore']].groupby('wbs').sum().reset_index()

        # Rinomina la colonna per consistenza
        combined_sil.rename(columns={'sil_valore': 'sil_tot'}, inplace=True)

        logging.info("🔗 Step 2: Mapping WBS → Attività P6")
        p6_data = self.df_p6.copy()
        p6_data[p6_cost_col] = pd.to_numeric(p6_data[p6_cost_col], errors='coerce').fillna(0)

        # Rinomina colonne per consistenza
        p6_renamed = p6_data.rename(columns={
            p6_wbs_col: 'wbs_id',
            p6_task_col: 'act_id',
            p6_cost_col: 'act_cost',
            p6_name_col: 'act_name' if p6_name_col else 'act_name',
            p6_phys_col: 'phys_complete_pct' if p6_phys_col else 'phys_complete_pct'
        })

        # Aggiungi valore di default per colonne mancanti
        if 'act_name' not in p6_renamed.columns:
            p6_renamed['act_name'] = ''
        if 'phys_complete_pct' not in p6_renamed.columns:
            p6_renamed['phys_complete_pct'] = 0

        # Esegue il merge
        bridge = pd.merge(combined_sil, p6_renamed, left_on='wbs', right_on='wbs_id', how='left')

        # Controlla WBS presenti nel Budget/SIL ma assenti nell'export P6
        unmapped_wbs = bridge[bridge['act_id'].isna()]
        if not unmapped_wbs.empty:
            self.alerts.append(f"🔴 {unmapped_wbs['wbs'].nunique()} WBS non trovate in Export P6.")
            logging.warning(f"WBS non trovate in P6: {unmapped_wbs['wbs'].unique()}")
        bridge = bridge[~bridge['act_id'].isna()]

        logging.info("🔗 Step 3: Distribuzione proporzionale SIL su act_cost")
        bridge['act_cost'] = pd.to_numeric(bridge['act_cost'], errors='coerce').fillna(0)

        # Calcola il costo totale per ogni WBS (somma di tutte le attività associate)
        bridge['wbs_total_cost'] = bridge.groupby('wbs')['act_cost'].transform('sum')

        # Calcola la percentuale di ripartizione per ogni attività
        # Gestione divisione per zero: se il WBS ha costo 0, la ratio diventa uniforme
        bridge['ratio'] = bridge.apply(
            lambda r: r['act_cost'] / r['wbs_total_cost'] if r['wbs_total_cost'] > 0 else 1.0 / len(bridge[bridge['wbs'] == r['wbs']]),
            axis=1
        )

        # Assegna il SIL proporzionale a ciascuna attività
        bridge['allocated_sil'] = bridge['sil_tot'] * bridge['ratio']

        # Calcola la differenza tra SIL allocato e costo P6
        bridge['delta'] = bridge['allocated_sil'] - bridge['act_cost']

        # Seleziona, ordina e pulisce il DataFrame finale
        self.df_bridge = bridge[[
            'wbs', 'act_id', 'act_name', 'act_cost', 'sil_tot', 'ratio', 'allocated_sil', 'delta'
        ]].copy().sort_values(['wbs', 'act_id']).reset_index(drop=True)

        logging.info(f"✅ Bridge completato: {len(self.df_bridge)} record creati")

    def show_status(self):
        """
        Stampa un riepilogo finale dell'elaborazione.
        Se non sono stati raccolti alert, conferma il successo.
        Altrimenti, elenca tutti i warnings/errori per facilitare il debug.
        """
        if not self.alerts:
            logging.info("✅ Bridge completo — Tutti gli articoli SIL mappati su attività P6.")
        else:
            for a in self.alerts:
                logging.warning(a)

    def export(self, output_path: str = "CPM_P6_Bridge_v13.1.xlsx"):
        """
        Esporta il DataFrame risultato in un file Excel.
        Verifica che ci siano dati da salvare prima di scrivere su disco.
        """
        if self.df_bridge is None or self.df_bridge.empty:
            logging.error("❌ Nessun dato da esportare.")
            return

        # Assicura che la cartella di destinazione esista
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            self.df_bridge.to_excel(writer, sheet_name='Bridge_Result', index=False)
            
            # Aggiungi foglio di riepilogo
            summary_data = {
                'Metrica': [
                    'Totale Attività Mappate',
                    'Totale WBS Coinvolte',
                    'SIL Totale Allocato (€)',
                    'Costo P6 Totale (€)',
                    'Delta Totale (€)',
                    'CPI (SIL/P6)'
                ],
                'Valore': [
                    len(self.df_bridge),
                    self.df_bridge['wbs'].nunique(),
                    round(self.df_bridge['allocated_sil'].sum(), 2),
                    round(self.df_bridge['act_cost'].sum(), 2),
                    round(self.df_bridge['delta'].sum(), 2),
                    round(self.df_bridge['allocated_sil'].sum() / self.df_bridge['act_cost'].sum(), 3) if self.df_bridge['act_cost'].sum() > 0 else 'N/A'
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Riepilogo', index=False)

        logging.info(f"📁 Export salvato: {output_path}")

    def run(self, output_path: str = "CPM_P6_Bridge_v13.1.xlsx"):
        """
        Metodo orchestratore: esegue il flusso completo in sequenza.
        Garantisce un ordine logico: Caricamento → Elaborazione → Report → Export.
        """
        self.load_sheets()
        self.build_bridge()
        self.show_status()
        self.export(output_path)


def main():
    if len(sys.argv) < 2:
        print("Utilizzo: python CMPTOP6.PY <percorso_file_excel> [percorso_output_opzionale]")
        print("Esempio: python CMPTOP6.PY ./dati.xlsx ./output/bridge_result.xlsx")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "CPM_P6_Bridge_v13.1.xlsx"
    
    bridge = CPMtoP6Bridge(input_path)
    bridge.run(output_path=output_path)


if __name__ == "__main__":
    main()