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
            "budget": {"articolo": "Articolo", "wbs": "Cod. WBS"},
            "sil_dir": {"articolo": "Articolo", "valore": "Importo"},
            "sil_ind": {"articolo": "Articolo", "valore": "Importo"},
            "p6": {"wbs_id": "wbs_id", "act_id": "activity_id", "act_cost": "act_cost", "act_name": "activity_name"}
        }

    def load_sheets(self):
        """
        Legge i 4 fogli Excel richiesti dal flusso CPM→P6.
        Verifica l'esistenza del file, carica i fogli con i nomi standard e gestisce
        errori di nomi mancanti o formattazione errata.
        """
        logging.info("📥 Lettura fogli Excel...")
        if not self.excel_path.exists():
            raise FileNotFoundError(f"File non trovato: {self.excel_path}")

        try:
            self.df_budget = pd.read_excel(self.excel_path, sheet_name="BUDGET")
            self.df_sil_dir = pd.read_excel(self.excel_path, sheet_name="SIL diretti")
            self.df_sil_ind = pd.read_excel(self.excel_path, sheet_name="SIL indiretti")
            self.df_p6 = pd.read_excel(self.excel_path, sheet_name="EXPORT_P6")
        except ValueError as e:
            logging.error(f"❌ Foglio mancante o nome errato: {e}")
            sys.exit(1)

    def standardize_columns(self, df: pd.DataFrame, mapping: dict, prefix: str) -> pd.DataFrame:
        """
        Rinomina le colonne del DataFrame secondo uno standard interno.
        - `mapping`: dizionario {nome_standard: nome_excel}
        - `prefix`: prefisso univoco (es. "b_", "sd_") per evitare collisioni durante i merge
        Restituisce una copia del DataFrame con le colonne rinominate.
        """
        df = df.copy()
        for std_name, excel_name in mapping.items():
            if excel_name in df.columns:
                df.rename(columns={excel_name: f"{prefix}_{std_name}"}, inplace=True)
        return df

    def build_bridge(self):
        """
        Cuore del processo. Implementa il mapping a 3 livelli descritto nel tool:
        1️⃣ SIL Articolo → Budget Articolo → ottieni Cod. WBS
        2️⃣ Cod. WBS → Export P6 wbs_id → lista attività P6
        3️⃣ Distribuisce il SIL tra le attività proporzionalmente al loro act_cost
        Gestisce anche SIL indiretti, casi di mancata corrispondenza e divisioni per zero.
        """
        if self.df_budget is None or self.df_sil_dir is None:
            self.alerts.append("⚠️ Budget CPM o SIL Diretti mancanti.")
            return

        # 🔹 Step 1: Mapping SIL → Budget per recuperare il WBS di riferimento
        logging.info("🔗 Step 1: Mapping Articolo → WBS")
        b = self.standardize_columns(self.df_budget, self.cols["budget"], "b")
        sd = self.standardize_columns(self.df_sil_dir, self.cols["sil_dir"], "sd")

        # Unisce SIL diretti e Budget tramite la colonna comune "Articolo"
        merged_sil = pd.merge(sd, b[["b_articolo", "b_wbs"]], left_on="sd_articolo", right_on="b_articolo", how="left")

        # Controlla articoli SIL senza corrispondenza nel Budget
        unmapped = merged_sil[merged_sil["b_wbs"].isna()]
        if not unmapped.empty:
            self.alerts.append(f"🔴 {len(unmapped)} articoli SIL diretti non mappati sul Budget.")
        merged_sil = merged_sil.dropna(subset=["b_wbs"])

        # 🔹 Gestione SIL Indiretti (opzionali ma sommati ai Diretti se presenti)
        if self.df_sil_ind is not None:
            si = self.standardize_columns(self.df_sil_ind, self.cols["sil_ind"], "si")
            merged_ind = pd.merge(si, b[["b_articolo", "b_wbs"]], left_on="si_articolo", right_on="b_articolo", how="left")
            merged_ind = merged_ind.dropna(subset=["b_wbs"])

            # Concatena Diretti e Indiretti e somma i valori monetari per ogni WBS
            merged_sil = pd.concat([
                merged_sil[["b_wbs", "sd_valore"]].rename(columns={"sd_valore": "sil_tot"}),
                merged_ind[["b_wbs", "si_valore"]].rename(columns={"si_valore": "sil_tot"})
            ], ignore_index=True).groupby("b_wbs").sum().reset_index()
        else:
            # Se non ci sono indiretti, rinomina semplicemente la colonna dei diretti
            merged_sil = merged_sil[["b_wbs", "sd_valore"]].rename(columns={"sd_valore": "sil_tot"})

        # 🔹 Step 2: Mapping WBS → Attività P6
        logging.info("🔗 Step 2: Mapping WBS → Attività P6")
        p6 = self.standardize_columns(self.df_p6, self.cols["p6"], "p6")
        bridge = pd.merge(merged_sil, p6, left_on="b_wbs", right_on="p6_wbs_id", how="left")

        # Controlla WBS presenti nel Budget/SIL ma assenti nell'export P6
        unmapped_wbs = bridge[bridge["p6_act_id"].isna()]
        if not unmapped_wbs.empty:
            self.alerts.append(f"🔴 {unmapped_wbs['b_wbs'].nunique()} WBS non trovate in Export P6.")
        bridge = bridge.dropna(subset=["p6_act_id"])

        # 🔹 Step 3: Distribuzione proporzionale del SIL basata su act_cost
        logging.info("🔗 Step 3: Distribuzione proporzionale SIL su act_cost")
        bridge["p6_act_cost"] = pd.to_numeric(bridge["p6_act_cost"], errors="coerce").fillna(0)

        # Calcola il costo totale per ogni WBS (somma di tutte le attività associate)
        bridge["wbs_total_cost"] = bridge.groupby("b_wbs")["p6_act_cost"].transform("sum")

        # Calcola la percentuale di ripartizione per ogni attività
        # Gestione divisione per zero: se il WBS ha costo 0, la ratio diventa 0
        bridge["ratio"] = bridge.apply(
            lambda r: r["p6_act_cost"] / r["wbs_total_cost"] if r["wbs_total_cost"] > 0 else 0, axis=1
        )

        # Assegna il SIL proporzionale a ciascuna attività
        bridge["allocated_sil"] = bridge["sil_tot"] * bridge["ratio"]

        # Seleziona, ordina e pulisce il DataFrame finale
        self.df_bridge = bridge[[
            "b_wbs", "p6_act_id", "p6_act_name", "p6_act_cost",
            "sil_tot", "ratio", "allocated_sil"
        ]].copy().sort_values(["b_wbs", "p6_act_id"]).reset_index(drop=True)

    def show_status(self):
        """
        Stampa un riepilogo finale dell'elaborazione.
        Se non sono stati raccolti alert, conferma il successo.
        Altrimenti, elenca tutti i warning/errori per facilitare il debug.
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
        self.df_bridge.to_excel(output_path, index=False)
        logging.info(f"📁 Export salvato: {output_path}")

    def run(self):
        """
        Metodo orchestratore: esegue il flusso completo in sequenza.
        Garantisce un ordine logico: Caricamento → Elaborazione → Report → Export.
        """
        self.load_sheets()
        self.build_bridge()
        self.show_status()
        self.export()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Utilizzo: python cpm_to_p6_bridge.py <percorso_file_excel>")
        sys.exit(1)
    bridge = CPMtoP6Bridge(sys.argv[1])
    bridge.run()