#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CPM → P6 Bridge v13.1 (Python Edition)
Logica: SIL Articolo → Budget WBS → P6 Activities → Distribuzione proporzionale su act_cost
"""

import pandas as pd
import numpy as np
import logging
import sys
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

class CPMtoP6Bridge:
    def __init__(self, excel_path: str):
        self.excel_path = Path(excel_path)
        self.df_budget = None
        self.df_sil_dir = None
        self.df_sil_ind = None
        self.df_p6 = None
        self.df_bridge = None
        self.alerts = []

        # ⚙️ CONFIGURAZIONE NOMI COLONNE (adattali al tuo Excel)
        self.cols = {
            "budget": {"articolo": "Articolo", "wbs": "Cod. WBS"},
            "sil_dir": {"articolo": "Articolo", "valore": "Importo"},
            "sil_ind": {"articolo": "Articolo", "valore": "Importo"},
            "p6": {"wbs_id": "wbs_id", "act_id": "activity_id", "act_cost": "act_cost", "act_name": "activity_name"}
        }

    def load_sheets(self):
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
        df = df.copy()
        for std_name, excel_name in mapping.items():
            if excel_name in df.columns:
                df.rename(columns={excel_name: f"{prefix}_{std_name}"}, inplace=True)
        return df

    def build_bridge(self):
        if self.df_budget is None or self.df_sil_dir is None:
            self.alerts.append("⚠️ Budget CPM o SIL Diretti mancanti.")
            return

        # 🔹 Step 1: SIL → Budget → WBS
        logging.info("🔗 Step 1: Mapping Articolo → WBS")
        b = self.standardize_columns(self.df_budget, self.cols["budget"], "b")
        sd = self.standardize_columns(self.df_sil_dir, self.cols["sil_dir"], "sd")
        
        merged_sil = pd.merge(sd, b[["b_articolo", "b_wbs"]], left_on="sd_articolo", right_on="b_articolo", how="left")
        
        unmapped = merged_sil[merged_sil["b_wbs"].isna()]
        if not unmapped.empty:
            self.alerts.append(f"🔴 {len(unmapped)} articoli SIL diretti non mappati sul Budget.")
        merged_sil = merged_sil.dropna(subset=["b_wbs"])

        # Unisci SIL Indiretti (se presenti)
        if self.df_sil_ind is not None:
            si = self.standardize_columns(self.df_sil_ind, self.cols["sil_ind"], "si")
            merged_ind = pd.merge(si, b[["b_articolo", "b_wbs"]], left_on="si_articolo", right_on="b_articolo", how="left")
            merged_ind = merged_ind.dropna(subset=["b_wbs"])
            # Somma SIL Diretti + Indiretti per stesso articolo/WBS
            merged_sil = pd.concat([
                merged_sil[["b_wbs", "sd_valore"]].rename(columns={"sd_valore": "sil_tot"}),
                merged_ind[["b_wbs", "si_valore"]].rename(columns={"si_valore": "sil_tot"})
            ], ignore_index=True).groupby("b_wbs").sum().reset_index()
        else:
            merged_sil = merged_sil[["b_wbs", "sd_valore"]].rename(columns={"sd_valore": "sil_tot"})

        # 🔹 Step 2: WBS → P6 Activities
        logging.info("🔗 Step 2: Mapping WBS → Attività P6")
        p6 = self.standardize_columns(self.df_p6, self.cols["p6"], "p6")
        bridge = pd.merge(merged_sil, p6, left_on="b_wbs", right_on="p6_wbs_id", how="left")
        
        unmapped_wbs = bridge[bridge["p6_act_id"].isna()]
        if not unmapped_wbs.empty:
            self.alerts.append(f"🔴 {unmapped_wbs['b_wbs'].nunique()} WBS non trovate in Export P6.")
        bridge = bridge.dropna(subset=["p6_act_id"])

        # 🔹 Step 3: Distribuzione proporzionale su act_cost
        logging.info("🔗 Step 3: Distribuzione proporzionale SIL su act_cost")
        bridge["p6_act_cost"] = pd.to_numeric(bridge["p6_act_cost"], errors="coerce").fillna(0)
        bridge["wbs_total_cost"] = bridge.groupby("b_wbs")["p6_act_cost"].transform("sum")
        
        # Evita divisione per zero
        bridge["ratio"] = bridge.apply(lambda r: r["p6_act_cost"] / r["wbs_total_cost"] if r["wbs_total_cost"] > 0 else 0, axis=1)
        bridge["allocated_sil"] = bridge["sil_tot"] * bridge["ratio"]

        self.df_bridge = bridge[[
            "b_wbs", "p6_act_id", "p6_act_name", "p6_act_cost", 
            "sil_tot", "ratio", "allocated_sil"
        ]].copy().sort_values(["b_wbs", "p6_act_id"]).reset_index(drop=True)

    def show_status(self):
        if not self.alerts:
            logging.info("✅ Bridge completo — Tutti gli articoli SIL mappati su attività P6.")
        else:
            for a in self.alerts:
                logging.warning(a)

    def export(self, output_path: str = "CPM_P6_Bridge_v13.1.xlsx"):
        if self.df_bridge is None or self.df_bridge.empty:
            logging.error("❌ Nessun dato da esportare.")
            return
        self.df_bridge.to_excel(output_path, index=False)
        logging.info(f"📁 Export salvato: {output_path}")

    def run(self):
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