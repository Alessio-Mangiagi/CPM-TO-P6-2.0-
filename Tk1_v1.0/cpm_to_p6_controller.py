"""
AppController — logica operativa separata dalla GUI.
Istanziato da App; riceve un riferimento alla finestra per schedulare
aggiornamenti UI tramite after(0, ...).
"""

import csv
import shutil
from pathlib import Path

import cpm_to_p6 as core

OK_COLOR   = "#2fa84f"
WARN_COLOR = "#e67e22"


class AppController:
    def __init__(self, app):
        self._app = app

    # ── Analisi ──────────────────────────────────────────────────────────────
    def do_analizza(self, wb_path: Path, sil: int):
        app = self._app
        app.after(0, lambda: app._set_status("Analisi in corso…"))

        wbs_to_acts, wbs_to_des, wbs_to_tipo, art_to_cost, act_info, sil_records, sil_max = \
            core._load_all(wb_path)

        app._sil_max = sil_max
        app.after(0, lambda: app._sil_info.configure(text=f"max disponibile: {sil_max}"))

        sil_cur = min(sil, sil_max)
        app.after(0, lambda: app._sil_var.set(str(sil_cur)))

        print(f"  SIL corrente: {sil_cur}")
        print("  Generazione BRIDGE_SIL in memoria …")
        bridge = core.generate_bridge(
            sil_records, wbs_to_acts, wbs_to_des, wbs_to_tipo, act_info, art_to_cost
        )
        print(f"  → {len(bridge)} righe BRIDGE_SIL generate\n")

        core.show_riepilogo(sil_records, bridge)
        core.show_alert(bridge, act_info, sil_cur)

        tot_per = sum(r[9] for r in bridge if int(r[0]) == sil_cur)
        tot_cum = sum(r[9] for r in bridge)
        quadra  = all(
            abs(sum(r[9] for r in bridge if str(r[2]) == w) -
                sum(rec["imp"] for rec in sil_records if rec["wbs"] == w)) <= 1.0
            for w in {str(r[2]) for r in bridge}
        )
        chk = "✓ QUADRA" if quadra else "⚠ NON QUADRA"
        app.after(0, lambda: app._update_kpi(chk, tot_per, tot_cum, sil_cur, quadra))
        app.after(0, lambda: app._set_status(
            f"Analisi completata — SIL {sil_cur}  |  Tot cumulativo: €{tot_cum:,.0f}",
            OK_COLOR if quadra else WARN_COLOR,
        ))

    # ── Rigenerazione Bridge ─────────────────────────────────────────────────
    def do_rigenera(self, wb_path: Path, sil: int):
        from openpyxl import load_workbook as lw
        app = self._app
        app.after(0, lambda: app._set_status("Rigenerazione in corso…"))

        wbs_to_acts, wbs_to_des, wbs_to_tipo, art_to_cost, act_info, sil_records, sil_max = \
            core._load_all(wb_path)

        app._sil_max = sil_max
        app.after(0, lambda: app._sil_info.configure(text=f"max disponibile: {sil_max}"))

        sil_cur = min(sil, sil_max)

        print("  Generazione BRIDGE_SIL …")
        bridge = core.generate_bridge(
            sil_records, wbs_to_acts, wbs_to_des, wbs_to_tipo, act_info, art_to_cost
        )
        print(f"  → {len(bridge)} righe\n")

        print("  Generazione P6_IMPORT_PULITO …")
        p6_import = core.generate_p6_import(bridge, sil_cur, act_info)
        print(f"  → {len(p6_import)} attività\n")

        backup = wb_path.with_stem(wb_path.stem + "_BACKUP")
        shutil.copy2(wb_path, backup)
        print(f"  Backup: {backup.name}")

        print("  Apertura Excel in scrittura …")
        wb = lw(str(wb_path))
        core._write_bridge(wb["BRIDGE_SIL"], bridge)
        core._write_p6_import(wb["P6_IMPORT_PULITO"], p6_import, sil_cur, sil_max)
        wb.save(str(wb_path))
        print(f"  ✓ Salvato: {wb_path.name}\n")

        core.show_riepilogo(sil_records, bridge)
        core.show_alert(bridge, act_info, sil_cur)

        tot_per = sum(r[9] for r in bridge if int(r[0]) == sil_cur)
        tot_cum = sum(r[9] for r in bridge)
        app.after(0, lambda: app._update_kpi("✓ Rigenera OK", tot_per, tot_cum, sil_cur, True))
        app.after(0, lambda: app._set_status(f"Bridge rigenerato — {wb_path.name}", OK_COLOR))
        app.after(0, lambda: app._btn_open_folder.configure(state="normal"))

    # ── Export CSV ───────────────────────────────────────────────────────────
    def do_export(self, wb_path: Path, sil: int, out_path: Path):
        app = self._app
        app.after(0, lambda: app._set_status("Export in corso…"))

        wbs_to_acts, wbs_to_des, wbs_to_tipo, art_to_cost, act_info, sil_records, sil_max = \
            core._load_all(wb_path)

        app._sil_max = sil_max
        app.after(0, lambda: app._sil_info.configure(text=f"max disponibile: {sil_max}"))

        sil_cur = min(sil, sil_max)

        print("  Generazione BRIDGE_SIL …")
        bridge = core.generate_bridge(
            sil_records, wbs_to_acts, wbs_to_des, wbs_to_tipo, act_info, art_to_cost
        )
        print("  Generazione P6_IMPORT_PULITO …")
        p6_import = core.generate_p6_import(bridge, sil_cur, act_info)

        with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(core.P6_IMPORT_HEADERS)
            writer.writerows(p6_import)

        tot_per = sum(r[2] for r in p6_import)
        tot_cum = sum(r[3] for r in p6_import)

        print(f"\n  ✓ CSV esportato ({len(p6_import)} righe):")
        print(f"    {out_path}")
        print(f"\n  Periodo  SIL {sil_cur}:      € {tot_per:>15,.2f}")
        print(f"  Cumulativo SIL 1→{sil_cur}:  € {tot_cum:>15,.2f}")

        app.after(0, lambda: app._update_kpi(
            f"✓ CSV SIL {sil_cur}", tot_per, tot_cum, sil_cur, True
        ))
        app.after(0, lambda: app._set_status(f"CSV esportato → {out_path.name}", OK_COLOR))
        app.after(0, lambda: app._btn_open_csv.configure(state="normal"))
        app.after(0, lambda: app._btn_open_folder.configure(state="normal"))
