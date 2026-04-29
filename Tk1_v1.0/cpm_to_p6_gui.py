#!/usr/bin/env python3
"""
CPM→P6 Bridge — Interfaccia grafica  ·  COSEDIL S.p.A. · PMO
Avvio: python cpm_to_p6_gui.py
"""

import io
import os
import sys
import csv
import shutil
import threading
import subprocess
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

# ── Import logica core ────────────────────────────────────────────────────────
_HERE = Path(__file__).parent
sys.path.insert(0, str(_HERE))
import cpm_to_p6 as core

# ── Tema e aspetto ────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT   = "#1f6aa5"
OK_COLOR = "#2fa84f"
ERR_COLOR = "#c0392b"
WARN_COLOR = "#e67e22"

DEFAULT_FILE = core.DEFAULT_FILE


# ══════════════════════════════════════════════════════════════════════════════
#  Redirect stdout → CTkTextbox
# ══════════════════════════════════════════════════════════════════════════════
class TextboxWriter(io.TextIOBase):
    """Scrive su un CTkTextbox catturando print() durante le operazioni."""
    def __init__(self, textbox: ctk.CTkTextbox):
        self._box = textbox

    def write(self, text: str) -> int:
        if text:
            self._box.configure(state="normal")
            self._box.insert("end", text)
            self._box.see("end")
            self._box.configure(state="disabled")
            self._box.update_idletasks()
        return len(text)

    def flush(self):
        pass


# ══════════════════════════════════════════════════════════════════════════════
#  Widget: card con titolo
# ══════════════════════════════════════════════════════════════════════════════
class Card(ctk.CTkFrame):
    def __init__(self, master, title: str, **kw):
        super().__init__(master, corner_radius=10, **kw)
        ctk.CTkLabel(
            self, text=title, font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray40", "gray70"),
        ).pack(anchor="w", padx=12, pady=(8, 2))


# ══════════════════════════════════════════════════════════════════════════════
#  Finestra principale
# ══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CPM→P6 Bridge  ·  COSEDIL S.p.A.")
        self.geometry("1100x680")
        self.minsize(900, 560)

        self._busy = False           # blocca doppio-click sui pulsanti
        self._sil_max = 0
        self._last_csv: Path | None = None

        self._build_ui()
        self._apply_file(str(DEFAULT_FILE))

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── Pannello sinistro ─────────────────────────────────────────────────
        left = ctk.CTkFrame(self, width=260, corner_radius=0)
        left.grid(row=0, column=0, sticky="nsew")
        left.grid_rowconfigure(5, weight=1)
        left.grid_propagate(False)

        # Logo / titolo
        ctk.CTkLabel(
            left, text="CPM → P6",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 2), sticky="w")
        ctk.CTkLabel(
            left, text="COSEDIL S.p.A.  ·  PMO",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")

        # File Bridge
        file_card = Card(left, "FILE BRIDGE")
        file_card.grid(row=2, column=0, padx=12, pady=(0, 8), sticky="ew")

        self._file_var = ctk.StringVar()
        file_row = ctk.CTkFrame(file_card, fg_color="transparent")
        file_row.pack(fill="x", padx=8, pady=(0, 8))
        self._file_entry = ctk.CTkEntry(
            file_row, textvariable=self._file_var,
            placeholder_text="Percorso file .xlsx",
            font=ctk.CTkFont(size=11),
        )
        self._file_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ctk.CTkButton(
            file_row, text="...", width=36,
            command=self._browse_file,
        ).pack(side="left")

        # SIL corrente
        sil_card = Card(left, "SIL CORRENTE")
        sil_card.grid(row=3, column=0, padx=12, pady=(0, 8), sticky="ew")

        sil_row = ctk.CTkFrame(sil_card, fg_color="transparent")
        sil_row.pack(fill="x", padx=8, pady=(0, 10))

        ctk.CTkButton(sil_row, text="−", width=32, command=self._sil_dec).pack(side="left")
        self._sil_var = ctk.StringVar(value="15")
        ctk.CTkEntry(
            sil_row, textvariable=self._sil_var,
            width=60, justify="center",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=6)
        ctk.CTkButton(sil_row, text="+", width=32, command=self._sil_inc).pack(side="left")

        self._sil_info = ctk.CTkLabel(
            sil_card, text="max: —",
            font=ctk.CTkFont(size=10), text_color=("gray50", "gray60"),
        )
        self._sil_info.pack(anchor="w", padx=14, pady=(0, 6))

        # Azioni
        action_card = Card(left, "AZIONI")
        action_card.grid(row=4, column=0, padx=12, pady=(0, 8), sticky="ew")

        btn_cfg = dict(
            corner_radius=8, height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self._btn_analizza = ctk.CTkButton(
            action_card, text="🔍  Analizza", fg_color=ACCENT,
            command=self._run_analizza, **btn_cfg,
        )
        self._btn_analizza.pack(fill="x", padx=10, pady=(4, 4))

        self._btn_rigenera = ctk.CTkButton(
            action_card, text="⚡  Rigenera Bridge",
            command=self._run_rigenera, **btn_cfg,
        )
        self._btn_rigenera.pack(fill="x", padx=10, pady=(0, 4))

        self._btn_export = ctk.CTkButton(
            action_card, text="💾  Export CSV per P6",
            fg_color="#2d6a4f",
            command=self._run_export, **btn_cfg,
        )
        self._btn_export.pack(fill="x", padx=10, pady=(0, 8))

        # Stato / KPI
        self._status_card = Card(left, "STATO ULTIMA ESECUZIONE")
        self._status_card.grid(row=5, column=0, padx=12, pady=(0, 12), sticky="sew")

        self._kpi_check   = self._kpi_label("—", self._status_card)
        self._kpi_periodo = self._kpi_label("—", self._status_card)
        self._kpi_cumul   = self._kpi_label("—", self._status_card)

        # Tema
        ctk.CTkLabel(left, text="Tema:", font=ctk.CTkFont(size=11)).grid(
            row=6, column=0, padx=16, pady=(0, 4), sticky="w"
        )
        self._theme_switch = ctk.CTkSwitch(
            left, text="Dark mode",
            command=self._toggle_theme, onvalue="dark", offvalue="light",
        )
        self._theme_switch.select()
        self._theme_switch.grid(row=7, column=0, padx=16, pady=(0, 14), sticky="w")

        # ── Pannello destro ───────────────────────────────────────────────────
        right = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        # Barra superiore destra
        top_bar = ctk.CTkFrame(right, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        ctk.CTkLabel(
            top_bar, text="Output", font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left")

        ctk.CTkButton(
            top_bar, text="Pulisci", width=80, height=28,
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray40"),
            text_color=("gray10", "gray90"),
            hover_color=("gray80", "gray30"),
            command=self._clear_log,
        ).pack(side="right", padx=(4, 0))

        self._btn_open_folder = ctk.CTkButton(
            top_bar, text="Apri cartella output", width=150, height=28,
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray40"),
            text_color=("gray10", "gray90"),
            hover_color=("gray80", "gray30"),
            command=self._open_output_folder,
            state="disabled",
        )
        self._btn_open_folder.pack(side="right", padx=(0, 4))

        self._btn_open_csv = ctk.CTkButton(
            top_bar, text="Apri CSV", width=80, height=28,
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray40"),
            text_color=("gray10", "gray90"),
            hover_color=("gray80", "gray30"),
            command=self._open_csv,
            state="disabled",
        )
        self._btn_open_csv.pack(side="right", padx=(0, 4))

        # Log textbox
        self._log = ctk.CTkTextbox(
            right, font=ctk.CTkFont(family="Consolas", size=12),
            state="disabled", wrap="none",
        )
        self._log.grid(row=1, column=0, sticky="nsew")

        # Barra di stato
        self._statusbar = ctk.CTkLabel(
            right, text="Pronto.", anchor="w",
            font=ctk.CTkFont(size=11), text_color=("gray50", "gray60"),
        )
        self._statusbar.grid(row=2, column=0, sticky="ew", pady=(4, 0))

        # Progress bar
        self._progress = ctk.CTkProgressBar(right, mode="indeterminate")
        self._progress.grid(row=3, column=0, sticky="ew", pady=(2, 0))
        self._progress.grid_remove()

    def _kpi_label(self, text: str, parent) -> ctk.CTkLabel:
        lbl = ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=12), anchor="w")
        lbl.pack(anchor="w", padx=14, pady=1)
        return lbl

    # ── Helpers UI ────────────────────────────────────────────────────────────
    def _set_busy(self, busy: bool):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for btn in (self._btn_analizza, self._btn_rigenera, self._btn_export):
            btn.configure(state=state)
        if busy:
            self._progress.grid()
            self._progress.start()
        else:
            self._progress.stop()
            self._progress.grid_remove()

    def _set_status(self, msg: str, color: str = ("gray50", "gray60")):
        self._statusbar.configure(text=msg, text_color=color)

    def _log_write(self, text: str):
        self._log.configure(state="normal")
        self._log.insert("end", text)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _apply_file(self, path: str):
        self._file_var.set(path)

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Seleziona il file Bridge Excel",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            initialdir=str(_HERE / "file"),
        )
        if path:
            self._apply_file(path)

    def _sil_dec(self):
        try:
            v = int(self._sil_var.get())
            if v > 1:
                self._sil_var.set(str(v - 1))
        except ValueError:
            pass

    def _sil_inc(self):
        try:
            v = int(self._sil_var.get())
            self._sil_var.set(str(v + 1))
        except ValueError:
            pass

    def _toggle_theme(self):
        mode = self._theme_switch.get()
        ctk.set_appearance_mode(mode)

    def _open_output_folder(self):
        p = Path(self._file_var.get()).parent
        if p.exists():
            os.startfile(str(p))

    def _open_csv(self):
        if self._last_csv and self._last_csv.exists():
            os.startfile(str(self._last_csv))

    # ── Preparazione esecuzione ───────────────────────────────────────────────
    def _get_params(self):
        """Valida i parametri e restituisce (wb_path, sil_corrente) o None."""
        path_str = self._file_var.get().strip()
        if not path_str:
            messagebox.showerror("Errore", "Nessun file selezionato.")
            return None
        wb_path = Path(path_str)
        if not wb_path.exists():
            messagebox.showerror("Errore", f"File non trovato:\n{wb_path}")
            return None
        try:
            sil = int(self._sil_var.get())
            if sil < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Errore", "SIL corrente deve essere un numero ≥ 1.")
            return None
        return wb_path, sil

    def _run_in_thread(self, fn, *args):
        """Esegue fn(*args) in un thread, catturando stdout nel log."""
        if self._busy:
            return
        self._set_busy(True)
        self._clear_log()

        writer = TextboxWriter(self._log)

        def worker():
            old_stdout = sys.stdout
            sys.stdout = writer
            try:
                fn(*args)
            except Exception as exc:
                print(f"\n[ERRORE] {exc}")
                self.after(0, lambda: self._set_status(f"Errore: {exc}", ERR_COLOR))
            finally:
                sys.stdout = old_stdout
                self.after(0, lambda: self._set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    # ── Comandi ───────────────────────────────────────────────────────────────
    def _run_analizza(self):
        params = self._get_params()
        if not params:
            return
        wb_path, sil = params
        self._run_in_thread(self._do_analizza, wb_path, sil)

    def _do_analizza(self, wb_path: Path, sil: int):
        self.after(0, lambda: self._set_status("Analisi in corso…"))
        wbs_to_acts, wbs_to_des, wbs_to_tipo, art_to_cost, act_info, sil_records, sil_max = \
            core._load_all(wb_path)

        self._sil_max = sil_max
        self.after(0, lambda: self._sil_info.configure(text=f"max disponibile: {sil_max}"))

        sil_cur = min(sil, sil_max)
        self.after(0, lambda: self._sil_var.set(str(sil_cur)))

        print(f"  SIL corrente: {sil_cur}")
        print("  Generazione BRIDGE_SIL in memoria …")
        bridge = core.generate_bridge(
            sil_records, wbs_to_acts, wbs_to_des, wbs_to_tipo, act_info, art_to_cost
        )
        print(f"  → {len(bridge)} righe BRIDGE_SIL generate\n")

        core.show_riepilogo(sil_records, bridge)
        core.show_alert(bridge, act_info, sil_cur)

        # KPI sidebar
        tot_per  = sum(r[9] for r in bridge if int(r[0]) == sil_cur)
        tot_cum  = sum(r[9] for r in bridge)
        quadra   = all(
            abs(sum(r[9] for r in bridge if str(r[2]) == w) -
                sum(rec["imp"] for rec in sil_records if rec["wbs"] == w)) <= 1.0
            for w in {str(r[2]) for r in bridge}
        )
        chk = "✓ QUADRA" if quadra else "⚠ NON QUADRA"
        self.after(0, lambda: self._update_kpi(chk, tot_per, tot_cum, sil_cur, quadra))
        self.after(0, lambda: self._set_status(
            f"Analisi completata — SIL {sil_cur}  |  Tot cumulativo: €{tot_cum:,.0f}",
            OK_COLOR if quadra else WARN_COLOR,
        ))

    def _run_rigenera(self):
        params = self._get_params()
        if not params:
            return
        wb_path, sil = params
        if not messagebox.askyesno(
            "Conferma",
            f"Stai per SOVRASCRIVERE i fogli BRIDGE_SIL e P6_IMPORT_PULITO nel file:\n\n"
            f"{wb_path.name}\n\n"
            f"Verrà creato un backup automatico.\nContinuare?",
        ):
            return
        self._run_in_thread(self._do_rigenera, wb_path, sil)

    def _do_rigenera(self, wb_path: Path, sil: int):
        self.after(0, lambda: self._set_status("Rigenerazione in corso…"))
        from openpyxl import load_workbook as lw

        wbs_to_acts, wbs_to_des, wbs_to_tipo, art_to_cost, act_info, sil_records, sil_max = \
            core._load_all(wb_path)

        self._sil_max = sil_max
        self.after(0, lambda: self._sil_info.configure(text=f"max disponibile: {sil_max}"))

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

        print(f"  Apertura Excel in scrittura …")
        wb = lw(str(wb_path))
        core._write_bridge(wb["BRIDGE_SIL"], bridge)
        core._write_p6_import(wb["P6_IMPORT_PULITO"], p6_import, sil_cur, sil_max)
        wb.save(str(wb_path))
        print(f"  ✓ Salvato: {wb_path.name}\n")

        core.show_riepilogo(sil_records, bridge)
        core.show_alert(bridge, act_info, sil_cur)

        tot_per = sum(r[9] for r in bridge if int(r[0]) == sil_cur)
        tot_cum = sum(r[9] for r in bridge)
        quadra = True
        self.after(0, lambda: self._update_kpi("✓ Rigenera OK", tot_per, tot_cum, sil_cur, quadra))
        self.after(0, lambda: self._set_status(
            f"Bridge rigenerato — {wb_path.name}", OK_COLOR
        ))
        self.after(0, lambda: self._btn_open_folder.configure(state="normal"))

    def _run_export(self):
        params = self._get_params()
        if not params:
            return
        wb_path, sil = params

        out_path = filedialog.asksaveasfilename(
            title="Salva CSV per P6",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialdir=str(wb_path.parent),
            initialfile=f"P6_IMPORT_SIL{sil:02d}.csv",
        )
        if not out_path:
            return
        self._last_csv = Path(out_path)
        self._run_in_thread(self._do_export, wb_path, sil, Path(out_path))

    def _do_export(self, wb_path: Path, sil: int, out_path: Path):
        self.after(0, lambda: self._set_status("Export in corso…"))
        wbs_to_acts, wbs_to_des, wbs_to_tipo, art_to_cost, act_info, sil_records, sil_max = \
            core._load_all(wb_path)

        self._sil_max = sil_max
        self.after(0, lambda: self._sil_info.configure(text=f"max disponibile: {sil_max}"))

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

        self.after(0, lambda: self._update_kpi(
            f"✓ CSV SIL {sil_cur}", tot_per, tot_cum, sil_cur, True
        ))
        self.after(0, lambda: self._set_status(f"CSV esportato → {out_path.name}", OK_COLOR))
        self.after(0, lambda: self._btn_open_csv.configure(state="normal"))
        self.after(0, lambda: self._btn_open_folder.configure(state="normal"))

    # ── Aggiorna KPI sidebar ──────────────────────────────────────────────────
    def _update_kpi(self, check: str, tot_per: float, tot_cum: float, sil_cur: int, ok: bool):
        color = OK_COLOR if ok else WARN_COLOR
        self._kpi_check.configure(
            text=check, text_color=color,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self._kpi_periodo.configure(
            text=f"Periodo SIL {sil_cur}:  €{tot_per:>13,.0f}"
        )
        self._kpi_cumul.configure(
            text=f"Cumulativo 1→{sil_cur}:  €{tot_cum:>13,.0f}"
        )


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
