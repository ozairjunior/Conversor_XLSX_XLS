import os
import sys
import math
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import xlwt

# ── Paleta de cores ──────────────────────────────────────────────────────────
BG         = "#1C1C1E"
BG_CARD    = "#2C2C2E"
BG_INPUT   = "#3A3A3C"
ACCENT     = "#30D158"   # verde sistema
ACCENT_DIM = "#1A7A32"
WARN       = "#FF9F0A"
ERR        = "#FF453A"
INFO       = "#0A84FF"
TEXT       = "#F2F2F7"
TEXT_DIM   = "#8E8E93"
BORDER     = "#48484A"

FONT_UI    = ("SF Pro Display", 10) if sys.platform == "darwin" else ("Segoe UI", 10)
FONT_MONO  = ("SF Mono", 9)        if sys.platform == "darwin" else ("Consolas", 9)
FONT_H1    = ("SF Pro Display", 16, "bold") if sys.platform == "darwin" else ("Segoe UI", 16, "bold")
FONT_H2    = ("SF Pro Display", 11, "bold") if sys.platform == "darwin" else ("Segoe UI", 11, "bold")

LIMITE_DEFAULT = 65535


# ── Lógica de conversão (roda em thread) ────────────────────────────────────

def salvar_dataframe_em_xls(df, nome_sheet, caminho):
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(nome_sheet[:31])
    for ci, col in enumerate(df.columns):
        ws.write(0, ci, str(col))
    for ri, row in enumerate(df.itertuples(index=False), 1):
        for ci, val in enumerate(row):
            if pd.isna(val):
                ws.write(ri, ci, None)
            elif isinstance(val, (int, float, bool)):
                ws.write(ri, ci, val)
            else:
                ws.write(ri, ci, str(val))
    wb.save(caminho)


def converter(pasta_entrada, pasta_saida, limite, auto_split, log_cb, prog_cb, done_cb):
    os.makedirs(pasta_saida, exist_ok=True)
    arquivos = [f for f in os.listdir(pasta_entrada) if f.lower().endswith(".xlsx")]

    if not arquivos:
        log_cb("Nenhum arquivo .xlsx encontrado na pasta de entrada.", "warn")
        done_cb(0, 0)
        return

    total = len(arquivos)
    convertidos = 0
    erros = 0

    log_cb(f"Encontrados {total} arquivo(s) .xlsx.", "info")
    log_cb("─" * 52, "dim")

    for idx, arquivo in enumerate(arquivos):
        caminho_xlsx = os.path.join(pasta_entrada, arquivo)
        nome_base    = os.path.splitext(arquivo)[0]
        try:
            sheets = pd.read_excel(caminho_xlsx, sheet_name=None, engine="openpyxl")
            precisa_dividir = auto_split and any(len(df) > limite for df in sheets.values())

            if not precisa_dividir:
                caminho_xls = os.path.join(pasta_saida, nome_base + ".xls")
                wb = xlwt.Workbook(encoding="utf-8")
                for sname, df in sheets.items():
                    ws = wb.add_sheet(sname[:31])
                    for ci, col in enumerate(df.columns):
                        ws.write(0, ci, str(col))
                    for ri, row in enumerate(df.itertuples(index=False), 1):
                        for ci, val in enumerate(row):
                            if pd.isna(val):  ws.write(ri, ci, None)
                            elif isinstance(val, (int, float, bool)): ws.write(ri, ci, val)
                            else: ws.write(ri, ci, str(val))
                wb.save(caminho_xls)
                log_cb(f"✔  {arquivo}  →  {nome_base}.xls", "ok")
                convertidos += 1
            else:
                log_cb(f"⚠  Planilha grande: '{arquivo}' — dividindo...", "warn")
                for sname, df in sheets.items():
                    total_rows = len(df)
                    if total_rows <= limite:
                        saida_ind = os.path.join(pasta_saida, f"{nome_base}_{sname}.xls")
                        salvar_dataframe_em_xls(df, sname, saida_ind)
                        log_cb(f"   ✔  '{sname}'  →  {os.path.basename(saida_ind)}", "ok")
                    else:
                        partes = math.ceil(total_rows / limite)
                        log_cb(f"   ·  '{sname}' ({total_rows:,} linhas) → {partes} partes", "info")
                        for i in range(partes):
                            df_parte = df.iloc[i*limite:(i+1)*limite]
                            nome_parte = f"{nome_base}_{sname}_parte_{i+1}.xls"
                            salvar_dataframe_em_xls(df_parte, sname, os.path.join(pasta_saida, nome_parte))
                            log_cb(f"      ✔  {nome_parte}", "ok")
                convertidos += 1

        except Exception as e:
            log_cb(f"✖  Erro em {arquivo}: {e}", "err")
            erros += 1

        prog_cb((idx + 1) / total * 100)

    log_cb("─" * 52, "dim")
    done_cb(convertidos, erros)


# ── Interface gráfica ────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conversor XLSX → XLS")
        self.geometry("720x680")
        self.minsize(640, 580)
        self.configure(bg=BG)
        self.resizable(True, True)

        self._build_ui()

    # ── helpers de layout ────────────────────────────────────────────────────

    def _label(self, parent, text, font=None, fg=TEXT, **kw):
        return tk.Label(parent, text=text, bg=parent["bg"], fg=fg,
                        font=font or FONT_UI, **kw)

    def _card(self, parent, **kw):
        f = tk.Frame(parent, bg=BG_CARD, bd=0, highlightthickness=1,
                     highlightbackground=BORDER, **kw)
        return f

    def _section(self, parent, title):
        tk.Label(parent, text=title, bg=BG, fg=TEXT_DIM,
                 font=(*FONT_UI[:1], 9), anchor="w").pack(fill="x", pady=(12, 4), padx=2)

    def _path_row(self, parent, label, var, badge_text, badge_color):
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill="x", padx=14, pady=(0, 10))
        tk.Label(row, text=label, bg=BG_CARD, fg=TEXT_DIM,
                 font=(*FONT_UI[:1], 9), anchor="w").pack(fill="x", pady=(6, 3))
        inner = tk.Frame(row, bg=BG_CARD)
        inner.pack(fill="x")
        entry = tk.Entry(inner, textvariable=var, bg=BG_INPUT, fg=TEXT,
                         insertbackground=TEXT, relief="flat",
                         font=(*FONT_MONO[:1], 10), bd=0,
                         highlightthickness=1, highlightbackground=BORDER,
                         highlightcolor=ACCENT)
        entry.pack(side="left", fill="x", expand=True, ipady=6, ipadx=8)
        tk.Button(inner, text="...", bg=BG_INPUT, fg=TEXT_DIM,
                  font=(*FONT_UI[:1], 9), relief="flat", bd=0, cursor="hand2",
                  activebackground=BORDER, activeforeground=TEXT,
                  command=lambda v=var: self._browse(v)).pack(side="left", padx=(4, 6))
        badge = tk.Label(inner, text=badge_text, bg=badge_color,
                         fg=BG, font=(*FONT_UI[:1], 8, "bold"),
                         padx=8, pady=2, relief="flat")
        badge.pack(side="left")

    def _browse(self, var):
        path = filedialog.askdirectory(initialdir=var.get() or "/")
        if path:
            var.set(path)

    # ── construção principal ─────────────────────────────────────────────────

    def _build_ui(self):
        # título
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=20, pady=(18, 0))
        tk.Label(header, text="Conversor  XLSX → XLS", bg=BG, fg=TEXT,
                 font=FONT_H1, anchor="w").pack(side="left")
        tk.Label(header, text="v1.0", bg=BG, fg=TEXT_DIM,
                 font=(*FONT_UI[:1], 9), anchor="e").pack(side="right", pady=6)

        tk.Label(self, text="Converte planilhas modernas para o formato legado Excel 97–2003",
                 bg=BG, fg=TEXT_DIM, font=(*FONT_UI[:1], 9), anchor="w").pack(fill="x", padx=22, pady=(2, 12))

        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=20)

        # ── card caminhos ────────────────────────────────────────────────────
        self._section(main, "CAMINHOS")
        card_paths = self._card(main)
        card_paths.pack(fill="x", pady=(0, 4))

        self.var_entrada = tk.StringVar(value=r"C:\Softcom\Clientes\D & C FITNESS\banco")
        self.var_saida   = tk.StringVar(value=r"C:\Softcom\Clientes\D & C FITNESS\tratado")

        tk.Frame(card_paths, bg=BG_CARD, height=10).pack()
        self._path_row(card_paths, "Pasta de entrada (.xlsx)", self.var_entrada, "ENTRADA", INFO)
        self._path_row(card_paths, "Pasta de saída (.xls)",   self.var_saida,   "SAÍDA",   ACCENT)
        tk.Frame(card_paths, bg=BG_CARD, height=4).pack()

        # ── card parâmetros ──────────────────────────────────────────────────
        self._section(main, "PARÂMETROS")
        card_params = self._card(main)
        card_params.pack(fill="x", pady=(0, 4))
        tk.Frame(card_params, bg=BG_CARD, height=10).pack()

        row_lim = tk.Frame(card_params, bg=BG_CARD)
        row_lim.pack(fill="x", padx=14, pady=(0, 10))

        left_lim = tk.Frame(row_lim, bg=BG_CARD)
        left_lim.pack(side="left", fill="x", expand=True)
        tk.Label(left_lim, text="Limite de linhas por arquivo",
                 bg=BG_CARD, fg=TEXT, font=FONT_UI, anchor="w").pack(anchor="w")
        tk.Label(left_lim, text="Formato .xls suporta no máximo 65.535 linhas de dados",
                 bg=BG_CARD, fg=TEXT_DIM, font=(*FONT_UI[:1], 8), anchor="w").pack(anchor="w")

        self.var_limite = tk.IntVar(value=LIMITE_DEFAULT)
        spin = tk.Spinbox(row_lim, from_=100, to=65535, textvariable=self.var_limite,
                          width=9, bg=BG_INPUT, fg=TEXT, insertbackground=TEXT,
                          buttonbackground=BG_INPUT, relief="flat",
                          font=(*FONT_MONO[:1], 10), bd=0,
                          highlightthickness=1, highlightbackground=BORDER,
                          highlightcolor=ACCENT)
        spin.pack(side="right", ipady=6, ipadx=4)

        tk.Frame(card_params, bg=BORDER, height=1).pack(fill="x", padx=14)

        row_toggle = tk.Frame(card_params, bg=BG_CARD)
        row_toggle.pack(fill="x", padx=14, pady=10)
        tk.Label(row_toggle, text="Dividir automaticamente planilhas grandes",
                 bg=BG_CARD, fg=TEXT, font=FONT_UI, anchor="w").pack(side="left")

        self.var_split = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(row_toggle, variable=self.var_split,
                             bg=BG_CARD, activebackground=BG_CARD,
                             fg=ACCENT, selectcolor=BG_INPUT,
                             relief="flat", bd=0, cursor="hand2")
        chk.pack(side="right")

        # ── métricas ─────────────────────────────────────────────────────────
        self._section(main, "PROGRESSO")
        metrics = tk.Frame(main, bg=BG)
        metrics.pack(fill="x", pady=(0, 6))
        metrics.columnconfigure((0, 1, 2), weight=1, uniform="m")

        def metric_card(col, label, color):
            c = tk.Frame(metrics, bg=BG_CARD, bd=0, highlightthickness=1,
                         highlightbackground=BORDER)
            c.grid(row=0, column=col, sticky="ew", padx=(0 if col else 0, 6 if col < 2 else 0))
            tk.Label(c, text=label, bg=BG_CARD, fg=TEXT_DIM,
                     font=(*FONT_UI[:1], 8)).pack(pady=(10, 0))
            v = tk.StringVar(value="—")
            tk.Label(c, textvariable=v, bg=BG_CARD, fg=color,
                     font=(*FONT_H2[:1], 20, "bold")).pack(pady=(0, 10))
            return v

        metrics.grid_columnconfigure(0, weight=1)
        metrics.grid_columnconfigure(1, weight=1)
        metrics.grid_columnconfigure(2, weight=1)
        self.var_queue = metric_card(0, "NA FILA",     INFO)
        self.var_done  = metric_card(1, "CONVERTIDOS", ACCENT)
        self.var_errs  = metric_card(2, "ERROS",       WARN)

        # barra de progresso
        self.prog_var = tk.DoubleVar(value=0)
        prog_bg = tk.Frame(main, bg=BORDER, height=4)
        prog_bg.pack(fill="x", pady=(8, 6))
        prog_bg.pack_propagate(False)
        self.prog_fill = tk.Frame(prog_bg, bg=ACCENT, height=4)
        self.prog_fill.place(x=0, y=0, relheight=1, relwidth=0)

        # ── botão executar ────────────────────────────────────────────────────
        self.btn_run = tk.Button(main, text="▶   Executar conversão",
                                 bg=ACCENT, fg=BG, font=(*FONT_H2[:1], 11, "bold"),
                                 relief="flat", bd=0, cursor="hand2",
                                 activebackground=ACCENT_DIM, activeforeground=BG,
                                 padx=20, pady=12, command=self._run)
        self.btn_run.pack(fill="x", pady=(2, 8))

        # ── log ───────────────────────────────────────────────────────────────
        self._section(main, "LOG DE EXECUÇÃO")
        log_frame = self._card(main)
        log_frame.pack(fill="both", expand=True, pady=(0, 16))

        scroll = tk.Scrollbar(log_frame, bg=BG_CARD, troughcolor=BG_CARD,
                              relief="flat", bd=0, width=10)
        self.log_text = tk.Text(log_frame, bg=BG_CARD, fg=TEXT_DIM,
                                font=(*FONT_MONO[:1], 9), relief="flat", bd=0,
                                wrap="word", state="disabled",
                                yscrollcommand=scroll.set,
                                insertbackground=TEXT,
                                selectbackground=BORDER)
        scroll.config(command=self.log_text.yview)
        scroll.pack(side="right", fill="y", padx=(0, 4), pady=6)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=8)

        # tags de cor no log
        for tag, color in [("ok", ACCENT), ("warn", WARN), ("err", ERR),
                           ("info", INFO), ("dim", TEXT_DIM)]:
            self.log_text.tag_config(tag, foreground=color)

        self._log("Pronto. Configure os caminhos e clique em Executar.", "info")

    # ── lógica de execução ───────────────────────────────────────────────────

    def _log(self, msg, tag=None):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n", tag or "")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _set_prog(self, pct):
        self.prog_fill.place(relwidth=pct / 100)

    def _run(self):
        entrada = self.var_entrada.get().strip()
        saida   = self.var_saida.get().strip()
        limite  = self.var_limite.get()
        split   = self.var_split.get()

        if not entrada or not os.path.isdir(entrada):
            messagebox.showerror("Erro", "A pasta de entrada não existe ou não foi informada.")
            return
        if not saida:
            messagebox.showerror("Erro", "Informe a pasta de saída.")
            return

        self.btn_run.config(state="disabled", text="Convertendo…")
        self.var_queue.set("…")
        self.var_done.set("0")
        self.var_errs.set("0")
        self._set_prog(0)

        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

        def log_cb(msg, tag=None):
            self.after(0, lambda: self._log(msg, tag))

        def prog_cb(pct):
            self.after(0, lambda: self._set_prog(pct))

        def done_cb(conv, errs):
            def _finish():
                self.btn_run.config(state="normal", text="▶   Executar conversão")
                self.var_done.set(str(conv))
                self.var_errs.set(str(errs))
                if errs == 0:
                    self._log(f"✔  Concluído com sucesso! {conv} arquivo(s) convertido(s).", "ok")
                else:
                    self._log(f"Concluído. {conv} convertido(s), {errs} erro(s).", "warn")
            self.after(0, _finish)

        # conta arquivos antes de iniciar
        try:
            n = len([f for f in os.listdir(entrada) if f.lower().endswith(".xlsx")])
            self.var_queue.set(str(n))
        except Exception:
            pass

        t = threading.Thread(target=converter,
                             args=(entrada, saida, limite, split, log_cb, prog_cb, done_cb),
                             daemon=True)
        t.start()


# ── entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
