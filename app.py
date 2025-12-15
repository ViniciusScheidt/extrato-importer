import os
import re
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import numbers


# =========================
# Utils
# =========================

def brl_to_float(s: str) -> float:
    """
    Converte string BRL para float.
    Aceita: "1.234,56", "-1.234,56", "R$ 1.234,56", "- R$ 0,29"
    """
    s = s.strip().replace("R$", "").replace(" ", "")
    neg = s.startswith("-")
    s = s.replace("-", "")
    s = s.replace(".", "").replace(",", ".")
    v = float(s)
    return -v if neg else v


def format_cpf_cnpj(num: str) -> str:
    num = re.sub(r"\D+", "", num)
    if len(num) == 11:
        return f"{num[0:3]}.{num[3:6]}.{num[6:9]}-{num[9:11]}"
    if len(num) == 14:
        return f"{num[0:2]}.{num[2:5]}.{num[5:8]}/{num[8:12]}-{num[12:14]}"
    return ""


def read_pdf_lines(pdf_path: str) -> list[str]:
    lines: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for ln in text.splitlines():
                ln = ln.strip()
                if ln:
                    lines.append(ln)
    return lines


def detect_bank(pdf_path: str) -> str:
    blob = "\n".join(read_pdf_lines(pdf_path)[:250]).upper()

    # heurísticas simples
    if "SICREDI" in blob:
        return "SICREDI"
    if "PAGBANK" in blob or "PAGSEGURO" in blob:
        return "PAGBANK"
    if "STONE" in blob:
        return "STONE"
    return "UNKNOWN"


# =========================
# Parsers
# =========================

def parse_sicredi(pdf_path: str, only_outgoing: bool = True) -> list[dict]:
    """
    Sicredi:
    - Aceita: PAGAMENTO PIX / LIQUIDACAO BOLETO / QUALQUER DEBITO CONVENIO
    - Valor pago = abs(valor)
    - Data baixa = data do extrato
    """
    lines = read_pdf_lines(pdf_path)
    txs: list[dict] = []

    date_re = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(.*)$")

    for ln in lines:
        m = date_re.match(ln)
        if not m:
            continue

        date_str, rest = m.group(1), m.group(2)
        full = ln.upper()

        # ✅ Regra pedida: qualquer débito convênio
        if (
            "PAGAMENTO PIX" not in full
            and "LIQUIDACAO BOLETO" not in full
            and "DEBITO CONVENIO" not in full
        ):
            continue

        # geralmente termina com "valor saldo" -> usa penúltimo número como valor
        nums = re.findall(r"-?\d{1,3}(?:\.\d{3})*,\d{2}", ln)
        if not nums:
            continue

        valor_str = nums[-2] if len(nums) >= 2 else nums[-1]
        valor = brl_to_float(valor_str)

        # saída geralmente vem negativa
        if only_outgoing and valor >= 0:
            continue

        # CPF/CNPJ pode estar contido no texto
        merged = rest.replace(" ", "")
        id_match = re.search(r"(\d{11}|\d{14})", merged) or re.search(r"(\d{11}|\d{14})", rest)
        cpf_cnpj = format_cpf_cnpj(id_match.group(1)) if id_match else ""

        # descrição sem os valores finais
        desc = rest
        desc = re.sub(r"\s+-?\d{1,3}(?:\.\d{3})*,\d{2}\s+-?\d{1,3}(?:\.\d{3})*,\d{2}$", "", desc)
        desc = re.sub(r"\s+-?\d{1,3}(?:\.\d{3})*,\d{2}$", "", desc)

        dt = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
        txs.append({
            "date": dt,
            "cpf_cnpj": cpf_cnpj,
            "desc": desc,
            "valor_pago": abs(valor),
        })

    return txs


def parse_pagbank(pdf_path: str, only_outgoing: bool = True, ignore_balance: bool = True) -> list[dict]:
    """
    PagBank (layout simples):
    01/11/2025 ... R$ 39,14
    """
    lines = read_pdf_lines(pdf_path)
    txs: list[dict] = []

    line_re = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(.*?)\s+R\$\s*([0-9\.\,]+)$")

    for ln in lines:
        m = line_re.match(ln)
        if not m:
            continue

        date_str, desc, val_str = m.group(1), m.group(2), m.group(3)
        if ignore_balance and ("SALDO DO DIA" in desc.upper() or "RENDIMENTO" in desc.upper()):
            continue

        dt = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
        valor = brl_to_float(val_str)

        if only_outgoing and valor >= 0:
            continue

        txs.append({
            "date": dt,
            "cpf_cnpj": "",
            "desc": desc,
            "valor_pago": abs(valor),
        })

    return txs


def parse_stone(pdf_path: str, only_outgoing: bool = True, ignore_fees: bool = True) -> list[dict]:
    """
    Stone (tabela):
    30/11/25 Saída ... - R$ 0,29 ...
    """
    lines = read_pdf_lines(pdf_path)
    txs: list[dict] = []

    row_re = re.compile(
        r"^(\d{2}/\d{2}/\d{2})\s+(Entrada|Saída)\s+(.*?)\s+(-?\s*R\$\s*[0-9\.\,]+)"
    )

    buf: list[str] = []
    for ln in lines:
        buf.append(ln)
        joined = " ".join(buf)

        m = row_re.match(joined)
        if not m:
            if len(buf) > 6:
                buf = buf[-3:]
            continue

        buf = []
        date_str, typ, desc, val_str = m.group(1), m.group(2), m.group(3), m.group(4)
        dt = datetime.datetime.strptime(date_str, "%d/%m/%y").date()

        if ignore_fees and "TARIFA" in desc.upper():
            continue

        valor = brl_to_float(val_str)

        if only_outgoing:
            if typ != "Saída":
                continue
            valor = -abs(valor)

        digits = re.sub(r"\D+", "", desc)
        cpf_cnpj = ""
        m_id = re.search(r"(\d{11}|\d{14})", digits)
        if m_id:
            cpf_cnpj = format_cpf_cnpj(m_id.group(1))

        txs.append({
            "date": dt,
            "cpf_cnpj": cpf_cnpj,
            "desc": desc,
            "valor_pago": abs(valor),
        })

    return txs


# =========================
# Writer
# =========================

REQUIRED_HEADERS = [
    "Número da Nota",
    "CPF/CNPJ do Fornecedor",
    "Data do Extrato Bancário",
    "Data da Baixa da Parcela",
    "Valor Pago",
    "Valor Juros",
    "Valor Multa",
    "Valor Desconto",
    "Código da Conta Banco/Caixa",
    "Caixa/Banco",
]


def fill_sheet(
    xlsx_path: str,
    out_path: str,
    sheet_name: str,
    txs: list[dict],
    bank_code: int,
    caixa_banco: str,
):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Aba '{sheet_name}' não existe na planilha.")

    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers) if h}

    for h in REQUIRED_HEADERS:
        if h not in col:
            raise ValueError(f"Coluna obrigatória não encontrada na aba '{sheet_name}': {h}")

    # limpa tudo abaixo do cabeçalho (primeiras 10 colunas do layout)
    for r in range(2, ws.max_row + 1):
        for c in range(1, 11):
            ws.cell(row=r, column=c).value = None

    row = 2
    for tx in txs:
        ws.cell(row=row, column=col["Número da Nota"]).value = None

        cpf = ws.cell(row=row, column=col["CPF/CNPJ do Fornecedor"])
        cpf.value = tx["cpf_cnpj"]
        cpf.number_format = "@"

        d1 = ws.cell(row=row, column=col["Data do Extrato Bancário"])
        d1.value = tx["date"]
        d1.number_format = "dd/mm/yyyy"

        d2 = ws.cell(row=row, column=col["Data da Baixa da Parcela"])
        d2.value = tx["date"]
        d2.number_format = "dd/mm/yyyy"

        v = ws.cell(row=row, column=col["Valor Pago"])
        v.value = float(tx["valor_pago"])
        v.number_format = numbers.FORMAT_NUMBER_00

        for k in ["Valor Juros", "Valor Multa", "Valor Desconto"]:
            z = ws.cell(row=row, column=col[k])
            z.value = 0.0
            z.number_format = numbers.FORMAT_NUMBER_00

        ws.cell(row=row, column=col["Código da Conta Banco/Caixa"]).value = bank_code
        ws.cell(row=row, column=col["Caixa/Banco"]).value = caixa_banco

        row += 1

    wb.save(out_path)


# =========================
# UI (mais moderna)
# =========================

class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master)
        self.master = master

        self.xlsx_path = tk.StringVar()
        self.pdf_path = tk.StringVar()
        self.sheet_name = tk.StringVar()

        self.only_outgoing = tk.BooleanVar(value=True)
        self.ignore_fees = tk.BooleanVar(value=True)
        self.ignore_balance = tk.BooleanVar(value=True)

        self.detected_bank = tk.StringVar(value="—")

        self._build_styles()
        self._build_layout()

    def _build_styles(self):
        self.master.title("Importador de Extratos → Planilha (Sicredi / PagBank / Stone)")
        self.master.geometry("860x420")
        self.master.minsize(860, 420)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background="#f5f6f8")
        style.configure("Card.TFrame", background="white", relief="flat")
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"), background="#f5f6f8")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#555", background="#f5f6f8")
        style.configure("CardTitle.TLabel", font=("Segoe UI", 11, "bold"), background="white")
        style.configure("Muted.TLabel", font=("Segoe UI", 9), foreground="#666")
        style.configure("Status.TLabel", font=("Segoe UI", 9), foreground="#333", background="#f5f6f8")

        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TCheckbutton", background="white")

        style.configure("TEntry", padding=6)
        style.configure("TCombobox", padding=4)

    def _build_layout(self):
        self.pack(fill="both", expand=True, padx=16, pady=16)

        # Header
        header = ttk.Frame(self)
        header.pack(fill="x", pady=(0, 12))
        ttk.Label(header, text="Importador de Extratos", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Selecione a planilha, escolha a aba do banco, selecione o PDF do extrato e processe.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        # Main cards container
        cards = ttk.Frame(self)
        cards.pack(fill="both", expand=True)

        # Card: arquivos
        card_files = ttk.Frame(cards, style="Card.TFrame")
        card_files.pack(fill="x", pady=(0, 12))
        card_files.configure(padding=14)

        ttk.Label(card_files, text="Arquivos", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", columnspan=3, pady=(0, 10))

        ttk.Label(card_files, text="Planilha (.xlsx)", style="Muted.TLabel").grid(row=1, column=0, sticky="w")
        self.ent_xlsx = ttk.Entry(card_files, textvariable=self.xlsx_path)
        self.ent_xlsx.grid(row=2, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(card_files, text="Selecionar…", command=self.pick_xlsx).grid(row=2, column=1, sticky="ew")

        ttk.Label(card_files, text="Aba (Banco)", style="Muted.TLabel").grid(row=1, column=2, sticky="w", padx=(12, 0))
        self.sheet_combo = ttk.Combobox(card_files, textvariable=self.sheet_name, state="readonly")
        self.sheet_combo.grid(row=2, column=2, sticky="ew", padx=(12, 0))

        ttk.Label(card_files, text="Extrato (.pdf)", style="Muted.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.ent_pdf = ttk.Entry(card_files, textvariable=self.pdf_path)
        self.ent_pdf.grid(row=4, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(card_files, text="Selecionar…", command=self.pick_pdf).grid(row=4, column=1, sticky="ew")

        ttk.Label(card_files, text="Banco detectado", style="Muted.TLabel").grid(row=3, column=2, sticky="w", padx=(12, 0), pady=(10, 0))
        ttk.Label(card_files, textvariable=self.detected_bank, style="CardTitle.TLabel").grid(row=4, column=2, sticky="w", padx=(12, 0))

        card_files.columnconfigure(0, weight=3)
        card_files.columnconfigure(1, weight=0)
        card_files.columnconfigure(2, weight=2)

        # Card: opções
        card_opts = ttk.Frame(cards, style="Card.TFrame")
        card_opts.pack(fill="x", pady=(0, 12))
        card_opts.configure(padding=14)

        ttk.Label(card_opts, text="Opções", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))

        ttk.Checkbutton(card_opts, text="Somente saídas (pagamentos)", variable=self.only_outgoing).pack(anchor="w")
        ttk.Checkbutton(card_opts, text="Ignorar tarifas (Stone)", variable=self.ignore_fees).pack(anchor="w")
        ttk.Checkbutton(card_opts, text="Ignorar “Saldo do dia”/Rendimento (PagBank)", variable=self.ignore_balance).pack(anchor="w")

        # Footer actions
        footer = ttk.Frame(self)
        footer.pack(fill="x")

        self.status = tk.StringVar(value="Pronto.")
        ttk.Label(footer, textvariable=self.status, style="Status.TLabel").pack(side="left")

        self.btn_process = ttk.Button(footer, text="Processar", style="Primary.TButton", command=self.process)
        self.btn_process.pack(side="right")

    def _set_status(self, msg: str):
        self.status.set(msg)
        self.master.update_idletasks()

    def pick_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return

        self.xlsx_path.set(path)
        self._set_status("Lendo abas da planilha…")

        # ✅ Correção: não usar read_only=True (evita combo vazio / falha silenciosa em Windows)
        try:
            wb = load_workbook(path, data_only=True)
            sheets = wb.sheetnames
        except Exception as e:
            messagebox.showerror(
                "Erro ao ler planilha",
                "Não foi possível carregar as abas da planilha.\n\n"
                "Dicas:\n"
                "- Feche a planilha no Excel e tente novamente\n"
                "- Copie para uma pasta local (evite OneDrive/rede)\n\n"
                f"Detalhe: {str(e)}"
            )
            self._set_status("Falha ao ler a planilha.")
            return

        if not sheets:
            messagebox.showerror("Erro", "Nenhuma aba encontrada na planilha.")
            self._set_status("Nenhuma aba encontrada.")
            return

        self.sheet_combo["values"] = sheets
        self.sheet_name.set(sheets[0])
        self._set_status("Planilha carregada. Selecione a aba e o PDF.")

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return

        self.pdf_path.set(path)
        self._set_status("Lendo PDF e detectando banco…")

        try:
            bank = detect_bank(path)
        except Exception as e:
            messagebox.showerror("Erro", f"Não consegui ler o PDF.\n\n{str(e)}")
            self._set_status("Falha ao ler PDF.")
            return

        self.detected_bank.set(bank)
        self._set_status(f"PDF selecionado. Banco detectado: {bank}")

    def process(self):
        xlsx = self.xlsx_path.get().strip()
        pdf = self.pdf_path.get().strip()
        sheet = self.sheet_name.get().strip()

        if not os.path.exists(xlsx):
            messagebox.showerror("Erro", "Selecione uma planilha .xlsx válida.")
            return
        if not os.path.exists(pdf):
            messagebox.showerror("Erro", "Selecione um extrato .pdf válido.")
            return
        if not sheet:
            messagebox.showerror("Erro", "Selecione uma aba.")
            return

        bank = self.detected_bank.get()
        if bank == "—" or bank == "UNKNOWN":
            # tenta detectar na hora, caso o usuário não tenha clicado no botão do PDF antes
            bank = detect_bank(pdf)
            self.detected_bank.set(bank)

        if bank == "UNKNOWN":
            messagebox.showerror("Erro", "Não consegui detectar o banco do PDF.")
            return

        # UI feedback
        self.btn_process.configure(state="disabled")
        self._set_status("Processando…")

        try:
            if bank == "SICREDI":
                txs = parse_sicredi(pdf, only_outgoing=self.only_outgoing.get())
            elif bank == "PAGBANK":
                txs = parse_pagbank(pdf, only_outgoing=self.only_outgoing.get(), ignore_balance=self.ignore_balance.get())
            elif bank == "STONE":
                txs = parse_stone(pdf, only_outgoing=self.only_outgoing.get(), ignore_fees=self.ignore_fees.get())
            else:
                raise ValueError("Banco não suportado.")

            out_path = os.path.splitext(xlsx)[0] + f"_PREENCHIDO_{bank}.xlsx"
            fill_sheet(
                xlsx_path=xlsx,
                out_path=out_path,
                sheet_name=sheet,
                txs=txs,
                bank_code=8,
                caixa_banco=sheet,
            )

            messagebox.showinfo("OK", f"Arquivo gerado:\n{out_path}\n\nRegistros importados: {len(txs)}")
            self._set_status(f"Concluído. Registros: {len(txs)}")
        except Exception as e:
            messagebox.showerror("Falhou", str(e))
            self._set_status("Falhou.")
        finally:
            self.btn_process.configure(state="normal")


def main():
    root = tk.Tk()
    # melhora a aparência no Windows (DPI / fonte)
    try:
        root.tk.call("tk", "scaling", 1.1)
    except Exception:
        pass

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()