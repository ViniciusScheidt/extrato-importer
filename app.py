import os
import re
import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import numbers


# =========================
# Utilitários
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
    lines = read_pdf_lines(pdf_path)[:200]
    blob = "\n".join(lines)

    if "Sicredi" in blob or ("SICREDI" in blob.upper() and "COOPERATIVA" in blob.upper()):
        return "SICREDI"
    if "PagSeguro" in blob or "PAGBANK" in blob.upper() or "PagBank" in blob:
        return "PAGBANK"
    if "Stone Instituição de Pagamento" in blob or "STONE" in blob.upper():
        return "STONE"
    return "UNKNOWN"


# =========================
# Parsers
# =========================

def parse_sicredi(pdf_path: str, only_outgoing: bool = True) -> list[dict]:
    lines = read_pdf_lines(pdf_path)
    txs: list[dict] = []

    date_re = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(.*)$")

    for ln in lines:
        m = date_re.match(ln)
        if not m:
            continue

        date_str, rest = m.group(1), m.group(2)

        # Regras do Sicredi: apenas PAGAMENTO PIX e LIQUIDACAO BOLETO
        if ("PAGAMENTO PIX" not in rest) and ("LIQUIDACAO BOLETO" not in rest):
            continue

        # Geralmente a linha termina com "valor saldo" → pega penúltimo número (valor)
        nums = re.findall(r"-?\d{1,3}(?:\.\d{3})*,\d{2}", ln)
        if not nums:
            continue
        valor_str = nums[-2] if len(nums) >= 2 else nums[-1]
        valor = brl_to_float(valor_str)

        # Pagamentos costumam vir negativos; se only_outgoing, filtra positivos
        if only_outgoing and valor >= 0:
            continue

        # CPF/CNPJ está na descrição em muitos casos
        merged = rest.replace(" ", "")
        id_match = re.search(r"(\d{11}|\d{14})", merged) or re.search(r"(\d{11}|\d{14})", rest)
        cpf_cnpj = format_cpf_cnpj(id_match.group(1)) if id_match else ""

        # Remove valores no final da descrição (se existirem)
        desc = rest
        desc = re.sub(r"\s+-?\d{1,3}(?:\.\d{3})*,\d{2}\s+-?\d{1,3}(?:\.\d{3})*,\d{2}$", "", desc)
        desc = re.sub(r"\s+-?\d{1,3}(?:\.\d{3})*,\d{2}$", "", desc)

        dt = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
        txs.append({"date": dt, "cpf_cnpj": cpf_cnpj, "desc": desc, "valor_pago": abs(valor)})

    return txs


def parse_pagbank(pdf_path: str, only_outgoing: bool = True, ignore_balance: bool = True) -> list[dict]:
    lines = read_pdf_lines(pdf_path)
    txs: list[dict] = []

    # Ex.: 01/11/2025 Vendas - Disponivel DEBITO MASTERCARD R$ 39,14
    line_re = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(.*?)\s+R\$\s*([0-9\.\,]+)$")

    for ln in lines:
        m = line_re.match(ln)
        if not m:
            continue

        date_str, desc, val_str = m.group(1), m.group(2), m.group(3)

        if ignore_balance and ("Saldo do dia" in desc or "Rendimento" in desc):
            continue

        dt = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
        valor = brl_to_float(val_str)

        # PagBank nesse layout normalmente é entrada (positiva).
        # Se only_outgoing, ele vai filtrar tudo que for positivo.
        if only_outgoing and valor >= 0:
            continue

        txs.append({"date": dt, "cpf_cnpj": "", "desc": desc, "valor_pago": abs(valor)})

    return txs


def parse_stone(pdf_path: str, only_outgoing: bool = True, ignore_fees: bool = True) -> list[dict]:
    lines = read_pdf_lines(pdf_path)
    txs: list[dict] = []

    # Ex.: 30/11/25 Saída Tarifa - R$ 0,29 ...
    row_re = re.compile(r"^(\d{2}/\d{2}/\d{2})\s+(Entrada|Saída)\s+(.*?)\s+(-?\s*R\$\s*[0-9\.\,]+)")

    buf: list[str] = []
    for ln in lines:
        buf.append(ln)
        joined = " ".join(buf)

        m = row_re.match(joined)
        if not m:
            # evita buffer crescer infinito se o PDF quebrar muito
            if len(buf) > 6:
                buf = buf[-3:]
            continue

        buf = []

        date_str, typ, desc, val_str = m.group(1), m.group(2), m.group(3), m.group(4)
        dt = datetime.datetime.strptime(date_str, "%d/%m/%y").date()

        if ignore_fees and "Tarifa" in desc:
            continue

        valor = brl_to_float(val_str)

        if only_outgoing:
            if typ != "Saída":
                continue
            # garante saída como negativa
            valor = -abs(valor)

        # tenta achar CPF/CNPJ na descrição (nem sempre tem)
        digits = re.sub(r"\D+", "", desc)
        cpf_cnpj = ""
        m_id = re.search(r"(\d{11}|\d{14})", digits)
        if m_id:
            cpf_cnpj = format_cpf_cnpj(m_id.group(1))

        txs.append({"date": dt, "cpf_cnpj": cpf_cnpj, "desc": desc, "valor_pago": abs(valor)})

    return txs


# =========================
# Writer (planilha)
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
# UI
# =========================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Extrato → Planilha Contábil (Sicredi/PagBank/Stone)")
        self.geometry("820x340")

        self.xlsx_path = tk.StringVar()
        self.pdf_path = tk.StringVar()
        self.sheet_name = tk.StringVar()

        self.only_outgoing = tk.BooleanVar(value=True)
        self.ignore_fees = tk.BooleanVar(value=True)
        self.ignore_balance = tk.BooleanVar(value=True)

        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 6}

        frm1 = tk.Frame(self)
        frm1.pack(fill="x", **pad)
        tk.Label(frm1, text="Planilha (.xlsx):", width=18, anchor="w").pack(side="left")
        tk.Entry(frm1, textvariable=self.xlsx_path).pack(side="left", fill="x", expand=True)
        tk.Button(frm1, text="Selecionar", command=self.pick_xlsx).pack(side="left", padx=8)

        frm2 = tk.Frame(self)
        frm2.pack(fill="x", **pad)
        tk.Label(frm2, text="Aba (Banco):", width=18, anchor="w").pack(side="left")
        self.sheet_combo = ttk.Combobox(frm2, textvariable=self.sheet_name, state="readonly")
        self.sheet_combo.pack(side="left", fill="x", expand=True)

        frm3 = tk.Frame(self)
        frm3.pack(fill="x", **pad)
        tk.Label(frm3, text="Extrato (.pdf):", width=18, anchor="w").pack(side="left")
        tk.Entry(frm3, textvariable=self.pdf_path).pack(side="left", fill="x", expand=True)
        tk.Button(frm3, text="Selecionar", command=self.pick_pdf).pack(side="left", padx=8)

        frmF = tk.Frame(self)
        frmF.pack(fill="x", **pad)
        tk.Checkbutton(frmF, text="Somente saídas (pagamentos)", variable=self.only_outgoing).pack(anchor="w")
        tk.Checkbutton(frmF, text="Ignorar tarifas (Stone)", variable=self.ignore_fees).pack(anchor="w")
        tk.Checkbutton(frmF, text="Ignorar “Saldo do dia”/Rendimento (PagBank)", variable=self.ignore_balance).pack(anchor="w")

        frm4 = tk.Frame(self)
        frm4.pack(fill="x", **pad)
        tk.Button(frm4, text="Processar", command=self.process, height=2).pack(side="right")

        self.lbl_info = tk.Label(self, text="", fg="gray")
        self.lbl_info.pack(fill="x", padx=10, pady=10)

    def pick_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return

        self.xlsx_path.set(path)

        # ✅ Correção: não usar read_only=True (no Windows pode dar falha silenciosa / combo vazio)
        try:
            wb = load_workbook(path, data_only=True)
            sheets = wb.sheetnames
        except Exception as e:
            messagebox.showerror(
                "Erro ao ler planilha",
                "Não foi possível carregar as abas da planilha.\n\n"
                "Dicas:\n"
                "- Feche a planilha no Excel e tente novamente\n"
                "- Evite abrir direto do OneDrive/rede; copie para uma pasta local\n\n"
                f"Detalhe técnico: {str(e)}"
            )
            return

        if not sheets:
            messagebox.showerror("Erro", "Nenhuma aba encontrada na planilha.")
            return

        self.sheet_combo["values"] = sheets
        self.sheet_name.set(sheets[0])

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return

        self.pdf_path.set(path)

        try:
            bank = detect_bank(path)
        except Exception as e:
            messagebox.showerror("Erro", f"Não consegui ler o PDF.\n\n{str(e)}")
            return

        self.lbl_info.config(text=f"Banco detectado: {bank}")

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

        bank = detect_bank(pdf)
        if bank == "UNKNOWN":
            messagebox.showerror("Erro", "Não consegui detectar o banco do PDF.")
            return

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

            messagebox.showinfo("OK", f"Arquivo gerado:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Falhou", str(e))


if __name__ == "__main__":
    App().mainloop()