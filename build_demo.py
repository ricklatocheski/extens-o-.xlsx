#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_demo.py
=============
Gera uma pasta de trabalho de demonstracao (dist/DemoTemaEscuro.xlsx) com o
tema escuro ja aplicado ao conteudo, usando as mesmas cores do
DarkTheme.thmx. Serve para visualizar o resultado do suplemento sem precisar
executar as macros.

Uso:
    python3 build_demo.py     (requer: pip install openpyxl)
"""

import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(BASE_DIR, "dist")
OUTPUT = os.path.join(DIST_DIR, "DemoTemaEscuro.xlsx")

# Paleta moderna (identica ao tema .thmx e ao suplemento VBA)
FUNDO = "18191D"
FUNDO_ALT = "1F2127"
CABECALHO = "212530"
TEXTO = "E4E6EB"
TEXTO_FORTE = "FFFFFF"
ACENTO = "3B82F6"
GRADE = "2A2E38"


def main():
    os.makedirs(DIST_DIR, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Demonstracao"
    ws.sheet_view.showGridLines = False   # grade propria desligada; usamos zebra
    ws.sheet_properties.tabColor = ACENTO

    fundo = PatternFill("solid", fgColor=FUNDO)
    fundo_alt = PatternFill("solid", fgColor=FUNDO_ALT)
    cab = PatternFill("solid", fgColor=CABECALHO)
    grade = Border(*[Side(style="thin", color=GRADE)] * 4)
    borda_acento = Border(bottom=Side(style="medium", color=ACENTO))

    dados = [
        ("Produto", "Categoria", "Preco", "Estoque"),
        ("Teclado", "Perifericos", 149.90, 32),
        ("Monitor", "Telas", 899.00, 12),
        ("Mouse", "Perifericos", 79.90, 54),
        ("Notebook", "Computadores", 4299.00, 7),
        ("Webcam", "Perifericos", 199.90, 21),
        ("Headset", "Audio", 349.90, 18),
        ("Cadeira", "Moveis", 1290.00, 5),
    ]

    for r, linha in enumerate(dados, start=1):
        for c, valor in enumerate(linha, start=1):
            cel = ws.cell(row=r, column=c, value=valor)
            cel.alignment = Alignment(vertical="center")
            if r == 1:
                cel.fill = cab
                cel.font = Font(name="Segoe UI", color=TEXTO_FORTE, bold=True, size=11)
                cel.border = borda_acento
            else:
                cel.fill = fundo_alt if r % 2 == 0 else fundo
                cel.font = Font(name="Segoe UI", color=TEXTO, size=10)
                cel.border = grade

    # Margem extra de celulas escuras para o efeito de "modo escuro"
    for r in range(1, 30):
        for c in range(1, 8):
            cel = ws.cell(row=r, column=c)
            if cel.fill.fgColor.rgb in (None, "00000000"):
                cel.fill = fundo

    ws.row_dimensions[1].height = 24
    widths = {"A": 16, "B": 16, "C": 12, "D": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.freeze_panes = "A2"   # congela o cabecalho

    wb.save(OUTPUT)
    print(f"[OK] Demonstracao gerada: {OUTPUT}")


if __name__ == "__main__":
    main()
