#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_theme.py
==============
Empacota os arquivos XML em src/theme/ no arquivo de Tema do Office 2007
(dist/DarkTheme.thmx). Um .thmx e simplesmente um pacote OPC (ZIP) com a
estrutura de partes definida pela Open Packaging Conventions.

Uso:
    python3 build_theme.py

Nao requer dependencias externas (apenas a biblioteca padrao do Python).
"""

import os
import sys
import zipfile
import xml.dom.minidom as minidom

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(BASE_DIR, "src", "theme")
DIST_DIR = os.path.join(BASE_DIR, "dist")
OUTPUT = os.path.join(DIST_DIR, "DarkTheme.thmx")

# Ordem canonica das partes dentro do pacote. O [Content_Types].xml
# deve ser sempre a primeira entrada de um pacote OPC.
PART_ORDER = [
    "[Content_Types].xml",
    "_rels/.rels",
    "theme/theme/themeManager.xml",
    "theme/theme/_rels/themeManager.xml.rels",
    "theme/theme/theme1.xml",
]


def validar_xml(caminho):
    """Garante que cada parte XML esta bem formada antes de empacotar."""
    with open(caminho, "rb") as f:
        dados = f.read()
    try:
        minidom.parseString(dados)
    except Exception as exc:  # noqa: BLE001
        raise SystemExit(f"[ERRO] XML invalido em {caminho}: {exc}")
    return dados


def main():
    if not os.path.isdir(SRC_DIR):
        raise SystemExit(f"[ERRO] Pasta de origem nao encontrada: {SRC_DIR}")

    os.makedirs(DIST_DIR, exist_ok=True)

    # Verifica que todas as partes esperadas existem.
    faltando = [
        p for p in PART_ORDER
        if not os.path.isfile(os.path.join(SRC_DIR, p))
    ]
    if faltando:
        raise SystemExit(
            "[ERRO] Partes ausentes em src/theme/:\n  - "
            + "\n  - ".join(faltando)
        )

    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        for parte in PART_ORDER:
            caminho = os.path.join(SRC_DIR, parte)
            dados = validar_xml(caminho)
            # Usa sempre "/" como separador dentro do ZIP (padrao OPC).
            z.writestr(parte.replace(os.sep, "/"), dados)

    tamanho = os.path.getsize(OUTPUT)
    print(f"[OK] Tema gerado: {OUTPUT} ({tamanho} bytes)")
    print("     Instale-o pela guia Layout da Pagina > Temas > Procurar Temas.")


if __name__ == "__main__":
    sys.exit(main())
