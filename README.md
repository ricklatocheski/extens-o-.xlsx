# Suplemento Tema Escuro para Excel 2007 (.xlsx)

Um suplemento (add-in) que traz um **tema escuro** para o Microsoft Excel
2007 e versões superiores. O Excel 2007 não possui modo escuro nativo, então
este projeto oferece **duas peças complementares** que, juntas, entregam a
experiência de tema escuro:

| Peça | Arquivo | O que faz |
|------|---------|-----------|
| **Tema de cores do Office** | `dist/DarkTheme.thmx` | Instala uma paleta escura no seletor de temas do Excel. Tabelas, gráficos, formas e realces passam a usar cores escuras. |
| **Suplemento VBA** | `addin/DarkThemeAddin.bas` + `addin/customUI/customUI.xml` | Adiciona a guia **"Tema Escuro"** na Faixa de Opções, com botões que pintam o fundo das células, clareiam a fonte, desligam as linhas de grade e colorem a aba — o "modo escuro" completo da planilha. |

Também há uma pasta de **demonstração** já pronta em
`dist/DemoTemaEscuro.xlsx`, para você ver o resultado imediatamente.

---

## Estrutura do projeto

```
.
├── README.md               ← este arquivo
├── build_theme.py          ← empacota o .thmx a partir de src/theme/
├── build_demo.py           ← gera a planilha de demonstração
├── src/theme/              ← fontes XML (Open Packaging Conventions) do tema
│   ├── [Content_Types].xml
│   ├── _rels/.rels
│   └── theme/theme/
│       ├── themeManager.xml
│       ├── theme1.xml       ← paleta de cores escura
│       └── _rels/themeManager.xml.rels
├── addin/
│   ├── DarkThemeAddin.bas   ← código VBA do suplemento
│   ├── customUI/customUI.xml← guia/botões na Faixa de Opções
│   └── INSTALL_ADDIN.md     ← como montar o arquivo .xlam
└── dist/                    ← saída gerada
    ├── DarkTheme.thmx
    └── DemoTemaEscuro.xlsx
```

---

## Paleta do tema escuro

| Uso                | Cor       | RGB              |
|--------------------|-----------|------------------|
| Fundo principal    | `#18191D` | 24, 25, 29       |
| Fundo alternado (zebra) | `#1F2127` | 31, 33, 39  |
| Cabeçalho          | `#212530` | 33, 37, 48       |
| Texto              | `#E4E6EB` | 228, 230, 235    |
| Texto forte        | `#FFFFFF` | 255, 255, 255    |
| Azul de acento     | `#3B82F6` | 59, 130, 246     |
| Grade sutil        | `#2A2E38` | 42, 46, 56       |

Estilo moderno: fundo grafite levemente azulado, **zebra** nas linhas,
cabeçalho com **borda de acento** e **congelamento**, texto off-white suave
e fonte **Segoe UI**.

---

## Como instalar

### 1. Tema de cores (`DarkTheme.thmx`)

1. No Excel: guia **Layout da Página → Temas → Procurar Temas...**
2. Selecione `dist/DarkTheme.thmx`.
   - Para que ele apareça permanentemente na galeria de temas, copie o
     arquivo para:
     `C:\Users\<seu_usuario>\AppData\Roaming\Microsoft\Templates\Document Themes\`

### 2. Suplemento VBA (`.xlam`)

O código VBA e a personalização da Faixa de Opções precisam ser reunidos em
um arquivo `.xlam`. O passo a passo completo está em
[`addin/INSTALL_ADDIN.md`](addin/INSTALL_ADDIN.md).

Resumo:
1. No Excel, abra o Editor do VBA (`Alt+F11`), importe `DarkThemeAddin.bas`.
2. Salve a pasta como **Suplemento do Excel (`*.xlam`)**.
3. Adicione a guia da Faixa de Opções injetando `customUI/customUI.xml` no
   `.xlam` com o *Custom UI Editor for Microsoft Office*.
4. Ative em **Arquivo → Opções → Suplementos → Ir... → Procurar**.

---

## Como recompilar

```bash
python3 build_theme.py     # gera dist/DarkTheme.thmx
python3 build_demo.py      # gera dist/DemoTemaEscuro.xlsx (requer openpyxl)
```

`build_theme.py` usa apenas a biblioteca padrão do Python. `build_demo.py`
precisa de `openpyxl` (`pip install openpyxl`).

---

## Observações técnicas

- Um arquivo `.thmx` é um pacote OPC (ZIP) com XML DrawingML — o mesmo
  formato de tema introduzido no Office 2007. Por isso é totalmente
  compatível com a extensão `.xlsx`.
- O Excel 2007 não permite escurecer a **interface do programa** (menus,
  faixa de opções) via arquivo ou macro — isso depende do tema do próprio
  Windows/Office. O suplemento atua sobre o **conteúdo das planilhas**, que é
  onde o tema escuro faz diferença visual.
- As três peças usam exatamente a **mesma paleta**, garantindo aparência
  consistente entre células, tabelas e gráficos.
