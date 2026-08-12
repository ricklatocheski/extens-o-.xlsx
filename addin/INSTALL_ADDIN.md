# Montar e instalar o suplemento `.xlam`

Estas instruções transformam o código-fonte (`DarkThemeAddin.bas` +
`customUI/customUI.xml`) em um suplemento do Excel (`.xlam`) com a guia
**"Tema Escuro"** na Faixa de Opções. Testado no Excel 2007; funciona
igualmente nas versões 2010/2013/2016/2019/365.

> **Por que não há um `.xlam` pronto no repositório?**
> Um `.xlam` embute um `vbaProject.bin` — um projeto VBA já compilado em um
> formato binário OLE que só o próprio Excel gera de forma confiável.
> Distribuir o **código-fonte** e montá-lo no Excel é a forma correta,
> segura e auditável de entregar a macro.

---

## Passo 1 — Importar o módulo VBA

1. Abra o Excel com uma pasta de trabalho em branco.
2. Pressione `Alt + F11` para abrir o **Editor do Visual Basic (VBE)**.
3. Menu **Arquivo → Importar Arquivo...** e selecione
   `addin/DarkThemeAddin.bas`.
4. Confirme que o módulo `DarkThemeAddin` apareceu no Explorador de Projetos.

> Se o Excel bloquear macros: **Arquivo → Opções → Central de
> Confiabilidade → Configurações da Central de Confiabilidade →
> Configurações de Macro → "Desabilitar todas as macros com notificação"**
> e habilite quando solicitado.

## Passo 2 — Salvar como suplemento

1. De volta ao Excel, **Arquivo → Salvar Como**.
2. Em "Tipo", escolha **Suplemento do Excel (`*.xlam`)**.
3. Nomeie como `DarkThemeAddin.xlam`. O Excel sugerirá a pasta padrão
   de suplementos (`...\AppData\Roaming\Microsoft\AddIns\`) — mantenha-a.
4. Feche o Excel.

## Passo 3 — Adicionar a guia da Faixa de Opções

O `.xlam` recém-salvo ainda não tem a guia personalizada. Para adicioná-la:

1. Baixe e instale o **Custom UI Editor for Microsoft Office** (gratuito).
2. Abra o `DarkThemeAddin.xlam` nesse editor.
3. Menu **Insert → Office 2007 Custom UI Part** (cria `customUI/customUI.xml`).
4. Copie todo o conteúdo de `addin/customUI/customUI.xml` para dentro dessa
   parte e salve.

> Alternativa manual: um `.xlam` também é um ZIP. É possível abrir com um
> descompactador, criar a pasta `customUI/` com o `customUI.xml` e adicionar
> a relação em `_rels/.rels`. O Custom UI Editor faz isso automaticamente e
> sem erros.

## Passo 4 — Ativar o suplemento

1. Abra o Excel novamente.
2. **Arquivo → Opções → Suplementos**.
3. No rodapé, em "Gerenciar", escolha **Suplementos do Excel** e clique
   **Ir...**.
4. Marque **DarkThemeAddin** (ou clique **Procurar...** e selecione o
   `.xlam`, caso não apareça).
5. Clique **OK**.

Pronto — a guia **"Tema Escuro"** aparece na Faixa de Opções com os botões:

- **Aplicar (planilha)** → escurece a planilha ativa.
- **Aplicar (tudo)** → escurece todas as planilhas da pasta.
- **Remover** → restaura o visual claro padrão.

---

## Uso via macro (sem a Faixa de Opções)

Se preferir não configurar a guia, chame as macros diretamente por
`Alt + F8`:

| Macro                        | Efeito                                   |
|------------------------------|------------------------------------------|
| `AplicarTemaEscuro`          | Tema escuro na planilha ativa            |
| `AplicarTemaEscuroTodas`     | Tema escuro em todas as planilhas        |
| `RemoverTemaEscuro`          | Volta ao visual claro                    |
