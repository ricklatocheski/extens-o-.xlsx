# Tema escuro automático em TODO arquivo que você abrir

Estas instruções fazem o Excel aplicar o tema escuro **sozinho**, toda vez
que você abrir qualquer planilha — sem clicar em nada. Usa a **Pasta de
Trabalho Pessoal de Macros** (`Personal.xlsb`), que carrega automaticamente
quando o Excel inicia.

## Passo 1 — Criar o Personal.xlsb (se ainda não existir)

1. Na guia **Exibir → Macros → Gravar Macro...**
2. Em **"Armazenar macro em"**, escolha **"Pasta de trabalho pessoal de
   macros"** → **OK**.
3. Clique em qualquer célula e depois **Exibir → Macros → Parar Gravação**.

Isso cria o `Personal.xlsb`. (Se você já tem macros pessoais, pule este passo.)

## Passo 2 — Abrir o editor VBA

Pressione **`Alt + F11`**. No painel esquerdo (Explorador de Projetos),
localize **VBAProject (PERSONAL.XLSB)**.

## Passo 3 — Adicionar o Módulo de Classe

1. Clique com o botão direito em **VBAProject (PERSONAL.XLSB) → Inserir →
   Módulo de Classe**.
2. Com o módulo selecionado, abra a janela **Propriedades** (`F4`) e mude o
   campo **(Name)** para **`clsAutoTema`**.
3. Cole todo o conteúdo de **`clsAutoTema.cls`** (a partir da linha
   `Option Explicit`; ignore o cabeçalho `VERSION/BEGIN/Attribute` se copiar
   o texto manualmente).

## Passo 4 — Adicionar o Módulo padrão

1. Botão direito em **VBAProject (PERSONAL.XLSB) → Inserir → Módulo**.
2. Cole todo o conteúdo de **`AutoTemaEscuro.bas`**.

## Passo 5 — Ativar agora

- Clique dentro do procedimento `Auto_Open` e aperte **`F5`** uma vez
  (ou apenas **feche e reabra o Excel**).
- A partir daí, **todo arquivo que você abrir fica escuro automaticamente**.

## Passo 6 — Salvar o Personal.xlsb

Ao fechar o Excel, ele vai perguntar se deseja salvar as alterações na
**Pasta de trabalho pessoal de macros** → clique **Salvar**. (Assim a
configuração fica permanente.)

---

## Observações

- **Desfazer em um arquivo específico:** rode a macro `RemoverTemaEscuro`
  (do módulo principal) ou desfaça com `Ctrl+Z` logo após abrir.
- Como o tema é aplicado ao abrir, o Excel marca o arquivo como "modificado".
  Se você só consultou e não quer salvar, feche escolhendo **"Não salvar"**.
- Se as macros estiverem bloqueadas, habilite em **Arquivo → Opções →
  Central de Confiabilidade → Configurações de Macro**.
- Para **parar** o tema automático: no `Personal.xlsb`, apague (ou renomeie)
  o `Auto_Open`, ou remova os dois módulos.
