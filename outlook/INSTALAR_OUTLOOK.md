# Tema Escuro Animado no Outlook 2019

Este guia cobre as duas partes: escurecer o **Outlook** (nativo) e criar
**e-mails com tema escuro colorido** (macro VBA).

---

## Parte 1 — Escurecer o Outlook (sem add-in, nativo)

1. **Interface (menus, faixa):** Arquivo → **Conta do Office** → **Tema do
   Office** → escolha **"Preto"**.
2. **Fundo da mensagem ao ler:** ao abrir um e-mail, use o botão **sol/lua**
   na faixa (Mensagem) para alternar o fundo entre claro e escuro.

Isso já deixa o Outlook em modo escuro. A Parte 2 é para o **conteúdo dos
e-mails que você envia**.

---

## Parte 2 — Macro que cria e-mails com tema escuro animado

### Instalar

1. No Outlook, pressione **`Alt + F11`** (abre o editor VBA).
2. No painel esquerdo, expanda **Project1 → Microsoft Outlook Objetos**.
3. Menu **Inserir → Módulo**.
4. Cole todo o conteúdo de **`DarkMailOutlook.bas`**.
5. Salve (`Ctrl + S`).

### Habilitar macros (necessário uma vez)

Arquivo → Opções → **Central de Confiabilidade** → Configurações da Central
de Confiabilidade → **Configurações de Macro** → marque
**"Notificações para todas as macros"** e reinicie o Outlook (habilite quando
perguntar). Em ambientes corporativos pode ser preciso assinar o projeto VBA.

### Usar

- Aperte **`Alt + F8`**, escolha **`NovoEmailEscuro`** → **Executar**.
  Abre um e-mail novo já com o tema escuro colorido.
- Para colocar num botão fixo: **Arquivo → Opções → Barra de Ferramentas de
  Acesso Rápido** → em "Escolher comandos" selecione **Macros** → adicione
  `NovoEmailEscuro`. Agora é 1 clique no topo do Outlook.

### Colocar um GIF animado (movimento de verdade)

Na macro `NovoEmailEscuro`, troque o último parâmetro `""` pela URL de um GIF:

```vba
mail.HTMLBody = TemplateEscuro("Ola! ", "Sua mensagem...", _
    "https://.../seu-gif.gif")
```

O GIF aparece animado no topo do e-mail, inclusive para quem recebe.

---

## Limitações honestas (importante)

- **E-mail não roda animação de CSS nem JavaScript.** O Outlook renderiza com
  o motor do Word e remove esses efeitos. Movimento real só com **GIF**.
- **Gradientes** não aparecem no Outlook (desktop): o modelo já traz uma
  **cor sólida de reserva**, então o visual não quebra — só fica sem o
  degradê para quem usa Outlook. Em Gmail/Apple Mail o degradê aparece.
- **Cantos arredondados** são ignorados pelo Outlook (ficam retos) — normal.

## Suplemento "oficial" (Office Web Add-in) — opcional

Existe o formato moderno de add-in do Outlook (manifesto + HTML/JS com
Office.js), que coloca um botão na faixa e um painel lateral. Ele é mais
poderoso, mas exige **hospedar os arquivos em HTTPS** e **fazer sideload** do
manifesto — bem mais trabalhoso que a macro. Se quiser seguir por esse
caminho, peça que eu monto o esqueleto do projeto.
