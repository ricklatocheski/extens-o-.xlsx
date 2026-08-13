Attribute VB_Name = "DarkMailOutlook"
'============================================================
' Tema Escuro Animado para e-mails do Outlook 2019
'------------------------------------------------------------
' Cria e-mails com um modelo ESCURO e COLORIDO (fundo grafite,
' texto claro, cabecalho com gradiente vibrante, botao de acao
' e faixa colorida) para deixar a mensagem mais "animada".
'
' Como instalar: veja outlook/INSTALAR_OUTLOOK.md
'
' OBS: e-mail nao executa animacao de CSS (o Outlook usa o motor
' do Word). Para movimento real, use um GIF animado - veja a
' funcao TemplateEscuro (parametro urlGif).
'============================================================
Option Explicit

' --- Paleta ---
Private Const C_FUNDO As String = "#0F1115"      ' fundo externo
Private Const C_CARTAO As String = "#171A21"     ' cartao da mensagem
Private Const C_RODAPE As String = "#12141A"     ' rodape
Private Const C_TEXTO As String = "#E4E6EB"      ' texto claro
Private Const C_SUAVE As String = "#9CA3AF"      ' texto secundario
Private Const C_BORDA As String = "#262B36"      ' borda do cartao
Private Const C_AZUL As String = "#3B82F6"       ' acento azul
Private Const C_ROXO As String = "#A855F7"       ' acento roxo

'------------------------------------------------------------
' Abre um e-mail NOVO ja com o tema escuro animado.
'------------------------------------------------------------
Public Sub NovoEmailEscuro()
    Dim mail As Outlook.MailItem
    Set mail = Application.CreateItem(olMailItem)
    mail.BodyFormat = olFormatHTML
    ' Deixe urlGif = "" ou informe a URL de um GIF para animar o topo.
    mail.HTMLBody = TemplateEscuro("Ola! ", _
        "Escreva sua mensagem aqui. Este modelo ja vem com fundo escuro, " & _
        "texto claro e um toque colorido para deixar tudo mais animado.", "")
    mail.Display
End Sub

'------------------------------------------------------------
' Aplica um fundo escuro ao e-mail que esta aberto (compose).
' Funciona melhor em e-mails novos escritos por voce.
'------------------------------------------------------------
Public Sub AplicarEscuroNoEmailAtual()
    Dim insp As Outlook.Inspector
    Set insp = Application.ActiveInspector
    If insp Is Nothing Then
        MsgBox "Abra um e-mail (nova mensagem) primeiro.", vbInformation
        Exit Sub
    End If
    If Not (TypeOf insp.CurrentItem Is Outlook.MailItem) Then Exit Sub

    Dim mail As Outlook.MailItem
    Set mail = insp.CurrentItem

    Dim css As String
    css = "<style>body,table,td,div,p,span,li{background-color:" & C_FUNDO & _
          " !important;color:" & C_TEXTO & " !important;}" & _
          "a{color:" & C_AZUL & " !important;}</style>"

    Dim s As String
    s = mail.HTMLBody
    If InStr(1, s, "<head", vbTextCompare) > 0 Then
        mail.HTMLBody = Replace(s, "<head>", "<head>" & css, 1, 1, vbTextCompare)
    Else
        mail.HTMLBody = css & s
    End If
End Sub

'------------------------------------------------------------
' Monta o HTML do modelo escuro. Usa aspas simples nos atributos
' HTML para nao precisar escapar aspas no VBA. Inclui cores
' solidas de reserva (fallback) porque o Outlook nao renderiza
' gradientes - assim o visual nao "quebra" para quem recebe.
'------------------------------------------------------------
Public Function TemplateEscuro(ByVal titulo As String, _
                               ByVal corpo As String, _
                               ByVal urlGif As String) As String
    Dim h As String
    Dim gif As String

    gif = ""
    If Len(urlGif) > 0 Then
        gif = "<div style='margin-top:14px;'>" & _
              "<img src='" & urlGif & "' width='120' alt='' " & _
              "style='border:0;display:inline-block;'></div>"
    End If

    h = ""
    h = h & "<div style='margin:0;padding:0;background:" & C_FUNDO & ";'>"
    h = h & "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:" & C_FUNDO & ";'>"
    h = h & "<tr><td align='center' style='padding:24px 12px;'>"
    h = h & "<table role='presentation' width='600' cellpadding='0' cellspacing='0' style='max-width:600px;width:100%;background:" & C_CARTAO & ";border:1px solid " & C_BORDA & ";border-radius:16px;'>"

    ' Cabecalho com gradiente vibrante (fallback solido roxo/azul)
    h = h & "<tr><td style='background:#6D5EF7;background:linear-gradient(135deg," & C_AZUL & "," & C_ROXO & ");padding:34px 28px;border-radius:16px 16px 0 0;'>"
    h = h & "<div style='font-size:26px;font-weight:700;color:#FFFFFF;font-family:Segoe UI,Arial,sans-serif;'>" & titulo & "&#127770;</div>"
    h = h & "<div style='font-size:14px;color:#EEF2FF;margin-top:6px;font-family:Segoe UI,Arial,sans-serif;'>Uma mensagem com visual escuro e moderno</div>"
    h = h & gif
    h = h & "</td></tr>"

    ' Corpo
    h = h & "<tr><td style='padding:28px;'>"
    h = h & "<div style='font-size:16px;line-height:1.7;color:" & C_TEXTO & ";font-family:Segoe UI,Arial,sans-serif;'>" & corpo & "</div>"
    ' Botao de acao (fallback solido azul)
    h = h & "<div style='margin:26px 0;'>"
    h = h & "<a href='#' style='display:inline-block;background:" & C_AZUL & ";background:linear-gradient(135deg," & C_AZUL & "," & C_ROXO & ");color:#FFFFFF;text-decoration:none;font-weight:600;font-size:15px;padding:13px 26px;border-radius:10px;font-family:Segoe UI,Arial,sans-serif;'>&#128073; Clique aqui</a>"
    h = h & "</div>"
    h = h & "<div style='font-size:13px;color:" & C_SUAVE & ";font-family:Segoe UI,Arial,sans-serif;'>Dica: para movimento de verdade, use um GIF animado (as animacoes de CSS nao funcionam na maioria dos leitores de e-mail).</div>"
    h = h & "</td></tr>"

    ' Faixa colorida (arco-iris) - fallback azul solido
    h = h & "<tr><td style='height:5px;background:" & C_AZUL & ";background:linear-gradient(90deg," & C_AZUL & ",#22C55E,#F59E0B," & C_ROXO & ");font-size:0;line-height:0;'>&nbsp;</td></tr>"

    ' Rodape
    h = h & "<tr><td style='padding:18px 28px;background:" & C_RODAPE & ";border-radius:0 0 16px 16px;'>"
    h = h & "<div style='font-size:12px;color:#6B7280;font-family:Segoe UI,Arial,sans-serif;'>Enviado com &#128153; &#8226; Tema Escuro Animado</div>"
    h = h & "</td></tr>"

    h = h & "</table></td></tr></table></div>"

    TemplateEscuro = h
End Function
