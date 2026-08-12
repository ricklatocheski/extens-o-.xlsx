Attribute VB_Name = "DarkThemeAddin"
'============================================================
' Suplemento Tema Escuro (Moderno) para Excel 2007+ / 2019
'------------------------------------------------------------
' Aplica um tema escuro moderno as planilhas do Excel:
'   - fundo grafite levemente azulado
'   - zebra (linhas alternadas) via formatacao condicional,
'     rapida mesmo em tabelas grandes
'   - texto off-white suave (mais confortavel que branco puro)
'   - cabecalho com borda de acento e congelamento (freeze)
'   - grade sutil e cor de aba de acento
'
' Compativel com Excel 2007, 2010, 2013, 2016, 2019 e 365.
' No Excel 2019/365, combine com Arquivo > Opcoes > Geral >
' Tema do Office = "Preto" para escurecer tambem a interface.
'============================================================

Option Explicit

'------------------------------------------------------------
' Paleta moderna (mesmas cores do DarkTheme.thmx e da demo)
'------------------------------------------------------------
Private Function corFundo() As Long:      corFundo = RGB(24, 25, 29):      End Function    ' #18191D
Private Function corFundoAlt() As Long:   corFundoAlt = RGB(31, 33, 39):   End Function    ' #1F2127
Private Function corCabecalho() As Long:  corCabecalho = RGB(33, 37, 48):  End Function    ' #212530
Private Function corTexto() As Long:      corTexto = RGB(228, 230, 235):   End Function    ' #E4E6EB
Private Function corTextoForte() As Long: corTextoForte = RGB(255, 255, 255): End Function ' #FFFFFF
Private Function corAcento() As Long:     corAcento = RGB(59, 130, 246):   End Function    ' #3B82F6
Private Function corGrade() As Long:      corGrade = RGB(42, 46, 56):      End Function    ' #2A2E38

Private Const FONTE As String = "Segoe UI"   ' cai para Calibri se nao existir

'------------------------------------------------------------
' Aplica o tema escuro moderno a planilha ativa.
'------------------------------------------------------------
Public Sub AplicarTemaEscuro()
    On Error GoTo Falha
    If ActiveSheet Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Activate

    Application.ScreenUpdating = False

    ' --- Fundo, fonte e grade de toda a planilha ---
    With ws.Cells
        .Interior.Color = corFundo
        .Font.Color = corTexto
        .Font.Name = FONTE
        .Font.Size = 10
    End With

    ActiveWindow.DisplayGridlines = True
    ActiveWindow.GridlineColor = corGrade
    ws.Tab.Color = corAcento

    Dim rng As Range
    Set rng = ws.UsedRange
    If rng Is Nothing Then GoTo Fim
    If rng.Cells.Count = 0 Then GoTo Fim

    ' --- Zebra via formatacao condicional (rapido em tabelas grandes) ---
    ' Aplicada apenas as linhas de dados (o cabecalho fica de fora).
    ' A formula do modelo VBA usa sempre ROW()/virgula (padrao EUA),
    ' independentemente do idioma do Excel.
    rng.FormatConditions.Delete
    If rng.Rows.Count > 1 Then
        Dim dados As Range
        Set dados = rng.Resize(rng.Rows.Count - 1).Offset(1, 0)
        Dim fc As FormatCondition
        Set fc = dados.FormatConditions.Add(Type:=xlExpression, _
                 Formula1:="=MOD(ROW(),2)=0")
        fc.Interior.Color = corFundoAlt
    End If

    ' --- Cabecalho (primeira linha usada) ---
    With rng.Rows(1)
        .Interior.Color = corCabecalho
        .Font.Color = corTextoForte
        .Font.Bold = True
        .Font.Size = 10
        .RowHeight = 22
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = corAcento
            .Weight = xlMedium
        End With
    End With

    ' --- Congela a linha de cabecalho ---
    If rng.Rows.Count > 1 Then
        ws.Cells(rng.Row + 1, rng.Column).Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
    End If

Fim:
    ws.Cells(1, 1).Select
    Application.ScreenUpdating = True
    Exit Sub

Falha:
    Application.ScreenUpdating = True
    MsgBox "Nao foi possivel aplicar o tema escuro: " & Err.Description, _
           vbExclamation, "Tema Escuro"
End Sub

'------------------------------------------------------------
' Aplica o tema escuro em TODAS as planilhas da pasta ativa.
'------------------------------------------------------------
Public Sub AplicarTemaEscuroTodas()
    On Error GoTo Falha
    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim ws As Worksheet, atual As Worksheet
    Set atual = ActiveSheet
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        AplicarTemaEscuro
    Next ws
    atual.Activate
    Application.ScreenUpdating = True
    Exit Sub
Falha:
    Application.ScreenUpdating = True
    MsgBox "Erro ao aplicar em todas: " & Err.Description, vbExclamation
End Sub

'------------------------------------------------------------
' Restaura o visual claro padrao da planilha ativa.
'------------------------------------------------------------
Public Sub RemoverTemaEscuro()
    On Error GoTo Falha
    If ActiveSheet Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Activate
    Application.ScreenUpdating = False

    ws.Cells.FormatConditions.Delete
    With ws.Cells
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
    End With
    On Error Resume Next
    ws.UsedRange.Rows(1).Borders(xlEdgeBottom).LineStyle = xlNone
    On Error GoTo 0

    ActiveWindow.DisplayGridlines = True
    ActiveWindow.GridlineColorIndex = xlColorIndexAutomatic
    ActiveWindow.FreezePanes = False
    ws.Tab.ColorIndex = xlColorIndexNone

    Application.ScreenUpdating = True
    Exit Sub
Falha:
    Application.ScreenUpdating = True
    MsgBox "Nao foi possivel remover: " & Err.Description, vbExclamation
End Sub

'============================================================
' Callbacks da Faixa de Opcoes (customUI.xml)
'============================================================
Public Sub RibbonAplicar(ByVal control As Object):      AplicarTemaEscuro:      End Sub
Public Sub RibbonAplicarTodas(ByVal control As Object): AplicarTemaEscuroTodas: End Sub
Public Sub RibbonRemover(ByVal control As Object):      RemoverTemaEscuro:      End Sub
