Attribute VB_Name = "AutoTemaEscuro"
'============================================================
' AutoTemaEscuro  -  Modulo padrao
'------------------------------------------------------------
' Ativa o tema escuro automatico para TODOS os arquivos que
' voce abrir no Excel. Funciona em conjunto com a classe
' clsAutoTema. Instale ambos dentro do Personal.xlsb.
'
' Passo a passo completo: addin/auto/INSTALAR_AUTO.md
'============================================================
Option Explicit

' Instancia da classe que escuta os eventos do Excel.
Dim gTema As New clsAutoTema

'------------------------------------------------------------
' Executado automaticamente quando o Personal.xlsb carrega
' (ou seja, toda vez que o Excel abre).
'------------------------------------------------------------
Public Sub Auto_Open()
    Set gTema.App = Application
End Sub

'------------------------------------------------------------
' Aplica o tema escuro moderno em uma planilha.
'------------------------------------------------------------
Public Sub AplicarTemaEscuroEm(ByVal ws As Worksheet)
    On Error Resume Next

    ' Fundo, fonte e cor do texto
    ws.Cells.Interior.Color = RGB(24, 25, 29)
    ws.Cells.Font.Color = RGB(228, 230, 235)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Tab.Color = RGB(59, 130, 246)

    Dim rng As Range
    Set rng = ws.UsedRange
    If rng Is Nothing Then Exit Sub

    ' Linhas entre as celulas (bordas claras)
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(120, 120, 120)
        .Weight = xlThin
    End With

    ' Zebra (linhas alternadas) nas linhas de dados
    rng.FormatConditions.Delete
    If rng.Rows.Count > 1 Then
        Dim dados As Range
        Set dados = rng.Resize(rng.Rows.Count - 1).Offset(1, 0)
        dados.FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0"
        dados.FormatConditions(dados.FormatConditions.Count).Interior.Color = RGB(31, 33, 39)
    End If

    ' Cabecalho com borda de acento
    With rng.Rows(1)
        .Interior.Color = RGB(33, 37, 48)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .RowHeight = 22
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(59, 130, 246)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

'------------------------------------------------------------
' Aplica manualmente na planilha ativa (atalho opcional).
'------------------------------------------------------------
Public Sub AplicarAgora()
    If Not ActiveSheet Is Nothing Then AplicarTemaEscuroEm ActiveSheet
End Sub
