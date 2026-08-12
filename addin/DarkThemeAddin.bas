Attribute VB_Name = "DarkThemeAddin"
'============================================================
' Suplemento Tema Escuro para Excel 2007 (.xlsx / .xlam)
'------------------------------------------------------------
' Aplica e remove um tema escuro nas planilhas do Excel.
' O tema escuro do Excel 2007 nao existe de forma nativa,
' entao este suplemento pinta o plano de fundo das celulas,
' ajusta a cor da fonte, desliga as linhas de grade e colore
' a aba da planilha para simular um modo escuro consistente.
'
' Compativel com Excel 2007 e versoes superiores.
'============================================================

Option Explicit

' --- Paleta do tema escuro (mesmas cores do DarkTheme.thmx) ---
Private Const CLR_FUNDO As Long = &H1E1E1E       ' RGB(30, 30, 30)   fundo principal
Private Const CLR_FUNDO_ALT As Long = &H302D2D   ' RGB(45, 45, 48)   faixas alternadas
Private Const CLR_TEXTO As Long = &HF2F2F2       ' RGB(242,242,242)  texto claro
Private Const CLR_CABECALHO As Long = &H2A2A2A   ' RGB(42, 42, 42)   cabecalho
Private Const CLR_ACENTO As Long = &HF7C34F      ' RGB(79,195,247)   azul de acento
Private Const CLR_BORDA As Long = &H3F3F3F       ' RGB(63, 63, 63)   bordas suaves

'------------------------------------------------------------
' Aplica o tema escuro a planilha ativa.
'------------------------------------------------------------
Public Sub AplicarTemaEscuro()
    On Error GoTo Falha
    If ActiveSheet Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False

    ' Fundo de toda a area usada + uma margem para novas celulas
    With ws.Cells
        .Interior.Color = CLR_FUNDO
        .Font.Color = CLR_TEXTO
    End With

    ' Desliga as linhas de grade e usa uma cor de grade escura
    ws.Parent.Windows(1).DisplayGridlines = False

    ' Colore a aba da planilha
    ws.Tab.Color = CLR_ACENTO

    ' Realca a linha de cabecalho (primeira linha usada)
    If ws.UsedRange.Rows.Count > 0 Then
        With ws.UsedRange.Rows(1)
            .Interior.Color = CLR_CABECALHO
            .Font.Color = CLR_ACENTO
            .Font.Bold = True
        End With
    End If

    Application.ScreenUpdating = True
    Exit Sub

Falha:
    Application.ScreenUpdating = True
    MsgBox "Nao foi possivel aplicar o tema escuro: " & Err.Description, _
           vbExclamation, "Tema Escuro"
End Sub

'------------------------------------------------------------
' Aplica o tema escuro a TODAS as planilhas da pasta ativa.
'------------------------------------------------------------
Public Sub AplicarTemaEscuroTodas()
    On Error GoTo Falha
    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Dim atual As Worksheet
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
    MsgBox "Erro ao aplicar em todas as planilhas: " & Err.Description, _
           vbExclamation, "Tema Escuro"
End Sub

'------------------------------------------------------------
' Restaura o visual claro padrao da planilha ativa.
'------------------------------------------------------------
Public Sub RemoverTemaEscuro()
    On Error GoTo Falha
    If ActiveSheet Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False

    With ws.Cells
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
    End With

    ws.Parent.Windows(1).DisplayGridlines = True
    ws.Tab.ColorIndex = xlColorIndexNone

    Application.ScreenUpdating = True
    Exit Sub

Falha:
    Application.ScreenUpdating = True
    MsgBox "Nao foi possivel remover o tema escuro: " & Err.Description, _
           vbExclamation, "Tema Escuro"
End Sub

'============================================================
' Callbacks da Faixa de Opcoes (Ribbon customUI.xml)
'============================================================
Public Sub RibbonAplicar(ByVal control As Object)
    AplicarTemaEscuro
End Sub

Public Sub RibbonAplicarTodas(ByVal control As Object)
    AplicarTemaEscuroTodas
End Sub

Public Sub RibbonRemover(ByVal control As Object)
    RemoverTemaEscuro
End Sub
