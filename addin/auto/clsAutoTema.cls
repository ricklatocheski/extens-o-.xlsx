VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAutoTema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'============================================================
' clsAutoTema  -  Modulo de Classe
'------------------------------------------------------------
' Escuta o evento "abrir pasta de trabalho" de TODO o Excel e
' aplica o tema escuro automaticamente em cada planilha do
' arquivo que for aberto.
'
' Instale dentro do Personal.xlsb (veja addin/auto/INSTALAR_AUTO.md).
'============================================================
Option Explicit

Public WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In Wb.Worksheets
        AplicarTemaEscuroEm ws
    Next ws
End Sub
