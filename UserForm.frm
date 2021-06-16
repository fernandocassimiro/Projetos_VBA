VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Formulario Análise de Operação Agro "
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14250
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Filtro()

Dim W As Worksheet
Dim i As Integer
Dim ln As Integer
Set W = Sheets("Filtro_Operacoes")



i = 0
ln = 2



    For i = 0 To Me.ListBox1.ListCount - 1

    If Me.ListBox1.Selected(i) Then
     
     W.Cells(ln, 1).Value = Me.ListBox1.List(i, 0)
    
     
     ln = ln + 1

    End If
    
Next
    

End Sub

Public Sub Analisar_Click()
Planilha6.LimpaFiltro
Filtro
Plan1.capturaLinha
End Sub
Sub Classificar_Op_Data()
'
' Classificar_Op_Data Macro
'

'
    ActiveWorkbook.Worksheets("Captura_Operacoes").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Captura_Operacoes").Sort.SortFields.Add Key:=Range _
        ("H2:H100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Captura_Operacoes").Sort
        .SetRange Range("A1:T100")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Tela").Select
    ActiveWorkbook.Save
End Sub

Public Sub CommandButton1_Click()
Planilha4.CapturaOperacoes
Classificar_Op_Data

End Sub






Private Sub Label5_Click()

End Sub
