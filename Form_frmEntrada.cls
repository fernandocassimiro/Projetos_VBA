VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub botaoIncluir_Click()


Dim data As DAO.Recordset
Set data = CurrentDb.OpenRecordset("Tbl - Data")

        
'bkp.Index = "PrimaryKey"
'bkp.Seek "=", Forms!frmIncluirProposta!pesquisa.Value 'FORMATA MCI PARA TEXTO E PESQUISA NA TABELA SE JÁ HÁ O REGISTRO
        
    data.AddNew
    data!dia1 = Forms!frmEntrada!dia1.Value
    data!mes1 = Forms!frmEntrada!mes1.Value
    data!ano1 = Forms!frmEntrada!ano1.Value
    data!dia2 = Forms!frmEntrada!dia2.Value
    data!Mes2 = Forms!frmEntrada!Mes2.Value
    data!Ano2 = Forms!frmEntrada!Ano2.Value
    data.Update
    Call FN_LimpaCampos(Me)
    

    Call GravaTaxa1


End Sub
