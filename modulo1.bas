Attribute VB_Name = "modulo1"
Option Compare Database


Public Sub GravaTaxa1()
If Not Conectar Then Exit Sub

 MsgBox "Logue no APLICATIVO CIC 50 e na tela INICIAL do sistema CDC"

Dim Dados As DAO.Recordset                                        '|
Set Dados = CurrentDb.OpenRecordset("Tbl - Dados")                '|
                                                                  '|    Declaração de variáveis
Dim data As DAO.Recordset                                         '|
Set data = CurrentDb.OpenRecordset("Tbl - Data")                  '|

    
   
    Dados.MoveFirst                                               '|    Captura o primeiro registro


    Do Until Dados.EOF
    
        TeclarTxt 22, 21, 20                                      '|
        ENTER                                                     '|
        Esperasystem                                              '|
        TeclarTxt 12, 19, 37                                      '|
        ENTER                                                     '|
        Esperasystem                                              '|
        TeclarTxt 1, 17, 30                                       '|
        TeclarTxt 2881, 18, 30                                    '|    Acessa o CDC-22-12-01
        TeclarTxt Dados!MCI, 19, 30                               '|
        ENTER                                                     '|
        Esperasystem                                              '|
        TeclarTxt Dados!Conv, 21, 30                              '|
        ENTER                                                     '|
        Esperasystem                                              '|
           
           If Trim(Copiar(23, 3, 4)) = "Conv" Then
 
           GoTo proximo01

           End If
        
        TeclarTxt 25, 9, 29                                       '|
        TeclarTxt Dados!PrazoMax, 9, 35                           '|
        TeclarTxt Replace(Dados!valor, ",", ""), 10, 29           '|    Grava dados
        TeclarTxt Replace(Dados!Teto, ",", ""), 11, 29            '|
        TeclarTxt 80, 12, 29                                      '|
        TeclarTxt 47, 13, 29                                      '|
        
        
        TeclarTxt data!dia1, 18, 29                               '|
        TeclarTxt data!mes1, 18, 34                               '|
        TeclarTxt data!ano1, 18, 39                               '|
        TeclarTxt data!dia2, 19, 29                               '|
        TeclarTxt data!Mes2, 19, 34                               '|    Inclui as datas inciais e finais
        TeclarTxt data!Ano2, 19, 39                               '|
        ENTER                                                     '|
        Esperasystem                                              '|
        F3                                                        '|
        Esperasystem
        
           If Trim(Copiar(23, 3, 4)) = "Cond" Then
 
           GoTo proximo01

           End If
                    
        
        TeclarTxt 25, 9, 25                                       '|
        TeclarTxt Dados!PrazoMax, 9, 36                           '|
        TeclarTxt Replace(Dados!Taxa, ",", ""), 9, 44             '|
        ENTER                                                     '|
        Esperasystem                                              '|
        TeclarTxt "S", 21, 26                                     '|
        ENTER                                                     '|
        Esperasystem                                              '|    Informa novamente prazo inicial, final e taxa
        ENTER                                                     '|
        Esperasystem
        
proximo01:
        
        While Trim(Copiar(1, 3, 8)) <> "CDCM0000"                 '|
        F3                                                        '|
        Esperasystem                                              '|
        Wend                                                      '|
       
        
        Dados.MoveNext                                            '|    Vai para captura do póximo registro
                                   
    Loop
    
    MsgBox "As taxas foram gravadas com sucesso!"
    
DoCmd.SetWarnings (False)
DoCmd.RunSQL "Delete * from [Tbl - Dados];"
DoCmd.RunSQL "Delete * from [Tbl - Data];"

End Sub


