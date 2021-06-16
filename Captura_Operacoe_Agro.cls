VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 Public Chave As String, Senha As String
 
 Public Sub CapturaOperacoes()

    If Not Conectar Then Exit Sub
    Dim linha As Long
    Dim linha1 As Long
    Dim ok As Integer
    Dim linha3 As Long
    Dim linha4 As Long
    Dim linha5 As Long
    Dim linha6 As Long
    Dim GrupAgro As Boolean
    Dim pagcop As Integer
    Dim PPAA As Boolean
    Dim Pdihip As Double
    Dim Gdihip As Double
    
    
      ok = Aguardar(1, 2, "BB30")
    
    If Chave = "" Then
        Chave = ChaveSenha(1)
        Senha = ChaveSenha(2)
    End If

DoEvents
    
    ok = cursor(15, 14)
    Teclar ("clientes")
    
    ok = cursor(16, 14)
    Teclar (Senha)
    
    ENTER
    Esperasystem

    
    
    linha = 2
        
        While linha < 101
        
           
            Range("B" & linha) = del
            Range("C" & linha) = del
            Range("D" & linha) = del
            Range("E" & linha) = del
            Range("F" & linha) = del
            Range("F" & linha) = del
            Range("G" & linha) = del
            Range("H" & linha) = del
            Range("I" & linha) = del
            Range("J" & linha) = del
            Range("K" & linha) = del
            Range("L" & linha) = del
            Range("O" & linha) = del
            Range("P" & linha) = del
            Range("Q" & linha) = del
            Range("R" & linha) = del
       
        linha = linha + 1
        Wend
    
    
 

    ok = Aguardar(1, 2, "SBBP6130")
    Esperasystem
    
    TeclarTxt "27", 21, 20
    ENTER
    Esperasystem
    
    TeclarTxt "03", 21, 20
    ENTER
    Esperasystem
    
    TeclarTxt "11", 19, 24
    ENTER
    Esperasystem
    TeclarTxt "14", 19, 27
    TeclarTxt "a", 19, 43
    ENTER
    Esperasystem
    TeclarTxt "11", 20, 24
    ENTER
    Esperasystem
    
    ' 14 a
    
    linha = 10
    linha1 = 2
        
        If Trim(Copiar(23, 11, 3)) = "Nao" Then
        F3
        Esperasystem
        linha = 22 'não passar no loop abaixo se caso nao tiver operacao
        Esperasystem
        End If
 

            
        Do While linha < 20
        
            If Trim(Copiar(linha, 5, 30)) <> "" Then
            
            TeclarTxt "x", linha, 3
            ENTER
            Esperasystem
                If Trim(Copiar(17, 22, 5)) = "ENTER" Then
                ENTER
                Esperasystem
                End If
                
                If Trim(Copiar(11, 18, 6)) = "Propos" Then
                F3
                Esperasystem
                End If
                
                F4
            Esperasystem
               If Trim(Copiar(11, 18, 8)) = "Proposta" Then
                F3
                Esperasystem
                End If
            
            linha5 = 15
                 
                                        Do While linha5 < 20
                                            
                                            If Trim(Copiar(linha5, 70, 1)) <> "_" Then
                                            
                                                If Trim(Copiar(linha5, 70, 1)) = "S" Then
                                                GoTo proximo3
                                                
                                                End If
                                                
                                                If Trim(Copiar(linha5, 70, 1)) = "N" Then
                                                GoTo proximo3
                                                End If
                        
                                            End If
                                            
                                                linha5 = linha5 + 1
                                                
                                                If linha5 > 19 Then
                                                F8
                                                Esperasystem
                                                F8
                                                Esperasystem
                                                    linha6 = 15
                                                          Do While linha6 < 20
                                                          If Trim(Copiar(linha6, 70, 1)) <> "_" Then
                                            
                                                                If Trim(Copiar(linha6, 70, 1)) = "S" Then
                                                                GoTo proximo3
                                                                End If
                                                
                                                                If Trim(Copiar(linha6, 70, 1)) = "N" Then
                                                                GoTo proximo3
                                                                End If
                                                           End If
                                                           linha6 = linha6 + 1
                                                         Loop
                                                    F8
                                                    Esperasystem
                                                    If Trim(Copiar(22, 3, 6)) = "Ultima" Then
                                                    Range("D" & linha1) = (Replace(Trim(Copiar(8, 23, 11)), ".", "")) 'numero da proposta
                                                    Range("G" & linha1) = (Trim(Copiar(9, 23, 22)))  'valor da operação
                                                    Range("J" & linha1) = "14a" 'recebe a opção do cop
                                                    linha1 = linha1 + 1
                                                    GoTo proximo3
                                                    
                                    
                                          
                                                    End If

                                                End If
                                     
                                                
                                        Loop
proximo3:
While (Trim(Copiar(1, 3, 8))) <> "COPM2422"
F3
Esperasystem
Wend
TeclarTxt "11", 20, 24
ENTER
Esperasystem
'linha = linha + 1
            If linha > 19 Then
            F8
            Esperasystem
            linha = linha + 1
            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
            F3
            Esperasystem
            End If
        
    End If
    End If
    linha = linha + 1
  
Loop

        While (Trim(Copiar(1, 3, 8))) <> "COPM1720"
        F3
        Esperasystem
        Wend
        ok = Aguardar(1, 3, "COPM1720")
        
        
        
'14 b

    TeclarTxt "14", 19, 27
    TeclarTxt "b", 19, 43
    ENTER
    Esperasystem
    TeclarTxt "11", 20, 24
    ENTER
    Esperasystem
    linha = 10
        
        If Trim(Copiar(23, 11, 3)) = "Nao" Then
        F3
        Esperasystem
        linha = 22 'não passar no loop abaixo
        Esperasystem
        End If
 
            
        Do While linha < 20
        
        If Trim(Copiar(linha, 5, 30)) <> "" Then
            
            TeclarTxt "x", linha, 3
            ENTER
            Esperasystem
                If Trim(Copiar(17, 22, 5)) = "ENTER" Then
                ENTER
                Esperasystem
                End If
                
           F4
            Esperasystem
                If Trim(Copiar(11, 18, 8)) = "Proposta" Then
                F3
                Esperasystem
                End If
            linha5 = 15
                 
                                        Do While linha5 < 20
                                            
                                            If Trim(Copiar(linha5, 70, 1)) <> "_" Then
                                            
                                                If Trim(Copiar(linha5, 70, 1)) = "S" Then
                                                GoTo proximo5
                                                
                                                End If
                                                
                                                If Trim(Copiar(linha5, 70, 1)) = "N" Then
                                                GoTo proximo5
                                                End If
                        
                                            End If
                                            
                                                linha5 = linha5 + 1
                                                
                                                If linha5 > 19 Then
                                                F8
                                                Esperasystem
                                                F8
                                                Esperasystem
                                                    linha6 = 15
                                                          Do While linha6 < 20
                                                          If Trim(Copiar(linha6, 70, 1)) <> "_" Then
                                            
                                                                If Trim(Copiar(linha6, 70, 1)) = "S" Then
                                                                GoTo proximo5
                                                                End If
                                                
                                                                If Trim(Copiar(linha6, 70, 1)) = "N" Then
                                                                GoTo proximo5
                                                                End If
                                                           End If
                                                           linha6 = linha6 + 1
                                                         Loop
                                                    F8
                                                    Esperasystem
                                                    If Trim(Copiar(22, 3, 6)) = "Ultima" Then
                                                    Range("D" & linha1) = Replace(Trim(Copiar(8, 23, 11)), ".", "") 'numero da proposta
                                                    Range("G" & linha1) = Trim(Copiar(9, 23, 22)) 'valor da operação
                                                    Range("J" & linha1) = "14b" 'recebe a opção do cop
                                                    linha1 = linha1 + 1
                                                    linha = linha + 1
                                                    GoTo proximo5
                                                    
                                    
                                          
                                                    End If

                                                End If
                                     
                                                
                                        Loop
proximo5:
While (Trim(Copiar(1, 3, 8))) <> "COPM2422"
F3
Esperasystem
Wend
TeclarTxt "11", 20, 24
ENTER
Esperasystem
'linha = linha + 1
            If linha > 19 Then
            F8
            Esperasystem
            linha = linha + 1
            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
            F3
            Esperasystem
            End If
        
    End If
    End If

linha = linha + 1
Loop

 While (Trim(Copiar(1, 3, 8))) <> "COPM1720"
        F3
        Esperasystem
        Wend
        ok = Aguardar(1, 3, "COPM1720")


'15a

    TeclarTxt "15", 19, 27
    TeclarTxt "a", 19, 43
    ENTER
    Esperasystem
    TeclarTxt "11", 20, 24
    ENTER
    Esperasystem
    
    
    linha = 10
    pagcop = "001"
   
            
        Do While linha < 20
        
        If Trim(Copiar(linha, 5, 30)) <> "" Then
            
            TeclarTxt "x", linha, 3
            ENTER
            Esperasystem
                If Trim(Copiar(17, 22, 5)) = "ENTER" Then
                ENTER
                Esperasystem
                End If
                
            F4
            Esperasystem
               If Trim(Copiar(11, 18, 8)) = "Proposta" Then
                F3
                Esperasystem
                End If

            
            linha5 = 15
                 
                                        Do While linha5 < 20
                                            
                                            If Trim(Copiar(linha5, 70, 1)) <> "_" Then
                                            
                                                If Trim(Copiar(linha5, 70, 1)) = "S" Then
                                                GoTo proximo4
                                                
                                                End If
                                                
                                                If Trim(Copiar(linha5, 70, 1)) = "N" Then
                                                GoTo proximo4
                                                End If
                        
                                            End If
                                            
                                                linha5 = linha5 + 1
                                                
                                                If linha5 > 19 Then
                                                F8
                                                Esperasystem
                                                F8
                                                Esperasystem
                                                    linha6 = 15
                                                          Do While linha6 < 20
                                                          If Trim(Copiar(linha6, 70, 1)) <> "_" Then
                                            
                                                                If Trim(Copiar(linha6, 70, 1)) = "S" Then
                                                                GoTo proximo4
                                                                End If
                                                
                                                                If Trim(Copiar(linha6, 70, 1)) = "N" Then
                                                                GoTo proximo4
                                                                End If
                                                           End If
                                                           linha6 = linha6 + 1
                                                         Loop
                                                    
                                                    F8
                                                    Esperasystem
                                                    F8
                                                    Esperasystem
                                                    If Trim(Copiar(22, 3, 6)) = "Ultima" Then
                                                    Range("D" & linha1) = Replace(Trim(Copiar(8, 23, 11)), ".", "") 'numero da proposta
                                                    Range("G" & linha1) = Trim(Copiar(9, 23, 22)) 'valor da operação
                                                    Range("J" & linha1) = "15a" 'recebe a opção do cop
                                                    linha1 = linha1 + 1
                                                    'linha = linha + 1
                                                    GoTo proximo4
                                                    
                                    
                                          
                                                    End If

                                                End If
                                     
                                                
                                        Loop
proximo4:
While (Trim(Copiar(1, 3, 8))) <> "COPM2403"
F3
Esperasystem
Wend
TeclarTxt "11", 20, 24
ENTER
Esperasystem
     While (Trim(Copiar(3, 74, 3))) <> pagcop
        F8
        Esperasystem
        Wend
        'ok = Aguardar(3, 3, pagcop)
    
    
    

Esperasystem

'linha = linha + 1
        If linha = 19 Then
        F8
        Esperasystem
        pagcop = Trim(Copiar(3, 74, 3))
        Esperasystem
        linha = 9
            
            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
            F3
            Esperasystem
            linha = 21
            End If
        
        End If
        
        If Trim(Copiar(linha, 5, 30)) = "" Then
         F3
         Esperasystem
         F3
         Esperasystem
        End If
    End If

    linha = linha + 1
Loop

 While (Trim(Copiar(1, 3, 8))) <> "COPM1720"
        F3
        Esperasystem
        Wend
        ok = Aguardar(1, 3, "COPM1720")
 
'15b

    TeclarTxt "15", 19, 27
    TeclarTxt "b", 19, 43
    ENTER
    Esperasystem
    TeclarTxt "11", 20, 24
    ENTER
    Esperasystem
    
    
    linha = 10
   
            If Trim(Copiar(23, 11, 3)) = "Nao" Then
            F3
            Esperasystem
            linha = 22 'não passar no loop abaixo
            Esperasystem
            End If
 
   
   
            
        Do While linha < 20
        
        If Trim(Copiar(linha, 5, 30)) <> "" Then
            
            TeclarTxt "x", linha, 3
            ENTER
            Esperasystem
                If Trim(Copiar(17, 22, 5)) = "ENTER" Then
                ENTER
                Esperasystem
                End If
                
              F4
            Esperasystem
               If Trim(Copiar(11, 18, 8)) = "Proposta" Then
                F3
                Esperasystem
                End If
            
            linha5 = 15
                 
                                        Do While linha5 < 20
                                            
                                            If Trim(Copiar(linha5, 70, 1)) <> "_" Then
                                            
                                                If Trim(Copiar(linha5, 70, 1)) = "S" Then
                                                GoTo proximo6
                                                
                                                End If
                                                
                                                If Trim(Copiar(linha5, 70, 1)) = "N" Then
                                                GoTo proximo6
                                                End If
                        
                                            End If
                                            
                                                linha5 = linha5 + 1
                                                
                                                If linha5 > 19 Then
                                                F8
                                                Esperasystem
                                                F8
                                                Esperasystem
                                                    linha6 = 15
                                                          Do While linha6 < 20
                                                          If Trim(Copiar(linha6, 70, 1)) <> "_" Then
                                            
                                                                If Trim(Copiar(linha6, 70, 1)) = "S" Then
                                                                GoTo proximo6
                                                    
                                                                End If
                                                
                                                                If Trim(Copiar(linha6, 70, 1)) = "N" Then
                                                                GoTo proximo6
                                                    
                                                                End If
                                                           End If
                                                           linha6 = linha6 + 1
                                                         Loop
                                                    F8
                                                    Esperasystem
                                                    F8
                                                    Esperasystem
                                                    If Trim(Copiar(22, 3, 6)) = "Ultima" Then
                                                    Range("D" & linha1) = Replace(Trim(Copiar(8, 23, 11)), ".", "") 'numero da proposta
                                                    Range("G" & linha1) = Trim(Copiar(9, 23, 22)) 'valor da operação
                                                    Range("J" & linha1) = "15b" 'recebe a opção do cop
                                                    linha1 = linha1 + 1
                                                    linha = linha + 1
                                                    GoTo proximo6
                                                    
                                    
                                          
                                                    End If

                                                End If
                                     
                                                
                                        Loop
proximo6:
While (Trim(Copiar(1, 3, 8))) <> "COPM2403"
F3
Esperasystem
Wend
TeclarTxt "11", 20, 24
ENTER
Esperasystem
'linha = linha + 1
            If linha > 19 Then
            F8
            Esperasystem
            linha = linha + 1
            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
            F3
            Esperasystem
            linha = 21
            End If
        
    End If
        
        If Trim(Copiar(linha, 3, 1)) = "" Then
         F3
         Esperasystem
         F3
         Esperasystem
        End If
    End If
    


    linha = linha + 1
Loop

 'While (Trim(Copiar(1, 3, 8))) <> "COPM1720"
  '      F3
   '     Esperasystem
    '    Wend
     '   ok = Aguardar(1, 3, "COPM1720")

'15e

   ' TeclarTxt "15", 19, 27
    'TeclarTxt "e", 19, 43
    'ENTER
    'Esperasystem
    'TeclarTxt "11", 20, 24
    'ENTER
   ' Esperasystem
    
    
    'linha = 10
    
     '   If Trim(Copiar(23, 11, 3)) = "Nao" Then
      '  F3
       ' linha = 22 'não passar no loop abaixo
        'Esperasystem
 '       End If
 
    
  '      If Trim(Copiar(23, 11, 3)) = "Nao" Then
   '     F3
    '    linha = 22 'não passar no loop abaixo
     '   Esperasystem
     '   End If
 
   
            
       ' Do While linha < 20
        
        'If Trim(Copiar(linha, 5, 30)) <> "" Then
            
           ' TeclarTxt "x", linha, 3
           ' ENTER
           ' Esperasystem
           '     If Trim(Copiar(17, 22, 5)) = "ENTER" Then
               ' ENTER
               ' Esperasystem
               ' End If
                
              ' F4
            'Esperasystem
              ' If Trim(Copiar(11, 18, 8)) = "Proposta" Then
               ' F3
               ' Esperasystem
                ''End If
            
            'linha5 = 15
                 
                                        'Do While linha5 < 20
                                            
                                           ' If Trim(Copiar(linha5, 70, 1)) <> "_" Then
                                            
                                               ' If Trim(Copiar(linha5, 70, 1)) = "S" Then
                                              '  GoTo proximo7
                                                
                                              '  End If
                                                
                                              '  If Trim(Copiar(linha5, 70, 1)) = "N" Then
                                              '  GoTo proximo7
                                             '   End If
                        
                                           ' End If
                                            
                                               ' linha5 = linha5 + 1
                                                
                                              '  If linha5 > 19 Then
                                              '  F8
                                              '  F8
                                               ' Esperasystem
                                                   ' linha6 = 15
                                                          'Do While linha6 < 20
                                                         ' If Trim(Copiar(linha6, 70, 1)) <> "_" Then
                                            
                                                              '  If Trim(Copiar(linha6, 70, 1)) = "S" Then
                                                              '  GoTo proximo7
                                                    
                                                              '  End If
                                                
                                                              '  If Trim(Copiar(linha6, 70, 1)) = "N" Then
                                                              '  GoTo proximo7
                                                    
                                                               ' End If
                                                          ' End If
                                                          ' linha6 = linha6 + 1
                                                        ' Loop
                                                    
                                                   ' If Trim(Copiar(22, 3, 6)) = "Ultima" Then
                                                   ' Range("D" & linha1) = Replace(Trim(Copiar(8, 23, 11)), ".", "") 'numero da proposta
                                                   ' Range("G" & linha1) = Trim(Copiar(9, 23, 22)) 'valor da operação
                                                   ' Range("J" & linha1) = "15e" 'recebe a opção do cop
                                                  '  linha1 = linha1 + 1
                                                  '  linha = linha + 1
                                                  '  GoTo proximo7
                                                    
                                    
                                          
                                                    'End If

                                                'End If
                                     
                                                
                                        'Loop
'proximo7:
'While (Trim(Copiar(1, 3, 8))) <> "COPM2403"
'F3
'Esperasystem
'Wend
'TeclarTxt "11", 20, 24
'ENTER
'Esperasystem

            'If linha > 19 Then
           ' F8
           ' Esperasystem
           ' linha = linha + 1
           ' If Trim(Copiar(23, 3, 6)) = "Ultima" Then
           ' F3
           ' End If
        
   ' End If
        
       ' If Trim(Copiar(linha, 3, 1)) = "" Then
        ' F3
        ' F3
        'End If
  '  End If
    


    'linha = linha + 1

'Loop

 While (Trim(Copiar(1, 3, 8))) <> "COPM0000"
        F3
        Esperasystem
        Wend
        
        ok = Aguardar(1, 3, "COPM0000")


linha1 = 2
    
    Do While Range("D" & linha1) <> 0
        
        TeclarTxt "13", 19, 24
        ENTER
        Esperasystem
        TeclarTxt Range("A" & linha1), 20, 24 'a celula n define o prefixo da agencia com base no numero da proposta
        Esperasystem
        ENTER
        Esperasystem
        TeclarTxt Range("D" & linha1), 7, 50
        ENTER
        Esperasystem
        
            linha = 10
   
            While Trim(Copiar(linha, 5, 6)) <> "SUMULA"
            linha = linha + 1
            Wend
            Esperasystem
            TeclarTxt "x", linha, 3
            Esperasystem
            ENTER
            Esperasystem
            
            Range("B" & linha1) = Replace(Trim(Copiar(14, 30, 11)), ".", "") 'captura o numero do mci do cliente
            Range("E" & linha1) = Trim(Copiar(14, 44, 35)) ' nome do proponente
            Range("C" & linha1) = Replace(Replace(Trim(Copiar(15, 30, 14)), ".", ""), "-", "")  ' CPF do Proponente
            Esperasystem
            F8
            Esperasystem
            
                If Trim(Copiar(13, 24, 5)) = "-----" Then
                Range("F" & linha1) = Trim(Copiar(19, 24, 40)) 'captura a linha de credito
                Range("K" & linha1) = Trim(Copiar(11, 48, 4))  'captura agencia do cliente
                Range("L" & linha1) = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(Copiar(11, 55, 15)), ".", ""), "-X", ""), "-0", ""), "-1", ""), "-2", ""), "-3", ""), "-4", ""), "-5", ""), "-6", ""), "-7", ""), "-8", ""), "-9", "")
                End If
                
                If Trim(Copiar(13, 24, 5)) <> "-----" Then
                Range("F" & linha1) = Trim(Copiar(13, 24, 40)) 'captura a linha de credito
                Range("K" & linha1) = Trim(Copiar(5, 48, 4))  'captura agencia do cliente
                Range("L" & linha1) = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(Copiar(5, 55, 15)), ".", ""), "-X", ""), "-0", ""), "-1", ""), "-2", ""), "-3", ""), "-4", ""), "-5", ""), "-6", ""), "-7", ""), "-8", ""), "-9", "")
                End If
                
                If Trim(Copiar(14, 3, 4)) = "2.1." Then
                Range("F" & linha1) = Trim(Copiar(16, 24, 40)) 'captura a linha de credito
                Range("K" & linha1) = Trim(Copiar(8, 48, 4))  'captura agencia do cliente
                Range("L" & linha1) = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(Copiar(8, 55, 15)), ".", ""), "-X", ""), "-0", ""), "-1", ""), "-2", ""), "-3", ""), "-4", ""), "-5", ""), "-6", ""), "-7", ""), "-8", ""), "-9", "")
                End If
                
                If Trim(Copiar(10, 3, 4)) = "1.3." Then
                
                Range("K" & linha1) = Trim(Copiar(14, 48, 4))  'captura agencia do cliente
                Range("L" & linha1) = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Trim(Copiar(14, 55, 15)), ".", ""), "-X", ""), "-0", ""), "-1", ""), "-2", ""), "-3", ""), "-4", ""), "-5", ""), "-6", ""), "-7", ""), "-8", ""), "-9", "")
                F8
                Esperasystem
                Range("F" & linha1) = Trim(Copiar(6, 24, 40))  'captura a linha de credito
                End If
                
                
            Esperasystem
            linha1 = linha1 + 1
            F3
            Esperasystem
            F3
            Esperasystem
            F3
            Esperasystem
    Loop
    
 
    
    
      linha1 = 2
    
      Do While Range("D" & linha1) <> 0
      
           
            
            TeclarTxt "25", 19, 24
            TeclarTxt WorksheetFunction.Text(Range("A" & linha1), "0000"), 20, 24
            ENTER
            Esperasystem
            TeclarTxt "11", 21, 26
            ENTER
            Esperasystem
            TeclarTxt WorksheetFunction.Text(Range("D" & linha1), "000000000"), 9, 46
            ENTER
            Esperasystem
            
              linha4 = 12
                                
            Do While linha4 < 21
                If Trim(Copiar(linha4, 5, 34)) <> Range("N" & linha1) Then
                   linha4 = linha4 + 1
                        If Trim(Copiar(linha4, 5, 34)) = Range("N" & linha1) Then
                        TeclarTxt "D", linha4, 3
                        ENTER
                        Esperasystem
                        Range("H" & linha1) = Replace(Trim(Copiar(15, 39, 21)), " -", "") 'pegando a data de encerramento do parecer na agencia
                        linha1 = linha1 + 1
                        GoTo Proximo28
                        End If
                End If
                                        
                        If linha4 > 20 Then
                        F8
                        Esperasystem
                            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
                            F3
                            Esperasystem
                            ENTER
                            Esperasystem
                            End If
                            linha4 = 12
                         End If
                                    
            Loop
                                
Proximo28:
        While (Trim(Copiar(1, 3, 8))) <> "COPM0000"
        F3
        Esperasystem
        Wend
'ENTER
'Esperasystem
                                
'                                linha4 = 12
'
'                                Do While linha4 < 21
'                                        If Trim(Copiar(linha4, 5, 34)) <> "EMISSAO DE DESPACHO           8515" Then
'                                        linha4 = linha4 + 1
'                                            If Trim(Copiar(linha4, 5, 34)) = "EMISSAO DE DESPACHO           8515" Then
'                                            TeclarTxt "D", linha4, 3
'                                            ENTER
'                                            Esperasystem
'                                            Range("U" & linha1) = Replace(Trim(Copiar(14, 39, 21)), " -", "")
'                                            F3
'                                            F3
'                                            ENTER
'                                            GoTo Proximo28
'                                            End If
'                                        End If
                                        
'                                        If linha4 > 20 Then
'                                        F8
'                                            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
'                                            F3
'                                            F3
'                                            F3
                                            'GoTo proximo29
'                                            End If
'                                            linha4 = 12
'                                        End If
 
'                                Loop

            
'    linha1 = linha + 1
    
    
    
    Loop
        While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
        F3
        Esperasystem
        Wend
    
    'Pegar informação se há grupo agropecuario.

    linha1 = 2
    
    ok = Aguardar(1, 2, "SBBP6130")

    TeclarTxt "01", 21, 20
    ENTER
    Esperasystem
            If Trim(Copiar(16, 16, 2)) = "F3" Then
            F3
            Esperasystem
            End If
            
    TeclarTxt "01", 19, 18
    ENTER
    Esperasystem
            
            Do While Range("D" & linha1) <> 0
            
                TeclarTxt WorksheetFunction.Text(Range("B" & linha1), "000000000"), 20, 57
                ENTER
                Esperasystem
                    If Trim(Copiar(9, 12, 19)) = "Informação restrita" Then
                    TeclarTxt "Analise de operacao Rural " & Range("B" & linha1), 11, 12
                    ENTER
                    Esperasystem
                    End If
        
                TeclarTxt "04", 21, 41
                ENTER
                Esperasystem
            
                TeclarTxt "5", 21, 26
                ENTER
                Esperasystem
            
                    If Trim(Copiar(10, 7, 9)) <> "" Then
                    GrupAgro = True
                        If GrupAgro = True Then
                        Range("I" & linha1) = "Sim"
                        End If
                    End If
                
                    If Trim(Copiar(10, 7, 9)) = "" Then
                    GrupAgro = False
                        If GrupAgro = False Then
                        Range("I" & linha1) = "Não"
                        End If
                    End If
                    linha1 = linha1 + 1
            
                    While (Trim(Copiar(1, 3, 8))) <> "MCIM001A"
                    F3
                    Esperasystem
                    Wend
                     'If Trim(Copiar(16, 16, 2)) = "F3" Then
                     'F3
                     'Esperasystem
                     'End If
    Loop
    
    While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
        F3
        Esperasystem
        Wend
    
        Range("B102") = Trim(Copiar(1, 70, 10))  'data da captura
        Range("C102") = Trim(Copiar(2, 72, 8)) ' captura hora
       
        
        While (Trim(Copiar(1, 2, 4))) <> "BB30"
        F5
        Esperasystem
        Wend
        
        
            DoEvents
    
    ok = cursor(15, 14)
    Teclar ("deb")
    
    ok = cursor(16, 14)
    Teclar (Senha)
    
    ENTER
    Esperasystem
        
        
        
         linha1 = 2
        
        Do While Range("K" & linha1) <> 0
        
        
               TeclarTxt "40", 21, 11
                ENTER
                Esperasystem
                TeclarTxt "03", 18, 16
                TeclarTxt WorksheetFunction.Text(Range("K" & linha1), "0000"), 20, 16
                TeclarTxt WorksheetFunction.Text(Range("L" & linha1), "0000000000"), 21, 16
                ENTER
                Esperasystem
                
                
                    
                If Trim(Copiar(8, 20, 30)) = "Conta sem aplicação automática" Then
                PPAA = False
                    If PPAA = False Then
                    Range("R" & linha1) = "Não"
                    End If
                End If
                If Trim(Copiar(8, 20, 30)) <> "Conta sem aplicação automática" Then
                PPAA = True
                    If PPAA = True Then
                    Range("R" & linha1) = "Sim"
                    End If
                End If
                
                While (Trim(Copiar(1, 3, 8))) <> "DEBM0050"
                F3
                Esperasystem
                Wend
    
                linha1 = linha1 + 1
                
        Loop
  
        While (Trim(Copiar(1, 2, 4))) <> "BB30"
        F5
        Esperasystem
        Wend
              
              
      
      
             
              
  
  MsgBox "Captura de Operações Finalizada"
  
  


End Sub





