VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'////////////////////////////////////////////////////////////////////
'//////////////// ATENÇÃO! NÃO REMOVA ESTE COMENTÁRIO \\\\\\\\\\\\\\\
'Automação 5272 para captura propostas COP
'Desenvolvimento do Codigo iniciado por
'Desenvolvimento do Codigo continuado e terminado por Fernando Cassimiro, mat. F3229518-9
'Banco do Brasil S/A - Super PF II- Janeiro de 2020
'////////////////////////////////////////////////////////////////////

Public Sub capturaLinha()
    If Not Conectar Then Exit Sub
    Dim linha As Long
    Dim linha1 As Long
    Dim linha2 As Long
    Dim linha3 As Long
    Dim linha4 As Long
    Dim linha5 As Long
    Dim linha6 As Long
    Dim linha7 As Long
    Dim Ano As Long
    Dim cont As Long
    Dim comp As Long
    Dim comp2 As Long
    Dim vlr As Long
    Dim Impac As Long
    Dim SFNLP As Long
    Dim BBLP As Long
    Dim RBAagri As Long
    Dim RBApec As Long
    Dim BBTOT As Long
    Dim SFNTOT As Long
    Dim APLIC As Long
    Dim RESPON As Long
    Dim MC As Long
    Dim HIP As Long
    Dim prop As Long
    Dim arren As Long
    Dim Sleep As Long
    Dim Atraso As Long
    Dim PSFNLP As Long
    Dim PBBLP As Long
    Dim PSFNTOT As Long
    Dim PBBTOT As Long
    Dim PTARIT As String
    Dim PVCHES As Long
    Dim PVCART As Long
    Dim VPENHOR As Long
    Dim VHIPOTE As Long
    Dim PRBAagri As Long
    Dim PRBAPEC As Long
    Dim PAPLIC As Long
    Dim PRESPON As Long
    Dim PMC As Long
    Dim pprop As Long
    Dim Parren As Long
    Dim PHIP As Long
    Dim Pvlr As Long
    Dim PImpac As Long
    Dim a As Integer
    Dim opcao As Integer
    Dim Pconta As String
    Dim PPAA As Boolean
    Dim PAgen As Long
    Dim ok As Integer
    Dim GrupAgro As Boolean
    Dim egrupo As String
    Dim ecpf As Double
    Dim eprop As Double
    Dim Processo As Boolean
    Dim eaplic As String
    Dim vtarr As Long
    Dim gcess As Long
    Dim Tarren As Long
    Dim TParren As Long
    Dim Pdihip As Double
    Dim Gdihip As Double
    Dim VRLAF As Double
    Dim MCIAF As Long
    Dim linhanome As String
    Dim locaract As String
    Dim testec As Boolean
    Dim VDEP As Double
    Dim VDEG As Double
    Dim TRCFA As Long
    Dim TRNCFA As Long
    Dim perprop As Double
    Dim vatenu As Double
    Dim SFNEXT As Double
    Dim VDAF As Double
    Dim VCAF As Double
    Dim TDCAF As Double
    Dim CONJCO As Boolean
    Dim cedente As Long
    Dim cedgrup As Boolean
    Dim Vagarp As Double
    Dim Vagarg As Double
    Dim nbem As Double
    Dim bem As Boolean
    

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


linha7 = 2

    ' LIMPA A Tabela Inserção Dados Macro

  Do While Planilha6.Range("A" & linha7) <> ""
  
        linha = 4
        While linha < 153
        If linha <> 6 And linha <> 10 And linha <> 13 And linha <> 9 And linha <> 21 And linha <> 79 And linha <> 84 And linha <> 85 And linha <> 93 Then
        Range("e" & linha) = del
        Range("f" & linha) = del
        End If
        linha = linha + 1
        Wend
  
  
 
        If Planilha6.Range("B" & linha7) <> "sim" Then
        
       
         eprop = Planilha6.Range("A" & linha7)


         Plan1.Range("O3") = eprop

    
             
    ' 01 - Rentabiliza Agro'
   
                     
     ok = Aguardar(1, 2, "SBBP6130")
        

        '01.1 - valor cheque especial '
                
        PVCHES = 0
        
        TeclarTxt "03", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "20", 20, 11
        ENTER
        Esperasystem
        TeclarTxt "44", 16, 16
        ENTER
        Esperasystem
        TeclarTxt "01", 18, 13
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 21, 13
        ENTER
        Esperasystem
        
                If Trim(Copiar(10, 18, 37)) = "Atencao! Cliente com divida ajuizada." Then
                F3
                Esperasystem
                F3
                Esperasystem
                End If
                
                
                If Trim(Copiar(10, 12, 19)) = "Informação restrita" Then
                TeclarTxt "Analise de operacao Rural " & Range("e10"), 12, 12
                ENTER
                Esperasystem
                End If
        
        ENTER
        Esperasystem
                  
        
                If Trim(Copiar(10, 8, 3)) = "CHE" Then
                TeclarTxt "01", 21, 30
                ENTER
                Esperasystem
                TeclarTxt "01", 21, 30
                ENTER
                Esperasystem
                PVCHES = Trim(Copiar(10, 71, 10))
                Range("E77") = PVCHES
                End If
                
                While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
                F3
                Esperasystem
                Wend
                ok = Aguardar(1, 2, "SBBP6130")
        
        
        'PEGANDO O VALOR DE DEBITOS PROPONENTE
        
        TeclarTxt "03", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "20", 20, 11
        ENTER
        Esperasystem
        TeclarTxt "44", 16, 16
        ENTER
        Esperasystem
        TeclarTxt "01", 18, 13
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 21, 13
        ENTER
        Esperasystem
        
                If Trim(Copiar(10, 14, 4)) = "Essa" Then
                F3
                Esperasystem
                End If
        
                If Trim(Copiar(10, 18, 37)) = "Atencao! Cliente com divida ajuizada." Then
                F3
                Esperasystem
                F3
                Esperasystem
                End If
        
                If Trim(Copiar(10, 12, 19)) = "Informação restrita" Then
                TeclarTxt "Analise de operacao Rural " & Range("e10"), 12, 12
                ENTER
                Esperasystem
                End If
        
        ENTER
        Esperasystem
                  
                
                If Trim(Copiar(9, 3, 7)) = "DEBITOS" Then
                VDEP = VDEP + Trim(Copiar(9, 65, 16))
                Range("E118") = VDEP
                End If
                    
                While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
                F3
                Esperasystem
                Wend
                ok = Aguardar(1, 2, "SBBP6130")
                
 
        ' 01.3 - valor cheque CARTAO '
        
        PVCART = 0
        
        TeclarTxt "22", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "11", 21, 19
        Esperasystem
        ENTER
        TeclarTxt "01", 21, 20
        Esperasystem
        ENTER
        Esperasystem
        
        
        TeclarTxt "01", 17, 20
        TeclarTxt "", 20, 20
        F4
        Esperasystem
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 20, 36
        ENTER
        Esperasystem
        ENTER
        Esperasystem
        
        If Trim(Copiar(1, 3, 8)) = "MCIM0102" Then
        TeclarTxt "X", 6, 3
        ENTER
        Esperasystem
        End If
        
        
            If Trim(Copiar(11, 6, 19)) = "Informação restrita" Then
            TeclarTxt "Analise de operacao Rural " & Range("E10"), 14, 7
            ENTER
            Esperasystem
            End If
        TeclarTxt "x", 15, 4
        ENTER
        Esperasystem
        PVCART = Trim(Copiar(15, 30, 10))
        Range("E78") = PVCART
            While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
            F3
            Esperasystem
            Wend
            ok = Aguardar(1, 2, "SBBP6130")
                
        
        ' 02 - MITIGADOR DE RISCO
        
        
        TeclarTxt "27", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "03", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "24", 19, 24
        TeclarTxt WorksheetFunction.Text(Range("L10"), "0000"), 20, 24
        ENTER
        Esperasystem
        TeclarTxt "11", 19, 26
        ENTER
        Esperasystem
        TeclarTxt "11", 15, 28
        TeclarTxt WorksheetFunction.Text(Range("E10"), "000000000"), 16, 28
        ENTER
        Esperasystem
        TeclarTxt "11", 21, 47
        ENTER
        Esperasystem
        F8
        Esperasystem
        
         
        Range("E88") = Trim(Copiar(17, 37, 15))
        Range("E89") = Trim(Copiar(15, 37, 15))
        Range("E91") = Trim(Copiar(16, 37, 15))
        Range("E92") = Trim(Copiar(18, 37, 15))
        
        
        While (Trim(Copiar(1, 3, 8))) <> "COPM112T"
        F8
        Esperasystem
        Wend
        ok = Aguardar(1, 3, "COPM112T")
        
        TAEOP = Trim(Copiar(11, 77, 4))
            
        Range("E75") = TAEOP
        
        While (Trim(Copiar(1, 3, 8))) <> "COPM0000"
        F3
        Esperasystem
        Wend
        ok = Aguardar(1, 3, "COPM0000")
            
            
        
        ' 3 - GARANTIA / PATRIMÔNIO
                

        TeclarTxt "13", 19, 24
        ENTER
        Esperasystem
        TeclarTxt WorksheetFunction.Text(Range("L10"), "0000"), 20, 24
        ENTER
        Esperasystem
        TeclarTxt WorksheetFunction.Text(Range("E10"), "000000000"), 7, 50
        Esperasystem
        ENTER
        Esperasystem
            If Trim(Copiar(23, 2, 19)) = "Proposta cancelada." Then
            MsgBox "Proposta Cancelada"
             While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
             F3
             Esperasystem
            Wend
            ok = Aguardar(1, 2, "SBBP6130")
            
            GoTo ProximoFim
            End If
            
           ' 3.1 - PEGANDO VALOR DA GARANTIA PENHOR
           
        linha = 10
              
            While Trim(Copiar(linha, 5, 6)) <> "SUMULA"
                linha = linha + 1
            Wend
                TeclarTxt "x", linha, 3
                ENTER
                Esperasystem
      
            linha = 4


           VPENHOR = 0
             
              
           Do While linha < 22 And Trim(Copiar(linha, 3, 4) <> "2.4.")
            
                linha = linha + 1
                
                Select Case linha
                 
                Case 5 To 16
                
                If Trim(Copiar(linha, 51, 6)) = "PENHOR" Then
                linha = linha + 4
                VPENHOR = VPENHOR + Trim(Copiar(linha, 40, 15))
                End If
                 
                Case 17
                If Trim(Copiar(linha, 51, 6)) = "PENHOR" Then
                F8
                Esperasystem
                linha = 5
                VPENHOR = VPENHOR + Trim(Copiar(linha, 40, 15))
                End If
                
                Case 18
                If Trim(Copiar(linha, 51, 6)) = "PENHOR" Then
                F8
                Esperasystem
                linha = 6
                VPENHOR = VPENHOR + Trim(Copiar(linha, 40, 15))
                End If
                
                
                Case 19
                If Trim(Copiar(linha, 51, 6)) = "PENHOR" Then
                F8
                Esperasystem
                linha = 7
                VPENHOR = VPENHOR + Trim(Copiar(linha, 40, 15))
                End If
                
                
                Case 20
                If Trim(Copiar(linha, 51, 6)) = "PENHOR" Then
                F8
                Esperasystem
                linha = 8
                VPENHOR = VPENHOR + Trim(Copiar(linha, 40, 15))
                End If
                
                Case 21
                F8
                Esperasystem
                linha = 4

                End Select
                
                Range("E82") = VPENHOR
                 
            Loop
                           
        While (Trim(Copiar(1, 3, 8))) <> "COPM2912"
        F3
        Esperasystem
        Wend
        ok = Aguardar(1, 3, "COPM2912")
        
        ' 3.2 - PEGANDO VALOR DA GARANTIA HIPOTECA
                                    
  
        linha = 10
              
            While Trim(Copiar(linha, 5, 6)) <> "SUMULA"
            linha = linha + 1
            Wend
            TeclarTxt "x", linha, 3
            ENTER
            Esperasystem
      
            linha = 4


            VHIPOTE = 0
             
             
           Do While linha < 22 And Trim(Copiar(linha, 3, 4) <> "2.4.")
            
                linha = linha + 1
                
                Select Case linha
                 
                Case 5 To 16
                
                If Trim(Copiar(linha, 51, 8)) = "HIPOTECA" Then
                linha = linha + 4
                VHIPOTE = VHIPOTE + Trim(Copiar(linha, 40, 15))
                End If
                 
                Case 17
                If Trim(Copiar(linha, 51, 8)) = "HIPOTECA" Then
                F8
                Esperasystem
                linha = 5
                VHIPOTE = VHIPOTE + Trim(Copiar(linha, 40, 15))
                End If
                
                Case 18
                If Trim(Copiar(linha, 51, 8)) = "HIPOTECA" Then
                F8
                Esperasystem
                linha = 6
                VHIPOTE = VHIPOTE + Trim(Copiar(linha, 40, 15))
                End If
                
                
                Case 19
                If Trim(Copiar(linha, 51, 8)) = "HIPOTECA" Then
                F8
                Esperasystem
                linha = 7
                VHIPOTE = VHIPOTE + Trim(Copiar(linha, 40, 15))
                End If
                
                
                Case 20
                If Trim(Copiar(linha, 51, 8)) = "HIPOTECA" Then
                F8
                Esperasystem
                linha = 8
                VHIPOTE = VHIPOTE + Trim(Copiar(linha, 40, 15))
                End If
                
                Case 21
                F8
                Esperasystem
                linha = 4

                End Select
                
                Range("E81") = VHIPOTE
                 
            Loop
  
  
  
  
 While (Trim(Copiar(1, 3, 8))) <> "COPM2912"
        F3
        Esperasystem
        Wend
        ok = Aguardar(1, 3, "COPM2912")
        
        
        ' 02.4 - Pegando mci dos  Avalistas e Fiadores

        
        linha = 10
              
            While Trim(Copiar(linha, 5, 6)) <> "SUMULA"
                linha = linha + 1
            Wend
                TeclarTxt "x", linha, 3
                ENTER
                Esperasystem
                
                
         linha = 4
         linha1 = 98

             
           Do While linha < 22 And Trim(Copiar(linha, 3, 4) <> "2.4.")
            
                linha = linha + 1
             
                
                Select Case linha
                 
                Case 5 To 20
                
                If Trim(Copiar(linha, 7, 3)) = "MCI" Then
                
                Range("E" & linha1) = Replace(Trim(Copiar(linha, 20, 11)), ".", "")
                linha1 = linha1 + 1
                               
                End If
                 
                
                Case 21
                F8
                Esperasystem
                linha = 4
                    If Trim(Copiar(23, 2, 6)) = "Ultima" Then
                    linha = 23
                    End If
                Esperasystem
                

                End Select
                  
                 
            Loop
  
  
  While (Trim(Copiar(1, 3, 8))) <> "COPM2912"
        F3
        Esperasystem
        Wend
        ok = Aguardar(1, 3, "COPM2912")
  

  
          locaract = "*"
          linha = 10
              
            While Trim(Copiar(linha, 5, 6)) <> "SUMULA"
                linha = linha + 1
            Wend
                TeclarTxt "x", linha, 3
                Esperasystem
                ENTER
                Esperasystem
                
                
         linha = 4
         linha2 = 108
  
  
  Do While linha < 22 And Trim(Copiar(linha, 3, 4) <> "2.4.")
            
                linha = linha + 1
             
                
                Select Case linha
                 
                Case 5 To 20
                
                If Trim(Copiar(linha, 7, 3)) = "MCI" Then
                
                
                    If linha = 5 Then
                    
                        F7
                        Esperasystem
                        linha = 20
                      
                        linhanome = Trim(Copiar(linha, 4, 76))
                    
                        testec = InStr(linhanome, locaract)
                    
                    
                        F8
                        Esperasystem
                        linha = 5
                   
                     
                            If testec = True Then
                    
                            Range("E" & linha2) = Replace(Trim(Copiar(linha, 20, 11)), ".", "")
                            
                            linha2 = linha2 + 1
                      
                            End If
                    
                    End If
                    
                    
                    If linha <> 5 Then
                        
                    
                        linha = linha - 1
                        
                        linhanome = Trim(Copiar(linha, 4, 76))
                    
                        testec = InStr(linhanome, locaract)
                      
                      
                        linha = linha + 1
                            
                            If testec = True Then
                            
                    
                             Range("E" & linha2) = Replace(Trim(Copiar(linha, 20, 11)), ".", "")
                    
                             linha2 = linha2 + 1
                            
                            End If
                                                   
                                                   
                        End If
                
            
                
                End If
                
                Case 21
                F8
                Esperasystem
                linha = 4

                End Select
                
                
                 
            Loop
            
    
  While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
        Esperasystem
        F3
        Esperasystem
        Wend
        ok = Aguardar(1, 2, "SBBP6130")
  
  
  ' 4 - Dados da operação
        
           
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
            If Trim(Copiar(4, 10, 3)) = "Olá" Then
            ENTER
            Esperasystem
            End If
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 20, 57
        ENTER
        Esperasystem
        
            If Trim(Copiar(9, 12, 19)) = "Informação restrita" Then
            TeclarTxt "Analise de operacao Rural " & Range("E10"), 11, 12
            ENTER
            Esperasystem
            End If
        
        If Trim(Copiar(17, 5, 2)) = "" Then
        TeclarTxt "x", 6, 3
        ENTER
        Esperasystem
        If Trim(Copiar(9, 12, 3)) = "Inf" Then
        TeclarTxt "Analise de Operacao Agro", 11, 12
        ENTER
        Esperasystem
        End If
        End If
        
        
'4.1 - Pegando nome'
        Range("E5") = Trim(Copiar(4, 18, 50))
        
        
'4.2 -  Pegando grupo'
        
        TeclarTxt "04", 21, 41
        ENTER
        Esperasystem
        TeclarTxt "5", 21, 26
        ENTER
        Esperasystem

        If (Trim(Copiar(10, 7, 1))) <> "" Then
            TeclarTxt "x", 10, 4
            ENTER
            Esperasystem
        linha1 = 9
        linha = 63
        While linha1 < 21
            Range("E" & linha) = Trim(Copiar(linha1, 71, 9))
            linha1 = linha1 + 1
            linha = linha + 1
        Wend
        F3
        Esperasystem
        End If
        If (Trim(Copiar(10, 7, 1))) = "" Then
        Range("E63") = TiraPonto(Trim(Copiar(6, 69, 11)))
        End If
        
        F3
        Esperasystem
        F3
        Esperasystem
        
'4.3 - pegando poderes'

        TeclarTxt "08", 21, 41
        ENTER
        Esperasystem
        TeclarTxt "x", 15, 44
        ENTER
        Esperasystem
        Range("E24") = Trim(Copiar(23, 3, 3))
        F3
        Esperasystem
        F3
        Esperasystem
        
'4.4 - pegando recursos computaveis'

        If Trim(Copiar(17, 5, 2)) = "" Then
        F3
        Esperasystem
        If Trim(Copiar(7, 23, 3)) = "Exi" Then
        F3
        Esperasystem
        End If
        ENTER
        Esperasystem
        End If
        
        TeclarTxt "01", 19, 18
        ENTER
        Esperasystem
        If Trim(Copiar(4, 10, 3)) = "Olá" Then
        F3
        Esperasystem
        End If
        If Trim(Copiar(16, 16, 2)) = "F3" Then
        F3
        Esperasystem
        End If
      
      
        comp = 0
        comp2 = 0
        cont = 0
        linha1 = 63
        While Range("E" & linha1) <> ""
        TeclarTxt WorksheetFunction.Text(Range("E" & linha1), "000000000"), 20, 57
        ENTER
        Esperasystem
        If Trim(Copiar(9, 12, 3)) = "Inf" Then
        TeclarTxt "Analise de Operacao Agro", 11, 12
        ENTER
        Esperasystem
        End If
        
        TeclarTxt "5", 21, 41
        ENTER
        Esperasystem
        While cont < 21
        For cont = 15 To 20
                If (Trim(Copiar(cont, 43, 14))) <> "" Then
                        comp = comp + CLng(Replace(Trim(Copiar(cont, 43, 14)), ".", ""))
                End If
                If (Trim(Copiar(cont, 67, 14))) <> "" Then
                        comp2 = comp2 + CLng(Replace(Trim(Copiar(cont, 67, 14)), ".", ""))
                End If
        Next
        Wend
        linha1 = linha1 + 1
        cont = 0
        F3
        Esperasystem
        F3
        Esperasystem
        Wend
        Range("E25") = comp
        Range("E26") = comp2

'4.5 -  pegando recursos computaveis - Proponente'

        If Trim(Copiar(17, 5, 2)) = "" Then
        F3
        Esperasystem
        If Trim(Copiar(7, 23, 3)) = "Exi" Then
        F3
        Esperasystem
        End If
        ENTER
        Esperasystem
        End If
        
        TeclarTxt "01", 19, 18
        ENTER
        Esperasystem
            If Trim(Copiar(4, 10, 3)) = "Olá" Then
            ENTER
            Esperasystem
            End If
        If Trim(Copiar(16, 16, 2)) = "F3" Then
        F3
        Esperasystem
        End If
      
        comp = 0
        comp2 = 0
        cont = 0
      
         TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 20, 57 ' aqui utiliza o mci do proponente
       
        ENTER
        Esperasystem
        If Trim(Copiar(9, 12, 3)) = "Inf" Then
        TeclarTxt "Analise de Operacao Agro", 11, 12
        ENTER
        Esperasystem
        End If
        
        TeclarTxt "5", 21, 41
        ENTER
        Esperasystem
        While cont < 21
        For cont = 15 To 20
                If (Trim(Copiar(cont, 43, 14))) <> "" Then
                        comp = comp + CLng(Replace(Trim(Copiar(cont, 43, 14)), ".", ""))
                End If
                If (Trim(Copiar(cont, 67, 14))) <> "" Then
                        comp2 = comp2 + CLng(Replace(Trim(Copiar(cont, 67, 14)), ".", ""))
                End If
        Next
        Wend
        linha1 = linha1 + 1
        cont = 0
        Range("F25") = comp
        Range("F26") = comp2
        F3
        Esperasystem
        F3
        Esperasystem
    
    '4.6 - pegando recursos computaveis Fiadores e avalistas

        If Trim(Copiar(17, 5, 2)) = "" Then
        F3
        Esperasystem
        If Trim(Copiar(7, 23, 3)) = "Exi" Then
        F3
        Esperasystem
        End If
        ENTER
        Esperasystem
        End If
        
        TeclarTxt "01", 19, 18
        ENTER
        Esperasystem
            If Trim(Copiar(4, 10, 3)) = "Olá" Then
            ENTER
            Esperasystem
            End If
             
             If Trim(Copiar(16, 16, 2)) = "F3" Then
             F3
             Esperasystem
             End If
      
        comp = 0
        comp2 = 0
        cont = 0
        linha1 = 98
        
            While Range("E" & linha1) <> ""
            TeclarTxt WorksheetFunction.Text(Range("E" & linha1), "000000000"), 20, 57
            ENTER
            Esperasystem
                If Trim(Copiar(9, 12, 3)) = "Inf" Then
                TeclarTxt "Analise de Operacao Agro", 11, 12
                 ENTER
                 Esperasystem
                End If
        
            TeclarTxt "5", 21, 41
            ENTER
            Esperasystem
                While cont < 21
                    For cont = 15 To 20
                    If (Trim(Copiar(cont, 43, 14))) <> "" Then
                        comp = comp + CLng(Replace(Trim(Copiar(cont, 43, 14)), ".", ""))
                    End If
                    If (Trim(Copiar(cont, 67, 14))) <> "" Then
                        comp2 = comp2 + CLng(Replace(Trim(Copiar(cont, 67, 14)), ".", ""))
                    End If
        Next
        Wend
        linha1 = linha1 + 1
        cont = 0
        F3
        Esperasystem
        F3
        Esperasystem
        Wend
        TRCFA = TRCFA + comp
        TRNCFA = TRNCFA + comp2
        
        Range("E120") = TRCFA
        Range("E121") = TRNCFA

'5 - pegando garantias hipotecárias, impacto, hipotecaveis'

' 5.1 - grupo

        prop = 0
   
        HIP = 0
        vlr = 0
        Impac = 0
        linha1 = 63
        Tarren = 0
        cont = 9
        While Range("E" & linha1) <> ""
        TeclarTxt WorksheetFunction.Text(Range("E" & linha1), "000000000"), 20, 57
        ENTER
        Esperasystem
        If Trim(Copiar(9, 12, 3)) = "Inf" Then
        TeclarTxt "Analise de Operacao Agro", 11, 12
        ENTER
        Esperasystem
        End If
        TeclarTxt "5", 21, 41
        ENTER
        Esperasystem
         If (Trim(Copiar(20, 32, 1))) <> "" Then
        TeclarTxt "x", 20, 4
        ENTER
        Esperasystem
       
        '5.1 - arrenndamento
        
        While Trim(Copiar(23, 3, 3)) <> "Últ"
        For cont = 9 To 20
        If Trim(Copiar(cont, 7, 3)) = "ARR" Or Trim(Copiar(cont, 7, 3)) = "COM" Then
        TeclarTxt "x", cont, 4
        ENTER
        Esperasystem
        arren = 0
        gcess = 1
        linha = 63
        linha2 = 63
            
            While linha < 73

            cedente = Trim(Copiar(15, 7, 9))
            
            Do While Range("e" & linha1) <> "" And linha2 < 73
                
                If cedente <> Range("e" & linha2) Then
                cedgrup = False
                End If
                
                If cedente = Range("e" & linha2) Then
                cedgrup = True
                linha2 = 74
                End If
                
              linha2 = linha2 + 1
                
            Loop
            

                If cedgrup = False Then
                
                If Trim(Copiar(10, 24, 11)) <> "" Then
                
                        If Trim(Copiar(24, 62, 6)) = "Outros" Then
                        F11
                        Esperasystem
                   
                            linha5 = 13
                            
            
                            If gcess < 2 Then
                                
                                While Trim(Copiar(linha5, 13, 9)) <> "" And Trim(Copiar(19, 8, 3)) <> "Últ"
                                    gcess = gcess + 1
                                    linha5 = linha5 + 1
                                    If linha5 = 18 Then
                                        F8
                                        Esperasystem
                                        linha5 = 13
                                    End If
                                Wend
                             End If
                            F3
                            Esperasystem

                         End If
                
                        arren = arren + CLng(Replace(Trim(Copiar(10, 24, 10)), ".", ""))
                      
                    End If


            End If
            
             If cedgrup = True Then
                
                
                linha = 74
                      
             End If

       
            linha = linha + 1
            
        Wend

        arren = arren / 10
        arren = arren / gcess
        Tarren = Tarren + arren
        
        F3
        Esperasystem
        End If
        
proximo4:
        Next
        F8
        Esperasystem
        Wend
        F3
        Esperasystem
        End If
        
        TeclarTxt "x", 17, 4
        ENTER
        Esperasystem
        If (Trim(Copiar(23, 3, 3))) = "Não" Then
            GoTo Proximo1
        End If
        While (Trim(Copiar(23, 3, 3))) <> "Últ"
        For cont = 10 To 20
            If (Trim(Copiar(cont, 16, 1))) <> "" Then
            prop = prop + (CLng(Replace(Trim(Copiar(cont, 75, 6)), ".", "")) / 100) * CLng(Replace(Trim(Copiar(cont, 43, 10)), ".", ""))
            
            If (Trim(Copiar(cont, 75, 6))) <> "" Then
            
            perprop = Trim(Copiar(cont, 75, 6)) / 100
            
            F12
            
            vatenu = Trim(Copiar(cont, 49, 14))
            
            vatenu = vatenu * perprop
            
            HIP = HIP + vatenu
            
            F11
            Esperasystem
            
            End If
            
            
            TeclarTxt "x", cont, 3
            ENTER
            Esperasystem
            F2
            Esperasystem
            TeclarTxt "9", 20, 20
            ENTER
            Esperasystem
            If (Trim(Copiar(23, 3, 1))) <> "O" Then
                    vlr = vlr + CLng(Replace(Trim(Copiar(5, 67, 14)), ".", ""))
                    Impac = Impac + CLng(Replace(Trim(Copiar(20, 67, 14)), ".", ""))
            F3
            Esperasystem
            End If
            F3
            Esperasystem
            F3
            Esperasystem
            End If
        Next
        F8
        Esperasystem
        Wend
Proximo1:
        linha1 = linha1 + 1
        F3
        Esperasystem
        F3
        Esperasystem
        F3
        Esperasystem
        Wend
        linha1 = linha1 + 1
        
        Range("E28") = (Impac)
        Range("E27") = vlr
        Range("E62") = HIP
        Range("e56") = prop
        Range("e57") = Tarren
   
        
' 5.3 - proponente

        pprop = 0
        TParren = 0
        PHIP = 0
        Pvlr = 0
        PImpac = 0
     
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 20, 57
        ENTER
        Esperasystem
        If Trim(Copiar(9, 12, 3)) = "Inf" Then
        TeclarTxt "Analise de Operacao Agro", 11, 12
        ENTER
        Esperasystem
        End If
        TeclarTxt "5", 21, 41
        ENTER
        Esperasystem
         If (Trim(Copiar(20, 32, 1))) <> "" Then
        TeclarTxt "x", 20, 4
        ENTER
        Esperasystem
       
        'arrenndamento
       While Trim(Copiar(23, 3, 3)) <> "Últ"
        For cont = 9 To 20
        If Trim(Copiar(cont, 7, 3)) = "ARR" Or Trim(Copiar(cont, 7, 3)) = "COM" Then
        TeclarTxt "x", cont, 4
        ENTER
        Esperasystem
        arren = 0
        gcess = 1
        linha = 63 'parei aqui
        linha2 = 63
        
            While linha < 73
            
             cedente = Trim(Copiar(15, 7, 9))
            
            Do While Range("e" & linha1) <> "" And linha2 < 73
                
                If cedente <> Range("e" & linha2) Then
                cedgrup = False
                End If
                
                If cedente = Range("e" & linha2) Then
                cedgrup = True
                linha2 = 74
                End If
                
              linha2 = linha2 + 1
                
            Loop
            

                If cedgrup = False Then
                
                
                    If Trim(Copiar(10, 24, 11)) <> "" Then
                
                        If Trim(Copiar(24, 62, 6)) = "Outros" Then
                        F11
                        Esperasystem
                   
                            linha5 = 13
                            
            
                            If gcess < 2 Then
                                
                                 While Trim(Copiar(linha5, 13, 9)) <> "" And Trim(Copiar(19, 8, 3)) <> "Últ"
                                    gcess = gcess + 1
                                    linha5 = linha5 + 1
                                    If linha5 = 18 Then
                                        F8
                                        Esperasystem
                                        linha5 = 13
                                    End If
                                Wend
                             End If
                            F3
                            Esperasystem

                         End If
                
                        arren = arren + CLng(Replace(Trim(Copiar(10, 24, 10)), ".", ""))
                      
                    End If

            End If
            
             If cedgrup = True Then
                
                
                linha = 74
                      
             End If

       
            linha = linha + 1
            
        Wend
        arren = arren / 10
        arren = arren / gcess
        TParren = TParren + arren
        F3
        Esperasystem
     
        End If
proximo22:
        Next
        F8
        Esperasystem
        Wend
        F3
        Esperasystem
        End If
        
        ' alteração de regra
        '-----------------------------------------
     
        TeclarTxt "x", 17, 4
        ENTER
        Esperasystem
        If (Trim(Copiar(23, 3, 3))) = "Não" Then
            GoTo Proximo23
        End If
        While (Trim(Copiar(23, 3, 3))) <> "Últ"
        For cont = 10 To 20
            If (Trim(Copiar(cont, 16, 1))) <> "" Then
            pprop = pprop + (CLng(Replace(Trim(Copiar(cont, 75, 6)), ".", "")) / 100) * CLng(Replace(Trim(Copiar(cont, 43, 10)), ".", ""))
            
                  
            If (Trim(Copiar(cont, 75, 6))) <> "" Then
            
            perprop = Trim(Copiar(cont, 75, 6)) / 100
            
            
            F12
            
            
            vatenu = Trim(Copiar(cont, 49, 14))
            
            vatenu = vatenu * perprop
            
            PHIP = PHIP + vatenu
         
            F11
            Esperasystem
            
            End If
            
            TeclarTxt "x", cont, 3
            ENTER
            Esperasystem
            F2
            Esperasystem
            TeclarTxt "9", 20, 20
            ENTER
            Esperasystem
            If (Trim(Copiar(23, 3, 1))) <> "O" Then
                    Pvlr = Pvlr + CLng(Replace(Trim(Copiar(5, 67, 14)), ".", ""))
                    PImpac = PImpac + CLng(Replace(Trim(Copiar(20, 67, 14)), ".", ""))
            F3
            Esperasystem
            End If
            F3
            Esperasystem
            F3
            Esperasystem
            End If
        Next
        F8
        Esperasystem
        Wend
Proximo23:
        linha1 = linha1 + 1
        F3
        Esperasystem
        F3
        Esperasystem
        F3
        Esperasystem
        linha1 = linha1 + 1
        Range("F28") = (PImpac)
        Range("F27") = Pvlr
        Range("F62") = PHIP
        Range("F56") = pprop
        Range("F57") = TParren
 
 
        While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
            F3
            Esperasystem
            Wend
            ok = Aguardar(1, 2, "SBBP6130")
        
                    
         'PEGANDO O VALOR DE DEBITOS GRUPO

        TeclarTxt "03", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "20", 20, 11
        ENTER
        Esperasystem
        TeclarTxt "44", 16, 16
        ENTER
        Esperasystem
        
         linha1 = 63
        
        Do While Range("e" & linha1) <> ""
        
         TeclarTxt "01", 18, 13
    
                    TeclarTxt WorksheetFunction.Text(Range("E" & linha1), "000000000"), 21, 13
                    ENTER
                    Esperasystem
                    
                        If Trim(Copiar(10, 18, 37)) = "Atencao! Cliente com divida ajuizada." Then
                        F3
                        Esperasystem
                        F3
                        Esperasystem
                        End If
                
                        If Trim(Copiar(10, 12, 19)) = "Informação restrita" Then
                        TeclarTxt "Analise de operacao Rural " & Range("e10"), 12, 12
                        ENTER
                        Esperasystem
                        End If
        
                        ENTER
                        Esperasystem
                  
                
                        If Trim(Copiar(9, 3, 7)) = "DEBITOS" Then
                        VDEG = VDEG + Trim(Copiar(9, 65, 16))
                        Range("E119") = VDEG
                        End If
                    
                        F3
                        Esperasystem
                         linha1 = linha1 + 1
                
             Loop
             
                While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
                F3
                Esperasystem
                Wend
                ok = Aguardar(1, 2, "SBBP6130")

        
  '6 - Pegando dados REL
        
'6.1 - Pegando dados REL cliente


        TeclarTxt "15", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "01", 21, 18
        ENTER
        Esperasystem
        If (Trim(Copiar(8, 16, 2))) = "AS" Then
        F3
        Esperasystem
        End If
        
        
        TeclarTxt "11", 19, 18
        ENTER
        Esperasystem
    
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 20, 57
        ENTER
        Esperasystem
        
        If Trim(Copiar(17, 5, 2)) = "" Then
        TeclarTxt "x", 6, 3
        ENTER
        Esperasystem
        If Trim(Copiar(9, 12, 3)) = "Inf" Then
        TeclarTxt "Analise de Operacao Agro", 11, 12
        ENTER
        Esperasystem
        End If
        End If
        
        If Trim(Copiar(7, 46, 3)) = "Cli" Then
        F3
        Esperasystem
        End If
        
        Range("e8") = Trim(Copiar(10, 79, 2))
        Range("e11") = Trim(Copiar(16, 80, 1))
        Range("e12") = Trim(Copiar(19, 80, 1))
        TeclarTxt "X", 19, 3
        ENTER
        Esperasystem
        If (Trim(Copiar(11, 48, 8))) = "SUSPENSO" Then
            Range("F25") = Trim(Copiar(11, 48, 8))
            F3
            Esperasystem
            F3
            Esperasystem
            F3
            Esperasystem
            MsgBox "Limite Suspenso"
            GoTo proximo3:
        End If
        Range("F73") = Trim(Copiar(9, 30, 15))
        Range("e7") = Trim(Copiar(12, 77, 1))
        F11
        Range("e73") = TiraPonto(Trim(Copiar(18, 28, 12)))
        F3
        Esperasystem
        F3
        Esperasystem
        F3
        Esperasystem
  
     
'6.2 -  Pegando dados REL do Grupo'

        linha1 = 63
        RBAagri = 0
        RBApec = 0
        APLIC = 0
        RESPON = 0
        MC = 0
       
        While Range("e" & linha1) <> ""
        TeclarTxt "11", 19, 18
        Esperasystem
        ENTER
        Esperasystem
        TeclarTxt WorksheetFunction.Text(Range("e" & linha1), "000000000"), 20, 57
        ENTER
        Esperasystem
        
        If Trim(Copiar(17, 5, 2)) = "" Then
        TeclarTxt "x", 6, 3
        ENTER
        Esperasystem
        If Trim(Copiar(9, 12, 3)) = "Inf" Then
        TeclarTxt "Analise de Operacao Agro", 11, 12
        ENTER
        Esperasystem
        End If
        End If
        
        If Trim(Copiar(7, 46, 3)) = "Cli" Then
        F3
        Esperasystem
        End If
        
        If (Trim(Copiar(7, 46, 1))) = "C" Then
        F3
        Esperasystem
        End If
        If (Trim(Copiar(1, 3, 8))) <> "RELM129D" Then
      
        
' 7 - APLICAÇÕES, RESPONSABILIDADES E MC'
        
        APLIC = APLIC + CLng(Replace(Trim(Copiar(14, 67, 14)), ".", ""))
        RESPON = RESPON + CLng(Replace(Trim(Copiar(15, 67, 14)), ".", ""))
        If Trim(Copiar(16, 27, 11)) = "0,00" Then
        MC = 0
        GoTo proximo5
        End If
        TeclarTxt "X", 16, 3
        ENTER
        Esperasystem
        F2
        Esperasystem
        MC = MC + CLng(Replace(Trim(Copiar(19, 49, 13)), ".", ""))
        F3
        Esperasystem
        F3
        Esperasystem
proximo5:
      
        'RBA'
    If Trim(Copiar(16, 80, 1)) <> "S" Then
        TeclarTxt "X", 19, 3
        ENTER
        Esperasystem
        If (Trim(Copiar(11, 48, 8))) = "SUSPENSO" Then
            F3
            Esperasystem
        End If
        F12
        TeclarTxt "06", 20, 30
        ENTER
        Esperasystem
   
        
        If Trim(Copiar(1, 3, 8)) <> "ANCMF511" Then
        RBAagri = RBAagri + CLng(Replace(Trim(Copiar(13, 40, 14)), ".", ""))
        RBApec = RBApec + CLng(Replace(Trim(Copiar(13, 67, 14)), ".", ""))
        End If
        End If
        End If
            
            While (Trim(Copiar(1, 3, 8))) <> "RELM3001"
            F3
            Esperasystem
        Wend
        linha1 = linha1 + 1
        Wend
        Range("E18") = RBAagri
        Range("E19") = RBApec
        Range("E15") = APLIC
        Range("E16") = RESPON
        Range("E20") = MC

        
        While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
        F3
        Esperasystem
        Wend

        '7.2 ) Pegando dados do PROPONENTE'
     
        PRBAagri = 0
        PRBAPEC = 0
        PAPLIC = 0
        PRESPON = 0
        PMC = 0
        
        TeclarTxt "15", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "01", 21, 18
        ENTER
        Esperasystem
        TeclarTxt "11", 19, 18
        ENTER
        Esperasystem
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 20, 57
        ENTER
        Esperasystem
        If Trim(Copiar(17, 5, 2)) = "" Then
        TeclarTxt "x", 6, 3
        ENTER
        Esperasystem
        If Trim(Copiar(9, 12, 3)) = "Inf" Then
        TeclarTxt "Analise de Operacao Agro", 11, 12
        ENTER
        Esperasystem
        End If
        End If
        
        If Trim(Copiar(7, 46, 3)) = "Cli" Then
        F3
        Esperasystem
        End If
        
        If (Trim(Copiar(7, 46, 1))) = "C" Then
        F3
        Esperasystem
        End If
        If (Trim(Copiar(1, 3, 8))) <> "RELM129D" Then
      
        'APLICAÇÕES, RESPONSABILIDADES E MC'
        
        PAPLIC = PAPLIC + CLng(Replace(Trim(Copiar(14, 67, 14)), ".", ""))
        PRESPON = PRESPON + CLng(Replace(Trim(Copiar(15, 67, 14)), ".", ""))
        If Trim(Copiar(16, 27, 11)) = "0,00" Then
        PMC = 0
        GoTo PROXIMO21
        End If
        TeclarTxt "X", 16, 3
        ENTER
        Esperasystem
        F2
        Esperasystem
        PMC = PMC + CLng(Replace(Trim(Copiar(19, 49, 13)), ".", ""))
        F3
        Esperasystem
        F3
        Esperasystem
PROXIMO21:
      
        'RBA'
        TeclarTxt "X", 19, 3
        ENTER
        Esperasystem
        If (Trim(Copiar(11, 48, 8))) = "SUSPENSO" Then
            F3
            Esperasystem
        End If
        F12
        TeclarTxt "06", 20, 30
        ENTER
        Esperasystem
        If Trim(Copiar(1, 3, 8)) <> "ANCMF511" Then
        PRBAagri = PRBAagri + CLng(Replace(Trim(Copiar(13, 40, 14)), ".", ""))
        PRBAPEC = PRBAPEC + CLng(Replace(Trim(Copiar(13, 67, 14)), ".", ""))
        End If
        End If
       
            While (Trim(Copiar(1, 3, 8))) <> "RELM3001"
            F3
            Esperasystem
            Wend
        
        linha1 = linha1 + 1
        
        Range("F18") = PRBAagri
        Range("F19") = PRBAPEC
        Range("F15") = PAPLIC
        Range("F16") = PRESPON
        Range("F20") = PMC
        
        While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
           F3
           Esperasystem
        Wend
        
        
'8 - Pegando dados SCR
        
'8.1 - Pegando dados SCR do Grupo'

        TeclarTxt "04", 21, 20
        ENTER
        Esperasystem
        SFNLP = 0
        BBLP = 0
        SFNTOT = 0
        BBTOT = 0
        linha1 = 63
        While Range("e" & linha1) <> ""
        TeclarTxt "04", 16, 20
        TeclarTxt WorksheetFunction.Text(Range("e" & linha1), "000000000"), 19, 20
        ENTER
        Esperasystem
        If Trim(Copiar(23, 3, 3)) = "Res" Then
        GoTo próximo7
        End If
        If Trim(Copiar(23, 3, 3)) = "Cli" Then
        GoTo próximo7
        End If
        TeclarTxt "01", 21, 27
        ENTER
        Esperasystem
            If Trim(Copiar(23, 3, 3)) = "Res" Then
            GoTo próximo7
            End If
            If Trim(Copiar(23, 3, 3)) = "Cli" Then
            GoTo próximo7
            End If
        TeclarTxt "04", 17, 24
        ENTER
        Esperasystem
            If Trim(Copiar(23, 3, 3)) = "Res" Then
            GoTo próximo7
            End If
            If Trim(Copiar(23, 3, 3)) = "Cli" Then
            GoTo próximo7
            End If
       
        SFNLP = SFNLP + CLng(Replace(Trim(Copiar(13, 45, 14)), ".", ""))
        BBLP = BBLP + CLng(Replace(Trim(Copiar(13, 67, 14)), ".", ""))
        F3
        Esperasystem
        SFNTOT = SFNTOT + CLng(Replace(Trim(Copiar(12, 36, 14)), ".", ""))
        BBTOT = BBTOT + CLng(Replace(Trim(Copiar(12, 64, 14)), ".", ""))
        
  
próximo7:
            While (Trim(Copiar(1, 3, 8))) <> "CERM3000"
            F3
            Esperasystem
            Wend

        
        linha1 = linha1 + 1
        Wend
        Range("E29") = SFNLP
        Range("E31") = BBLP
        Range("E30") = SFNTOT - SFNLP
        Range("E32") = BBTOT - BBLP
        F3
        Esperasystem
        
 '8.2 - Pegando dados SCR do Proponente'
 
        TeclarTxt "04", 21, 20
        ENTER
        Esperasystem
        Esperasystem
        PSFNLP = 0
        PBBLP = 0
        PSFNTOT = 0
        PBBTOT = 0
        linha1 = 63
        TeclarTxt "04", 16, 20
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 19, 20
        ENTER
        Esperasystem
        TeclarTxt "01", 21, 27
        ENTER
        Esperasystem
        TeclarTxt "04", 17, 24
        ENTER
        Esperasystem
        
        PSFNLP = PSFNLP + CLng(Replace(Trim(Copiar(13, 45, 14)), ".", ""))
        PBBLP = PBBLP + CLng(Replace(Trim(Copiar(13, 67, 14)), ".", ""))
        F3
        Esperasystem
        PSFNTOT = PSFNTOT + CLng(Replace(Trim(Copiar(12, 36, 14)), ".", ""))
        PBBTOT = PBBTOT + CLng(Replace(Trim(Copiar(12, 64, 14)), ".", ""))
        F3
        Esperasystem
        F3
        Esperasystem
        Range("F29") = PSFNLP
        Range("F31") = PBBLP
        Range("F30") = PSFNTOT - PSFNLP
        Range("F32") = PBBTOT - PBBLP
        F3
        Esperasystem
        
  '8.3 ) Pegando dados SCR do Avalista e Fiadores

        TeclarTxt "04", 21, 20
        ENTER
        Esperasystem
        SFNLP = 0
        BBLP = 0
        SFNTOT = 0
        BBTOT = 0
        
        linha1 = 98
        
        While Range("e" & linha1) <> ""
        TeclarTxt "04", 16, 20
        TeclarTxt WorksheetFunction.Text(Range("e" & linha1), "000000000"), 19, 20
        ENTER
        Esperasystem
        If Trim(Copiar(23, 3, 3)) = "Res" Then
        GoTo próximo70
        End If
        If Trim(Copiar(23, 3, 3)) = "Cli" Then
        GoTo próximo70
        End If
        TeclarTxt "01", 21, 27
        ENTER
        Esperasystem

        SFNTOT = SFNTOT + CLng(Replace(Trim(Copiar(12, 36, 14)), ".", ""))
        BBTOT = BBTOT + CLng(Replace(Trim(Copiar(12, 64, 14)), ".", ""))
        F3
        Esperasystem
        F3
        Esperasystem

próximo70:
        linha1 = linha1 + 1
       
        Wend
        
        SFNEXT = SFNTOT - BBTOT
        Range("e122") = SFNEXT
    
        F3
        Esperasystem
    ' 9. PEGANDO DEBITOS 3.20.44 E COOBRIGAÇÃO DO AVALISTAS E FIADORES

        linha1 = 98
        
        VDAF = 0
        VCAF = 0
        
        TeclarTxt "03", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "20", 20, 11
        ENTER
        Esperasystem
        TeclarTxt "44", 16, 16
        ENTER
        Esperasystem
        
        
        While Range("e" & linha1) <> ""
        
        TeclarTxt "01", 18, 13
        TeclarTxt WorksheetFunction.Text(Range("e" & linha1), "000000000"), 21, 13
        ENTER
        Esperasystem
        
                        If Trim(Copiar(10, 18, 37)) = "Atencao! Cliente com divida ajuizada." Then
                        F3
                        Esperasystem
                        F3
                        Esperasystem
                        End If
        
                If Trim(Copiar(10, 14, 4)) = "Essa" Then
                F3
                Esperasystem
                End If
        
                If Trim(Copiar(10, 12, 19)) = "Informação restrita" Then
                TeclarTxt "Analise de operacao Rural " & Range("e10"), 12, 12
                ENTER
                Esperasystem
                End If
                
                linha = 9
                
                If Trim(Copiar(9, 3, 7)) = "DEBITOS" Then
                VDAF = VDAF + Trim(Copiar(linha, 65, 16))
                End If
               
                linha2 = 108
                linha = 8
      
             Do While linha < 20
              
                      
                linha = linha + 1
   
                        If Trim(Copiar(linha, 3, 12)) = "COOBRIGACOES" Then

                        
                            Do While Range("e" & linha2) <> ""
                        
                                    If Range("e" & linha2) <> Range("e" & linha1) Then
                                                                          
                                    CONJCO = False
                                    
                                    End If
                                    
                                    If Range("e" & linha2) = Range("e" & linha1) Then
                                                                          
                                    CONJCO = True
                                    
                                    End If
                                    
                                    linha2 = linha2 + 1
                         
                            
                            Loop
                            
                            If CONJCO = False Then
                            
                            VCAF = VCAF + Trim(Copiar(linha, 65, 15))
                            
                            Else
                            
                            VCAF = VCAF + 0
                            
                            End If
                            
   
                            
                        End If
  

                    If linha > 19 Then
                        F8
                        Esperasystem
                        linha = 8
                        If Trim(Copiar(24, 3, 6)) = "Ultima" Then
                            linha = 21
                            linha1 = linha1 + 1
                            F3
                            Esperasystem
                        End If
                        
                    
                    End If
                
                   
                
                Loop

        Wend
         
         TDCAF = VDAF + VCAF
         
         Range("e123") = TDCAF
         
         
  
               
                
        While (Trim(Copiar(1, 2, 8))) <> "SBBP6130"
           F3
           Esperasystem
        Wend

        
'10 - Pegando dados Operação'

If Range("e3") = 0 Then
TeclarTxt "27", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "03", 21, 20
        ENTER
        Esperasystem
GoTo Proximo2
End If

        TeclarTxt "27", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "03", 21, 20
        ENTER
        Esperasystem
        TeclarTxt "13", 19, 24
        TeclarTxt WorksheetFunction.Text(Range("E10"), "000000000"), 20, 24
        ENTER
        Esperasystem
        TeclarTxt WorksheetFunction.Text(Range("E10"), "000000000"), 7, 50
        ENTER
        Esperasystem
       'pegando risco, cobertura e customizacao'
        linha = 6
        While Trim(Copiar(linha, 5, 6)) <> "SUMULA"
        linha = linha + 1
        Wend
        TeclarTxt "x", linha, 3
        ENTER
        Esperasystem
        While (Trim(Copiar(23, 2, 3))) <> "Ult"
        For cont = 4 To 22
            If (Trim(Copiar(cont, 10, 8))) = "original" Then
            Range("E55") = Trim(Copiar(cont, 24, 2))
            End If
            If (Trim(Copiar(cont, 3, 12))) = "CUSTOMIZACAO" Then
            Range("E53") = Copiar(cont + 1, 3, 3)
            End If
            If (Trim(Copiar(cont, 13, 10))) = "DESCOBERTO" Then
            Range("E54") = Trim(Copiar(cont, 40, 14))
            End If
            If (Trim(Copiar(cont, 51, 9))) = "Sublimite" Then
            Range("e73") = Trim(Copiar(cont, 66, 14))
            End If
            Next
            F8
            Esperasystem
            Wend
        F3
        Esperasystem
        F3
        Esperasystem
        F3
        Esperasystem
        TeclarTxt "22", 19, 24
        TeclarTxt WorksheetFunction.Text(Range("E10"), "000000000"), 20, 24
        ENTER
        Esperasystem
        TeclarTxt "11", 17, 24
        TeclarTxt WorksheetFunction.Text(Range("E10"), "0000"), 19, 57
        ENTER
        Esperasystem
        TeclarTxt "11", 21, 47
        ENTER
        Esperasystem
        Range("E17") = Trim(Copiar(4, 21, 55))
        Range("e14") = Trim(Copiar(9, 28, 14))
        F3
        Esperasystem
        'analise'
        TeclarTxt "21", 21, 47
        ENTER
        Esperasystem
        linha1 = 13
        While linha1 <> 19 And Trim(Copiar(linha1, 5, 7)) <> "ANALISE"
        linha1 = linha1 + 1
        Wend
        Range("e23") = Trim(Copiar(linha1 + 1, 35, 4))
        F8
        Esperasystem
        linha1 = 13
        While linha1 <> 19 And Trim(Copiar(linha1, 5, 7)) <> "PARECER"
        linha1 = linha1 + 1
        Wend
        Range("e4") = Trim(Copiar(linha1, 66, 10))
        F3
        Esperasystem
        
        'indicadores de atenção'
       
        'Garantias'
        TeclarTxt "15", 21, 47
        ENTER
        Esperasystem
        linha1 = 10
        linha = 33
        While Trim(Copiar(linha1, 5, 1)) <> "" And linha1 < 20
        Range("E" & linha) = Trim(Copiar(linha1, 5, 66))
        linha1 = linha1 + 1
        linha = linha + 1
        Wend
        F3
        Esperasystem
        
        'ITENS FINANCIADOS'
        TeclarTxt "14", 21, 47
        ENTER
        Esperasystem
        linha1 = 10
        linha = 43
        While Trim(Copiar(linha1, 5, 1)) <> "" And linha1 < 20
        Range("E" & linha) = Trim(Copiar(linha1, 5, 40))
        linha1 = linha1 + 1
        linha2 = linha2 + 1
        Wend
       
        F3
        Esperasystem
        F3
        Esperasystem
        F3
        Esperasystem
        
Proximo2:
       
'11 - Pegando tempo de Operação'

        TeclarTxt "99", 19, 24
        TeclarTxt WorksheetFunction.Text(Range("e10"), "000000000"), 20, 24
        ENTER
        Esperasystem
        TeclarTxt "19", 21, 18
        ENTER
        Esperasystem
        TeclarTxt "17", 20, 18
        ENTER
        Esperasystem
        TeclarTxt WorksheetFunction.Text(Range("E93"), "000000000"), 13, 48
        ENTER
        Esperasystem
        
            If Trim(Copiar(16, 24, 3)) = "Nao" Then
            GoTo proximo90
            End If
            
            
        
        Ano = 0
        While Trim(Copiar(23, 3, 3)) <> "Ult"
            For cont = 10 To 19
                If cont = 10 And Ano = 0 Then
                    Ano = Trim(Copiar(cont, 45, 4))
                End If
                If Trim(Copiar(cont, 45, 4)) < Ano Then
                    Ano = Trim(Copiar(cont, 45, 4))
                End If
            Next
            F8
            Esperasystem
        Wend
        Range("e22") = Ano
                      
If Range("e3") = 0 Then
GoTo proximo3
proximo90:
End If

                     
        While (Trim(Copiar(1, 3, 8))) <> "COPM0000"
           F3
           Esperasystem
        Wend
        
       
        
        '12 - preenchendo Despacho
        
        If Range("e3") = "15a" Then
        
            TeclarTxt "11", 19, 24
            ENTER
            Esperasystem
                'If Trim(Copiar(14, 22, 3)) = "PF3" Then
                'F3
                'Esperasystem
            
            TeclarTxt "15", 19, 27
            TeclarTxt "a", 19, 43
            ENTER
            Esperasystem
            TeclarTxt "11", 20, 24
            ENTER
            Esperasystem
                    
                                        
                    linha1 = 10
                    
                    Do While linha1 < 21
                    
                    
                        If Replace(Trim(Copiar(linha1, 5, 11)), ".", "") = Range("e10") Then
                        TeclarTxt "x", linha1, 3
                        ENTER
                        Esperasystem
                        ENTER
                        Esperasystem
                        linha1 = 22
                        End If
                        
                        If linha1 = 20 Then
                        F8
                        Esperasystem
                        linha1 = 9
                            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
                            linha1 = 22
                            End If
                            
                        End If
                                           
                        linha1 = linha1 + 1
                    
                    Loop
                    
                    F2
                    Esperasystem
                    F3
                    Esperasystem
                    
                    TeclarTxt "S", 6, 53
                    TeclarTxt "N", 7, 79
                    
                    TeclarTxt "1. DO PARECER DA AGÊNCIA: ", 9, 10
                    TeclarTxt "A. CÁLCULO CAPACIDADE DE PAGAMENTO: CONFORME SÚMULA DA PROPOSTA.", 10, 10
                    TeclarTxt "B. APRESENTAÇÃO DAS CERTIDÕES PREVISTAS NA IN 528-1 DISPENSADAS PELA", 11, 10
                    TeclarTxt "AGÊNCIA.", 12, 10
                    TeclarTxt "2. CONSIDERAÇÕES FINAIS:", 13, 10
                    TeclarTxt "CABERÁ A AGÊNCIA A VERIFICAÇÃO DOS SEGUINTES ITENS A ÉPOCA DA LIBE-", 14, 10
                    TeclarTxt "RAÇÃO DOS RECURSOS:", 15, 10
                    TeclarTxt "A. MANTER ATUALIZADO O CADASTRO DO PROPONENTE E/OU DO GRUPO AGROPECU-", 16, 10
                    TeclarTxt "ÁRIO, BEM COMO DE TODOS INTERVENIENTES, CONFORME IN 529-1 1.2.1.2.", 17, 10
                    TeclarTxt "B. NA OCORRÊNCIA DE FATOS QUE POSSAM ALTERAR O RISCO DO CLIENTE OU DA", 18, 10
                    TeclarTxt "OPERAÇÃO, A AGÊNCIA DEVERÁ REANALISAR OU SOLICITAR À INSTÂNCIA DEVIDA", 19, 10
                    TeclarTxt "A REANÁLISE DO LIMITE DE CRÉDITO IN 677-1 1.2.2.", 20, 10
                                        
                    F8
                    Esperasystem
                    
                    TeclarTxt "C. LEMBRAMOS QUE CABE AO COMITÊ DE CRÉDITO DA DEPENDÊNCIA A RESPONSA-", 9, 10
                    TeclarTxt "SABILIDADE PELA GARANTIA DA QUALIDADE DO PROCESSO DECISÓRIO, CONFORME", 10, 10
                    TeclarTxt "IN 285-1 4.2.3.", 11, 10
                    TeclarTxt "D. ADOTAR MEDIDAS DA IN 677-1 NO QUE TANGE AO GERENCIAMENTO DA OPERA-", 12, 10
                    TeclarTxt "ÇÃO E DOCUMENTAÇÃO NECESSÁRIA PARA SUA FORMALIZAÇÃO. TAL PROCEDIMENTO", 13, 10
                    TeclarTxt "DIMINUI OS RISCOS AOS QUAIS O BANCO ESTÁ EXPOSTO NO DESENVOLVIMENTO DE", 14, 10
                    TeclarTxt "SUAS ATIVIDADES, GARANTINDO O RETORNO DE SEU PATRIMÔNIO.", 15, 10
                    TeclarTxt "E. OBSERVAR AS ORIENTAÇÕES ESTRATÉGICAS, ORÇAMENTÁRIAS E DE MITIGADO-", 16, 10
                    TeclarTxt "RES DE RISCO PARA AS OPERAÇÕES CONTRATADAS COM RECURSOS CONTROLADOS,", 17, 10
                    TeclarTxt "EXPLICITADAS NA IN 607-1.", 18, 10
                    TeclarTxt "F. ATENTAR PARA TODAS AS NORMAS GERAIS E ESPECÍFICAS DA LINHA DE CRÉ-", 19, 10
                    TeclarTxt "DITO.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "G. ATENDER, OBRIGATORIAMENTE, A TODAS AS CONDICIONANTES PREVISTAS NO", 9, 10
                    TeclarTxt "LIMITE DE CRÉDITO, CONFORME IN 672-1 1.4.1.", 10, 10
                    TeclarTxt "H. ATENTAR PARA IN 425-1: PREVENÇÃO E COMBATE À LAVAGEM DE DINHEIRO.", 11, 10
                    TeclarTxt "I. ATENTAR PARA IN 417-1, ESPECIALMENTE QUANDO HOUVER PROCURAÇÃO(ÕES).", 12, 10
                    TeclarTxt "J. CONTRATAÇÃO CONDICIONADA À INEXISTÊNCIA DE OPERAÇÕES VENCIDAS,", 13, 10
                    TeclarTxt "RESTRIÇÕES IMPEDITIVAS EM NOME DO CLIENTE, DE PESSOAS JURÍDICAS DAS", 14, 10
                    TeclarTxt "QUAIS O CLIENTE SEJA SÓCIO/PROPRIETÁRIO E/OU DIRIGENTE E INTERVENIEN-", 15, 10
                    TeclarTxt "TE, INCLUSIVE SE BAIXADAS PELO CÓDIGO DE DEFESA CONSUMIDOR, CONFORME", 16, 10
                    TeclarTxt "IN 528-2 1.1.1.6.1.", 17, 10
                    TeclarTxt "K. A CONTRATAÇÃO EXIGE AUSÊNCIA DE ANOTAÇÕES CADASTRAIS IMPEDITIVAS", 18, 10
                    TeclarTxt "RELATIVAS OU ABSOLUTAS DO PROPONENTE E INTERVENIENTES, COMO COOBRIGA-", 19, 10
                    TeclarTxt "DOS, CÔNJUGES, MEMBROS DO GRUPO AGROPECUÁRIO, ETC.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "L. A AGÊNCIA DEVERÁ CONDICIONAR A CONTRATAÇÃO DO CRÉDITO AO CUMPRIMEN-", 9, 10
                    TeclarTxt "TO DE TODAS CONDICIONANTES DA NOTA TÉCNICA, CASO EXISTA.", 10, 10
                    TeclarTxt "M. LICENCIAMENTO AMBIENTAL E OUTORGA D ÁGUA: OBSERVAR O PRECONIZADO NA", 11, 10
                    TeclarTxt "IN 671-1, TANTO NO QUE SE REFERE A SUA EXIGÊNCIA OU NÃO, BEM COMO SI-", 12, 10
                    TeclarTxt "TUAÇÕES PASSÍVEIS DE DISPENSA E DE CADASTRO NO ÓRGÃO AMBIENTAL.", 13, 10
                    TeclarTxt "N. DOCUMENTOS: A AGÊNCIA É RESPONSÁVEL PELA INCLUSÃO DE TODOS DOCUMEN-", 14, 10
                    TeclarTxt "TOS OBRIGATÓRIOS NO DEOC, BEM COMO DAQUELES QUE, EMBORA PASSÍVEIS DE", 15, 10
                    TeclarTxt "DISPENSA, NÃO FORAM DISPENSADOS NA NOTA TÉCNICA, PARECER OU DESPACHO,", 16, 10
                    TeclarTxt "SOBRETUDO OS CITADOS NAS INSTRUÇÕES NORMATIVAS 1, 528 E 673.", 17, 10
                    TeclarTxt "O. A AGÊNCIA É RESPONSÁVEL PELA EXATIDÃO DOS DADOS CADASTRADOS NO SIS-", 18, 10
                    TeclarTxt "TEMA, DEVENDO OBSERVAR SUA ESTRITA CONSONÂNCIA COM OS DOCUMENTOS PER-", 19, 10
                    TeclarTxt "TINENTES, RESPEITANDO O CONTIDO NA IN 529-1 ITEM 1.1.8.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "P. TARIFA DE ESTUDO: OBSERVAR A OBRIGATORIEDADE OU NÃO DE SUA COBRANÇA", 9, 10
                    TeclarTxt "CONFORME INSTRUÇÕES NORMATIVAS 1, 4, 6, 7, 10 e 849.", 10, 10
                       
                    ENTER
                    Esperasystem
                                        
                    TeclarTxt "s", 21, 45
                    ENTER
                    Esperasystem
                   
            While (Trim(Copiar(1, 3, 8))) <> "COPM1720"
            F3
            Esperasystem
            Wend
                       
        
        End If

        
        If Range("e3") = "14a" Then
        
            TeclarTxt "11", 19, 24
            ENTER
            Esperasystem
                'If Trim(Copiar(14, 22, 3)) = "PF3" Then
                'F3
                'Esperasystem
            
            TeclarTxt "14", 19, 27
            TeclarTxt "a", 19, 43
            ENTER
            Esperasystem
            TeclarTxt "11", 20, 24
           
            ENTER
            Esperasystem
                    
                 
                    linha1 = 10
                    
                     Do While linha1 < 21
                    
                    
                        If Replace(Trim(Copiar(linha1, 5, 11)), ".", "") = Range("e10") Then
                        TeclarTxt "x", linha1, 3
                        ENTER
                        Esperasystem
                        ENTER
                        Esperasystem
                            If Trim(Copiar(11, 18, 8)) = "Proposta" Then
                            F3
                            Esperasystem
                            End If
                        
                        
                        linha1 = 22
                        End If
                        
                        If linha1 = 20 Then
                        F8
                        Esperasystem
                        linha1 = 9
                            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
                            linha1 = 22
                            End If
                            
                        End If
                                           
                        linha1 = linha1 + 1
                    
                    Loop
                    
                    F2
                    Esperasystem
                    F3
                    Esperasystem
                    
                    TeclarTxt "S", 6, 53
                    TeclarTxt "N", 7, 79
                    
                    TeclarTxt "1. DO PARECER DA AGÊNCIA: ", 9, 10
                    TeclarTxt "A. CÁLCULO CAPACIDADE DE PAGAMENTO: CONFORME SÚMULA DA PROPOSTA.", 10, 10
                    TeclarTxt "B. APRESENTAÇÃO DAS CERTIDÕES PREVISTAS NA IN 528-1 DISPENSADAS PELA", 11, 10
                    TeclarTxt "AGÊNCIA.", 12, 10
                    TeclarTxt "2. CONSIDERAÇÕES FINAIS:", 13, 10
                    TeclarTxt "CABERÁ A AGÊNCIA A VERIFICAÇÃO DOS SEGUINTES ITENS A ÉPOCA DA LIBE-", 14, 10
                    TeclarTxt "RAÇÃO DOS RECURSOS:", 15, 10
                    TeclarTxt "A. MANTER ATUALIZADO O CADASTRO DO PROPONENTE E/OU DO GRUPO AGROPECU-", 16, 10
                    TeclarTxt "ÁRIO, BEM COMO DE TODOS INTERVENIENTES, CONFORME IN 529-1 1.2.1.2.", 17, 10
                    TeclarTxt "B. NA OCORRÊNCIA DE FATOS QUE POSSAM ALTERAR O RISCO DO CLIENTE OU DA", 18, 10
                    TeclarTxt "OPERAÇÃO, A AGÊNCIA DEVERÁ REANALISAR OU SOLICITAR À INSTÂNCIA DEVIDA", 19, 10
                    TeclarTxt "A REANÁLISE DO LIMITE DE CRÉDITO IN 677-1 1.2.2.", 20, 10
                                        
                    F8
                    Esperasystem
                    
                    TeclarTxt "C. LEMBRAMOS QUE CABE AO COMITÊ DE CRÉDITO DA DEPENDÊNCIA A RESPONSA-", 9, 10
                    TeclarTxt "SABILIDADE PELA GARANTIA DA QUALIDADE DO PROCESSO DECISÓRIO, CONFORME", 10, 10
                    TeclarTxt "IN 285-1 4.2.3.", 11, 10
                    TeclarTxt "D. ADOTAR MEDIDAS DA IN 677-1 NO QUE TANGE AO GERENCIAMENTO DA OPERA-", 12, 10
                    TeclarTxt "ÇÃO E DOCUMENTAÇÃO NECESSÁRIA PARA SUA FORMALIZAÇÃO. TAL PROCEDIMENTO", 13, 10
                    TeclarTxt "DIMINUI OS RISCOS AOS QUAIS O BANCO ESTÁ EXPOSTO NO DESENVOLVIMENTO DE", 14, 10
                    TeclarTxt "SUAS ATIVIDADES, GARANTINDO O RETORNO DE SEU PATRIMÔNIO.", 15, 10
                    TeclarTxt "E. OBSERVAR AS ORIENTAÇÕES ESTRATÉGICAS, ORÇAMENTÁRIAS E DE MITIGADO-", 16, 10
                    TeclarTxt "RES DE RISCO PARA AS OPERAÇÕES CONTRATADAS COM RECURSOS CONTROLADOS,", 17, 10
                    TeclarTxt "EXPLICITADAS NA IN 607-1.", 18, 10
                    TeclarTxt "F. ATENTAR PARA TODAS AS NORMAS GERAIS E ESPECÍFICAS DA LINHA DE CRÉ-", 19, 10
                    TeclarTxt "DITO.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "G. ATENDER, OBRIGATORIAMENTE, A TODAS AS CONDICIONANTES PREVISTAS NO", 9, 10
                    TeclarTxt "LIMITE DE CRÉDITO, CONFORME IN 672-1 1.4.1.", 10, 10
                    TeclarTxt "H. ATENTAR PARA IN 425-1: PREVENÇÃO E COMBATE À LAVAGEM DE DINHEIRO.", 11, 10
                    TeclarTxt "I. ATENTAR PARA IN 417-1, ESPECIALMENTE QUANDO HOUVER PROCURAÇÃO(ÕES).", 12, 10
                    TeclarTxt "J. CONTRATAÇÃO CONDICIONADA À INEXISTÊNCIA DE OPERAÇÕES VENCIDAS,", 13, 10
                    TeclarTxt "RESTRIÇÕES IMPEDITIVAS EM NOME DO CLIENTE, DE PESSOAS JURÍDICAS DAS", 14, 10
                    TeclarTxt "QUAIS O CLIENTE SEJA SÓCIO/PROPRIETÁRIO E/OU DIRIGENTE E INTERVENIEN-", 15, 10
                    TeclarTxt "TE, INCLUSIVE SE BAIXADAS PELO CÓDIGO DE DEFESA CONSUMIDOR, CONFORME", 16, 10
                    TeclarTxt "IN 528-2 1.1.1.6.1.", 17, 10
                    TeclarTxt "K. A CONTRATAÇÃO EXIGE AUSÊNCIA DE ANOTAÇÕES CADASTRAIS IMPEDITIVAS", 18, 10
                    TeclarTxt "RELATIVAS OU ABSOLUTAS DO PROPONENTE E INTERVENIENTES, COMO COOBRIGA-", 19, 10
                    TeclarTxt "DOS, CÔNJUGES, MEMBROS DO GRUPO AGROPECUÁRIO, ETC.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "L. A AGÊNCIA DEVERÁ CONDICIONAR A CONTRATAÇÃO DO CRÉDITO AO CUMPRIMEN-", 9, 10
                    TeclarTxt "TO DE TODAS CONDICIONANTES DA NOTA TÉCNICA, CASO EXISTA.", 10, 10
                    TeclarTxt "M. LICENCIAMENTO AMBIENTAL E OUTORGA D ÁGUA: OBSERVAR O PRECONIZADO NA", 11, 10
                    TeclarTxt "IN 671-1, TANTO NO QUE SE REFERE A SUA EXIGÊNCIA OU NÃO, BEM COMO SI-", 12, 10
                    TeclarTxt "TUAÇÕES PASSÍVEIS DE DISPENSA E DE CADASTRO NO ÓRGÃO AMBIENTAL.", 13, 10
                    TeclarTxt "N. DOCUMENTOS: A AGÊNCIA É RESPONSÁVEL PELA INCLUSÃO DE TODOS DOCUMEN-", 14, 10
                    TeclarTxt "TOS OBRIGATÓRIOS NO DEOC, BEM COMO DAQUELES QUE, EMBORA PASSÍVEIS DE", 15, 10
                    TeclarTxt "DISPENSA, NÃO FORAM DISPENSADOS NA NOTA TÉCNICA, PARECER OU DESPACHO,", 16, 10
                    TeclarTxt "SOBRETUDO OS CITADOS NAS INSTRUÇÕES NORMATIVAS 1, 528 E 673.", 17, 10
                    TeclarTxt "O. A AGÊNCIA É RESPONSÁVEL PELA EXATIDÃO DOS DADOS CADASTRADOS NO SIS-", 18, 10
                    TeclarTxt "TEMA, DEVENDO OBSERVAR SUA ESTRITA CONSONÂNCIA COM OS DOCUMENTOS PER-", 19, 10
                    TeclarTxt "TINENTES, RESPEITANDO O CONTIDO NA IN 529-1 ITEM 1.1.8.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "P. TARIFA DE ESTUDO: OBSERVAR A OBRIGATORIEDADE OU NÃO DE SUA COBRANÇA", 9, 10
                    TeclarTxt "CONFORME INSTRUÇÕES NORMATIVAS 1, 4, 6, 7, 10 e 849.", 10, 10
                     
                    
                    ENTER
                    Esperasystem
                                        
                    TeclarTxt "s", 21, 45
                    ENTER
                    Esperasystem
                   
            While (Trim(Copiar(1, 3, 8))) <> "COPM1720"
            F3
            Esperasystem
            Wend
                
                       
        
        End If
        
                                                                                
        If Range("e3") = "15b" Then
        
            TeclarTxt "11", 19, 24
            ENTER
            Esperasystem
                'If Trim(Copiar(14, 22, 3)) = "PF3" Then
                'F3
                'Esperasystem
            
            TeclarTxt "15", 19, 27
            TeclarTxt "b", 19, 43
            ENTER
            Esperasystem
            TeclarTxt "11", 20, 24
            'TeclarTxt WorksheetFunction.Text(Range("e10"), "000000000"), 21, 24
            ENTER
            Esperasystem
                    
                'End If
                                    
                    linha1 = 10
                    
                     Do While linha1 < 21
                    
                    
                        If Replace(Trim(Copiar(linha1, 5, 11)), ".", "") = Range("e10") Then
                        TeclarTxt "x", linha1, 3
                        ENTER
                        Esperasystem
                        ENTER
                        Esperasystem
                        linha1 = 22
                        End If
                        
                        If linha1 = 20 Then
                        F8
                        Esperasystem
                        linha1 = 9
                            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
                            linha1 = 22
                            End If
                            
                        End If
                                           
                        linha1 = linha1 + 1
                    
                    Loop
                    
                    F2
                    Esperasystem
                    F3
                    Esperasystem
                    
                    TeclarTxt "S", 6, 53
                    TeclarTxt "N", 7, 79
                    
                    TeclarTxt "1. DO PARECER DA AGÊNCIA: ", 9, 10
                    TeclarTxt "A. CÁLCULO CAPACIDADE DE PAGAMENTO: CONFORME SÚMULA DA PROPOSTA.", 10, 10
                    TeclarTxt "B. APRESENTAÇÃO DAS CERTIDÕES PREVISTAS NA IN 528-1 DISPENSADAS PELA", 11, 10
                    TeclarTxt "AGÊNCIA.", 12, 10
                    TeclarTxt "2. CONSIDERAÇÕES FINAIS:", 13, 10
                    TeclarTxt "CABERÁ A AGÊNCIA A VERIFICAÇÃO DOS SEGUINTES ITENS A ÉPOCA DA LIBE-", 14, 10
                    TeclarTxt "RAÇÃO DOS RECURSOS:", 15, 10
                    TeclarTxt "A. MANTER ATUALIZADO O CADASTRO DO PROPONENTE E/OU DO GRUPO AGROPECU-", 16, 10
                    TeclarTxt "ÁRIO, BEM COMO DE TODOS INTERVENIENTES, CONFORME IN 529-1 1.2.1.2.", 17, 10
                    TeclarTxt "B. NA OCORRÊNCIA DE FATOS QUE POSSAM ALTERAR O RISCO DO CLIENTE OU DA", 18, 10
                    TeclarTxt "OPERAÇÃO, A AGÊNCIA DEVERÁ REANALISAR OU SOLICITAR À INSTÂNCIA DEVIDA", 19, 10
                    TeclarTxt "A REANÁLISE DO LIMITE DE CRÉDITO IN 677-1 1.2.2.", 20, 10
                                        
                    F8
                    Esperasystem
                    
                    TeclarTxt "C. LEMBRAMOS QUE CABE AO COMITÊ DE CRÉDITO DA DEPENDÊNCIA A RESPONSA-", 9, 10
                    TeclarTxt "SABILIDADE PELA GARANTIA DA QUALIDADE DO PROCESSO DECISÓRIO, CONFORME", 10, 10
                    TeclarTxt "IN 285-1 4.2.3.", 11, 10
                    TeclarTxt "D. ADOTAR MEDIDAS DA IN 677-1 NO QUE TANGE AO GERENCIAMENTO DA OPERA-", 12, 10
                    TeclarTxt "ÇÃO E DOCUMENTAÇÃO NECESSÁRIA PARA SUA FORMALIZAÇÃO. TAL PROCEDIMENTO", 13, 10
                    TeclarTxt "DIMINUI OS RISCOS AOS QUAIS O BANCO ESTÁ EXPOSTO NO DESENVOLVIMENTO DE", 14, 10
                    TeclarTxt "SUAS ATIVIDADES, GARANTINDO O RETORNO DE SEU PATRIMÔNIO.", 15, 10
                    TeclarTxt "E. OBSERVAR AS ORIENTAÇÕES ESTRATÉGICAS, ORÇAMENTÁRIAS E DE MITIGADO-", 16, 10
                    TeclarTxt "RES DE RISCO PARA AS OPERAÇÕES CONTRATADAS COM RECURSOS CONTROLADOS,", 17, 10
                    TeclarTxt "EXPLICITADAS NA IN 607-1.", 18, 10
                    TeclarTxt "F. ATENTAR PARA TODAS AS NORMAS GERAIS E ESPECÍFICAS DA LINHA DE CRÉ-", 19, 10
                    TeclarTxt "DITO.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "G. ATENDER, OBRIGATORIAMENTE, A TODAS AS CONDICIONANTES PREVISTAS NO", 9, 10
                    TeclarTxt "LIMITE DE CRÉDITO, CONFORME IN 672-1 1.4.1.", 10, 10
                    TeclarTxt "H. ATENTAR PARA IN 425-1: PREVENÇÃO E COMBATE À LAVAGEM DE DINHEIRO.", 11, 10
                    TeclarTxt "I. ATENTAR PARA IN 417-1, ESPECIALMENTE QUANDO HOUVER PROCURAÇÃO(ÕES).", 12, 10
                    TeclarTxt "J. CONTRATAÇÃO CONDICIONADA À INEXISTÊNCIA DE OPERAÇÕES VENCIDAS,", 13, 10
                    TeclarTxt "RESTRIÇÕES IMPEDITIVAS EM NOME DO CLIENTE, DE PESSOAS JURÍDICAS DAS", 14, 10
                    TeclarTxt "QUAIS O CLIENTE SEJA SÓCIO/PROPRIETÁRIO E/OU DIRIGENTE E INTERVENIEN-", 15, 10
                    TeclarTxt "TE, INCLUSIVE SE BAIXADAS PELO CÓDIGO DE DEFESA CONSUMIDOR, CONFORME", 16, 10
                    TeclarTxt "IN 528-2 1.1.1.6.1.", 17, 10
                    TeclarTxt "K. A CONTRATAÇÃO EXIGE AUSÊNCIA DE ANOTAÇÕES CADASTRAIS IMPEDITIVAS", 18, 10
                    TeclarTxt "RELATIVAS OU ABSOLUTAS DO PROPONENTE E INTERVENIENTES, COMO COOBRIGA-", 19, 10
                    TeclarTxt "DOS, CÔNJUGES, MEMBROS DO GRUPO AGROPECUÁRIO, ETC.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "L. A AGÊNCIA DEVERÁ CONDICIONAR A CONTRATAÇÃO DO CRÉDITO AO CUMPRIMEN-", 9, 10
                    TeclarTxt "TO DE TODAS CONDICIONANTES DA NOTA TÉCNICA, CASO EXISTA.", 10, 10
                    TeclarTxt "M. LICENCIAMENTO AMBIENTAL E OUTORGA D ÁGUA: OBSERVAR O PRECONIZADO NA", 11, 10
                    TeclarTxt "IN 671-1, TANTO NO QUE SE REFERE A SUA EXIGÊNCIA OU NÃO, BEM COMO SI-", 12, 10
                    TeclarTxt "TUAÇÕES PASSÍVEIS DE DISPENSA E DE CADASTRO NO ÓRGÃO AMBIENTAL.", 13, 10
                    TeclarTxt "N. DOCUMENTOS: A AGÊNCIA É RESPONSÁVEL PELA INCLUSÃO DE TODOS DOCUMEN-", 14, 10
                    TeclarTxt "TOS OBRIGATÓRIOS NO DEOC, BEM COMO DAQUELES QUE, EMBORA PASSÍVEIS DE", 15, 10
                    TeclarTxt "DISPENSA, NÃO FORAM DISPENSADOS NA NOTA TÉCNICA, PARECER OU DESPACHO,", 16, 10
                    TeclarTxt "SOBRETUDO OS CITADOS NAS INSTRUÇÕES NORMATIVAS 1, 528 E 673.", 17, 10
                    TeclarTxt "O. A AGÊNCIA É RESPONSÁVEL PELA EXATIDÃO DOS DADOS CADASTRADOS NO SIS-", 18, 10
                    TeclarTxt "TEMA, DEVENDO OBSERVAR SUA ESTRITA CONSONÂNCIA COM OS DOCUMENTOS PER-", 19, 10
                    TeclarTxt "TINENTES, RESPEITANDO O CONTIDO NA IN 529-1 ITEM 1.1.8.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "P. TARIFA DE ESTUDO: OBSERVAR A OBRIGATORIEDADE OU NÃO DE SUA COBRANÇA", 9, 10
                    TeclarTxt "CONFORME INSTRUÇÕES NORMATIVAS 1, 4, 6, 7, 10 e 849.", 10, 10
                    
                    ENTER
                    Esperasystem
                                        
                    TeclarTxt "s", 21, 45
                    ENTER
                    Esperasystem
                   
            While (Trim(Copiar(1, 3, 8))) <> "COPM1720"
            F3
            Esperasystem
            Wend
                
        
        End If
        
      
        If Range("e3") = "14b" Then
        
            TeclarTxt "11", 19, 24
            ENTER
            Esperasystem
                'If Trim(Copiar(14, 22, 3)) = "PF3" Then
                'F3
                'Esperasystem
            
            TeclarTxt "14", 19, 27
            TeclarTxt "b", 19, 43
            ENTER
            Esperasystem
            TeclarTxt "11", 20, 24
            ENTER
            Esperasystem
                    
                                      
                    linha1 = 10
                    
                   Do While linha1 < 21
                    
                    
                        If Replace(Trim(Copiar(linha1, 5, 11)), ".", "") = Range("e10") Then
                        TeclarTxt "x", linha1, 3
                        ENTER
                        Esperasystem
                        ENTER
                        Esperasystem
                        linha1 = 22
                        End If
                        
                        If linha1 = 20 Then
                        F8
                        Esperasystem
                        linha1 = 9
                            If Trim(Copiar(23, 3, 6)) = "Ultima" Then
                            linha1 = 22
                            End If
                            
                        End If
                                           
                        linha1 = linha1 + 1
                    
                    Loop
                    
                    F2
                    Esperasystem
                    F3
                    Esperasystem
                    
                    TeclarTxt "S", 6, 53
                    TeclarTxt "N", 7, 79
                    
                    TeclarTxt "1. DO PARECER DA AGÊNCIA: ", 9, 10
                    TeclarTxt "A. CÁLCULO CAPACIDADE DE PAGAMENTO: CONFORME SÚMULA DA PROPOSTA.", 10, 10
                    TeclarTxt "B. APRESENTAÇÃO DAS CERTIDÕES PREVISTAS NA IN 528-1 DISPENSADAS PELA", 11, 10
                    TeclarTxt "AGÊNCIA.", 12, 10
                    TeclarTxt "2. CONSIDERAÇÕES FINAIS:", 13, 10
                    TeclarTxt "CABERÁ A AGÊNCIA A VERIFICAÇÃO DOS SEGUINTES ITENS A ÉPOCA DA LIBE-", 14, 10
                    TeclarTxt "RAÇÃO DOS RECURSOS:", 15, 10
                    TeclarTxt "A. MANTER ATUALIZADO O CADASTRO DO PROPONENTE E/OU DO GRUPO AGROPECU-", 16, 10
                    TeclarTxt "ÁRIO, BEM COMO DE TODOS INTERVENIENTES, CONFORME IN 529-1 1.2.1.2.", 17, 10
                    TeclarTxt "B. NA OCORRÊNCIA DE FATOS QUE POSSAM ALTERAR O RISCO DO CLIENTE OU DA", 18, 10
                    TeclarTxt "OPERAÇÃO, A AGÊNCIA DEVERÁ REANALISAR OU SOLICITAR À INSTÂNCIA DEVIDA", 19, 10
                    TeclarTxt "A REANÁLISE DO LIMITE DE CRÉDITO IN 677-1 1.2.2.", 20, 10
                                        
                    F8
                    Esperasystem
                    
                    TeclarTxt "C. LEMBRAMOS QUE CABE AO COMITÊ DE CRÉDITO DA DEPENDÊNCIA A RESPONSA-", 9, 10
                    TeclarTxt "SABILIDADE PELA GARANTIA DA QUALIDADE DO PROCESSO DECISÓRIO, CONFORME", 10, 10
                    TeclarTxt "IN 285-1 4.2.3.", 11, 10
                    TeclarTxt "D. ADOTAR MEDIDAS DA IN 677-1 NO QUE TANGE AO GERENCIAMENTO DA OPERA-", 12, 10
                    TeclarTxt "ÇÃO E DOCUMENTAÇÃO NECESSÁRIA PARA SUA FORMALIZAÇÃO. TAL PROCEDIMENTO", 13, 10
                    TeclarTxt "DIMINUI OS RISCOS AOS QUAIS O BANCO ESTÁ EXPOSTO NO DESENVOLVIMENTO DE", 14, 10
                    TeclarTxt "SUAS ATIVIDADES, GARANTINDO O RETORNO DE SEU PATRIMÔNIO.", 15, 10
                    TeclarTxt "E. OBSERVAR AS ORIENTAÇÕES ESTRATÉGICAS, ORÇAMENTÁRIAS E DE MITIGADO-", 16, 10
                    TeclarTxt "RES DE RISCO PARA AS OPERAÇÕES CONTRATADAS COM RECURSOS CONTROLADOS,", 17, 10
                    TeclarTxt "EXPLICITADAS NA IN 607-1.", 18, 10
                    TeclarTxt "F. ATENTAR PARA TODAS AS NORMAS GERAIS E ESPECÍFICAS DA LINHA DE CRÉ-", 19, 10
                    TeclarTxt "DITO.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "G. ATENDER, OBRIGATORIAMENTE, A TODAS AS CONDICIONANTES PREVISTAS NO", 9, 10
                    TeclarTxt "LIMITE DE CRÉDITO, CONFORME IN 672-1 1.4.1.", 10, 10
                    TeclarTxt "H. ATENTAR PARA IN 425-1: PREVENÇÃO E COMBATE À LAVAGEM DE DINHEIRO.", 11, 10
                    TeclarTxt "I. ATENTAR PARA IN 417-1, ESPECIALMENTE QUANDO HOUVER PROCURAÇÃO(ÕES).", 12, 10
                    TeclarTxt "J. CONTRATAÇÃO CONDICIONADA À INEXISTÊNCIA DE OPERAÇÕES VENCIDAS,", 13, 10
                    TeclarTxt "RESTRIÇÕES IMPEDITIVAS EM NOME DO CLIENTE, DE PESSOAS JURÍDICAS DAS", 14, 10
                    TeclarTxt "QUAIS O CLIENTE SEJA SÓCIO/PROPRIETÁRIO E/OU DIRIGENTE E INTERVENIEN-", 15, 10
                    TeclarTxt "TE, INCLUSIVE SE BAIXADAS PELO CÓDIGO DE DEFESA CONSUMIDOR, CONFORME", 16, 10
                    TeclarTxt "IN 528-2 1.1.1.6.1.", 17, 10
                    TeclarTxt "K. A CONTRATAÇÃO EXIGE AUSÊNCIA DE ANOTAÇÕES CADASTRAIS IMPEDITIVAS", 18, 10
                    TeclarTxt "RELATIVAS OU ABSOLUTAS DO PROPONENTE E INTERVENIENTES, COMO COOBRIGA-", 19, 10
                    TeclarTxt "DOS, CÔNJUGES, MEMBROS DO GRUPO AGROPECUÁRIO, ETC.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "L. A AGÊNCIA DEVERÁ CONDICIONAR A CONTRATAÇÃO DO CRÉDITO AO CUMPRIMEN-", 9, 10
                    TeclarTxt "TO DE TODAS CONDICIONANTES DA NOTA TÉCNICA, CASO EXISTA.", 10, 10
                    TeclarTxt "M. LICENCIAMENTO AMBIENTAL E OUTORGA D ÁGUA: OBSERVAR O PRECONIZADO NA", 11, 10
                    TeclarTxt "IN 671-1, TANTO NO QUE SE REFERE A SUA EXIGÊNCIA OU NÃO, BEM COMO SI-", 12, 10
                    TeclarTxt "TUAÇÕES PASSÍVEIS DE DISPENSA E DE CADASTRO NO ÓRGÃO AMBIENTAL.", 13, 10
                    TeclarTxt "N. DOCUMENTOS: A AGÊNCIA É RESPONSÁVEL PELA INCLUSÃO DE TODOS DOCUMEN-", 14, 10
                    TeclarTxt "TOS OBRIGATÓRIOS NO DEOC, BEM COMO DAQUELES QUE, EMBORA PASSÍVEIS DE", 15, 10
                    TeclarTxt "DISPENSA, NÃO FORAM DISPENSADOS NA NOTA TÉCNICA, PARECER OU DESPACHO,", 16, 10
                    TeclarTxt "SOBRETUDO OS CITADOS NAS INSTRUÇÕES NORMATIVAS 1, 528 E 673.", 17, 10
                    TeclarTxt "O. A AGÊNCIA É RESPONSÁVEL PELA EXATIDÃO DOS DADOS CADASTRADOS NO SIS-", 18, 10
                    TeclarTxt "TEMA, DEVENDO OBSERVAR SUA ESTRITA CONSONÂNCIA COM OS DOCUMENTOS PER-", 19, 10
                    TeclarTxt "TINENTES, RESPEITANDO O CONTIDO NA IN 529-1 ITEM 1.1.8.", 20, 10
                    
                    F8
                    Esperasystem
                    
                    TeclarTxt "P. TARIFA DE ESTUDO: OBSERVAR A OBRIGATORIEDADE OU NÃO DE SUA COBRANÇA", 9, 10
                    TeclarTxt "CONFORME INSTRUÇÕES NORMATIVAS 1, 4, 6, 7, 10 e 849.", 10, 10

                    
                    ENTER
                    Esperasystem
                                        
                    TeclarTxt "s", 21, 45
                    ENTER
                    Esperasystem
                   
            While (Trim(Copiar(1, 3, 8))) <> "COPM1720"
            F3
            Esperasystem
            Wend
                       
        
        End If

        
'--------------------------------------------------------------------
' pegando dados no sistema GRT - inoformações sobre as garantias.
ProximoFim:
proximo80:
proximo3:
     While (Trim(Copiar(1, 2, 4))) <> "BB30"
           F5
           Esperasystem
        
        Wend
        
      capturagrt
        
      End If
      
      
      
      Call Gerar_PDF("Analise_Op_Agro_Grupo", Planilha6.Range("A" & linha7))
      
      linha7 = linha7 + 1
      
      Loop
      
      MsgBox "Captura finalizada"
End Sub
      
Public Sub capturagrt()

    If Not Conectar Then Exit Sub
    Dim linha As Long
    Dim linha1 As Long
    Dim linha2 As Long
    Dim linha3 As Long
    Dim linha4 As Long
    Dim linha5 As Long
    Dim linha6 As Long
    Dim linha7 As Long
    Dim Pdihip As Double
    Dim Gdihip As Double
    
    
        DoEvents
    
    ok = cursor(15, 14)
    Teclar ("grt")
    
    ok = cursor(16, 14)
    Teclar (Senha)
    
    ENTER
    Esperasystem
        
      
            
                TeclarTxt "21", 19, 18
                ENTER
                Esperasystem
                    
                    linha = 63
                    Gdihip = 0
        
                    Do While Range("E" & linha) <> 0
                
                        TeclarTxt "01", 21, 18
                        ENTER
                        Esperasystem
                        TeclarTxt WorksheetFunction.Text(Range("e" & linha), "000000000"), 20, 36
                        ENTER
                        Esperasystem
                
                        linha1 = 2
                        linha3 = 9
                        linha4 = 12
                
                            If Trim(Copiar(23, 3, 3)) = "Não" Then
                            linha3 = 22
                            F3
                            Esperasystem
                            End If
       
                
                
                Do While linha3 < 21
                
                    For linha3 = 9 To 20
                    
                    If Trim(Copiar(linha3, 15, 20)) <> "" Then

                            TeclarTxt "x", linha3, 4
                            ENTER
                            Esperasystem
                        
                            linha4 = 12
                        
                        Do While linha4 < 22
                        
                                If Trim(Copiar(linha4, 25, 12)) <> "IMOVEL RURAL" Then
                                linha4 = linha4 + 1
                                
                                    If linha4 > 21 Then
                                    F8
                                    Esperasystem
                                    linha4 = 12
                                        If Trim(Copiar(23, 3, 1)) = "Ú" Then
                                        linha4 = 23
                                        F3
                                        Esperasystem
                                        End If
                                    End If
                                Else
                                Gdihip = Gdihip + Trim(Copiar(8, 61, 20))
                                F3
                                Esperasystem
                                linha4 = 23
          
                                End If

                    
                        Loop
  
                    Else
                    linha3 = 21
                    End If
                          
                Next
                    
                    F8
                    Esperasystem
                    linha3 = 9
                        If Trim(Copiar(23, 3, 1)) = "Ú" Then
                        linha3 = 22
                        F3
                        
                        Esperasystem
                        End If
                 Loop
           linha = linha + 1
           Loop
           
           Range("E95") = Gdihip
           
           
           
                TeclarTxt "01", 21, 18
                ENTER
                Esperasystem
                TeclarTxt WorksheetFunction.Text(Range("S3"), "000000000"), 20, 36
                ENTER
                Esperasystem
                
                linha1 = 2
                linha3 = 9
                linha4 = 12
                Pdihip = 0
                
                If Trim(Copiar(23, 3, 3)) = "Não" Then
                linha3 = 22
                End If
                
                Do While linha3 < 21
                
                    For linha3 = 9 To 20
                    
                    If Trim(Copiar(linha3, 15, 20)) <> "" Then

                            TeclarTxt "x", linha3, 4
                            ENTER
                            Esperasystem
                        
                            linha4 = 12
                        
                        Do While linha4 < 22
                        
                                If Trim(Copiar(linha4, 25, 12)) <> "IMOVEL RURAL" Then
                                linha4 = linha4 + 1
                                
                                    If linha4 > 21 Then
                                    F8
                                    Esperasystem
                                    linha4 = 12
                                        If Trim(Copiar(23, 3, 1)) = "Ú" Then
                                        linha4 = 23
                                        F3
                                        Esperasystem
                                        End If
                                    End If
                                Else
                                Pdihip = Pdihip + Trim(Copiar(8, 61, 20))
                                F3
                                Esperasystem
                                linha4 = 23
          
                                End If

                    
                        Loop
  
                    Else
                    linha3 = 21
                    End If
                          
                Next
                    
                    F8
                    Esperasystem
                    linha3 = 9
                        If Trim(Copiar(23, 3, 1)) = "Ú" Then
                        linha3 = 22
                        F3
                        Esperasystem
                        
                        End If
            Loop
        
        Range("E94") = Pdihip
            
                'alteração no valor dos bens hipotecados

                TeclarTxt "01", 21, 18
                ENTER
                Esperasystem
                TeclarTxt WorksheetFunction.Text(Range("S3"), "000000000"), 20, 36
                ENTER
                Esperasystem
                
                linha1 = 2
                linha3 = 9
                linha4 = 12
                linha5 = 124
                linha6 = 124
                nbem = 0
                Vagarp = 0
                
                If Trim(Copiar(23, 3, 3)) = "Não" Then
                linha3 = 22
                End If
                
                Do While linha3 < 21
                
                    For linha3 = 9 To 20
                    
                    If Trim(Copiar(linha3, 15, 20)) <> "" Then

                            TeclarTxt "x", linha3, 4
                            ENTER
                            Esperasystem
                        
                            linha4 = 12
                        
                        Do While linha4 < 22
                        
                                If Trim(Copiar(linha4, 25, 12)) <> "IMOVEL RURAL" Then
                                linha4 = linha4 + 1
                                
                                    If linha4 > 21 Then
                                    F8
                                    Esperasystem
                                    linha4 = 12
                                        If Trim(Copiar(23, 3, 1)) = "Ú" Then
                                        linha4 = 23
                                        F3
                                        Esperasystem
                                        End If
                                    End If
                                Else
                                
                                TeclarTxt "x", linha4, 4
                                ENTER
                                Esperasystem
                                
                               
                                linha6 = 124
                                
                                
                                nbem = Trim(Copiar(17, 22, 9))
                                
                                    Do While Range("E" & linha6) <> "0" And linha6 < 139
                                     
                    
                                        If nbem <> Range("E" & linha6) Then
                                        
                                        linha6 = linha6 + 1
                                        
                                        bem = False
                                        
                                        End If
                                    
                                        If nbem = Range("E" & linha6) Then
                                        
                                        linha6 = 140
                                        
                                        bem = True
                                        
                                        End If
                                     
                                    
                                    Loop
            
                                        If bem = False Then
                                        
                                        Range("E" & linha5) = Trim(Copiar(17, 22, 9))
                                            
                                        Vagarp = Vagarp + Trim(Copiar(17, 63, 17))
                                        
                                        linha5 = linha5 + 1
                                        
                                        End If
                                        
                                        F3
                                        Esperasystem
                                        
                                        linha4 = linha4 + 1
                                            
                                    
                                    
                                                                    
          
                                End If
                                
                        Loop
  
                    Else
                    linha3 = 21
                    End If
                          
                Next
                    
                    F8
                    Esperasystem
                    linha3 = 9
                        If Trim(Copiar(23, 3, 1)) = "Ú" Then
                        linha3 = 22
                        F3
                        Esperasystem
                        
                        End If
            Loop
        
        Range("E151") = Vagarp
        
        linha6 = 124
        
        While linha6 < 150
        Range("e" & linha6) = del
        linha6 = linha6 + 1
        
        Wend
        
        linha = 63
        linha1 = 2
        
        linha4 = 12
        linha5 = 124
        linha6 = 124
        nbem = 0
        Vagarg = 0
        
        
        Do While Range("E" & linha) <> 0
          
                TeclarTxt "01", 21, 18
                ENTER
                Esperasystem
                            
                
                TeclarTxt WorksheetFunction.Text(Range("e" & linha), "000000000"), 20, 36
                
                linha3 = 9
                
                ENTER
                Esperasystem
                
             
                
                If Trim(Copiar(23, 3, 3)) = "Não" Then
                linha3 = 22
                F3
                Esperasystem
                End If
                
                Do While linha3 < 21
                
                    For linha3 = 9 To 20
                    
                    If Trim(Copiar(linha3, 15, 20)) <> "" Then

                            TeclarTxt "x", linha3, 4
                            ENTER
                            Esperasystem
                        
                            linha4 = 12
                        
                        Do While linha4 < 22
                        
                                If Trim(Copiar(linha4, 25, 12)) <> "IMOVEL RURAL" Then
                                linha4 = linha4 + 1
                                
                                    If linha4 > 21 Then
                                    F8
                                    Esperasystem
                                    linha4 = 12
                                        If Trim(Copiar(23, 3, 1)) = "Ú" Then
                                        linha4 = 23
                                        F3
                                        Esperasystem
                                        End If
                                    End If
                                Else
                                
                                TeclarTxt "x", linha4, 4
                                ENTER
                                Esperasystem
                                
                               
                                linha6 = 124
                                
                                
                                nbem = Trim(Copiar(17, 22, 9))
                                
                                    Do While Range("E" & linha6) <> "0" And linha6 < 139
                                     
                    
                                        If nbem <> Range("E" & linha6) Then
                                        
                                        linha6 = linha6 + 1
                                        
                                        bem = False
                                        
                                        End If
                                    
                                        If nbem = Range("E" & linha6) Then
                                        
                                        linha6 = 140
                                        
                                        bem = True
                                        
                                        End If
                                     
                                    
                                    Loop
            
                                        If bem = False Then
                                        
                                        Range("E" & linha5) = Trim(Copiar(17, 22, 9))
                                            
                                        Vagarg = Vagarg + Trim(Copiar(17, 63, 17))
                                        
                                        linha5 = linha5 + 1
                                        
                                        End If
                                        
                                        F3
                                        Esperasystem
                                        
                                        linha4 = linha4 + 1
                                            
                                    
                                    
                                                                    
          
                                End If
                                
                        Loop
  
                    Else
                    linha3 = 21
                    End If
                          
                Next
                    
                    F8
                    Esperasystem
                    linha3 = 9
                        If Trim(Copiar(23, 3, 1)) = "Ú" Then
                        linha3 = 22
                        F3
                        Esperasystem
                        
                        End If
            Loop
           linha = linha + 1
        Loop
        
        Range("E152") = Vagarg
        
        
        While (Trim(Copiar(1, 2, 4))) <> "BB30"
        F5
        Esperasystem
        Wend

          
      'pegando o valor de pacote de serviços
          
          DoEvents
    
    ok = cursor(15, 14)
    Teclar ("Tarifas")
    
    ok = cursor(16, 14)
    Teclar (Senha)
    
    ENTER
    Esperasystem
    
    ok = Aguardar(1, 3, "TFAM0000")
    
    TeclarTxt "44", 21, 20
    Esperasystem
    ENTER
    Esperasystem
    F3
    Esperasystem
    TeclarTxt "91", 21, 20
    ENTER
    Esperasystem
            
    TeclarTxt WorksheetFunction.Text(Planilha4.Range("K" & linha1), "0000"), 4, 30
    TeclarTxt WorksheetFunction.Text(Planilha4.Range("L" & linha1), "0000000000"), 5, 30
    ENTER
    Esperasystem
    PTARIT = Trim(Copiar(7, 46, 5))
    Range("E76") = PTARIT
    
    While (Trim(Copiar(1, 2, 4))) <> "BB30"
    F3
    Esperasystem
    Wend
    ok = Aguardar(1, 2, "BB30")

    'Range("E170") = Trim(Copiar(1, 71, 10))
    
    ENTER
    Esperasystem
  
    'Call Gerar_PDF("Analise_Op_Agro_Grupo", Planilha6.Range("A" & linha7))
     
    'Planilha6.Range("B" & linha7) = ""
    
    'End If
    
    
   ' linha7 = linha7 + 1
    
     'If Planilha6.Range("B" & linha7) = "sim" Then
     'MsgBox "Operação já foi analisada"
     'End If
    
'Loop



  While (Trim(Copiar(1, 2, 4))) <> "BB30"
    F5
    Esperasystem
  Wend

'MsgBox "Captura finalizada"

End Sub




