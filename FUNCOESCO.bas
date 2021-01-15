Attribute VB_Name = "FUNCOESCO"
'------------------------------------------------------------------------------------
'Descrição    :.    MÓDULO API PCCOM E FUNÇÃOES GERAIS
'Criação      :.
'Compilação   :.    07/04/2017 - F3314437 - FLÁVIO STUDART WERNIK
'Atualização  :.    07/04/2017 - F3314437 - FLÁVIO STUDART WERNIK
'------------------------------------------------------------------------------------

'If Not Conectar Then Exit Sub

Option Explicit

Global Sess
Global ConnMgr

Global sessoes_abertas
Global escolha_da_sessao
Global sessao
Global pf As Long
Global conex As Long
Global UsarCIC As String


Global hThread As Long, hProcess As Long
Public Const THREAD_BASE_PRIORITY_LOWRT = 15
Public Const THREAD_BASE_PRIORITY_MIN = -2
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_NORMAL = 0
Public Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long



'Tipo para def
Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Function pastawin() As String
Dim lpIDList As Long
Dim pasta As String
Dim browse As BrowseInfo

lpIDList = SHBrowseForFolder(browse)
If (lpIDList) Then
    pasta = Space(260)
    SHGetPathFromIDList lpIDList, pasta
    pasta = Left(pasta, InStr(pasta, vbNullChar))
    pastawin = pasta
End If
End Function

Public Function Copiar(linha As Long, coluna As Long, tam As Long) As String
      Copiar = Sess.autECLPS.GetText(linha, coluna, tam)
End Function
  
Public Function Conectar() As Boolean

    Set ConnMgr = CreateObject("PCOMM.autECLConnMgr")
    Set Sess = CreateObject("PCOMM.autECLSession")
    
    ConnMgr.autECLConnList.Refresh
    sessoes_abertas = ConnMgr.autECLConnList.Count
    Dim textosessao As String
    Dim idsessao As Long

    
    If sessoes_abertas > 1 Then
        For idsessao = 1 To sessoes_abertas
            On Error Resume Next
            textosessao = textosessao & Chr(13) & idsessao & " >>>>>> Sessão " & ConnMgr.autECLConnList(idsessao).Name
        Next
        escolha_da_sessao = InputBox("Digite um número de sessão entre 1 e " & sessoes_abertas & Chr(13) & Chr(13) & textosessao)
        
        If escolha_da_sessao = "" Then
            MsgBox ("Número inválido")
        ElseIf IsNumeric(escolha_da_sessao) Then
            sessao = CInt(escolha_da_sessao)
        Else
            MsgBox ("A entrada deve ser numérica" & escolha_da_sessao)
        End If
    ElseIf sessoes_abertas = 1 Then
        sessao = 1
    Else
        MsgBox "  APARENTEMENTE  nenhuma  SESSÃO  está  aberta...  ", 48, "  ERRO DE ABERTURA DE SESSÕES. "
        Conectar = False
        Exit Function
    End If
    
    Sess.SetConnectionByHandle (ConnMgr.autECLConnList(sessao).Handle)
    Conectar = True

End Function

Public Function Esperasystem()
     Do While True
        If Sess.autECLOIA.InputInhibited = 0 Then
           Exit Do
        End If
    
    DoEvents
     
     Loop
End Function

Public Function TeclarTxt(texto As String, linha As Long, coluna As Long, Optional tela As String, Optional plinha As Long, Optional pcoluna As Long) As Long
    Sess.autECLPS.SetCursorPos linha, coluna
    Sess.autECLPS.SendKeys (texto)
    DoEvents
    
    If UsarCIC = "" Then
    UsarCIC = "SIM"
    
    End If
    
    If UsarCIC = "NAO" Then
    'espera?
    Randomize
    esperar Rnd * 3
    End If

End Function

Public Function pressionar(tecla As String)
'On Error GoTo fim
Sess.autECLPS.SendKeys (tecla)
Esperasystem
DoEvents

End Function

Public Function Titulo(nometitulo As String)
Sess.autECLWinMetrics.WindowTitle = nometitulo
End Function

Public Function baixa()
AppActivate "Processando..."
hThread = GetCurrentThread
hProcess = GetCurrentProcess
SetThreadPriority hThread, THREAD_PRIORITY_LOWEST
SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
End Function

Public Function rep(valor As String) As Double
Dim letra As Long
For letra = 1 To Len(valor)
    If Right(Left(valor, letra), 1) = "." Then
        valor = Left(valor, letra - 1) & Right(valor, Len(valor) - letra)
    End If
Next
For letra = 1 To Len(valor)
    If Right(Left(valor, letra), 1) = "," Then
        valor = Left(valor, letra - 1) & "." & Right(valor, Len(valor) - letra)
    End If
Next
rep = Val(valor)
End Function

Public Function repl(texto As String, Simb1 As String, Simb2 As String) As String
Dim letra As Long
If Len(Simb1) > 1 Or Len(Simb2) > 1 Then
    MsgBox "Erro repl. 'simb1' e 'simb2' devem ter, no máximo, 1 caracter."
End If
For letra = 1 To Len(texto)
    If Right(Left(texto, letra), 1) = Simb1 Then
        texto = Left(texto, letra - 1) & Simb2 & Right(texto, Len(texto) - letra)
    End If
Next
repl = texto
End Function

Public Function reconect(aplicativo As String, usuario As String, Senha As String, Optional opcao1 As String, Optional opcao2 As String, Optional opcao3 As String, Optional opcao4 As String, Optional opcao5 As String) As Long
If conex = True Then
    Exit Function
End If
pf = 3
While Not Copiar(14, 22, 7) = "SISTEMA"
    If Not pf = 3 Then
        pf = 3
    Else
        pf = 5
    End If
    pressionar "[pf" & pf & "]"
    esperar 2
Wend
    
    If Copiar(14, 22, 7) = "SISTEMA" Then
        pressionar "[ERASEEOF]"
        escrever "SISBB", 20, 39
        pressionar "[enter]"
        esperar 3
        While Not Copiar(14, 5, 5) = "Senha"
            DoEvents
        Wend
        Sess.autECLPS.SetCursorPos 13, 24
        Sess.autECLPS.SendKeys (usuario)
        Sess.autECLPS.SetCursorPos 14, 24
        Sess.autECLPS.SendKeys (Senha)
        Sess.autECLPS.SetCursorPos 15, 24
        Sess.autECLPS.SendKeys (aplicativo)
        pressionar "[enter]"
        esperar 2
        If Copiar(1, 3, 8) = "COEM7010" Then
            pressionar "[pf3]"
        End If
        esperar 3
        If Not IsMissing(opcao1) And Not IsNull(opcao1) And Not opcao1 = "" And Not Len(opcao1) = 0 Then
        pressionar opcao1
        pressionar "[enter]"
        End If
        pressionar "[enter]"
        pressionar "[enter]"
        esperar 1
        If Not IsMissing(opcao2) And Not IsNull(opcao2) And Not opcao2 = "" And Not Len(opcao2) = 0 Then
            pressionar opcao2
            pressionar "[enter]"
        End If
        If Not IsMissing(opcao3) And Not IsNull(opcao3) And Not opcao3 = "" And Not Len(opcao3) = 0 Then
            pressionar opcao3
            pressionar "[enter]"
        End If
        If Not IsMissing(opcao4) And Not IsNull(opcao4) And Not opcao4 = "" And Not Len(opcao4) = 0 Then
            pressionar opcao4
            pressionar "[enter]"
        End If
        If Not IsMissing(opcao5) And Not IsNull(opcao5) And Not opcao5 = "" And Not Len(opcao5) = 0 Then
            pressionar opcao5
            pressionar "[enter]"
        End If
        GoTo fim
    End If
fim:
conex = True
End Function

Public Function esperar(Tempo As Long)
Dim Início
Início = Timer  ' Define a hora inicial.
Do While Timer < Início + Tempo
    DoEvents    ' Submete-se a outros processos.
Loop
End Function
Function ENTER()
    pressionar "[enter]"
End Function
Function F1()
    pressionar "[pf1]"
End Function
Function F2()
    pressionar "[pf2]"
End Function
Function F3()
    pressionar "[pf3]"
End Function
Function F4()
    pressionar "[pf4]"
End Function
Function F5()
    pressionar "[pf5]"
End Function
Function F6()
    pressionar "[pf6]"
End Function
Function F7()
    pressionar "[pf7]"
End Function
Function F8()
    pressionar "[pf8]"
End Function
Function F9()
    pressionar "[pf9]"
End Function
Function F10()
    pressionar "[pf10]"
End Function
Function F11()
    pressionar "[pf11]"
End Function
Function F12()
    pressionar "[pf12]"
End Function
Public Function TiraCaract(strNumero As String, caract As String) As String
    Dim i As Integer
    For i = 1 To Len(strNumero)
        If Mid(strNumero, i, 1) <> caract Then
            TiraCaract = TiraCaract + Mid(strNumero, i, 1)
        End If
    Next
End Function

Public Function TrocaVirgulaPorPonto(strNumber As String) As String
    Dim i As Integer
    For i = 1 To Len(strNumber)
        If Mid(strNumber, i, 1) <> "," And Mid(strNumber, i, 1) <> "." Then
            TrocaVirgulaPorPonto = TrocaVirgulaPorPonto + Mid(strNumber, i, 1)
        Else
            If Mid(strNumber, i, 1) = "," Then
                TrocaVirgulaPorPonto = TrocaVirgulaPorPonto + "."
            End If
        End If
    Next
End Function
Public Function TrocaPontoPorBarra(strNumber As String) As String
    Dim i As Integer
    For i = 1 To Len(strNumber)
        If Mid(strNumber, i, 1) <> "." And Mid(strNumber, i, 1) <> "/" Then
            TrocaPontoPorBarra = TrocaPontoPorBarra + Mid(strNumber, i, 1)
        Else
            If Mid(strNumber, i, 1) = "." Then
                TrocaPontoPorBarra = TrocaPontoPorBarra + "/"
            End If
        End If
    Next
End Function
Public Function TiraAcento(strTexto As String) As String
    Dim i As Integer
    strTexto = UCase(strTexto)
    For i = 1 To Len(strTexto)
        If (Mid(strTexto, i, 1) <> "é") And (Mid(strTexto, i, 1) <> "É") And (Mid(strTexto, i, 1) <> "á") And (Mid(strTexto, i, 1) <> "Á") And (Mid(strTexto, i, 1) <> "ç") And (Mid(strTexto, i, 1) <> "Ç") And (Mid(strTexto, i, 1) <> "Ã") And (Mid(strTexto, i, 1) <> "ã") And (Mid(strTexto, i, 1) <> "ó") And (Mid(strTexto, i, 1) <> "Ó") Then
            TiraAcento = TiraAcento + Mid(strTexto, i, 1)
        Else
            If Mid(strTexto, i, 1) = "É" Or Mid(strTexto, i, 1) = "Ê" Or Mid(strTexto, i, 1) = "È" Then
                TiraAcento = TiraAcento + "E"
            End If
            If Mid(strTexto, i, 1) = "Á" Or Mid(strTexto, i, 1) = "Ã" Or Mid(strTexto, i, 1) = "Â" Or Mid(strTexto, i, 1) = "À" Then
                TiraAcento = TiraAcento + "A"
            End If
            If Mid(strTexto, i, 1) = "Ç" Then
                TiraAcento = TiraAcento + "C"
            End If
            If Mid(strTexto, i, 1) = "Ó" Or Mid(strTexto, i, 1) = "Õ" Then
                TiraAcento = TiraAcento + "O"
            End If
            If Mid(strTexto, i, 1) = "Í" Then
                TiraAcento = TiraAcento + "O"
            End If
        End If
    Next
End Function
Public Function TiraTraco(strTexto As String) As String
    Dim i As Integer
    For i = 1 To Len(strTexto)
        If Mid(strTexto, i, 1) <> "-" Then
            TiraTraco = TiraTraco + Mid(strTexto, i, 1)
        End If
    Next
End Function
Public Function TiraEspaco(strTexto As String) As String
    Dim i As Integer
    For i = 1 To Len(strTexto)
        If Mid(strTexto, i, 1) <> " " Then
            TiraEspaco = TiraEspaco + Mid(strTexto, i, 1)
        End If
    Next
End Function
Public Function TiraUnderline(strTexto As String) As String
    Dim i As Integer
    For i = 1 To Len(strTexto)
        If Mid(strTexto, i, 1) <> "_" Then
            TiraUnderline = TiraUnderline + Mid(strTexto, i, 1)
        End If
    Next
End Function
Public Function TiraPonto(strNumero As String) As String
    Dim i As Integer
    For i = 1 To Len(strNumero)
        If Mid(strNumero, i, 1) <> "." Then
            TiraPonto = TiraPonto + Mid(strNumero, i, 1)
        End If
    Next
End Function
Public Function TiraBarra(strNumero As String) As String
    Dim i As Integer
    For i = 1 To Len(strNumero)
        If Mid(strNumero, i, 1) <> "/" Then
            TiraBarra = TiraBarra + Mid(strNumero, i, 1)
        End If
    Next
End Function
Public Function UltimoDiaMes(Mes, Ano As Integer) As Integer
    Dim i As Integer
    For i = 27 To 32
        If IsDate(CStr(i) & "/" & CStr(Mes) & "/" & CStr(Ano)) Then
            UltimoDiaMes = i
        End If
    Next
End Function
Public Sub Atraso()
    DoEvents
    Randomize
    DoEvents
    Sleep Rnd * 1000
    Randomize
    DoEvents
    Randomize
End Sub
Public Function ObtemNumero(strTexto As String) As String
    Dim i As Integer
    For i = 1 To Len(strTexto)
        If IsNumeric(Mid(strTexto, i, 1)) Then
            ObtemNumero = ObtemNumero + Mid(strTexto, i, 1)
        End If
    Next
End Function
Public Function ConverteDataReduzida(DataReduzida As String)
' converte data no formato ddmmyy para dd/mm/yyyy
    Dim tmp As String
    If Len(DataReduzida) = 6 Then
        tmp = DataReduzida
        DataReduzida = Mid(tmp, 1, 4) & "20" & Mid(tmp, 5, 2)
        tmp = DataReduzida
        DataReduzida = Mid(tmp, 1, 2) & "/" & Mid(tmp, 3, 2) & "/" & Mid(tmp, 5, 9)
    End If
    ConverteDataReduzida = DataReduzida
End Function
Function GetUserName() As String
    Dim lpBuff As String * 25
    Get_User_Name lpBuff, 25
    GetUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function
Public Function CalculaDigito(conta As String) As String
    Dim digito As String
    Dim i, peso, dv As Integer

    dv = 0
    peso = 9

    For i = Len(conta) To 1 Step -1
        dv = dv + (peso * CInt(Mid(conta, i, 1)))
        peso = peso - 1
        If peso = 2 Then
            peso = 9
        End If
        Next i
    dv = dv Mod 11
    digito = CStr(dv)
    If digito = "10" Then
        digito = "X"
    End If
    CalculaDigito = conta & "-" & digito
End Function
Public Function StripChar(Nome As String, caracter As String) As String
    Dim tam As Byte
    Dim i As Byte
    Dim novonome As String
    
    novonome = ""
    tam = Len(Nome)

    For i = 1 To tam
        If Mid(Nome, i, 1) <> caracter Then
            novonome = novonome + Mid(Nome, i, 1)
        End If
    Next i
    StripChar = Trim(novonome)
End Function
Public Function RetornaDiretorioMDB() As String
    Dim sPath As String
    sPath = CurrentDb.Name
    While Right$(sPath, 1) <> "\"
      sPath = Left$(sPath, Len(sPath) - 1)
    Wend
    RetornaDiretorioMDB = sPath
End Function
' Obtem arquivos de um diretorio passado como parâmetro
' e os retorna em um array de strings
Public Function obtemArquivosDiretorio(diretorio As String, retiraExtensao As Boolean) As String()
    Dim Arquivo As Variant
    Dim arquivos() As String, tmp() As String
    Dim nro As Integer, tam As Integer

    ReDim arquivos(1000)

    nro = 0
    Arquivo = Dir(diretorio)
    While Arquivo <> ""
        If Arquivo <> "." And Arquivo <> ".." Then
                arquivos(nro) = Arquivo
                nro = nro + 1
        End If
        Arquivo = Dir
    Wend
    ReDim Preserve arquivos(nro - 1)
    If retiraExtensao = True Then
        For nro = 0 To UBound(arquivos)
            tmp = Split(arquivos(nro), ".")
            arquivos(nro) = tmp(0)
        Next nro
    End If
    obtemArquivosDiretorio = arquivos()
End Function
Public Function RetiraCaracterInvalido(Dado As String, Optional RetiraNumero)
    Dim tam As Byte
    Dim i As Byte
    Dim NovoDado As String
    NovoDado = ""
    tam = Len(Dado)
    If IsMissing(RetiraNumero) Then
        RetiraNumero = False
    End If
    For i = 1 To tam
        If RetiraNumero = False Then
            Select Case Asc(Mid(Dado, i, 1))
                Case 48 To 57 ' 0 a 9
                    NovoDado = NovoDado + Mid(Dado, i, 1)
                Case 65 To 90 ' A a Z
                    NovoDado = NovoDado + Mid(Dado, i, 1)
                Case 97 To 122 ' a a z
                    NovoDado = NovoDado + Mid(Dado, i, 1)
                Case 192 To 254 ' caracteres acentuados
                    Select Case Mid(Dado, i, 1)
                        Case "á"
                            NovoDado = NovoDado + "a"
                        Case "à"
                            NovoDado = NovoDado + "a"
                        Case "ã"
                            NovoDado = NovoDado + "a"
                        Case "â"
                            NovoDado = NovoDado + "a"
                        Case "Á"
                            NovoDado = NovoDado + "A"
                        Case "À"
                            NovoDado = NovoDado + "A"
                        Case "Ã"
                            NovoDado = NovoDado + "A"
                        Case "Â"
                            NovoDado = NovoDado + "A"
                        Case "é"
                            NovoDado = NovoDado + "e"
                        Case "ê"
                            NovoDado = NovoDado + "e"
                        Case "É"
                            NovoDado = NovoDado + "E"
                        Case "Ê"
                            NovoDado = NovoDado + "E"
                        Case "í"
                            NovoDado = NovoDado + "i"
                        Case "Í"
                            NovoDado = NovoDado + "I"
                        Case "ó"
                            NovoDado = NovoDado + "o"
                        Case "õ"
                            NovoDado = NovoDado + "o"
                        Case "Ó"
                            NovoDado = NovoDado + "O"
                        Case "Õ"
                            NovoDado = NovoDado + "O"
                        Case "ú"
                            NovoDado = NovoDado + "u"
                        Case "Ú"
                            NovoDado = NovoDado + "U"
                        Case "ç"
                            NovoDado = NovoDado + "c"
                        Case "Ç"
                            NovoDado = NovoDado + "C"
                        Case Else
                            NovoDado = NovoDado + Mid(Dado, i, 1)
                    End Select
                Case 32 ' espaço
                    NovoDado = NovoDado + Mid(Dado, i, 1)
            End Select
        Else
            Select Case Asc(Mid(Dado, i, 1))
                Case 65 To 90 ' A a Z
                    NovoDado = NovoDado + Mid(Dado, i, 1)
                Case 97 To 122 ' a a z
                    NovoDado = NovoDado + Mid(Dado, i, 1)
                Case 192 To 254 ' caracteres acentuados
                    Select Case Mid(Dado, i, 1)
                        Case "á"
                            NovoDado = NovoDado + "a"
                        Case "à"
                            NovoDado = NovoDado + "a"
                        Case "ã"
                            NovoDado = NovoDado + "a"
                        Case "â"
                            NovoDado = NovoDado + "a"
                        Case "Á"
                            NovoDado = NovoDado + "A"
                        Case "À"
                            NovoDado = NovoDado + "A"
                        Case "Ã"
                            NovoDado = NovoDado + "A"
                        Case "Â"
                            NovoDado = NovoDado + "A"
                        Case "é"
                            NovoDado = NovoDado + "e"
                        Case "ê"
                            NovoDado = NovoDado + "e"
                        Case "É"
                            NovoDado = NovoDado + "E"
                        Case "Ê"
                            NovoDado = NovoDado + "E"
                        Case "í"
                            NovoDado = NovoDado + "i"
                        Case "Í"
                            NovoDado = NovoDado + "I"
                        Case "ó"
                            NovoDado = NovoDado + "o"
                        Case "õ"
                            NovoDado = NovoDado + "o"
                        Case "Ó"
                            NovoDado = NovoDado + "O"
                        Case "Õ"
                            NovoDado = NovoDado + "O"
                        Case "ú"
                            NovoDado = NovoDado + "u"
                        Case "Ú"
                            NovoDado = NovoDado + "U"
                        Case "ç"
                            NovoDado = NovoDado + "c"
                        Case "Ç"
                            NovoDado = NovoDado + "C"
                        Case Else
                            NovoDado = NovoDado + Mid(Dado, i, 1)
                    End Select
                Case 32 ' espaço
                    NovoDado = NovoDado + Mid(Dado, i, 1)
            End Select
        End If
    Next i
    RetiraCaracterInvalido = Trim(NovoDado)
End Function
Public Function CalcPercDesvio(ByVal dblOrc As Double, ByVal dblReal As Double, FormaDeApuração As Integer, Optional blModal As Boolean) As Double
    Select Case FormaDeApuração
        Case 1
            If dblOrc = 0 Then
                CalcPercDesvio = 0
            Else
                CalcPercDesvio = ((dblReal - dblOrc) / IIf(blModal, IIf((dblOrc) < 0, dblOrc * (-1), dblOrc), (dblOrc))) * 100
            End If
        Case 2
            CalcPercDesvio = dblReal - dblOrc
        Case 3
            CalcPercDesvio = dblReal / IIf(dblOrc <> 0, dblOrc, IIf(dblReal = 0, 1, dblReal)) * 100
    End Select
End Function
' conta quantidade de ocorrências de um determinado caractere
Public Function ContaOcorrencias(Nome As String, caracter As String) As Integer
    Dim pos As Integer
    Dim cont As Integer
    
    If Len(Nome) = 0 Then
        ContaOcorrencias = 0
        Exit Function
    End If
    For pos = 1 To Len(Nome)
        If Mid(Nome, pos, 1) = caracter Then
            cont = cont + 1
        End If
    Next pos
    ContaOcorrencias = cont
End Function
'Esta função inclui "n" espaços a esquerda do texto informado
'"n" é a diferença entre o tamanho do campo e o tamanho do texto
Public Function EspaçosEsquerda(texto As String, TamanhoCampo As Integer) As String
    Dim i, Diferenca As Integer
    Diferenca = TamanhoCampo - Len(texto)
    For i = 1 To Diferenca
        texto = " " + texto
    Next
    EspaçosEsquerda = texto
End Function
'Esta função inclui "n" espaços a direita do texto informado
'"n" é a diferença entre o tamanho do campo e o tamanho do texto
Public Function EspaçosDireita(texto As String, TamanhoCampo As Integer) As String
    Dim i, Diferenca As Integer
    
    Diferenca = TamanhoCampo - Len(texto)
    For i = 1 To Diferenca
        texto = texto + " "
    Next
    EspaçosDireita = texto
End Function
'Esta função inclui "n" zeros a esquerda do texto informado
'"n" é a diferença entre o tamanho do campo e o tamanho do texto
Public Function ZerosEsquerda(texto As String, TamanhoCampo As Integer) As String
    Dim i, Diferenca As Integer
    Diferenca = TamanhoCampo - Len(texto)
    For i = 1 To Diferenca
        texto = "0" + texto
    Next
    ZerosEsquerda = texto
End Function
Public Function FiltrarCaracter(Dado As String)
    Dim tam As Byte
    Dim i As Byte
    Dim NovoDado As String
    NovoDado = ""
    tam = Len(Dado)
    For i = 1 To tam
        Select Case Asc(Mid(Dado, i, 1))
            Case 32 To 126
                NovoDado = NovoDado + Mid(Dado, i, 1)
            Case 192 To 254 ' caracteres acentuados
                NovoDado = NovoDado + Mid(Dado, i, 1)
            Case 166 To 167 ' caracteres ª º
                NovoDado = NovoDado + Mid(Dado, i, 1)
            Case 171 To 172 ' caracteres ½ ¼
                NovoDado = NovoDado + Mid(Dado, i, 1)
            Case 184 ' caracteres ©
                NovoDado = NovoDado + Mid(Dado, i, 1)
            Case 243 ' caracteres ¾
                NovoDado = NovoDado + Mid(Dado, i, 1)
            Case 251 To 253 ' caracteres ¹ ² ³
                NovoDado = NovoDado + Mid(Dado, i, 1)
        End Select
    Next i
    FiltrarCaracter = Trim(NovoDado)
End Function






