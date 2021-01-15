Attribute VB_Name = "FUN��ES"
Option Compare Database   'Usar ordem do banco de dados para compara��es
Option Explicit

'Declara fun��es da API do Windows (SetCaption)
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Sub SetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)

'Declara fun��es avulsas da API do Windows

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: IsIconic
' Finalidade: Retorna True se o form (report) estiver maximizado
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Declare Function IsIconic Lib "User" (ByVal hWnd As Integer) As Integer

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: IsZoomed
' Finalidade: Retorna True se o form (report) estiver minimizado
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Declare Function IsZoomed Lib "User" (ByVal hWnd As Integer) As Integer

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: aaa
' Finalidade: teste
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function aaa(strN�mero As String) As Integer
   'Define vari�veis
   Dim strNum As String
   Dim intC As Integer
   Dim strDV As String
   Dim intAcum As Integer
   
   Dim intResto As Integer
   Dim strDV1 As String
   Dim strDV2 As String
   'Tira espa�os em branco do n�mero
   strNum = Trim(strN�mero)
   'Verifica se s� possui n�meros
   If Not IsDig(strNum) Then
      aaa = False
      Exit Function
   End If
   'Verifica tamanho da string
   If Len(strNum) < 3 Or Len(strNum) > 11 Then
      aaa = False
      Exit Function
   End If
   'Inclui zeros para ficar com 11 d�gitos
   Do While Len(strNum) <> 11
      strNum = "0" & strNum
   Loop
   'Separa o n�mero dos d�gitos
   strDV = Right$(strNum, 2)
   strNum = Left$(strNum, 9)
   'Multiplica cada n�mero e acumula
   intAcum = 0
   For intC = 1 To 9
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * (11 - intC))
   Next
   'Acha o resto da divis�o por 11 que � o DV1
   intResto = intAcum Mod 11
   'Primeiro DV
   strDV1 = Trim(IIf(intResto = 10, "0", Str$(intResto)))
   'Soma DV1 a strNum para repetir o c�lculo
   strNum = strNum & strDV1
   'Multiplica cada n�mero e acumula
   intAcum = 0
   For intC = 1 To 10
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * (12 - intC))
   Next
   'Acha o resto da divis�o por 11 que � o DV2
   intResto = intAcum Mod 11
   'Retorna o d�gito verificador
   strDV2 = Trim(IIf(intResto = 10, "0", Str$(intResto)))
   'Compara com os DVs e retorna True
   If strDV = (strDV1 & strDV2) Then
      aaa = True
   Else
      aaa = False
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: CalcDV
' Finalidade: Calcula d�gito verificador (m�dulo 11) de strN�mero
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CalcDV(strN�mero As String) As String
   'Define vari�veis
   Dim strNum As String
   Dim strCalc As String
   Dim intC As Integer
   Dim intAcum As Integer
   Dim intResto As Integer
   'Tira espa�os em branco do n�mero
   strNum = Trim(strN�mero)
   'Retorna string nula se foi passada string nula
   If Len(strNum) = 0 Then
      CalcDV = ""
      Exit Function
   End If
   'Verifica se s� possui n�meros
   If Not IsDig(strNum) Then
      CalcDV = ""
      Exit Function
   End If
   'String para c�lculo
   strCalc = "23456789"
   'Aumenta tamanho da string de c�lculo
   Do While Len(strNum) > Len(strCalc)
      strCalc = strCalc & strCalc
   Loop
   'Deixa a string para c�lculo com o mesmo tamanho do n�mero
   strCalc = Right$(strCalc, Len(strNum))
   'Multiplica string para c�lculo com n�mero e acumula
   For intC = 1 To Len(strNum)
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * Val(Mid$(strCalc, intC, 1)))
   Next
   'Acha o resto da divis�o por 11 que � o DV
   intResto = intAcum Mod 11
   'Retorna o d�gito verificador
   CalcDV = Trim(IIf(intResto = 10, "X", Str$(intResto)))
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: CloseBD
' Finalidade: Fecha Banco de Dados, verificando se est� no ADT
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CloseBD() As Integer
   Dim varRet As Variant
   'Verifica se ADT est� sob ADT e sai
   If IsADT() Then
      'Encerra sess�o do Access
      DoCmd.Quit
   Else
      'Habilita barras de ferramentas
      varRet = EnableToolBar(True)
      'Retorna t�tulo original na janela
      varRet = SetCaption("")
      'Fecha aplica��o
      varRet = FechaBD()
   End If
   'Retorna True
   CloseBD = True
End Function

Function C�dGMR(txtParam As String) As String

Dim compr As Integer, inttam As Integer, aux  As String

inttam = Len(Trim(txtParam))

Rem intTam = intNivel * 2
aux = "000000000000000000000000000000"
C�dGMR = Trim(txtParam) & Mid$(aux, 1, (26 - inttam))

End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: CompactMDB
' Finalidade: Compacta Banco de Dados
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CompactMDB(strOldName As String, strNewName As String) As Integer
   'Desvia para pr�xima linha em caso de erro
   On Error Resume Next
   'Compacta Banco de Dados
   DBEngine.CompactDatabase strOldName, strNewName
   'Se n�o houver erro retorna True
   If Err = 0 Then CompactMDB = True Else CompactMDB = False
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: CountStr
' Finalidade: Calcula quantas vezes strProc aparece em strAlvo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CountStr(strProc As String, strAlvo As String) As Long
   'Defini��o de vari�veis
   Dim intC As Long
   Dim intAcum As Long
   'Inicializa valor acumulado
   intAcum = 0
   'Se um dos par�metros for uma string nula retorna 0
   If strProc = "" Or strAlvo = "" Then
      CountStr = intAcum
      Exit Function
   End If
   'Conta quantos strProc h� em strAlvo e acumula
   For intC = 1 To Len(strAlvo)
      If InStr(intC, strAlvo, strProc, 2) > 0 Then
         intAcum = intAcum + 1
         intC = InStr(intC, strAlvo, strProc, 2) + Len(strProc) - 1
      End If
   Next
   'Retorna valor acumulado
   CountStr = intAcum
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: CurDrive
' Finalidade: Retorna a letra do drive corrente
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CurDrive() As String
   CurDrive = Left$(CurDir, 1)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: EnableStatusBar
' Finalidade: Habilita/desabilita a barra de status do sistema
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function EnableStatusBar(intEnable As Integer) As Integer
   'Habilita/desabilita a barra de status do sistema
   Application.SetOption "Show Status Bar", intEnable
   'Redesenha objeto
   DoCmd.RepaintObject
   'Retorna intEnable
   EnableStatusBar = intEnable
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: EnableToolBar
' Finalidade: Habilita/desabilita as barras de ferramentas do
'             sistema, usando o menu Exibir/Op��es
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function EnableToolBar(intEnable As Integer) As Integer
   'Habilita/desabilita barras de ferramentas do sistema
   Application.SetOption "Barras de ferramentas incorporadas dispon�veis", intEnable
   'Redesenha objeto
   DoCmd.RepaintObject
   'Retorna intEnable
   EnableToolBar = intEnable
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: FechaBD
' Finalidade: Fecha janela banco de dados semelhante ao menu
'             Arquivo/Fechar Banco de Dados
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function FechaBD() As Integer
   'Fecha banco de dados
   SendKeys "{F11}", True
   SendKeys "^{F4}", True
   FechaBD = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: Fechar
' Finalidade: Fecha janela ativa
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function Fechar() As Integer
   DoCmd.Close
   Fechar = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: FecharConf
' Finalidade: Fecha janela ativa pedindo confirma��o
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function FecharConf(strMsg As String, strT�tulo As String) As Integer
   If MsgBox(strMsg, 4 + 32 + 256, strT�tulo) = 6 Then
      DoCmd.Close
      FecharConf = True
   Else
      FecharConf = False
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: File
' Finalidade: Verifica se arquivo existe
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function File(strArq As String) As Integer
   'Desvia para pr�xima linha em caso de erro
   On Error Resume Next
   'Verifica se existe arquivo especificado e retorna True
   If Len(Dir$(strArq)) > 0 Then
      File = True
      Exit Function
   End If
   'Retorna False se n�o encontrar arquivo
   File = False
End Function

Function getconfiguracao(strCodigo As String) As String
Dim strAcha As Variant

strAcha = DLookup("ValorConfig", "tblConfig", "CodConfig = '" & strCodigo & "'")
If IsNull(strAcha) Then
    MsgBox "C�digo de configura��o n�o encontrado! Verifique a tabela de configuracao.", 16, "Configura�ao"
    strAcha = ""
End If
getconfiguracao = strAcha
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: GotoFirstRecord
' Finalidade: Vai para o primeiro registro da tabela, consulta ou
'             formul�rio que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoFirstRecord() As Integer
   'Vai para o primeiro registro
   GotoFirstRecord = GotoRecord(A_FIRST)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: GotoLastRecord
' Finalidade: Vai para o �ltimo registro da tabela, consulta ou
'             formul�rio que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoLastRecord() As Integer
   'Vai para o �ltimo registro
   GotoLastRecord = GotoRecord(A_LAST)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: GotoNextRecord
' Finalidade: Vai para o pr�ximo registro da tabela, consulta ou
'             formul�rio que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoNextRecord() As Integer
   'Vai para o pr�ximo registro
   GotoNextRecord = GotoRecord(A_NEXT)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: GotoPrevRecord
' Finalidade: Vai para o registro anterior da tabela, consulta ou
'             formul�rio que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoPrevRecord() As Integer
   'Vai para o registro anterior
   GotoPrevRecord = GotoRecord(A_PREVIOUS)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: GotoRecord
' Finalidade: Vai para o registro indicado por intDire��o da
'             tabela, consulta ou formul�rio que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoRecord(intDire��o As Integer) As Integer
   'Desvia para pr�xima linha em caso de erro
   On Error Resume Next
   'Vai para o registro indicado por intDire��o
   DoCmd.GotoRecord , , intDire��o
   'Redesenha objeto
   DoCmd.RepaintObject
   'Retorna True
   GotoRecord = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: InvStr
' Finalidade: Inverte os caracteres de uma string
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function InvStr(strTexto As String) As String
   'Define vari�veis
   Dim lngC As Long
   Dim strInv As String
   strInv = ""
   'Inverte a string e a retorna
   For lngC = Len(strTexto) To 1 Step -1
      strInv = strInv & Mid$(strTexto, lngC, 1)
   Next
   InvStr = strInv
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: IsADT
' Finalidade: Verifica se o MDB est� sob o ADT ou Access
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsADT() As Integer
   'Retorna True se o MDB estiver sob o ADT
   IsADT = SysCmd(SYSCMD_RUNTIME)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: IsDig
' Finalidade: Verifica se strN�mero � composto somente por
'             d�gitos, ou seja, pelos algarismos de 0 a 9
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsDig(strN�mero As String) As Integer
   Dim strNum As String
   Dim intC As Integer
   'Tira espa�os
   strNum = Trim(strN�mero)
   'Testa tamanho
   If Len(strNum) = 0 Then
      IsDig = False
      Exit Function
   End If
   'Verifica todos os d�gitos
   For intC = 1 To Len(strNum)
      If InStr(1, "0123456789", Mid$(strNum, intC, 1)) = 0 Then
         IsDig = False
         Exit Function
      End If
   Next
   'Retorna True
   IsDig = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: IsLoaded
' Finalidade: Verifica se formul�rio est� carregado
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsLoaded(strForm As String) As Integer
   Dim intI As Integer
   'Localiza formul�rio na cole��o Forms
   For intI = 0 To Forms.Count - 1
      'Se encontrar retorna True
      If Forms(intI).FormName = strForm Then
         IsLoaded = True
         Exit Function
      End If
   Next intI
   'Retorna False se n�o encontrar
   IsLoaded = False
End Function

Function IsQuery(vQuery As String)
   Dim X As Integer
    For X = 0 To meubd.QueryDefs.Count - 1
      'Se encontrar retorna True
      If meubd.QueryDefs(X).Name = vQuery Then
         IsQuery = True
         Exit Function
      End If
   Next

   'Retorna False se n�o encontrar
   IsQuery = False


End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: IsTable
' Finalidade: Verifica se tabela existe no Banco de Dados atual
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsTable(strTable As String) As Integer
   Dim intI As Integer
   'Localiza tabela na cole��o TableDefs
   For intI = 0 To meubd.TableDefs.Count - 1
      'Se encontrar retorna True
      If meubd.TableDefs(intI).Name = strTable Then
         IsTable = True
         Exit Function
      End If
   Next intI
   'Retorna False se n�o encontrar
   IsTable = False
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: IsWindowed
' Finalidade: Retorna True se o form/report n�o estiver
'             maximizado nem minimizado
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsWindowed(ByVal hWnd As Integer) As Integer
   'Verifica se a janela est� minimizada ou maximizada
   If IsZoomed(hWnd) Or IsIconic(hWnd) Then
      IsWindowed = False
   Else
      IsWindowed = True
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: LDM
' Finalidade: Retorna �ltima data do m�s
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function LDM(varDtAtual As Variant) As Variant
   Dim varDtAux As Variant
   'Acrescenta 1 m�s � data original
   varDtAux = DateAdd("m", 1, varDtAtual)
   'Retorna 1 dia da data
   LDM = DateAdd("d", -Day(varDtAux), varDtAux)
End Function

Function MesExtenso(vMeses) As String
Dim sMes As String
    sMes = "JANFEVMARABRMAIJUNJULAGOSETOUTNOVDEZJANFEVMARABRMAIJUN"
    'MesExtenso = Mid(sMes, (vMeses + (vMeses - 1) * 2), 3)
    MesExtenso = Mid(sMes, (vMeses + (vMeses - 1) * 2), 3)

End Function

Function mm(vMeses)
Select Case vMeses
    Case "JAN"
        mm = "01"
    Case "FEV"
        mm = "02"
    Case "MAR"
        mm = "03"
    Case "ABR"
        mm = "04"
    Case "MAI"
        mm = "05"
    Case "JUN"
        mm = "06"
    Case "JUL"
        mm = "07"
    Case "AGO"
        mm = "08"
    Case "SET"
        mm = "09"
    Case "OUT"
        mm = "10"
    Case "NOV"
        mm = "11"
    Case "DEZ"
        mm = "12"

End Select

End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: MsgErro
' Finalidade: Exibe mensagem de erro
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function MsgErro(strMsg As String, intCancela) As Integer
   'Defini��o de vari�veis
   Dim intRet As Integer
   'Verifica se mensagem � vazia
   If Trim(strMsg) = "" Then
      'Mensagem de erro
      intRet = MsgErro("Mensagem n�o pode ser vazia", False)
      'Retorna False
      MsgErro = False
   Else
      'Aviso sonoro
      Beep
      'Exibe mensagem de erro com �cone de aviso
      MsgErro = MsgBox(strMsg, 48 + Abs(intCancela))
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: NewMDB
' Finalidade: Cria um novo MDB
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function NewMDB(strName As String) As Integer
   'Desvia para pr�xima linha em caso de erro
   On Error Resume Next
   Dim db As Database
   'Abre Banco de Dados
   Set db = DBEngine.Workspaces(0).CreateDatabase(strName, DB_LANG_GENERAL)
   'Se n�o houver erro retorna True
   If Err = 0 Then NewMDB = True Else NewMDB = False
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: NomeArq
' Finalidade: Retorna o nome completo do arquivo strArq no
'             diret�rio strDir
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function NomeArq(strDir As String, strArq As String) As String
   'Retorna nome do arquivo com diret�rio
   NomeArq = Trim(strDir) & IIf(Right$(Trim(strDir), 1) = "\", "", "\") & Trim(strArq)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: OcultaJanela
' Finalidade: Oculta janela ativa
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function OcultaJanela() As Integer
   'Oculta janela ativa
   DoCmd.DoMenuItem 1, 4, 3
   'Retorna True
   OcultaJanela = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: OpenBD
' Finalidade: Usada ao abrir MDB para padronizar apresenta��o e
'             verificar presen�a do ADT
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function OpenBD(strCaption As String) As Integer
   Dim varRet As Variant
   'Verifica se est� sob ADT
   If Not IsADT() Then
      'Oculta janela Banco de Dados
      SendKeys "{F11}", True
      DoCmd.MoveSize 3969, 2268, 0, 0
      varRet = OcultaJanela()
      'Inibe barras de ferramentas do sistema
      varRet = EnableToolBar(False)
   End If
   'Retira menu do Access
   Application.MenuBar = "mnuVazio"
   'T�tulo da janela principal
   varRet = SetCaption(strCaption)
   'Retira mensagem da barra de status
   varRet = SetStatusMsg("")
   'Retorna True
   OpenBD = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: SetCaption
' Finalidade: ALtera caption (t�tulo) da janela do Access
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function SetCaption(strCaption As String) As Integer
   'Define vari�veis
   Dim intWnd As Integer
   'Identifica janela do Access
   intWnd = FindWindow("OMain", 0&)
   'Altera caption para strCaption ou retorna ao normal
   If strCaption = "" Then strCaption = "Microsoft Access"
   Call SetWindowText(intWnd, strCaption)
   'Retorna True
   SetCaption = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: SetStatusMsg
' Finalidade: Altera mensagem na barra de status do Access
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function SetStatusMsg(strMsg As String) As Integer
   'Define vari�veis
   Dim varRet As Variant
   'Altera mensagem para strMsg
   If strMsg = "" Then strMsg = " "
   varRet = SysCmd(SYSCMD_SETSTATUS, strMsg)
   'Retorna True
   SetStatusMsg = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: ShowToolBar
' Finalidade: Exibe/oculta as barras de ferramentas do sistema.
'             S� funciona no MS-Access 2.0 em Portugu�s se o
'             mesmo estiver habilitado para exibi-las. Veja
'             fun��o EnableToolBar
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function ShowToolBar(intAtiva As Integer) As Integer
   'Exibe/oculta barras de ferramentas do sistema
   If intAtiva Then
      DoCmd.ShowToolBar "Banco de Dados", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Relacionamentos", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Estrutura da Tabela", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Folha de Dados da Tabela", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Estrutura da Consulta", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Folha de Dados da Consulta", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Estrutura do Formul�rio", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Modo Formul�rio", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Filtro/Classifica��o", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Estrutura do Relat�rio", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Visualizar Impress�o", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Caixa de Ferramentas", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Paleta", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Macro", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "M�dulo", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Microsoft", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Utilit�rio 1", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Utilit�rio 2", A_TOOLBAR_WHERE_APPROP
   Else
      DoCmd.ShowToolBar "Banco de Dados", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Relacionamentos", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Estrutura da Tabela", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Folha de Dados da Tabela", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Estrutura da Consulta", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Folha de Dados da Consulta", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Estrutura do Formul�rio", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Modo Formul�rio", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Filtro/Classifica��o", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Estrutura do Relat�rio", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Visualizar Impress�o", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Caixa de Ferramentas", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Paleta", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Macro", A_TOOLBAR_NO
      DoCmd.ShowToolBar "M�dulo", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Microsoft", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Utilit�rio 1", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Utilit�rio 2", A_TOOLBAR_NO
   End If
   'Retorna o par�metro passado
   ShowToolBar = intAtiva
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: VerCGC
' Finalidade: Verifica se a string passada corresponde a um CGC
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function VerCGC(strN�mero As String) As Integer
   'Define vari�veis
   Dim strNum As String
   Dim strDV1 As String
   Dim strDV2 As String
   'Separa n�mero dos DVs
   strNum = Left$(Trim(strN�mero), Len(Trim(strN�mero)) - 2)
   'Calcula primeiro d�gito
   strDV1 = CalcDV(strNum)
   If strDV1 = "X" Then strDV1 = "0"
   'Calcula segundo d�gito
   strDV2 = CalcDV(strNum & strDV1)
   If strDV2 = "X" Then strDV2 = "0"
   'Compara com os DVs e retorna True
   If Right$(Trim(strN�mero), 2) = (strDV1 & strDV2) Then
      VerCGC = True
   Else
      VerCGC = False
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: VerCPF
' Finalidade: Verifica se a string passada corresponde a um CPF
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function VerCPF(strN�mero As String) As Integer
   'Define vari�veis
   Dim strNum As String
   Dim intC As Integer
   Dim strDV As String
   Dim intAcum As Integer
   Dim intResto As Integer
   Dim strDV1 As String
   Dim strDV2 As String
   'Tira espa�os em branco do n�mero
   strNum = Trim(strN�mero)
   'Verifica se s� possui n�meros
   If Not IsDig(strNum) Then
      VerCPF = False
      Exit Function
   End If
   'Verifica tamanho da string
   If Len(strNum) < 3 Or Len(strNum) > 11 Then
      VerCPF = False
      Exit Function
   End If
   'Inclui zeros para ficar com 11 d�gitos
   Do While Len(strNum) <> 11
      strNum = "0" & strNum
   Loop
   'Separa o n�mero dos d�gitos
   strDV = Right$(strNum, 2)
   strNum = Left$(strNum, 9)
   'Multiplica cada n�mero e acumula
   intAcum = 0
   For intC = 1 To 9
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * (11 - intC))
   Next
   'Acha o resto da divis�o por 11 que � o DV1
   intResto = intAcum Mod 11
   'Primeiro DV
   strDV1 = Trim(IIf(intResto = 10, "0", Str$(intResto)))
   'Soma DV1 a strNum para repetir o c�lculo
   strNum = strNum & strDV1
   'Multiplica cada n�mero e acumula
   intAcum = 0
   For intC = 1 To 10
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * (12 - intC))
   Next
   'Acha o resto da divis�o por 11 que � o DV2
   intResto = intAcum Mod 11
   'Retorna o d�gito verificador
   strDV2 = Trim(IIf(intResto = 10, "0", Str$(intResto)))
   Stop
   'Compara com os DVs e retorna True
   If strDV = (strDV1 & strDV2) Then
      VerCPF = True
   Else
      VerCPF = False
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Fun��o....: VerDV
' Finalidade: Verifica o d�gito verificador (m�dulo 11) de
'             strN�mero
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function VerDV(strN�mero As String) As Integer
   'Define vari�veis
   Dim strNum As String
   Dim strDV As String
   'Tira espa�os a direita e a esquerda
   strNum = Trim(strN�mero)
   'Verifica se n�mero passado � v�lido
   If Len(strNum) < 2 Then
      VerDV = False
      Exit Function
   End If
   'Calcula d�gito e retorna
   If Right$(strNum, 1) = CalcDV(Left$(strNum, Len(strNum) - 1)) Then
      VerDV = True
   Else
      VerDV = False
   End If
End Function

Function VerificaSuper()
Dim db As Database
Dim dep As Recordset
Dim criterio As String
Dim pref As String
Dim Nome As String

pref = getconfiguracao("07")
Nome = getconfiguracao("05")

Set db = DBEngine.Workspaces(0).Databases(0)
Set dep = db.OpenRecordset("DepAux")

criterio = "PREFDEP=pref"

dep.FindFirst "PREFDEP= GetConfiguracao('07')"
If dep.NoMatch Then
   dep.AddNew
   dep![prefdep] = pref
   dep![NOMEDEP] = Nome
   dep.Update
End If
dep.Close

End Function


Public Function Receitas_Super()
Dim X
Dim LC As New IBM3270
Dim item As String
Dim Orca, Reali As Integer


Call AbreBd("Tbl - Receitas", "Tbl - Super", "", "", "")
' tabela1.Index = "primarykey"
tabela2.Index = "primarykey"
With LC
  Do While Not tabela2.EOF
    .Sess�o = "A"
    .Desconectar
    .Conectar
    .Atualizar
    If .Aguardar(1, 3, "ORCM7000", 1, 2) Then
        .Colar 18, 27, "5"
        .Colar 18, 29, "b"
        .Colar 18, 42, tabela2("PREFSUPER")
        .Colar 19, 37, "06"
        .Teclar "@E"
    End If
    .Atualizar
    If .Aguardar(1, 3, "ORCM7510", 1, 2) Then
        .Colar 15, 3, "X"
        .Teclar "@E"
    End If
    .Atualizar
     If .Aguardar(1, 3, "ORCM7511", 1, 2) Then
          X = .Posicionar(10, 4)
          item = .Copiar(10, 4, 25)
          X = .Posicionar(10, 37)
          Orca = .Copiar(10, 37, 12)
          X = .Posicionar(11, 58)
          Reali = item = .Copiar(11, 58, 12)
      End If
          tabela1.AddNew
          tabela1("ItemAval") = item
          tabela1("Super") = tabela2("PREFSUPER")
          tabela1("OrcJun") = Orca
          tabela1("MetaJun") = Reali
          tabela1.Update
          tabela2.MoveNext
     Loop
          tabela1.Close
          tabela2.Close
          
End With
End Function

Function cvalor(texto As String) As Double
   Do While InStr(1, texto, ".") > 0
      texto = Left(texto, InStr(1, texto, ".") - 1) & Mid(texto, InStr(1, texto, ".") + 1)
   Loop
   Do While InStr(1, texto, ",") > 0
      Mid(texto, InStr(1, texto, ","), 1) = "."
   Loop
   
   cvalor = Val(texto)
End Function

Function cdata(ByVal texto As String) As Variant
   If Mid(texto, 3, 1) = "." Then
      Mid(texto, 3, 1) = "/"
   End If
   If Mid(texto, 6, 1) = "." Then
      Mid(texto, 6, 1) = "/"
   End If
'   cdata = CVDate(Texto)
End Function

Public Function FN_LimpaCampos(nForm As Form)
Dim ctl As Control 'varialvel que assume o controle(textbox,combobox,listbox e etc) do form

    For Each ctl In nForm.Controls 'Para cada controle no formul�rio
        'Se o tipo do controle for Caixa de Texto ent�o
        If TypeOf ctl Is TextBox Then
            ctl = Empty
        End If
        
        'Se o tipo do controle for Caixa de Combina��o ent�o
        If TypeOf ctl Is ComboBox Then
            ctl = Empty
        End If
        
        'Se o tipo do controle for Caixa de Listagem ent�o
        If TypeOf ctl Is ListBox Then
            ctl = Empty
        End If
        
        'Se o tipo do controle for Caixa de Sele��o ent�o
        If TypeOf ctl Is CheckBox Then
            ctl = Empty
        End If
        
    Next

End Function


