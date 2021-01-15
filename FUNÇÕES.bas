Attribute VB_Name = "FUNÇÕES"
Option Compare Database   'Usar ordem do banco de dados para comparações
Option Explicit

'Declara funções da API do Windows (SetCaption)
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Sub SetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)

'Declara funções avulsas da API do Windows

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: IsIconic
' Finalidade: Retorna True se o form (report) estiver maximizado
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Declare Function IsIconic Lib "User" (ByVal hWnd As Integer) As Integer

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: IsZoomed
' Finalidade: Retorna True se o form (report) estiver minimizado
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Declare Function IsZoomed Lib "User" (ByVal hWnd As Integer) As Integer

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: aaa
' Finalidade: teste
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function aaa(strNúmero As String) As Integer
   'Define variáveis
   Dim strNum As String
   Dim intC As Integer
   Dim strDV As String
   Dim intAcum As Integer
   
   Dim intResto As Integer
   Dim strDV1 As String
   Dim strDV2 As String
   'Tira espaços em branco do número
   strNum = Trim(strNúmero)
   'Verifica se só possui números
   If Not IsDig(strNum) Then
      aaa = False
      Exit Function
   End If
   'Verifica tamanho da string
   If Len(strNum) < 3 Or Len(strNum) > 11 Then
      aaa = False
      Exit Function
   End If
   'Inclui zeros para ficar com 11 dígitos
   Do While Len(strNum) <> 11
      strNum = "0" & strNum
   Loop
   'Separa o número dos dígitos
   strDV = Right$(strNum, 2)
   strNum = Left$(strNum, 9)
   'Multiplica cada número e acumula
   intAcum = 0
   For intC = 1 To 9
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * (11 - intC))
   Next
   'Acha o resto da divisão por 11 que é o DV1
   intResto = intAcum Mod 11
   'Primeiro DV
   strDV1 = Trim(IIf(intResto = 10, "0", Str$(intResto)))
   'Soma DV1 a strNum para repetir o cálculo
   strNum = strNum & strDV1
   'Multiplica cada número e acumula
   intAcum = 0
   For intC = 1 To 10
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * (12 - intC))
   Next
   'Acha o resto da divisão por 11 que é o DV2
   intResto = intAcum Mod 11
   'Retorna o dígito verificador
   strDV2 = Trim(IIf(intResto = 10, "0", Str$(intResto)))
   'Compara com os DVs e retorna True
   If strDV = (strDV1 & strDV2) Then
      aaa = True
   Else
      aaa = False
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: CalcDV
' Finalidade: Calcula dígito verificador (módulo 11) de strNúmero
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CalcDV(strNúmero As String) As String
   'Define variáveis
   Dim strNum As String
   Dim strCalc As String
   Dim intC As Integer
   Dim intAcum As Integer
   Dim intResto As Integer
   'Tira espaços em branco do número
   strNum = Trim(strNúmero)
   'Retorna string nula se foi passada string nula
   If Len(strNum) = 0 Then
      CalcDV = ""
      Exit Function
   End If
   'Verifica se só possui números
   If Not IsDig(strNum) Then
      CalcDV = ""
      Exit Function
   End If
   'String para cálculo
   strCalc = "23456789"
   'Aumenta tamanho da string de cálculo
   Do While Len(strNum) > Len(strCalc)
      strCalc = strCalc & strCalc
   Loop
   'Deixa a string para cálculo com o mesmo tamanho do número
   strCalc = Right$(strCalc, Len(strNum))
   'Multiplica string para cálculo com número e acumula
   For intC = 1 To Len(strNum)
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * Val(Mid$(strCalc, intC, 1)))
   Next
   'Acha o resto da divisão por 11 que é o DV
   intResto = intAcum Mod 11
   'Retorna o dígito verificador
   CalcDV = Trim(IIf(intResto = 10, "X", Str$(intResto)))
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: CloseBD
' Finalidade: Fecha Banco de Dados, verificando se está no ADT
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CloseBD() As Integer
   Dim varRet As Variant
   'Verifica se ADT está sob ADT e sai
   If IsADT() Then
      'Encerra sessão do Access
      DoCmd.Quit
   Else
      'Habilita barras de ferramentas
      varRet = EnableToolBar(True)
      'Retorna título original na janela
      varRet = SetCaption("")
      'Fecha aplicação
      varRet = FechaBD()
   End If
   'Retorna True
   CloseBD = True
End Function

Function CódGMR(txtParam As String) As String

Dim compr As Integer, inttam As Integer, aux  As String

inttam = Len(Trim(txtParam))

Rem intTam = intNivel * 2
aux = "000000000000000000000000000000"
CódGMR = Trim(txtParam) & Mid$(aux, 1, (26 - inttam))

End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: CompactMDB
' Finalidade: Compacta Banco de Dados
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CompactMDB(strOldName As String, strNewName As String) As Integer
   'Desvia para próxima linha em caso de erro
   On Error Resume Next
   'Compacta Banco de Dados
   DBEngine.CompactDatabase strOldName, strNewName
   'Se não houver erro retorna True
   If Err = 0 Then CompactMDB = True Else CompactMDB = False
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: CountStr
' Finalidade: Calcula quantas vezes strProc aparece em strAlvo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CountStr(strProc As String, strAlvo As String) As Long
   'Definição de variáveis
   Dim intC As Long
   Dim intAcum As Long
   'Inicializa valor acumulado
   intAcum = 0
   'Se um dos parâmetros for uma string nula retorna 0
   If strProc = "" Or strAlvo = "" Then
      CountStr = intAcum
      Exit Function
   End If
   'Conta quantos strProc há em strAlvo e acumula
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
' Função....: CurDrive
' Finalidade: Retorna a letra do drive corrente
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function CurDrive() As String
   CurDrive = Left$(CurDir, 1)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: EnableStatusBar
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
' Função....: EnableToolBar
' Finalidade: Habilita/desabilita as barras de ferramentas do
'             sistema, usando o menu Exibir/Opções
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function EnableToolBar(intEnable As Integer) As Integer
   'Habilita/desabilita barras de ferramentas do sistema
   Application.SetOption "Barras de ferramentas incorporadas disponíveis", intEnable
   'Redesenha objeto
   DoCmd.RepaintObject
   'Retorna intEnable
   EnableToolBar = intEnable
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: FechaBD
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
' Função....: Fechar
' Finalidade: Fecha janela ativa
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function Fechar() As Integer
   DoCmd.Close
   Fechar = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: FecharConf
' Finalidade: Fecha janela ativa pedindo confirmação
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function FecharConf(strMsg As String, strTítulo As String) As Integer
   If MsgBox(strMsg, 4 + 32 + 256, strTítulo) = 6 Then
      DoCmd.Close
      FecharConf = True
   Else
      FecharConf = False
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: File
' Finalidade: Verifica se arquivo existe
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function File(strArq As String) As Integer
   'Desvia para próxima linha em caso de erro
   On Error Resume Next
   'Verifica se existe arquivo especificado e retorna True
   If Len(Dir$(strArq)) > 0 Then
      File = True
      Exit Function
   End If
   'Retorna False se não encontrar arquivo
   File = False
End Function

Function getconfiguracao(strCodigo As String) As String
Dim strAcha As Variant

strAcha = DLookup("ValorConfig", "tblConfig", "CodConfig = '" & strCodigo & "'")
If IsNull(strAcha) Then
    MsgBox "Código de configuração não encontrado! Verifique a tabela de configuracao.", 16, "Configuraçao"
    strAcha = ""
End If
getconfiguracao = strAcha
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: GotoFirstRecord
' Finalidade: Vai para o primeiro registro da tabela, consulta ou
'             formulário que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoFirstRecord() As Integer
   'Vai para o primeiro registro
   GotoFirstRecord = GotoRecord(A_FIRST)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: GotoLastRecord
' Finalidade: Vai para o último registro da tabela, consulta ou
'             formulário que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoLastRecord() As Integer
   'Vai para o último registro
   GotoLastRecord = GotoRecord(A_LAST)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: GotoNextRecord
' Finalidade: Vai para o próximo registro da tabela, consulta ou
'             formulário que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoNextRecord() As Integer
   'Vai para o próximo registro
   GotoNextRecord = GotoRecord(A_NEXT)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: GotoPrevRecord
' Finalidade: Vai para o registro anterior da tabela, consulta ou
'             formulário que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoPrevRecord() As Integer
   'Vai para o registro anterior
   GotoPrevRecord = GotoRecord(A_PREVIOUS)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: GotoRecord
' Finalidade: Vai para o registro indicado por intDireção da
'             tabela, consulta ou formulário que estiver ativo
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function GotoRecord(intDireção As Integer) As Integer
   'Desvia para próxima linha em caso de erro
   On Error Resume Next
   'Vai para o registro indicado por intDireção
   DoCmd.GotoRecord , , intDireção
   'Redesenha objeto
   DoCmd.RepaintObject
   'Retorna True
   GotoRecord = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: InvStr
' Finalidade: Inverte os caracteres de uma string
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function InvStr(strTexto As String) As String
   'Define variáveis
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
' Função....: IsADT
' Finalidade: Verifica se o MDB está sob o ADT ou Access
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsADT() As Integer
   'Retorna True se o MDB estiver sob o ADT
   IsADT = SysCmd(SYSCMD_RUNTIME)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: IsDig
' Finalidade: Verifica se strNúmero é composto somente por
'             dígitos, ou seja, pelos algarismos de 0 a 9
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsDig(strNúmero As String) As Integer
   Dim strNum As String
   Dim intC As Integer
   'Tira espaços
   strNum = Trim(strNúmero)
   'Testa tamanho
   If Len(strNum) = 0 Then
      IsDig = False
      Exit Function
   End If
   'Verifica todos os dígitos
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
' Função....: IsLoaded
' Finalidade: Verifica se formulário está carregado
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsLoaded(strForm As String) As Integer
   Dim intI As Integer
   'Localiza formulário na coleção Forms
   For intI = 0 To Forms.Count - 1
      'Se encontrar retorna True
      If Forms(intI).FormName = strForm Then
         IsLoaded = True
         Exit Function
      End If
   Next intI
   'Retorna False se não encontrar
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

   'Retorna False se não encontrar
   IsQuery = False


End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: IsTable
' Finalidade: Verifica se tabela existe no Banco de Dados atual
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsTable(strTable As String) As Integer
   Dim intI As Integer
   'Localiza tabela na coleção TableDefs
   For intI = 0 To meubd.TableDefs.Count - 1
      'Se encontrar retorna True
      If meubd.TableDefs(intI).Name = strTable Then
         IsTable = True
         Exit Function
      End If
   Next intI
   'Retorna False se não encontrar
   IsTable = False
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: IsWindowed
' Finalidade: Retorna True se o form/report não estiver
'             maximizado nem minimizado
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function IsWindowed(ByVal hWnd As Integer) As Integer
   'Verifica se a janela está minimizada ou maximizada
   If IsZoomed(hWnd) Or IsIconic(hWnd) Then
      IsWindowed = False
   Else
      IsWindowed = True
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: LDM
' Finalidade: Retorna última data do mês
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function LDM(varDtAtual As Variant) As Variant
   Dim varDtAux As Variant
   'Acrescenta 1 mês à data original
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
' Função....: MsgErro
' Finalidade: Exibe mensagem de erro
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function MsgErro(strMsg As String, intCancela) As Integer
   'Definição de variáveis
   Dim intRet As Integer
   'Verifica se mensagem é vazia
   If Trim(strMsg) = "" Then
      'Mensagem de erro
      intRet = MsgErro("Mensagem não pode ser vazia", False)
      'Retorna False
      MsgErro = False
   Else
      'Aviso sonoro
      Beep
      'Exibe mensagem de erro com ícone de aviso
      MsgErro = MsgBox(strMsg, 48 + Abs(intCancela))
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: NewMDB
' Finalidade: Cria um novo MDB
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function NewMDB(strName As String) As Integer
   'Desvia para próxima linha em caso de erro
   On Error Resume Next
   Dim db As Database
   'Abre Banco de Dados
   Set db = DBEngine.Workspaces(0).CreateDatabase(strName, DB_LANG_GENERAL)
   'Se não houver erro retorna True
   If Err = 0 Then NewMDB = True Else NewMDB = False
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: NomeArq
' Finalidade: Retorna o nome completo do arquivo strArq no
'             diretório strDir
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function NomeArq(strDir As String, strArq As String) As String
   'Retorna nome do arquivo com diretório
   NomeArq = Trim(strDir) & IIf(Right$(Trim(strDir), 1) = "\", "", "\") & Trim(strArq)
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: OcultaJanela
' Finalidade: Oculta janela ativa
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function OcultaJanela() As Integer
   'Oculta janela ativa
   DoCmd.DoMenuItem 1, 4, 3
   'Retorna True
   OcultaJanela = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: OpenBD
' Finalidade: Usada ao abrir MDB para padronizar apresentação e
'             verificar presença do ADT
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function OpenBD(strCaption As String) As Integer
   Dim varRet As Variant
   'Verifica se está sob ADT
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
   'Título da janela principal
   varRet = SetCaption(strCaption)
   'Retira mensagem da barra de status
   varRet = SetStatusMsg("")
   'Retorna True
   OpenBD = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: SetCaption
' Finalidade: ALtera caption (título) da janela do Access
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function SetCaption(strCaption As String) As Integer
   'Define variáveis
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
' Função....: SetStatusMsg
' Finalidade: Altera mensagem na barra de status do Access
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function SetStatusMsg(strMsg As String) As Integer
   'Define variáveis
   Dim varRet As Variant
   'Altera mensagem para strMsg
   If strMsg = "" Then strMsg = " "
   varRet = SysCmd(SYSCMD_SETSTATUS, strMsg)
   'Retorna True
   SetStatusMsg = True
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: ShowToolBar
' Finalidade: Exibe/oculta as barras de ferramentas do sistema.
'             Só funciona no MS-Access 2.0 em Português se o
'             mesmo estiver habilitado para exibi-las. Veja
'             função EnableToolBar
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
      DoCmd.ShowToolBar "Estrutura do Formulário", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Modo Formulário", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Filtro/Classificação", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Estrutura do Relatório", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Visualizar Impressão", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Caixa de Ferramentas", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Paleta", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Macro", A_TOOLBAR_WHERE_APPROP
      DoCmd.ShowToolBar "Módulo", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Microsoft", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Utilitário 1", A_TOOLBAR_WHERE_APPROP
      'DoCmd ShowToolbar "Utilitário 2", A_TOOLBAR_WHERE_APPROP
   Else
      DoCmd.ShowToolBar "Banco de Dados", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Relacionamentos", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Estrutura da Tabela", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Folha de Dados da Tabela", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Estrutura da Consulta", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Folha de Dados da Consulta", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Estrutura do Formulário", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Modo Formulário", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Filtro/Classificação", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Estrutura do Relatório", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Visualizar Impressão", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Caixa de Ferramentas", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Paleta", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Macro", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Módulo", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Microsoft", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Utilitário 1", A_TOOLBAR_NO
      DoCmd.ShowToolBar "Utilitário 2", A_TOOLBAR_NO
   End If
   'Retorna o parâmetro passado
   ShowToolBar = intAtiva
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: VerCGC
' Finalidade: Verifica se a string passada corresponde a um CGC
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function VerCGC(strNúmero As String) As Integer
   'Define variáveis
   Dim strNum As String
   Dim strDV1 As String
   Dim strDV2 As String
   'Separa número dos DVs
   strNum = Left$(Trim(strNúmero), Len(Trim(strNúmero)) - 2)
   'Calcula primeiro dígito
   strDV1 = CalcDV(strNum)
   If strDV1 = "X" Then strDV1 = "0"
   'Calcula segundo dígito
   strDV2 = CalcDV(strNum & strDV1)
   If strDV2 = "X" Then strDV2 = "0"
   'Compara com os DVs e retorna True
   If Right$(Trim(strNúmero), 2) = (strDV1 & strDV2) Then
      VerCGC = True
   Else
      VerCGC = False
   End If
End Function

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
' Função....: VerCPF
' Finalidade: Verifica se a string passada corresponde a um CPF
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function VerCPF(strNúmero As String) As Integer
   'Define variáveis
   Dim strNum As String
   Dim intC As Integer
   Dim strDV As String
   Dim intAcum As Integer
   Dim intResto As Integer
   Dim strDV1 As String
   Dim strDV2 As String
   'Tira espaços em branco do número
   strNum = Trim(strNúmero)
   'Verifica se só possui números
   If Not IsDig(strNum) Then
      VerCPF = False
      Exit Function
   End If
   'Verifica tamanho da string
   If Len(strNum) < 3 Or Len(strNum) > 11 Then
      VerCPF = False
      Exit Function
   End If
   'Inclui zeros para ficar com 11 dígitos
   Do While Len(strNum) <> 11
      strNum = "0" & strNum
   Loop
   'Separa o número dos dígitos
   strDV = Right$(strNum, 2)
   strNum = Left$(strNum, 9)
   'Multiplica cada número e acumula
   intAcum = 0
   For intC = 1 To 9
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * (11 - intC))
   Next
   'Acha o resto da divisão por 11 que é o DV1
   intResto = intAcum Mod 11
   'Primeiro DV
   strDV1 = Trim(IIf(intResto = 10, "0", Str$(intResto)))
   'Soma DV1 a strNum para repetir o cálculo
   strNum = strNum & strDV1
   'Multiplica cada número e acumula
   intAcum = 0
   For intC = 1 To 10
      intAcum = intAcum + (Val(Mid$(strNum, intC, 1)) * (12 - intC))
   Next
   'Acha o resto da divisão por 11 que é o DV2
   intResto = intAcum Mod 11
   'Retorna o dígito verificador
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
' Função....: VerDV
' Finalidade: Verifica o dígito verificador (módulo 11) de
'             strNúmero
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Function VerDV(strNúmero As String) As Integer
   'Define variáveis
   Dim strNum As String
   Dim strDV As String
   'Tira espaços a direita e a esquerda
   strNum = Trim(strNúmero)
   'Verifica se número passado é válido
   If Len(strNum) < 2 Then
      VerDV = False
      Exit Function
   End If
   'Calcula dígito e retorna
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
    .Sessão = "A"
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

    For Each ctl In nForm.Controls 'Para cada controle no formulário
        'Se o tipo do controle for Caixa de Texto então
        If TypeOf ctl Is TextBox Then
            ctl = Empty
        End If
        
        'Se o tipo do controle for Caixa de Combinação então
        If TypeOf ctl Is ComboBox Then
            ctl = Empty
        End If
        
        'Se o tipo do controle for Caixa de Listagem então
        If TypeOf ctl Is ListBox Then
            ctl = Empty
        End If
        
        'Se o tipo do controle for Caixa de Seleção então
        If TypeOf ctl Is CheckBox Then
            ctl = Empty
        End If
        
    Next

End Function


