Attribute VB_Name = "ActiveLib"
Option Explicit
Global xAmbiente     As Object 'XLib.Ambiente
Global xBanco        As Object 'XLib.Banco
Global xGeneral      As Object 'XLib.General
Global xMensagem     As Object 'XLib.Mensagem
Global xObjCommand   As Object 'XLib.xObjCommand
Global xObjiGrid     As Object 'XLib.ObjiGrid
Global xObjPane      As Object 'XLib.ObjPane
Global xObjRC        As Object 'XLib.ObjReportControl
Global xObjCmbBox    As Object 'XLib.ObjComboBox
Global xObjTreeView  As Object 'XLib.ObjTreeView
Global xObjZip       As Object 'XLib.ObjZip
Public Function AbrirTxt(Arq As String) As Integer
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   AbrirTxt = xAmbiente.AbrirTxt(Arq)
End Function
Public Sub AcoplarForm(pForm As Form, nPane As Integer, pSys As Object, Optional bDefineFoco As Boolean = True, Optional pMDI As Object)
   If xObjPane Is Nothing Then Set xObjPane = CreateObject("xLIB.ObjPane")
   Call xObjPane.AcoplarForm(pForm, nPane, pSys, bDefineFoco, pMDI)
End Sub
Public Function ExisteNo(pTree As Object, pKey As String) As Boolean
   If xObjTreeView Is Nothing Then Set xObjTreeView = CreateObject("xLIB.ObjTreeView")
   ExisteNo = xObjTreeView.ExisteNo(pTree, pKey)
End Function
Public Function AjustaTextoComboCodeJock(ByRef pCmb As Object, ByVal pFrm As Form) As Boolean
   If xObjCmbBox Is Nothing Then Set xObjCmbBox = CreateObject("xLIB.ObjComboBox")
   AjustaTextoComboCodeJock = xObjCmbBox.AjustaTextoComboCodeJock(pCmb, pFrm)
End Function
Public Function Between(Vl, Min, Max) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   Between = xGeneral.Between(Vl, Min, Max)
End Function
Public Function BinToBoolean(nVal As Long) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   BinToBoolean = xGeneral.BinToBoolean(nVal)
End Function
Public Function BooleanToBin(bVal As Boolean) As Long
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   BooleanToBin = xGeneral.BooleanToBin(bVal)
End Function
Public Function BuscaPeriodo(ByVal pSemana As String, ByRef pDataIni As Date, ByRef pDataFim As Date, Optional pExibeMensagem As Boolean = True) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   BuscaPeriodo = xGeneral.BuscaPeriodo(pSemana, pDataIni, pDataFim, pExibeMensagem)
End Function
Public Function BuscaSemana(pData As Date) As String
      If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   BuscaSemana = xGeneral.BuscaSemana(pData)
End Function
Public Sub CapturarTela(ByRef PicBox As PictureBox)
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Call xAmbiente.CapturarTela(PicBox)
End Sub
Public Function CapturarTelaSis(ByVal pDestinoJPG As String) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   CapturarTelaSis = xAmbiente.CapturarTelaSis(pDestinoJPG)
End Function
Public Function ClonarRS(ByVal pRecordSet As Object, Optional pFiltro As String) As Object
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   Set ClonarRS = xBanco.ClonarRS(pRecordSet, pFiltro)
End Function
Public Function ComputerName() As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ComputerName = xAmbiente.ComputerName()
End Function
Public Sub ConvertBMPtoJPG(pOriBMP As String, pDestJPG As String)
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Call xAmbiente.ConvertBMPtoJPG(pOriBMP, pDestJPG)
End Sub
Public Function CopiarArquivo(Orig As String, Dest As String) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   CopiarArquivo = xAmbiente.CopiarArquivo(Orig, Dest)
End Function
Public Function CriarDiretorio(pPath As String, Optional bViewMsg As Boolean = False) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   CriarDiretorio = xAmbiente.CriarDiretorio(pPath, bViewMsg)
End Function
Public Function CriarObjeto(sObjeto As String) As Object
   Dim MyObj As Object
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   
   Set MyObj = xAmbiente.CriarObjeto(sObjeto)
   If MyObj Is Nothing Then
      On Error Resume Next
      Set MyObj = CreateObject(sObjeto)
   End If
   Set CriarObjeto = MyObj
End Function
Public Function CriarRS(pColl As Collection) As Object
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   Set CriarRS = xBanco.CriarRS(pColl)
End Function
Public Function CruzRef(pRecordSet As Object, pClsDetahles As Object, pCampo As String _
         , pQtdCampo As Integer, Optional pTotalLinha As Boolean = False _
         , Optional pTotalColuna As Boolean = False, Optional pQtdValor As Integer = 1 _
         ) As Object
   
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   Set CruzRef = xBanco.CruzRef(pRecordSet, pClsDetahles, pCampo, pQtdCampo, pTotalLinha, pTotalColuna, pQtdValor)
End Function
Public Function Decrypt2(ByVal Password As String, Optional Key As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Decrypt2 = xAmbiente.Decrypt2(Password, Key)
End Function
Public Sub DockBarRightOf(pBarToDock As Variant, pBarOnLeft As Variant, Optional pSys As Object)
   If xObjCommand Is Nothing Then Set xObjCommand = CreateObject("xLIB.ObjCommand")
   Call xObjCommand.DockBarRightOf(pBarToDock, pBarOnLeft, pSys)
End Sub
Public Function Encrypt2(ByVal Password As String, Optional Key As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Encrypt2 = xAmbiente.Encrypt2(Password, Key)
End Function
Public Function ExcluirDiretorio(Diretorio As String) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ExcluirDiretorio = xAmbiente.ExcluirDiretorio(Diretorio)
End Function
Public Function ExcluirArquivo(File As String, Optional ViewError As Boolean = True) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ExcluirArquivo = xAmbiente.ExcluirArquivo(File, ViewError)
End Function
Public Sub ExibirInformacao(pTexto As String, Optional pTITULO As String)
   If xMensagem Is Nothing Then Set xMensagem = CreateObject("XLib.Mensagem")
   Call xMensagem.ExibirInformacao(pTexto, pTITULO)
End Sub
Public Sub ExibirAviso(pTexto As String, Optional pTITULO As String)
   If xMensagem Is Nothing Then Set xMensagem = CreateObject("XLib.Mensagem")
   Call xMensagem.ExibirAviso(pTexto, pTITULO)
End Sub
Public Function ExibirPergunta(pTexto As String, Optional pTITULO As String, Optional pDefaultYes = True) As Integer
   If xMensagem Is Nothing Then Set xMensagem = CreateObject("XLib.Mensagem")
   ExibirPergunta = xMensagem.ExibirPergunta(pTexto, pTITULO, pDefaultYes)
End Function
Public Sub ExibirResultado(pSys As Object, Optional pResultado As Boolean = True, Optional pNumPisca As Integer, Optional pMsg As String = "")
   If xMensagem Is Nothing Then Set xMensagem = CreateObject("XLib.Mensagem")
   Call xMensagem.ExibirResultado(pSys, pResultado, pNumPisca, pMsg)
End Sub
Public Sub ExibirStop(pTexto As String, Optional pTITULO As String)
   If xMensagem Is Nothing Then Set xMensagem = CreateObject("XLib.Mensagem")
   Call xMensagem.ExibirStop(pTexto, pTITULO)
End Sub
Function ExisteArquivo(ByVal strPathName As String) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ExisteArquivo = xAmbiente.ExisteArquivo(strPathName)
End Function
Public Function ExisteItem(pColl As Collection, pItem As String) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   ExisteItem = xGeneral.ExisteItem(pColl, pItem)
End Function
Public Sub ExecuteLink(ByVal sLinkTo As String)
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   Call xGeneral.ExecuteLink(sLinkTo)
End Sub
Public Function ExtractResData(Id, Tipo, Arquivo As String, Optional pFileBuf) As Boolean
'   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
'   ExtractResData = xGeneral.ExtractResData(Id, Tipo, Arquivo, pFileBuf)
   Dim nInt As Integer
   Dim byteFileBuf() As Byte 'This must be byte rather than String, so no Unicode conversion takes place
   Dim nVez As Integer
   Dim sPath   As String
   
   On Error GoTo Fim
   
   Call GetNameFromPath(Arquivo, sPath)
   If sPath <> "" Then
      Call CriarDiretorio(sPath)
   End If
   Call ExcluirArquivo(Arquivo, False)
   
   nInt = FreeFile
   Open Arquivo$ For Binary Access Write As nInt
      If IsMissing(pFileBuf) Then
         byteFileBuf = LoadResData(Id, Tipo)
      End If
      Put nInt, , byteFileBuf
   GoTo Saida
Fim:
   nVez = nVez + 1
   If nVez < 5 Then
      Resume
   Else
      Resume Next
   End If
   
Saida:
    Close nInt
    Err = 0
    ExtractResData = ExisteArquivo(Arquivo$)
    Exit Function
End Function
Public Sub FecharTxt(Arq As Integer)
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Call xAmbiente.FecharTxt(Arq)
End Sub
Public Function FillRCFromRS(ByRef pRecordSet As Object, ByRef pReportControl As Object, Optional bDoEvents As Boolean = False, Optional ByRef pCollColumn As Collection, Optional pCurrency As Boolean = False)
   If xObjRC Is Nothing Then Set xObjRC = CreateObject("xLIB.ObjReportControl")
   Call xObjRC.FillRCFromRS(pRecordSet, pReportControl, bDoEvents, pCollColumn, pCurrency)
End Function
Public Function FormatarData(pStrDate As String, Optional pNull As Boolean = False) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   FormatarData = xGeneral.FormatarData(pStrDate, pNull)
End Function
Public Function FormatarHora(pStrHour As String, Optional pSegundo As Boolean = False) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   FormatarHora = xGeneral.FormatarHora(pStrHour, pSegundo)
End Function
Public Function FormatarNome(pNome As String, Optional Somente1Maiuscula As Boolean) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   FormatarNome = xGeneral.FormatarNome(pNome, Somente1Maiuscula)
End Function
Public Function GetFileExtension(ByVal pFilename As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetFileExtension = xAmbiente.GetFileExtension(pFilename)
End Function
Public Function GetFileVersion(ByVal pFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetFileVersion = xAmbiente.GetFileVersion(pFilename, fIsRemoteServerSupportFile)
End Function
Public Function GetFileVersionNumber(pFilename As String) As Double
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetFileVersionNumber = xAmbiente.GetFileVersionNumber(pFilename)
End Function
Public Function GetGrdColumnIndex(pGrd As Object, pCaption As String) As Integer
   If xObjRC Is Nothing Then Set xObjRC = CreateObject("xLIB.ObjReportControl")
   GetGrdColumnIndex = xObjRC.GetGrdColumnIndex(pGrd, pCaption)
End Function
Public Function GetNameFromPath(PathFile As String, Optional ByRef PathReturn As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetNameFromPath = xAmbiente.GetNameFromPath(PathFile, PathReturn)
End Function
Public Function GetShortName(sFile As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetShortName = xAmbiente.GetShortName(sFile)
End Function
Public Function GetSpecialFolder(CSIDL As Long) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetSpecialFolder = xAmbiente.GetSpecialFolder(CSIDL)
End Function
Public Function GetSerialNumber(Optional sDrive As String = "C:\") As Long
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetSerialNumber = xAmbiente.GetSerialNumber(sDrive)
End Function
Public Function GetUserName() As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetUserName = xAmbiente.GetUserName()
End Function
Public Function GetTag(ByRef pControle As Variant, ByVal pNome As String, Optional pPadrao As String) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   GetTag = xGeneral.GetTag(pControle, pNome, pPadrao)
End Function
Public Function GetTempFolder() As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GetTempFolder = xAmbiente.GetTempFolder()
End Function
Public Function GetTypeField(pFieldName As String, pRecordSet As Object) As VbVarType
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   GetTypeField = xBanco.GetTypeField(pFieldName, pRecordSet)
End Function
Public Function GetWords(ByVal StrLinha As String) As Collection
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   Set GetWords = xGeneral.GetWords(StrLinha)
End Function
Public Function GetWords_AndOR(pTexto As String, Optional ByRef Palavras_And As Collection, Optional ByRef Palavras_Or As Collection, Optional pCampo) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   GetWords_AndOR = xGeneral.GetWords_AndOR(pTexto, Palavras_And, Palavras_Or, pCampo)
End Function
Public Function GravarArquivoLog(pPath As String, pNomeArq As String, pTITULO As String, pConteudo As String, Optional bHora As Boolean = True)
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   GravarArquivoLog = xAmbiente.GravarArquivoLog(pPath, pNomeArq, pTITULO, pConteudo, bHora)
End Function
Public Function iGridToRecordset(ByVal pIGrid As Object, Optional pSomenteSelecao, Optional pRsDados As Object) As Object
   If xObjiGrid Is Nothing Then Set xObjiGrid = CreateObject("xLIB.ObjiGrid")
   Set iGridToRecordset = xObjiGrid.iGridToRecordset(pIGrid, pSomenteSelecao, pRsDados)
End Function
Public Function InArray(Valor As Variant, VETOR As Variant) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   InArray = xGeneral.InArray(Valor, VETOR)
End Function
Public Function InputBoxPassword(prompt, Optional Title, Optional Default) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   InputBoxPassword = xGeneral.InputBoxPassword(prompt, Title, Default)
End Function
Public Function isAlfaNum(Character As String) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   isAlfaNum = xGeneral.isAlfaNum(Character)
End Function
Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   IsWebConnected = xAmbiente.IsWebConnected(ConnType)
End Function
Public Function LocalizarCombo(Cmb, Chave As String, Optional SetCombo = True, Optional PorItemData As Boolean = False) As Integer
   If xObjCmbBox Is Nothing Then Set xObjCmbBox = CreateObject("xLIB.ObjComboBox")
   LocalizarCombo = xObjCmbBox.LocalizarCombo(Cmb, Chave, SetCombo, PorItemData)
End Function
Public Sub OrdenarGrd(pReportControl As Object, pColChave As String, pColPai As String, pColTree As String)
   If xObjRC Is Nothing Then Set xObjRC = CreateObject("xLIB.ObjReportControl")
   Call xObjRC.OrdenarGrd(pReportControl, pColChave, pColPai, pColTree)
End Sub
Public Function ProcuraArquivo(ByVal pPath As String, ByVal pArq As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ProcuraArquivo = xAmbiente.ProcuraArquivo(pPath, pArq)
End Function
Public Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, Optional DefaultValue As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ReadIniFile = xAmbiente.ReadIniFile(strIniFile, strSection, strKey, DefaultValue)
End Function
Public Function ReadTextFile(strPath As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ReadTextFile = xAmbiente.ReadTextFile(strPath)
End Function
Public Function RecordSetToExcel(ByRef pRs As Object, Optional ByVal pNome, Optional ByVal isVisible As Boolean = False, Optional ByRef pForm, Optional ByVal TopFlood, Optional ByVal ExcluiArq As Boolean = True, Optional ByVal NomeArq, Optional ByVal ExibeMsg As Boolean = True) As Boolean
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   RecordSetToExcel = xBanco.RecordSetToExcel(pRs, pNome, isVisible, pForm, TopFlood, ExcluiArq, NomeArq, ExibeMsg)
End Function
Public Function RegServer(sServerPath As String, Optional fRegister = True, Optional fMsg As Boolean = True, Optional isActivexExe As Boolean = False) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   RegServer = xAmbiente.RegServer(sServerPath, fRegister, fMsg, isActivexExe)
End Function
Public Function ResolvePathName(ByVal sPath As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ResolvePathName = xAmbiente.ResolvePathName(sPath)
End Function
Public Sub RetiraPreposicao(ByRef pString As String, Optional ByRef pClString As Collection)
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   Call xGeneral.RetiraPreposicao(pString, pClString)
End Sub
Public Function RichWordOver(ByVal RchTxt As Variant, x As Single, y As Single, Optional Posicao = 1, Optional VerifDclImplicta = True) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   RichWordOver = xGeneral.RichWordOver(RchTxt, x, y, Posicao, VerifDclImplicta)
End Function
Public Function SetMDI(ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
  If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
  SetMDI = xAmbiente.SetMDI(hWndChild, hWndNewParent)
End Function
Public Function SetTag(ByRef pControle As Variant, ByVal pNome As String, ByVal pValor As String) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   SetTag = xGeneral.SetTag(pControle, pNome, pValor)
End Function
Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   SetTopMostWindow = xAmbiente.SetTopMostWindow(hwnd, Topmost)
End Function
Public Sub SelecionarTexto(ByRef Obj As Object)
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   Call xGeneral.SelecionarTexto(Obj)
End Sub
Public Function SendSMS(ByVal pUrl As String) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   SendSMS = xGeneral.SendSMS(pUrl)
End Function

Public Function SendTab(frm As Object, ByVal Key As Integer, Optional Tipo As Variant, _
                        Optional Obj As Variant, Optional Maiuscula = True, _
                        Optional Tamanho As Integer = 13, _
                        Optional Qtd_Dec As Integer = 2) As Integer
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   SendTab = xGeneral.SendTab(frm, Key, Tipo, Obj, Maiuscula, Tamanho, Qtd_Dec)
End Function
Public Sub SincShell(Comando As String, Optional Modo As VbAppWinStyle = vbMaximizedFocus, Optional EsperaProcesso = True)
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Call xAmbiente.SincShell(Comando, Modo, EsperaProcesso)
End Sub
Public Function SqlDate(ByVal DT As String, Optional Format_Date As Integer = 3, Optional InsereNull As Boolean = True, Optional pDbTipo As Integer = 1) As String
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   SqlDate = xBanco.SqlDate(DT, Format_Date, InsereNull, pDbTipo)
End Function
Function SqlNum(ByVal Num As String, Optional InsereNull As Boolean = False) As String
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   SqlNum = xBanco.SqlNum(Num, InsereNull)
End Function
Public Function SqlStr(ByVal Txt As String, Optional InsereNull As Boolean = False, Optional pDbTipo As Integer = 1) As String
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   SqlStr = xBanco.SqlStr(Txt, InsereNull, pDbTipo)
End Function
Public Sub ShowError(Optional TxtAux = "", Optional pExibeMsg As Boolean = True)
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("XLib.General")
   Call xGeneral.ShowError(TxtAux, pExibeMsg)
End Sub
Public Function StrZero(pValor As Variant, pQtd As Integer, Optional pCaracter = "0") As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   StrZero = xGeneral.StrZero(pValor, pQtd, pCaracter)
End Function
Public Function TratarMoeda(Key%, ByRef Obj As Object, Optional Tamanho As Integer, Optional Qtd_Dec As Integer = 2) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   TratarMoeda = xGeneral.TratarMoeda(Key, Obj, Tamanho, Qtd_Dec)
End Function
'Public Function Traduzir(pString As String, Optional pIdioma As Double) As String
'   If xMensagem Is Nothing Then Set xMensagem = CreateObject("XLib.Mensagem")
'   Traduzir = xMensagem.Traduzir(pString, pIdioma)
'End Function
Function UnFormat(ByVal Codigo As String) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   UnFormat = xGeneral.UnFormat(Codigo)
End Function
Public Function UrlEncode(ByVal urlText As String) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   UrlEncode = xGeneral.UrlEncode(urlText)
End Function
Function ValBr(ByVal pNum As String, Optional pCasaDec As Integer = 2, Optional pArredTruncar As Integer = 1) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   ValBr = xGeneral.ValBr(pNum, pCasaDec, pArredTruncar)
End Function
Public Function ValidarCNPJ(ByVal NumCNPJ) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   ValidarCNPJ = xGeneral.ValidarCNPJ(NumCNPJ)
End Function
Public Function ValidarCPF(ByVal NumCPF As String) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   ValidarCPF = xGeneral.ValidarCPF(NumCPF)
End Function
Function ValorReal(pValor As String) As Currency
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   ValorReal = xGeneral.ValorReal(pValor)
End Function
Public Sub Wait(pSecond As Integer)
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   Call xGeneral.Wait(pSecond)
End Sub
Public Function WriteIniFile(ByVal strIniFile As String, strSection As String, strKey As String, strValue As String) As Boolean
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   WriteIniFile = xAmbiente.WriteIniFile(strIniFile, strSection, strKey, strValue)
End Function
Function xVal(ByVal pNum As String, Optional pQtdCasaDec = 5) As Double
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   xVal = xGeneral.xVal(pNum, pQtdCasaDec)
End Function
Function Zip(pCollOrFile As String, pZipFileName As String) As Long
   If xObjZip Is Nothing Then Set xObjZip = CreateObject("xLIB.ObjZip")
   Zip = xObjZip.Zip(pCollOrFile, pZipFileName)
End Function
Function Unzip(pPath As String, pFile As String, Optional pPathDest As String, Optional pHonorDir As Boolean = True) As Boolean
   If xObjZip Is Nothing Then Set xObjZip = CreateObject("xLIB.ObjZip")
   Unzip = xObjZip.Unzip(pPath, pFile, pPathDest, pHonorDir)
End Function


