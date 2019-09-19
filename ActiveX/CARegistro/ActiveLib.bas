Attribute VB_Name = "ActiveLib"
Option Explicit
Global xAmbiente  As Object 'XLib.Ambiente
Global xBanco     As Object 'XLib.Banco
Global xGeneral   As Object 'XLib.General
Global xMensagem  As Object 'XLib.Mensagem
Global xObjiGrid  As Object 'XLib.ObjiGrid
Global xObjRC     As Object 'XLib.ObjReportControl
Global xObjCmbBox As Object 'XLib.ObjComboBox
Public Function AbrirTxt(Arq As String) As Integer
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   AbrirTxt = xAmbiente.AbrirTxt(Arq)
End Function
Public Function AjustaTextoComboCodeJock(ByRef pCmb As Object, ByVal pFrm As Form) As Boolean
   If xObjCmbBox Is Nothing Then Set xObjCmbBox = CreateObject("xLIB.ObjComboBox")
   AjustaTextoComboCodeJock = xObjCmbBox.AjustaTextoComboCodeJock(pCmb, pFrm)
End Function
Public Function BuscaPeriodo(ByVal pSemana As String, ByRef pDataIni As Date, ByRef pDataFim As Date, Optional pExibeMensagem As Boolean = True) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   BuscaPeriodo = xGeneral.BuscaPeriodo(pSemana, pDataIni, pDataFim, pExibeMensagem)
End Function
Public Function BuscaSemana(pData As Date) As String
      If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   BuscaSemana = xGeneral.BuscaSemana(pData)
End Function
Public Function ClonarRS(ByVal pRecordSet As Object, Optional pFiltro As String) As Object
   If xBanco Is Nothing Then Set xBanco = CreateObject("xLIB.Banco")
   Set ClonarRS = xBanco.ClonarRS(pRecordSet, pFiltro)
End Function
Public Function ComputerName() As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ComputerName = xAmbiente.ComputerName()
End Function
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
Public Function ExtractResData(Id, Tipo, Arquivo As String, Optional pFileBuf) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   ExtractResData = xGeneral.ExtractResData(Id, Tipo, Arquivo, pFileBuf)
End Function
Public Sub FecharTxt(Arq As Integer)
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Call xAmbiente.FecharTxt(Arq)
End Sub
Public Function FillRCFromRS(ByRef pRecordSet As Object, ByRef pReportControl As Object, Optional bDoEvents As Boolean = False)
   If xObjRC Is Nothing Then Set xObjRC = CreateObject("xLIB.ObjReportControl")
   Call xObjRC.FillRCFromRS(pRecordSet, pReportControl, bDoEvents)
End Function
Function GetFileExtension(ByVal pFilename As String) As String
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
Public Function isAlfaNum(Character As String) As Boolean
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   isAlfaNum = xGeneral.isAlfaNum(Character)
End Function
Public Function LocalizarCombo(Cmb, Chave As String, Optional SetCombo = True, Optional PorItemData As Boolean = False) As Integer
   If xObjCmbBox Is Nothing Then Set xObjCmbBox = CreateObject("xLIB.ObjComboBox")
   LocalizarCombo = xObjCmbBox.LocalizarCombo(Cmb, Chave, SetCombo, PorItemData)
End Function
Public Sub OrdenarGrd(pReportControl As Object, pColChave As String, pColPai As String, pColTree As String)
   If xObjRC Is Nothing Then Set xObjRC = CreateObject("xLIB.ObjReportControl")
   Call xObjRC.OrdenarGrd(pReportControl, pColChave, pColPai, pColTree)
End Sub
Public Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, Optional DefaultValue As String) As String
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   ReadIniFile = xAmbiente.ReadIniFile(strIniFile, strSection, strKey, DefaultValue)
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
Public Function RichWordOver(ByVal RchTxt As Variant, x As Single, y As Single, Optional Posicao = 1, Optional VerifDclImplicta = True) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   RichWordOver = xGeneral.RichWordOver(RchTxt, x, y, Posicao, VerifDclImplicta)
End Function
Public Function SetTag(ByRef pControle As Variant, ByVal pNome As String, ByVal pValor As String) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   SetTag = xGeneral.SetTag(pControle, pNome, pValor)
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
Function ValBr(ByVal pNum As String, Optional pCasaDec As Integer = 2, Optional pArredTruncar As Integer = 1) As String
   If xGeneral Is Nothing Then Set xGeneral = CreateObject("xLIB.General")
   ValBr = xGeneral.ValBr(pNum, pCasaDec, pArredTruncar)
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

