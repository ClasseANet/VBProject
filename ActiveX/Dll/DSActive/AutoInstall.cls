VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoInstall"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const SubFolder = "ClasseA"
Const Chave = 4582
Public Function VerificaSetup(pCODSIS As String, AppPath As String, pExeName As String, pCommand As String, Optional PathSetup As String, Optional SetupFile As String = "SETUP.INI") As Boolean
   Dim i             As Integer
   Dim bDebug        As Boolean
   Dim sPath         As String
   Dim sPathSis      As String
   Dim sPathNewFile  As String
   Dim sPathTemp     As String
   Dim sPathTempExe  As String
   Dim sPathExe      As String
   Dim sVersaoInst   As String
   Dim sVersaoMaq    As String
   Dim nVersaoInst   As Double
   Dim nVersaoTemp   As Double
   Dim nVersaoMaq    As Double
   Dim sArq          As String
   Dim sArqExe       As String
   Dim sStatus       As String
   Dim LerIni        As Boolean
   Dim SetupIni      As String
   Dim LocalSetupIni As String
   Dim sCommand      As String
   Dim sMsg          As String
   
   On Error Resume Next

   '***********
   '* Possíveis valores de pCommand
   '***********
   'pCommand = "|SourcePath=C:\PROGRA~1\CLASSEA\PAC\| /DEBUG"
   'pCommand = " |ClearTemp=True|"

'pCommand = "|SOURCEPATH=C:\PROGRA~1\CLASSEA\PAC\|PATHSETUP=C:\PROGRA~1\CLASSEA\PAC\ /DEBUG"

   VerificaSetup = True
   pExeName = UCase(pExeName)
   pCODSIS = UCase(pCODSIS)
   pCommand = UCase(pCommand)
   bDebug = ((InStr(pCommand, "DEBUG") <> 0))
   
   If bDebug Then MsgBox "In VerificaSetup"
   
   sPathTemp = UCase(GetSpecialFolder(CSIDL_TEMPORARY))
   sPathTemp = UCase(sPathTemp & "ClasseA\")
      
   '***********
   '* AppPath = "C:\Program Files\ClasseA\Producao"
   '* PathSetup = "\\guarani\sistemas\admin\"
   AppPath = UCase(ClsDsr.ResolvePathName(AppPath))
   sPathExe = UCase(ClsCtrl.GetTag(pCommand, "SOURCEPATH", ""))
   SetupFile = UCase(IIf(SetupFile = "", "SETUP.INI", SetupFile))
   PathSetup = UCase(IIf(PathSetup = "", AppPath, PathSetup))
   If Trim(ClsCtrl.GetTag(pCommand, "PATHSETUP", "")) <> "" Then
      PathSetup = UCase(Trim(ClsCtrl.GetTag(pCommand, "PATHSETUP", "")))
   End If
   
   LocalSetupIni = AppPath & SetupFile
   SetupIni = PathSetup & SetupFile
   
   If ClsDos.FileExists(LocalSetupIni) Then
      PathSetup = AppPath
   Else
      If Not ClsDos.FileExists(SetupIni) Then
         PathSetup = UCase(ClsDsr.ResolvePathName(GetSetting(pCODSIS, "Setup", "PATHSETUP")))
      End If
   End If
   
   SetupIni = PathSetup & SetupFile
   
   If bDebug Then MsgBox "If ClsDos.FileExists(" & SetupIni & ") Then"
   If ClsDos.FileExists(SetupIni) Then
      If bDebug Then
         MsgBox "If ClsDos.FileExists(" & SetupIni & ") Then 'R: True"
      End If

      '************
      '* Define Conexão Padrão [Database Format]
      '************
      Call RegistroConn(pCODSIS, SetupIni)

      '* sStatus = 1 -> Atualiza somente versões novas
      '* sStatus = 2 -> Atualiza coam a versão existente no servidor
      sStatus = ReadIniFile(SetupIni, "General", "Status", "0")
      If sStatus = "0" Then
         LerIni = False
      Else
         sStatus = ReadIniFile(SetupIni, "AutoInstall", "Status", "0")
         LerIni = (sStatus = "1" Or sStatus = "2")
      End If
      If LerIni Then
         '************
         '* Verifica Executável
         '************
         sPathNewFile = ClsDsr.ResolvePathName(ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "Path", PathSetup))
         sArqExe = pExeName & IIf(Right(pExeName, 4) = ".exe", "", ".exe")
         If Not ClsDos.FileExists(sPathNewFile & sArqExe) Then
            sPathNewFile = ClsDsr.ResolvePathName(ReadIniFile(SetupIni, "AutoInstall Files", "Path", PathSetup))
         End If
         
         If bDebug Then
            MsgBox "ClsDos.FileExists(" & sPathNewFile & sArqExe & ") " & IIf(ClsDos.FileExists(SetupIni), "= True ", "= False")
         End If
         
         If ClsDos.FileExists(sPathNewFile & sArqExe) Then
            nVersaoInst = GetVersao(GetFileVersion(sPathNewFile & sArqExe))
            If sPathExe = "" Then
               nVersaoMaq = GetVersao(GetFileVersion(AppPath & sArqExe))
            Else
               nVersaoMaq = GetVersao(GetFileVersion(sPathExe & sArqExe))
            End If
            If bDebug Then
               MsgBox sPathNewFile & sArqExe & " Vs." & CStr(nVersaoInst) & vbNewLine & AppPath & "\" & sArqExe & " Vs." & CStr(nVersaoMaq)
            End If
            If (nVersaoInst > nVersaoMaq And nVersaoMaq <> 0 And nVersaoInst <> 0) Or (sStatus = "2" And nVersaoInst <> 0) Then
               Err = 0
               If sPathExe = "" Then
                  '***************
                  '* Cria Executável Temporário
                  '***************
                  sPathTempExe = sPathTemp
                  If CreatePathTempExe(sPathTempExe, sArqExe, nVersaoInst) Then
                     i = 0
                     nVersaoTemp = GetVersao(GetFileVersion(sPathTempExe & sArqExe))
                     While (nVersaoInst > nVersaoTemp) And (i < 10)
                        If bDebug Then MsgBox " Tentativa " & i & " - Call FileCopy(" & sPathNewFile & sArqExe & ", " & sPathTempExe & sArqExe & ")"
                        Call FileCopy(sPathNewFile & sArqExe, sPathTempExe & sArqExe)
                        nVersaoTemp = GetVersao(GetFileVersion(sPathTempExe & sArqExe))
                        i = i + 1
                     Wend
'                     If Err <> 0 Then '* Path not found
'                        MsgBox Err.Number & " - " & Err.Description, vbCritical, "AutoInstall [" & sPathNewFile & sArqExe & "] To [" & sPathTempExe & sArqExe & "]"
'                     End If
                     ClsDsr.Wait 3
                     sCommand = sPathTempExe & sArqExe & " |SOURCEPATH=" & AppPath & "|PATHSETUP=" & PathSetup & "|"
                     If bDebug Then
                        sCommand = sCommand & " /Debug"
                        MsgBox "SincShell(" & sCommand & ", False)"
                     End If
                     'Call SincShell(sCommand, False)
                     Call Shell(sCommand, vbMinimizedNoFocus)
                     VerificaSetup = False
                     Exit Function
                  End If
               Else
                  '***************
                  '* Troca de Executável
                  '***************
                  i = 1
                  sMsg = "O Sistema detectou uma nova versão e irá atualizar automaticamente sua máquina." & vbNewLine
                  sMsg = sMsg & vbNewLine
                  sMsg = sMsg & "Após o [Ok] esta operação levará alguns segundos." & vbNewLine
                  sMsg = sMsg & "Por favor aguarde..."
                  MsgBox sMsg
                  While (nVersaoInst > nVersaoMaq) And (i < 10)
                     If bDebug Then MsgBox i & " - Call FileCopy(" & sPathNewFile & sArqExe & ", " & sPathExe & sArqExe & ")"
                     Call ClsDos.Del(sPathExe & sArqExe, False)
                     Call FileCopy(sPathNewFile & sArqExe, sPathExe & sArqExe)
                     nVersaoMaq = GetVersao(GetFileVersion(sPathExe & sArqExe))
                     i = i + 1
                  Wend
                  ClsDsr.Wait 3
                  sCommand = sPathExe & sArqExe & " |CLEARTEMP=TRUE|PATHSETUP=" & PathSetup & "|"
                  If bDebug Then
                     sCommand = sCommand & " /Debug"
                     MsgBox "SincShell(" & sCommand & ", False)"
                  End If
                  'Call SincShell(sCommand, False)
                  Call Shell(sCommand, vbNormalFocus)
                  VerificaSetup = False
               End If
               Exit Function
            End If
            '***************
            '* Limpa Temporário
            '***************
            ClsDsr.Wait 3
            If ClsCtrl.GetTag(pCommand, "CLEARTEMP", "") = "TRUE" Then
               Call DeletePathTempExe(sPathTemp, sArqExe, nVersaoInst)
            End If
         End If
         
         i = 1
         '************
         '* Verifica Objetos Comuns
         '************
         While ReadIniFile(SetupIni, "AutoInstall Files", "File" & CStr(i)) <> ""
            sArq = ReadIniFile(SetupIni, "AutoInstall Files", "File" & CStr(i))
            '**********
            '* Recupera pasta do arquivo a ser registrado
            sPathSis = ""
            sPathSis = GetRegisterDir(sArq, AppPath)
            If sPathSis = "" Then
               sPathSis = ReadIniFile(SetupIni, "AutoInstall Files", "Path" & CStr(i))
            End If
            sPathSis = ClsDsr.ResolvePathName(sPathSis)
            If sPathSis <> "" Then
               If bDebug Then
                  MsgBox "i=" & CStr(i) & vbNewLine & sPathNewFile & sArq & " Vs." & GetFileVersion(sPathNewFile & sArq) & vbNewLine & GetWindowsSysDir() & sArq & " Vs." & GetFileVersion(GetWindowsSysDir() & sArq)
               End If
               If Not AtualizaDLL(sPathSis & sArq, sPathNewFile & sArq, (sStatus = "2")) Then
                  Call AtualizaDLL(sPathSis & sArq, sPathNewFile & sArq, (sStatus = "2"))
               End If
            End If
            i = i + 1
         Wend
         
         i = 1
         '************
         '* Verifica Objetos do Módulo
         '************
         While ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "File" & CStr(i)) <> ""
            sArq = ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "File" & CStr(i))
            '**********
            '* Recupera pasta do arquivo a ser registrado
            sPathSis = ""
            sPathSis = GetRegisterDir(sArq, AppPath)
            If sPathSis = "" Then
               sPathSis = ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "Path" & CStr(i))
            End If
            sPathSis = ClsDsr.ResolvePathName(sPathSis)
            If sPathSis <> "" Then
               If bDebug Then
                  MsgBox "i=" & CStr(i) & vbNewLine & sPathNewFile & sArq & " Vs." & GetFileVersion(sPathNewFile & sArq) & vbNewLine & GetWindowsSysDir() & sArq & " Vs." & GetFileVersion(GetWindowsSysDir() & sArq)
               End If
               If Not AtualizaDLL(sPathSis & sArq, sPathNewFile & sArq, (sStatus = "2")) Then
                  Call AtualizaDLL(sPathSis & sArq, sPathNewFile & sArq, (sStatus = "2"))
               End If
            End If
            i = i + 1
         Wend
      End If
   Else
      If bDebug Then
         MsgBox "If ClsDos.FileExists(" & SetupIni & ") Then 'R: False"
      End If
   End If
   VerificaSetup = True
   Exit Function
   
OpenError:
   Resume Next
End Function
Public Function GetSpecialFolder(CSIDL As Long) As String
    Dim sPath  As String
    Dim IDL    As ITEMIDLIST
    Dim nhWnd  As Long
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
   If CSIDL = CSIDL_TEMPORARY Then
      GetSpecialFolder = ClsDsr.ResolvePathName(GetTempFolder)
   Else
      GetSpecialFolder = ""
      If SHGetSpecialFolderLocation(nhWnd, CSIDL, IDL) = 0 Then
         '
         ' Get the path from the ID list, and return the folder.
         '
         sPath = Space$(MAX_PATH)
         If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            GetSpecialFolder = ClsDsr.ResolvePathName(Left$(sPath, InStr(sPath, vbNullChar) - 1) & "")
         End If
     End If
   End If
End Function
Public Function GetTempFolder() As String
   ' Returns the path to the user's Temp folder. To boot, Windows
   ' requires that a temporary folder exist, so this should always
   ' safely return a path to one. Just in case, though, check the
   ' return value of GetTempPath.
   
   Dim strTempPath As String
   Dim lngTempPath As Long
   
   ' Fill string with null characters.
   strTempPath = String(144, vbNullChar)
   ' Get length of string.
   lngTempPath = Len(strTempPath)
   ' Call GetTempPath, passing in string length and string.
   If (GetTempPath(lngTempPath, strTempPath) > 0) Then
      ' GetTempPath returns path into string.
      ' Truncate string at first null character.
      GetTempFolder = ClsDsr.ResolvePathName(VBA.Left(strTempPath, InStr(1, strTempPath, vbNullChar) - 1))
   Else
      GetTempFolder = ""
   End If
End Function

Private Sub RegistroConn(pCODSIS As String, pSetupIni As String)
   Dim sSecao     As String
   Dim SecaoBase  As String
   Dim Index      As Integer
   
   Dim sALIAS     As String
   Dim bisODBC    As Boolean
   Dim bisADO     As Boolean
   Dim nDBTIPO    As Integer
   Dim sDBVERSAO  As String
   Dim sSERVER    As String
   Dim sDBNAME    As String
   Dim sDSN       As String
   Dim sUID       As String
   Dim sPWD       As String
   Dim nLast      As Integer
   
   SecaoBase = "\Conections\Conection_"
   Index = 0
   sSecao = SecaoBase & Index
   
   If ClsDos.FileExists(pSetupIni) Then
      If ReadIniFile(pSetupIni, "General", "Status", "0") = "0" Then Exit Sub
      If ReadIniFile(pSetupIni, "Database", "Status", "0") = "0" Then Exit Sub
   End If
   
   sALIAS = "CLASSEA"
   nDBTIPO = 1
   sDBVERSAO = ""
   bisODBC = False
   bisADO = True
   sSERVER = ComputerName & "\DSVM"
   sDBNAME = "CLASSEA"
   sDSN = ""
   sUID = "USU_VERIF"
   sPWD = ClsSeg.Decrypt("iH,pf067^")
   
   If Trim(GetSetting(pCODSIS, sSecao, "ALIAS", "")) <> "" Then
      sALIAS = GetSetting(pCODSIS, sSecao, "ALIAS", sALIAS)
      nDBTIPO = GetSetting(pCODSIS, sSecao, "DBTIPO", CStr(nDBTIPO))
      sDBVERSAO = GetSetting(pCODSIS, sSecao, "DBVERSAO")
      bisADO = GetSetting(pCODSIS, sSecao, "isADO", CStr(bisADO))
      bisODBC = GetSetting(pCODSIS, sSecao, "isODBC", CStr(bisODBC))
      sSERVER = GetSetting(pCODSIS, sSecao, "SERVER", sSERVER)
      sDBNAME = GetSetting(pCODSIS, sSecao, "DBNAME", sDBNAME)
      sDSN = GetSetting(pCODSIS, sSecao, "DSN", sDSN)
      sUID = GetSetting(pCODSIS, sSecao, "UID", sUID)
      sPWD = ClsSeg.Decrypt(GetSetting(pCODSIS, sSecao, "PWD", IIf(sUID = "USU_VERIF", ClsSeg.Encrypt(sPWD), "")))
   End If
   If ClsDos.FileExists(pSetupIni) Then
      sALIAS = ReadIniFile(pSetupIni, "Database Format", "ALIAS", sALIAS)
      nDBTIPO = ReadIniFile(pSetupIni, "Database Format", "DBTIPO", CStr(nDBTIPO))
      sDBVERSAO = ReadIniFile(pSetupIni, "Database Format", "DBVERSAO")
      bisADO = ReadIniFile(pSetupIni, "Database Format", "isADO", CStr(bisADO))
      bisODBC = ReadIniFile(pSetupIni, "Database Format", "isODBC", CStr(bisODBC))
      sSERVER = ReadIniFile(pSetupIni, "Database Format", "SERVER", sSERVER)
      sDBNAME = ReadIniFile(pSetupIni, "Database Format", "DBNAME", sDBNAME)
      sDSN = ReadIniFile(pSetupIni, "Database Format", "DSN", sDSN)
      sUID = ReadIniFile(pSetupIni, "Database Format", "UID", sUID)
      sPWD = ClsSeg.Decrypt(ReadIniFile(pSetupIni, "Database Format", "PWD", IIf(sUID = "USU_VERIF", ClsSeg.Encrypt(sPWD), "")))
   End If
   
   While GetSetting(pCODSIS, SecaoBase & Index, "ALIAS") <> ""
      Index = Index + 1
   Wend
   nLast = GetSetting(pCODSIS, "\Conections", "Last", 0)
   nLast = IIf(nLast > Index, 0, nLast)
   Call SaveSetting(pCODSIS, "\Conections", "Last", CStr(nLast))
         
   If Trim(sALIAS) <> "" Then Call SaveSetting(pCODSIS, sSecao, "ALIAS", sALIAS)
   Call SaveSetting(pCODSIS, sSecao, "DBTIPO", nDBTIPO)
   'Call SaveSetting(pCODSIS, sSecao, "DBVERSAO", sDBVERSAO)
   Call SaveSetting(pCODSIS, sSecao, "isODBC", bisODBC)
   Call SaveSetting(pCODSIS, sSecao, "isADO", bisADO)
   If Trim(sSERVER) <> "" Then Call SaveSetting(pCODSIS, sSecao, "SERVER", sSERVER)
   Call SaveSetting(pCODSIS, sSecao, "DBNAME", sDBNAME)
   'Call SaveSetting(pCODSIS, sSecao, "DSN", sDSN)
   Call SaveSetting(pCODSIS, sSecao, "UID", sUID)
   Call SaveSetting(pCODSIS, sSecao, "PWD", ClsSeg.Encrypt(sPWD))
End Sub
Public Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, Optional DefaultValue As String) As String
    Dim strBuffer As String
    
    strBuffer = Space$(gintMAX_SIZE)
    If GetPrivateProfileString(strSection, strKey, gstrNULL, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = DefaultValue
    End If
End Function
Public Function WriteIniFile(ByVal strIniFile As String, strSection As String, strKey As String, strValue As String) As Boolean
   Dim intLen As Integer
   
   If Not ClsDos.FileExists(strIniFile) Then
      intLen = ClsDos.AbrirTxt(strIniFile)
      Call ClsDos.FecharTxt(intLen)
   End If
   intLen = 0
   intLen = WritePrivateProfileString(strSection, strKey, strValue, strIniFile)
   WriteIniFile = (intLen > 0)
End Function
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = VBA.Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Private Function CreatePathTempExe(ByRef pPathTemp As String, ByVal pArq As String, ByVal pVersaoInst As Double) As Boolean
   On Error GoTo TrataErro
   
   If Dir(pPathTemp, vbDirectory) = "" Then
      Call MkDir(pPathTemp)
   End If
   
   pPathTemp = pPathTemp & Mid(pArq, 1, IIf(InStr(pArq, ".") = 0, Len(pArq), InStr(pArq, ".") - 1))
   pPathTemp = ClsDsr.ResolvePathName(pPathTemp)
   If Dir(pPathTemp, vbDirectory) = "" Then
      Call MkDir(pPathTemp)
   End If
   
   pPathTemp = ClsDsr.ResolvePathName(pPathTemp & pVersaoInst)
   If Dir(pPathTemp, vbDirectory) = "" Then
      Call MkDir(pPathTemp)
   End If
   CreatePathTempExe = (Dir(pPathTemp, vbDirectory) <> "")
   Exit Function
TrataErro:
   CreatePathTempExe = False
End Function
Private Function DeletePathTempExe(ByRef pPathTemp As String, ByVal pArq As String, ByVal pVersaoInst As Double) As Boolean
   Dim sPath As String
   On Error GoTo TrataErro
   
   sPath = pPathTemp
   sPath = ClsDsr.ResolvePathName(sPath & Mid(pArq, 1, IIf(InStr(pArq, ".") = 0, Len(pArq), InStr(pArq, ".") - 1)))
   sPath = ClsDsr.ResolvePathName(sPath & pVersaoInst)
   
   If Dir(sPath, vbDirectory) <> "" Then
      While Dir(sPath, vbArchive) <> ""
         Call ClsDos.Del(sPath & Dir(sPath, vbArchive), False)
      Wend
      Call RmDir(sPath)
   End If
      
   sPath = pPathTemp
   sPath = ClsDsr.ResolvePathName(sPath & Mid(pArq, 1, IIf(InStr(pArq, ".") = 0, Len(pArq), InStr(pArq, ".") - 1)))
   If Dir(sPath, vbDirectory) <> "" Then
      While Dir(sPath, vbArchive) <> ""
         Call ClsDos.Del(sPath & Dir(sPath, vbArchive), False)
      Wend
      Call RmDir(sPath)
   End If
   
   sPath = pPathTemp
   If Dir(sPath, vbDirectory) <> "" Then
      While Dir(sPath, vbArchive) <> ""
         Call ClsDos.Del(sPath & Dir(sPath, vbArchive), False)
      Wend
      Call RmDir(sPath)
   End If
   
   DeletePathTempExe = (Dir(pPathTemp, vbDirectory) = "")
   Exit Function
TrataErro:
   DeletePathTempExe = False
End Function
Private Sub GravaAtualizaDll_Log(PathArquivoLog As String)
   Dim sArqLog As String
   
   If PathArquivoLog = "" Then PathArquivoLog = App.Path
   If InStr(PathArquivoLog, ".") = 0 Then
      sArqLog = ClsDsr.ResolvePathName(PathArquivoLog) & "AutoInstall.Log"
   Else
      sArqLog = PathArquivoLog
   End If
   
End Sub
Private Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub
Public Function AtualizaDLL(pDLLMaq As String, pDLLNova As String, Optional pObrigatorio As Boolean = False, Optional ArquivoLog As String) As Boolean
   Dim nVersaoMaq    As Double
   Dim nVersaoNova   As Double
   
   If ClsDos.FileExists(pDLLNova) Then
      nVersaoNova = GetVersao(GetFileVersion(pDLLNova))
      
      If ClsDos.FileExists(pDLLMaq) Then
         nVersaoMaq = GetVersao(GetFileVersion(pDLLMaq))
      Else
         nVersaoMaq = -1
      End If
   Else
      nVersaoNova = 0
      nVersaoMaq = 0
   End If
   If nVersaoNova > nVersaoMaq Or pObrigatorio Then
      If ClsDos.FileExists(pDLLMaq) Then
         Call ClsDsr.RegServer(pDLLMaq, False, False)
         Call Kill(pDLLMaq)
      End If
      
      Call FileCopy(pDLLNova, pDLLMaq)
      AtualizaDLL = ClsDsr.RegServer(pDLLMaq, True, False)
      Call GravaAtualizaDll_Log(ArquivoLog)
      
   ElseIf nVersaoNova <= nVersaoMaq Then
      AtualizaDLL = True
   End If
End Function
Public Function AtualizaExe(pEXEMaq As String, pEXENovo As String, Optional pRegistro As Boolean = False) As Boolean
   Dim nVersaoMaq    As Double
   Dim nVersaoNova   As Double
   
   On Error Resume Next
   nVersaoNova = -1
   nVersaoMaq = -1
   If ClsDos.FileExists(pEXENovo) Then
      nVersaoNova = GetVersao(GetFileVersion(pEXENovo))
      
      If ClsDos.FileExists(pEXEMaq) Then
         nVersaoMaq = GetVersao(GetFileVersion(pEXEMaq))
      End If
   Else
      nVersaoNova = 0
      nVersaoMaq = 0
   End If
   If nVersaoNova > nVersaoMaq Then
      If pRegistro Then
         Call ClsDsr.RegServer(pEXEMaq, False, False)
      End If
      Call Kill(ClsDsr.GetPathName(pEXEMaq) & "\*.rpt")
      
      Call ClsDos.Del(pEXEMaq)
      Call FileCopy(pEXENovo, pEXEMaq)
      
      AtualizaExe = True
      If pRegistro Then
         AtualizaExe = ClsDsr.RegServer(pEXEMaq, True, False)
      End If
   End If
End Function
Public Function GetRegisterDir(ByVal sArq As String, Optional AppPath As String) As String
   Dim sPath      As String
   Dim bAchou     As Boolean
   Dim sAux       As String
   
   
   sPath = ClsDsr.ResolvePathName(AppPath)
   bAchou = ClsDos.FileExists(AppPath & sArq)
   
   If Not bAchou Then
      sPath = ClsDsr.ResolvePathName(GetWindowsSysDir())
      sPath = ClsDsr.ResolvePathName(sPath & SubFolder)
      bAchou = ClsDos.FileExists(sPath & sArq)
   End If
   If Not bAchou Then
      sPath = ClsDsr.ResolvePathName(GetWindowsSysDir())
      bAchou = ClsDos.FileExists(sPath & sArq)
   End If
   If Not bAchou Then
      sPath = ClsDsr.ResolvePathName(GetSpecialFolder(CSIDL_PROGRAM_FILES))
      sPath = ClsDsr.ResolvePathName(sPath & SubFolder)
      bAchou = ClsDos.FileExists(sPath & sArq)
      If Not bAchou Then
         sPath = ProcuraArquivo(sPath, sArq)
         bAchou = ClsDos.FileExists(sPath & sArq)
      End If
   End If
   If Not bAchou Then
      sPath = ClsDsr.ResolvePathName(GetSpecialFolder(CSIDL_WINDOWS))
      bAchou = ClsDos.FileExists(sPath & sArq)
   End If
   If Not bAchou Then
      sPath = ClsDsr.ResolvePathName(GetSpecialFolder(CSIDL_COMMON))
      bAchou = ClsDos.FileExists(sPath & sArq)
   End If
   If Not bAchou Then
      sPath = ClsDsr.ResolvePathName(GetSpecialFolder(CSIDL_PROGRAM_FILES))
      bAchou = ClsDos.FileExists(sPath & sArq)
   End If
   
   sPath = IIf(bAchou, sPath, "")
   
   GetRegisterDir = ClsDsr.ResolvePathName(sPath)
End Function
Public Function ProcuraArquivo(ByVal pPath As String, ByVal pArq As String) As String
   Dim sAux    As String
   Dim sPath   As String
   Dim bAchou  As Boolean
   Dim sPath0  As String
   Dim i       As Integer
   Dim nVezes  As Integer
   
   On Error GoTo TrataErro
   nVezes = 10000
   
   sAux = ClsDsr.ResolvePathName(pPath)
   ChDir sAux
   sPath0 = sAux
   If ClsDos.FileExists(sAux & pArq) Then
      sAux = sPath
   Else
      sAux = Dir(sAux, vbDirectory)
      While sAux <> ""
         sAux = Dir(Attributes:=VbFileAttribute.vbDirectory)
         If InStr(sAux, ".") = 0 And sAux <> "" Then
            sAux = ClsDsr.ResolvePathName(pPath & sAux)
            If (GetAttr(sAux) And vbDirectory) = vbDirectory Then
               bAchou = ClsDos.FileExists(sAux & pArq)
               If bAchou Then
                  sPath = sAux
                  sAux = ""
               Else
                  sPath = ProcuraArquivo(sAux, pArq)
                  If sPath = "" Then
                     '*********************
                     '* Retorna ao diretório anterior
                     ChDir sPath0
                     If sAux <> Dir(pPath, vbDirectory) Then
                        i = 0
                        While i < nVezes
                           i = i + 1
                           If sAux = ClsDsr.ResolvePathName(pPath & Dir(Attributes:=vbDirectory)) Then
                              i = nVezes + 1
                           End If
                        Wend
                     End If
                  Else
                    sAux = ""
                  End If
               End If
            End If
         End If
      Wend
   End If
   ProcuraArquivo = sPath
Exit Function
TrataErro:
   If Err = 16 Then
      Resume Next
   Else
      MsgBox Err.Number & " - " & Err.Description, vbCritical, "TrataErro AutoInstall"
      Resume Next
   End If
End Function
Public Function GetDriveType(ByVal intDriveNum As Integer) As Integer
    '
    ' This function expects an integer drive number in Win16 or a string in Win32
    '
    Dim strDriveName As String

    strDriveName = Chr$(Asc("A") + intDriveNum) & gstrSEP_DRIVE & gstrSEP_DIR
    GetDriveType = CInt(GetDriveType32(strDriveName))
End Function
Public Function Extension(ByVal strFilename As String) As String
    Dim intPos As Integer

    Extension = gstrNULL

    intPos = Len(strFilename)

    Do While intPos > 0
        Select Case Mid$(strFilename, intPos, 1)
            Case gstrSEP_EXT
                Extension = Mid$(strFilename, intPos + 1)
                Exit Do
            Case gstrSEP_DIR, gstrSEP_DIRALT
                Exit Do
            'End Case
        End Select

        intPos = intPos - 1
    Loop
End Function
Public Function GetWindowsSysDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetSystemDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator(strBuf)
        AddDirSep strBuf

        GetWindowsSysDir = strBuf
    Else
        GetWindowsSysDir = gstrNULL
    End If
End Function
Public Function GetFileVersion(ByVal strFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
    Dim sVerInfo As VERINFO
    Dim strVer As String

    On Error GoTo GFVError

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If

    '
    'Get the file version into a VERINFO struct, and then assemble a version string
    'from the appropriate elements.
    '
    If GetFileVerStruct(strFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
        strVer = Format$(sVerInfo.nMSHi) & gstrDECIMAL & Format$(sVerInfo.nMSLo) & gstrDECIMAL
        strVer = strVer & Format$(sVerInfo.nLSHi) & gstrDECIMAL & Format$(sVerInfo.nLSLo)
        GetFileVersion = strVer
    Else
        GetFileVersion = gstrNULL
    End If

    Exit Function

GFVError:
    GetFileVersion = gstrNULL
    Err = 0
End Function
Private Function GetFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
    Const strFIXEDFILEINFO$ = "\"

    Dim lVerSize As Long
    Dim lVerHandle As Long
    Dim lpBufPtr As Long
    Dim byteVerData() As Byte
    Dim fFoundVer As Boolean

    GetFileVerStruct = False
    fFoundVer = False

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If

    If fIsRemoteServerSupportFile Then
        GetFileVerStruct = GetRemoteSupportFileVerStruct(strFilename, sVerInfo)
        fFoundVer = True
    Else
        '
        'Get the size of the file version info, allocate a buffer for it, and get the
        'version info.  Next, we query the Fixed file info portion, where the internal
        'file version used by the Windows VerInstallFile API is kept.  We then copy
        'the fixed file info into a VERINFO structure.
        '
        lVerSize = GetFileVersionInfoSize(strFilename, lVerHandle)
        If lVerSize > 0 Then
            ReDim byteVerData(lVerSize)
            If GetFileVersionInfo(strFilename, lVerHandle, lVerSize, byteVerData(0)) <> 0 Then ' (Pass byteVerData array via reference to first element)
                If VerQueryValue(byteVerData(0), strFIXEDFILEINFO & "", lpBufPtr, lVerSize) <> 0 Then
                    lmemcpy sVerInfo, lpBufPtr, lVerSize
                    fFoundVer = True
                    GetFileVerStruct = True
                End If
            End If
        End If
    End If

    If Not fFoundVer Then
        '
        ' We were unsuccessful in finding the version info from the file.
        ' One possibility is that this is a dependency file.
        '
        If UCase(Extension(strFilename)) = gstrEXT_DEP Then
            GetFileVerStruct = GetDepFileVerStruct(strFilename, sVerInfo)
        End If
    End If
End Function
Public Function GetVersao(sVer As String) As Double
   Dim Pos As Integer
   Pos = InStr(sVer, ".")
   While Pos <> 0
      sVer = Mid(sVer, 1, Pos - 1) & Mid(sVer, Pos + 1)
      Pos = InStr(sVer, ".")
   Wend
   sVer = IIf(sVer = "", sVer = "0", sVer)
   GetVersao = Val(sVer)
End Function
'Function ResolvePathName(ByVal sPath As String) As String
'   Dim PosIni As Integer
'   Dim PosFim As Integer
'
'   If Right(sPath, 1) <> "\" And Trim(sPath) <> "" Then
'      sPath = sPath & "\"
'   End If
'   If InStr(sPath, "%") <> 0 Then
'      PosIni = InStr(sPath, "%")
'      PosFim = InStr(PosIni + 1, sPath, "%")
'      sPath = Mid(sPath, 1, PosIni - 1) & Environ(Mid(sPath, PosIni + 1, PosFim - PosIni - 1)) & Mid(sPath, PosFim + 1)
'   End If
'
'   ResolvePathName = sPath
'End Function
'Public Function Encrypt(Psenha As String) As Variant
'    Dim v_sqlerrm    As String
'    Dim SenhaCript   As String
'    Dim var1         As String
'    Dim to_text      As String
'    Const MIN_ASC = 32
'    Const MAX_ASC = 126
'    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
'
'    'Chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
'    Dim offset As Long
'    Dim str_len As Integer
'    Dim i As Integer
'    Dim Ch As Integer
'
'    to_text = ""
'    offset = ClsSeg.NumericPassword(Chave)
'    Rnd -1
'    Randomize offset
'    str_len = Len(Psenha)
'    For i = 1 To str_len
'        Ch = Asc(Mid$(Psenha, i, 1))
'        If Ch >= MIN_ASC And Ch <= MAX_ASC Then
'            Ch = Ch - MIN_ASC
'            offset = Int((NUM_ASC + 1) * Rnd)
'            Ch = ((Ch + offset) Mod NUM_ASC)
'            Ch = Ch + MIN_ASC
'            to_text = to_text & Chr$(Ch)
'        End If
'    Next i
'
'    Encrypt = to_text
'End Function
'Public Function ComputerName() As String
'   Dim lpBuffer As String * 255
'   Dim lRet As Long
'   Dim i As Integer
'
'   lRet = GetComputerName(lpBuffer, 255)
'   If lRet <> 0 Then
'     ComputerName = UCase$(Left$(Trim$(lpBuffer), Len(Trim$(lpBuffer)) - 1))
'     ComputerName = ClsDsr.EliminarString(ComputerName, Chr(0))
'   Else
'     ComputerName = "Não Fornecido."
'   End If
'End Function
'Public Function EliminarString(ByVal Palavra$, ByVal Caracter$, Optional CaseSensitive = True) As String
'    Dim Pos%, Com_Carac$
'
'    If CaseSensitive Then
'       Pos% = InStr(Palavra, Caracter)
'    Else
'      Pos% = InStr(UCase(Palavra), UCase(Caracter))
'    End If
'
'    Com_Carac = Palavra$
'    While Pos% <> 0
'        Com_Carac = Left$(Com_Carac, Pos% - 1) + Mid$(Com_Carac, Pos% + Len(Caracter$))
'        If CaseSensitive Then
'           Pos% = InStr(Com_Carac$, Caracter)
'        Else
'           Pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
'        End If
'    Wend
'
'    If CaseSensitive Then
'       Pos% = InStr(Com_Carac$, Caracter)
'    Else
'       Pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
'    End If
'
'    If Pos% > 0 Then
'       If CaseSensitive Then
'          Pos% = InStr(Pos% + Len(Caracter$), Com_Carac, Caracter)
'       Else
'          Pos% = InStr(Pos% + Len(Caracter$), UCase(Com_Carac), UCase(Caracter))
'       End If
'       If Pos% > 0 Then Com_Carac = Left$(Com_Carac, Pos% - 1)
'    End If
'
'    EliminarString = Com_Carac
'End Function
'Public Function GetTag(ByRef Controle As Variant, ByVal VarName As String, Optional VarDefault As String) As String
'   Dim PosIni As Long, PosFim As Long
'   Dim StrTAG As String
'   Dim i%
'
'   On Error GoTo Saida
'
'   VarName = "|" & Trim(VarName) & "="
'
'   If UCase(TypeName(Controle)) = "STRING" Then
'      StrTAG = Controle
'   Else
'      StrTAG = Controle.Tag
'   End If
'
'   PosIni = InStr(StrTAG, Trim(VarName))
'   If PosIni > 0 Then
'      PosIni = PosIni + Len(Trim(VarName))
'      PosFim = InStr(PosIni, StrTAG$, "|")
'      i = 0
'      While Mid(StrTAG$, PosIni + i, 1) = "|"
'         i = i + 1
'      Wend
'      If i > 0 Then
'         PosFim = InStr(PosIni + (i - 1), StrTAG$, "|")
'      End If
'      PosFim = IIf(PosFim = 0, Len(StrTAG$), PosFim - 1)
'      StrTAG$ = Mid(StrTAG$, PosIni, PosFim - PosIni + 1)
'   Else
'      StrTAG$ = ""
'   End If
'   GetTag = StrTAG$
'Saida:
'   If StrTAG$ = "" Then
'      GetTag = VarDefault
'   End If
'End Function
'Public Function Decrypt(Psenha As String) As Variant
'
'Dim v_sqlerrm  As String
'Dim SenhaCript As String
'Dim var1       As String
'Dim to_text    As String
'
'Const MIN_ASC = 32  ' Space.
'Const MAX_ASC = 126 ' ~.
'Const NUM_ASC = MAX_ASC - MIN_ASC + 1
'
''Chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
'Dim offset As Long
'Dim str_len As Integer
'Dim i As Integer
'Dim ch As Integer
'
'to_text = ""
'offset = NumericPassword(Chave)
'Rnd -1
'Randomize offset
'str_len = Len(Psenha)
'For i = 1 To str_len
'    ch = Asc(Mid$(Psenha, i, 1))
'    If ch >= MIN_ASC And ch <= MAX_ASC Then
'        ch = ch - MIN_ASC
'        offset = Int((NUM_ASC + 1) * Rnd)
'        ch = ((ch - offset) Mod NUM_ASC)
'        If ch < 0 Then ch = ch + NUM_ASC
'        ch = ch + MIN_ASC
'        to_text = to_text & Chr$(ch)
'    End If
'Next i
'
'Decrypt = to_text
'
'End Function
'Private Function NumericPassword(ByVal password As String) As Long
'    Dim value As Long
'    Dim ch As Long
'    Dim shift1 As Long
'    Dim shift2 As Long
'    Dim i As Integer
'    Dim str_len As Integer
'
'    str_len = Len(password)
'    For i = 1 To str_len
'        ' Adiciona a próxima letra
'        ch = Asc(Mid$(password, i, 1))
'        value = value Xor (ch * 2 ^ shift1)
'        value = value Xor (ch * 2 ^ shift2)
'
'        ' Change the shift offse
'        shift1 = (shift1 + 7) Mod 19
'        shift2 = (shift2 + 13) Mod 23
'    Next i
'    NumericPassword = value
'End Function
'Public Sub Wait(Second As Integer)
'   Dim TimeIni&
'   TimeIni& = Timer
'   While Timer - TimeIni& < Second
''      DoEvents
'   Wend
'End Sub
'Function ResolveDestDir(ByVal strDestDir As String, Optional fAssumeDir As Variant) As String
'    Const strMACROSTART$ = "$("
'    Const strMACROEND$ = ")"
'
'    Dim intPos As Integer
'    Dim strResolved As String
'    Dim hKey As Long
'    Dim strPathsKey As String
'    Dim fQuoted As Boolean
'
'    If IsMissing(fAssumeDir) Then
'        fAssumeDir = True
'    End If
'
'    strPathsKey = RegPathWinCurrentVersion()
'    strDestDir = Trim(strDestDir)
'    '
'    ' If strDestDir is quoted when passed to this routine, it
'    ' should be quoted when it's returned.  The quotes need
'    ' to be temporarily removed, though, for processing.
'    '
'    If VBA.Left(strDestDir, 1) = gstrQUOTE Then
'        fQuoted = True
'        strDestDir = strUnQuoteString(strDestDir)
'    End If
'    '
'    ' We take the first part of destdir, and if its $( then we need to get the portion
'    ' of destdir up to and including the last paren.  We then test against this for
'    ' macro expansion.  If no ) is found after finding $(, then must assume that it's
'    ' just a normal file name and do no processing.  Only enter the case statement
'    ' if strDestDir starts with $(.
'    '
'    If VBA.Left$(strDestDir, 2) = strMACROSTART Then
'        intPos = InStr(strDestDir, strMACROEND)
'
'        Select Case VBA.Left$(strDestDir, intPos)
'            Case gstrAPPDEST
'                If gstrDestDir <> gstrNULL Then
'                    strResolved = gstrDestDir
'                Else
'                    strResolved = "c:"
'                End If
'            Case gstrWINDEST
'                strResolved = gstrWinDir
'            Case gstrWINSYSDEST, gstrWINSYSDESTSYSFILE
'                strResolved = gstrWinSysDir
'            Case gstrPROGRAMFILES
'                If TreatAsWin95() Then
'                    Const strProgramFilesKey = "ProgramFilesDir"
'
'                    If RegOpenKey(HKEY_LOCAL_MACHINE, strPathsKey, hKey) Then
'                        RegQueryStringValue hKey, strProgramFilesKey, strResolved
'                        RegCloseKey hKey
'                    End If
'                End If
'
'                If strResolved = "" Then
'                    'If not otherwise set, let strResolved be the root of the first fixed disk
'                    strResolved = strRootDrive()
'                End If
'            Case gstrCOMMONFILES
'                'First determine the correct path of Program Files\Common Files, if under Win95
'                strResolved = strGetCommonFilesPath()
'                If strResolved = "" Then
'                    'If not otherwise set, let strResolved be the Windows directory
'                    strResolved = gstrWinDir
'                End If
'            Case gstrCOMMONFILESSYS
'                'First determine the correct path of Program Files\Common Files, if under Win95
'                Dim strCommonFiles As String
'
'                strCommonFiles = strGetCommonFilesPath()
'                If strCommonFiles <> "" Then
'                    'Okay, now just add \System, and we're done
'                    strResolved = strCommonFiles & "System\"
'                Else
'                    'If Common Files isn't in the registry, then map the
'                    'entire macro to the Windows\{system,system32} directory
'                    strResolved = gstrWinSysDir
'                End If
'            Case gstrDAODEST
'                strResolved = strGetDAOPath()
'            Case Else
'                intPos = 0
'            'End Case
'        End Select
'    End If
'
'    If intPos <> 0 Then
'        AddDirSep strResolved
'    End If
'
'    If fAssumeDir = True Then
'        If intPos = 0 Then
'            '
'            'if no drive spec, and doesn't begin with any root path indicator ("\"),
'            'then we assume that this destination is relative to the app dest dir
'            '
'            If Mid$(strDestDir, 2, 1) <> gstrCOLON Then
'                If VBA.Left$(strDestDir, 1) <> gstrSEP_DIR Then
'                    strResolved = gstrDestDir
'                End If
'            End If
'        Else
'            If Mid$(strDestDir, intPos + 1, 1) = gstrSEP_DIR Then
'                intPos = intPos + 1
'            End If
'        End If
'    End If
'
'    If fQuoted = True Then
'        ResolveDestDir = strQuoteString(strResolved & Mid$(strDestDir, intPos + 1), True, False)
'    Else
'        ResolveDestDir = strResolved & Mid$(strDestDir, intPos + 1)
'    End If
'End Function
'Function RegPathWinCurrentVersion() As String
'    RegPathWinCurrentVersion = "SOFTWARE\Microsoft\Windows\CurrentVersion"
'End Function
'Public Function strUnQuoteString(ByVal strQuotedString As String)
''
'' This routine tests to see if strQuotedString is wrapped in quotation
'' marks, and, if so, remove them.
''
'    strQuotedString = Trim(strQuotedString)
'
'    If Mid$(strQuotedString, 1, 1) = gstrQUOTE And Right$(strQuotedString, 1) = gstrQUOTE Then
'        '
'        ' It's quoted.  Get rid of the quotes.
'        '
'        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
'    End If
'    strUnQuoteString = strQuotedString
'End Function
'Function TreatAsWin95() As Boolean
'    If IsWindows95() Then
'        TreatAsWin95 = True
'    ElseIf fNTWithShell() Then
'        TreatAsWin95 = True
'    Else
'        TreatAsWin95 = False
'    End If
'End Function
'Function IsWindows95() As Boolean
'    Const dwMask95 = &H2&
'    If GetWinPlatform() And dwMask95 Then
'        IsWindows95 = True
'    Else
'        IsWindows95 = False
'    End If
'End Function
'Function RegOpenKey(ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Boolean
'    Dim lResult As Long
'    Dim strHkey As String
'
'    On Error GoTo 0
'
'    strHkey = strGetHKEYString(hKey)
'
'    lResult = OSRegOpenKey(hKey, lpszSubKey, phkResult)
'    If lResult = ERROR_SUCCESS Then
'        RegOpenKey = True
'        AddHkeyToCache phkResult, strHkey & "\" & lpszSubKey
'    Else
'        RegOpenKey = False
'    End If
'End Function
'Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strData As String) As Boolean
'    Dim lResult As Long
'    Dim lValueType As Long
'    Dim strBuf As String
'    Dim lDataBufSize As Long
'
'    RegQueryStringValue = False
'    On Error GoTo 0
'    ' Get length/data type
'    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
'    If lResult = ERROR_SUCCESS Then
'        If lValueType = REG_SZ Then
'            strBuf = String(lDataBufSize, " ")
'            lResult = OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
'            If lResult = ERROR_SUCCESS Then
'                RegQueryStringValue = True
'                strData = StripTerminator(strBuf)
'            End If
'        End If
'    End If
'End Function
'
'Function RegCloseKey(ByVal hKey As Long) As Boolean
'    Dim lResult As Long
'
'    On Error GoTo 0
'    lResult = OSRegCloseKey(hKey)
'    RegCloseKey = (lResult = ERROR_SUCCESS)
'End Function
'Function strRootDrive() As String
'    Dim intDriveNum As Integer
'
'    For intDriveNum = 0 To Asc("Z") - Asc("A") - 1
'        If GetDriveType(intDriveNum) = intDRIVE_FIXED Then
'            strRootDrive = Chr$(Asc("A") + intDriveNum) & gstrCOLON & gstrSEP_DIR
'            Exit Function
'        End If
'    Next intDriveNum
'
'    strRootDrive = "C:\"
'End Function
'Public Function strGetCommonFilesPath() As String
'    Dim hKey As Long
'    Dim strPath As String
'
'    If TreatAsWin95() Then
'        Const strCommonFilesKey = "CommonFilesDir"
'
'        If RegOpenKey(HKEY_LOCAL_MACHINE, RegPathWinCurrentVersion(), hKey) Then
'            RegQueryStringValue hKey, strCommonFilesKey, strPath
'            RegCloseKey hKey
'        End If
'    End If
'
'    If strPath <> "" Then
'        AddDirSep strPath
'    End If
'
'    strGetCommonFilesPath = strPath
'End Function
'Private Function strGetDAOPath() As String
'    Const strMSAPPS$ = "MSAPPS\"
'    Const strDAO3032$ = "DAO3032.DLL"
'
'    'first look in the registry
'    Const strKey = "SOFTWARE\Microsoft\Shared Tools\DAO"
'    Const strValueName = "Path"
'    Dim hKey As Long
'    Dim strPath As String
'
'    If RegOpenKey(HKEY_LOCAL_MACHINE, strKey, hKey) Then
'        RegQueryStringValue hKey, strValueName, strPath
'        RegCloseKey hKey
'    End If
'
'    If strPath <> "" Then
'        strPath = GetPathName(strPath)
'        AddDirSep strPath
'        strGetDAOPath = strPath
'        Exit Function
'    End If
'
'    'It's not yet in the registry, so we need to decide
'    'where the directory should be, and then need to place
'    'that location in the registry.
'
'    If TreatAsWin95() Then
'        'For Win95, use "Common Files\Microsoft Shared\DAO"
'        ' strPath = strGetCommonFilesPath() & ResolveResString(resMICROSOFTSHARED) & "DAO\"
'        strPath = strGetCommonFilesPath() & "Microsoft Shared\" & "DAO\"
'    Else
'        'Otherwise use Windows\MSAPPS\DAO
'        strPath = gstrWinDir & strMSAPPS & "DAO\"
'    End If
'
'    'Place this information in the registry (note that we point to DAO3032.DLL
'    'itself, not just to the directory)
'    If RegCreateKey(HKEY_LOCAL_MACHINE, strKey, "", hKey) Then
'        RegSetStringValue hKey, strValueName, strPath & strDAO3032, False
'        RegCloseKey hKey
'    End If
'
'    strGetDAOPath = strPath
'End Function
'Public Function strQuoteString(strUnQuotedString As String, Optional vForce As Variant, Optional vTrim As Variant)
''
'' This routine adds quotation marks around an unquoted string, by default.  If the string is already quoted
'' it returns without making any changes unless vForce is set to True (vForce defaults to False) except that white
'' space before and after the quotes will be removed unless vTrim is False.  If the string contains leading or
'' trailing white space it is trimmed unless vTrim is set to False (vTrim defaults to True).
''
'    Dim strQuotedString As String
'
'    If IsMissing(vForce) Then
'        vForce = False
'    End If
'    If IsMissing(vTrim) Then
'        vTrim = True
'    End If
'
'    strQuotedString = strUnQuotedString
'    '
'    ' Trim the string if necessary
'    '
'    If vTrim = True Then
'        strQuotedString = Trim(strQuotedString)
'    End If
'    '
'    ' See if the string is already quoted
'    '
'    If vForce = False Then
'        If (VBA.Left(strQuotedString, 1) = gstrQUOTE) And (Right(strQuotedString, 1) = gstrQUOTE) Then
'            '
'            ' String is already quoted.  We are done.
'            '
'            GoTo DoneQuoteString
'        End If
'    End If
'    '
'    ' Add the quotes
'    '
'    strQuotedString = gstrQUOTE & strQuotedString & gstrQUOTE
'DoneQuoteString:
'    strQuoteString = strQuotedString
'End Function
'Private Function strGetHKEYString(ByVal hKey As Long) As String
'    Dim strKey As String
'
'    'Is the hkey predefined?
'    strKey = strGetPredefinedHKEYString(hKey)
'    If strKey <> "" Then
'        strGetHKEYString = strKey
'        Exit Function
'    End If
'
'    'It is not predefined.  Look in the cache.
'    Dim intIdx As Integer
'    intIdx = intGetHKEYIndex(hKey)
'    If intIdx >= 0 Then
'        strGetHKEYString = hkeyCache(intIdx).strHkey
'    Else
'        strGetHKEYString = ""
'    End If
'End Function
'Private Sub AddHkeyToCache(ByVal hKey As Long, ByVal strHkey As String)
'    Dim intIdx As Integer
'
'    intIdx = intGetHKEYIndex(hKey)
'    If intIdx < 0 Then
'        'The key does not already exist.  Add it to the end.
'        On Error Resume Next
'        ReDim Preserve hkeyCache(0 To UBound(hkeyCache) + 1)
'        If Err Then
'            'If there was an error, it means the cache was empty.
'            On Error GoTo 0
'            ReDim hkeyCache(0 To 0)
'        End If
'        On Error GoTo 0
'
'        intIdx = UBound(hkeyCache)
'    Else
'        'The key already exists.  It will be replaced.
'    End If
'
'    hkeyCache(intIdx).hKey = hKey
'    hkeyCache(intIdx).strHkey = strHkey
'End Sub
'Function GetPathName(ByVal strFilename As String) As String
'    Dim intPos As Integer
'    Dim strPathOnly As String
'    Dim dirTmp As DirListBox
'    Dim i As Integer
'
'    On Error Resume Next
'
'
'    Err = 0
'
'    intPos = Len(strFilename)
'
'    '
'    'Change all '/' chars to '\'
'    '
'
'    For i = 1 To Len(strFilename)
'        If Mid$(strFilename, i, 1) = gstrSEP_DIRALT Then
'            Mid$(strFilename, i, 1) = gstrSEP_DIR
'        End If
'    Next i
'
'    If InStr(strFilename, gstrSEP_DIR) = intPos Then
'        If intPos > 1 Then
'            intPos = intPos - 1
'        End If
'    Else
'        Do While intPos > 0
'            If Mid$(strFilename, intPos, 1) <> gstrSEP_DIR Then
'                intPos = intPos - 1
'            Else
'                Exit Do
'            End If
'        Loop
'    End If
'
'    If intPos > 0 Then
'        strPathOnly = VBA.Left$(strFilename, intPos)
'        If Right$(strPathOnly, 1) = gstrCOLON Then
'            strPathOnly = strPathOnly & gstrSEP_DIR
'        End If
'    Else
'        strPathOnly = CurDir$
'    End If
'
'    If Right$(strPathOnly, 1) = gstrSEP_DIR Then
'        strPathOnly = VBA.Left$(strPathOnly, Len(strPathOnly) - 1)
'    End If
'
'    GetPathName = UCase16(strPathOnly)
'
'    Err = 0
'End Function
'Function RegCreateKey(ByVal hKey As Long, ByVal lpszSubKeyPermanent As String, ByVal lpszSubKeyRemovable As String, phkResult As Long) As Boolean
'    Dim lResult As Long
'    Dim strHkey As String
'    Dim fLog As Boolean
'    Dim strSubKeyFull As String
'
'    On Error GoTo 0
'
'    If lpszSubKeyPermanent = "" Then
'        RegCreateKey = False 'Error: lpszSubKeyPermanent must not = ""
'        Exit Function
'    End If
'
'    If VBA.Left$(lpszSubKeyRemovable, 1) = "\" Then
'        lpszSubKeyRemovable = Mid$(lpszSubKeyRemovable, 2)
'    End If
'
'    If lpszSubKeyRemovable = "" Then
'        fLog = False
'    Else
'        fLog = True
'    End If
'
'    If lpszSubKeyRemovable <> "" Then
'        strSubKeyFull = lpszSubKeyPermanent & "\" & lpszSubKeyRemovable
'    Else
'        strSubKeyFull = lpszSubKeyPermanent
'    End If
'    strHkey = strGetHKEYString(hKey)
'
'    If fLog Then
'        NewAction _
'          gstrKEY_REGKEY, _
'          """" & strHkey & "\" & lpszSubKeyPermanent & """" _
'            & ", " & """" & lpszSubKeyRemovable & """"
'    End If
'
'    lResult = OSRegCreateKey(hKey, strSubKeyFull, phkResult)
'    If lResult = ERROR_SUCCESS Then
'        RegCreateKey = True
'        If fLog Then
'            CommitAction
'        End If
'        AddHkeyToCache phkResult, strHkey & "\" & strSubKeyFull
'    Else
'        RegCreateKey = False
'        MsgBox "An error occurred trying to update the Windows registration database.", vbOKOnly Or vbExclamation, gstrTitle
'        'MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
'        If fLog Then
'            AbortAction
'        End If
'        If gfNoUserInput Then
'            'ExitSetup frmSetup1, gintRET_FATAL
'        End If
'    End If
'End Function
'Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strData As String, Optional ByVal fLog) As Boolean
'    Dim lResult As Long
'    Dim strHkey As String
'
'    On Error GoTo 0
'
'    If IsMissing(fLog) Then fLog = True
'
'    If hKey = 0 Then
'        Exit Function
'    End If
'
'    strHkey = strGetHKEYString(hKey)
'
'    If fLog Then
'        NewAction _
'          gstrKEY_REGVALUE, _
'          """" & strHkey & """" _
'            & ", " & """" & strValueName & """"
'    End If
'
'    lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)) + 1)
'
'    If lResult = ERROR_SUCCESS Then
'        RegSetStringValue = True
'        If fLog Then
'            CommitAction
'        End If
'    Else
'        RegSetStringValue = False
'        MsgBox "An error occurred trying to update the Windows registration database.", vbOKOnly Or vbExclamation, gstrTitle
'        'MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
'        If fLog Then
'            AbortAction
'        End If
'        If gfNoUserInput Then
'            'ExitSetup frmSetup1, gintRET_FATAL
'        End If
'    End If
'End Function
'Private Function strGetPredefinedHKEYString(ByVal hKey As Long) As String
'    Select Case hKey
'        Case HKEY_CLASSES_ROOT
'            strGetPredefinedHKEYString = "HKEY_CLASSES_ROOT"
'        Case HKEY_CURRENT_USER
'            strGetPredefinedHKEYString = "HKEY_CURRENT_USER"
'        Case HKEY_LOCAL_MACHINE
'            strGetPredefinedHKEYString = "HKEY_LOCAL_MACHINE"
'        Case HKEY_USERS
'            strGetPredefinedHKEYString = "HKEY_USERS"
'        'End Case
'    End Select
'End Function
'Private Function intGetHKEYIndex(ByVal hKey As Long) As Integer
'    Dim intUBound As Integer
'
'    On Error Resume Next
'    intUBound = UBound(hkeyCache)
'    If Err Then
'        'If there was an error accessing the ubound of the array,
'        'then the cache is empty
'        GoTo NotFound
'    End If
'    On Error GoTo 0
'
'    Dim intIdx As Integer
'    For intIdx = 0 To intUBound
'        If hkeyCache(intIdx).hKey = hKey Then
'            intGetHKEYIndex = intIdx
'            Exit Function
'        End If
'    Next intIdx
'
'NotFound:
'    intGetHKEYIndex = -1
'End Function
'Function UCase16(ByVal str As String)
'    UCase16 = str
'End Function
'Sub NewAction(ByVal strKey As String, ByVal strData As String)
'    ShowLoggingError DllNewAction(strKey, strData), LogErrFatal
'End Sub
'Sub CommitAction()
'    ShowLoggingError DllCommitAction(), LogErrFatal
'End Sub
'Sub AbortAction()
'    ShowLoggingError DllAbortAction(), LogErrFatal
'End Sub
'Sub ShowLoggingError(ByVal lErr As Long, ByVal lErrSeverity As Long)
'    If lErr = LOGERR_SUCCESS Then
'        Exit Sub
'    End If
'
'    Dim strErrMsg As String
'    Static fRecursive As Boolean
'
'    If fRecursive Then
'        'If we're getting called recursively, we're likely
'        'getting errors while trying to write out errors to
'        'the logfile.  Nothing to do but turn off logging
'        'and abort setup.
'        DisableLogging
'        'MsgError ResolveResString(resUNEXPECTED), vbExclamation Or vbOKOnly, gstrTitle
'        MsgBox "An unexpected setup error has occurred!", vbExclamation Or vbOKOnly, gstrTitle
'        'ExitSetup frmSetup1, gintRET_FATAL
'    End If
'
'    fRecursive = True
'
'    Select Case lErr
'        Case LOGERR_OUTOFMEMORY, LOGERR_WRITEERROR, LOGERR_UNEXPECTED, LOGERR_FILENOTFOUND
'            'strErrMsg = ResolveResString(resUNEXPECTED)
'            strErrMsg = "An unexpected setup error has occurred!"
'
'        Case LOGERR_INVALIDARGS, LOGERR_EXCEEDEDCAPACITY, LOGERR_NOCURRENTACTION
'            'Note: These errors are most likely the result of improper customization
'            'of this project.  Make certain that any changes you have made to these
'            'files are valid and bug-free.
'            'LOGERR_INVALIDARGS -- some parameter to a logging function was invalid or improper
'            'LOGERR_EXCEEDEDCAPACITY -- the stacking depth of actions has probably been
'            '   exceeded.  This most likely means that CommitAction or AbortAction statements
'            '   are missing from your code.
'            'LOGERR_NOCURRENTACTION -- the logging function you tried to use requires that
'            '   there be a current action, but there was none.  Check for a missing NewAction
'            '   statement.
'            'strErrMsg = ResolveResString(resUNEXPECTED)
'            strErrMsg = "An unexpected setup error has occurred!"
'        Case Else
'            'strErrMsg = ResolveResString(resUNEXPECTED)
'            strErrMsg = "An unexpected setup error has occurred!"
'        'End Case
'    End Select
'
'    Dim iRet As Integer
'    Dim fAbort As Boolean
'
'    fAbort = False
'    If lErrSeverity = LogErrOK Then
'        ' User can select whether or not to continue
'        iRet = MsgBox(strErrMsg, vbOKCancel Or vbExclamation, gstrTitle)
'        If gfNoUserInput Then iRet = vbCancel ' can't continue if silent install.
'        Select Case iRet
'            Case vbOK
'            Case vbCancel
'                fAbort = True
'            Case Else
'                fAbort = True
'            'End Case
'        End Select
'    Else
'        ' Fatal
'        MsgBox strErrMsg, vbOK Or vbExclamation, gstrTitle
'        fAbort = True
'    End If
'
'    If fAbort Then
'        'ExitSetup frmCopy, gintRET_ABORT
'    End If
'
'    fRecursive = False
'
'End Sub
'Sub DisableLogging()
'    ShowLoggingError DllDisableLogging(), LogErrFatal
'End Sub
'Function FileExists(ByVal strPathName As String) As Integer
'    Dim intFileNum As Integer
'
'    On Error Resume Next
'
'    '
'    ' If the string is quoted, remove the quotes.
'    '
'    strPathName = strUnQuoteString(strPathName)
'    '
'    'Remove any trailing directory separator character
'    '
'    If Right$(strPathName, 1) = gstrSEP_DIR Then
'        strPathName = VBA.Left$(strPathName, Len(strPathName) - 1)
'    End If
'
'    '
'    'Attempt to open the file, return value of this function is False
'    'if an error occurs on open, True otherwise
'    '
'    intFileNum = FreeFile
'    Open strPathName For Input As intFileNum
'
'    FileExists = IIf(Err = 0, True, False)
'
'    Close intFileNum
'
'    Err = 0
'End Function
'Function GetWindowsDir() As String
'    Dim strBuf As String
'
'    strBuf = Space$(gintMAX_SIZE)
'
'    '
'    'Get the windows directory and then trim the buffer to the exact length
'    'returned and add a dir sep (backslash) if the API didn't return one
'    '
'    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
'        strBuf = StripTerminator$(strBuf)
'        AddDirSep strBuf
'
'        GetWindowsDir = strBuf
'    Else
'        GetWindowsDir = gstrNULL
'    End If
'End Function
''Public Function SincShell(Comando As String, Optional Modo = vbMinimizedFocus)
''    Dim Handle As Long
''    Dim Inst1 As Long
''    Dim Inst2 As Long
''    Dim Res As Long
''    Dim StrWinModule As String
''    'Objetivo: Fazer um loop para controlar o fim da última aplicacao DOS gerada
''    StrWinModule = "WINOLDAP"
''    Handle = GetModuleHandle(StrWinModule) 'Handle do Módulo WINOLDAP que carrega as apps DOS
''    Inst1 = GetModuleUsage(Handle)         'Número de aplicações DOS ativas antes de executar a atual
''
''    'SHELL -> roda um pgm executável
''
''    Res = Shell(Comando, Modo)
''    Handle = GetModuleHandle(StrWinModule)  'Handle do Módulo WINOLDAP que carrega as apps DOS
''    Inst2 = GetModuleUsage(Handle)        'Novo numero de aplicacoes DOS (inst2=inst1+1)  (depois de executar o shell acima)
''    While Not (Inst2 = Inst1)
''        Inst2 = GetModuleUsage(Handle)    'Aguarda o fim da última aplicação gerada (que no caso e a chamada do shell acima)
''        DoEvents
''    Wend
''    SincShell = Res
''
''End Function
'Public Sub AutoInstala(Aplic As App)
'   On Error Resume Next
'   Dim strFilename As String
'
'   Dim FileLST As String, strKey As String, sLinha As String
'
'   Dim i As Integer, J As Integer
'   Dim Pos As Integer, PosFim As Integer
'   Dim sExt As String, bExtensao As Boolean
'   Dim bInstala As Boolean
'
'   Dim sNome As String, sPath As String
'   Dim sVersaoInst As String, sVersaoMaq As String
'   Dim lComando As String, lFileComando As String
'   Dim NumProp As Integer
'   Dim nArq As Integer, sArq As String
'   Dim gstrINI_FILES$
'   gstrINI_FILES$ = "Setup1 Files"
'
'   FileLST = Aplic.Path & "\SETUP\SETUP.LST"
'
'   If Not FileExists(FileLST) Then
'      FileLST = Aplic.Path & "\Package\SETUP.LST"
'      If Not FileExists(FileLST) Then
'         Exit Sub
'      End If
'   End If
'
'   gstrWinDir = GetWindowsDir()
'   gstrWinSysDir = GetWindowsSysDir()
'   gstrDestDir = ResolveDestDir(ClsAutoIns.ReadIniFile(FileLST, "Setup", "DefaultDir"))
'
'   i = 0
'   nArq = FreeFile()
'   sArq = GetPathName(FileLST) & "\AutoVer.log"
'   Kill sArq
'   Open sArq For Output As #nArq
'   Print #nArq, "Auto-Instalação em " & Format(Now, "dd/mm/yyyy hh:mm")
'   Print #nArq, "|----------------|----------------|----------------|"
'   Print #nArq, "| Arquivo        | Versão Atual   | Versão Setup   |"
'   Print #nArq, "|----------------|----------------|----------------|"
'
'   Do
'      i = i + 1
'      strKey = "File" & Trim(CStr(i))
'      sLinha = ClsAutoIns.ReadIniFile(FileLST, gstrINI_FILES$, strKey)
'      If Trim(sLinha) <> "" Then
'         Pos = 2
'         PosFim = InStr(Pos, sLinha, ",")
'         NumProp = 6
'         For J = 0 To NumProp
'            If J = 0 Then sNome = Mid(sLinha, Pos, PosFim - Pos)
'            If J = 1 Then sPath = ResolveDestDir(Mid(sLinha, Pos, PosFim - Pos))
'            If J = 6 Then sVersaoInst = Mid(sLinha, Pos, PosFim - Pos)
'            Pos = InStr(Pos, sLinha, ",")
'            Pos = Pos + 1
'               If Pos <= 1 Then Exit For
'            If InStr(Pos, sLinha, ",") = 0 Then
'               PosFim = Len(sLinha) + 1
'            Else
'               PosFim = InStr(Pos, sLinha, ",")
'            End If
'         Next
'         sExt = UCase(Extension(sNome))
'         bExtensao = (sExt = "DLL" Or sExt = "OCX" Or sExt = "TLB")
'         If Pos > 0 And bExtensao Then
'            If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
'            strFilename = sPath & sNome
'            sVersaoMaq = GetFileVersion(strFilename)
'            If sVersaoInst > sVersaoMaq Or Not FileExists(strFilename) Then
'               '* Gravalog
'               Print #nArq, "| " & Mid(sNome & Space(15), 1, 15) & "| " & Mid(sVersaoMaq & Space(15), 1, 15) & "| " & Mid(sVersaoInst & Space(15), 1, 15) & "|"
'               Print #nArq, "|----------------|----------------|----------------|"
'               bInstala = True
'            End If
'         End If
'      End If
'   Loop Until Trim(sLinha) = ""
'   Close #nArq
'   If bInstala Then
'      lFileComando = GetPathName(FileLST) & "\Instala.exe"
'      If Not FileExists(lFileComando) Then
'         lFileComando = GetPathName(FileLST) & "\Setup.exe"
'      End If
'      lComando = lFileComando & " /s AutoIns.log"
'      If FileExists(lFileComando) Then
'         Call SincShell(lComando)
'      End If
'   Else
'      Kill sArq
'   End If
'Exit Sub
'OpenError:
''   CmDialog.filename = ""
'   Resume Next
'End Sub
'


