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
   Dim sLocalReg     As String
   Dim LerSecaoIni   As Boolean
   Dim sStatusSecao  As String
   
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
   sPathTemp = UCase(sPathTemp & SubFolder & "\")
      
   '***********
   '* AppPath = "C:\Program Files\ClasseA\Producao"
   '* PathSetup = "\\guarani\sistemas\admin\"
   AppPath = UCase(ResolvePathName(AppPath))
   sPathExe = UCase(GetTag(pCommand, "SOURCEPATH", ""))
   SetupFile = UCase(IIf(SetupFile = "", "SETUP.INI", SetupFile))
   PathSetup = UCase(IIf(PathSetup = "", AppPath, PathSetup))
   If Trim(GetTag(pCommand, "PATHSETUP", "")) <> "" Then
      PathSetup = UCase(Trim(GetTag(pCommand, "PATHSETUP", "")))
   End If
   
   LocalSetupIni = AppPath & SetupFile
   SetupIni = PathSetup & SetupFile
   
   sLocalReg = ResolvePathName(AppPath) & pExeName & ".reg"
   
   If ExisteArquivo(LocalSetupIni) Then
      PathSetup = AppPath
   Else
      If Not ExisteArquivo(SetupIni) Then
         If ExisteArquivo(sLocalReg) Then
            PathSetup = UCase(ResolvePathName(ReadIniFile(sLocalReg, "Setup", "PATHSETUP")))
         Else
            PathSetup = UCase(ResolvePathName(GetSetting(pCODSIS, "Setup", "PATHSETUP")))
         End If
      End If
   End If
   
   SetupIni = PathSetup & SetupFile
   
   If bDebug Then MsgBox "SetupIni = " & SetupIni
   If ExisteArquivo(SetupIni) Then
      If bDebug Then
         MsgBox "If ExisteArquivo(" & SetupIni & ") Then 'R: True"
      End If

      '************
      '* Define Conexão Padrão [Database Format]
      '************
      If ExisteArquivo(sLocalReg) Then
         Call RegistroConn(sLocalReg, SetupIni)
      End If

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
         sPathNewFile = ResolvePathName(ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "Path", PathSetup))
         sArqExe = pExeName & IIf(Right(pExeName, 4) = ".exe", "", ".exe")
         If Not ExisteArquivo(sPathNewFile & sArqExe) Then
            sPathNewFile = ResolvePathName(ReadIniFile(SetupIni, "AutoInstall Files", "Path", PathSetup))
         End If
         
         If bDebug Then
            MsgBox "ExisteArquivo(" & sPathNewFile & sArqExe & ") " & IIf(ExisteArquivo(SetupIni), "= True ", "= False")
         End If
         
         If ExisteArquivo(sPathNewFile & sArqExe) Then
            nVersaoInst = GetFileVersionNumber(sPathNewFile & sArqExe)
            If sPathExe = "" Then
               nVersaoMaq = GetFileVersionNumber(AppPath & sArqExe)
            Else
               nVersaoMaq = GetFileVersionNumber(sPathExe & sArqExe)
            End If
            If bDebug Then
               MsgBox sPathNewFile & sArqExe & " Vs." & CStr(nVersaoInst) & vbNewLine & AppPath & sArqExe & " Vs." & CStr(nVersaoMaq)
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
                     nVersaoTemp = GetFileVersionNumber(sPathTempExe & sArqExe)
                     While (nVersaoInst > nVersaoTemp) And (i < 10)
                        If bDebug Then MsgBox " Tentativa " & i & " - Call FileCopy(" & sPathNewFile & sArqExe & ", " & sPathTempExe & sArqExe & ")"
                        Call FileCopy(sPathNewFile & sArqExe, sPathTempExe & sArqExe)
                        nVersaoTemp = GetFileVersionNumber(sPathTempExe & sArqExe)
                        i = i + 1
                     Wend
'                     If Err <> 0 Then '* Path not found
'                        MsgBox Err.Number & " - " & Err.Description, vbCritical, "AutoInstall [" & sPathNewFile & sArqExe & "] To [" & sPathTempExe & sArqExe & "]"
'                     End If
                     Wait 3
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
                     Call Del(sPathExe & sArqExe, False)
                     Call FileCopy(sPathNewFile & sArqExe, sPathExe & sArqExe)
                     nVersaoMaq = GetFileVersionNumber(sPathExe & sArqExe)
                     i = i + 1
                  Wend
                  Wait 3
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
            If GetTag(pCommand, "CLEARTEMP", "") = "TRUE" Then
               Wait 3
               Call DeletePathTempExe(sPathTemp, sArqExe, nVersaoInst)
            End If
         End If
         
         i = 1
         '************
         '* Verifica Objetos Comuns
         '************
         sStatusSecao = ""
         LerSecaoIni = False
         sStatusSecao = ReadIniFile(SetupIni, "AutoInstall Files", "Status", "")
         sStatusSecao = IIf(sStatusSecao = "", sStatus, sStatusSecao)
         LerSecaoIni = (sStatusSecao = "1" Or sStatusSecao = "2")
         If LerSecaoIni Then
            While ReadIniFile(SetupIni, "AutoInstall Files", "File" & CStr(i)) <> ""
               sArq = ReadIniFile(SetupIni, "AutoInstall Files", "File" & CStr(i))
               '**********
               '* Recupera pasta do arquivo a ser registrado
               sPathSis = ""
               sPathSis = GetRegisterDir(sArq, AppPath)
               If sPathSis = "" Then
                  If bDebug Then MsgBox "Não Achou Arquivo '" & sArq & "'"
                  
                  sPathSis = ReadIniFile(SetupIni, "AutoInstall Files", "Path" & CStr(i))
                  If bDebug Then MsgBox "ResolvePathName(" & sPathSis & ") R: " & ResolvePathName(sPathSis)
               Else
                  If bDebug Then MsgBox "Achou Arquivo '" & sArq & "' em  " & sPathSis
               End If
               sPathSis = ResolvePathName(sPathSis)
               If sPathSis <> "" Then
               
                  If bDebug Then
                     sMsg = "Arquivo " & CStr(i) & vbNewLine
                     sMsg = sMsg & sPathSis & sArq & " Vs." & GetFileVersion(sPathSis & sArq) & vbNewLine
                     sMsg = sMsg & sPathNewFile & sArq & " Vs." & GetFileVersion(sPathNewFile & sArq)
                     MsgBox sMsg
                  End If
                  If Not AtualizaDLL(sPathSis & sArq, sPathNewFile & sArq, (sStatusSecao = "2")) Then
                     Call AtualizaDLL(sPathSis & sArq, sPathNewFile & sArq, (sStatusSecao = "2"))
                  End If
               End If
               i = i + 1
            Wend
         End If
         
         i = 1
         '************
         '* Verifica Objetos do Módulo
         '************
         sStatusSecao = ""
         LerSecaoIni = False
         sStatusSecao = ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "Status", "")
         sStatusSecao = IIf(sStatusSecao = "", sStatus, sStatusSecao)
         LerSecaoIni = (sStatusSecao = "1" Or sStatusSecao = "2")
         While ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "File" & CStr(i)) <> ""
            sArq = ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "File" & CStr(i))
            '**********
            '* Recupera pasta do arquivo a ser registrado
            sPathSis = ""
            sPathSis = GetRegisterDir(sArq, AppPath)
            If sPathSis = "" Then
               If bDebug Then MsgBox "Não Achou Arquivo '" & sArq & "'"
               
               sPathSis = ReadIniFile(SetupIni, pCODSIS & " AutoInstall Files", "Path" & CStr(i))
               If bDebug Then MsgBox "ResolvePathName(" & sPathSis & ") R: " & ResolvePathName(sPathSis)
            Else
               If bDebug Then MsgBox "Achou Arquivo '" & sArq & "' em  " & sPathSis
            End If
            sPathSis = ResolvePathName(sPathSis)
            If sPathSis <> "" Then
               If bDebug Then
                  sMsg = "Arquivo " & CStr(i) & vbNewLine
                  sMsg = sMsg & sPathSis & sArq & " Vs." & GetFileVersion(sPathSis & sArq) & vbNewLine
                  sMsg = sMsg & sPathNewFile & sArq & " Vs." & GetFileVersion(sPathNewFile & sArq)
                  MsgBox sMsg
               End If
               If Not AtualizaDLL(sPathSis & sArq, sPathNewFile & sArq, (sStatusSecao = "2")) Then
                  Call AtualizaDLL(sPathSis & sArq, sPathNewFile & sArq, (sStatusSecao = "2"))
               End If
            End If
            i = i + 1
         Wend
         '************
         '* Verifica Objetos Apagados
         '************
         i = 1
         While ReadIniFile(SetupIni, "Deleted Files", "File" & CStr(i)) <> ""
            sArq = ReadIniFile(SetupIni, "Deleted Files", "File" & CStr(i))
            '**********
            '* Recupera pasta do arquivo a ser apagado
            sPathSis = ""
            sPathSis = ResolvePathName(ReadIniFile(SetupIni, "Deleted Files", "Path" & CStr(i)))
            
            If sPathSis <> "" Then
               If ExisteArquivo(sPathSis & sArq) Then
                  If bDebug Then MsgBox "Kill(" & sPathSis & sArq & ")"
                  Call RegServer(sPathSis & sArq, False, False)
                  Call Kill(sPathSis & sArq)
                  
                  '**************
                  '* Localiza Arquivo e registra em local correto
                  sPathSis = GetRegisterDir(sArq, AppPath)
                  sPathSis = ResolvePathName(sPathSis)
                  If sPathSis <> "" Then
                     Call RegServer(sPathSis & sArq, True, False)
                  End If
               End If
            End If
            i = i + 1
         Wend
      End If
   Else
      If bDebug Then
         MsgBox "If ExisteArquivo(" & SetupIni & ") Then 'R: False"
      End If
   End If
   VerificaSetup = True
   Exit Function
   
OpenError:
   Resume Next
End Function
Private Sub RegistroConn(pLocalReg As String, pSetupIni As String)
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
   
   SecaoBase = "Conection "
   Index = 0
   sSecao = SecaoBase & Index
   
   If ExisteArquivo(pSetupIni) Then
      If ReadIniFile(pSetupIni, "General", "Status", "0") = "0" Then Exit Sub
      If ReadIniFile(pSetupIni, "Database", "Status", "0") = "0" Then Exit Sub
   End If
   
   sALIAS = UCase(SubFolder)
   nDBTIPO = 1
   sDBVERSAO = ""
   bisODBC = False
   bisADO = True
   sSERVER = ComputerName & "\DSVM"
   sDBNAME = sALIAS
   sDSN = ""
   sUID = "USU_VERIF"
   sPWD = Decrypt2("757C717D64747978756C74")
   
   If Trim(ReadIniFile(pLocalReg, sSecao, "ALIAS", "")) <> "" Then
      sALIAS = ReadIniFile(pLocalReg, sSecao, "ALIAS", sALIAS)
      nDBTIPO = ReadIniFile(pLocalReg, sSecao, "DBTIPO", CStr(nDBTIPO))
      sDBVERSAO = ReadIniFile(pLocalReg, sSecao, "DBVERSAO")
      bisADO = ReadIniFile(pLocalReg, sSecao, "isADO", CStr(bisADO))
      bisODBC = ReadIniFile(pLocalReg, sSecao, "isODBC", CStr(bisODBC))
      sSERVER = ReadIniFile(pLocalReg, sSecao, "SERVER", sSERVER)
      sDBNAME = ReadIniFile(pLocalReg, sSecao, "DBNAME", sDBNAME)
      sDSN = ReadIniFile(pLocalReg, sSecao, "DSN", sDSN)
      sUID = ReadIniFile(pLocalReg, sSecao, "UID", sUID)
      sPWD = Decrypt2(ReadIniFile(pLocalReg, sSecao, "PWD", IIf(sUID = "USU_VERIF", Encrypt2(sPWD), "")))
   End If
   If ExisteArquivo(pSetupIni) Then
      sALIAS = ReadIniFile(pSetupIni, "Database Format", "ALIAS", sALIAS)
      nDBTIPO = ReadIniFile(pSetupIni, "Database Format", "DBTIPO", CStr(nDBTIPO))
      sDBVERSAO = ReadIniFile(pSetupIni, "Database Format", "DBVERSAO")
      bisADO = ReadIniFile(pSetupIni, "Database Format", "isADO", CStr(bisADO))
      bisODBC = ReadIniFile(pSetupIni, "Database Format", "isODBC", CStr(bisODBC))
      sSERVER = ReadIniFile(pSetupIni, "Database Format", "SERVER", sSERVER)
      sDBNAME = ReadIniFile(pSetupIni, "Database Format", "DBNAME", sDBNAME)
      sDSN = ReadIniFile(pSetupIni, "Database Format", "DSN", sDSN)
      sUID = ReadIniFile(pSetupIni, "Database Format", "UID", sUID)
      sPWD = Decrypt2(ReadIniFile(pSetupIni, "Database Format", "PWD", IIf(sUID = "USU_VERIF", Encrypt2(sPWD), "")))
   End If
   
   While ReadIniFile(pLocalReg, SecaoBase & Index, "ALIAS") <> ""
      Index = Index + 1
   Wend
   nLast = ReadIniFile(pLocalReg, "Conections", "Last", 0)
   nLast = IIf(nLast > Index, 0, nLast)
   Call WriteIniFile(pLocalReg, "Conections", "Last", CStr(nLast))
         
   If Trim(sALIAS) <> "" Then Call WriteIniFile(pLocalReg, sSecao, "ALIAS", sALIAS)
   Call WriteIniFile(pLocalReg, sSecao, "DBTIPO", CStr(nDBTIPO))
   'Call WriteIniFile(pLocalReg, sSecao, "DBVERSAO", sDBVERSAO)
   Call WriteIniFile(pLocalReg, sSecao, "isODBC", CStr(bisODBC))
   Call WriteIniFile(pLocalReg, sSecao, "isADO", CStr(bisADO))
   If Trim(sSERVER) <> "" Then Call WriteIniFile(pLocalReg, sSecao, "SERVER", sSERVER)
   Call WriteIniFile(pLocalReg, sSecao, "DBNAME", sDBNAME)
   'Call WriteIniFile(pLocalReg, sSecao, "DSN", sDSN)
   Call WriteIniFile(pLocalReg, sSecao, "UID", sUID)
   Call WriteIniFile(pLocalReg, sSecao, "PWD", Encrypt2(sPWD))
End Sub
Public Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, Optional DefaultValue As String) As String
   Dim strBuffer As String
   
   strBuffer = Space$(255)
   If GetPrivateProfileString(strSection, strKey, "", strBuffer, 255, strIniFile) > 0 Then
      If InStr(strBuffer, Chr$(0)) > 0 Then
        strBuffer = VBA.Left$(strBuffer, InStr(strBuffer, Chr$(0)) - 1)
      End If
      ReadIniFile = RTrim$(strBuffer)
   Else
      ReadIniFile = DefaultValue
   End If
End Function
Public Function RegServer(sServerPath As String, Optional fRegister = True, Optional fMsg As Boolean = True, Optional isActivexExe As Boolean = False) As Boolean
   Dim hMod       As Long    ' module handle
   Dim lpfn       As Long    ' reg/unreg function address
   Dim sCmd       As String  ' msgbox string
   Dim lpThreadID As Long    ' unused, receives the thread ID
   Dim hThread    As Long    ' thread handle
   Dim fSuccess   As Boolean ' if things worked
   Dim dwExitCode As Long    ' thread's exit code if it doesn't finish
   
   ' Load the server into memory
   hMod = LoadLibrary(sServerPath)
   
   ' Get the specified function's address and our msgbox string.
   If fRegister Then
      If isActivexExe Then
         lpfn = GetProcAddress(hMod, "ExeRegisterServer")
      Else
         lpfn = GetProcAddress(hMod, "DllRegisterServer")
      End If
     sCmd = "register"
   Else
      If isActivexExe Then
         lpfn = GetProcAddress(hMod, "ExeUnregisterServer")
      Else
         lpfn = GetProcAddress(hMod, "DllUnregisterServer")
      End If
      sCmd = "unregister"
   End If
   
   ' If we got a function address...
   If lpfn Then
     
     ' Create an alive thread and execute the function.
     hThread = CreateThread(ByVal 0, 0, ByVal lpfn, ByVal 0, 0, lpThreadID)
     
     ' If we got the thread handle...
     If hThread Then
       
       ' Wait 10 secs for the thread to finish (the function may take a while...)
       fSuccess = (WaitForSingleObject(hThread, 10000) = WAIT_OBJECT_0)
       
       ' If it didn't finish in 10 seconds...
       If Not fSuccess Then
         ' Something unlikely happened, lose the thread.
         Call GetExitCodeThread(hThread, dwExitCode)
         Call ExitThread(dwExitCode)
       End If
       
       ' Lose the thread handle
       Call CloseHandle(hThread)
     
     End If   ' hThread
   End If   ' lpfn
   
   ' Free the server if we loaded it.
   If hMod Then Call FreeLibrary(hMod)
   
   RegServer = fSuccess
   
   If fMsg Then
      If fSuccess Then
         MsgBox "Successfully " & sCmd & "ed " & sServerPath   ' past tense
      Else
        MsgBox "Failed To " & sCmd & " " & sServerPath, vbExclamation
      End If
   End If
End Function
Public Function Encrypt2(ByVal Password As String, Optional Key As String) As String
   Dim P    As String
   Dim b    As String
   Dim S    As String
   Dim i    As Integer
   Dim j    As Integer
   Dim A1   As Integer
   Dim A2   As Integer
   Dim A3   As Integer
   
   Dim sAux As String
   
   If Len(Trim(Password)) < 2 Then
      sAux = "PI"
   Else
      sAux = Right$(Password, 1) & Left$(Password, 1)
   End If
   Password = sAux & Password
   If Trim(Key) = "" Then Key = Encrypt2("231072150500", "DIO")
   j = 1
   For i = 1 To Len(Key)
     P = P & Asc(Mid$(Key, i, 1))
   Next
    
   For i = 1 To Len(Password)
     A1 = Asc(Mid$(P, j, 1))
     j = j + 1
     If j > Len(P) Then j = 1
     A2 = Asc(Mid$(Password, i, 1))
     A3 = A1 Xor A2
     b = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b
     S = S + b
   Next
   Encrypt2 = S
End Function
Public Function Decrypt2(ByVal Password As String, Optional Key As String) As String
   Dim P    As String
   Dim b    As String
   Dim S    As String
   Dim i    As Integer
   Dim j    As Integer
   Dim A1   As Integer
   Dim A2   As Integer
   Dim A3   As Integer
   
   If Trim(Key) = "" Then Key = Encrypt2("231072150500", "DIO")
   j = 1
   For i = 1 To Len(Key)
     P = P & Asc(Mid$(Key, i, 1))
   Next
   For i = 1 To Len(Password) Step 2
     A1 = Asc(Mid$(P, j, 1))
     j = j + 1
     If j > Len(P) Then j = 1
     b = Mid$(Password, i, 2)
     A3 = Val("&H" + b)
     A2 = A1 Xor A3
     S = S + Chr$(A2)
   Next
   Decrypt2 = Mid(S, 3)
End Function
Public Function WriteIniFile(ByVal strIniFile As String, strSection As String, strKey As String, strValue As String) As Boolean
   Dim intLen As Integer
   
   If Not ExisteArquivo(strIniFile) Then
      intLen = AbrirTxt(strIniFile)
      Call FecharTxt(intLen)
   End If
   intLen = 0
   intLen = WritePrivateProfileString(strSection, strKey, strValue, strIniFile)
   WriteIniFile = (intLen > 0)
End Function
Public Function ComputerName() As String
   Dim lpBuffer As String * 255
   Dim lRet As Long
   Dim i As Integer
   
   lRet = GetComputerName(lpBuffer, 255)
   If lRet <> 0 Then
     ComputerName = UCase$(Left$(Trim$(lpBuffer), Len(Trim$(lpBuffer)) - 1))
     ComputerName = EliminarString(ComputerName, Chr(0))
   Else
     ComputerName = "Não Fornecido."
   End If
End Function
Public Function EliminarString(ByVal Palavra$, ByVal Caracter$, Optional CaseSensitive = True) As String
    Dim Pos%, Com_Carac$
    
    If CaseSensitive Then
       Pos% = InStr(Palavra, Caracter)
    Else
      Pos% = InStr(UCase(Palavra), UCase(Caracter))
    End If
    
    Com_Carac = Palavra$
    While Pos% <> 0
        Com_Carac = Left$(Com_Carac, Pos% - 1) + Mid$(Com_Carac, Pos% + Len(Caracter$))
        If CaseSensitive Then
           Pos% = InStr(Com_Carac$, Caracter)
        Else
           Pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
        End If
    Wend

    If CaseSensitive Then
       Pos% = InStr(Com_Carac$, Caracter)
    Else
       Pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
    End If
    
    If Pos% > 0 Then
       If CaseSensitive Then
          Pos% = InStr(Pos% + Len(Caracter$), Com_Carac, Caracter)
       Else
          Pos% = InStr(Pos% + Len(Caracter$), UCase(Com_Carac), UCase(Caracter))
       End If
       If Pos% > 0 Then Com_Carac = Left$(Com_Carac, Pos% - 1)
    End If

    EliminarString = Com_Carac
End Function
Private Function CreatePathTempExe(ByRef pPathTemp As String, ByVal pArq As String, ByVal pVersaoInst As Double) As Boolean
   On Error GoTo TrataErro
   
   If Dir(pPathTemp, vbDirectory) = "" Then
      Call MkDir(pPathTemp)
   End If
   
   pPathTemp = pPathTemp & Mid(pArq, 1, IIf(InStr(pArq, ".") = 0, Len(pArq), InStr(pArq, ".") - 1))
   pPathTemp = ResolvePathName(pPathTemp)
   If Dir(pPathTemp, vbDirectory) = "" Then
      Call MkDir(pPathTemp)
   End If
   
   pPathTemp = ResolvePathName(pPathTemp & pVersaoInst)
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
   sPath = ResolvePathName(sPath & Mid(pArq, 1, IIf(InStr(pArq, ".") = 0, Len(pArq), InStr(pArq, ".") - 1)))
   sPath = ResolvePathName(sPath & pVersaoInst)
   
   If Dir(sPath, vbDirectory) <> "" Then
      While Dir(sPath, vbArchive) <> ""
         Call Del(sPath & Dir(sPath, vbArchive), False)
      Wend
      Call RmDir(sPath)
   End If
      
   sPath = pPathTemp
   sPath = ResolvePathName(sPath & Mid(pArq, 1, IIf(InStr(pArq, ".") = 0, Len(pArq), InStr(pArq, ".") - 1)))
   If Dir(sPath, vbDirectory) <> "" Then
      While Dir(sPath, vbArchive) <> ""
         Call Del(sPath & Dir(sPath, vbArchive), False)
      Wend
      Call RmDir(sPath)
   End If
   
   sPath = pPathTemp
   If Dir(sPath, vbDirectory) <> "" Then
      While Dir(sPath, vbArchive) <> ""
         Call Del(sPath & Dir(sPath, vbArchive), False)
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
      sArqLog = ResolvePathName(PathArquivoLog) & "AutoInstall.Log"
   Else
      sArqLog = PathArquivoLog
   End If
   
End Sub
Public Function AtualizaDLL(pDLLMaq As String, pDLLNova As String, Optional pObrigatorio As Boolean = False, Optional ArquivoLog As String) As Boolean
   Dim nVersaoMaq    As Double
   Dim nVersaoNova   As Double
   Dim sVersaoMaq    As String
   Dim sVersaoNova   As String
   Dim sPathMaq      As String
   
   On Error Resume Next
   
   nVersaoNova = 0
   nVersaoMaq = 0
   sVersaoNova = "0"
   sVersaoMaq = "0"
   
   sPathMaq = StrReverse(Mid(StrReverse(pDLLMaq), InStr(StrReverse(pDLLMaq), "\") + 1))
   If ExisteArquivo(pDLLNova) Then
      nVersaoNova = GetFileVersionNumber(pDLLNova)
      sVersaoNova = GetFileVersion(pDLLNova)
      
      If ExisteArquivo(pDLLMaq) Then
         nVersaoMaq = GetFileVersionNumber(pDLLMaq)
         sVersaoMaq = GetFileVersion(pDLLMaq)
      Else
         nVersaoMaq = -1
         sVersaoMaq = "-1"
      End If
   End If
   
   If nVersaoNova > nVersaoMaq Or pObrigatorio Then
      If ExisteArquivo(pDLLMaq) Then
         Call RegServer(pDLLMaq, False, False)
         Call Kill(pDLLMaq)
      Else
         Call CreatePath(sPathMaq)
      End If
      
      Call FileCopy(pDLLNova, pDLLMaq)
      AtualizaDLL = RegServer(pDLLMaq, True, False)
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
   If ExisteArquivo(pEXENovo) Then
      nVersaoNova = GetFileVersionNumber(pEXENovo)
      
      If ExisteArquivo(pEXEMaq) Then
         nVersaoMaq = GetFileVersionNumber(pEXEMaq)
      End If
   Else
      nVersaoNova = 0
      nVersaoMaq = 0
   End If
   If nVersaoNova > nVersaoMaq Then
      If pRegistro Then
         Call RegServer(pEXEMaq, False, False)
      End If
      Call Kill(GetNameFromPath(pEXEMaq, 1) & "*.rpt")
      
      Call Del(pEXEMaq)
      Call FileCopy(pEXENovo, pEXEMaq)
      
      AtualizaExe = True
      If pRegistro Then
         AtualizaExe = RegServer(pEXEMaq, True, False)
      End If
   End If
End Function
Public Function GetRegisterDir(ByVal sArq As String, Optional AppPath As String) As String
   Dim sPath      As String
   Dim bAchou     As Boolean
   Dim sAux       As String
   
   
   sPath = ResolvePathName(AppPath)
   bAchou = ExisteArquivo(AppPath & sArq)
   
   If Not bAchou Then
      sPath = ResolvePathName(GetWindowsSysDir())
      sPath = ResolvePathName(sPath & SubFolder)
      bAchou = ExisteArquivo(sPath & sArq)
   End If
   If Not bAchou Then
      sPath = ResolvePathName(GetWindowsSysDir())
      bAchou = ExisteArquivo(sPath & sArq)
   End If
   If Not bAchou Then
      sPath = ResolvePathName(GetSpecialFolder(CSIDL_PROGRAM_FILES))
      sPath = ResolvePathName(sPath & SubFolder)
      bAchou = ExisteArquivo(sPath & sArq)
      If Not bAchou Then
         sPath = ProcuraArquivo(sPath, sArq)
         bAchou = ExisteArquivo(sPath & sArq)
      End If
   End If
   If Not bAchou Then
      sPath = ResolvePathName(GetSpecialFolder(CSIDL_WINDOWS))
      bAchou = ExisteArquivo(sPath & sArq)
   End If
   If Not bAchou Then
      sPath = ResolvePathName(GetSpecialFolder(CSIDL_COMMON))
      bAchou = ExisteArquivo(sPath & sArq)
   End If
   If Not bAchou Then
      sPath = ResolvePathName(GetSpecialFolder(CSIDL_PROGRAM_FILES))
      bAchou = ExisteArquivo(sPath & sArq)
   End If
   
   sPath = IIf(bAchou, sPath, "")
   
   GetRegisterDir = ResolvePathName(sPath)
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
   
   sAux = ResolvePathName(pPath)
   ChDir sAux
   sPath0 = sAux
   
   bAchou = ExisteArquivo(sAux & pArq)
   If bAchou Then
      sAux = sPath
   Else
      sAux = Dir(sAux, vbDirectory)
      While sAux <> ""
         sAux = Dir(Attributes:=VbFileAttribute.vbDirectory)
         If InStr(sAux, ".") = 0 And sAux <> "" Then
            sAux = ResolvePathName(pPath & sAux)
            If (GetAttr(sAux) And vbDirectory) = vbDirectory Then
               bAchou = ExisteArquivo(sAux & pArq)
               
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
                           If sAux = ResolvePathName(pPath & Dir(Attributes:=vbDirectory)) Then
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
Public Function GetFileVersion(ByVal pFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
   Dim sVerInfo As VERINFO
   Dim strVer As String
   
   On Error GoTo GFVError
   
   If IsMissing(fIsRemoteServerSupportFile) Then
      fIsRemoteServerSupportFile = False
   End If
   
   '
   'Get the file version into a VERINFO struct, and then assemble a version string
   'from the appropriate elemen
   '
   If GetFileVerStruct(pFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
      strVer = ""
      strVer = strVer & Format$(sVerInfo.nMSHi, "000") & "."
      strVer = strVer & Format$(sVerInfo.nMSLo, "000") & "."
      strVer = strVer & Format$(sVerInfo.nLSHi, "000") & "."
      strVer = strVer & Format$(sVerInfo.nLSLo, "000")
      GetFileVersion = strVer
   Else
      GetFileVersion = ""
   End If
   
   Exit Function
    
GFVError:
   GetFileVersion = ""
   If Err = 48 Then
      MsgBox "ERRO : " & Err & " - " & Error
   End If
   Err = 0
End Function
Public Function GetFileVersionNumber(pFilename As String) As Double
   Dim Pos  As Integer
   Dim nVer As Double
   Dim sVer As String
   Dim sAux As String
   Dim PosA As Integer
   
   On Error Resume Next
   
   sAux = ""
   PosA = 0
   
   sVer = GetFileVersion(pFilename)
   Pos = InStr(sVer, ".")
   If Pos <> 0 Then
      While Pos <> 0
         Pos = InStr(PosA + 1, sVer, ".")
         sAux = sAux & Right("000" + Mid(sVer, PosA + 1, IIf(Pos = 0, Len(sVer) + 1, Pos) - PosA - 1), 3)
         PosA = Pos
      Wend
   End If
   sAux = IIf(Trim(sAux) = "", "0", Trim(sAux))
   GetFileVersionNumber = Val(sAux)
   
'   Pos = InStr(sVer, ".")
'   While Pos <> 0
'      sVer = Mid(sVer, 1, Pos - 1) & Mid(sVer, Pos + 1)
'      Pos = InStr(sVer, ".")
'   Wend
'   sVer = IIf(sVer = "", sVer = "0", sVer)
'   GetFileVersionNumber = Val(sVer)
End Function
'********************************************************************************************
'**************************************** PRIVATE *******************************************
'********************************************************************************************
Private Function ResolvePathName(ByVal sPath As String, Optional bDebug As Boolean) As String
   Dim PosIni As Integer
   Dim PosFim As Integer
   Dim sMsg   As String
   
   If Right(sPath, 1) <> "\" And Trim(sPath) <> "" Then
      sPath = sPath & "\"
   End If
   If InStr(sPath, "%") <> 0 Then
      PosIni = InStr(sPath, "%")
      PosFim = InStr(PosIni + 1, sPath, "%")
      
      If bDebug Then
         sMsg = "ResolvePathName(sPath)" & vbNewLine
         sMsg = sMsg & "Inicio : " & Mid(sPath, 1, PosIni - 1) & vbNewLine
         sMsg = sMsg & "Meio   : " & Mid(sPath, PosIni + 1, PosFim - PosIni - 1) & vbNewLine
         sMsg = sMsg & "Fim : " & Mid(sPath, PosFim + 1) & vbNewLine
         sMsg = sMsg & "Environ : " & Environ(Mid(sPath, PosIni + 1, PosFim - PosIni - 1)) & vbNewLine
         MsgBox sMsg
      End If
      sPath = Mid(sPath, 1, PosIni - 1) & Environ(Mid(sPath, PosIni + 1, PosFim - PosIni - 1)) & Mid(sPath, PosFim + 1)
   End If
   
   ResolvePathName = sPath
End Function
Private Function GetSpecialFolder(CSIDL As Long) As String
    Dim sPath  As String
    Dim IDL    As ITEMIDLIST
    Dim nhWnd  As Long
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
   If CSIDL = CSIDL_TEMPORARY Then
      GetSpecialFolder = ResolvePathName(GetTempFolder)
   Else
      GetSpecialFolder = ""
      If SHGetSpecialFolderLocation(nhWnd, CSIDL, IDL) = 0 Then
         '
         ' Get the path from the ID list, and return the folder.
         '
         sPath = Space$(260)
         If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            GetSpecialFolder = ResolvePathName(Left$(sPath, InStr(sPath, vbNullChar) - 1) & "")
         End If
     End If
   End If
End Function
Private Function GetTempFolder() As String
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
      GetTempFolder = ResolvePathName(VBA.Left(strTempPath, InStr(1, strTempPath, vbNullChar) - 1))
   Else
      GetTempFolder = ""
   End If
End Function
Private Function ExisteArquivo(ByVal strPathName As String) As Boolean
   Dim intFileNum   As Integer
   Dim sArq         As String
   Dim sPath        As String
   
   On Error Resume Next
   
   strPathName = Trim(strPathName)
   strPathName = Replace(strPathName, """", "")
   
   Call GetNameFromPath(strPathName, sPath)
   sArq = Mid(strPathName, Len(sPath) + 1)
   If Len(Dir(strPathName, vbArchive)) > 4 And sPath <> "" And sArq <> "" Then
      ExisteArquivo = True
   Else
      If Right$(strPathName, 1) = "\" Then
          strPathName = VBA.Left$(strPathName, Len(strPathName) - 1)
      End If
      '
      'Attempt to open the file, return value of this function is False
      'if an error occurs on open, True otherwise
      '
      intFileNum = FreeFile
      Open strPathName For Input As intFileNum
      ExisteArquivo = IIf(Err = 0, True, False)
      Close intFileNum
   End If
   
   Err = 0
End Function
Private Function GetNameFromPath(PathFile As String, Optional ByRef PathReturn As String) As String
   Dim i As Integer
   
   i = InStrRev(PathFile, "\")
   i = IIf(i = 0, 1, i)
   If PathReturn = "1" Then
      GetNameFromPath = VBA.Left$(PathFile, i)
   Else
      GetNameFromPath = VBA.Mid$(PathFile, Len(VBA.Left$(PathFile, i)) + 1)
   End If
   PathReturn = ResolvePathName(VBA.Left$(PathFile, i))
End Function
Private Sub Wait(Second As Integer)
   Dim TimeIni&
   TimeIni& = Timer
   While Timer - TimeIni& < Second
'      DoEvents
   Wend
End Sub
Private Sub Del(File As String, Optional ViewError As Boolean = True)
   If ExisteArquivo(File) Then
      On Error GoTo Fim
      Call Kill(File)
   End If
   Exit Sub
Fim:
   If ViewError Then
      MsgBox "O seguinte erro ocorreu : " & vbNewLine & vbNewLine & _
            "Number : " & Err.Number & vbNewLine & _
            "Description : " & Err.Description & _
            "Help File : " & Err.HelpFile
   End If
End Sub
Private Function AbrirTxt%(Arq$)
     Dim Hnd%
  On Error GoTo CopyErr
   Call Del(Arq$)
   AbrirTxt% = FreeFile()
   Open Arq$ For Output As #AbrirTxt%
Exit Function
CopyErr:
  Select Case Err
     Case 55: Err = 0
     Case Else: MsgBox Err.Number & " - " & Err.Description
  End Select
End Function
Private Sub FecharTxt(Arq%)
      Close #Arq%
End Sub
Private Sub CreatePath(pPath As String)
   Dim sAux    As String
   Dim sPath   As String
   
   pPath = ResolvePathName(pPath)
   
   If InStr(pPath, ":\") > 0 Then
      sPath = Mid(pPath, 1, InStr(pPath, ":\") + 1)
      ChDir sPath
      If Trim(Dir(pPath, vbDirectory)) = "" Then
         sAux = Mid(pPath, InStr(pPath, "\") + 1)
         While InStr(sAux, "\") <> 0
            sPath = Mid(sAux, 1, InStr(sAux, "\") - 1)
            If Trim(Dir(sPath, vbDirectory)) = "" Then
               Call MkDir(sPath)
            End If
            ChDir sPath
            sAux = Mid(sAux, InStr(sAux, "\") + 1)
         Wend
      End If
   End If
End Sub
Private Function GetTag(ByRef Controle As Variant, ByVal VarName As String, Optional VarDefault As String) As String
   Dim PosIni  As Long
   Dim PosFim  As Long
   Dim StrTAG  As String
   Dim i       As Integer
   
   On Error GoTo Saida
   
   VarName = "|" & Trim(VarName) & "="
   
   If UCase(TypeName(Controle)) = "STRING" Then
      StrTAG = Controle
   Else
      StrTAG = Controle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(VarName))
   If PosIni > 0 Then
      PosIni = PosIni + Len(Trim(VarName))
      PosFim = InStr(PosIni, StrTAG$, "|")
      i = 0
      While Mid(StrTAG$, PosIni + i, 1) = "|"
         i = i + 1
      Wend
      If i > 0 Then
         PosFim = InStr(PosIni + (i - 1), StrTAG$, "|")
      End If
      PosFim = IIf(PosFim = 0, Len(StrTAG$), PosFim - 1)
      StrTAG$ = Mid(StrTAG$, PosIni, PosFim - PosIni + 1)
   Else
      StrTAG$ = ""
   End If
   GetTag = StrTAG$
Saida:
   If StrTAG$ = "" Then
      GetTag = VarDefault
   End If
End Function
Private Function NumericPassword(ByVal Password As String) As Long
    Dim value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    str_len = Len(Password)
    For i = 1 To str_len
        ' Adiciona a próxima letra
        ch = Asc(Mid$(Password, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)

        ' Change the shift offse
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = value
End Function

