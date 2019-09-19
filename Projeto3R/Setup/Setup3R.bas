Attribute VB_Name = "SetupBas"
Option Explicit
Global xAmbiente  As Object 'XLib.Ambiente

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long '(lpThreadAttributes As SECURITY_ATTRIBUTES,
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFilename As String, lVerHandle As Long) As Long
Public Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFilename As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
Public Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Sub lmemcpy Lib "VB5STKIT.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)

Type VERINFO                  'Version FIXEDFILEINFO
    strPad1 As Long           'Pad out struct version
    strPad2 As Long           'Pad out struct signature
    nMSLo As Integer          'Low word of ver # MS DWord
    nMSHi As Integer          'High word of ver # MS DWord
    nLSLo As Integer          'Low word of ver # LS DWord
    nLSHi As Integer          'High word of ver # LS DWord
    strPad3(1 To 16) As Byte  'Skip some of VERINFO struct (16 bytes)
    FileOS As Long            'Information about the OS this file is targeted for.
    strPad4(1 To 16) As Byte  'Pad out the resto of VERINFO struct (16 bytes)
End Type

Public Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
Public Const STATUS_WAIT_0 = &H0
Public Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)             ' The State of the specified object is signaled (success)
Public Const CONNECT_LAN As Long = &H2
Public Const CONNECT_MODEM As Long = &H1
Public Const CONNECT_PROXY As Long = &H4
Public Const CONNECT_OFFLINE As Long = &H20
Public Const CONNECT_CONFIGURED As Long = &H40

Public Const PROCESS_QUERY_INFORMATION = &H400&
Public Const STATUS_PENDING = &H103
Public Const STILL_ACTIVE = STATUS_PENDING

Global gLocalPath As String
Global gLocalPathSetup As String
Global gFileTotal  As String
Global gDebug     As Boolean
Global gMetodo1   As Boolean
Global gFileMSDE  As String
Global gFtpURL    As String
Global gFtpUSU    As String
Global gFtpPWD    As String

Public Function ApagarDiretorio(pDir As String) As Boolean
   Dim sDir As String
   Dim sAux As String
   
   On Error Resume Next
   
   sDir = ResolvePathName(pDir)
   Call ExcluirArquivo(sDir & "*.*", False)
   sAux = Dir(sDir, Attributes:=VbFileAttribute.vbDirectory)
   While sAux <> ""
      If sAux <> "" And sAux <> "." And sAux <> ".." Then
         Call ApagarDiretorio(sDir & sAux & "\")
         sAux = Dir(sDir, Attributes:=VbFileAttribute.vbDirectory)
      End If
      sAux = Dir(Attributes:=VbFileAttribute.vbDirectory)
   Wend
   
   RmDir pDir
   
   ApagarDiretorio = (Dir(pDir) = "")
End Function
Public Sub Main()
   Dim MyVerif As Object
   Dim bExiste As Boolean
   Dim bBaixar As Boolean
   
   gDebug = (InStr(Command$, "Debug") <> 0)
   
   gFtpURL = "ftp.classeanet.com.br"
   gFtpUSU = "classean"
   gFtpPWD = "d0lphin72"
   
   
   gLocalPath = Environ("ProgramFiles") & "\ClasseA\Admin\Instalacao\"
   gLocalPathSetup = gLocalPath & "Setup\"
   gFileTotal = "Total.zip" '"MSDE.zip"
   gFileMSDE = gFileTotal
   gMetodo1 = ExisteArquivo(App.Path & "\" & gFileTotal) 'Or ExisteArquivo(gLocalPathSetup & gFileTotal)
   
   If gDebug Then MsgBox "Extrair Custom"
   
   Call RegistrarDependencia
       
   If Not ExisteArquivo(gLocalPathSetup & gFileTotal) Then
      If ExisteArquivo(gLocalPath & gFileTotal) Then
         Call CopiarArquivo(gLocalPath & gFileTotal, gLocalPathSetup & gFileTotal)
      ElseIf ExisteArquivo(App.Path & "\" & gFileTotal) Then
         Call CopiarArquivo(App.Path & "\" & gFileTotal, gLocalPathSetup & gFileTotal)
      End If
   End If
   
   If gMetodo1 Then
      bExiste = ExisteArquivo(gLocalPathSetup & gFileTotal)
      If bExiste Then
         bBaixar = (vbYes = MsgBox("Arquivo para instalação já se encontra em seu computador." & vbNewLine & vbNewLine & "Deseja baixar novamente?", vbQuestion + vbYesNo + vbDefaultButton2, "Setup 3R"))
         If bBaixar Then
            Call ExcluirArquivo(gLocalPath & gFileTotal)
            Call CopiarArquivo(gLocalPathSetup & gFileTotal, gLocalPath & gFileTotal)
            Call ExcluirArquivo(gLocalPathSetup & gFileTotal)
         End If
      Else
      
      End If
      bExiste = ExisteArquivo(gLocalPathSetup & gFileTotal)
      If Not bExiste Then
         '* Veficar existência de Conexão
         If Not IsWebConnected Then
            MsgBox "Você precisa estar conecatado à intenet para continuar a instalação.", vbInformation + vbOKOnly, "Atenção!"
            GoTo Saida
         End If
         If gDebug Then MsgBox "Conexão Ativa"
      
         Set MyVerif = CreateObject("VersaoFTP.TL_VerifVersao")
         With MyVerif
            
            'If .ConectarFTP(gFTPURL, "pub", "free4", False) Then
            If .ConectarFTP(gFtpURL, gFtpUSU, gFtpPWD, False) Then
            
               If App.Path = "C:\Sistemas\Dsr\Projeto3R\Setup" Then
                  If vbYes = MsgBox("Baixar Arquivos?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção!") Then
                     'MsgBox "Baixararquivo"
                     If .BaixarArquivo(pLocalOrig:="/anon_ftp/pub/", _
                                       pArqOrig:="P3R.zip", _
                                       pLocalDest:=Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup", _
                                       pArqDest:="P3R.zip", _
                                       pReWrite:=True, _
                                       bViewFlood:=True) Then
                        bExiste = bExiste
                     Else
                        bExiste = bExiste
                     End If
                  End If
               Else
                  Call .BaixarArquivo(pLocalOrig:="/anon_ftp/pub/", _
                                       pArqOrig:="P3R.zip", _
                                       pLocalDest:=Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup", _
                                       pArqDest:="P3R.zip", _
                                       pReWrite:=True, _
                                       bViewFlood:=True)
   
               End If
            End If
         End With
         Set MyVerif = Nothing
      End If
   End If
   On Error Resume Next
'   If ExisteArquivo(gLocalPathSetup & gFileTotal) Then
'      Call ExcluirArquivo(gLocalPathSetup & sAux)
'   End If
   FrmDownload.Show 'vbModal
'   Else
'      MsgBox "Erro ao baixar arquivo de instalação." & vbNewLine & "Por favor tente novamente.", vbInformation
'   End If
   
Saida:
   On Error Resume Next
   Set MyVerif = Nothing
   Call ExcluirArquivo(App.Path & "\Zip32.dll")
   Call ExcluirArquivo(App.Path & "\DEP.zip")
   Call ExcluirArquivo(App.Path & "\Unzip32.dll")
End Sub
Public Sub RegistrarDependencia()
   Call ExtractResData("UNZIP", "CUSTOM", App.Path & "\Unzip32.dll")
   Call ExtractResData("DEP", "CUSTOM", App.Path & "\DEP.zip")
   
   If ExisteArquivo(App.Path & "\Unzip32.dll") Then
      Call RegSetupFile(App.Path, "\Unzip32.dll")
      If ExisteArquivo(App.Path & "\DEP.ZIP") Then
         
         If gDebug Then MsgBox "Descompactar DEP.zip em: " & gLocalPathSetup
         Call ExcluirArquivo(App.Path & "\DEP\*.*")
         Call DescompactarArquivo(App.Path, "DEP.ZIP", gLocalPathSetup)
         Call ExcluirArquivo(App.Path & "\DEP.zip")
         
         If gDebug Then MsgBox "Registrar Arquivos"
         Call RegistrarArquivos
      End If
   End If
End Sub
Private Sub RegistrarArquivos()
   Dim sWinDir As String
   Dim sClaDir As String
   Dim sCodeDir As String
   Dim sCommDir As String
   
   sWinDir = Environ("SystemRoot") & "\System32\"
   sClaDir = sWinDir & "ClasseA\"
   sCodeDir = Environ("ProgramFiles") & "\Codejock Software\ActiveX\Xtreme SuitePro ActiveX v11.2.2\Bin\"
   sCommDir = Environ("ProgramFiles") & "\Common Files\Business Objects\3.0\Bin\"
   
   On Error Resume Next
   Call RegSetupFile(sClaDir, "xLib.dll")
   Call RegSetupFile(sWinDir, "msvbvm50.dll")
   Call RegSetupFile(sWinDir, "msvbvm60.dll")
   Call RegSetupFile(sWinDir, "asycfilt.dll")
   Call RegSetupFile(sWinDir, "comcat.dll")
   Call RegSetupFile(sWinDir, "oleaut32.dll")
   Call RegSetupFile(sWinDir, "olepro32.dll")
   Call RegSetupFile(sWinDir, "stdole2.tlb")
   Call RegSetupFile(sWinDir, "version.dll")
   Call RegSetupFile(sClaDir, "VersaoFTP.dll")
   Call RegSetupFile(sCodeDir, "Codejock.Controls.v11.2.2.ocx")
   Call RegSetupFile(sWinDir, "MSINET.OCX")
   Call RegSetupFile(sClaDir, "Zip32.dll")
   Call RegSetupFile(sClaDir, "Unzip32.dll")
   
   On Error Resume Next
   Call RegSetupFile(sWinDir, "smtpctrs.dll")
   Call RegSetupFile(sCommDir, "RegistryWrapper.dll")
End Sub
Private Sub RegSetupFile(sPath As String, sFile As String)
   Dim bCopia As Boolean
   
   If ExisteArquivo(gLocalPathSetup & sFile) Then
      bCopia = True
      If ExisteArquivo(sPath & sFile) Then
         bCopia = GetFileVersion(gLocalPathSetup & sFile) > GetFileVersion(sPath & sFile)
      End If
   End If
   
   If bCopia Then
      Call CopiarArquivo(gLocalPathSetup & sFile, sPath & sFile)
   End If
   
   If ExisteArquivo(sPath & sFile) Then
      Call RegServer(sPath & sFile, True, False)
   End If
   
   Call ExcluirArquivo(gLocalPathSetup & sFile)
End Sub
Public Function DescompactarArquivo(pPath As String, pFile As String, Optional pPathDest As String, Optional pHonorDir As Boolean = True) As Boolean
   Dim pMessage  As String
   
   On Error GoTo TrataErro
   Dim oUnZip As CGUnzipFiles
   
   'Call RegServer(App.Path & "\Unzip.dll")
        
   If Trim(pPathDest) = "" Then pPathDest = pPath
   Call CriarDiretorio(pPathDest)
   
    Set oUnZip = New CGUnzipFiles
    With oUnZip
      .ZipFileName = ResolvePathName(pPath) & pFile
      .ExtractDir = ResolvePathName(pPathDest) 'GetTempPathName
      .HonorDirectories = pHonorDir
      If .Unzip <> 0 Then
         pMessage = .GetLastMessage
         DescompactarArquivo = True
      End If
    End With
    Set oUnZip = Nothing
    
    'MsgBox "\ZIPTEST.ZIP Extracted Successfully to " & GetTempPathName

    Exit Function

TrataErro:
    MsgBox Err.Number & " " & "Form1::cmdUnZip_Click" & " " & Err.Description
End Function
Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
    Dim dwflags As Long
    Dim WebTest As Boolean
    ConnType = ""
    WebTest = InternetGetConnectedState(dwflags, 0&)
    Select Case WebTest
        Case dwflags And CONNECT_LAN: ConnType = "LAN"
        Case dwflags And CONNECT_MODEM: ConnType = "Modem"
        Case dwflags And CONNECT_PROXY: ConnType = "Proxy"
        Case dwflags And CONNECT_OFFLINE: ConnType = "Offline"
        Case dwflags And CONNECT_CONFIGURED: ConnType = "Configurada"
        'Case dwflags And CONNECT_RAS: ConnType = "Remota"
    End Select
   IsWebConnected = WebTest
End Function
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
Public Sub SincShell(Comando As String, Optional Modo As VbAppWinStyle = vbMaximizedFocus, Optional EsperaProcesso = True)
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Call xAmbiente.SincShell(Comando, Modo, EsperaProcesso)
End Sub
Public Sub L_SincShell(Comando As String, Optional Modo As VbAppWinStyle = vbMaximizedFocus, Optional EsperaProcesso = True)
   Dim IDProcess  As Long
   Dim hProcess   As Long
   Dim ExitCode   As Long
   Dim Ret        As Long
   
   On Error GoTo TrataErro

   IDProcess = Shell(Comando, Modo)
   If EsperaProcesso Then
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 1, IDProcess)
      Do
         Ret = GetExitCodeProcess(hProcess, ExitCode)
         DoEvents
      Loop While (ExitCode = STILL_ACTIVE)
      Ret = CloseHandle(hProcess)
   End If
Exit Sub
TrataErro:
   MsgBox CStr(Err) & " - " & CStr(Error)
End Sub
Public Function GetNameFromPath(PathFile As String, Optional ByRef PathReturn As String) As String
   Dim i As Integer
   
   i = InStrRev(PathFile, "\")
   'i = IIf(i = 0, 1, i)
   If PathReturn = "1" Then
      GetNameFromPath = VBA.Left$(PathFile, i)
   Else
      GetNameFromPath = VBA.Mid$(PathFile, Len(VBA.Left$(PathFile, i)) + 1)
   End If
   PathReturn = ResolvePathName(VBA.Left$(PathFile, i))
End Function
Public Function CriarDiretorio(pPath As String, Optional bViewMsg As Boolean = False) As Boolean
   Dim sAux     As String
   Dim sPath    As String
   Dim sOldPath As String
   Dim sIniPath As String
   Dim PosIni   As Integer
   Dim PosAux   As Integer
   Dim nResp    As Integer
   Dim bResult  As Boolean
   
   On Error Resume Next
   
   sOldPath = CurDir$
   
   pPath = ResolvePathName(pPath)
   bResult = (Trim(Dir(pPath, vbDirectory)) <> "")
   
   If Not bResult Then
      PosIni = InStr(1, pPath, "\")
      While PosIni <> 0
         PosIni = InStr(PosIni + 1, pPath, "\")
         If PosIni > 0 Then
            sPath = Left$(pPath, PosIni - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir sPath
            If Err Then
               ' We must create this directory
               Err = 0
               #If Win32 And LOGGING Then
                  NewAction gstrKEY_CREATEDIR, """" & sPath & """"
               #End If
               MkDir sPath
               #If Win32 And LOGGING Then
                  If Err Then
                     LogError ResolveResString(resMAKEDIR) & " " & sPath
                     AbortAction
                     GoTo Done
                  Else
                     CommitAction
                  End If
               #End If
            End If
         End If
      Wend
      bResult = (Trim(Dir(pPath, vbDirectory)) <> "")
   End If
   
   ChDir sOldPath
   
   If Trim(Dir(pPath, vbDirectory)) = "" Then
      nResp = MsgBox("Erro ao criar pasta:" & pPath, vbRetryCancel Or vbExclamation Or vbDefaultButton2, "Ambiente")
      Select Case nResp
         Case vbIgnore, vbAbort, vbCancel
            bResult = False
         Case vbRetry
            bResult = CriarDiretorio(pPath, bViewMsg)
      End Select
   End If
   CriarDiretorio = bResult
   Err = 0
End Function
Public Function ExcluirArquivo(File As String, Optional ViewError As Boolean = True) As Boolean
   If ExisteArquivo(File) Then
      On Error GoTo Fim
      Call Kill(File)
   End If
   ExcluirArquivo = Not ExisteArquivo(File)
   Exit Function
Fim:
   If ViewError Then
      'ClsMensagem.ExibirErro
   End If
End Function
Public Function ExisteArquivo(ByVal strPathName As String) As Boolean
   Dim intFileNum   As Integer
   Dim sArq         As String
   Dim sPath        As String
   
   On Error Resume Next
   
   strPathName = Trim(strPathName)
   strPathName = Replace(strPathName, """", "")
   
   Call GetNameFromPath(strPathName, sPath)
   sArq = Mid(strPathName, Len(sPath) + 1)
   If Len(Dir(strPathName, vbArchive)) > 4 And sPath <> "" And sArq <> "" Then
      ExisteArquivo = IIf(Err = 0, True, False)
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
Public Function GetFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
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
        If UCase(GetFileExtension(strFilename)) = "DEP" Then
            GetFileVerStruct = GetDepFileVerStruct(strFilename, sVerInfo)
        End If
    End If
End Function
Public Function GetFileExtension(ByVal pFilename As String) As String
    Dim nPos As Integer

    nPos = InStrRev(pFilename, ".")
    If nPos > 0 Then
      GetFileExtension = Mid(pFilename, nPos + 1)
    End If
End Function

Public Function ResolvePathName(ByVal sPath As String) As String
   Dim PosIni As Integer
   Dim PosFim As Integer

   If Right(sPath, 1) <> "\" And Trim(sPath) <> "" Then
      sPath = sPath & "\"
   End If
   If InStr(sPath, "%") <> 0 Then
      PosIni = InStr(sPath, "%")
      PosFim = InStr(PosIni + 1, sPath, "%")
      sPath = Mid(sPath, 1, PosIni - 1) & Environ(Mid(sPath, PosIni + 1, PosFim - PosIni - 1)) & Mid(sPath, PosFim + 1)
   End If

   ResolvePathName = sPath
End Function

Function GetRemoteSupportFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = 32767
    
    On Error GoTo Failed
    
    iFile = FreeFile

    Open strFilename For Input Access Read Lock Read Write As #iFile
    
    ' Loop through each line, looking for the key
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If VBA.Left$(strLine, cchVersionKey) = strVersionKey Then
            ' We've found the version key.  Copy everything after the equals sign
            Dim strVersion As String
            
            strVersion = Mid$(strLine, cchVersionKey + 1)
            
            'Parse and store the version information
            PackVerInfo strVersion, sVerInfo

            'Convert the format 1.2.3 from the .VBR into
            '1.2.0.3, which is really want we want
            sVerInfo.nLSLo = sVerInfo.nLSHi
            sVerInfo.nLSHi = 0
            
            GetRemoteSupportFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend
    
    Close iFile
    Exit Function

Failed:
    GetRemoteSupportFileVerStruct = False
End Function
Function GetDepFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    GetDepFileVerStruct = False
    
    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = 32767
    
    On Error GoTo Failed
    
    iFile = FreeFile

    Open strFilename For Input Access Read Lock Read Write As #iFile
    
    ' Loop through each line, looking for the key
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If VBA.Left$(strLine, cchVersionKey) = strVersionKey Then
            ' We've found the version key.  Copy everything after the equals sign
            Dim strVersion As String
            
            strVersion = Mid$(strLine, cchVersionKey + 1)
            
            'Parse and store the version information
            PackVerInfo strVersion, sVerInfo

            GetDepFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend
    
    Close iFile
    Exit Function

Failed:
    GetDepFileVerStruct = False
End Function

Sub PackVerInfo(ByVal strVersion As String, sVerInfo As VERINFO)
    Dim intOffset As Integer
    Dim intAnchor As Integer

    On Error GoTo PVIError

    intOffset = InStr(strVersion, ".")
    If intOffset = 0 Then
        sVerInfo.nMSHi = Val(strVersion)
        GoTo PVIMSLo
    Else
        sVerInfo.nMSHi = Val(VBA.Left$(strVersion, intOffset - 1))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, ".")
    If intOffset = 0 Then
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSHi
    Else
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, ".")
    If intOffset = 0 Then
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSLo
    Else
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, ".")
    If intOffset = 0 Then
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor))
    Else
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
    End If

    Exit Sub

PVIError:
    sVerInfo.nMSHi = 0
PVIMSLo:
    sVerInfo.nMSLo = 0
PVILSHi:
    sVerInfo.nLSHi = 0
PVILSLo:
    sVerInfo.nLSLo = 0
End Sub
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
       
       ' L_Wait 10 secs for the thread to finish (the function may take a while...)
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
Public Function CopiarArquivo(Orig As String, Dest As String) As Boolean
   Dim nMsg       As String
   Dim nTipoBox   As Long
   Dim Resp       As Integer

   On Error Resume Next
   
   If ExisteArquivo(Orig) Then
      If ExisteArquivo(Dest) Then
         Call Kill(Dest)
      Else
         Call CriarDiretorio(GetNameFromPath(Dest, 1))
      End If
      FileCopy Orig, Dest

   Else
      Call MsgBox("Arquivo não encontrado: " + UCase(Orig), "Importação")
      Resp = vbCancel
      Exit Function
   End If
   
   
   Resp = vbYes
   Select Case Err
      Case 71
         While Resp = vbYes
            nTipoBox = vbYesNo + vbCritical + vbDefaultButton1
            nMsg = "Drive ou arquivo inválido" + vbNewLine + vbNewLine
            nMsg = nMsg & "Insira um disco no drive ou verifieu o arquivo." + vbNewLine
            nMsg = nMsg & "Deseja continuar?"
            Resp = MsgBox(nMsg, nTipoBox, "Erro!")
            If Resp = vbYes Then
               Err = 0
               FileCopy Orig, Dest
            End If
         Wend
      Case 70
         While Resp = vbOK
            nTipoBox = vbOK + vbCritical + vbDefaultButton1
            nMsg = "Usuário não tem permissão a esta operação." + vbNewLine + vbNewLine
            nMsg = nMsg & "Algum recurso está compartilhando esta informação." + vbNewLine
            Resp = MsgBox(nTipoBox, nTipoBox, "Erro!")
            If Resp = vbYes Then
               Err = 0
               FileCopy Orig, Dest
            End If
         Wend
   End Select
   CopiarArquivo = (Resp = vbYes)
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
Public Function AbrirTxt(Arq As String) As Integer
   Dim Hnd As Integer
  
   On Error GoTo CopyErr
   Call ExcluirArquivo(Arq)
   AbrirTxt = FreeFile()
   Open Arq For Output As #AbrirTxt
Exit Function
CopyErr:
  Select Case Err
     Case 55: Err = 0
     Case Else: MsgBox Err.Number & " - " & Err.Description
  End Select
End Function
Public Sub FecharTxt(Arq As Integer)
      Close #Arq
End Sub
Public Function ReadTextFile(strPath As String) As String
    On Error GoTo ErrTrap
    Dim intFileNumber As Integer
    
    If Dir(strPath) = "" Then Exit Function
    intFileNumber = FreeFile
    Open strPath For Input As #intFileNumber
    
    ReadTextFile = Input(LOF(intFileNumber), #intFileNumber)
ErrTrap:
    Close #intFileNumber
End Function
Public Function ProcuraArquivo(ByVal pPath As String, ByVal pArq As String, Optional pAdmin As Boolean = True) As String
   Dim sAux    As String
   Dim sPath   As String
   Dim bAchou  As Boolean
   Dim sPath0  As String
   Dim i       As Integer
   Dim nVezes  As Integer
   Dim bAdmin  As Boolean
   
   On Error GoTo TrataErro
   nVezes = 10000
   
   
   sAux = ResolvePathName(pPath)
   ChDir sAux
   sPath0 = sAux
   
   bAchou = ExisteArquivo(sAux & pArq)
   If bAchou Then
      sPath = pPath
   Else
      sAux = Dir(sAux, vbDirectory)
      While sAux <> ""
         sAux = Dir(Attributes:=VbFileAttribute.vbDirectory)
         bAdmin = IIf(UCase(sAux) = "ADMIN", pAdmin, True)
         
         If InStr(sAux, ".") = 0 And sAux <> "" And bAdmin Then
            If (GetAttr(pPath & sAux) And vbDirectory) = vbDirectory Then
               sAux = ResolvePathName(pPath & sAux)
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
   ElseIf Err = 76 Then   '*Path not Fouund
      Resume Next
   Else
      MsgBox Err.Number & " - " & Err.Description, vbCritical, "TrataErro AutoInstall"
      Resume Next
   End If
End Function
Public Function BaixarArquivo(pArquivo As String)
   Dim MyVerif As Object
   Dim bExiste As Boolean
   
   If ExisteArquivo(Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup\" & pArquivo) Then
      Exit Function
   End If
   
   Set MyVerif = CreateObject("VersaoFTP.TL_VerifVersao")
   With MyVerif
      
      'If .ConectarFTP(gFTPURL, "pub", "free4", False) Then
      'If .ConectarFTP(gFTPURL, "classeanet", "ramos10", False) Then
      
      If InStr(App.Path, "\Sistemas\Dsr\Projeto3R\Setup") <> 0 Then
         'If vbYes = MsgBox("Baixar Arquivo '" & pArquivo & "' ?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção!") Then
            If .ConectarFTP(gFtpURL, gFtpUSU, gFtpPWD, False) Then
            'MsgBox "Baixararquivo"
               If .BaixarArquivo(pLocalOrig:="/anon_ftp/pub/", _
                              pArqOrig:=pArquivo, _
                              pLocalDest:=Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup", _
                              pArqDest:=pArquivo, _
                              pReWrite:=True, _
                              bViewFlood:=True) Then
                  bExiste = bExiste
               Else
                  bExiste = bExiste
               End If
          '  End If
         End If
      Else
         If ExisteArquivo(Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup\" & Mid(pArquivo, 1, Len(pArquivo) - 3) & "zip") Then
            Call CopiarArquivo(Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup\" & Mid(pArquivo, 1, Len(pArquivo) - 3) & "zip", Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup\" & pArquivo)
         End If
         If Not ExisteArquivo(Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup\" & pArquivo) Then
            If .ConectarFTP(gFtpURL, gFtpUSU, gFtpPWD, False) Then
               Call .BaixarArquivo(pLocalOrig:="/anon_ftp/pub/", _
                                 pArqOrig:=pArquivo, _
                                 pLocalDest:=Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup", _
                                 pArqDest:=pArquivo, _
                                 pReWrite:=True, _
                                 bViewFlood:=True)
            End If
         End If
      End If
      'End If
   End With
   Set MyVerif = Nothing
   
End Function
