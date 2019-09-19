Attribute VB_Name = "AutoInstall"
Option Explicit
Public Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" (ByVal Flags&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Public Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFilename As String, lVerHandle As Long) As Long
Public Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFilename As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
Public Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long

Public Declare Sub lmemcpy Lib "VB5STKIT.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
Public Declare Function fNTWithShell Lib "VB5STKIT.DLL" () As Boolean
Public Declare Function GetWinPlatform Lib "VB5STKIT.DLL" () As Long
Public Declare Function DllAbortAction Lib "VB5STKIT.DLL" Alias "AbortAction" () As Long
Public Declare Function DllCommitAction Lib "VB5STKIT.DLL" Alias "CommitAction" () As Long
Public Declare Function DllNewAction Lib "VB5STKIT.DLL" Alias "NewAction" (ByVal lpszKey As String, ByVal lpszData As String) As Long
Public Declare Function DllDisableLogging Lib "VB5STKIT.DLL" Alias "DisableLogging" () As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetDriveType32 Lib "kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempFilename32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal strWhichDrive As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFilename As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetModuleHandle Lib "Kernel" (ByVal lpModuleName As String) As Integer
Public Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer



Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long
Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, lpbData As Any, ByVal cbData As Long) As Long

Type VERINFO                                            'Version FIXEDFILEINFO
    strPad1 As Long                                     'Pad out struct version
    strPad2 As Long                                     'Pad out struct signature
    nMSLo As Integer                                    'Low word of ver # MS DWord
    nMSHi As Integer                                    'High word of ver # MS DWord
    nLSLo As Integer                                    'Low word of ver # LS DWord
    nLSHi As Integer                                    'High word of ver # LS DWord
    strPad3(1 To 16) As Byte                            'Skip some of VERINFO struct (16 bytes)
    FileOS As Long                                      'Information about the OS this file is targeted for.
    strPad4(1 To 16) As Byte                            'Pad out the resto of VERINFO struct (16 bytes)
End Type

Global gstrDestDir As String                                'dest dir for application files
Global gstrWinDir As String                                 'windows directory
Global gstrWinSysDir As String                              'windows\system directory
Global gstrTitle As String                                  '"setup" name of app being installed
Public gfNoUserInput As Boolean                         ' True if either gfSMS or gfSilent is True

Global Const gstrEXT_DEP$ = "DEP"
Global Const gintMAX_SIZE% = 255                        'Maximum buffer size
Global Const gintMAX_PATH_LEN% = 260                    ' Maximum allowed path length including path, filename,
Global Const gstrSEP_DIR$ = "\"                         ' Directory separator character
Global Const gstrSEP_DRIVE$ = ":"                       ' Driver separater character, e.g., C:\
Global Const gstrSEP_DIRALT$ = "/"                      ' Alternate directory separator character
Global Const gstrSEP_EXT$ = "."                         ' Filename extension separator character
Global Const gstrDECIMAL$ = "."
Global Const gstrNULL$ = ""                             'Empty string
Global Const gintNOVERINFO% = 32767                     'flag indicating no version info
Global Const gstrQUOTE$ = """"
Global Const gstrCOLON$ = ":"
Global Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.
Global Const gstrKEY_REGKEY = "RegKey"
Global Const gstrKEY_REGVALUE = "RegValue"

'Setup information file macros
Global Const gstrAPPDEST$ = "$(AppPath)"
Global Const gstrWINDEST$ = "$(WinPath)"
Global Const gstrWINSYSDEST$ = "$(WinSysPath)"
Global Const gstrWINSYSDESTSYSFILE$ = "$(WinSysPathSysFile)"
Global Const gstrPROGRAMFILES$ = "$(ProgramFiles)"
Global Const gstrCOMMONFILES$ = "$(CommonFiles)"
Global Const gstrCOMMONFILESSYS$ = "$(CommonFilesSys)"
Global Const gstrDAODEST$ = "$(MSDAOPath)"
Global Const gstrDONOTINSTALL$ = "$(DoNotInstall)"

' Registry manipulation API's (32-bit)
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const ERROR_SUCCESS = 0&
Global Const ERROR_NO_MORE_ITEMS = 259&

Global Const resMICROSOFTSHARED = 463

Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4

Global Const intDRIVE_REMOVABLE% = 2                    'Constants for GetDriveType
Global Const intDRIVE_FIXED% = 3
Global Const intDRIVE_REMOTE% = 4
Global Const intDRIVE_CDROM% = 5
Global Const intDRIVE_RAMDISK% = 6

'VB5STKIT.DLL logging errors
Private Const LOGERR_SUCCESS = 0
Private Const LOGERR_INVALIDARGS = 1
Private Const LOGERR_OUTOFMEMORY = 2
Private Const LOGERR_EXCEEDEDCAPACITY = 3
Private Const LOGERR_WRITEERROR = 4
Private Const LOGERR_NOCURRENTACTION = 5
Private Const LOGERR_UNEXPECTED = 6
Private Const LOGERR_FILENOTFOUND = 7

'Logging error Severities
Private Const LogErrOK = 1 ' OK to continue upon this error
Private Const LogErrFatal = 2 ' Must terminate install upon this error


' Hkey cache (used for logging purposes)
Private Type HKEY_CACHE
    hKey As Long
    strHkey As String
End Type
Private hkeyCache() As HKEY_CACHE

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400&
Private Const STILL_ACTIVE = &H103&

'Global ClsLoad As DS_LOAD

'**********
'* RegServer
Public Const STATUS_WAIT_0 = &H0
Public Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long '(lpThreadAttributes As SECURITY_ATTRIBUTES,
Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Sub Main()
'   Dim bDebug As Boolean
'
'   On Error GoTo Saida
'
'   bDebug = ((InStr(UCase(Command$), "DEBUG") <> 0))
'
'   Screen.MousePointer = vbHourglass
'   Call AutoInstall(App)
'
'   Screen.MousePointer = vbHourglass
'   Set ClsLoad = New DS_LOAD
'   With ClsLoad
'      .GetAppDateStart
'      .AnoDsvm = App.Comments           '* Ano de Desenvolvimento da Versão1.0
'      .APLIC = App                      '* Onjeto de Aplicação
'      .Name = App.FileDescription
'      If bDebug Then
'         MsgBox "Inicio : Exibir DS_LOAD" & vbNewLine & "Função : Main"
'      End If
'      .Show
'      'Wait 1
'      .Name = App.ProductName
'
'      Sys.AppExeName = .EXEName         '* SS
'      Sys.AppName = .Name               '* Suprimentos
'      Sys.AppTitle = .Title             '* Sistema de Suprimentos
'      Sys.AppVer = .Versao              '* Versão 1.0
'      Sys.AppDate = .GetAppDateStart    '* 30/09/99273 - 10:16:03.
'      Sys.NomeEmpresa = .Empresa        '* Marítima Petróleo Engenharia LTDA.
'
'      Sys.AppName = App.ProductName
'
'      If Not .SetFormat Or .Ativa Then '* Testa se já existe uma cópia da aplicação rodando.
'         End                           '* Define formato Data e número.
'      End If
'   End With
'   Set ClsLoad = Nothing
'
'   If bDebug Then
'      MsgBox "Inicio : MDI.Show" & vbNewLine & "Função : Main"
'   End If
'
'   MDI.Show
'
'   Screen.MousePointer = vbDefault
'GoTo Fim
'Saida:
'   MsgBox Err & " - " & Error & vbNewLine & "Function : Main"
'Fim:
'End Sub
Public Sub SincShell(Comando As String, Optional EsperaProcesso = True)
   Dim IDProcess As Long
   Dim hProcess As Long
   Dim ExitCode As Long
   Dim ret As Long
   
   On Error GoTo TrataErro

   IDProcess = Shell(Comando$, 1)
   If EsperaProcesso Then
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 1, IDProcess)
      Do
         ret = GetExitCodeProcess(hProcess, ExitCode)
         DoEvents
      Loop While (ExitCode = STILL_ACTIVE)
      ret = CloseHandle(hProcess)
   End If
Exit Sub
TrataErro:
   MsgBox CStr(Err) & " - " & CStr(Error)
End Sub
Function GetFileVersion(ByVal strFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
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
    If Err = 48 Then
       MsgBox "ERRO : " & Err & " - " & Error
    End If
    Err = 0
End Function

Function GetFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
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
Function GetRemoteSupportFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = gintNOVERINFO
    
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
Function Extension(ByVal strFilename As String) As String
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
Function GetDepFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    GetDepFileVerStruct = False
    
    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = gintNOVERINFO
    
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

    intOffset = InStr(strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSHi = Val(strVersion)
        GoTo PVIMSLo
    Else
        sVerInfo.nMSHi = Val(VBA.Left$(strVersion, intOffset - 1))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSHi
    Else
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSLo
    Else
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
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
Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String) As String
    Dim strBuffer As String
    Dim intPos As Integer

    '
    'If successful read of .INI file, strip any trailing zero returned by the Windows API GetPrivateProfileString
    '
    strBuffer = Space$(gintMAX_SIZE)
    
    If GetPrivateProfileString(strSection, strKey, gstrNULL, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = gstrNULL
    End If
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
Function ResolveDestDir(ByVal strDestDir As String, Optional fAssumeDir As Variant) As String
    Const strMACROSTART$ = "$("
    Const strMACROEND$ = ")"

    Dim intPos As Integer
    Dim strResolved As String
    Dim hKey As Long
    Dim strPathsKey As String
    Dim fQuoted As Boolean
    
    If IsMissing(fAssumeDir) Then
        fAssumeDir = True
    End If
    
    strPathsKey = RegPathWinCurrentVersion()
    strDestDir = Trim(strDestDir)
    '
    ' If strDestDir is quoted when passed to this routine, it
    ' should be quoted when it's returned.  The quotes need
    ' to be temporarily removed, though, for processing.
    '
    If VBA.Left(strDestDir, 1) = gstrQUOTE Then
        fQuoted = True
        strDestDir = strUnQuoteString(strDestDir)
    End If
    '
    ' We take the first part of destdir, and if its $( then we need to get the portion
    ' of destdir up to and including the last paren.  We then test against this for
    ' macro expansion.  If no ) is found after finding $(, then must assume that it's
    ' just a normal file name and do no processing.  Only enter the case statement
    ' if strDestDir starts with $(.
    '
    If VBA.Left$(strDestDir, 2) = strMACROSTART Then
        intPos = InStr(strDestDir, strMACROEND)

        Select Case VBA.Left$(strDestDir, intPos)
            Case gstrAPPDEST
                If gstrDestDir <> gstrNULL Then
                    strResolved = gstrDestDir
                Else
                    strResolved = "c:"
                End If
            Case gstrWINDEST
                strResolved = gstrWinDir
            Case gstrWINSYSDEST, gstrWINSYSDESTSYSFILE
                strResolved = gstrWinSysDir
            Case gstrPROGRAMFILES
                If TreatAsWin95() Then
                    Const strProgramFilesKey = "ProgramFilesDir"
    
                    If RegOpenKey(HKEY_LOCAL_MACHINE, strPathsKey, hKey) Then
                        RegQueryStringValue hKey, strProgramFilesKey, strResolved
                        RegCloseKey hKey
                    End If
                End If
    
                If strResolved = "" Then
                    'If not otherwise set, let strResolved be the root of the first fixed disk
                    strResolved = strRootDrive()
                End If
            Case gstrCOMMONFILES
                'First determine the correct path of Program Files\Common Files, if under Win95
                strResolved = strGetCommonFilesPath()
                If strResolved = "" Then
                    'If not otherwise set, let strResolved be the Windows directory
                    strResolved = gstrWinDir
                End If
            Case gstrCOMMONFILESSYS
                'First determine the correct path of Program Files\Common Files, if under Win95
                Dim strCommonFiles As String
                
                strCommonFiles = strGetCommonFilesPath()
                If strCommonFiles <> "" Then
                    'Okay, now just add \System, and we're done
                    strResolved = strCommonFiles & "System\"
                Else
                    'If Common Files isn't in the registry, then map the
                    'entire macro to the Windows\{system,system32} directory
                    strResolved = gstrWinSysDir
                End If
            Case gstrDAODEST
                strResolved = strGetDAOPath()
            Case Else
                intPos = 0
            'End Case
        End Select
    End If
    
    If intPos <> 0 Then
        AddDirSep strResolved
    End If

    If fAssumeDir = True Then
        If intPos = 0 Then
            '
            'if no drive spec, and doesn't begin with any root path indicator ("\"),
            'then we assume that this destination is relative to the app dest dir
            '
            If Mid$(strDestDir, 2, 1) <> gstrCOLON Then
                If VBA.Left$(strDestDir, 1) <> gstrSEP_DIR Then
                    strResolved = gstrDestDir
                End If
            End If
        Else
            If Mid$(strDestDir, intPos + 1, 1) = gstrSEP_DIR Then
                intPos = intPos + 1
            End If
        End If
    End If

    If fQuoted = True Then
        ResolveDestDir = strQuoteString(strResolved & Mid$(strDestDir, intPos + 1), True, False)
    Else
        ResolveDestDir = strResolved & Mid$(strDestDir, intPos + 1)
    End If
End Function
Function RegPathWinCurrentVersion() As String
    RegPathWinCurrentVersion = "SOFTWARE\Microsoft\Windows\CurrentVersion"
End Function
Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE And Right$(strQuotedString, 1) = gstrQUOTE Then
        '
        ' It's quoted.  Get rid of the quotes.
        '
        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
    End If
    strUnQuoteString = strQuotedString
End Function
Function TreatAsWin95() As Boolean
    If IsWindows95() Then
        TreatAsWin95 = True
    ElseIf fNTWithShell() Then
        TreatAsWin95 = True
    Else
        TreatAsWin95 = False
    End If
End Function
Public Sub TrocarVersaoExe(ObjMe As Object)
   Dim IniFile As String
   Dim sPath As String
   Dim LerIni  As Boolean
   Dim sArq  As String
   
   Dim nVersaoInst  As String
   Dim nVersaoMaq As String
   Dim SourcePath As String
   
   On Error Resume Next
   
   SourcePath = GetTag(ObjMe, "SourcePath")
'MsgBox SourcePath
   If SourcePath = "" Then Exit Sub
   SourcePath = IIf(Right(SourcePath, 1) = "\", SourcePath, SourcePath & "\")
   
   IniFile = App.EXEName & ".ini"
   IniFile = SourcePath & IniFile
   
   If FileExists(IniFile) Then
      LerIni = (ReadIniFile(IniFile, "AutoInstall", "Status") = "1")
      If LerIni Then
      
         sPath = ReadIniFile(IniFile, "AutoIntall Files", "Path")
         sPath = IIf(sPath = "", App.Path, sPath)
         sPath = sPath & IIf(Right(sPath, 1) = "/", "\", "\")
         sArq = App.EXEName & ".exe"
'MsgBox "FileExists(" & sPath & sArq & ")"
         If FileExists(sPath & sArq) And SourcePath <> "" Then

            nVersaoInst = GetVersao(GetFileVersion(sPath & sArq))
            nVersaoMaq = GetVersao(GetFileVersion(SourcePath & sArq))
'MsgBox "nVersaoInst = " & CStr(nVersaoInst) & vbNewLine & "nVersaoMaq = " & CStr(nVersaoMaq)
            If nVersaoInst <> nVersaoMaq And nVersaoMaq <> 0 And nVersaoInst <> 0 Then
               Err = 0
'MsgBox "FileCopy(" & sPath & sArq & ", " & SourcePath & sArq & ")"
               Call FileCopy(sPath & sArq, SourcePath & sArq)
            End If
         End If
      End If
   End If
End Sub
Function IsWindows95() As Boolean
    Const dwMask95 = &H2&
    If GetWinPlatform() And dwMask95 Then
        IsWindows95 = True
    Else
        IsWindows95 = False
    End If
End Function
Function RegOpenKey(ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Boolean
    Dim lResult As Long
    Dim strHkey As String

    On Error GoTo 0

    strHkey = strGetHKEYString(hKey)

    lResult = OSRegOpenKey(hKey, lpszSubKey, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegOpenKey = True
        AddHkeyToCache phkResult, strHkey & "\" & lpszSubKey
    Else
        RegOpenKey = False
    End If
End Function
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strData As String) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    
    RegQueryStringValue = False
    On Error GoTo 0
    ' Get length/data type
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = True
                strData = StripTerminator(strBuf)
            End If
        End If
    End If
End Function

Function RegCloseKey(ByVal hKey As Long) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = OSRegCloseKey(hKey)
    RegCloseKey = (lResult = ERROR_SUCCESS)
End Function
Function strRootDrive() As String
    Dim intDriveNum As Integer
    
    For intDriveNum = 0 To Asc("Z") - Asc("A") - 1
        If GetDriveType(intDriveNum) = intDRIVE_FIXED Then
            strRootDrive = Chr$(Asc("A") + intDriveNum) & gstrCOLON & gstrSEP_DIR
            Exit Function
        End If
    Next intDriveNum
    
    strRootDrive = "C:\"
End Function
Public Function strGetCommonFilesPath() As String
    Dim hKey As Long
    Dim strPath As String
    
    If TreatAsWin95() Then
        Const strCommonFilesKey = "CommonFilesDir"

        If RegOpenKey(HKEY_LOCAL_MACHINE, RegPathWinCurrentVersion(), hKey) Then
            RegQueryStringValue hKey, strCommonFilesKey, strPath
            RegCloseKey hKey
        End If
    End If

    If strPath <> "" Then
        AddDirSep strPath
    End If
    
    strGetCommonFilesPath = strPath
End Function
Private Function strGetDAOPath() As String
    Const strMSAPPS$ = "MSAPPS\"
    Const strDAO3032$ = "DAO3032.DLL"
    
    'first look in the registry
    Const strKey = "SOFTWARE\Microsoft\Shared Tools\DAO"
    Const strValueName = "Path"
    Dim hKey As Long
    Dim strPath As String

    If RegOpenKey(HKEY_LOCAL_MACHINE, strKey, hKey) Then
        RegQueryStringValue hKey, strValueName, strPath
        RegCloseKey hKey
    End If

    If strPath <> "" Then
        strPath = GetPathName(strPath)
        AddDirSep strPath
        strGetDAOPath = strPath
        Exit Function
    End If
    
    'It's not yet in the registry, so we need to decide
    'where the directory should be, and then need to place
    'that location in the registry.

    If TreatAsWin95() Then
        'For Win95, use "Common Files\Microsoft Shared\DAO"
        ' strPath = strGetCommonFilesPath() & ResolveResString(resMICROSOFTSHARED) & "DAO\"
        strPath = strGetCommonFilesPath() & "Microsoft Shared\" & "DAO\"
    Else
        'Otherwise use Windows\MSAPPS\DAO
        strPath = gstrWinDir & strMSAPPS & "DAO\"
    End If
    
    'Place this information in the registry (note that we point to DAO3032.DLL
    'itself, not just to the directory)
    If RegCreateKey(HKEY_LOCAL_MACHINE, strKey, "", hKey) Then
        RegSetStringValue hKey, strValueName, strPath & strDAO3032, False
        RegCloseKey hKey
    End If

    strGetDAOPath = strPath
End Function
Public Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub
Public Function strQuoteString(strUnQuotedString As String, Optional vForce As Variant, Optional vTrim As Variant)
'
' This routine adds quotation marks around an unquoted string, by default.  If the string is already quoted
' it returns without making any changes unless vForce is set to True (vForce defaults to False) except that white
' space before and after the quotes will be removed unless vTrim is False.  If the string contains leading or
' trailing white space it is trimmed unless vTrim is set to False (vTrim defaults to True).
'
    Dim strQuotedString As String
    
    If IsMissing(vForce) Then
        vForce = False
    End If
    If IsMissing(vTrim) Then
        vTrim = True
    End If
    
    strQuotedString = strUnQuotedString
    '
    ' Trim the string if necessary
    '
    If vTrim = True Then
        strQuotedString = Trim(strQuotedString)
    End If
    '
    ' See if the string is already quoted
    '
    If vForce = False Then
        If (VBA.Left(strQuotedString, 1) = gstrQUOTE) And (Right(strQuotedString, 1) = gstrQUOTE) Then
            '
            ' String is already quoted.  We are done.
            '
            GoTo DoneQuoteString
        End If
    End If
    '
    ' Add the quotes
    '
    strQuotedString = gstrQUOTE & strQuotedString & gstrQUOTE
DoneQuoteString:
    strQuoteString = strQuotedString
End Function
Private Function strGetHKEYString(ByVal hKey As Long) As String
    Dim strKey As String

    'Is the hkey predefined?
    strKey = strGetPredefinedHKEYString(hKey)
    If strKey <> "" Then
        strGetHKEYString = strKey
        Exit Function
    End If
    
    'It is not predefined.  Look in the cache.
    Dim intIdx As Integer
    intIdx = intGetHKEYIndex(hKey)
    If intIdx >= 0 Then
        strGetHKEYString = hkeyCache(intIdx).strHkey
    Else
        strGetHKEYString = ""
    End If
End Function
Private Sub AddHkeyToCache(ByVal hKey As Long, ByVal strHkey As String)
    Dim intIdx As Integer
    
    intIdx = intGetHKEYIndex(hKey)
    If intIdx < 0 Then
        'The key does not already exist.  Add it to the end.
        On Error Resume Next
        ReDim Preserve hkeyCache(0 To UBound(hkeyCache) + 1)
        If Err Then
            'If there was an error, it means the cache was empty.
            On Error GoTo 0
            ReDim hkeyCache(0 To 0)
        End If
        On Error GoTo 0

        intIdx = UBound(hkeyCache)
    Else
        'The key already exists.  It will be replaced.
    End If

    hkeyCache(intIdx).hKey = hKey
    hkeyCache(intIdx).strHkey = strHkey
End Sub
Function GetDriveType(ByVal intDriveNum As Integer) As Integer
    '
    ' This function expects an integer drive number in Win16 or a string in Win32
    '
    Dim strDriveName As String
    
    strDriveName = Chr$(Asc("A") + intDriveNum) & gstrSEP_DRIVE & gstrSEP_DIR
    GetDriveType = CInt(GetDriveType32(strDriveName))
End Function
Public Function GetPathName(ByVal strFilename As String) As String
    Dim intPos As Integer
    Dim strPathOnly As String
    Dim dirTmp As DirListBox
    Dim i As Integer

    On Error Resume Next


    Err = 0
    
    intPos = Len(strFilename)

    '
    'Change all '/' chars to '\'
    '

    For i = 1 To Len(strFilename)
        If Mid$(strFilename, i, 1) = gstrSEP_DIRALT Then
            Mid$(strFilename, i, 1) = gstrSEP_DIR
        End If
    Next i

    If InStr(strFilename, gstrSEP_DIR) = intPos Then
        If intPos > 1 Then
            intPos = intPos - 1
        End If
    Else
        Do While intPos > 0
            If Mid$(strFilename, intPos, 1) <> gstrSEP_DIR Then
                intPos = intPos - 1
            Else
                Exit Do
            End If
        Loop
    End If

    If intPos > 0 Then
        strPathOnly = VBA.Left$(strFilename, intPos)
        If Right$(strPathOnly, 1) = gstrCOLON Then
            strPathOnly = strPathOnly & gstrSEP_DIR
        End If
    Else
        strPathOnly = CurDir$
    End If

    If Right$(strPathOnly, 1) = gstrSEP_DIR Then
        strPathOnly = VBA.Left$(strPathOnly, Len(strPathOnly) - 1)
    End If

    GetPathName = UCase16(strPathOnly)
    
    Err = 0
End Function
Function RegCreateKey(ByVal hKey As Long, ByVal lpszSubKeyPermanent As String, ByVal lpszSubKeyRemovable As String, phkResult As Long) As Boolean
    Dim lResult As Long
    Dim strHkey As String
    Dim fLog As Boolean
    Dim strSubKeyFull As String

    On Error GoTo 0

    If lpszSubKeyPermanent = "" Then
        RegCreateKey = False 'Error: lpszSubKeyPermanent must not = ""
        Exit Function
    End If
    
    If VBA.Left$(lpszSubKeyRemovable, 1) = "\" Then
        lpszSubKeyRemovable = Mid$(lpszSubKeyRemovable, 2)
    End If

    If lpszSubKeyRemovable = "" Then
        fLog = False
    Else
        fLog = True
    End If
    
    If lpszSubKeyRemovable <> "" Then
        strSubKeyFull = lpszSubKeyPermanent & "\" & lpszSubKeyRemovable
    Else
        strSubKeyFull = lpszSubKeyPermanent
    End If
    strHkey = strGetHKEYString(hKey)

    If fLog Then
        NewAction _
          gstrKEY_REGKEY, _
          """" & strHkey & "\" & lpszSubKeyPermanent & """" _
            & ", " & """" & lpszSubKeyRemovable & """"
    End If

    lResult = OSRegCreateKey(hKey, strSubKeyFull, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegCreateKey = True
        If fLog Then
            CommitAction
        End If
        AddHkeyToCache phkResult, strHkey & "\" & strSubKeyFull
    Else
        RegCreateKey = False
        MsgBox "An error occurred trying to update the Windows registration database.", vbOKOnly Or vbExclamation, gstrTitle
        'MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            'ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function
Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strData As String, Optional ByVal fLog) As Boolean
    Dim lResult As Long
    Dim strHkey As String
    
    On Error GoTo 0
    
    If IsMissing(fLog) Then fLog = True

    If hKey = 0 Then
        Exit Function
    End If
    
    strHkey = strGetHKEYString(hKey)

    If fLog Then
        NewAction _
          gstrKEY_REGVALUE, _
          """" & strHkey & """" _
            & ", " & """" & strValueName & """"
    End If

    lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)) + 1)
    
    If lResult = ERROR_SUCCESS Then
        RegSetStringValue = True
        If fLog Then
            CommitAction
        End If
    Else
        RegSetStringValue = False
        MsgBox "An error occurred trying to update the Windows registration database.", vbOKOnly Or vbExclamation, gstrTitle
        'MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            'ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function
Private Function strGetPredefinedHKEYString(ByVal hKey As Long) As String
    Select Case hKey
        Case HKEY_CLASSES_ROOT
            strGetPredefinedHKEYString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER
            strGetPredefinedHKEYString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            strGetPredefinedHKEYString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            strGetPredefinedHKEYString = "HKEY_USERS"
        'End Case
    End Select
End Function
Private Function intGetHKEYIndex(ByVal hKey As Long) As Integer
    Dim intUBound As Integer
    
    On Error Resume Next
    intUBound = UBound(hkeyCache)
    If Err Then
        'If there was an error accessing the ubound of the array,
        'then the cache is empty
        GoTo NotFound
    End If
    On Error GoTo 0

    Dim intIdx As Integer
    For intIdx = 0 To intUBound
        If hkeyCache(intIdx).hKey = hKey Then
            intGetHKEYIndex = intIdx
            Exit Function
        End If
    Next intIdx
    
NotFound:
    intGetHKEYIndex = -1
End Function
Function UCase16(ByVal str As String)
    UCase16 = str
End Function
Sub NewAction(ByVal strKey As String, ByVal strData As String)
    ShowLoggingError DllNewAction(strKey, strData), LogErrFatal
End Sub
Sub CommitAction()
    ShowLoggingError DllCommitAction(), LogErrFatal
End Sub
Sub AbortAction()
    ShowLoggingError DllAbortAction(), LogErrFatal
End Sub
Sub ShowLoggingError(ByVal lErr As Long, ByVal lErrSeverity As Long)
    If lErr = LOGERR_SUCCESS Then
        Exit Sub
    End If
    
    Dim strErrMsg As String
    Static fRecursive As Boolean
    
    If fRecursive Then
        'If we're getting called recursively, we're likely
        'getting errors while trying to write out errors to
        'the logfile.  Nothing to do but turn off logging
        'and abort setup.
        DisableLogging
        'MsgError ResolveResString(resUNEXPECTED), vbExclamation Or vbOKOnly, gstrTitle
        MsgBox "An unexpected setup error has occurred!", vbExclamation Or vbOKOnly, gstrTitle
        'ExitSetup frmSetup1, gintRET_FATAL
    End If

    fRecursive = True

    Select Case lErr
        Case LOGERR_OUTOFMEMORY, LOGERR_WRITEERROR, LOGERR_UNEXPECTED, LOGERR_FILENOTFOUND
            'strErrMsg = ResolveResString(resUNEXPECTED)
            strErrMsg = "An unexpected setup error has occurred!"
        
        Case LOGERR_INVALIDARGS, LOGERR_EXCEEDEDCAPACITY, LOGERR_NOCURRENTACTION
            'Note: These errors are most likely the result of improper customization
            'of this project.  Make certain that any changes you have made to these
            'files are valid and bug-free.
            'LOGERR_INVALIDARGS -- some parameter to a logging function was invalid or improper
            'LOGERR_EXCEEDEDCAPACITY -- the stacking depth of actions has probably been
            '   exceeded.  This most likely means that CommitAction or AbortAction statements
            '   are missing from your code.
            'LOGERR_NOCURRENTACTION -- the logging function you tried to use requires that
            '   there be a current action, but there was none.  Check for a missing NewAction
            '   statement.
            'strErrMsg = ResolveResString(resUNEXPECTED)
            strErrMsg = "An unexpected setup error has occurred!"
        Case Else
            'strErrMsg = ResolveResString(resUNEXPECTED)
            strErrMsg = "An unexpected setup error has occurred!"
        'End Case
    End Select
    
    Dim iRet As Integer
    Dim fAbort As Boolean
    
    fAbort = False
    If lErrSeverity = LogErrOK Then
        ' User can select whether or not to continue
        iRet = MsgBox(strErrMsg, vbOKCancel Or vbExclamation, gstrTitle)
        If gfNoUserInput Then iRet = vbCancel ' can't continue if silent install.
        Select Case iRet
            Case vbOK
            Case vbCancel
                fAbort = True
            Case Else
                fAbort = True
            'End Case
        End Select
    Else
        ' Fatal
        MsgBox strErrMsg, vbOK Or vbExclamation, gstrTitle
        fAbort = True
    End If

    If fAbort Then
        'ExitSetup frmCopy, gintRET_ABORT
    End If

    fRecursive = False

End Sub
Sub DisableLogging()
    ShowLoggingError DllDisableLogging(), LogErrFatal
End Sub
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    ' If the string is quoted, remove the quotes.
    '
    strPathName = strUnQuoteString(strPathName)
    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = gstrSEP_DIR Then
        strPathName = VBA.Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function
Function GetWindowsDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf

        GetWindowsDir = strBuf
    Else
        GetWindowsDir = gstrNULL
    End If
End Function
Function GetWindowsSysDir() As String
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
'Public Function SincShell(Comando As String, Optional Modo = vbMinimizedFocus)
'    Dim Handle As Long
'    Dim Inst1 As Long
'    Dim Inst2 As Long
'    Dim Res As Long
'    Dim StrWinModule As String
'    'Objetivo: Fazer um loop para controlar o fim da última aplicacao DOS gerada
'    StrWinModule = "WINOLDAP"
'    Handle = GetModuleHandle(StrWinModule) 'Handle do Módulo WINOLDAP que carrega as apps DOS
'    Inst1 = GetModuleUsage(Handle)         'Número de aplicações DOS ativas antes de executar a atual
'
'    'SHELL -> roda um pgm executável
'
'    Res = Shell(Comando, Modo)
'    Handle = GetModuleHandle(StrWinModule)  'Handle do Módulo WINOLDAP que carrega as apps DOS
'    Inst2 = GetModuleUsage(Handle)        'Novo numero de aplicacoes DOS (inst2=inst1+1)  (depois de executar o shell acima)
'    While Not (Inst2 = Inst1)
'        Inst2 = GetModuleUsage(Handle)    'Aguarda o fim da última aplicação gerada (que no caso e a chamada do shell acima)
'        DoEvents
'    Wend
'    SincShell = Res
'
'End Function
Public Sub AutoInstall(Aplic As App)
   Dim strFilename   As String
   Dim FileLST       As String
   Dim strKey        As String
   Dim sLinha        As String
   Dim i             As Integer
   Dim j             As Integer
   Dim Pos           As Integer
   Dim posfim        As Integer
   Dim sExt          As String
   Dim bExtensao     As Boolean
   Dim bInstala      As Boolean
   Dim bDebug        As Boolean
   
   Dim sNome         As String
   Dim sPath         As String
   Dim sVersaoInst   As String
   Dim sVersaoMaq    As String
   Dim nVersaoInst   As Double
   Dim nVersaoMaq    As Double
   
   Dim lComando      As String
   
   Dim lFileComando  As String
   Dim NumProp       As Integer
   Dim nArq          As Integer
   Dim sArq          As String
   Dim gstrINI_FILES As String
   Dim LstDate       As String
      
   Dim IniFile       As String
   Dim sStatus       As String
   Dim LerIni        As Boolean
   
   On Error Resume Next

   bDebug = ((InStr(UCase(Command$), "DEBUG") <> 0))

   IniFile = App.EXEName & ".ini"
   IniFile = App.Path & "\" & IniFile
    
   If bDebug Then MsgBox "FileExists(" & IniFile & ") "
   
   
   Call Kill("C:\TMP\" & App.ProductName & "\" & App.EXEName & ".ini")
   
   If bDebug Then MsgBox "Command$=" & Command$
   If bDebug Then MsgBox "InStr(UCase(Command$), UCase(|SourcePath)=) = " & CStr(InStr(UCase(Command$), UCase("|SourcePath=")))
   If InStr(UCase(Command$), UCase("|SourcePath=")) <> 0 Then
      sPath = Mid(UCase(Command$), InStr(UCase(Command$), UCase("|SourcePath=")) + 12)
      sPath = Mid(sPath, 1, InStr(sPath, "|") - 1)
      sPath = sPath & IIf(Right(sPath, 1) = "\", "", "\")
      
      If bDebug Then MsgBox "sPath=" & sPath
      
      Call FileCopy(sPath & App.EXEName & ".ini", "C:\TMP\" & App.ProductName & "\" & App.EXEName & ".ini")
      sPath = ""
   End If
   
   If App.Path <> "C:\TMP\" & App.ProductName Then
      If bDebug Then MsgBox "Kill(C:\TMP\" & App.ProductName & "\" & App.EXEName & ".exe)"

      Call Kill("C:\TMP\" & App.ProductName & "\" & App.EXEName & ".exe")
   End If
   
   If bDebug Then MsgBox "If FileExists(" & IniFile & ") Then"
   If FileExists(IniFile) Then
      If bDebug Then
         MsgBox "If FileExists(" & IniFile & ") Then 'R: True"
      End If

      sStatus = ReadIniFile(IniFile, "AutoInstall", "Status")
      LerIni = (sStatus = "1" Or sStatus = "2")
      If LerIni Then
      
         sPath = ReadIniFile(IniFile, "AutoIntall Files", "Path")
         sPath = IIf(sPath = "", App.Path, sPath)
         sPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "")
         sArq = App.EXEName & ".exe"
         If bDebug Then
            MsgBox "FileExists(" & sPath & sArq & ") " & IIf(FileExists(IniFile), "= True ", "= False")
         End If
         If FileExists(sPath & sArq) Then
            nVersaoInst = GetVersao(GetFileVersion(sPath & sArq))
            nVersaoMaq = GetVersao(GetFileVersion(App.Path & "\" & sArq))
            If bDebug Then
               MsgBox sPath & sArq & " Vs." & CStr(nVersaoInst) & vbNewLine & App.Path & "\" & sArq & " Vs." & CStr(nVersaoMaq)
            End If
            If nVersaoInst <> nVersaoMaq And nVersaoMaq <> 0 And nVersaoInst <> 0 Then
               Err = 0
               Call FileCopy(sPath & sArq, "C:\TMP\" & App.ProductName & "\" & sArq)
               If Err <> 0 Then '* Path not found
                  Err = 0
                  Call MkDir("C:\TMP")
                  If Err <> 0 Then
                     Call MkDir("C:\TMP\" & App.ProductName)
                  End If
                  Call FileCopy(sPath & sArq, "C:\TMP\" & App.ProductName & "\" & sArq)
               End If

               If bDebug Then
                  MsgBox "SincShell(C:\TMP\" & App.ProductName & "\" & sArq & " |SourcePath=" & App.Path & "| /Debug, False)"
                  Call SincShell("C:\TMP\" & App.ProductName & "\" & sArq & " |SourcePath=" & App.Path & "| /Debug", False)
               Else
                  Call SincShell("C:\TMP\" & App.ProductName & "\" & sArq & " |SourcePath=" & App.Path & "|", False)
               End If
               End
               
            End If
         Else
         
         End If
         
         i = 1
         While ReadIniFile(IniFile, "AutoIntall Files", "File" & CStr(i)) <> ""

            sArq = ReadIniFile(IniFile, "AutoIntall Files", "File" & CStr(i))
            If FileExists(sPath & sArq) Then
               nVersaoInst = GetVersao(GetFileVersion(sPath & sArq))
               If FileExists(GetWindowsSysDir() & sArq) Then
                  nVersaoMaq = GetVersao(GetFileVersion(GetWindowsSysDir() & sArq))
               Else
                  nVersaoMaq = -1
               End If
            Else
               nVersaoInst = 0
               nVersaoMaq = 0
            End If
            If bDebug Then
               MsgBox "i=" & CStr(i) & vbNewLine & sPath & sArq & " Vs." & nVersaoInst & vbNewLine & GetWindowsSysDir() & sArq & " Vs." & nVersaoMaq
            End If
            If (nVersaoInst <> nVersaoMaq And nVersaoMaq <> 0 And nVersaoInst <> 0) Or sStatus = "2" Then
               If FileExists(GetWindowsSysDir() & sArq) Then
                  Call RegServer(GetWindowsSysDir() & sArq, False, False)
               End If
               If bDebug And Err <> 0 Then
                  MsgBox "Call RegServer(GetWindowsSysDir() & sArq, False, False) " & vbNewLine & "Erro :" & CStr(Err) & "-" & Error
                  Err = 0
               End If
               
               If FileExists(GetWindowsSysDir() & sArq) Then
                  Call Kill(GetWindowsSysDir() & sArq)
               End If
               If bDebug And Err <> 0 Then
                  MsgBox "Call Kill(GetWindowsSysDir() & sArq) " & vbNewLine & "Erro :" & CStr(Err) & "-" & Error
                  Err = 0
               End If
               
               Call FileCopy(sPath & sArq, GetWindowsSysDir() & sArq)
               If bDebug And Err <> 0 Then
                  MsgBox "Call FileCopy(sPath & sArq, GetWindowsSysDir() & sArq) " & vbNewLine & "Erro :" & CStr(Err) & "-" & Error
                  Err = 0
               End If
               
               Call RegServer(GetWindowsSysDir() & sArq, True, False)
               If bDebug And Err <> 0 Then
                  MsgBox "Call RegServer(GetWindowsSysDir() & sArq, True, False) " & vbNewLine & "Erro :" & CStr(Err) & "-" & Error
                  Err = 0
               End If
               
            End If
            i = i + 1
         Wend
         'Exit Sub
      End If
   Else
      If bDebug Then
         MsgBox "If FileExists(" & IniFile & ") Then 'R: False"
      End If
   
      'If Mid(Command(), 1, Len("|SourcePath=")) = "|SourcePath=" Then
      '   sArq = App.EXEName & ".exe"
      '   sPath = SourcePath
      '   sPath = sPath & IIf(Right(sPath) = "\", "", "\")
      '   Err = 0
      '   Call FileCopy( "C:\TMP\" & App.ProductName & "\" & sArq, sPath & sArq)
      '   If Err = 0 Then
      '      Call SincShell(sPath & sArq, False)
      '      End
      '   End If
      'End If
   End If
   
   gstrINI_FILES$ = "Setup1 Files"
   
   LerIni = LerIni Or (Val("0" & ReadIniFile(IniFile, "AutoInstall", "Status")) <> 0)
   
   If bDebug Then
      MsgBox "LerIni : " & IIf(LerIni, "True", "False") & vbNewLine & "Função : AutoInstall"
   End If
   
   If LerIni Then
      sPath = ReadIniFile(IniFile, "AutoIntall Files", "Path")
      sPath = sPath & IIf(Right(sPath, 1) = "\", "", "\")
      sPath = sPath & "Package\"
      
      FileLST = sPath & "SETUP.LST"
      If Not FileExists(FileLST) Then
         sPath = Aplic.Path & "\Setup"
      End If
   Else
      sPath = Aplic.Path & "\Setup"
   End If
   
   sPath = sPath & IIf(Right(sPath, 1) = "\", "", "\")
   
   FileLST = sPath & "SETUP.LST"
   If Not FileExists(FileLST) Then
      FileLST = Aplic.Path & "\Package\SETUP.LST"
      If Not FileExists(FileLST) Then
         Exit Sub
      End If
   End If

   gstrWinDir = GetWindowsDir()
   gstrWinSysDir = GetWindowsSysDir()
   If Not FileExists(gstrWinSysDir & "VB5STKIT.DLL") Then
      Exit Sub
   End If
   
   gstrDestDir = ResolveDestDir(ReadIniFile(FileLST, "Setup", "DefaultDir"))
   
   If Not FileExists(FileLST) Then
      FileLST = Aplic.Path & "\Package\SETUP.LST"
      If Not FileExists(FileLST) Then
         Exit Sub
      End If
   End If
      
   
'   LstDate = GetSetting(App.EXEName, "Update System", "FileDate", "01/01/1000")
'   If FileDateTime(FileLST) <= CVDate(LstDate) Then
'      Exit Sub
'   End If

   i = 0
   nArq = FreeFile()
   sArq = GetPathName(FileLST) & "\AutoVer.log"
   sArq = "C:\TMP\" & App.ProductName & "\AutoVer.log"
   
   If bDebug Then
      MsgBox "Kill " & sArq & vbNewLine & "Função : AutoInstall"
   End If
   
   Kill sArq
   Open sArq For Output As #nArq

   If bDebug Then
      MsgBox "Kill " & sArq & vbNewLine & "Função : AutoInstall"
   End If
   Print #nArq, "Auto-Instalação em " & Format(Now, "dd/mm/yyyy hh:mm")
   Print #nArq, "|----------------|----------------|----------------|"
   Print #nArq, "| Arquivo        | Versão Atual   | Versão Setup   |"
   Print #nArq, "|----------------|----------------|----------------|"
   
   Do
      i = i + 1
      strKey = "File" & Trim(CStr(i))
      sLinha = ReadIniFile(FileLST, gstrINI_FILES$, strKey)
      If Trim(sLinha) <> "" Then
         Pos = 2
         posfim = InStr(Pos, sLinha, ",")
         NumProp = 6
         For j = 0 To NumProp
            If j = 0 Then sNome = Mid(sLinha, Pos, posfim - Pos)
            If j = 1 Then sPath = ResolveDestDir(Mid(sLinha, Pos, posfim - Pos))
            If j = 6 Then
               sVersaoInst = Mid(sLinha, Pos, posfim - Pos)
            End If
            Pos = InStr(Pos, sLinha, ",")
            Pos = Pos + 1
               If Pos <= 1 Then Exit For
            If InStr(Pos, sLinha, ",") = 0 Then
               posfim = Len(sLinha) + 1
            Else
               posfim = InStr(Pos, sLinha, ",")
            End If
         Next
         sExt = UCase(Extension(sNome))
         bExtensao = (sExt = "DLL" Or sExt = "OCX" Or sExt = "TLB")
         If Pos > 0 And bExtensao Then
            If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
            strFilename = sPath & sNome
            sVersaoMaq = GetFileVersion(strFilename)
            
            Pos = InStr(sVersaoInst, ".")
            While Pos <> 0
               sVersaoInst = Mid(sVersaoInst, 1, Pos - 1) & Mid(sVersaoInst, Pos + 1)
               Pos = InStr(sVersaoInst, ".")
            Wend
            Pos = InStr(sVersaoMaq, ".")
            While Pos <> 0
               sVersaoMaq = Mid(sVersaoMaq, 1, Pos - 1) & Mid(sVersaoMaq, Pos + 1)
               Pos = InStr(sVersaoMaq, ".")
            Wend
            
            If Val(sVersaoInst) > Val(sVersaoMaq) Or Not FileExists(strFilename) Then
               '* Gravalog
               Print #nArq, "|S " & Mid(sNome & Space(15), 1, 14) & "| " & Mid(sVersaoMaq & Space(15), 1, 15) & "| " & Mid(sVersaoInst & Space(15), 1, 15) & "|"
               bInstala = True
            Else
               Print #nArq, "|N " & Mid(sNome & Space(15), 1, 14) & "| " & Mid(sVersaoMaq & Space(15), 1, 15) & "| " & Mid(sVersaoInst & Space(15), 1, 15) & "|"
            End If
            Print #nArq, "|----------------|----------------|----------------|"
         End If
      End If
   Loop Until Trim(sLinha) = ""
   Close #nArq
   
   
   If bDebug Then MsgBox "bInstala = " & IIf(bInstala, "True", "False")
   
   If bInstala Then
      lComando = ""
      Call SaveSetting(App.EXEName, "Update System", "FileDate", FileDateTime(FileLST))
      lFileComando = GetPathName(FileLST) & "\Instala.exe"
      If Not FileExists(lFileComando) Then
         lFileComando = GetPathName(FileLST) & "\Setup.exe"
      End If
      lComando = lFileComando & " /s AutoIns.log"
      If FileExists(lFileComando) Then
         Call SincShell(lComando)
      End If
   Else
      'Kill sArq
   End If
Exit Sub
OpenError:
'   CmDialog.filename = ""
   Resume Next
End Sub
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
Public Function RegServer(sServerPath As String, Optional fRegister = True, Optional fMsg As Boolean = True, Optional isActivexExe As Boolean = False) As Boolean
  
   Dim hMod As Long               ' module handle
   Dim lpfn As Long                  ' reg/unreg function address
   Dim sCmd As String             ' msgbox string
   Dim lpThreadID As Long        ' unused, receives the thread ID
   Dim hThread As Long            ' thread handle
   Dim fSuccess As Boolean     ' if things worked
   Dim dwExitCode As Long      ' thread's exit code if it doesn't finish
   
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
   
   If fMsg Then
      If fSuccess Then
        MsgBox "Successfully " & sCmd & "ed " & sServerPath   ' past tense
        RegServer = True
      Else
        MsgBox "Failed To " & sCmd & " " & sServerPath, vbExclamation
      End If
   End If
End Function
Public Sub AtualizaDLL(pDLLMaq As String, pDLLNova As String)
   Dim nVersaoMaq    As Double
   Dim nVersaoNova   As Double
   
   If FileExists(pDLLNova) Then
      nVersaoNova = GetVersao(GetFileVersion(pDLLNova))
      
      If FileExists(pDLLMaq) Then
         nVersaoMaq = GetVersao(GetFileVersion(pDLLMaq))
      Else
         nVersaoMaq = -1
      End If
   Else
      nVersaoNova = 0
      nVersaoMaq = 0
   End If
   If nVersaoNova > nVersaoMaq Then
      If FileExists(pDLLMaq) Then
         Call RegServer(pDLLMaq, False, False)
         Call Kill(pDLLMaq)
      End If
      
      Call FileCopy(pDLLNova, pDLLMaq)
      Call RegServer(pDLLMaq, True, False)
   End If
End Sub

