Attribute VB_Name = "Install"
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

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetDriveType32 Lib "kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
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

Global MyActive  As Object

'****************************************
'****************************************
'****************************************
Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public Const MAX_PATH As Integer = 260

Public Const CSIDL_DESKTOP = &H0       '// The Desktop - virtual folder
Public Const CSIDL_TEMPORARY = 1
Public Const CSIDL_PROGRAMS = 2        '// Program Files
Public Const CSIDL_CONTROLS = 3        '// Control Panel - virtual folder
Public Const CSIDL_PRINTERS = 4        '// Printers - virtual folder
Public Const CSIDL_DOCUMENTS = 5       '// C:\Documents and Settings\dsr\My Documents
Public Const CSIDL_FAVORITES = 6       '// C:\Documents and Settings\dsr\Favorites
Public Const CSIDL_STARTUP = 7         '// C:\Documents and Settings\dsr\Start Menu\Programs\Startup
Public Const CSIDL_RECENT = 8          '// C:\Documents and Settings\dsr\Recent
Public Const CSIDL_SENDTO = 9          '// C:\Documents and Settings\dsr\SendTo
Public Const CSIDL_BITBUCKET = 10      '// Recycle Bin - virtual folder
Public Const CSIDL_STARTMENU = 11      '// C:\Documents and Settings\dsr\Start Menu
Public Const CSIDL_DESKTOPFOLDER = 16  '// C:\Documents and Settings\dsr\Desktop
Public Const CSIDL_DRIVES = 17         '// My Computer - virtual folder
Public Const CSIDL_NETWORK = 18        '// Network Neighbourhood - virtual folder
Public Const CSIDL_NETHOOD = 19        '// C:\Documents and Settings\dsr\NetHood
Public Const CSIDL_FONTS = 20          '// C:\WINNT\Fonts
Public Const CSIDL_SHELLNEW = 21       '// C:\Documents and Settings\dsr\Templates
'22 - C:\Documents and Settings\All Users\Start Menu
'23 - C:\Documents and Settings\All Users\Start Menu\Programs
'24 - C:\Documents and Settings\All Users\Start Menu\Programs\Startup
'25 - C:\Documents and Settings\All Users\Desktop
'26 - C:\Documents and Settings\dsr\Application Data
'27 - C:\Documents and Settings\dsr\PrintHood
'28 - C:\Documents and Settings\dsr\Local Settings\Application Data
'31 - C:\Documents and Settings\All Users\Favorites
'32 - C:\Documents and Settings\dsr\Local Settings\Temporary Internet Files
'33 - C:\Documents and Settings\dsr\Cookies
'34 - C:\Documents and Settings\dsr\Local Settings\History
'35 - C:\Documents and Settings\All Users\Application Data
Public Const CSIDL_WINDOWS = 36        '// C:\WINNT
Public Const CSIDL_SYSTEM32 = 37       '// C:\WINNT\system32
Public Const CSIDL_PROGRAM_FILES = 38  '// C:\Program Files
'39 - C:\Documents and Settings\dsr\My Documents\My Pictures
'40 - C:\Documents and Settings\dsr
'Public Const CSIDL_SYSTEM32 = 41       '// C:\WINNT\system32
Public Const CSIDL_COMMON = 43         '// C:\Program Files\Common Files
'45 - C:\Documents and Settings\All Users\Templates
'46 - C:\Documents and Settings\All Users\Documents
'47 - C:\Documents and Settings\All Users\Start Menu\Programs\Administrative Tools
'48 - C:\Documents and Settings\dsr\Start Menu\Programs\Administrative Tools
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
   Dim Ret As Long
   
   On Error GoTo TrataErro

   IDProcess = Shell(Comando$, 1)
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

    Extension = ""

    intPos = Len(strFilename)

    Do While intPos > 0
        Select Case Mid$(strFilename, intPos, 1)
            Case "."
                Extension = Mid$(strFilename, intPos + 1)
                Exit Do
            Case "\", "/"
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
   
   If ExisteArquivo(IniFile) Then
      LerIni = (ReadIniFile(IniFile, "AutoInstall", "Status") = "1")
      If LerIni Then
      
         sPath = ReadIniFile(IniFile, "AutoIntall Files", "Path")
         sPath = IIf(sPath = "", App.Path, sPath)
         sPath = sPath & IIf(Right(sPath, 1) = "/", "\", "\")
         sArq = App.EXEName & ".exe"
'MsgBox "ExisteArquivo(" & sPath & sArq & ")"
         If ExisteArquivo(sPath & sArq) And SourcePath <> "" Then

            nVersaoInst = GetFileVersionNumber(sPath & sArq)
            nVersaoMaq = GetFileVersionNumber(SourcePath & sArq)
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
            If InStr(strBuf, Chr$(0)) > 0 Then
               strBuf = VBA.Left$(strBuf, InStr(strBuf, Chr$(0)) - 1)
            End If
            strData = strBuf
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
        strPath = GetNameFromPath(strPath, 1)
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
        'The key already exis  It will be replaced.
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
Function GetWindowsDir() As String
   Dim strBuf As String

   strBuf = Space$(255)

   '
   'Get the windows directory and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   '
   If GetWindowsDirectory(strBuf, 255) > 0 Then
      If InStr(strBuf, Chr$(0)) > 0 Then
         strBuf = VBA.Left$(strBuf, InStr(strBuf, Chr$(0)) - 1)
      End If
        
      AddDirSep strBuf

      GetWindowsDir = strBuf
   Else
     GetWindowsDir = gstrNULL
   End If
End Function
Function GetWindowsSysDir() As String
   Dim strBuf As String
   
   strBuf = Space$(255)
   
   '
   'Get the system directory and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   '
   If GetSystemDirectory(strBuf, 255) > 0 Then
      If InStr(strBuf, Chr$(0)) > 0 Then
         strBuf = VBA.Left$(strBuf, InStr(strBuf, Chr$(0)) - 1)
      End If
      
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
   Dim PosFim        As Integer
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
    
   If bDebug Then MsgBox "ExisteArquivo(" & IniFile & ") "
   
   
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
   
   If bDebug Then MsgBox "If ExisteArquivo(" & IniFile & ") Then"
   If ExisteArquivo(IniFile) Then
      If bDebug Then
         MsgBox "If ExisteArquivo(" & IniFile & ") Then 'R: True"
      End If

      sStatus = ReadIniFile(IniFile, "AutoInstall", "Status")
      LerIni = (sStatus = "1" Or sStatus = "2")
      If LerIni Then
      
         sPath = ReadIniFile(IniFile, "AutoIntall Files", "Path")
         sPath = IIf(sPath = "", App.Path, sPath)
         sPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "")
         sArq = App.EXEName & ".exe"
         If bDebug Then
            MsgBox "ExisteArquivo(" & sPath & sArq & ") " & IIf(ExisteArquivo(IniFile), "= True ", "= False")
         End If
         If ExisteArquivo(sPath & sArq) Then
            nVersaoInst = GetFileVersionNumber(sPath & sArq)
            nVersaoMaq = GetFileVersionNumber(App.Path & "\" & sArq)
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
'XXX               End
               
            End If
         Else
         
         End If
         
         i = 1
         While ReadIniFile(IniFile, "AutoIntall Files", "File" & CStr(i)) <> ""

            sArq = ReadIniFile(IniFile, "AutoIntall Files", "File" & CStr(i))
            If ExisteArquivo(sPath & sArq) Then
               nVersaoInst = GetFileVersionNumber(sPath & sArq)
               If ExisteArquivo(GetWindowsSysDir() & sArq) Then
                  nVersaoMaq = GetFileVersionNumber(GetWindowsSysDir() & sArq)
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
               If ExisteArquivo(GetWindowsSysDir() & sArq) Then
                  Call RegServer(GetWindowsSysDir() & sArq, False, False)
               End If
               If bDebug And Err <> 0 Then
                  MsgBox "Call RegServer(GetWindowsSysDir() & sArq, False, False) " & vbNewLine & "Erro :" & CStr(Err) & "-" & Error
                  Err = 0
               End If
               
               If ExisteArquivo(GetWindowsSysDir() & sArq) Then
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
         MsgBox "If ExisteArquivo(" & IniFile & ") Then 'R: False"
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
      If Not ExisteArquivo(FileLST) Then
         sPath = Aplic.Path & "\Setup"
      End If
   Else
      sPath = Aplic.Path & "\Setup"
   End If
   
   sPath = sPath & IIf(Right(sPath, 1) = "\", "", "\")
   
   FileLST = sPath & "SETUP.LST"
   If Not ExisteArquivo(FileLST) Then
      FileLST = Aplic.Path & "\Package\SETUP.LST"
      If Not ExisteArquivo(FileLST) Then
         Exit Sub
      End If
   End If

   gstrWinDir = GetWindowsDir()
   gstrWinSysDir = GetWindowsSysDir()
   If Not ExisteArquivo(gstrWinSysDir & "VB5STKIT.DLL") Then
      Exit Sub
   End If
   
   gstrDestDir = ResolveDestDir(ReadIniFile(FileLST, "Setup", "DefaultDir"))
   
   If Not ExisteArquivo(FileLST) Then
      FileLST = Aplic.Path & "\Package\SETUP.LST"
      If Not ExisteArquivo(FileLST) Then
         Exit Sub
      End If
   End If
      
   
'   LstDate = GetSetting(App.EXEName, "Update System", "FileDate", "01/01/1000")
'   If FileDateTime(FileLST) <= CVDate(LstDate) Then
'      Exit Sub
'   End If

   i = 0
   nArq = FreeFile()
   sArq = GetNameFromPath(FileLST, 1) & "AutoVer.log"
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
         PosFim = InStr(Pos, sLinha, ",")
         NumProp = 6
         For j = 0 To NumProp
            If j = 0 Then sNome = Mid(sLinha, Pos, PosFim - Pos)
            If j = 1 Then sPath = ResolveDestDir(Mid(sLinha, Pos, PosFim - Pos))
            If j = 6 Then
               sVersaoInst = Mid(sLinha, Pos, PosFim - Pos)
            End If
            Pos = InStr(Pos, sLinha, ",")
            Pos = Pos + 1
               If Pos <= 1 Then Exit For
            If InStr(Pos, sLinha, ",") = 0 Then
               PosFim = Len(sLinha) + 1
            Else
               PosFim = InStr(Pos, sLinha, ",")
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
            
            If Val(sVersaoInst) > Val(sVersaoMaq) Or Not ExisteArquivo(strFilename) Then
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
      lFileComando = GetNameFromPath(FileLST, 1) & "Instala.exe"
      If Not ExisteArquivo(lFileComando) Then
         lFileComando = GetNameFromPath(FileLST, 1) & "Setup.exe"
      End If
      lComando = lFileComando & " /s AutoIns.log"
      If ExisteArquivo(lFileComando) Then
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
Private Function GetTag(ByRef Controle As Variant, ByVal VarName As String, Optional VarDefault As String) As String
   Dim PosIni As Long, PosFim As Long
   Dim StrTAG As String
   Dim i%
   
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
Private Function SetTag(ByRef Controle As Variant, ByVal VarName As String, ByVal VarValor As String) As String
   Dim StrTAG As String, StrAux As String
   Dim PosIni As Long, PosFim As Long
   
   VarName = "|" & Trim(VarName) & "="
   
   If UCase(TypeName(Controle)) = "STRING" Then
      StrTAG = Controle
   Else
      StrTAG = Controle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(VarName))
   If PosIni > 0 Then
      PosFim = InStr(PosIni + 1, StrTAG$, "|")
      PosFim = IIf(PosFim = 0, Len(StrTAG) + 1, PosFim)
      StrAux = Mid(StrTAG, 1, PosIni - 1) & Mid(StrTAG, PosIni, Len(VarName)) & Trim(VarValor)
      StrAux = StrAux & Mid(StrTAG, PosFim, (Len(StrTAG) - PosFim) + 1)
      StrTAG = StrAux
   Else
      If Trim(StrTAG) = "" Then
         StrTAG = VarName & VarValor
      Else
         If UCase(TypeName(Controle)) = "STRING" Then
            StrTAG = Controle & VarName & VarValor
         Else
            StrTAG = Controle.Tag & VarName & VarValor
         End If
      End If
   End If
   If UCase(TypeName(Controle)) = "STRING" Then
      Controle = StrTAG
   Else
      Controle.Tag = StrTAG
   End If
   SetTag = StrTAG
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
Private Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, Optional DefaultValue As String) As String
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
Private Function GetFileVersionNumber(pFilename As String) As Double
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
'   L_GetFileVersionNumber = Val(sVer)
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
Private Function GetFileVersion(ByVal pFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
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

