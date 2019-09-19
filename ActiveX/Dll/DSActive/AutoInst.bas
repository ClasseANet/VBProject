Attribute VB_Name = "AutoInstallBas"
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
Public Const CSIDL_SYSTEM64 = 41       '// C:\WINNT\SysWOW64
Public Const CSIDL_COMMON = 43         '// C:\Program Files\Common Files

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
Public Sub SincShell(Comando As String, Optional Modo As VbAppWinStyle = vbMaximizedFocus, Optional EsperaProcesso = True)
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

