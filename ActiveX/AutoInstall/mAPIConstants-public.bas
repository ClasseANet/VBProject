Attribute VB_Name = "mAPIConstants"
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' This code was written by The Frog Prince
'
' If you have questions or comments, I can be reached at
'        TheFrogPrince@hotmail.com
' If you wanna see more cool vb user controls, classes, code,
' and add-ins like this one, or updates to this code, go to
' my web page at
'        http://members.tripod.com/the__frog__prince/
' You are free to use, re-write, or otherwise do as you wish
' with this code.  However, if you do a cool enhancement, I
' would appreciate it if you could e-mail it to me.  I like
' to see what people do with my stuff.  =)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


Option Explicit


Public Enum enumBorderFlags
    BF_ADJUST = &H2000
    BF_BOTTOM = &H8
    BF_DIAGONAL = &H10
    BF_FLAT = &H4000
    BF_LEFT = &H1
    BF_MIDDLE = &H800
    BF_MONO = &H8000
    BF_RIGHT = &H4
    BF_SOFT = &H1000
    BF_TOP = &H2
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
End Enum

Public Enum enumBorderEdges
    BDR_RAISEDINNER = &H4
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENINNER = &H8
    BDR_SUNKENOUTER = &H2
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum


Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Declare Function DrawEdge _
                                    Lib "user32" ( _
                                ByVal hdc As Long, _
                                qrc As Rect, _
                                ByVal edge As enumBorderEdges, _
                                ByVal grfFlags As enumBorderFlags) _
                            As Long


' List of different styles of keyboard entry allowed.
' Goes with the function ctlKeyPress()
Public Enum enumKeyPressAllowTypes
    NumbersOnly = 2 ^ 0
    Uppercase = 2 ^ 1
    NoSpaces = 2 ^ 2
    NoSingleQuotes = 2 ^ 3
    NoDoubleQuotes = 2 ^ 4
    AllowDecimal = 2 ^ 5
    AllowNegative = 2 ^ 6
    DatesOnly = 2 ^ 7
    TimesOnly = 2 ^ 8
    LettersOnly = 2 ^ 9
    AllowSpaces = 2 ^ 10
    AllowStars = 2 ^ 11
    AllowPounds = 2 ^ 12
End Enum


Public Const OFS_MAXPATHNAME = 128


Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Enum enumFileAttributes
    efaARCHIVE = &H20
    efaCOMPRESSED = &H800
    efaDIRECTORY = &H10
    efaHIDDEN = &H2
    efaNORMAL = &H80
    efaREADONLY = &H1
    efaSYSTEM = &H4
    efaTEMPORARY = &H100
End Enum

Public Enum enumDriveTypes
    DRIVE_CDROM = 5
    DRIVE_FIXED = 3
    DRIVE_RAMDISK = 6
    DRIVE_REMOTE = 4
    DRIVE_REMOVABLE = 2
End Enum

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Enum enumFileNameParts
    efpFileName = 2 ^ 0
    efpFileExt = 2 ^ 1
    efpFilePath = 2 ^ 2
    efpFileNameAndExt = efpFileName + efpFileExt
    efpFileNameAndPath = efpFilePath + efpFileName
    
End Enum

Public Type typeVolumeInformation
    sRootPathName As String
    sVolumeName As String
    lVolumeSerialNo As Long
    lMaximumComponentLength As Long
    lFileSystemFlags As Long
    sFileSystemName As String
End Type

Public Const MAX_PATH = 400

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function PathGetDriveNumber Lib "SHLWAPI.DLL" Alias "PathGetDriveNumberA" (ByVal pszPath As String) As Long
Public Declare Function PathStripToRoot Lib "SHLWAPI.DLL" Alias "PathStripToRootA" (ByVal pszPath As String) As Long
Public Declare Function PathIsNetworkPath Lib "SHLWAPI.DLL" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Boolean
Public Declare Function PathIsUNCServerShare Lib "SHLWAPI.DLL" Alias "PathIsUNCServerShareA" (ByVal pszPath As String) As Boolean
Public Declare Function PathIsUNCServer Lib "SHLWAPI.DLL" Alias "PathIsUNCServerA" (ByVal pszPath As String) As Boolean
Public Declare Function PathIsUNC Lib "SHLWAPI.DLL" Alias "PathIsUNCA" (ByVal pszPath As String) As Boolean
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function GetExpandedName Lib "lz32.dll" Alias "GetExpandedNameA" (ByVal lpszSource As String, ByVal lpszBuffer As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As enumFileAttributes
'Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
