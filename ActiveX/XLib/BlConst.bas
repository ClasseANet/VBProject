Attribute VB_Name = "BlConst"
Option Explicit
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
Public Const CSIDL_SYSTEM64 = 41       '// C:\WINNT\SysWOW64 ou system32
Public Const CSIDL_COMMON = 43         '// C:\Program Files\Common Files
'45 - C:\Documents and Settings\All Users\Templates
'46 - C:\Documents and Settings\All Users\Documents
'47 - C:\Documents and Settings\All Users\Start Menu\Programs\Administrative Tools
'48 - C:\Documents and Settings\dsr\Start Menu\Programs\Administrative Tools

Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_CHARFROMPOS& = &HD7
Public Const EM_SETCHARFORMAT = (WM_USER + 68)


'* WAIT FOR SINGLE OBJECT RETURN VALUES
'* Constants that are used by the API
Public Const CREATE_SUSPENDED = &H4
Public Const INFINITE = &HFFFFFFFF                             ' Infinite timeout
Public Const PROCESS_QUERY_INFORMATION = &H400&
Public Const STATUS_WAIT_0 = &H0
Public Const STATUS_ABANDONED_WAIT_0 = &H80
Public Const STATUS_TIMEOUT = &H102
Public Const STATUS_PENDING = &H103
Public Const STILL_ACTIVE = STATUS_PENDING
Public Const SYNCHRONIZE = &H100000
Public Const WAIT_FAILED = &HFFFFFFFF
Public Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)             ' The State of the specified object is signaled (success)
Public Const WAIT_ABANDONED = ((STATUS_ABANDONED_WAIT_0) + 0)  ' Thread went away before the mutex got signaled
Public Const WAIT_TIMEOUT = STATUS_TIMEOUT                     ' dwMilliseconds timed out
Public Const WM_CLOSE = &H10

Public Const CB_FINDSTRING = &H14C                  ' Used to search a Combo
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETItemHEIGHT = &H154
Public Const LB_FINDSTRING = &H18F                  ' Used to search a List box

Public Const GHND = &H42
Public Const MAXSIZE = 4096
Public Const CF_TEXT = 1

Public Const CONNECT_LAN As Long = &H2
Public Const CONNECT_MODEM As Long = &H1
Public Const CONNECT_PROXY As Long = &H4
Public Const CONNECT_OFFLINE As Long = &H20
Public Const CONNECT_CONFIGURED As Long = &H40
'Public Const PROCESS_QUERY_INFORMATION = &H400&

Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16&)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4

Public Const WS_BORDER As Long = &H800000
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_EX_LAYERED = &H80000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_THICKFRAME As Long = &H40000

Public Const SWP_NOSENDCHANGING = &H400
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2&
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const WMVSCROLL = &H115

Global Const gSubFolder = "ClasseA"
