Attribute VB_Name = "API_BAS"
Option Explicit
'**************************************************************
'****                                                      ****
'**             API Function Declarations                    **
'****                                                      ****
'**************************************************************
' Common to most controls
Public Enum WinMsgs
  WM_NULL = &H0
  WM_CREATE = &H1
  WM_DESTROY = &H2
  WM_MOVE = &H3
  WM_SIZE = &H5
  WM_ACTIVATE = &H6
  WM_SETFOCUS = &H7
  WM_KILLFOCUS = &H8
  WM_ENABLE = &HA
  WM_SETREDRAW = &HB
  WM_PAINT = &HF
  WM_ERASEBKGND = &H14
  WM_SYSCOLORCHANGE = &H15
  WM_WININICHANGE = &H1A
  WM_SETCURSOR = &H20
  WM_NEXTDLGCTL = &H28
  WM_DRAWItem = &H2B
  WM_MEASUREItem = &H2C
  
  WM_SETFONT = &H30
  WM_GETFONT = &H31
  WM_WINDOWPOSCHANGED = &H47
  WM_NOTIFY = &H4E
  WM_NCCREATE = &H81
  WM_NCDESTROY = &H82
  WM_NCCALCSIZE = &H83
  WM_GETDLGCODE = &H87
  WM_KEYDOWN = &H100
  WM_KEYUP = &H101
  WM_CHAR = &H102
  WM_COMMAND = &H111
  WM_TIMER = &H113
  WM_HSCROLL = &H114
  WM_VSCROLL = &H115
  WM_INITMENUPOPUP = &H117

  WM_MOUSEMOVE = &H200
  WM_LBUTTONDOWN = &H201
  WM_LBUTTONUP = &H202
  WM_LBUTTONDBLCLK = &H203
  WM_RBUTTONDOWN = &H204
  WM_RBUTTONUP = &H205
  WM_RBUTTONDBLCLK = &H206
  WM_MBUTTONDOWN = &H207
  WM_MBUTTONUP = &H208
  WM_MBUTTONDBLCLK = &H209
  WM_USER = &H400
End Enum   ' WinMsgs




Public Const CREATE_SUSPENDED = &H4
Public Const INFINITE = &HFFFFFFFF   ' Infinite timeout
' WaitForSingleObject rtn vals
Public Const STATUS_WAIT_0 = &H0
Public Const STATUS_ABANDONED_WAIT_0 = &H80
Public Const STATUS_TIMEOUT = &H102
Public Const WAIT_FAILED = &HFFFFFFFF
Public Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0) ' The State of the specified object is signaled (success)
Public Const WAIT_ABANDONED = ((STATUS_ABANDONED_WAIT_0) + 0) ' Thread went away before the mutex got signaled
Public Const WAIT_TIMEOUT = STATUS_TIMEOUT ' dwMilliseconds timed out
Public Const PROCESS_QUERY_INFORMATION = &H400&
Public Const STATUS_PENDING = &H103
Public Const STILL_ACTIVE = STATUS_PENDING

'* Cursor
Public Const IDC_WAIT = 32514&   ' Hourglass
Public Const IDC_ARROW = 32512&

'* ShowProgress
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const SB_GETRECT = (WM_USER + 10)
'************************************
'* Misc Windows Messages and Styles *
'************************************
Public Const SM_CXVSCROLL As Long = 2 ' Get Width Of Vertical ScrollBar
Public Const HDS_BUTTONS As Long = &H2
Public Const SW_SHOWNORMAL As Long = 1

'*********************************
'* Treeview Messages and Styles  *
'*********************************
Public Const GWL_Style As Long = (-16)
Public Const COLOR_WINDOW As Long = 5
Public Const COLOR_WINDOWTEXT As Long = 8


Public Const TVI_ROOT   As Long = &HFFFF0000
Public Const TVI_FIRST  As Long = &HFFFF0001
Public Const TVI_LAST   As Long = &HFFFF0002
Public Const TVI_SORT   As Long = &HFFFF0003


'treeview Styles
Public Const TVS_HASLINES As Long = 2
Public Const TVS_TRACKSELECT As Long = &H200&
Public Const TVS_FULLROWSELECT As Long = &H1000


Public Const TV_FIRST As Long = &H1100
Public Const TVM_DELETEItem As Long = (TV_FIRST + 1)
Public Const TVM_EXPAND = (TV_FIRST + 2)
Public Const TVM_GETItemRECT = (TV_FIRST + 4)
Public Const TVM_SELECTItem = (TV_FIRST + 11)
Public Const TVM_GETEDITCONTROL As Long = (TV_FIRST + 15)
Public Const TVM_HITTEST = (TV_FIRST + 17)
Public Const TVM_CREATEDRAGIMAGE = (TV_FIRST + 18)
Public Const TVM_GETNEXTItem As Long = (TV_FIRST + 10)
Public Const TVM_GETItem As Long = (TV_FIRST + 12)
Public Const TVM_SETItem As Long = (TV_FIRST + 13)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)

Public Const TVGN_ROOT                As Long = &H0
Public Const TVGN_NEXT                As Long = &H1
Public Const TVGN_PREVIOUS            As Long = &H2
Public Const TVGN_PARENT              As Long = &H3
Public Const TVGN_CHILD               As Long = &H4
Public Const TVGN_FIRSTVISIBLE        As Long = &H5
Public Const TVGN_NEXTVISIBLE         As Long = &H6
Public Const TVGN_PREVIOUSVISIBLE     As Long = &H7
Public Const TVGN_DROPHILITE          As Long = &H8
Public Const TVGN_CARET               As Long = &H9
Public Const EM_LIMITTEXT = &HC5
Public Type TV_Item
   Mask As Long
   hItem As Long
   State As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

'**********************
'* ListView Structure *
'**********************
Public Type LVBKIMAGE
   uFlags As Long
   hBmp As Long
   pszImage As String
   cchImageMax As Long
   xOffsetPercent As Long
   yOffsetPercent  As Long
End Type

Public Type TagInitCommonControlsEx
    dwSize As Long
    dwICC As Long
End Type

Public Const ICC_LISTVIEW_CLASSES As Long = &H1
Public Const CLR_NONE = &HFFFFFFFF
Public Const LVM_FIRST As Long = &H1000
Public Const LVBKIF_SOURCE_NONE = &H0
Public Const LVBKIF_SOURCE_HBITMAP = &H1
Public Const LVBKIF_SOURCE_URL = &H2
Public Const LVBKIF_SOURCE_Mask = &H3
Public Const LVBKIF_Style_NORMAL = &H0
Public Const LVBKIF_Style_TILE = &H10
Public Const LVBKIF_Style_Mask = &H10
Public Const LVM_SETBKIMAGEA = (LVM_FIRST + 68)
Public Const LVM_SETBKIMAGEW = (LVM_FIRST + 138)
Public Const LVM_GETBKIMAGEA = (LVM_FIRST + 69)
Public Const LVM_GETBKIMAGEW = (LVM_FIRST + 139)
Public Const LVM_SETBKIMAGE = LVM_SETBKIMAGEA
Public Const LVM_GETBKIMAGE = LVM_GETBKIMAGEA
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const LVM_SETEXTENDEDLISTVIEWStyle As Long = (LVM_FIRST + 54)
Public Const LVS_EX_FULLROWSELECT As Long = &H20
   
Public Const LVM_GETEXTENDEDLISTVIEWStyle As Long = (LVM_FIRST + 55)

Public Const LVS_EX_GRIDLINES As Long = &H1
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVS_EX_HEADERDRAGDROP As Long = &H10
Public Const LVS_EX_TRACKSELECT As Long = &H8
Public Const LVS_EX_ONECLICKACTIVATE As Long = &H40
Public Const LVS_EX_TWOCLICKACTIVATE As Long = &H80

'****************************
'* general window definitions
'****************************
Public Type SIZE
  cx As Long
  cy As Long
End Type
Public Type PointAPI   ' pt
  x As Long
  y As Long
End Type

Public Const SM_CXDRAG = 68
Public Const SM_CYDRAG = 69

'**************************************
'* Scroll Bar Commands for WM_H/VSCROLL
'**************************************
Public Const SB_LINEUP = 0
Public Const SB_LINELEFT = 0
Public Const SB_LINEDOWN = 1
Public Const SB_LINERIGHT = 1


'******************************************************************
'******************************************************************
'* TreeView definitions
'*********************************
'* User-defined as the maximum treeview Item text length.
'* If an Items text exceeds this value when calling GetTVItemText, there will be problems...
'*********************************
'Public Const MAX_Item = 256
'*********************************
'* Callback constants
'*********************************
'*********************************
'* T/LVItem.pszText
'*********************************
'Public Const LPSTR_TEXTCALLBACK = (-1)
'*********************************
'* TVItem.cChildren
'*********************************
'Public Const I_CHILDRENCALLBACK = (-1&)
'*********************************
''* TVItem.iImage/iSelectedImage, LVItem.iImage
'*********************************
Public Const I_IMAGECALLBACK = (-1)

'*********************************
'* TVM_EXPAND wParam action flags
'*********************************
Public Const TVE_EXPAND = &H2

'****************************
'* TVItem State, stateMask
'****************************
Public Enum TVHITTESTINFO_Flags
  TVHT_NOWHERE = &H1   ' In the client area, but below the last Item
  TVHT_ONItemICON = &H2
  TVHT_ONItemLABEL = &H4
  TVHT_ONItemINDENT = &H8
  TVHT_ONItemBUTTON = &H10
  TVHT_ONItemRIGHT = &H20
  TVHT_ONItemSTATEICON = &H40
  TVHT_ONItem = (TVHT_ONItemICON Or TVHT_ONItemLABEL Or TVHT_ONItemSTATEICON)
'****************
'* user-defined
'****************
  TVHT_ONItemLINE = (TVHT_ONItem Or TVHT_ONItemINDENT Or TVHT_ONItemBUTTON Or TVHT_ONItemRIGHT)
  TVHT_ABOVE = &H100
  TVHT_BELOW = &H200
  TVHT_TORIGHT = &H400
  TVHT_TOLEFT = &H800
End Enum
Public Type TVHITTESTINFO   ' was TV_HITTESTINFO
  pt As PointAPI
  Flags As TVHITTESTINFO_Flags
  hItem As Long
End Type
'
'Public Enum TVHITTESTINFO_Flags
'  TVHT_NOWHERE = &H1   ' In the client area, but below the last Item
'  TVHT_ONItemICON = &H2
'  TVHT_ONItemLABEL = &H4
'  TVHT_ONItemINDENT = &H8
'  TVHT_ONItemBUTTON = &H10
'  TVHT_ONItemRIGHT = &H20
'  TVHT_ONItemSTATEICON = &H40
'  TVHT_ONItem = (TVHT_ONItemICON Or TVHT_ONItemLABEL Or TVHT_ONItemSTATEICON)
''****************
''* user-defined
''****************
'  TVHT_ONItemLINE = (TVHT_ONItem Or TVHT_ONItemINDENT Or TVHT_ONItemBUTTON Or TVHT_ONItemRIGHT)
'  TVHT_ABOVE = &H100
'  TVHT_BELOW = &H200
'  TVHT_TORIGHT = &H400
'  TVHT_TOLEFT = &H800
'End Enum
'Public Type NMHDR
'    hwndFrom As Long
'    idfrom As Long
'    code As Long
'End Type
'Public Const WM_SIZE = &H5
'Public Const WM_NOTIFY = &H4E
'Public Const GWL_USERDATA = (-21)
'Public Const GWL_WNDPROC = -4

'******************************************************************
'* TollBar definitions
'******************************************************************
Public Const TB_SETStyle As Long = WM_USER + 56
Public Const TB_GETStyle As Long = WM_USER + 57
Public Const TBStyle_WRAPABLE As Long = &H200  'buttons to wrap when form resized
Public Const TBStyle_FLAT As Long = &H800      'flat IE3+ Style toolbar
Public Const TBStyle_LIST As Long = &H1000     'places captions beside buttons




Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (pDest As Any, ByVal dwLength As Long, ByVal bFill As Byte)
Public Declare Sub ImageList_EndDrag Lib "Comctl32.dll" ()
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowExA Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function GetModuleHandle Lib "Kernel" (ByVal lpModuleName As String) As Integer
Public Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal strUser As String, lngBuffer As Long) As Long
Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal strUser As String, lngBuffer As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function IsChildAPI Lib "user32" Alias "IsChild" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowAPI Lib "user32" Alias "IsWindow" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowEnabledAPI Lib "user32" Alias "IsWindowEnabled" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisibleAPI Lib "user32" Alias "IsWindowVisible" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As WinMsgs, wParam As Any, lParam As Any) As Long
Public Declare Function PtInRect Lib "user32" (lprc As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As WinMsgs, wParam As Any, lParam As Any) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpResultCallBack As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long

Global Handle


'*************************
'* ScrollBar Definitions
'*************************
Public Type SCROLLINFO
  cbSize As Long
  fMask As SIF_Mask
  nMin As Long
  nMax As Long
  nPage As Long
  nPos As Long
  nTrackPos As Long
End Type
Public Enum SIF_Mask
  SIF_RANGE = &H1
  SIF_PAGE = &H2
  SIF_POS = &H4
  SIF_DISABLENOSCROLL = &H8
  SIF_TRACKPOS = &H10
  SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
End Enum



' ==============================================================
' SHBrowseForFolder
Public Type BROWSEINFO
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String ' Return display name of Item selected.
  lpszTitle As String              ' text to go in the banner over the tree.
  ulFlags As Long                 ' Flags that control the return stuff
  lpfn As Long
  lParam As Long      ' extra info that's passed back in callbacks
  iImage As Long      ' output var: where to return the Image index.
End Type
Public Declare Function SHBrowseForFolder Lib "Shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal lItem As Long, ByVal sDir As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

' typedef int (CALLBACK* BFFCALLBACK)(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData) as long

Public Enum BF_Flags
' Browsing for directory.
  BIF_RETURNONLYFSDIRS = &H1      ' For finding a folder to start document searching
  BIF_DONTGOBELOWDOMAIN = &H2     ' For starting the Find Computer
  ' Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if
  ' this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the
  ' rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets
  ' all three lines of text.
  BIF_STATUSTEXT = &H4
  BIF_RETURNFSANCESTORS = &H8

#If (WIN32_IE >= &H400) Then
  BIF_EDITBOX = &H10               ' Add an editbox to the dialog.  Always on with BIF_USENEWUI
  BIF_VALIDATE = &H20              ' insist on valid result (or CANCEL)
  BIF_USENEWUI = &H40              ' Use the new dialog layout with the ability to resize.
#End If  ' // WIN32_IE >= &H400

  BIF_BROWSEFORCOMPUTER = &H1000  ' Browsing for Computers.
  BIF_BROWSEFORPRINTER = &H2000   ' Browsing for Printers
  BIF_BROWSEINCLUDEFILES = &H4000 ' Browsing for Everything
End Enum

' message from browser
Public Enum BFFM_FromDlg
  BFFM_INITIALIZED = 1
  BFFM_SELCHANGED = 2

#If (WIN32_IE >= &H400) Then
' If the user types an invalid name into the edit box, the browse dialog will call the
' application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message.
' This flag is ignored if BIF_EDITBOX is not specified.
  BFFM_VALIDATEFAILEDA = 3     ' lParam:szPath ret:1(cont),0(EndDialog)
  BFFM_VALIDATEFAILEDW = 4     ' lParam:wzPath ret:1(cont),0(EndDialog)
#End If  ' // WIN32_IE >= &H400
End Enum

'   http://www.mvps.org/ccrp
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font
'
' ==============================================================
' A fairly comprehensive wrapping of the IShellFolder and IEnumIDList interfaces with
' some IUnknown thrown in. Also will do about anything that can be done with a pidl...
'
' Note that "IShellFolder Extended Type Library v1.1" (ISHF_Ex.tlb) included with this
' project, must be present and correctly registered on your system, and referenced by
' this project to allow use of these interfaces.
' ==============================================================

' Retrieves a pointer to the shell's IMalloc interface.
' Returns NOERROR if successful or or E_FAIL otherwise.
'Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long

' Retrieves the IShellFolder interface for the desktop folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
'Declare Function SHGetDesktopFolder Lib "shell32" (ppshf As IShellFolder) As Long

' Frees memory allocated by the shell
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' GetItemID Item ID retrieval constants
Public Const GIID_FIRST = 1
Public Const GIID_LAST = -1


' ==============================================================
' SHGetFileInfo

Public Const MAX_PATH = 260

Public Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

' Retrieves information about an object in the file system, such as a file,
' a folder, a directory, or a drive root.

Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'***************
'* TVItem Mask
'***************
Public Enum TVItem_Mask
  TVIF_TEXT = &H1
  TVIF_IMAGE = &H2
  TVIF_PARAM = &H4
  TVIF_State = &H8
  TVIF_HANDLE = &H10
  TVIF_SELECTEDIMAGE = &H20
  TVIF_CHILDREN = &H40

#If (WIN32_IE >= &H400) Then
  TVIF_INTEGRAL = &H80
#End If
  
  TVIF_DI_SETItem = &H1000   ' Notification
End Enum

Public Type TVItemDATA
   isfParent As IShellFolder  ' pointer to the parent Item's shell folder interface
   pidlRel As Long            ' Item's pidl (relative to it's parent Item)
   pidlFQ As Long             ' Item's fully qualified pidl (relative to the desktop)
End Type

Public Type TVItem   ' was TV_Item
  Mask As TVItem_Mask
  hItem As Long
  State As TVItem_State
  stateMask As Long
  pszText As Long   ' pointer
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type
Public Type TVINSERTSTRUCT   ' was TV_INSERTSTRUCT
  hParent As Long
  hInsertAfter As Long
'#If (WIN32_IE >= &H400) Then
'    Union
'    {
'        TVItemEX Itemex;
'        TVItem  Item;
'    } DUMMYUNIONNAME;
'#Else
  Item As TVItem
'#End If
End Type
Public Enum TVItem_State
  TVIS_FOCUSED = &H1   ' no more than one Item
  TVIS_SELECTED = &H2   ' highlight, more than one?!
  TVIS_CUT = &H4
  TVIS_DROPHILITED = &H8
  TVIS_BOLD = &H10
  TVIS_EXPANDED = &H20
  TVIS_EXPANDEDONCE = &H40

#If (WIN32_IE >= &H300) Then
  TVIS_EXPANDPARTIAL = &H80
#End If
  
  TVIS_OVERLAYMask = &HF00
  TVIS_STATEIMAGEMask = &HF000
  TVIS_USERMask = &HF000
End Enum

Public Enum SHGFI_flags
  SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
  SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
  ' Indicates that the function should not attempt to access the file specified by pszPath.
  ' Rather, it should act as if the file specified by pszPath exists with the file attributes
  ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
  ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
  SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
  SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
  SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                        ' containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
  SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
End Enum



' ============================================================================
' window

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2&
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOREPOSITION = &H200
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const SPI_GETWORKAREA = 48&
Public Const GWL_EXStyle = (-20&)
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
' ============================================================================
' window creation
Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As WinStylesEx, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As WinStyles, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Public Enum WinStyles
  WS_OVERLAPPED = &H0
  WS_TABSTOP = &H10000
  WS_MAXIMIZEBOX = &H10000
  WS_MINIMIZEBOX = &H20000
  WS_GROUP = &H20000
  WS_THICKFRAME = &H40000
  WS_SYSMENU = &H80000
  WS_HSCROLL = &H100000
  WS_VSCROLL = &H200000
  WS_DLGFRAME = &H400000
  WS_BORDER = &H800000
  WS_CAPTION = (WS_BORDER Or WS_DLGFRAME)
  WS_MAXIMIZE = &H1000000
  WS_CLIPCHILDREN = &H2000000
  WS_CLIPSIBLINGS = &H4000000
  WS_DISABLED = &H8000000
  WS_VISIBLE = &H10000000
  WS_MINIMIZE = &H20000000
  WS_CHILD = &H40000000
  WS_POPUP = &H80000000
  
  WS_TILED = WS_OVERLAPPED
  WS_ICONIC = WS_MINIMIZE
  WS_SIZEBOX = WS_THICKFRAME
  
  ' Common Window Styles
  WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
  WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
  WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
  WS_CHILDWINDOW = WS_CHILD
End Enum   ' WinStyles

Public Enum WinStylesEx
  WS_EX_DLGMODALFRAME = &H1
  WS_EX_NOPARENTNOTIFY = &H4
  WS_EX_TOPMOST = &H8
  WS_EX_ACCEPTFILES = &H10
  WS_EX_TRANSPARENT = &H20& '&H20
  
  WS_EX_MDICHILD = &H40
  WS_EX_TOOLWINDOW = &H80
  WS_EX_WINDOWEDGE = &H100
  WS_EX_CLIENTEDGE = &H20000 '&H200
  WS_EX_CONTEXTHELP = &H400
  
  WS_EX_RIGHT = &H1000
  WS_EX_LEFT = &H0
  WS_EX_RTLREADING = &H2000
  WS_EX_LTRREADING = &H0
  WS_EX_LEFTSCROLLBAR = &H4000
  WS_EX_RIGHTSCROLLBAR = &H0
  
  WS_EX_CONTROLPARENT = &H10000
  WS_EX_STATICEDGE = &H20000
  WS_EX_APPWINDOW = &H40000
  
  WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
  WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum   ' WinStylesEx
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
' dwFlags
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_Mask = &HFF
' dwLanguageId
Public Const LANG_USER_DEFAULT = &H400&

' ============================================================================
' window messages

Public Const CB_FINDSTRING = &H14C                  ' Used to search a Combo
Public Const LB_FINDSTRING = &H18F                  ' Used to search a List box
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETItemHEIGHT = &H154



' WM_ACTIVATE State values
Public Enum WA_StateValues
  WA_INACTIVE = 0
  WA_ACTIVE = 1
  WA_CLICKACTIVE = 2
End Enum

' Key State Masks for Mouse Messages
Public Enum MouseKeys
  MK_LBUTTON = &H1
  MK_MBUTTON = &H10
  MK_RBUTTON = &H2
  MK_CONTROL = &H8
  MK_SHIFT = &H4
End Enum


Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_CHARFROMPOS& = &HD7
Public Const EM_SETCHARFORMAT = (WM_USER + 68)


'''Public Sub DEFINE_OLEGUID(name As Guid, L As Long, w1 As Integer, w2 As Integer)
'''  DEFINE_GUID name, L, w1, w2, &HC0, 0, 0, 0, 0, 0, 0, &H46
'''End Sub

'''Public Sub DEFINE_GUID(name As Guid, L As Long, w1 As Integer, w2 As Integer, b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
'''  With name
'''    .Data1 = L
'''    .Data2 = w1
'''    .Data3 = w2
'''    .Data4(0) = b0
'''    .Data4(1) = b1
'''    .Data4(2) = b2
'''    .Data4(3) = b3
'''    .Data4(4) = b4
'''    .Data4(5) = b5
'''    .Data4(6) = b6
'''    .Data4(7) = b7
'''  End With
'''End Sub

'''Public Function IID_IShellFolder() As Guid
'''  ' Returns the IShellFolder IID, {000214E6-0000-0000-C000-000000046}
'''  Static iidISF As Guid
'''  If (iidISF.Data1 = 0) Then Call DEFINE_OLEGUID(iidISF, &H214E6, 0, 0)
'''  IID_IShellFolder = iidISF
'''End Function

Public Function GetFileInfo(ByVal pszPath As Variant, uFlags As Long, sfi As SHFILEINFO) As Long
   'Returns the system-defined description of an API error code
   If (VarType(pszPath) = vbString) Then
      ' Must be an explicit path (not a display name).
      GetFileInfo = SHGetFileInfo(CStr(pszPath), 0, sfi, Len(sfi), uFlags)
   Else   ' assume good pidl
      GetFileInfo = SHGetFileInfo(CLng(pszPath), 0, sfi, Len(sfi), uFlags Or SHGFI_PIDL)
   End If
End Function
' Returns a file's SFGAO_ attributes
'   pszPath  - must be either an absolute path or absolute pidl
Public Function GetFileAttribs(ByVal pszPath As Variant) As Long
  Dim sfi As SHFILEINFO
  If GetFileInfo(pszPath, SHGFI_ATTRIBUTES, sfi) Then
    GetFileAttribs = sfi.dwAttributes
  End If
End Function
