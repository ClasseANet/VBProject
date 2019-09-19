Attribute VB_Name = "mHtmlEditor"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Mewsoft Visual Html Editor
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================
Option Explicit

'----------------------------------------------------------
Public FormIndex As Long
Public TextboxIndex As Long
Public TextareaIndex As Long
Public CheckboxIndex As Long
Public OptionButtonIndex As Long
Public ListBoxIndex As Long
Public DropDownBoxIndex As Long
Public PushButtonIndex As Long
Public HiddenDataIndex As Long
Public PasswordIndex As Long
Public SubmitButtonIndex As Long
Public ResetButtonIndex As Long
Public ImageButtonIndex As Long
Public FileUploadIndex As Long
'----------------------------------------------------------

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Any, ByVal fuRedraw As Long) As Long

'//***************************************************************************************
'// Cursor API
'//***************************************************************************************
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZE = 32640&
Public Const IDC_CROSS = 32515&
Public Const IDC_APPSTARTING = 32650&
Public Const IDC_NO = 32648&
Public Const IDC_WAIT = 32514&

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PointAPI
    X As Long
    y As Long
End Type

Public Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
End Type

Public Type UUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Public Type GUID
     Data1 As Long
     Data2 As Long
     Data3 As Long
     Data4(0 To 7) As Byte
End Type

'//---------------------------------------------------------------------------------------

'//***************************************************************************************
'// Version info
'//***************************************************************************************
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'//--- Platform ID ---
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32s = 0

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
'//---------------------------------------------------------------------------------------

'//--- RedrawWindow flags ---
Public Const RDW_ERASE = &H4
Public Const RDW_ERASENOW = &H200
Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100

Public Const IDM_RESPECTVISIBILITY_INDESIGN = 2405
Public Const IDM_PROTECTMETATAGS = 7101
Public Const IDM_DISABLE_EDITFOCUS_UI = 2404


Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As PointAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......

Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Const R2_BLACK = 1 ' 0
Public Const R2_COPYPEN = 13 ' P
Public Const R2_LAST = 16
Public Const R2_MASKNOTPEN = 3 ' DPna
Public Const R2_MASKPEN = 9 ' DPa
Public Const R2_MASKPENNOT = 5 ' PDna
Public Const R2_MERGENOTPEN = 12    ' DPno
Public Const R2_MERGEPEN = 15 ' DPo
Public Const R2_MERGEPENNOT = 14    ' PDno
Public Const R2_NOP = 11    ' D
Public Const R2_NOT = 6 ' Dn
Public Const R2_NOTCOPYPEN = 4 ' PN
Public Const R2_NOTMASKPEN = 8 ' DPan
Public Const R2_NOTMERGEPEN = 2 ' DPon
Public Const R2_NOTXORPEN = 10 ' DPxn
Public Const R2_WHITE = 16 ' 1
Public Const R2_XORPEN = 7 ' DPx

'Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

'//***************************************************************************************
'// Window API
'//***************************************************************************************
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'//--- Windows messages ---
Public Const WM_USER = &H400
Public Const WM_DDE_FIRST = &H3E0
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_MOUSEMOVE = &H200
Public Const WM_ERASEBKGND = &H14
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_DESTROY = &H2
Public Const WM_LBUTTONUP = &H202
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORSTATIC = &H138
Public Const EM_SETSEL = &HB1
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_NCHITTEST = &H84
Public Const WM_PAINT = &HF
Public Const WM_PRINTCLIENT = &H318
Public Const WM_NCPAINT = &H85
Public Const WM_PRINT = &H317
Public Const WM_ACTIVATE = &H6
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_CLOSE = &H10
Public Const WM_CREATE = &H1
Public Const WM_ENABLE = &HA
Public Const WM_HOTKEY = &H312
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOVING = &H216
Public Const WM_SIZING = &H214

Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_MOUSEWHEEL = &H20A

Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209


Public Const WM_NULL = &H0
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_CTLCOLOR = &H19
Public Const WM_WININICHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_QUERYDRAGICON = &H37

Public Const WM_COMPAREITEM = &H39
Public Const WM_COMPACTING = &H41

Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_KEYFIRST = &H100
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311

Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Public Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
Public Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Public Const WM_COMMNOTIFY = &H44 'no longer suported
Public Const WM_CONVERTREQUESTEX = &H108
Public Const WM_COPYDATA = &H4A
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORMsgBox = &H132
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Public Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Public Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Public Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Public Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Public Const WM_DDE_LAST = (WM_DDE_FIRST + 8)
Public Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Public Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Public Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Public Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Public Const WM_DROPFILES = &H233
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_GETHOTKEY = &H33
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_IME_KEYUP = &H291
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_OTHERWINDOWCREATED = &H42 'no longer suported
Public Const WM_OTHERWINDOWDESTROYED = &H43 'no longer suported
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_POWER = &H48
Public Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
Public Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
Public Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
Public Const WM_PSD_MARGINRECT = (WM_USER + 3)
Public Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
Public Const WM_PSD_PAGESETUPDLG = (WM_USER)
Public Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
Public Const WM_SETHOTKEY = &H32
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_WINDOWPOSCHANGING = &H46

'//--- WM_NCHITTEST return code ---
Public Const HTTRANSPARENT = (-1)
'//---------------------------------------------------------------------------------------

'//***************************************************************************************
'// Drawing API
'//***************************************************************************************
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hrgn As Long) As Long
Public Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As PointAPI) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As PointAPI) As Long
Public Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

'//--- Color codes ---
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14

'====================================================================
'       Subclass the Microsoft Web Browser control
'====================================================================
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
'API parameter constants
Public Const GWL_WNDPROC As Long = -4&
Public Const GW_CHILD As Long = 5&
Public Const API_FAILED As Long = 0&
Public Const API_NULL As Long = 0&
Public Const API_TRUE As Long = 1&
Public Const API_FALSE As Long = 0&

'====================================================================
Public IID_IInternetSecurityManager As GUID
Public IID_IElementBehaviorFactory As GUID
Public IID_IUnknown As GUID

Public Enum ELEMENT_CORNER
    ELEMENT_CORNER_NONE = 0
    ELEMENT_CORNER_TOP = 1
    ELEMENT_CORNER_LEFT = 2
    ELEMENT_CORNER_BOTTOM = 3
    ELEMENT_CORNER_RIGHT = 4
    ELEMENT_CORNER_TOPLEFT = 5
    ELEMENT_CORNER_TOPRIGHT = 6
    ELEMENT_CORNER_BOTTOMLEFT = 7
    ELEMENT_CORNER_BOTTOMRIGHT = 8
End Enum


'====================================================================
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'====================================================================
'               Subclass web browser control
'====================================================================
Public m_lWebBrowserhWnd As Long
Public m_lOriginalWebBrowserWindowProc As Long

Public Sub HookWebBrowser()
    If m_lWebBrowserhWnd = 0 Then Exit Sub
    m_lOriginalWebBrowserWindowProc = SetWindowLong(m_lWebBrowserhWnd, GWL_WNDPROC, AddressOf WebBrowserWindowProc)
End Sub

Public Function HookedWebBrowser() As Boolean
    'Give a way to tell if we are currently subclassed
    HookedWebBrowser = CBool(m_lOriginalWebBrowserWindowProc)
End Function

Public Sub UnHookWebBrowser()
    'Define local variables
    Dim lngReturnValue As Long
    'Reset the window procedure to the original value
    lngReturnValue = SetWindowLong(m_lWebBrowserhWnd, GWL_WNDPROC, m_lOriginalWebBrowserWindowProc)
    'Reset the local window procedure address
    m_lOriginalWebBrowserWindowProc = 0
End Sub

Public Function WebBrowserWindowProc(ByVal WindowHandle As Long, _
                                      ByVal Message As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
    
    'Debug.Print "WebBrowserWindowProc: &H"; Hex$(Message); " &H"; Hex$(wParam); " &H"; Hex$(lParam)
    WebBrowserWindowProc = CallWindowProc(m_lOriginalWebBrowserWindowProc, WindowHandle, Message, wParam, lParam)
    Exit Function
    
    
    Dim ClientRc As RECT
    Dim UpdateRc As RECT
    Dim hBrush As Long
    Dim PrevOrg As PointAPI
    Dim DC  As Long
    
    Select Case Message
        Case WM_DRAWITEM:
            Debug.Print "WM_DRAWITEM"
        Case WM_SIZING:
            Debug.Print "WM_SIZING"
    
'        Case WM_RBUTTONUP, WM_RBUTTONDOWN:
'            'Don't let it see the right mouse button up or down
'            WebBrowserWindowProc = 0
'            Debug.Print "WM_RBUTTONUP"
'            Exit Function
'
'        Case WM_LBUTTONDOWN, WM_LBUTTONUP:
'            Debug.Print "WM_LBUTTONDOWN"
'            WebBrowserWindowProc = 0
        
'        Case WM_CONTEXTMENU:
'            'Debug.Print "WM_CONTEXTMENU"
'            WebBrowserWindowProc = 0
'            Exit Function
        
        Case WM_ERASEBKGND: '&H14
            Debug.Print "WM_ERASEBKGND: "; Message; " "; wParam; " "; lParam
            
'WebBrowserWindowProc = CallWindowProc(m_lOriginalWebBrowserWindowProc, WindowHandle, Message, wParam, lParam)
'Call GetClientRect(WindowHandle, ClientRc)
'hBrush = CreateGridBrush(CLng(5))
'Call GetUpdateRect(WindowHandle, UpdateRc, True)
'Call ValidateRect(WindowHandle, UpdateRc)
''//**** Ask vb wich color to use ****
'Call SendMessage(WindowHandle, WM_CTLCOLORSTATIC, wParam, ByVal WindowHandle)
''//**** Set the brush draw origin to (-1,-1) ****
'Call SetBrushOrgEx(wParam, -1, -1, PrevOrg)
''//**** Swap background and foreground colors ****
'Call SwapBkColors(wParam)
'Call FillRect(wParam, UpdateRc, hBrush)
'DeleteObject hBrush
'RefreshContainer WindowHandle

            WebBrowserWindowProc = 0
            'Debug.Print "UpdateRc: "; UpdateRc.Left, UpdateRc.Top, UpdateRc.Bottom, UpdateRc.Right
            'Debug.Print "ClientRc: "; ClientRc.Left, ClientRc.Top, ClientRc.Bottom, ClientRc.Right
            'WebBrowserWindowProc = CallWindowProc(m_lOriginalWebBrowserWindowProc, WindowHandle, Message, wParam, lParam)
            Exit Function
        
        Case WM_PAINT: '&HF
            Debug.Print "WM_PAINT: &H"; Hex$(wParam); " &H"; Hex$(lParam)
            Dim ps As PAINTSTRUCT
            
            'Call BeginPaint(WindowHandle, ps)
            'Call EndPaint(WindowHandle, ps)
            'WebBrowserWindowProc = 0
            'RefreshContainer WindowHandle
            'Exit Function
            
            WebBrowserWindowProc = CallWindowProc(m_lOriginalWebBrowserWindowProc, WindowHandle, Message, wParam, lParam)
            
'        Case WM_DESTROY:
'        Case WM_RBUTTONDOWN:
'        Case WM_MBUTTONDOWN:
'        Case WM_LBUTTONDOWN:
'        Case WM_LBUTTONUP:
'        Case WM_LBUTTONDBLCLK:
'        Case WM_RBUTTONDBLCLK:
'        Case WM_MBUTTONDBLCLK:
'        Case WM_MOVE, WM_MOVING:
'            Debug.Print "WM_MOVE"
'            WebBrowserWindowProc = CallWindowProc(m_lOriginalWebBrowserWindowProc, WindowHandle, Message, wParam, lParam)
            
'        Case WM_SETCURSOR:
'        Case WM_CTLCOLOREDIT:
        
        Case Else:
            'Pass anything else through unchanged
            WebBrowserWindowProc = CallWindowProc(m_lOriginalWebBrowserWindowProc, WindowHandle, Message, wParam, lParam)
    End Select
    
End Function
'====================================================================

Public Function CreateGridBrush(Size As Long) As Long
    
    Dim nBytes As Long
    nBytes = Int((Size * Size))
    
    '//**** Define pattern bits ****
    Dim bits() As Integer
    ReDim bits(1 To nBytes)
    bits(1) = &H80 '//&H80 = 128 = [1000 0000 0000 0000]
    
    '//**** Create the pattern bitmap ****
    Dim hBmp As Long
    hBmp = CreateBitmap(Size, Size, 1, 1, bits(1))
    If (hBmp = 0) Then Exit Function
    
    '//**** Create a brush from the bitmap ****
    CreateGridBrush = CreatePatternBrush(hBmp)
    
End Function

Public Function SwapBkColors(hdc As Long)
    
    Dim TempBkColor As Long
    
    TempBkColor = GetBkColor(hdc)
    
    Call SetBkColor(hdc, GetTextColor(hdc))
    Call SetTextColor(hdc, TempBkColor)
    
End Function

'====================================================================
Public Sub RefreshContainer(hWnd As Long)
    
    Call RedrawWindow(hWnd, ByVal 0&, ByVal 0&, RDW_ERASE Or RDW_ERASENOW Or RDW_INVALIDATE)
    
End Sub

