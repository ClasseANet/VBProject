Attribute VB_Name = "Declare"
Option Explicit
Public Type PointAPI   ' pt
  x As Long
  y As Long
End Type
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

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As WinMsgs, wParam As Any, lParam As Any) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

