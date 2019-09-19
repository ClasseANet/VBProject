VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProgBar32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
 
Private ProgressStyle As Long
Private mvarValue As Integer
Private mvarMinValue As Integer
Private mvarMaxValue As Integer

Dim ProgBarWnd As Long
Private TempParent As Object
Private Const WM_COMMAND = &H111
Private Const WM_COMMNOTIFY = &H44
Private NoObjectParent As Long
Private Type TagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type
Const ICC_PROGRESS_CLASS = &H20
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As TagInitCommonControlsEx) As Boolean
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SendStringMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Const MF_OWNERDRAW& = &H100&

Const HWND_TOPMOST = -1
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOREDRAW = &H8
Const SWP_SHOWWINDOW = &H40

Private Type tagTBADDBITMAP
        hinst As Long
        nID As Long
End Type

Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNTEXT = 18

' Window Style constants
Const WS_VISIBLE = &H10000000
Const WS_CHILD = &H40000000
Const WS_POPUP = &H80000000

' CreateWindow constants
Const CW_USEDEFAULT = &H80000000

Private Const WM_PAINT = &HF
 
Private Const WM_USER = &H400
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const WS_BORDER = &H800000
Private Const WM_DRAWITEM = &H2B
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WM_SETREDRAW = &HB
'//Common Control Constants
Private Const CCS_TOP = &H1
Private Const CCS_NOMOVEY = &H2
Private Const CCS_BOTTOM = &H3
Private Const CCS_NORESIZE = &H4
Private Const CCS_NOPARENTALIGN = &H8
'Private Const CCS_ADJUSTABLE          0x00000020L
Private Const CCS_NODIVIDER = &H40
 
 
Private Const PROGRESS_CLASSA = "msctls_progress32"

'Style
Private Const PBS_SMOOTH = &H1
Private Const PBS_VERTICAL = &H4
Private Const PBM_SETRANGE = (WM_USER + 1)
Private Const PBM_SETPOS = (WM_USER + 2)
Private Const PBM_DELTAPOS = (WM_USER + 3)
Private Const PBM_SETSTEP = (WM_USER + 4)
Private Const PBM_STEPIT = (WM_USER + 5)
Private Const PBM_SETRANGE32 = (WM_USER + 6)
Private Const PBM_GETRANGE = (WM_USER + 7)
Private Const PBM_GETPOS = (WM_USER + 8)
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR
 

Private Type PPBRange
  iLow As Integer
  iHigh As Integer
End Type
Public Property Get Value() As Integer
   Value = mvarValue
End Property
Public Property Let Value(vData As Integer)
   mvarValue = CInt((vData * 100) / (mvarMaxValue - mvarMinValue))
   Call Me.SetProgBarPos(mvarValue)
End Property
Public Property Get MinValue() As Integer
   MinValue = mvarMinValue
End Property
Public Property Let MinValue(vData As Integer)
   mvarMinValue = vData
End Property
Public Property Get MaxValue() As Integer
   MaxValue = mvarMaxValue
End Property
Public Property Let MaxValue(vData As Integer)
   mvarMaxValue = vData
End Property

Public Function SetBackColor(NewBackColor As Long)
   Call SendMessage(ProgBarWnd, SB_SETBKCOLOR, 0, ByVal NewBackColor)
End Function
'SetBarColor
Public Function SetBarColor(NewBarColor As Long)
Call SendMessage(ProgBarWnd, PBM_SETBARCOLOR, 0, ByVal NewBarColor)
End Function

Public Sub SetProgVert(Vertical As Boolean)
   If Vertical = True Then
      ProgressStyle = PBS_VERTICAL
   Else
      ProgressStyle = 0
   End If
End Sub
Private Sub Class_Initialize()
   Dim iccex As TagInitCommonControlsEx
    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_PROGRESS_CLASS
    End With
    Call InitCommonControlsEx(iccex)
    ProgBarWnd = 0
    mvarValue = 0
    mvarMinValue = 0
    mvarMaxValue = 100
End Sub
Public Function Create(Optional Left As Variant, Optional Top As Variant, _
                       Optional Width As Variant, Optional Height As Variant, _
                       Optional Smooth = True) As Boolean
     
   Dim SmoothVal As Long

   If Smooth = True Then SmoothVal = PBS_SMOOTH
   
   If NoObjectParent <> 0 Then
      ProgBarWnd = CreateWindowEX(0, "msctls_progress32", "", _
             WS_VISIBLE Or WS_CHILD Or ProgressStyle Or SmoothVal, 0, 0, 0, 0, _
             NoObjectParent, 0&, App.hInstance, 0&)
      Call SetParent(ProgBarHwnd, NoObjectParent)
   Else
      If Parent Is Nothing Then
      Create = False
      Exit Function
   End If
   
   If IsMissing(Left) Then Left = 0
   If IsMissing(Top) Then Top = 0
   If IsMissing(Width) Then Width = Parent.Width \ Screen.TwipsPerPixelX
   If IsMissing(Height) Then Height = 20
   
   Left = Left \ Screen.TwipsPerPixelX
   Top = Top \ Screen.TwipsPerPixelY
   
   ProgBarWnd = CreateWindowEX(0, "msctls_progress32", "", _
         WS_VISIBLE Or WS_CHILD Or ProgressStyle Or SmoothVal, 0, 0, 0, 0, _
         Parent.hwnd, 0&, App.hInstance, 0&)
         Call SetParent(ProgBarHwnd, Parent.hwnd)
   End If
   
   Call MoveWindow(ProgBarWnd, CLng(Left), CLng(Top), CLng(Width), CLng(Height), True)
   
   Call ShowWindow(ProgBarWnd, SW_SHOWNORMAL)
   
   
   Create = (ProgBarWnd <> 0)
   
End Function
Public Property Get Parent() As Object
   Set Parent = TempParent
End Property
Public Property Set Parent(Frm As Object)
   Set TempParent = Frm
End Property
Private Sub Class_Terminate()
   Exit Sub
   If ProgBarWnd <> 0 Then
      Call DestroyWindow(ProgBarWnd)
   End If
End Sub
Public Sub DestroyProgBar()
   On Error Resume Next
   If ProgBarWnd <> 0 Then
      Call DestroyWindow(ProgBarWnd)
   End If
End Sub
Public Sub ClearProgBar()
On Error Resume Next
'Set Position to Zero
Call SendMessage(ProgBarWnd, PBM_SETPOS, 0, 0)
End Sub

Public Sub SetProgBarPos(ProgPos As Integer)
   DoEvents
   Call SendMessage(ProgBarWnd, PBM_SETPOS, ProgPos, 0)
   DoEvents
End Sub

Public Sub DelayProgBar(itime As Integer)
   DoEvents
   Call Sleep(itime)
   DoEvents
End Sub
Public Property Get SethWndParent() As Long
   SethWndParent = NoObjectParent
End Property
Public Property Get ProgBarHwnd() As Long
   ProgBarHwnd = ProgBarWnd
End Property
Public Property Let SethWndParent(ByVal vNewValue As Long)
   NoObjectParent = vNewValue
End Property
