VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProgBar32"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Private ProgressStyle As Long
Private mvarValue As Long
Private mvarMinValue As Integer
Private mvarMaxValue As Integer

Private mvarLeft As Long
Private mvarTop As Long
Private mvarWidth As Long
Private mvarHeight As Long

Dim ProgBarWnd As Long
Dim LblWnd As Long

Private TempParent As Object
Private Const WM_COMMAND = &H111
Private Const WM_COMMNOTIFY = &H44
Private NoObjectParent As Long
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type
Const ICC_LISTVIEW_CLASSES = &H1       ' listview, header
Const ICC_TREEVIEW_CLASSES = &H2       ' treeview, tooltips
Const ICC_BAR_CLASSES = &H4            ' toolbar, statusbar, trackbar, tooltips
Const ICC_TAB_CLASSES = &H8            ' tab, tooltips
Const ICC_UPDOWN_CLASS = &H10          ' updown
Const ICC_PROGRESS_CLASS = &H20        ' progress
Const ICC_HOTKEY_CLASS = &H40          ' hotkey
Const ICC_ANIMATE_CLASS = &H80         ' animate
Const ICC_WIN95_CLASSES = &HFF
Const ICC_DATE_CLASSES = &H100         ' month picker, date picker, time picker, updown
Const ICC_USEREX_CLASSES = &H200       ' comboex
Const ICC_COOL_CLASSES = &H400         ' rebar (coolbar) control
Const ICC_INTERNET_CLASSES = &H800
Const ICC_PAGESCROLLER_CLASS = &H1000  ' page scroller
Const ICC_NATIVEFNTCTL_CLASS = &H2000  ' native font control

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SendStringMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
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
 
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_Style = (-16)
Private Const WS_BORDER = &H800000
Private Const WM_DRAWItem = &H2B
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
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

Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    Style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type

Public Property Get Value() As Long
   Value = mvarValue
End Property
Public Property Let Value(vData As Long)
   mvarValue = CInt((vData * 100) / ((mvarMaxValue - mvarMinValue) + 1))
   Call Me.SetProgBarPos(CInt(mvarValue))
   DoEvents
'   Me.Parent.CurrentX = 40
'   Me.Parent.Print Space(82) & CStr(mvarValue) & "%"
'   Me.Parent.Refresh
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
   Dim iccex As tagInitCommonControlsEx
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
Public Function Create(Optional Left As Variant, Optional Top As Variant, Optional Width As Variant, Optional Height As Variant, _
                       Optional Smooth = True) As Boolean
     
   Dim SmoothVal As Long
   Dim lStyle As Long
   Dim CS As CREATESTRUCT

   If Smooth = True Then
      SmoothVal = PBS_SMOOTH
   End If
   
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
      If IsMissing(Width) Then Width = Parent.Width
      If IsMissing(Height) Then Height = 20
      
      mvarLeft = Left
      mvarTop = Top
      mvarWidth = Width
      mvarHeight = Height
   
      Left = Left \ Screen.TwipsPerPixelX
      Top = Top \ Screen.TwipsPerPixelY
      Width = Width \ Screen.TwipsPerPixelX
   
      lStyle = WS_VISIBLE Or WS_CHILD Or ProgressStyle Or SmoothVal
      ProgBarWnd = CreateWindowEX(0, "msctls_progress32", "", lStyle, 0, 0, 0, 0, _
                   Parent.hWnd, 0&, App.hInstance, 0&)
      
'      lStyle = WS_VISIBLE Or WS_CHILD
'      LblWnd = CreateWindowEX(WS_EX_STATICEDGE Or WS_EX_TRANSPARENT, "STATIC", "100%", lStyle, 0, 0, 0, 0, _
'             Parent.hWnd, 0&, App.hInstance, 0&)  'CS)
'
'      Call SetParent(LblWnd, Parent.hWnd)
      Call SetParent(ProgBarHwnd, Parent.hWnd)
   End If
   
'   Call MoveWindow(LblWnd, CLng(Left) - 40, CLng(Top), 40, CLng(Height), True)
'   Call ShowWindow(LblWnd, SW_SHOWNORMAL)
   
   Call MoveWindow(ProgBarWnd, CLng(Left), CLng(Top), CLng(Width), CLng(Height), True)
   Call ShowWindow(ProgBarWnd, SW_SHOWNORMAL)
   
   Call SetBarColor(QBColor(3))
   
   Create = (ProgBarWnd <> 0)
   
   
'   Me.Parent.CurrentY = 400
'   Me.Parent.FontSize = 8
'   For i = 0 To 38
'      Me.Parent.CurrentX = i
'      If i = 38 Then
'         Me.Parent.Print Space(2 * (Height + Top) / 13) & "40%"
'      Else
'         Me.Parent.Print "  "
'      End If
'      Me.Parent.Refresh
'   Next
Dim n As Variant
   On Error Resume Next
'   Call ClsOffice.SetContained(Parent)
   For Each n In Parent.Controls
      If n.Container.Name = Parent.Name Then
         If (n.Top \ Screen.TwipsPerPixelX) >= Top Then
            If n.Visible Then
               Call ClsCtrl.SetTag(n, "VISIBLE", n.Visible)
               n.Visible = False
            End If
         End If
      End If
   Next
End Function
Public Property Get Parent() As Object
   Set Parent = TempParent
End Property
Public Property Set Parent(frm As Object)
   Set TempParent = frm
End Property
Private Sub Class_Terminate()
   DestroyProgBar
End Sub
Public Sub DestroyProgBar()
   On Error Resume Next
   Dim n As Variant
   
   If ProgBarWnd <> 0 Then
      Call DestroyWindow(ProgBarWnd)
   End If
   If LblWnd <> 0 Then
      Call DestroyWindow(LblWnd)
   End If
   
   For Each n In Parent.Controls
      If n.Container.Name = Parent.Name Then
         If ClsCtrl.GetTag(n, "VISIBLE") <> "" Then
            n.Visible = ClsCtrl.GetTag(n, "VISIBLE")
         End If
      End If
   Next
End Sub
Public Sub ClearProgBar()
   On Error Resume Next
   'Set Position to Zero
   Call SendMessage(ProgBarWnd, PBM_SETPOS, 0, 0)
End Sub

Public Sub SetProgBarPos(ProgPos As Integer)
   DoEvents
   Call SendMessage(ProgBarWnd, PBM_SETPOS, ProgPos, 0)
'   Call SendMessage(LblWnd, PBM_SETSTEP, ProgPos, 0)
   DoEvents
'   Dim lStyle As Long
'   If LblWnd <> 0 Then Call DestroyWindow(LblWnd)
'   lStyle = WS_VISIBLE Or WS_CHILD
'   LblWnd = CreateWindowEX(WS_EX_STATICEDGE Or WS_EX_TRANSPARENT, "STATIC", CStr(ProgPos) & "%", lStyle, 0, 0, 0, 0, _
'            Parent.hWnd, 0&, App.hInstance, 0&)  'CS)
'   Call MoveWindow(LblWnd, CLng(mvarLeft) - 40, CLng(mvarTop), 40, CLng(mvarHeight), True)
'   Call ShowWindow(LblWnd, SW_SHOWNORMAL)
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
