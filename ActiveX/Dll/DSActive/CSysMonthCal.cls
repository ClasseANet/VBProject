VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DS_Calendario"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents mvarMe As FrmMonth
Attribute mvarMe.VB_VarHelpID = -1
Public WithEvents mvarCtrl As VBControlExtender
Attribute mvarCtrl.VB_VarHelpID = -1

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Const ICC_DATE_CLASSES = &H100&

Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Const MONTHCAL_CLASSA = "SysMonthCal32"

Private Const H_MAX As Long = &HFFFF + 1
Private Const DTM_FIRST = &H1000
Private Const DTN_FIRST = (H_MAX - 760&)
Private Const DTN_LAST = (H_MAX - 799&)
Private Const MCM_FIRST = &H1000
Private Const MCM_HITTEST = (MCM_FIRST + 14)
Private Const MCN_FIRST = (H_MAX - 750&)
Private Const MCN_LAST = (H_MAX - 759&)
Private Const MCM_SETRANGE = (MCM_FIRST + 18)
Private Const MCN_SELECT = (MCN_FIRST + 4)
Private Const MCM_GETCURSEL = (MCM_FIRST + 1)
Private Const MCM_GETTODAY = (MCM_FIRST + 13)

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

 
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
 
Private Declare Function SendStringMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
 
 
Private Const DTN_DATETIMECHANGE = (DTN_FIRST + 1)
  
Private MonthCalHwnd    As Long
Private MonthCalParent  As Object
Private mvarData        As String
Private mvarLeft        As Long
Private mvarTop         As Long
Private mvarWidth       As Long
Private mvarHeight      As Long
Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
End Property
Public Property Get Height() As Long
    Height = mvarHeight
End Property
Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
End Property
Public Property Get Width() As Long
    Width = mvarWidth
End Property
Public Property Let Top(ByVal vData As Long)
    mvarTop = vData
End Property
Public Property Get Top() As Long
    Top = mvarTop
End Property
Public Property Let Left(ByVal vData As Long)
    mvarLeft = vData
End Property
Public Property Get Left() As Long
    Left = mvarLeft
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Public Property Let Data(ByVal vData As String)
    mvarData = vData
End Property
Public Property Get Data() As String
    Data = mvarData
End Property
Public Property Get hWnd() As Long
    hWnd = MonthCalHwnd
End Property
Public Function Create(Optional Left As Integer = 5, Optional Top As Integer = 15, _
                       Optional Width As Integer = 190, Optional Height As Integer = 155) As Boolean
   If Parent Is Nothing Then
      Create = False
      Exit Function
   End If
   MonthCalHwnd = CreateWindowEX(0, "SysMonthCal32", "", _
                                 WS_CHILD Or WS_VISIBLE, 0, 0, 0, 0, _
                                 Parent.hWnd, 0, App.hInstance, 0)
   Call ShowWindow(hWnd, SW_SHOWNORMAL)
   Call MoveWindow(MonthCalHwnd, Left, Top, Width, Height, True)
End Function
Public Property Set Parent(frm As Object)
    Set MonthCalParent = frm
End Property
Public Property Get Parent() As Object
    Set Parent = MonthCalParent
End Property
Private Sub Class_Initialize()
   Dim iccex As tagInitCommonControlsEx
   With iccex
      .lngSize = LenB(iccex)
      .lngICC = ICC_DATE_CLASSES
   End With
   Call InitCommonControlsEx(iccex)
   MonthCalHwnd = 0
   Set mvarMe = FrmMonth
   mvarLeft = 0
   mvarTop = 0
End Sub
Private Sub Class_Terminate()
    If MonthCalHwnd <> 0 Then
        Call DestroyWindow(MonthCalHwnd)
    End If
End Sub
Public Function GetCalendarDate() As Date
    Dim systime As SYSTEMTIME
    Dim CalDate As Date
   
    Call SendMessage(MonthCalHwnd, MCM_GETCURSEL, 0, systime)
  
    With systime
       CalDate = DateSerial(.wYear, .wMonth, .wDay)
    End With
    
    GetCalendarDate = CalDate
End Function
Private Sub mvarCtrl_ObjectEvent(Info As EventInfo)
   Dim i As Integer
   i = i
   Select Case Info.Name
      Case "Change"
      Case "Click"
      Case "DblClick"
      Case "KeyDown"
      Case Else
  End Select
End Sub

Private Sub mvarMe_CmdOperClick(index As Integer)
   If index = 0 Then
      mvarData = Format(Me.GetCalendarDate, "dd/mm/yyyy")
   Else
      mvarData = ""
   End If
   mvarMe.Hide
   UnLoad mvarMe
End Sub
Private Sub mvarMe_Load()
  Dim Point As PointAPI
'Written by Ramon Guerrero
'ZoneCorp@dallas.net
'ZoneCorp@Aol.com
'ZoneCOrp@Compuserve.com
   With Me
      Set .Parent = mvarMe
      .Create
      mvarMe.LblDt = Format(Me.GetCalendarDate, "LONG DATE")
   End With
   mvarMe.SubClass mvarMe.hWnd
   Call GetCursorPos(Point)
   If mvarTop = 0 Then
      Point.Y = IIf(Point.Y > 360, 200, Point.Y)
   Else
      Point.Y = (mvarTop / Screen.TwipsPerPixelX) - 18
   End If
   If mvarLeft = 0 Then
      Point.X = IIf(Point.X > 600, 590, Point.X)
   Else
      Point.X = (mvarLeft / Screen.TwipsPerPixelY) + 1
   End If
   mvarMe.Move (Point.X - 1) * Screen.TwipsPerPixelY, (Point.Y + 18) * Screen.TwipsPerPixelX
  ' Set mvarCtrl = New VBControlExtender
   'Set mvarCtrl
End Sub

