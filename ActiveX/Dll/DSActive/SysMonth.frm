VERSION 5.00
Begin VB.Form FrmMonth 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "SysMonthCal32 Control"
   ClientHeight    =   3090
   ClientLeft      =   5940
   ClientTop       =   3990
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOper 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdOper 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label LblDt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "FrmMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Load()
Event CmdOperClick(index As Integer)

Dim Calendar As DS_Calendario
Private Const H_MAX As Long = &HFFFF + 1
Const DTN_FIRST = (H_MAX - 760&)
Const DTN_DATETIMECHANGE = (DTN_FIRST + 1)
Public Sub ProcMsg(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Result As Long)
   Dim hdrX As NMHDR
   On Error Resume Next
   Select Case uMsg
      Case WM_NOTIFY
         CopyMemory hdrX, ByVal lParam, Len(hdrX)
         'If it's our window then get the date
         
         
         If hdrX.hwndFrom = Calendar.hWnd Or hdrX.code = DTN_DATETIMECHANGE Then
            LblDt = Format(Calendar.GetCalendarDate, "Long Date")
         End If
'         If hdrX.hwndFrom = Me.hWnd Or hdrX.code = DTN_DATETIMECHANGE Then
'            mvarMe.LblDt = Format(Me.GetCalendarDate, "Long Date")
'         End If
   End Select
End Sub
Public Sub SubClass(hWnd As Long)
   On Error Resume Next
   NextProcs = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnSubClass()
   Dim hWndCur As Long
   hWndCur = Me.hWnd
   If NextProcs Then
      SetWindowLong hWndCur, GWL_WNDPROC, NextProcs
      NextProcs = 0
   End If
End Sub
Private Sub CmdOper_Click(index As Integer)
   RaiseEvent CmdOperClick(index)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Unload(Cancel As Integer)
   UnSubClass
End Sub
