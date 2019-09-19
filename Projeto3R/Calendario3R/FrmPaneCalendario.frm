VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#11.2#0"; "Codejock.Calendar.v11.2.2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.TaskPanel.v11.2.2.ocx"
Begin VB.Form FrmPaneCalendario 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7515
   ClientLeft      =   16620
   ClientTop       =   1125
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1320
      ScaleHeight     =   885
      ScaleWidth      =   1530
      TabIndex        =   9
      Top             =   6480
      Width           =   1530
      Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   64
         VisualTheme     =   6
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin VB.PictureBox PctPrev0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1080
      Picture         =   "FrmPaneCalendario.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PctPrev 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   720
      Picture         =   "FrmPaneCalendario.frx":014A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   6480
      Width           =   240
   End
   Begin VB.PictureBox PctReal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      Picture         =   "FrmPaneCalendario.frx":0294
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   6480
      Width           =   240
   End
   Begin VB.PictureBox PrgGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   230
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   6720
      Width           =   1095
   End
   Begin VB.PictureBox PrgYellow 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   230
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   6960
      Width           =   1095
   End
   Begin VB.PictureBox PrgRed 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   3
      Top             =   7200
      Width           =   1095
   End
   Begin XtremeCalendarControl.DatePicker wndDatePicker 
      Height          =   3975
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
      _Version        =   720898
      _ExtentX        =   3836
      _ExtentY        =   7011
      _StockProps     =   64
      AutoSize        =   0   'False
      FirstDayOfWeek  =   1
      ShowNoneButton  =   0   'False
      Show3DBorder    =   0
      RowCount        =   2
      TextNoneButton  =   "Nenhum"
      TextTodayButton =   "Hoje"
      BoldDaysPerIdleStep=   1
   End
   Begin XtremeSuiteControls.GroupBox FraCliente 
      Height          =   1500
      Left            =   3000
      TabIndex        =   11
      Top             =   5280
      Width           =   4215
      _Version        =   720898
      _ExtentX        =   7435
      _ExtentY        =   2646
      _StockProps     =   79
      Caption         =   "Nome "
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin VB.Label LblTel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome "
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label LblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome "
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   465
      End
   End
   Begin XtremeSuiteControls.GroupBox FraResumoDia 
      Height          =   1500
      Left            =   3000
      TabIndex        =   14
      Top             =   3480
      Width           =   4215
      _Version        =   720898
      _ExtentX        =   7435
      _ExtentY        =   2646
      _StockProps     =   79
      Caption         =   "Nome "
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome "
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome "
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   360
         Width           =   1185
      End
   End
   Begin XtremeSuiteControls.GroupBox FraFiltro 
      Height          =   960
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   1693
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox ChkAgendado 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agendados"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   5
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox ChkCancelados 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cancelados"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   5
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox ChkEmEspera 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Excluídos"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   5
         Value           =   1
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption SccCalendario2 
      Height          =   280
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   2655
      _Version        =   720898
      _ExtentX        =   4683
      _ExtentY        =   494
      _StockProps     =   14
      Caption         =   "Meu Calendario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption SccCalendario 
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2640
      _Version        =   720898
      _ExtentX        =   4657
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "Calendario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPaneCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Activate()
Event Load()
Event Resize()
Event PctPrevDblClick()
Event PctPrev0DblClick()
Event PctRealDblClick()
Event PctPrevMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event PctPrev0MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event PctRealMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event LstScheduleItemCheck(ByVal Item As Long)
Event WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Event wndDatePickerDayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
Event ChkEmEsperaClick()
Event ChkCanceladosClick()
Event ChkAgendadoClick()
Private Sub ChkAgendado_Click()
   RaiseEvent ChkAgendadoClick
End Sub
Private Sub ChkCancelados_Click()
   RaiseEvent ChkCanceladosClick
End Sub
Private Sub ChkEmEspera_Click()
   RaiseEvent ChkEmEsperaClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
  RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
'   Me.PrgRed.Orientation = xtpProgressBarVertical
'   Me.PrgYellow.Orientation = xtpProgressBarVertical
'   Me.PrgBlue.Orientation = xtpProgressBarVertical
'   Me.PrgRed.Move 0, Me.Height - Me.PrgRed.Height, 255, 1500
'   Me.PrgYellow.Move 0, Me.PrgRed.Top - Me.PrgYellow.Height, 255, 750
'   Me.PrgBlue.Move 0, Me.PrgYellow.Top - Me.PrgBlue.Height, 255, 450
End Sub
Private Sub LstSchedule_ItemCheck(ByVal Item As Long)
   RaiseEvent LstScheduleItemCheck(Item)
End Sub
Private Sub PctPrev_DblClick()
   RaiseEvent PctPrevDblClick
End Sub
Private Sub PctPrev_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent PctPrevMouseDown(Button, Shift, x, y)
End Sub
Private Sub PctPrev0_DblClick()
   RaiseEvent PctPrev0DblClick
End Sub
Private Sub PctPrev0_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent PctPrev0MouseDown(Button, Shift, x, y)
End Sub
Private Sub PctReal_DblClick()
   RaiseEvent PctRealDblClick
End Sub
Private Sub PctReal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent PctRealMouseDown(Button, Shift, x, y)
End Sub
Private Sub wndDatePicker_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
   RaiseEvent wndDatePickerDayMetrics(Day, Metrics)
End Sub

Private Sub wndDatePicker_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'   If Me.wndDatePicker.HitTestEx(x, y) > Now() Then
'      MsgBox Year(Me.wndDatePicker.HitTestEx(x, y))
'   End If
End Sub

Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   RaiseEvent WndTaskPanelItemClick(Item)
End Sub

