VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.TaskPanel.v11.2.2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#11.2#0"; "Codejock.Calendar.v11.2.2.ocx"
Begin VB.Form FrmPaneTarefa 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Contatos"
   ClientHeight    =   7230
   ClientLeft      =   7965
   ClientTop       =   2865
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox FraCalendario 
      Height          =   3240
      Left            =   840
      TabIndex        =   9
      Top             =   3480
      Width           =   2775
      _Version        =   720898
      _ExtentX        =   4895
      _ExtentY        =   5715
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.TabControl TabPeriodo 
         Height          =   3135
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   2640
         _Version        =   720898
         _ExtentX        =   4657
         _ExtentY        =   5530
         _StockProps     =   68
         Appearance      =   2
         Color           =   4
         PaintManager.Layout=   2
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   2
         Item(0).Caption =   "Dia"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "TabControlPage1"
         Item(1).Caption =   "Período"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "TabControlPage2"
         Begin XtremeSuiteControls.TabControlPage TabControlPage2 
            Height          =   2790
            Left            =   -69970
            TabIndex        =   11
            Top             =   315
            Visible         =   0   'False
            Width           =   2580
            _Version        =   720898
            _ExtentX        =   4551
            _ExtentY        =   4921
            _StockProps     =   1
            Page            =   1
            Begin XtremeSuiteControls.PushButton CmdPeriodo 
               Height          =   375
               Left            =   480
               TabIndex        =   18
               Top             =   1920
               Width           =   1215
               _Version        =   720898
               _ExtentX        =   2143
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "&Consultar"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.DateTimePicker CmbDTIni 
               Height          =   375
               Left            =   480
               TabIndex        =   14
               Top             =   600
               Width           =   1335
               _Version        =   720898
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   68
               MinDate         =   40787
               Format          =   1
               CurrentDate     =   42423.7471759259
            End
            Begin XtremeSuiteControls.DateTimePicker CmbDTFim 
               Height          =   375
               Left            =   480
               TabIndex        =   15
               Top             =   1320
               Width           =   1335
               _Version        =   720898
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   68
               MinDate         =   40787
               Format          =   1
               CurrentDate     =   42423.7471759259
            End
            Begin XtremeSuiteControls.Label LblFim 
               Height          =   255
               Left            =   480
               TabIndex        =   17
               Top             =   1080
               Width           =   1215
               _Version        =   720898
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Data Fim"
            End
            Begin XtremeSuiteControls.Label LblIni 
               Height          =   255
               Left            =   480
               TabIndex        =   16
               Top             =   360
               Width           =   1215
               _Version        =   720898
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Data Início"
            End
         End
         Begin XtremeSuiteControls.TabControlPage TabControlPage1 
            Height          =   2790
            Left            =   30
            TabIndex        =   12
            Top             =   315
            Width           =   2580
            _Version        =   720898
            _ExtentX        =   4551
            _ExtentY        =   4921
            _StockProps     =   1
            Page            =   0
            Begin XtremeCalendarControl.DatePicker DpiCalendario 
               Height          =   2640
               Left            =   0
               TabIndex        =   13
               Top             =   120
               Width           =   2535
               _Version        =   720898
               _ExtentX        =   4471
               _ExtentY        =   4657
               _StockProps     =   64
               FirstDayOfWeek  =   1
               Show3DBorder    =   2
               TextNoneButton  =   "Nenhum"
               TextTodayButton =   "Hoje"
            End
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox FraFiltro 
      Height          =   1800
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   3175
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox ChkN100 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1695
         _Version        =   720898
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Não Concluidos"
         UseVisualStyle  =   -1  'True
         Appearance      =   5
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox ChkAndamento 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Em Andamento"
         UseVisualStyle  =   -1  'True
         Appearance      =   5
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox Chk100 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Concluídos"
         UseVisualStyle  =   -1  'True
         Appearance      =   5
      End
      Begin XtremeSuiteControls.CheckBox ChkDelete 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Excluídos"
         UseVisualStyle  =   -1  'True
         Appearance      =   5
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   0
      ScaleHeight     =   2085
      ScaleWidth      =   4650
      TabIndex        =   3
      Top             =   720
      Width           =   4650
      Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
         Height          =   1455
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   2415
         _Version        =   720898
         _ExtentX        =   4260
         _ExtentY        =   2566
         _StockProps     =   64
         VisualTheme     =   6
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   3840
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneTarefa.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneTarefa.frx":01A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneTarefa.frx":033C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneTarefa.frx":03E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneTarefa.frx":0839
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneTarefa.frx":0C8B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   3000
      Top             =   6480
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   2
      VisualTheme     =   2
   End
   Begin XtremeShortcutBar.ShortcutCaption SccConta2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      _Version        =   720898
      _ExtentX        =   8281
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "Minhas Tarefas"
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
   Begin XtremeShortcutBar.ShortcutCaption SccConta 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _Version        =   720898
      _ExtentX        =   8255
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Tarefas"
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
Attribute VB_Name = "FrmPaneTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event Resize()
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CheckClick()
Event CmbDTIniChange()
Event CmbDTFimChange()
Event CmdPeriodoClick()
Event CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Event DpiCalendarioSelectionChanged()
Event DpiCalendarioDayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
Event SccConta2MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event TabPeriodoSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Event WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Event wndTaskPanelFocusedItemChanged()
Private Sub Chk100_Click()
   RaiseEvent CheckClick
End Sub
Private Sub ChkAndamento_Click()
   RaiseEvent CheckClick
End Sub
Private Sub ChkDelete_Click()
   RaiseEvent CheckClick
End Sub
Private Sub ChkN100_Click()
   RaiseEvent CheckClick
End Sub

Private Sub CmbDTFim_Change()
   RaiseEvent CmbDTFimChange
End Sub
Private Sub CmbDTIni_Change()
   RaiseEvent CmbDTIniChange
End Sub

Private Sub CmdPeriodo_Click()
  RaiseEvent CmdPeriodoClick
End Sub
Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   RaiseEvent CommandBarsExecute(Control)
End Sub
Private Sub DpiCalendario_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
   RaiseEvent DpiCalendarioDayMetrics(Day, Metrics)
End Sub

Private Sub DpiCalendario_SelectionChanged()
   RaiseEvent DpiCalendarioSelectionChanged
End Sub
Private Sub Form_Activate()
   Me.CommandBars.DeleteAll
   RaiseEvent Activate
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub SccConta2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent SccConta2MouseDown(Button, Shift, x, y)
End Sub
Private Sub TabPeriodo_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   RaiseEvent TabPeriodoSelectedChanged(Item)
End Sub
Private Sub wndTaskPanel_FocusedItemChanged()
   RaiseEvent wndTaskPanelFocusedItemChanged
End Sub
Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   RaiseEvent WndTaskPanelItemClick(Item)
End Sub
