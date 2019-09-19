VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#11.2#0"; "Codejock.Calendar.v11.2.2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmFecharDia 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fechamento do Dia"
   ClientHeight    =   8520
   ClientLeft      =   135
   ClientTop       =   1020
   ClientWidth     =   9585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl GrdFecharDia 
      Height          =   7335
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   6495
      _Version        =   720898
      _ExtentX        =   11456
      _ExtentY        =   12938
      _StockProps     =   64
      BorderStyle     =   4
      ShowGroupBox    =   -1  'True
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9000
      Top             =   7800
   End
   Begin XtremeSuiteControls.GroupBox GrpValoresDia 
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2775
      _Version        =   720898
      _ExtentX        =   4895
      _ExtentY        =   7646
      _StockProps     =   79
      Caption         =   " Valores do Dia "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox GrpValoresDia2 
         Height          =   3855
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2535
         _Version        =   720898
         _ExtentX        =   4471
         _ExtentY        =   6800
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.Label LblTotCT 
            Height          =   330
            Left            =   60
            TabIndex        =   21
            Top             =   960
            Width           =   1365
            _Version        =   720898
            _ExtentX        =   2408
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "CARTÃO   : R$"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblVLCaixa 
            Height          =   330
            Left            =   1460
            TabIndex        =   28
            Top             =   2520
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "1.100,00"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblVLTotCT 
            Height          =   330
            Left            =   1460
            TabIndex        =   27
            Top             =   960
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "1.100,00"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LbVLDia 
            Height          =   210
            Left            =   1460
            TabIndex        =   26
            Top             =   2040
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "1.100,00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblVLChe 
            Height          =   210
            Left            =   1460
            TabIndex        =   25
            Top             =   1560
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "1.100,00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblVLDin 
            Height          =   210
            Left            =   1460
            TabIndex        =   24
            Top             =   1320
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "1.100,00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblVLDeb 
            Height          =   210
            Left            =   1460
            TabIndex        =   23
            Top             =   600
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "1.100,00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblVLCred 
            Height          =   210
            Left            =   1460
            TabIndex        =   22
            Top             =   360
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "1.100,00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblCaixa 
            Height          =   330
            Left            =   60
            TabIndex        =   20
            Top             =   2520
            Width           =   1365
            _Version        =   720898
            _ExtentX        =   2408
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "VL. CAIXA: R$"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblDia 
            Height          =   210
            Left            =   60
            TabIndex        =   19
            Top             =   2040
            Width           =   1365
            _Version        =   720898
            _ExtentX        =   2408
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "TOTAL DIA: R$"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   210
            Left            =   60
            TabIndex        =   18
            Top             =   1800
            Width           =   2310
            _Version        =   720898
            _ExtentX        =   4075
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "          ============"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblChe 
            Height          =   210
            Left            =   60
            TabIndex        =   17
            Top             =   1560
            Width           =   1365
            _Version        =   720898
            _ExtentX        =   2408
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "CHEQUE   : R$"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblDin 
            Height          =   210
            Left            =   60
            TabIndex        =   16
            Top             =   1320
            Width           =   1365
            _Version        =   720898
            _ExtentX        =   2408
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "DINHEIRO : R$"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblDeb 
            Height          =   210
            Left            =   60
            TabIndex        =   15
            Top             =   600
            Width           =   1365
            _Version        =   720898
            _ExtentX        =   2408
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "DÉBITO   : R$"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblCred 
            Height          =   210
            Left            =   60
            TabIndex        =   14
            Top             =   360
            Width           =   1365
            _Version        =   720898
            _ExtentX        =   2408
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "CRÉDITO  : R$"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   210
            Left            =   60
            TabIndex        =   13
            Top             =   840
            Width           =   2310
            _Version        =   720898
            _ExtentX        =   4075
            _ExtentY        =   370
            _StockProps     =   79
            Caption         =   "          ============"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpTempo 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _Version        =   720898
      _ExtentX        =   4895
      _ExtentY        =   5318
      _StockProps     =   79
      Caption         =   "Data de Fechamento"
      UseVisualStyle  =   -1  'True
      Appearance      =   4
      Begin XtremeCalendarControl.DatePicker wndDatePicker 
         Height          =   2535
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   2535
         _Version        =   720898
         _ExtentX        =   4471
         _ExtentY        =   4471
         _StockProps     =   64
         AutoSize        =   0   'False
         FirstDayOfWeek  =   1
         ShowNoneButton  =   0   'False
         Show3DBorder    =   0
         TextNoneButton  =   "Nenhum"
         TextTodayButton =   "Hoje"
         BoldDaysOnIdle  =   0   'False
         BoldDaysPerIdleStep=   1
      End
      Begin MSComCtl2.MonthView MvwDTFECHAR 
         Height          =   2310
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         StartOfWeek     =   134021121
         TrailingForeColor=   -2147483632
         CurrentDate     =   41906
      End
      Begin XtremeSuiteControls.DateTimePicker CmbDTFECHAR 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
         _Version        =   720898
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   68
         CustomFormat    =   "dddd ',' dd/MMM"
         Format          =   3
         CurrentDate     =   40356.1843055556
      End
   End
   Begin XtremeSuiteControls.CheckBox ChkErros 
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   7455
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Erros"
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   8280
      Visible         =   0   'False
      Width           =   9015
      _Version        =   720898
      _ExtentX        =   15901
      _ExtentY        =   317
      _StockProps     =   93
      Enabled         =   0   'False
      Scrolling       =   2
      Appearance      =   1
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   7800
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sai&r"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox ChkAvisos 
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   7455
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Avisos"
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox ChkOk 
      Height          =   255
      Left            =   7440
      TabIndex        =   6
      Top             =   7455
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Sem problemas"
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.PushButton CmdFecharDia 
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   7800
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Fechar Dia"
      ForeColor       =   16384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdVerificar 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   7800
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Verificar"
      ForeColor       =   4210752
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.TaskDialog TaskFlood 
      Left            =   6720
      Top             =   7800
      _Version        =   720898
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      WindowTitle     =   "TaskDialog1"
   End
   Begin XtremeSuiteControls.TaskDialog TaskDialog 
      Left            =   720
      Top             =   8520
      _Version        =   720898
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   240
      Top             =   7800
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmFecharDia.frx":0000
   End
End
Attribute VB_Name = "FrmFecharDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bExeFechamento As Boolean
Dim nTotalErros As Integer

Event Load()
Event Activate()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Resize()
Event Unload()
Event CmbDTFECHARChange()
Event wndDatePickerSelectionChanged()
Event CmdVerificarClick()
Event CmdFecharDiaClick()
Event CmdSairClick()
Event ChkExibirClick()
Event GrdFecharDiaColumnClick(ByVal Column As XtremeReportControl.IReportColumn)
Event GrdFecharDiaHyperlinkClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal HyperlinkIndex As Long)
Event MvwDTFECHARDateClick(ByVal DateClicked As Date)
'Event MvwDTFECHARValidate(Cancel As Boolean)
Event TaskDialogButtonClicked(ByVal Id As Long, CloseDialog As Variant)
Event TaskDialogConstructed()
Event TaskFloodConstructed()
Event TaskDialogTimer(ByVal MilliSeconds As Long, Reset As Variant)
Event TaskFloodTimer(ByVal MilliSeconds As Long, Reset As Variant)
Event TimerTimer()
Private Sub ChkAvisos_Click()
   RaiseEvent ChkExibirClick
End Sub
Private Sub ChkErros_Click()
   RaiseEvent ChkExibirClick
End Sub
Private Sub ChkOk_Click()
   RaiseEvent ChkExibirClick
End Sub
Private Sub CmbDTFECHAR_Change()
   RaiseEvent CmbDTFECHARChange
End Sub
Private Sub CmdFecharDia_Click()
   RaiseEvent CmdFecharDiaClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdVerificar_Click()
   RaiseEvent CmdVerificarClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
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
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload
End Sub
Private Sub GrdFecharDia_ColumnClick(ByVal Column As XtremeReportControl.IReportColumn)
   RaiseEvent GrdFecharDiaColumnClick(Column)
End Sub
Private Sub GrdFecharDia_HyperlinkClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal HyperlinkIndex As Long)
   RaiseEvent GrdFecharDiaHyperlinkClick(Row, Item, HyperlinkIndex)
End Sub
Private Sub MvwDTFECHAR_DateClick(ByVal DateClicked As Date)
   RaiseEvent MvwDTFECHARDateClick(DateClicked)
End Sub
Private Sub MvwDTFECHAR_Validate(Cancel As Boolean)
'   RaiseEvent MvwDTFECHARValidate(Cancel)
End Sub
Private Sub taskDialog_ButtonClicked(ByVal Id As Long, CloseDialog As Variant)
   RaiseEvent TaskDialogButtonClicked(Id, CloseDialog)
End Sub
Private Sub TaskDialog_Constructed()
   RaiseEvent TaskDialogConstructed
End Sub
Private Sub TaskDialog_Timer(ByVal MilliSeconds As Long, Reset As Variant)
   RaiseEvent TaskDialogTimer(MilliSeconds, Reset)
End Sub
Private Sub TaskFlood_Constructed()
   RaiseEvent TaskFloodConstructed
End Sub
Private Sub TaskFlood_Timer(ByVal MilliSeconds As Long, Reset As Variant)
   RaiseEvent TaskFloodTimer(MilliSeconds, Reset)
End Sub
Private Sub Timer_Timer()
   RaiseEvent TimerTimer
End Sub
Private Sub wndDatePicker_SelectionChanged()
   RaiseEvent wndDatePickerSelectionChanged
End Sub

