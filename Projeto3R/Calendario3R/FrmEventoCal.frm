VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEventoCal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sem título - Compromisso"
   ClientHeight    =   8040
   ClientLeft      =   5985
   ClientTop       =   1185
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkMeeting 
      Caption         =   "&Exibir Serviços"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   6255
      _Version        =   720898
      _ExtentX        =   11033
      _ExtentY        =   1931
      _StockProps     =   79
      Appearance      =   4
      BorderStyle     =   1
      Begin VB.CommandButton BtnRecurrence 
         Caption         =   "&Recorrência..."
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox ChkAllDayEvent 
         Caption         =   "O &dia inteiro"
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CmbEndTime 
         Height          =   315
         Left            =   3240
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox CmbStartTime 
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin XtremeSuiteControls.DateTimePicker CmbStartDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1695
         _Version        =   720898
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40312.8454282407
      End
      Begin XtremeSuiteControls.DateTimePicker CmbEndDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   1695
         _Version        =   720898
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40312.8454282407
      End
      Begin VB.Label LblDuracao 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Duração : "
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   2880
         TabIndex        =   47
         Top             =   900
         Width           =   1725
      End
      Begin VB.Label lblEndTime 
         AutoSize        =   -1  'True
         Caption         =   "Hora de &Término:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         Caption         =   "Hora de &Início:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   1065
      End
   End
   Begin VB.ComboBox CmbSchedule 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   5295
   End
   Begin XtremeSuiteControls.TabControl TabEvento 
      Height          =   5175
      Left            =   120
      TabIndex        =   42
      Top             =   2280
      Width           =   6375
      _Version        =   720898
      _ExtentX        =   11245
      _ExtentY        =   9128
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.MultiRowJustified=   0   'False
      PaintManager.FixedTabWidth=   80
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Compromisso"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "CmbLabel"
      Item(0).Control(1)=   "CmbShowTimeAs"
      Item(0).Control(2)=   "txtBody"
      Item(0).Control(3)=   "ChkPrivate"
      Item(0).Control(4)=   "CmbReminder"
      Item(0).Control(5)=   "ChkReminder"
      Item(0).Control(6)=   "BtnCustomProperties"
      Item(0).Control(7)=   "LblCategory"
      Item(0).Control(8)=   "lblShowTimeAs"
      Item(1).Caption =   "Serviços"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "GrpSessao"
      Item(2).Caption =   "Tarefas"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "GrdTarefas"
      Begin VB.TextBox txtBody 
         Height          =   4095
         Left            =   -69640
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.CommandButton BtnCustomProperties 
         Caption         =   "Custom &Properties ..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69640
         TabIndex        =   37
         Top             =   4680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox ChkReminder 
         Caption         =   "Lembrete"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -69640
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox CmbReminder 
         Height          =   315
         ItemData        =   "FrmEventoCal.frx":0000
         Left            =   -69640
         List            =   "FrmEventoCal.frx":0002
         TabIndex        =   30
         Text            =   "15 minutos"
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox ChkPrivate 
         Alignment       =   1  'Right Justify
         Caption         =   "&Particular"
         Height          =   255
         Left            =   -65080
         TabIndex        =   35
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox CmbShowTimeAs 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67720
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox CmbLabel 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -65680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin XtremeSuiteControls.GroupBox GrpSessao 
         Height          =   4680
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   6135
         _Version        =   720898
         _ExtentX        =   10821
         _ExtentY        =   8255
         _StockProps     =   79
         Appearance      =   4
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit TxtNOME 
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   3960
            _Version        =   720898
            _ExtentX        =   6985
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin VB.CheckBox ChkFLGAVALIACAO 
            Caption         =   "&Avaliação"
            Height          =   255
            Left            =   1320
            TabIndex        =   48
            Top             =   0
            Width           =   1095
         End
         Begin XtremeSuiteControls.PushButton CmdIDCLIENTE 
            Height          =   315
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "Cadastro de Cliente"
            Top             =   360
            Width           =   855
            _Version        =   720898
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Cliente"
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            TextAlignment   =   10
            ImageAlignment  =   0
            TextImageRelation=   0
         End
         Begin iGrid251_75B4A91C.iGrid GrdSERVICOEVT 
            Height          =   2535
            Left            =   0
            TabIndex        =   28
            Top             =   1680
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   4471
            BorderStyle     =   1
            HighlightBackColorNoFocus=   14737632
         End
         Begin VB.CheckBox ChkFLGCANCELADO 
            Caption         =   "&Cancelado"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1200
            TabIndex        =   27
            Top             =   4320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox ChkFLGCONFIRMADO 
            Caption         =   "&Confirmado"
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   4320
            Width           =   1215
         End
         Begin XtremeSuiteControls.FlatEdit TxtEMAIL 
            Height          =   315
            Left            =   1200
            TabIndex        =   25
            Top             =   1080
            Width           =   4335
            _Version        =   720898
            _ExtentX        =   7646
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtTEL1 
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Top             =   720
            Width           =   1575
            _Version        =   720898
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtTEL2 
            Height          =   315
            Left            =   3840
            TabIndex        =   23
            Top             =   720
            Width           =   1695
            _Version        =   720898
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdLov 
            Height          =   315
            Left            =   5160
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   360
            Width           =   375
            _Version        =   720898
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "FrmEventoCal.frx":0004
         End
         Begin VB.CheckBox ChkFLGREMARCADO 
            Caption         =   "Remarcado"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2520
            TabIndex        =   44
            Top             =   4320
            Visible         =   0   'False
            Width           =   3495
         End
         Begin XtremeSuiteControls.FlatEdit TxtIDCLIENTE 
            Height          =   315
            Left            =   360
            TabIndex        =   46
            Top             =   360
            Visible         =   0   'False
            Width           =   840
            _Version        =   720898
            _ExtentX        =   1482
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            BackColor       =   -2147483633
            Alignment       =   2
            Appearance      =   2
         End
         Begin XtremeSuiteControls.ComboBox CmbNOME 
            Height          =   315
            Left            =   1200
            TabIndex        =   18
            Top             =   360
            Width           =   3960
            _Version        =   720898
            _ExtentX        =   6985
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label LblObsCli 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sessão:01/01 (45 dias)  / Creme: 0 (23/10)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Left            =   1080
            TabIndex        =   45
            Top             =   1395
            Width           =   3090
         End
         Begin VB.Label LbleMail 
            Caption         =   "&e-Mail:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   24
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label LblTel2 
            Caption         =   "Ce&lular:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   720
            Width           =   525
         End
         Begin VB.Label LblTel1 
            Caption         =   "&Telefone:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3000
            TabIndex        =   22
            Top             =   720
            Width           =   675
         End
      End
      Begin iGrid251_75B4A91C.iGrid GrdTarefas 
         Height          =   4575
         Left            =   -69880
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   8070
         BorderStyle     =   1
         HighlightBackColorNoFocus=   14737632
      End
      Begin VB.Label lblShowTimeAs 
         Caption         =   "&Mostrar como:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -67720
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblCategory 
         Caption         =   "Cate&goria:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -65680
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtLocation 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox TxtSubject 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   5295
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   2640
      TabIndex        =   40
      Top             =   7560
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   4920
      TabIndex        =   41
      Top             =   7560
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdAtendimento 
      Height          =   375
      Left            =   360
      TabIndex        =   39
      Top             =   7560
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Atendimento"
      ForeColor       =   32768
      BackColor       =   12648384
      UseVisualStyle  =   -1  'True
   End
   Begin MSComctlLib.ImageList lstImage 
      Left            =   120
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEventoCal.frx":0187
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEventoCal.frx":0521
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEventoCal.frx":067B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEventoCal.frx":1145
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label ctrlColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblLocation 
      Caption         =   "&Local:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   645
      Width           =   855
   End
   Begin VB.Label lblSubject 
      Caption         =   "&Título:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   285
      Width           =   855
   End
End
Attribute VB_Name = "FrmEventoCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Load()
Event Unload(Cancel As Integer)
Event Resize()
Event CmdAtendimentoClick()
Event CmdIDCLIENTEClick()
Event CmbLabelClick()
Event CmbEndTimeChange()
Event CmbEndTimeClick()
Event CmbEndTimeLostFocus()
Event CmbNOMEClick()
Event CmbStartTimeChange()
Event CmbStartTimeClick()
Event CmbStartTimeLostFocus()
Event CmbShowTimeAsClick()
Event CmdOkClick()
Event CmdCancelClick()
Event CmdLovClick()
Event ChkMeetingClick()
Event ChkReminderClick()
Event ChkAllDayEventClick()
Event ChkFLGCANCELADOClick()
Event ChkFLGCONFIRMADOClick()
Event ChkFLGAVALIACAOClick()
Event GrdSERVICOEVTAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdSERVICOEVTBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdSERVICOEVTColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdSERVICOEVTMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdSERVICOEVTLostFocus()
Event GrdSERVICOEVTRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdTarefasDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
Event TxtIDCLIENTEChange()
Event TxtNOMEChange()
Event CmbNOMEChange()
'Event TxtNOMEKeyPress(KeyAscii As Integer)
'Event TxtNOMELostFocus()
Event TxtTEL1Change()
Event TxtTEL1LostFocus()
Event TxtTEL2Change()
Event TxtTEL2LostFocus()
Event BtnCustomPropertiesClick()
Event BtnRecurrenceClick()
Public isLoad  As Boolean
Private Sub BtnCustomProperties_Click()
   RaiseEvent BtnCustomPropertiesClick
End Sub
Private Sub BtnRecurrence_Click()
   RaiseEvent BtnRecurrenceClick
End Sub
Private Sub chkAllDayEvent_Click()
   RaiseEvent ChkAllDayEventClick
End Sub

Private Sub ChkFLGCANCELADO_Click()
   RaiseEvent ChkFLGCANCELADOClick
End Sub
Private Sub ChkFLGCONFIRMADO_Click()
   RaiseEvent ChkFLGCONFIRMADOClick
End Sub
Private Sub ChkMeeting_Click()
   RaiseEvent ChkMeetingClick
End Sub
Private Sub ChkReminder_Click()
   RaiseEvent ChkReminderClick
End Sub
Private Sub CmbEndDate_GotFocus()
   Call SetTag(Me.CmbStartDate, "VALUE", Me.CmbStartDate.Value)
   Call SetTag(Me.CmbEndDate, "VALUE", Me.CmbEndDate.Value)
End Sub
Private Sub CmbEndTime_Change()
   RaiseEvent CmbEndTimeChange
End Sub
Private Sub CmbEndTime_Click()
   RaiseEvent CmbEndTimeClick
End Sub
Private Sub CmbEndTime_LostFocus()
   RaiseEvent CmbEndTimeLostFocus
End Sub
Private Sub CmbLabel_Click()
   RaiseEvent CmbLabelClick
End Sub
Private Sub CmbNOME_Change()
   RaiseEvent CmbNOMEChange
End Sub
Private Sub CmbNOME_Click()
   RaiseEvent CmbNOMEClick
End Sub
Private Sub CmbShowTimeAs_Click()
   RaiseEvent CmbShowTimeAsClick
End Sub
Private Sub CmbStartDate_Change()
   Dim nDif As Long
'   If Me.CmbEndDate.Value < Me.CmbStartDate.Value Then
'      Me.CmbEndDate.Value = DateAdd("n", Val(GetTag(Me.CmbEndDate, "DIFF", 0)), Me.CmbStartDate.Value)
'   End If
   If Me.CmbStartDate.Value <> GetTag(Me.CmbStartDate, "VALUE", "") And GetTag(Me.CmbStartDate, "VALUE", "") <> "" Then
      nDif = DateDiff("D", GetTag(Me.CmbStartDate, "VALUE", ""), Me.CmbStartDate.Value)
      If nDif <> 0 Then
         Me.CmbEndDate.Value = DateAdd("D", nDif, Me.CmbEndDate.Value)
      End If
   End If
   Call SetTag(Me.CmbStartDate, "VALUE", Me.CmbStartDate.Value)
   Call SetTag(Me.CmbEndDate, "VALUE", Me.CmbEndDate.Value)
End Sub
Private Sub CmbStartDate_GotFocus()
   Call SetTag(Me.CmbStartDate, "VALUE", Me.CmbStartDate.Value)
   Call SetTag(Me.CmbEndDate, "VALUE", Me.CmbEndDate.Value)
End Sub
Private Sub CmbStartDate_Validate(Cancel As Boolean)
   If Me.CmbEndDate.Value < Me.CmbStartDate.Value Then
      Me.CmbEndDate.Value = DateAdd("n", Val(GetTag(Me.CmbEndDate, "DIFF", 0)), Me.CmbStartDate.Value)
   End If
End Sub
Private Sub CmbStartTime_Change()
   RaiseEvent CmbStartTimeChange
'   If DateFromString(Me.CmbEndDate.Value, CmbEndTime.Text) < DateFromString(Me.CmbStartDate.Value, Me.CmbStartTime.Text) Then
'      If Me.CmbStartTime.ListIndex = Me.CmbStartTime.ListCount Then
'      Else
'         Me.CmbEndTime.ListIndex = Me.CmbStartTime.ListIndex + 1
'      End If
'   End If
End Sub

Private Sub CmbStartTime_Click()
   RaiseEvent CmbStartTimeClick
End Sub
Private Sub CmbStartTime_LostFocus()
   RaiseEvent CmbStartTimeLostFocus
End Sub
Private Sub CmdAtendimento_Click()
   Me.CmdAtendimento.Enabled = False
   RaiseEvent CmdAtendimentoClick
   Me.CmdAtendimento.Enabled = True
End Sub
Private Sub cmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdIDCLIENTE_Click()
   Me.CmbNOME.UseVisualStyle = True
   Me.CmbNOME.Style = xtpComboSimple
   
   Me.CmdIDCLIENTE.Enabled = False
   RaiseEvent CmdIDCLIENTEClick
   Me.CmdIDCLIENTE.Enabled = True
End Sub

Private Sub CmdIDCLIENTE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.CmdIDCLIENTE.ToolTipText = "Id.: " & Me.TxtIDCLIENTE.Text
End Sub

Private Sub CmdIDCLIENTE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.CmdIDCLIENTE.ToolTipText = "Id.: " & Me.TxtIDCLIENTE.Text
End Sub
Private Sub CmdLov_Click()
   RaiseEvent CmdLovClick
End Sub
Private Sub cmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub ChkFLGAVALIACAO_Click()
   RaiseEvent ChkFLGAVALIACAOClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   isLoad = True
   RaiseEvent Load
   Call SetTag(Me.CmbEndDate, "DIFF", DateDiff("n", DateFromString(Me.CmbStartDate.Value, Me.CmbStartTime.Text), DateFromString(Me.CmbEndDate.Value, Me.CmbEndTime.Text)))
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub Form_Unload(Cancel As Integer)
   
   RaiseEvent Unload(Cancel)
   If Cancel = 0 Then
      isLoad = False
   End If
End Sub
Private Sub GrdSERVICOEVT_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdSERVICOEVTAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdSERVICOEVT_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdSERVICOEVTBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdSERVICOEVT_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdSERVICOEVTColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdSERVICOEVT_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   
   With Me.GrdSERVICOEVT
      .RowMode = (lRow = .RowCount)
      If .RowCount > 0 Then
         For i = 1 To .ColCount
            If .ColVisible(i) Then
               .CellForeColor(.RowCount, i) = IIf(lRow = .RowCount, vbHighlightText, vbGrayText)
               Exit For
            End If
         Next
      End If
   End With
End Sub
Private Sub GrdSERVICOEVT_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdSERVICOEVT.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdSERVICOEVT_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyDelete Then
      If Me.GrdSERVICOEVT.CurRow <> Me.GrdSERVICOEVT.RowCount Then
         Me.GrdSERVICOEVT.RemoveRow Me.GrdSERVICOEVT.CurRow
      End If
   End If
End Sub
Private Sub GrdSERVICOEVT_LostFocus()
   RaiseEvent GrdSERVICOEVTLostFocus
End Sub
Private Sub GrdSERVICOEVT_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   RaiseEvent GrdSERVICOEVTMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdSERVICOEVT_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdSERVICOEVTRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdTarefas_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   RaiseEvent GrdTarefasDblClick(lRow, lCol, bRequestEdit)
End Sub

Private Sub TxtEMAIL_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtIDCLIENTE_Change()
   RaiseEvent TxtIDCLIENTEChange
End Sub
Private Sub TxtLocation_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtNOME_Change()
   RaiseEvent TxtNOMEChange
End Sub
Private Sub TxtNOME_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
'Private Sub TxtNOME_KeyPress(KeyAscii As Integer)
'   RaiseEvent TxtNOMEKeyPress(KeyAscii)
'End Sub
'Private Sub TxtNOME_LostFocus()
'   RaiseEvent TxtNOMELostFocus
'End Sub
Private Sub TxtSubject_Gotfocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtTEL1_Change()
   RaiseEvent TxtTEL1Change
End Sub
Private Sub TxtTEL1_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtTEL1_LostFocus()
   RaiseEvent TxtTEL1LostFocus
End Sub
Private Sub TxtTEL2_Change()
   RaiseEvent TxtTEL2Change
End Sub
Private Sub TxtTEL2_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtTEL2_LostFocus()
   RaiseEvent TxtTEL2LostFocus
End Sub
