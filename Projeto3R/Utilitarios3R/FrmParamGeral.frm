VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmParamGeral 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GrpBoxTop 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      _Version        =   720898
      _ExtentX        =   16536
      _ExtentY        =   1508
      _StockProps     =   79
      BackColor       =   16777215
      Begin XtremeSuiteControls.Label LblDSCTitulo 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6855
         _Version        =   720898
         _ExtentX        =   12091
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Parâmetros gerais da aplicação"
         ForeColor       =   8421504
         UseMnemonic     =   0   'False
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblTitulo 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   750
         _Version        =   720898
         _ExtentX        =   1323
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Geral"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpTela 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   9375
      _Version        =   720898
      _ExtentX        =   16536
      _ExtentY        =   9128
      _StockProps     =   79
      Begin XtremeSuiteControls.GroupBox GrpCalendario 
         Height          =   1455
         Left            =   240
         TabIndex        =   29
         Top             =   3600
         Width           =   7935
         _Version        =   720898
         _ExtentX        =   13996
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   " Calendário "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.GroupBox GrpSemanaTrab 
            Height          =   975
            Left            =   1680
            TabIndex        =   35
            Top             =   360
            Width           =   5775
            _Version        =   720898
            _ExtentX        =   10186
            _ExtentY        =   1720
            _StockProps     =   79
            Caption         =   " Semana de Trabalho"
            UseVisualStyle  =   -1  'True
            Begin VB.CheckBox chkWorkDay 
               Caption         =   "Dom"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chkWorkDay 
               Caption         =   "Seg"
               Height          =   195
               Index           =   1
               Left            =   1020
               TabIndex        =   37
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chkWorkDay 
               Caption         =   "Ter"
               Height          =   195
               Index           =   2
               Left            =   1860
               TabIndex        =   38
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chkWorkDay 
               Caption         =   "Qua"
               Height          =   195
               Index           =   3
               Left            =   2700
               TabIndex        =   39
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chkWorkDay 
               Caption         =   "Qui"
               Height          =   195
               Index           =   4
               Left            =   3540
               TabIndex        =   40
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chkWorkDay 
               Caption         =   "Sex"
               Height          =   195
               Index           =   5
               Left            =   4380
               TabIndex        =   41
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chkWorkDay 
               Caption         =   "Sab"
               Height          =   195
               Index           =   6
               Left            =   5100
               TabIndex        =   42
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox CmbDia1Semana 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label lblFirstDayOfWeek 
               Caption         =   "1º Dia da Semana:"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   645
               Width           =   1335
            End
         End
         Begin XtremeSuiteControls.GroupBox GrpExpediente 
            Height          =   975
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   1720
            _StockProps     =   79
            Caption         =   " Expediente "
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.FlatEdit TxtStartTime 
               Height          =   315
               Left            =   600
               TabIndex        =   32
               Top             =   240
               Width           =   690
               _Version        =   720898
               _ExtentX        =   1217
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   16777215
               Text            =   "00:00"
               BackColor       =   16777215
               MaxLength       =   5
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit TxtEndTime 
               Height          =   315
               Left            =   600
               TabIndex        =   34
               Top             =   600
               Width           =   690
               _Version        =   720898
               _ExtentX        =   1217
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   16777215
               Text            =   "00:00"
               BackColor       =   16777215
               MaxLength       =   5
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   600
               Width           =   615
               _Version        =   720898
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Fim: "
               ForeColor       =   0
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   375
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   615
               _Version        =   720898
               _ExtentX        =   1085
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Início: "
               ForeColor       =   0
            End
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpSenhaMestre 
         Height          =   2295
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2775
         _Version        =   720898
         _ExtentX        =   4895
         _ExtentY        =   4048
         _StockProps     =   79
         Caption         =   " Metas Mensais "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkFaixaMeta 
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Visualizar Faixa de Metas"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.FlatEdit TxtFX1 
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   600
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   8421631
            Text            =   "260"
            BackColor       =   8421631
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtFX2 
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Top             =   960
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   8454143
            Text            =   "300"
            BackColor       =   8454143
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtFX3 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Top             =   1320
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   12648384
            Text            =   "350"
            BackColor       =   12648384
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkFXMetaQTD 
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1800
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Visualizar Qtd. Sessões"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   1320
            Width           =   1215
            _Version        =   720898
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Faixa Verde: "
            ForeColor       =   0
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   1215
            _Version        =   720898
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Faixa Amarela: "
            ForeColor       =   0
         End
         Begin XtremeSuiteControls.Label LblSenhaAnt 
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   1215
            _Version        =   720898
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Faixa Vermelha: "
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpCadastroUnico 
         Height          =   1095
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   2775
         _Version        =   720898
         _ExtentX        =   4895
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   " Cadastro Único Entre Lojas "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkOCLIENTE1 
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Clientes"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkOCONTATO1 
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   480
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Contatos"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkRFUNCIONARIO1 
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Funcionários"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpIdentificacao 
         Height          =   855
         Left            =   240
         TabIndex        =   21
         Top             =   2640
         Width           =   2775
         _Version        =   720898
         _ExtentX        =   4895
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   " Identificação "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkBIOMETRIA 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   2535
            _Version        =   720898
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " Uso de Biometria (Hamster DX)"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkPonto 
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   2535
            _Version        =   720898
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " Controle de Ponto"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpSalas 
         Height          =   855
         Left            =   3360
         TabIndex        =   24
         Top             =   2640
         Width           =   2775
         _Version        =   720898
         _ExtentX        =   4895
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   " Exibição de Local de Serviço"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkExibeSala 
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2535
            _Version        =   720898
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " Exibir palavra 'Sala:' em pastas"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpEstruturaServ 
         Height          =   1095
         Left            =   3360
         TabIndex        =   17
         Top             =   1440
         Width           =   2775
         _Version        =   720898
         _ExtentX        =   4895
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   " Estrutura de Serviços "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkTPSERV 
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Serviços"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkTPTRAT 
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tratamentos"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkTPAREA 
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Áreas"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpFoto 
         Height          =   855
         Left            =   6240
         TabIndex        =   26
         Top             =   2640
         Width           =   1935
         _Version        =   720898
         _ExtentX        =   3413
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   " Fotodepilação "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkTPDIR 
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Direção"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkTPDISP 
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   480
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Disparos"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
      End
   End
End
Attribute VB_Name = "FrmParamGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Activate()
Event Resize()
Event Unload()
Event CmdOkSenha()
Event CmdCancelar()
Event CmdPadrao()
Event ChkBackupDiaClick()
Event ChkFaixaMetaClick()
Event ChkPontoClick()
Event ChkTPAREAClick()
Event TxtStartTimeLostFocus()
Event TxtEndTimeLostFocus()
Private Sub CmdCancelar_Click()
   RaiseEvent CmdCancelar
End Sub
Private Sub CmdPadrao_Click()
   RaiseEvent CmdPadrao
End Sub
Private Sub ChkBackupDia_Click()
   RaiseEvent ChkBackupDiaClick
End Sub
Private Sub ChkFaixaMeta_Click()
   RaiseEvent ChkFaixaMetaClick
End Sub

Private Sub ChkPonto_Click()
   RaiseEvent ChkPontoClick
End Sub
Private Sub ChkTPAREA_Click()
   RaiseEvent ChkTPAREAClick
End Sub
Private Sub Form_Activate()
 RaiseEvent Activate
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
Private Sub TxtEndTime_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtEndTime_LostFocus()
   RaiseEvent TxtEndTimeLostFocus
End Sub
Private Sub TxtFX1_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtFX2_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtFX3_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtStartTime_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtStartTime_LostFocus()
   RaiseEvent TxtStartTimeLostFocus
End Sub
