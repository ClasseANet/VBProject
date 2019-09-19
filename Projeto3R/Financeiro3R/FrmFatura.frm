VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmFatura 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Fatura"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GrpMemoCalc 
      Height          =   1575
      Left            =   3840
      TabIndex        =   15
      Top             =   1920
      Width           =   5295
      _Version        =   720898
      _ExtentX        =   9340
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   " Observação "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XtremeSuiteControls.FlatEdit TxtHISTORICO 
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5025
         _Version        =   720898
         _ExtentX        =   8864
         _ExtentY        =   2143
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblObs 
         Height          =   285
         Left            =   600
         TabIndex        =   28
         Top             =   480
         Width           =   960
         _Version        =   720898
         _ExtentX        =   1693
         _ExtentY        =   494
         _StockProps     =   79
         Caption         =   "Observação:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTPREV 
      Height          =   315
      Left            =   7320
      TabIndex        =   11
      Top             =   1320
      Width           =   1395
      _Version        =   720898
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDFATURA 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTPREV 
      Height          =   315
      Left            =   7320
      TabIndex        =   14
      Top             =   1320
      Width           =   1695
      _Version        =   720898
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin XtremeSuiteControls.PushButton CmdLov 
      Height          =   345
      Left            =   1200
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisar"
      Top             =   360
      Visible         =   0   'False
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   609
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmFatura.frx":0000
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   29
      Top             =   3840
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Height          =   375
      Left            =   480
      TabIndex        =   17
      ToolTipText     =   "Excluir"
      Top             =   3840
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Excluir"
      ForeColor       =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmFatura.frx":0183
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmFatura.frx":0C4D
   End
   Begin XtremeSuiteControls.FlatEdit TxtValor 
      Height          =   345
      Left            =   7320
      TabIndex        =   9
      Top             =   840
      Width           =   1305
      _Version        =   720898
      _ExtentX        =   2302
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0,00"
      Alignment       =   1
      MaxLength       =   12
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdPagar 
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "&Pagar"
      ForeColor       =   32768
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmFatura.frx":2517
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTEMISSAO 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   1275
      _Version        =   720898
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTEMISSAO 
      Height          =   315
      Left            =   2640
      TabIndex        =   31
      Top             =   360
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin XtremeSuiteControls.GroupBox GrpSessao 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   5895
      _Version        =   720898
      _ExtentX        =   10398
      _ExtentY        =   1085
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit TxtNOME 
         Height          =   345
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   4080
         _Version        =   720898
         _ExtentX        =   7197
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Patricia Moreira"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdLovCli 
         Height          =   345
         Left            =   5520
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
         _Version        =   720898
         _ExtentX        =   661
         _ExtentY        =   609
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmFatura.frx":28B1
      End
      Begin XtremeSuiteControls.PushButton CmdIDCLIENTE 
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   825
         _Version        =   720898
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   0
         Appearance      =   2
         ImageAlignment  =   6
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.FlatEdit TxtIDCLIENTE 
         Height          =   345
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   495
         _Version        =   720898
         _ExtentX        =   873
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "8888"
         BackColor       =   16777215
         Alignment       =   2
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit TxtAtend 
      Height          =   315
      Left            =   2040
      TabIndex        =   21
      Top             =   2160
      Width           =   1680
      _Version        =   720898
      _ExtentX        =   2963
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Text            =   "88/88/8888 88:88"
      MaxLength       =   80
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDATEND 
      Height          =   315
      Left            =   1200
      TabIndex        =   20
      Top             =   2160
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "888888"
      BackColor       =   14737632
      Locked          =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTVENDAP 
      Height          =   315
      Left            =   2040
      TabIndex        =   27
      Top             =   2880
      Width           =   1680
      _Version        =   720898
      _ExtentX        =   2963
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Text            =   "88/88/8888 88:88"
      MaxLength       =   80
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDVENDAP 
      Height          =   315
      Left            =   1200
      TabIndex        =   26
      Top             =   2880
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "888888"
      BackColor       =   14737632
      Locked          =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdAtend 
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   1065
      _Version        =   720898
      _ExtentX        =   1879
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Atendimento:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      TextAlignment   =   0
      Appearance      =   2
      ImageAlignment  =   6
      TextImageRelation=   0
   End
   Begin XtremeSuiteControls.PushButton CmdVendaP 
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   1065
      _Version        =   720898
      _ExtentX        =   1879
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Pagamento: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      TextAlignment   =   0
      Appearance      =   2
      ImageAlignment  =   6
      TextImageRelation=   0
   End
   Begin XtremeSuiteControls.PushButton CmdDividir 
      Height          =   375
      Left            =   8640
      TabIndex        =   34
      ToolTipText     =   "Dividir/Parcelar valor em outra Fatura"
      Top             =   840
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmFatura.frx":2A34
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTVENDA 
      Height          =   315
      Left            =   2040
      TabIndex        =   24
      Top             =   2520
      Width           =   1680
      _Version        =   720898
      _ExtentX        =   2963
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Text            =   "88/88/8888 88:88"
      MaxLength       =   80
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDVENDA 
      Height          =   315
      Left            =   1200
      TabIndex        =   23
      Top             =   2520
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "888888"
      BackColor       =   14737632
      Locked          =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdVenda 
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   1065
      _Version        =   720898
      _ExtentX        =   1879
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Venda: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      TextAlignment   =   0
      Appearance      =   2
      ImageAlignment  =   6
      TextImageRelation=   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   240
      TabIndex        =   35
      Top             =   3600
      Width           =   8775
      _Version        =   720898
      _ExtentX        =   15478
      _ExtentY        =   1508
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label LblSITFAT 
      Height          =   525
      Left            =   6600
      TabIndex        =   33
      Top             =   0
      Width           =   2400
      _Version        =   720898
      _ExtentX        =   4233
      _ExtentY        =   926
      _StockProps     =   79
      Caption         =   "Em Aberto"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   720
      _Version        =   720898
      _ExtentX        =   1270
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Emissão:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblVenc 
      Height          =   285
      Left            =   6360
      TabIndex        =   10
      Top             =   1320
      Width           =   960
      _Version        =   720898
      _ExtentX        =   1693
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Vencimento:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblVALOR 
      Height          =   285
      Left            =   6360
      TabIndex        =   8
      Top             =   840
      Width           =   600
      _Version        =   720898
      _ExtentX        =   1058
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Valor:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label LblVenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº &Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "FrmFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.guaru.com.br/sistemas/document/pdvtef_06.asp
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Unload(Cancel As Integer)
Event Resize()
Event CmdAtendClick()
Event CmbDTPREVChange()
Event CmdOkClick()
Event CmdCancelClick()
Event CmdDividirClick()
Event CmdPagarClick()
Event CmdSalvarClick()
Event CmdVendaClick()
Event CmdVendaPClick()
Event CmdExcluirClick()
Event CmdLovClick()
Event CmdLovCliClick()
Event CmdIDCLIENTEClick()

Event TxtIDFATURAGotFocus()
Event TxtIDFATURALostFocus()
Event TxtIDCLIENTELostFocus()
Event TxtNOMEChange()
Event TxtNOMEKeyPress(KeyAscii As Integer)
Event TxtDTPREVLostFocus()
Private Sub CmbDTPREV_Change()
  RaiseEvent CmbDTPREVChange
End Sub
Private Sub CmdAtend_Click()
   RaiseEvent CmdAtendClick
End Sub
Private Sub CmdDividir_Click()
   RaiseEvent CmdDividirClick
End Sub
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub
Private Sub CmdPagar_Click()
   RaiseEvent CmdPagarClick
End Sub

Private Sub FlatEdit2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub PushButton1_Click()

End Sub

Private Sub CmdVendaP_Click()
   RaiseEvent CmdVendaPClick
End Sub
Private Sub TxtAtend_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtDTPREV_GotFocus()
   If Me.TxtDTPREV.Enabled Then
      Me.TxtDTPREV.SelStart = 0
      Me.TxtDTPREV.SelLength = Len(Me.TxtDTPREV.Text)
      Call SelecionarTexto(Me.TxtDTPREV)
   End If
End Sub
Private Sub TxtDTPREV_LostFocus()
   RaiseEvent TxtDTPREVLostFocus
End Sub
Private Sub CmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdIDCLIENTE_Click()
   Me.CmdIDCLIENTE.Enabled = False
   RaiseEvent CmdIDCLIENTEClick
   Me.CmdIDCLIENTE.Enabled = True
End Sub
Private Sub CmdLov_Click()
   RaiseEvent CmdLovClick
End Sub
Private Sub CmdLovCli_Click()
   RaiseEvent CmdLovCliClick
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
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

Private Sub TxtHISTORICO_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtIDATEND_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtIDCLIENTE_LostFocus()
   RaiseEvent TxtIDCLIENTELostFocus
End Sub
Private Sub TxtIDFATURA_GotFocus()
   RaiseEvent TxtIDFATURAGotFocus
End Sub
Private Sub TxtIDFATURA_LostFocus()
   RaiseEvent TxtIDFATURALostFocus
End Sub
Private Sub TxtNOME_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtValor_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtValor_LostFocus()
   Me.TxtValor.Text = ValBr(Me.TxtValor.Text)
End Sub
Private Sub CmdVenda_Click()
   RaiseEvent CmdVendaClick
End Sub
