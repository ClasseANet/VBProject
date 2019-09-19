VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmVenda 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venda"
   ClientHeight    =   9495
   ClientLeft      =   2595
   ClientTop       =   2760
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmVenda.frx":0000
   ScaleHeight     =   9495
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   750
      Left            =   0
      TabIndex        =   49
      Top             =   -120
      Width           =   12600
      _Version        =   720898
      _ExtentX        =   22225
      _ExtentY        =   1323
      _StockProps     =   79
      Appearance      =   1
      Begin XtremeSuiteControls.PushButton CmdExcluir 
         Height          =   600
         Left            =   2700
         TabIndex        =   50
         ToolTipText     =   "Excluir"
         Top             =   135
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   1058
         _StockProps     =   79
         Caption         =   "Excluir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   10
         Picture         =   "FrmVenda.frx":0258
         ImageAlignment  =   6
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton CmdRecibo 
         Height          =   600
         Left            =   1365
         TabIndex        =   51
         ToolTipText     =   "Excluir"
         Top             =   135
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   1058
         _StockProps     =   79
         Caption         =   "Recibo 000000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   10
         Picture         =   "FrmVenda.frx":0D22
         ImageAlignment  =   6
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton CmdSaldo 
         Height          =   600
         Left            =   0
         TabIndex        =   52
         Top             =   120
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   1058
         _StockProps     =   79
         Caption         =   "&Saldo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   10
         Picture         =   "FrmVenda.frx":10BC
         ImageAlignment  =   6
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton CmdChave 
         Height          =   600
         Left            =   4035
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Abrir Venda"
         Top             =   135
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   1058
         _StockProps     =   79
         Caption         =   "Abrir Venda"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   10
         Picture         =   "FrmVenda.frx":1244
         ImageAlignment  =   6
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton CmdFaturas 
         Height          =   600
         Left            =   5370
         TabIndex        =   54
         Top             =   135
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   1058
         _StockProps     =   79
         Caption         =   "&Faturas"
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
         TextAlignment   =   10
         Picture         =   "FrmVenda.frx":17DE
         ImageAlignment  =   6
         TextImageRelation=   0
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpPedido 
      Height          =   3135
      Left            =   7560
      TabIndex        =   17
      Top             =   1800
      Width           =   4815
      _Version        =   720898
      _ExtentX        =   8493
      _ExtentY        =   5530
      _StockProps     =   79
      Caption         =   " VALORES "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XtremeSuiteControls.FlatEdit TxtVLVENDA 
         Height          =   345
         Left            =   1920
         TabIndex        =   19
         Top             =   2280
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   16711680
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "999,99"
         BackColor       =   12648384
         Alignment       =   1
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLPGTO0 
         Height          =   345
         Left            =   1920
         TabIndex        =   24
         Top             =   3360
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "99,99"
         BackColor       =   14737632
         Alignment       =   1
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLDESC 
         Height          =   345
         Left            =   1920
         TabIndex        =   21
         Top             =   1080
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   192
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "99,99"
         BackColor       =   14737632
         Alignment       =   1
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLTROCO0 
         Height          =   345
         Left            =   1920
         TabIndex        =   26
         Top             =   3960
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "99,99"
         BackColor       =   14737632
         Alignment       =   1
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdFechar 
         Height          =   495
         Left            =   4560
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "&Fechar Pagamento"
         ForeColor       =   32768
         BackColor       =   65280
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
      Begin XtremeSuiteControls.PushButton CmdDesc 
         Height          =   345
         Left            =   3960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   345
         _Version        =   720898
         _ExtentX        =   609
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "..."
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLACRESC 
         Height          =   345
         Left            =   1920
         TabIndex        =   41
         Top             =   1680
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "99,99"
         Alignment       =   1
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLSUBTOTAL 
         Height          =   345
         Left            =   1920
         TabIndex        =   43
         Top             =   480
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "99,99"
         BackColor       =   14737632
         Alignment       =   1
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label LblSubTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label LblAcresc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACRÉSCIMO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   300
         TabIndex        =   42
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label LblTroco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TROCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Label LblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCONTOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label LblPgto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAGAMENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   300
         TabIndex        =   23
         Top             =   3480
         Width           =   1500
      End
      Begin VB.Label LblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   300
         TabIndex        =   18
         Top             =   2400
         Width           =   1500
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpAtend 
      Height          =   3495
      Left            =   1560
      TabIndex        =   37
      Top             =   9720
      Visible         =   0   'False
      Width           =   7335
      _Version        =   720898
      _ExtentX        =   12938
      _ExtentY        =   6165
      _StockProps     =   79
      Caption         =   " Atendimentos Vinculados"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin iGrid251_75B4A91C.iGrid GrdAtendimento 
         Height          =   2535
         Left            =   840
         TabIndex        =   38
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4471
         BorderStyle     =   1
         HighlightBackColorNoFocus=   14737632
      End
      Begin XtremeSuiteControls.PushButton CmdOk2 
         Height          =   375
         Left            =   5400
         TabIndex        =   39
         Top             =   3000
         Width           =   975
         _Version        =   720898
         _ExtentX        =   1720
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   5655
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   7335
      _Version        =   720898
      _ExtentX        =   12938
      _ExtentY        =   9975
      _StockProps     =   79
      Caption         =   " ITENS "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin iGrid251_75B4A91C.iGrid GrdVenda 
         Height          =   4935
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   8705
         BorderStyle     =   1
         HighlightBackColorNoFocus=   14737632
      End
      Begin XtremeSuiteControls.PushButton CmdPacotes 
         Height          =   270
         Left            =   5400
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   5280
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   476
         _StockProps     =   79
         Caption         =   "Adicionar &Pacotes"
         ForeColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDVENDA 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   14737632
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTVENDA 
      Height          =   345
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   609
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin XtremeSuiteControls.GroupBox GrpSessao 
      Height          =   735
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   5055
      _Version        =   720898
      _ExtentX        =   8916
      _ExtentY        =   1296
      _StockProps     =   79
      Appearance      =   4
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit TxtNOME 
         Height          =   345
         Left            =   1200
         TabIndex        =   8
         Top             =   220
         Width           =   3495
         _Version        =   720898
         _ExtentX        =   6165
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
      Begin XtremeSuiteControls.FlatEdit TxtTEL1 
         Height          =   345
         Left            =   5280
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdIDCLIENTE 
         Height          =   345
         Left            =   105
         TabIndex        =   6
         Top             =   220
         Width           =   615
         _Version        =   720898
         _ExtentX        =   1085
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
         TextAlignment   =   6
         ImageAlignment  =   1
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton CmdLovCli 
         Height          =   345
         Left            =   4680
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   220
         Width           =   375
         _Version        =   720898
         _ExtentX        =   661
         _ExtentY        =   609
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmVenda.frx":1AC3
      End
      Begin XtremeSuiteControls.FlatEdit TxtIDCLIENTE 
         Height          =   345
         Left            =   720
         TabIndex        =   7
         Top             =   220
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
      Begin VB.Label LblTel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Celular/Tel."
         Height          =   240
         Left            =   5280
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin iGrid251_75B4A91C.iGrid GrdPagamento 
      Height          =   2175
      Left            =   7560
      TabIndex        =   31
      Top             =   5280
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3836
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin XtremeSuiteControls.PushButton CmdLov 
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisar"
      Top             =   960
      Width           =   345
      _Version        =   720898
      _ExtentX        =   609
      _ExtentY        =   609
      _StockProps     =   79
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmVenda.frx":1C46
   End
   Begin XtremeSuiteControls.PushButton CmdFatura 
      Height          =   240
      Left            =   11640
      TabIndex        =   33
      Top             =   4930
      Visible         =   0   'False
      Width           =   720
      _Version        =   720898
      _ExtentX        =   1270
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "000123"
      ForeColor       =   8421504
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
      TextAlignment   =   1
      Appearance      =   2
      TextImageRelation=   0
   End
   Begin XtremeSuiteControls.ComboBox CmbIDFUNCIONARIO 
      Height          =   315
      Left            =   9120
      TabIndex        =   13
      Top             =   960
      Width           =   3255
      _Version        =   720898
      _ExtentX        =   5741
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
      Style           =   2
      Appearance      =   4
      UseVisualStyle  =   -1  'True
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit TxtOBS 
      Height          =   855
      Left            =   120
      TabIndex        =   29
      Top             =   7800
      Width           =   12375
      _Version        =   720898
      _ExtentX        =   21828
      _ExtentY        =   1508
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      MultiLine       =   -1  'True
      ScrollBars      =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdChave0 
      Height          =   345
      Left            =   1560
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Abrir Venda"
      Top             =   960
      Visible         =   0   'False
      Width           =   345
      _Version        =   720898
      _ExtentX        =   609
      _ExtentY        =   609
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmVenda.frx":1DC9
   End
   Begin XtremeSuiteControls.FlatEdit TxtVLPGTO 
      Height          =   225
      Left            =   10920
      TabIndex        =   45
      Top             =   6720
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   397
      _StockProps     =   77
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "99.999,99"
      BackColor       =   14737632
      Alignment       =   1
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtVLTROCO 
      Height          =   225
      Left            =   10920
      TabIndex        =   46
      Top             =   6960
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   397
      _StockProps     =   77
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "99,99"
      BackColor       =   14737632
      Alignment       =   1
      Locked          =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   735
      Left            =   120
      TabIndex        =   34
      Top             =   8640
      Width           =   12375
      _Version        =   720898
      _ExtentX        =   21828
      _ExtentY        =   1296
      _StockProps     =   79
      Appearance      =   4
      Begin XtremeSuiteControls.PushButton CmdCancel 
         Height          =   375
         Left            =   7920
         TabIndex        =   36
         Top             =   240
         Width           =   1455
         _Version        =   720898
         _ExtentX        =   2566
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
         Height          =   375
         Left            =   10440
         TabIndex        =   35
         Top             =   240
         Width           =   1455
         _Version        =   720898
         _ExtentX        =   2566
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
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   12360
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dinheiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Left            =   9720
      TabIndex        =   48
      Top             =   6720
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Troco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Left            =   9720
      TabIndex        =   47
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label LblFuncionario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Vendedor:"
      Height          =   240
      Left            =   9120
      TabIndex        =   12
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Observação:"
      Height          =   240
      Left            =   120
      TabIndex        =   28
      Top             =   7560
      Width           =   1170
   End
   Begin VB.Label LblFatura 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Verifique pagamento na Fatura: Nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   8640
      TabIndex        =   32
      Top             =   4930
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   12360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label LblVenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Venda"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   600
   End
   Begin VB.Label LblPagamento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGAMENTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   7560
      TabIndex        =   30
      Top             =   5040
      Width           =   1590
   End
   Begin VB.Label LblDTATEND 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   240
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   435
   End
End
Attribute VB_Name = "FrmVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.guaru.com.br/sistemas/document/pdvtef_06.asp
Option Explicit
Event Activate()
Event DblClick()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Unload(Cancel As Integer)
Event Resize()
Event CmdOkClick()
Event CmdCancelClick()
Event CmdDescClick()
Event CmdFaturaClick()
Event CmdFaturasClick()
Event CmdFecharClick()
Event CmdExcluirClick()
Event CmdChaveClick()
Event CmdLovClick()
Event CmdLovCliClick()
Event CmdIDCLIENTEClick()
Event CmdReciboClick()
Event CmdPacotesClick()
Event CmdSaldoClick()

Event GrdVendaAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdVendaBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdVendaColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdVendaColHeaderDblClick(ByVal lCol As Long)
Event GrdVendaMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdVendaLostFocus()
Event GrdVendaRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdVendaValidate(Cancel As Boolean)

Event GrdPagamentoAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdPagamentoBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdPagamentoColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdPagamentoColHeaderDblClick(ByVal lCol As Long)
Event GrdPagamentoMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdPagamentoLostFocus()
Event GrdPagamentoRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdPagamentoValidate(Cancel As Boolean)

Event GrdAtendimentoAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdAtendimentoBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdAtendimentoColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdAtendimentoColHeaderDblClick(ByVal lCol As Long)
Event GrdAtendimentoCustomDrawCell(ByVal lRow As Long, ByVal lCol As Long, ByVal hdc As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal bSelected As Boolean)
Event GrdAtendimentoMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdAtendimentoMouseEnter(ByVal lRow As Long, ByVal lCol As Long)
Event GrdAtendimentoMouseLeave(ByVal lRow As Long, ByVal lCol As Long)
Event GrdAtendimentoMouseDown(ByVal Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean, ByVal bUnderControl As Boolean)
Event GrdAtendimentoLostFocus()
Event GrdAtendimentoRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdAtendimentoValidate(Cancel As Boolean)
Event GrdAtendimentoDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

Event TxtIDVENDAGotFocus()
Event TxtIDVENDALostFocus()
Event TxtIDCLIENTELostFocus()
Event TxtNOMEChange()
Event TxtNOMEKeyPress(KeyAscii As Integer)
Event TxtTEL1Change()
Event TxtTEL1KeyPress(KeyAscii As Integer)
Event TxtTEL1LostFocus()
Event TxtVLACRESCChange()
Event TxtVLDESCChange()
Event TxtVLPGTOChange()
Private Sub CmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdChave_Click()
   RaiseEvent CmdChaveClick
End Sub
Private Sub CmdDesc_Click()
   RaiseEvent CmdDescClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdFatura_Click()
   RaiseEvent CmdFaturaClick
End Sub
Private Sub CmdFaturas_Click()
   RaiseEvent CmdFaturasClick
End Sub
Private Sub CmdFechar_Click()
   RaiseEvent CmdFecharClick
End Sub
Private Sub CmdIDCLIENTE_Click()
   RaiseEvent CmdIDCLIENTEClick
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

Private Sub CmdOk2_Click()
 GrpAtend.Visible = False
End Sub

Private Sub CmdPacotes_Click()
   RaiseEvent CmdPacotesClick
End Sub
Private Sub CmdRecibo_Click()
   RaiseEvent CmdReciboClick
End Sub
Private Sub CmdSaldo_Click()
   RaiseEvent CmdSaldoClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_DblClick()
   RaiseEvent DblClick
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub GrdAtendimento_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendimentoAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdAtendimento_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdAtendimentoBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdAtendimento_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdAtendimentoColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdAtendimento_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdAtendimentoColHeaderDblClick(lCol)
End Sub
Private Sub GrdAtendimento_CustomDrawCell(ByVal lRow As Long, ByVal lCol As Long, ByVal hdc As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal bSelected As Boolean)
   RaiseEvent GrdAtendimentoCustomDrawCell(lRow, lCol, hdc, lLeft, lTop, lRight, lBottom, bSelected)
End Sub
Private Sub GrdAtendimento_LostFocus()
   RaiseEvent GrdAtendimentoLostFocus
End Sub
Private Sub GrdAtendimento_MouseDown(ByVal Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean, ByVal bUnderControl As Boolean)
   RaiseEvent GrdAtendimentoMouseDown(Button, Shift, x, y, lRow, lCol, bDoDefault, bUnderControl)
End Sub
Private Sub GrdAtendimento_MouseEnter(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendimentoMouseEnter(lRow, lCol)
End Sub
Private Sub GrdAtendimento_MouseLeave(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendimentoMouseLeave(lRow, lCol)
End Sub
Private Sub GrdAtendimento_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   RaiseEvent GrdAtendimentoMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdAtendimento_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdAtendimentoRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdAtendimento_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   With Me.GrdAtendimento
      .RowMode = True '(lRow = .RowCount)
      If .RowCount > 0 And lRow > 0 Then
         For i = 1 To .ColCount
            If .ColVisible(i) Then
               If lRow = .RowCount And Mid(.CellValue(.RowCount, 1), 1, 6) = "Clique" Then
                  .CellForeColor(lRow, i) = vbGrayText
               Else
                  .CellForeColor(lRow, i) = vbHighlightText
               End If
               Exit For
            End If
         Next
      End If
   End With
End Sub
Private Sub GrdAtendimento_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   RaiseEvent GrdAtendimentoDblClick(lRow, lCol, bRequestEdit)
End Sub
Private Sub GrdAtendimento_Validate(Cancel As Boolean)
   RaiseEvent GrdAtendimentoValidate(Cancel)
End Sub
Private Sub GrdPagamento_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdPagamentoAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdPagamento_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdPagamentoBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdPagamento_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdPagamentoColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdPagamento_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdPagamentoColHeaderDblClick(lCol)
End Sub
Private Sub GrdPagamento_LostFocus()
   RaiseEvent GrdPagamentoLostFocus
End Sub
Private Sub GrdPagamento_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdPagamentoMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdPagamento_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdPagamentoRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdPagamento_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   
   With Me.GrdPagamento
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
Private Sub GrdPagamento_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdPagamento.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdPagamento_Validate(Cancel As Boolean)
   RaiseEvent GrdPagamentoValidate(Cancel)
End Sub
Private Sub GrdVenda_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdVendaAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdVenda_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdVendaBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdVenda_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdVendaColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdVenda_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdVendaColHeaderDblClick(lCol)
End Sub
Private Sub GrdVenda_LostFocus()
   RaiseEvent GrdVendaLostFocus
End Sub
Private Sub GrdVenda_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdVendaMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdVenda_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdVendaRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdVenda_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   With Me.GrdVenda
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
Private Sub GrdVenda_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdVenda.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdVenda_Validate(Cancel As Boolean)
   RaiseEvent GrdVendaValidate(Cancel)
End Sub
Private Sub TxtIDCLIENTE_LostFocus()
   RaiseEvent TxtIDCLIENTELostFocus
End Sub
Private Sub TxtIDVENDA_GotFocus()
   RaiseEvent TxtIDVENDAGotFocus
End Sub
Private Sub TxtIDVENDA_LostFocus()
   RaiseEvent TxtIDVENDALostFocus
End Sub
Private Sub TxtNOME_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtTEL1_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLACRESC_Change()
   RaiseEvent TxtVLACRESCChange
End Sub
Private Sub TxtVLACRESC_GotFocus()
   If xVal(Me.TxtVLACRESC.Text) = 0 Then Me.TxtVLACRESC.Text = ""
   'Call SelecionarTexto(Me.ActiveControl)
   Me.TxtVLACRESC.SelStart = 0
   Me.TxtVLACRESC.SelLength = Len(Me.TxtVLACRESC.Text)
End Sub
Private Sub TxtVLACRESC_LostFocus()
   Me.TxtVLACRESC.Text = ValBr(Me.TxtVLACRESC.Text)
End Sub
Private Sub TxtVLDESC_Change()
   RaiseEvent TxtVLDESCChange
End Sub
Private Sub TxtVLDESC_GotFocus()
   If xVal(Me.TxtVLDESC.Text) = 0 Then Me.TxtVLDESC.Text = ""
   'Call SelecionarTexto(Me.ActiveControl)
   Me.TxtVLDESC.SelStart = 0
   Me.TxtVLDESC.SelLength = Len(Me.TxtVLDESC.Text)
End Sub
Private Sub TxtVLDESC_LostFocus()
   Me.TxtVLDESC.Text = ValBr(Me.TxtVLDESC.Text)
End Sub
Private Sub TxtVLPGTO_Change()
   'RaiseEvent TxtVLPGTOChange
End Sub
Private Sub TxtVLPGTO_GotFocus()
  Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLPGTO_LostFocus()
   Me.TxtVLPGTO.Text = ValBr(Me.TxtVLPGTO.Text)
End Sub
Private Sub TxtVLTROCO_GotFocus()
  Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLVENDA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtVLVENDA_LostFocus()
   Me.TxtVLVENDA.Text = ValBr(Me.TxtVLVENDA.Text)
End Sub
