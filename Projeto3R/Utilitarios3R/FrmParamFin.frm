VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "CODEJO~1.OCX"
Begin VB.Form FrmParamFin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GrpBoxTop 
      Height          =   855
      Left            =   0
      TabIndex        =   3
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
         TabIndex        =   5
         Top             =   480
         Width           =   6855
         _Version        =   720898
         _ExtentX        =   12091
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Parâmetros do módulo Financeiro."
         ForeColor       =   8421504
         UseMnemonic     =   0   'False
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblTitulo 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1500
         _Version        =   720898
         _ExtentX        =   2646
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Financeiro"
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
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   9375
      _Version        =   720898
      _ExtentX        =   16536
      _ExtentY        =   8493
      _StockProps     =   79
      Begin XtremeSuiteControls.GroupBox GrpExibe 
         Height          =   1215
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   8655
         _Version        =   720898
         _ExtentX        =   15266
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   " Exibição "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkNOMEFAVORECIDO 
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   2895
            _Version        =   720898
            _ExtentX        =   5106
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   " Exibir nome na coluna [Favorecido]"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpIntegra 
         Height          =   1575
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   8655
         _Version        =   720898
         _ExtentX        =   15266
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   " Integração"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkNFE 
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   2895
            _Version        =   720898
            _ExtentX        =   5106
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   " Nota Fiscal Eletrônica (Carioca)"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChKNFE_CLI 
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   5655
            _Version        =   720898
            _ExtentX        =   9975
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Enviar NF-e para Cliente (Não marcar se o envio já é feito pela prefeitura.)"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtUltFat 
            Height          =   315
            Left            =   7560
            TabIndex        =   10
            Top             =   600
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "1"
            BackColor       =   16777215
            Alignment       =   2
            MaxLength       =   2
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkNFE_CPF 
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   960
            Width           =   2535
            _Version        =   720898
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Recibo por CPF"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblUltFat2 
            Height          =   255
            Left            =   7560
            TabIndex        =   18
            Top             =   360
            Width           =   855
            _Version        =   720898
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fat.Enviado"
         End
         Begin XtremeSuiteControls.Label LblUltFat 
            Height          =   255
            Left            =   7560
            TabIndex        =   9
            Top             =   160
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ult. Mês"
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpTaxas 
         Height          =   975
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   8655
         _Version        =   720898
         _ExtentX        =   15266
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   " Taxas do Cartão"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit TxtTXSERV3 
            Height          =   315
            Left            =   1200
            TabIndex        =   15
            Top             =   600
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "2,7"
            BackColor       =   16777215
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtTXSERV2 
            Height          =   315
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "1,7"
            BackColor       =   16777215
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtTXSERV4 
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Top             =   600
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "3,0"
            BackColor       =   16777215
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblParcelado 
            Height          =   375
            Left            =   2160
            TabIndex        =   16
            Top             =   240
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Parcelado"
         End
         Begin XtremeSuiteControls.Label LblCredito 
            Height          =   375
            Left            =   1200
            TabIndex        =   14
            Top             =   240
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Crédito"
         End
         Begin XtremeSuiteControls.Label LblDebito 
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Débito"
         End
      End
   End
End
Attribute VB_Name = "FrmParamFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Activate()
Event Resize()
Event Unload()
Event CmdCancelar()
Event CmdPadrao()
Event LblUltFatDblClick()
Private Sub CmdCancelar_Click()
   RaiseEvent CmdCancelar
End Sub
Private Sub CmdPadrao_Click()
   RaiseEvent CmdPadrao
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
Private Sub LblUltFat_DblClick()
  RaiseEvent LblUltFatDblClick
End Sub
Private Sub LblUltFat2_DblClick()
   RaiseEvent LblUltFatDblClick
End Sub
Private Sub TxtTXSERV2_GotFocus()
   With Me.ActiveControl
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub TxtTXSERV3_GotFocus()
   With Me.ActiveControl
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub TxtTXSERV4_GotFocus()
   With Me.ActiveControl
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub TxtUltFat_GotFocus()
   With Me.ActiveControl
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
