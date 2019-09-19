VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmEmpresa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Cadastro de Unidades"
   ClientHeight    =   4080
   ClientLeft      =   2865
   ClientTop       =   2955
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PictBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9855
      TabIndex        =   9
      Top             =   3405
      Width           =   9885
      Begin XtremeSuiteControls.GroupBox GrpBoxBottom 
         Height          =   975
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   9735
         _Version        =   720898
         _ExtentX        =   17171
         _ExtentY        =   1720
         _StockProps     =   79
         Transparent     =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton CmdSair 
            Height          =   375
            Left            =   8400
            TabIndex        =   59
            Top             =   240
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Sai&r"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdExcluir 
            Height          =   375
            Left            =   6600
            TabIndex        =   58
            Top             =   240
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Excluir"
            ForeColor       =   192
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdSalvar 
            Height          =   375
            Left            =   3120
            TabIndex        =   56
            Top             =   240
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Salvar"
            ForeColor       =   12582912
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdNovo 
            Height          =   375
            Left            =   4680
            TabIndex        =   57
            Top             =   240
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Novo"
            ForeColor       =   4210752
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.TabControlPage TabPgBotton 
            Height          =   855
            Left            =   240
            TabIndex        =   60
            Top             =   240
            Width           =   495
            _Version        =   720898
            _ExtentX        =   873
            _ExtentY        =   1508
            _StockProps     =   1
         End
      End
   End
   Begin XtremeSuiteControls.TreeView TreeView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      _Version        =   720898
      _ExtentX        =   4260
      _ExtentY        =   5106
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      LabelEdit       =   1
      Appearance      =   4
   End
   Begin XtremeSuiteControls.TabControl TabEmpresa 
      Height          =   2895
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   6975
      _Version        =   720898
      _ExtentX        =   12303
      _ExtentY        =   5106
      _StockProps     =   68
      Appearance      =   5
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   80
      ItemCount       =   4
      Item(0).Caption =   "COLIGADA"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TbCColigadas"
      Item(1).Caption =   "LOJA"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TbCLoja"
      Item(2).Caption =   "SALA"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TbCSala"
      Item(3).Caption =   "Endereço"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage1"
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   2520
         Left            =   -69970
         TabIndex        =   65
         Top             =   345
         Visible         =   0   'False
         Width           =   6915
         _Version        =   720898
         _ExtentX        =   12197
         _ExtentY        =   4445
         _StockProps     =   1
         Page            =   3
         Begin XtremeSuiteControls.FlatEdit TxtESTADO 
            Height          =   285
            Left            =   6360
            TabIndex        =   37
            Top             =   705
            Width           =   465
            _Version        =   720898
            _ExtentX        =   820
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   2
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtCidade 
            Height          =   285
            Left            =   3840
            TabIndex        =   35
            Top             =   705
            Width           =   2025
            _Version        =   720898
            _ExtentX        =   3572
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtBairro 
            Height          =   285
            Left            =   960
            TabIndex        =   33
            Top             =   705
            Width           =   2145
            _Version        =   720898
            _ExtentX        =   3784
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtEndereco 
            Height          =   300
            Left            =   960
            TabIndex        =   31
            Top             =   225
            Width           =   5865
            _Version        =   720898
            _ExtentX        =   10345
            _ExtentY        =   529
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtCEP 
            Height          =   285
            Left            =   960
            TabIndex        =   39
            Top             =   1065
            Width           =   1185
            _Version        =   720898
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtTELEFONE1 
            Height          =   285
            Left            =   960
            TabIndex        =   41
            Top             =   1440
            Width           =   1185
            _Version        =   720898
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "8888888-"
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtTELEFONE2 
            Height          =   285
            Left            =   4200
            TabIndex        =   43
            Top             =   1440
            Width           =   1185
            _Version        =   720898
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtFromDisplayName 
            Height          =   285
            Left            =   1320
            TabIndex        =   45
            Top             =   1800
            Width           =   2025
            _Version        =   720898
            _ExtentX        =   3572
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtEMAIL 
            Height          =   285
            Left            =   1320
            TabIndex        =   47
            Top             =   2160
            Width           =   5385
            _Version        =   720898
            _ExtentX        =   9499
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label25 
            Height          =   285
            Left            =   120
            TabIndex        =   46
            Top             =   2160
            Width           =   1215
            _Version        =   720898
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Endereço e-Mail:"
         End
         Begin XtremeSuiteControls.Label Label24 
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Conta e-Mail:"
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   1440
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Telefone 1:"
         End
         Begin XtremeSuiteControls.Label Label9 
            Height          =   285
            Left            =   3000
            TabIndex        =   42
            Top             =   1440
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Telefone 2:"
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Top             =   1065
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "C.EP.:"
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   285
            Left            =   6000
            TabIndex        =   36
            Top             =   705
            Width           =   375
            _Version        =   720898
            _ExtentX        =   661
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "UF:"
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   285
            Left            =   3240
            TabIndex        =   34
            Top             =   705
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Cidade:"
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   705
            Width           =   855
            _Version        =   720898
            _ExtentX        =   1508
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Bairro:"
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   225
            Width           =   855
            _Version        =   720898
            _ExtentX        =   1508
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Endereço:"
         End
      End
      Begin XtremeSuiteControls.TabControlPage TbCColigadas 
         Height          =   2520
         Left            =   30
         TabIndex        =   61
         Top             =   345
         Width           =   6915
         _Version        =   720898
         _ExtentX        =   12197
         _ExtentY        =   4445
         _StockProps     =   1
         Page            =   0
         Begin XtremeSuiteControls.FlatEdit TxtIDCOLIGADA 
            Height          =   285
            Left            =   960
            TabIndex        =   3
            Top             =   120
            Width           =   795
            _Version        =   720898
            _ExtentX        =   1402
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            Text            =   "01"
            MaxLength       =   5
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtNMCOLIGADA 
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   480
            Width           =   2745
            _Version        =   720898
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label16 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Nome:"
         End
         Begin XtremeSuiteControls.Label Label15 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Id."
         End
      End
      Begin XtremeSuiteControls.TabControlPage TbCLoja 
         Height          =   2520
         Left            =   -69970
         TabIndex        =   62
         Top             =   345
         Visible         =   0   'False
         Width           =   6915
         _Version        =   720898
         _ExtentX        =   12197
         _ExtentY        =   4445
         _StockProps     =   1
         Page            =   1
         Begin XtremeSuiteControls.FlatEdit TxtCNPJ 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1665
            _Version        =   720898
            _ExtentX        =   2937
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "88.888.888/0001-72"
            MaxLength       =   18
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtNOME 
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            Top             =   120
            Width           =   2745
            _Version        =   720898
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtDTOPERACAO 
            Height          =   285
            Left            =   3960
            TabIndex        =   28
            Top             =   1920
            Width           =   1065
            _Version        =   720898
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "88/88/8888"
            MaxLength       =   14
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkAtivoEmp 
            Height          =   255
            Left            =   5640
            TabIndex        =   8
            Top             =   120
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ativo"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.FlatEdit TxtRazao 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   3315
            _Version        =   720898
            _ExtentX        =   5847
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtRzabrev 
            Height          =   285
            Left            =   5640
            TabIndex        =   22
            Top             =   1320
            Width           =   1065
            _Version        =   720898
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtFantasia 
            Height          =   285
            Left            =   3960
            TabIndex        =   20
            Top             =   1320
            Width           =   1185
            _Version        =   720898
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtInscMunic 
            Height          =   285
            Left            =   2400
            TabIndex        =   13
            Top             =   720
            Width           =   1065
            _Version        =   720898
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtInscEst 
            Height          =   285
            Left            =   3960
            TabIndex        =   15
            Top             =   720
            Width           =   1065
            _Version        =   720898
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtCodServMunic 
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   1920
            Width           =   1305
            _Version        =   720898
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtCodServFederal 
            Height          =   285
            Left            =   2160
            TabIndex        =   26
            Top             =   1920
            Width           =   1305
            _Version        =   720898
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkSimples 
            Height          =   255
            Left            =   5640
            TabIndex        =   29
            Top             =   1920
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Simples"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkMatriz 
            Height          =   255
            Left            =   5640
            TabIndex        =   16
            Top             =   720
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Matriz"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.Label Label23 
            Height          =   285
            Left            =   2160
            TabIndex        =   25
            Top             =   1680
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Cod.Serv.Federal"
         End
         Begin XtremeSuiteControls.Label Label22 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   1680
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Cod. Serv. Munic"
         End
         Begin XtremeSuiteControls.Label Label21 
            Height          =   285
            Left            =   3960
            TabIndex        =   14
            Top             =   480
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Insc. Estadual"
         End
         Begin XtremeSuiteControls.Label Label20 
            Height          =   285
            Left            =   2400
            TabIndex        =   12
            Top             =   480
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Insc. Munic."
         End
         Begin XtremeSuiteControls.Label Label19 
            Height          =   285
            Left            =   3960
            TabIndex        =   19
            Top             =   1080
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Nome Fantasia"
         End
         Begin XtremeSuiteControls.Label Label18 
            Height          =   285
            Left            =   5640
            TabIndex        =   21
            Top             =   1080
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Razão Abrev."
         End
         Begin XtremeSuiteControls.Label Label17 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Razão Social:"
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   285
            Left            =   3960
            TabIndex        =   27
            Top             =   1680
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Inauguração:"
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Unidade:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   285
            Left            =   120
            TabIndex        =   63
            Top             =   480
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "C.N.P.J."
         End
      End
      Begin XtremeSuiteControls.TabControlPage TbCSala 
         Height          =   2520
         Left            =   -69970
         TabIndex        =   64
         Top             =   345
         Visible         =   0   'False
         Width           =   6915
         _Version        =   720898
         _ExtentX        =   12197
         _ExtentY        =   4445
         _StockProps     =   1
         Page            =   2
         Begin XtremeSuiteControls.FlatEdit TxtIDSALA 
            Height          =   285
            Left            =   1320
            TabIndex        =   49
            Top             =   120
            Width           =   795
            _Version        =   720898
            _ExtentX        =   1402
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            Text            =   "01"
            MaxLength       =   5
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkAtivo 
            Height          =   255
            Left            =   5640
            TabIndex        =   55
            Top             =   360
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ativo"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.FlatEdit TxtDTOPERSALA 
            Height          =   285
            Left            =   1320
            TabIndex        =   54
            Top             =   1080
            Width           =   1065
            _Version        =   720898
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "88/88/8888"
            MaxLength       =   14
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtCODSALA 
            Height          =   285
            Left            =   1320
            TabIndex        =   52
            Top             =   600
            Width           =   1035
            _Version        =   720898
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "01"
            MaxLength       =   5
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label14 
            Height          =   285
            Left            =   240
            TabIndex        =   48
            Top             =   120
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Id."
         End
         Begin XtremeSuiteControls.Label Label13 
            Height          =   285
            Left            =   240
            TabIndex        =   53
            Top             =   1080
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Inauguração:"
         End
         Begin XtremeSuiteControls.Label Label12 
            Height          =   285
            Left            =   840
            TabIndex        =   51
            Top             =   600
            Width           =   495
            _Version        =   720898
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Sala"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label11 
            Height          =   285
            Left            =   240
            TabIndex        =   50
            Top             =   600
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Código:"
         End
      End
   End
End
Attribute VB_Name = "FrmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Activate()
Event Load()
Event ChkAtivoEmpClick()
Event CmdExcluirClick()
Event CmdNovoClick()
Event CmdSairClick()
Event CmdSalvarClick()
Event TxtDTOPERACAOLostFocus()
Event TxtDTOPERSALALostFocus()
Event TxtNOMELostFocus()
Event TreeView1NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
Event TreeView1DblClick()
Private Sub ChkAtivoEmp_Click()
   RaiseEvent ChkAtivoEmpClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub TreeView1_DblClick()
  ' RaiseEvent TreeView1DblClick
End Sub
Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   DoEvents
   Me.TreeView1.SelectedItem.Expanded = True
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
   RaiseEvent TreeView1NodeClick(Node)
End Sub
Private Sub TxtDTOPERACAO_LostFocus()
   RaiseEvent TxtDTOPERACAOLostFocus
End Sub
Private Sub TxtDTOPERSALA_LostFocus()
   RaiseEvent TxtDTOPERSALALostFocus
End Sub
Private Sub TxtNOME_LostFocus()
   RaiseEvent TxtNOMELostFocus
End Sub
