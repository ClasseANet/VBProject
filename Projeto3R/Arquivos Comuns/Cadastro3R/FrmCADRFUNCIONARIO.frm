VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "CODEJO~1.OCX"
Begin VB.Form FrmCADRFUNCIONARIO 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Funcionário"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.FlatEdit txtID 
      Height          =   315
      Left            =   690
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483637
      Enabled         =   0   'False
      BackColor       =   -2147483637
      Locked          =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdLov 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADRFUNCIONARIO.frx":0000
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   62
      Top             =   5880
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Excluir"
      ForeColor       =   64
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADRFUNCIONARIO.frx":0183
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   5760
      TabIndex        =   58
      Top             =   5880
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Sair"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtNome 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   645
      Width           =   3960
      _Version        =   720898
      _ExtentX        =   6985
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   50
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdNovo 
      Height          =   375
      Left            =   1560
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Novo"
      ForeColor       =   32768
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADRFUNCIONARIO.frx":0C4D
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   4560
      TabIndex        =   59
      Top             =   5880
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   32768
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADRFUNCIONARIO.frx":0DA7
   End
   Begin XtremeSuiteControls.RadioButton OptATIVO 
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inativo"
      ForeColor       =   192
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
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton OptATIVO 
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ativo"
      ForeColor       =   12582912
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
      UseVisualStyle  =   -1  'True
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl TabContato 
      Height          =   4455
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   6735
      _Version        =   720898
      _ExtentX        =   11880
      _ExtentY        =   7858
      _StockProps     =   68
      AutoResizeClient=   0   'False
      Appearance      =   2
      Color           =   16
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.MinTabWidth=   70
      ItemCount       =   3
      Item(0).Caption =   "Principal"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Endereço"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "Valores"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   4080
         Left            =   -69970
         TabIndex        =   63
         Top             =   345
         Visible         =   0   'False
         Width           =   6675
         _Version        =   720898
         _ExtentX        =   11774
         _ExtentY        =   7197
         _StockProps     =   1
         Page            =   1
         Begin XtremeSuiteControls.FlatEdit txtEndereco 
            Height          =   945
            Left            =   945
            TabIndex        =   25
            Top             =   240
            Width           =   5040
            _Version        =   720898
            _ExtentX        =   8890
            _ExtentY        =   1667
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtBairro 
            Height          =   330
            Left            =   945
            TabIndex        =   27
            Top             =   1320
            Width           =   3360
            _Version        =   720898
            _ExtentX        =   5927
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCidade 
            Height          =   330
            Left            =   945
            TabIndex        =   29
            Top             =   1800
            Width           =   5040
            _Version        =   720898
            _ExtentX        =   8890
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbPais 
            Height          =   315
            Left            =   5160
            TabIndex        =   33
            Top             =   1320
            Visible         =   0   'False
            Width           =   825
            _Version        =   720898
            _ExtentX        =   1455
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCEP 
            Height          =   330
            Left            =   2520
            TabIndex        =   35
            Top             =   2280
            Width           =   1320
            _Version        =   720898
            _ExtentX        =   2328
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   9
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtEstado 
            Height          =   330
            Left            =   960
            TabIndex        =   31
            Top             =   2280
            Width           =   600
            _Version        =   720898
            _ExtentX        =   1058
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   2
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GrpContato 
            Height          =   1335
            Left            =   120
            TabIndex        =   36
            Top             =   2640
            Width           =   5895
            _Version        =   720898
            _ExtentX        =   10398
            _ExtentY        =   2355
            _StockProps     =   79
            Caption         =   "Contatos"
            UseVisualStyle  =   -1  'True
            Appearance      =   1
            Begin XtremeSuiteControls.FlatEdit txtEmail 
               Height          =   285
               Left            =   1200
               TabIndex        =   40
               Top             =   600
               Width           =   3435
               _Version        =   720898
               _ExtentX        =   6059
               _ExtentY        =   503
               _StockProps     =   77
               BackColor       =   -2147483643
               MaxLength       =   50
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit txtCelular 
               Height          =   285
               Left            =   1200
               TabIndex        =   42
               Top             =   960
               Width           =   1905
               _Version        =   720898
               _ExtentX        =   3360
               _ExtentY        =   503
               _StockProps     =   77
               BackColor       =   -2147483643
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit txtTel 
               Height          =   285
               Left            =   1200
               TabIndex        =   38
               Top             =   240
               Width           =   1905
               _Version        =   720898
               _ExtentX        =   3360
               _ExtentY        =   503
               _StockProps     =   77
               BackColor       =   -2147483643
               MaxLength       =   20
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   41
               Top             =   960
               Width           =   525
               _Version        =   720898
               _ExtentX        =   926
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Celular:"
               AutoSize        =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   330
               Index           =   5
               Left            =   240
               TabIndex        =   39
               Top             =   600
               Width           =   465
               _Version        =   720898
               _ExtentX        =   820
               _ExtentY        =   582
               _StockProps     =   79
               Caption         =   "Email:"
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   330
               Index           =   0
               Left            =   240
               TabIndex        =   37
               Top             =   240
               Width           =   720
               _Version        =   720898
               _ExtentX        =   1270
               _ExtentY        =   582
               _StockProps     =   79
               Caption         =   "Telefone:"
            End
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Endereço:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   7
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Bairro:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   8
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   600
            _Version        =   720898
            _ExtentX        =   1058
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Cidade:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   9
            Left            =   120
            TabIndex        =   30
            Top             =   2280
            Width           =   645
            _Version        =   720898
            _ExtentX        =   1138
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Estado:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   10
            Left            =   4560
            TabIndex        =   32
            Top             =   1320
            Visible         =   0   'False
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "País:"
            Enabled         =   0   'False
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   11
            Left            =   2040
            TabIndex        =   34
            Top             =   2280
            Width           =   480
            _Version        =   720898
            _ExtentX        =   847
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "CEP:"
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   4080
         Left            =   30
         TabIndex        =   64
         Top             =   345
         Width           =   6675
         _Version        =   720898
         _ExtentX        =   11774
         _ExtentY        =   7197
         _StockProps     =   1
         Page            =   0
         Begin XtremeSuiteControls.CheckBox ChkFLGCERTIFICADO 
            Height          =   255
            Left            =   3840
            TabIndex        =   12
            Top             =   180
            Width           =   1575
            _Version        =   720898
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Certificado"
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
         Begin XtremeSuiteControls.FlatEdit TxtCHAPA 
            Height          =   330
            Left            =   1080
            TabIndex        =   11
            Top             =   600
            Width           =   1905
            _Version        =   720898
            _ExtentX        =   3360
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   6
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtObs 
            Height          =   825
            Left            =   120
            TabIndex        =   23
            Top             =   3120
            Width           =   6360
            _Version        =   720898
            _ExtentX        =   11218
            _ExtentY        =   1455
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtDTADMISSAO 
            Height          =   330
            Left            =   1080
            TabIndex        =   14
            Top             =   1080
            Width           =   1065
            _Version        =   720898
            _ExtentX        =   1879
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "88/88/8888"
            MaxLength       =   14
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtDTDEMISSAO 
            Height          =   330
            Left            =   1080
            TabIndex        =   17
            Top             =   1560
            Width           =   1065
            _Version        =   720898
            _ExtentX        =   1879
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "88/88/8888"
            MaxLength       =   14
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtDTNASC 
            Height          =   330
            Left            =   1080
            TabIndex        =   19
            Top             =   2040
            Width           =   1065
            _Version        =   720898
            _ExtentX        =   1879
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "88/88/8888"
            MaxLength       =   14
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtSenha 
            Height          =   330
            Left            =   3840
            TabIndex        =   21
            Top             =   2040
            Width           =   1905
            _Version        =   720898
            _ExtentX        =   3360
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            PasswordChar    =   "?"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkFLGVENDA 
            Height          =   255
            Left            =   3840
            TabIndex        =   15
            Top             =   720
            Width           =   1575
            _Version        =   720898
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Vendedor"
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
         Begin XtremeSuiteControls.ComboBox CmbIDLOJA0 
            Height          =   315
            Left            =   1080
            TabIndex        =   9
            Top             =   120
            Width           =   1935
            _Version        =   720898
            _ExtentX        =   3413
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            UseVisualStyle  =   -1  'True
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   12
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1080
            _Version        =   720898
            _ExtentX        =   1905
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Unidade:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   3
            Left            =   3840
            TabIndex        =   20
            Top             =   1680
            Width           =   1080
            _Version        =   720898
            _ExtentX        =   1905
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Senha:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   13
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Width           =   945
            _Version        =   720898
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Nascimento:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   16
            Left            =   120
            TabIndex        =   16
            Top             =   1560
            Width           =   825
            _Version        =   720898
            _ExtentX        =   1455
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Demissão:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   15
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1185
            _Version        =   720898
            _ExtentX        =   2090
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Admissão:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   1080
            _Version        =   720898
            _ExtentX        =   1905
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Registro:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   17
            Left            =   120
            TabIndex        =   22
            Top             =   2760
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Observação:"
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   4080
         Left            =   -69970
         TabIndex        =   65
         Top             =   345
         Visible         =   0   'False
         Width           =   6675
         _Version        =   720898
         _ExtentX        =   11774
         _ExtentY        =   7197
         _StockProps     =   1
         Page            =   2
         Begin XtremeSuiteControls.TabControl TabValores 
            Height          =   3855
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   6375
            _Version        =   720898
            _ExtentX        =   11245
            _ExtentY        =   6800
            _StockProps     =   68
            AutoResizeClient=   0   'False
            Appearance      =   2
            Color           =   16
            PaintManager.BoldSelected=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.FixedTabWidth=   80
            PaintManager.MinTabWidth=   70
            ItemCount       =   3
            Item(0).Caption =   "Salário"
            Item(0).ControlCount=   1
            Item(0).Control(0)=   "TabControlPage4"
            Item(1).Caption =   "Comissão"
            Item(1).ControlCount=   1
            Item(1).Control(0)=   "TabControlPage6"
            Item(2).Caption =   "Expediente"
            Item(2).ControlCount=   1
            Item(2).Control(0)=   "TabControlPage5"
            Begin XtremeSuiteControls.TabControlPage TabControlPage5 
               Height          =   3480
               Left            =   -69970
               TabIndex        =   68
               Top             =   345
               Visible         =   0   'False
               Width           =   6315
               _Version        =   720898
               _ExtentX        =   11139
               _ExtentY        =   6138
               _StockProps     =   1
               Page            =   2
               Begin XtremeSuiteControls.ComboBox CmbDIAFOLGA 
                  Height          =   315
                  Left            =   600
                  TabIndex        =   69
                  Top             =   480
                  Width           =   1695
                  _Version        =   720898
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   77
                  BackColor       =   -2147483643
                  Style           =   2
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.Label Label3 
                  Height          =   330
                  Left            =   600
                  TabIndex        =   70
                  Top             =   120
                  Width           =   1440
                  _Version        =   720898
                  _ExtentX        =   2540
                  _ExtentY        =   582
                  _StockProps     =   79
                  Caption         =   "Folga Semanal"
               End
            End
            Begin XtremeSuiteControls.TabControlPage TabControlPage4 
               Height          =   3480
               Left            =   30
               TabIndex        =   66
               Top             =   345
               Width           =   6315
               _Version        =   720898
               _ExtentX        =   11139
               _ExtentY        =   6138
               _StockProps     =   1
               Page            =   0
               Begin XtremeSuiteControls.FlatEdit TxtSALARIO 
                  Height          =   330
                  Left            =   1080
                  TabIndex        =   45
                  Top             =   225
                  Width           =   1200
                  _Version        =   720898
                  _ExtentX        =   2117
                  _ExtentY        =   582
                  _StockProps     =   77
                  BackColor       =   -2147483643
                  MaxLength       =   50
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.Label Label2 
                  Height          =   330
                  Index           =   14
                  Left            =   120
                  TabIndex        =   44
                  Top             =   225
                  Width           =   855
                  _Version        =   720898
                  _ExtentX        =   1508
                  _ExtentY        =   582
                  _StockProps     =   79
                  Caption         =   "Valor Base:"
               End
            End
            Begin XtremeSuiteControls.TabControlPage TabControlPage6 
               Height          =   3480
               Left            =   -69970
               TabIndex        =   67
               Top             =   345
               Visible         =   0   'False
               Width           =   6315
               _Version        =   720898
               _ExtentX        =   11139
               _ExtentY        =   6138
               _StockProps     =   1
               Page            =   1
               Begin XtremeSuiteControls.GroupBox GrpComissao 
                  Height          =   1560
                  Left            =   600
                  TabIndex        =   46
                  Top             =   120
                  Width           =   5175
                  _Version        =   720898
                  _ExtentX        =   9128
                  _ExtentY        =   2752
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
                  Appearance      =   1
                  Begin XtremeSuiteControls.CheckBox ChkCOMPROD 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   47
                     Top             =   0
                     Width           =   1095
                     _Version        =   720898
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Produto"
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
                  Begin XtremeSuiteControls.RadioButton OptTPCOMPROD 
                     Height          =   255
                     Index           =   1
                     Left            =   360
                     TabIndex        =   51
                     Top             =   1200
                     Width           =   3855
                     _Version        =   720898
                     _ExtentX        =   6800
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Apenas nos produtos vendidos pelo funcionário"
                     ForeColor       =   0
                     Transparent     =   -1  'True
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton OptTPCOMPROD 
                     Height          =   255
                     Index           =   0
                     Left            =   360
                     TabIndex        =   50
                     Top             =   840
                     Width           =   3135
                     _Version        =   720898
                     _ExtentX        =   5530
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Rateio em todos os produtos vendidos"
                     ForeColor       =   0
                     Transparent     =   -1  'True
                     UseVisualStyle  =   -1  'True
                     Value           =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit TxtVLCOMPROD 
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   49
                     Top             =   360
                     Width           =   1305
                     _Version        =   720898
                     _ExtentX        =   2302
                     _ExtentY        =   503
                     _StockProps     =   77
                     BackColor       =   -2147483643
                     Enabled         =   0   'False
                     MaxLength       =   20
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.Label LblVLCOMPROD 
                     Height          =   330
                     Left            =   600
                     TabIndex        =   48
                     Top             =   360
                     Width           =   720
                     _Version        =   720898
                     _ExtentX        =   1270
                     _ExtentY        =   582
                     _StockProps     =   79
                     Caption         =   "Valor: R$"
                  End
               End
               Begin XtremeSuiteControls.GroupBox GroupBox1 
                  Height          =   1560
                  Left            =   600
                  TabIndex        =   52
                  Top             =   1800
                  Width           =   5175
                  _Version        =   720898
                  _ExtentX        =   9128
                  _ExtentY        =   2752
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
                  Appearance      =   1
                  Begin XtremeSuiteControls.CheckBox ChkCOMSERV 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   53
                     Top             =   0
                     Width           =   1095
                     _Version        =   720898
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Serviço"
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
                  Begin XtremeSuiteControls.RadioButton OptTPCOMSERV 
                     Height          =   255
                     Index           =   1
                     Left            =   360
                     TabIndex        =   57
                     Top             =   1200
                     Width           =   3735
                     _Version        =   720898
                     _ExtentX        =   6588
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Apenas nos serviços vendidos pelo funcionário"
                     ForeColor       =   0
                     Transparent     =   -1  'True
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton OptTPCOMSERV 
                     Height          =   255
                     Index           =   0
                     Left            =   360
                     TabIndex        =   56
                     Top             =   840
                     Width           =   3135
                     _Version        =   720898
                     _ExtentX        =   5530
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Rateio em todos os serviços vendidos"
                     ForeColor       =   0
                     Transparent     =   -1  'True
                     UseVisualStyle  =   -1  'True
                     Value           =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit TxtVLCOMSERV 
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   55
                     Top             =   360
                     Width           =   1305
                     _Version        =   720898
                     _ExtentX        =   2302
                     _ExtentY        =   503
                     _StockProps     =   77
                     BackColor       =   -2147483643
                     Enabled         =   0   'False
                     MaxLength       =   20
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.Label LblVLCOMSERV 
                     Height          =   330
                     Left            =   600
                     TabIndex        =   54
                     Top             =   360
                     Width           =   720
                     _Version        =   720898
                     _ExtentX        =   1270
                     _ExtentY        =   582
                     _StockProps     =   79
                     Caption         =   "Valor: R$"
                  End
               End
            End
         End
      End
   End
   Begin XtremeSuiteControls.PushButton CmdBiometria 
      Height          =   375
      Left            =   3000
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Biometria"
      ForeColor       =   8388608
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADRFUNCIONARIO.frx":2671
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   510
      _Version        =   720898
      _ExtentX        =   900
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Id.:"
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   330
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   645
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nome:"
   End
End
Attribute VB_Name = "FrmCADRFUNCIONARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event CmdBiometriaClick()
Event CmdSalvarClick()
Event CmdSairClick()
Event CmdNovoClick()
Event CmdExcluirClick()
Event ChkCOMPRODClick()
Event ChkCOMSERVClick()
Private Sub ChkCOMPROD_Click()
   RaiseEvent ChkCOMPRODClick
End Sub
Private Sub ChkCOMSERV_Click()
   RaiseEvent ChkCOMSERVClick
End Sub

Private Sub CmdBiometria_Click()
   RaiseEvent CmdBiometriaClick
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

Private Sub TxtDTADMISSAO_LostFocus()
   TxtDTADMISSAO.Text = FormatarData(TxtDTADMISSAO.Text)
End Sub
Private Sub TxtDTDEMISSAO_LostFocus()
   TxtDTDEMISSAO.Text = FormatarData(TxtDTDEMISSAO.Text, True)
End Sub
Private Sub TxtDTNASC_LostFocus()
   TxtDTNASC.Text = FormatarData(TxtDTNASC.Text, True)
End Sub
Private Sub TxtVLCOMPROD_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLCOMPROD_LostFocus()
   TxtVLCOMPROD.Text = ValBr(TxtVLCOMPROD.Text)
End Sub
Private Sub TxtVLCOMSERV_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLCOMSERV_LostFocus()
   TxtVLCOMSERV.Text = ValBr(TxtVLCOMSERV.Text)
End Sub
