VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmCadCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   7755
   ClientLeft      =   1815
   ClientTop       =   1245
   ClientWidth     =   12240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl TabContato 
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   5415
      _Version        =   720898
      _ExtentX        =   9551
      _ExtentY        =   7435
      _StockProps     =   68
      AutoResizeClient=   0   'False
      Appearance      =   2
      Color           =   8
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.MinTabWidth=   70
      ItemCount       =   4
      Item(0).Caption =   "Principal"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Endereço"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "Classificação"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Item(3).Caption =   "Outros"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage4"
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   3840
         Left            =   -69970
         TabIndex        =   71
         Top             =   345
         Visible         =   0   'False
         Width           =   5355
         _Version        =   720898
         _ExtentX        =   9446
         _ExtentY        =   6773
         _StockProps     =   1
         Page            =   3
         Begin iGrid251_75B4A91C.iGrid GrdCupom 
            Height          =   2775
            Left            =   240
            TabIndex        =   79
            Top             =   960
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   4895
            BorderStyle     =   1
            HighlightBackColorNoFocus=   14737632
         End
         Begin XtremeSuiteControls.FlatEdit txtEmpresa 
            Height          =   330
            Left            =   1170
            TabIndex        =   56
            Top             =   2880
            Visible         =   0   'False
            Width           =   2625
            _Version        =   720898
            _ExtentX        =   4630
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCargo 
            Height          =   330
            Left            =   1200
            TabIndex        =   58
            Top             =   3360
            Visible         =   0   'False
            Width           =   2715
            _Version        =   720898
            _ExtentX        =   4789
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkIsento 
            Height          =   255
            Left            =   360
            TabIndex        =   77
            Top             =   240
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Isento"
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkNFE 
            Height          =   255
            Left            =   3120
            TabIndex        =   78
            Top             =   240
            Width           =   1815
            _Version        =   720898
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Solicita NF por e-Mail"
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   20
            Left            =   240
            TabIndex        =   80
            Top             =   600
            Width           =   1575
            _Version        =   720898
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Promoções / Cupons:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   13
            Left            =   360
            TabIndex        =   55
            Top             =   2880
            Visible         =   0   'False
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Empresa:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   14
            Left            =   360
            TabIndex        =   57
            Top             =   3360
            Visible         =   0   'False
            Width           =   555
            _Version        =   720898
            _ExtentX        =   979
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Cargo:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   3840
         Left            =   -69970
         TabIndex        =   31
         Top             =   345
         Visible         =   0   'False
         Width           =   5355
         _Version        =   720898
         _ExtentX        =   9446
         _ExtentY        =   6773
         _StockProps     =   1
         Page            =   1
         Begin XtremeSuiteControls.FlatEdit txtEndereco 
            Height          =   945
            Left            =   940
            TabIndex        =   33
            Top             =   240
            Width           =   4080
            _Version        =   720898
            _ExtentX        =   7197
            _ExtentY        =   1667
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbPais 
            Height          =   315
            Left            =   945
            TabIndex        =   43
            Top             =   3240
            Visible         =   0   'False
            Width           =   1545
            _Version        =   720898
            _ExtentX        =   2725
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   5
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCEP 
            Height          =   330
            Left            =   945
            TabIndex        =   41
            Top             =   2760
            Width           =   1320
            _Version        =   720898
            _ExtentX        =   2328
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   9
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cmbEstado 
            Height          =   315
            Left            =   940
            TabIndex        =   39
            Top             =   2280
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   5
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbBairro 
            Height          =   315
            Left            =   960
            TabIndex        =   35
            Top             =   1320
            Width           =   2400
            _Version        =   720898
            _ExtentX        =   4233
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   5
            UseVisualStyle  =   -1  'True
            AutoComplete    =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbCidade 
            Height          =   315
            Left            =   960
            TabIndex        =   37
            Top             =   1800
            Width           =   2400
            _Version        =   720898
            _ExtentX        =   4233
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   5
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   11
            Left            =   120
            TabIndex        =   40
            Top             =   2760
            Width           =   600
            _Version        =   720898
            _ExtentX        =   1058
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "CEP:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   10
            Left            =   120
            TabIndex        =   42
            Top             =   3240
            Visible         =   0   'False
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "País:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   9
            Left            =   120
            TabIndex        =   38
            Top             =   2280
            Width           =   645
            _Version        =   720898
            _ExtentX        =   1138
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Estado:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   8
            Left            =   120
            TabIndex        =   36
            Top             =   1800
            Width           =   600
            _Version        =   720898
            _ExtentX        =   1058
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Cidade:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   7
            Left            =   120
            TabIndex        =   34
            Top             =   1320
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Bairro:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   6
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Endereço:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   3840
         Left            =   30
         TabIndex        =   8
         Top             =   345
         Width           =   5355
         _Version        =   720898
         _ExtentX        =   9446
         _ExtentY        =   6773
         _StockProps     =   1
         Page            =   0
         Begin XtremeSuiteControls.GroupBox GrpConhec 
            Height          =   1215
            Left            =   120
            TabIndex        =   25
            Top             =   2625
            Width           =   5130
            _Version        =   720898
            _ExtentX        =   9049
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Como conheceu a Empresa? "
            Transparent     =   -1  'True
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox CmbConhecimento 
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Top             =   280
               Width           =   2130
               _Version        =   720898
               _ExtentX        =   3757
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Style           =   2
               Appearance      =   5
               UseVisualStyle  =   -1  'True
               DropDownItemCount=   15
            End
            Begin XtremeSuiteControls.CheckBox ChkFLGAGENDA 
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   930
               Width           =   3495
               _Version        =   720898
               _ExtentX        =   6165
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "NÃO recebe e-Mail de Agenda"
               Transparent     =   -1  'True
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChkFLGMARKETING 
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   680
               Width           =   3495
               _Version        =   720898
               _ExtentX        =   6165
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "NÃO recebe e-Mail Marketing"
               Transparent     =   -1  'True
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ComboBox CmbDSCTPCONHEC 
               Height          =   315
               Left            =   2880
               TabIndex        =   28
               Top             =   280
               Width           =   2115
               _Version        =   720898
               _ExtentX        =   3731
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Appearance      =   5
               UseVisualStyle  =   -1  'True
               AutoComplete    =   -1  'True
               DropDownItemCount=   15
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   330
               Index           =   16
               Left            =   120
               TabIndex        =   74
               Top             =   240
               Width           =   1455
               _Version        =   720898
               _ExtentX        =   2566
               _ExtentY        =   582
               _StockProps     =   79
               Caption         =   "Tp. Conhecimento:"
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   330
               Index           =   18
               Left            =   2400
               TabIndex        =   27
               Top             =   280
               Width           =   1095
               _Version        =   720898
               _ExtentX        =   1931
               _ExtentY        =   582
               _StockProps     =   79
               Caption         =   "Qual?"
               Transparent     =   -1  'True
            End
         End
         Begin XtremeSuiteControls.FlatEdit TxtDTNASC 
            Height          =   330
            Left            =   1080
            TabIndex        =   16
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
         Begin XtremeSuiteControls.FlatEdit txtTel 
            Height          =   330
            Left            =   3680
            TabIndex        =   12
            Top             =   240
            Width           =   1425
            _Version        =   720898
            _ExtentX        =   2514
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCelular 
            Height          =   330
            Left            =   1080
            TabIndex        =   10
            Top             =   240
            Width           =   1425
            _Version        =   720898
            _ExtentX        =   2514
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtEmail 
            Height          =   330
            Left            =   1080
            TabIndex        =   14
            Top             =   645
            Width           =   4025
            _Version        =   720898
            _ExtentX        =   7100
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   50
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtFax 
            Height          =   330
            Left            =   1080
            TabIndex        =   20
            Top             =   1485
            Width           =   1425
            _Version        =   720898
            _ExtentX        =   2514
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbSexo 
            Height          =   315
            Left            =   3600
            TabIndex        =   18
            Top             =   1080
            Width           =   1515
            _Version        =   720898
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Sorted          =   -1  'True
            Style           =   2
            Appearance      =   5
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GrpFototipo 
            Height          =   735
            Left            =   120
            TabIndex        =   21
            Top             =   1920
            Width           =   2370
            _Version        =   720898
            _ExtentX        =   4180
            _ExtentY        =   1296
            _StockProps     =   79
            Caption         =   "Fototipo "
            Transparent     =   -1  'True
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox CmbFototipo 
               Height          =   315
               Left            =   120
               TabIndex        =   22
               Top             =   280
               Width           =   2115
               _Version        =   720898
               _ExtentX        =   3731
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Style           =   2
               Appearance      =   5
               UseVisualStyle  =   -1  'True
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   735
            Left            =   2880
            TabIndex        =   23
            Top             =   1920
            Width           =   2370
            _Version        =   720898
            _ExtentX        =   4180
            _ExtentY        =   1296
            _StockProps     =   79
            Caption         =   "Profissão "
            Transparent     =   -1  'True
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox CmbProfissao 
               Height          =   315
               Left            =   0
               TabIndex        =   24
               Top             =   285
               Width           =   2235
               _Version        =   720898
               _ExtentX        =   3942
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Appearance      =   5
               UseVisualStyle  =   -1  'True
               AutoComplete    =   -1  'True
            End
         End
         Begin XtremeSuiteControls.ComboBox CmbFunc 
            Height          =   315
            Left            =   3600
            TabIndex        =   91
            Top             =   1485
            Width           =   1515
            _Version        =   720898
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Sorted          =   -1  'True
            Style           =   2
            Appearance      =   5
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   15
            Left            =   2640
            TabIndex        =   92
            Top             =   1485
            Width           =   945
            _Version        =   720898
            _ExtentX        =   1667
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Preferência:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblMenor2 
            Height          =   195
            Left            =   2160
            TabIndex        =   83
            Top             =   1240
            Visible         =   0   'False
            Width           =   540
            _Version        =   720898
            _ExtentX        =   953
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "   Idade"
            ForeColor       =   255
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblMenor 
            Height          =   195
            Left            =   2160
            TabIndex        =   82
            Top             =   1080
            Visible         =   0   'False
            Width           =   675
            _Version        =   720898
            _ExtentX        =   1191
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Menor de"
            ForeColor       =   255
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   19
            Left            =   2880
            TabIndex        =   17
            Top             =   1080
            Width           =   585
            _Version        =   720898
            _ExtentX        =   1032
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Sexo:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   12
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   1185
            _Version        =   720898
            _ExtentX        =   2090
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Data Nasc.:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   1485
            Width           =   705
            _Version        =   720898
            _ExtentX        =   1244
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Outro Tel:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   5
            Left            =   120
            TabIndex        =   13
            Top             =   645
            Width           =   465
            _Version        =   720898
            _ExtentX        =   820
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Email:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   0
            Left            =   2880
            TabIndex        =   11
            Top             =   240
            Width           =   840
            _Version        =   720898
            _ExtentX        =   1482
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Tel. Res.:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Tel. Celular:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   3840
         Left            =   -69970
         TabIndex        =   68
         Top             =   345
         Visible         =   0   'False
         Width           =   5355
         _Version        =   720898
         _ExtentX        =   9446
         _ExtentY        =   6773
         _StockProps     =   1
         Page            =   2
         Begin XtremeSuiteControls.TreeView TrvClasse 
            CausesValidation=   0   'False
            Height          =   1095
            Left            =   240
            TabIndex        =   51
            Top             =   2400
            Width           =   4920
            _Version        =   720898
            _ExtentX        =   8678
            _ExtentY        =   1931
            _StockProps     =   77
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            LabelEdit       =   1
            Appearance      =   6
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkClasse 
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   3520
            Visible         =   0   'False
            Width           =   2055
            _Version        =   720898
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Somente selecionados"
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdIncluirClasse 
            Height          =   300
            Left            =   2760
            TabIndex        =   53
            Top             =   3520
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "&Incluir"
            ForeColor       =   12582912
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdEditarClasse 
            Height          =   300
            Left            =   3960
            TabIndex        =   54
            Top             =   3520
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "&Editar"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1740
            Left            =   120
            TabIndex        =   73
            Top             =   120
            Width           =   5130
            _Version        =   720898
            _ExtentX        =   9049
            _ExtentY        =   3069
            _StockProps     =   79
            Caption         =   "    Situação"
            Transparent     =   -1  'True
            BorderStyle     =   2
            Begin XtremeSuiteControls.RadioButton OptATIVO 
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   45
               Top             =   0
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
            Begin XtremeSuiteControls.FlatEdit TxtMotivoInat 
               Height          =   1095
               Left            =   120
               TabIndex        =   48
               Top             =   600
               Width           =   4920
               _Version        =   720898
               _ExtentX        =   8678
               _ExtentY        =   1931
               _StockProps     =   77
               BackColor       =   -2147483643
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton OptATIVO 
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   44
               Top             =   0
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
            Begin XtremeSuiteControls.RadioButton OptATIVO 
               Height          =   255
               Index           =   2
               Left            =   3600
               TabIndex        =   46
               Top             =   0
               Width           =   1215
               _Version        =   720898
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Em Espera"
               ForeColor       =   4210752
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
            Begin XtremeSuiteControls.Label LblMotivo 
               Height          =   330
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   1455
               _Version        =   720898
               _ExtentX        =   2566
               _ExtentY        =   582
               _StockProps     =   79
               Caption         =   "Motivo:"
               Transparent     =   -1  'True
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   135
            Left            =   240
            TabIndex        =   49
            Top             =   1920
            Width           =   4935
            _Version        =   720898
            _ExtentX        =   8705
            _ExtentY        =   238
            _StockProps     =   79
            Transparent     =   -1  'True
            BorderStyle     =   1
         End
         Begin XtremeSuiteControls.Label LblClasse 
            Height          =   330
            Left            =   240
            TabIndex        =   50
            Top             =   2040
            Width           =   1695
            _Version        =   720898
            _ExtentX        =   2990
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Classificação Diversa:"
            Transparent     =   -1  'True
         End
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtID 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   615
      _Version        =   720898
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   77
      BackColor       =   -2147483643
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   67
      Top             =   7200
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Excluir"
      ForeColor       =   64
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCadCliente.frx":0000
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   9360
      TabIndex        =   66
      Top             =   7200
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
      Left            =   6960
      TabIndex        =   65
      Top             =   7200
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtNome 
      Height          =   330
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   4680
      _Version        =   720898
      _ExtentX        =   8255
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   50
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdNovo 
      Height          =   375
      Left            =   1920
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Novo"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCadCliente.frx":0ACA
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   3720
      TabIndex        =   70
      Top             =   7200
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   32768
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCadCliente.frx":0E64
   End
   Begin XtremeSuiteControls.FlatEdit TxtRegistro 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   1665
      _Version        =   720898
      _ExtentX        =   2937
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   14
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdLov 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCadCliente.frx":272E
   End
   Begin XtremeSuiteControls.FlatEdit TxtObs 
      Height          =   945
      Left            =   120
      TabIndex        =   60
      Top             =   6000
      Width           =   5400
      _Version        =   720898
      _ExtentX        =   9525
      _ExtentY        =   1667
      _StockProps     =   77
      BackColor       =   -2147483643
      MultiLine       =   -1  'True
      ScrollBars      =   2
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdExpHis 
      Height          =   240
      Left            =   6480
      TabIndex        =   84
      ToolTipText     =   "Exportar Histórico"
      Top             =   1320
      Width           =   240
      _Version        =   720898
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   8
      Appearance      =   2
      Picture         =   "FrmCadCliente.frx":28B1
      TextImageRelation=   0
   End
   Begin XtremeSuiteControls.GroupBox GrpSessao 
      Height          =   5775
      Left            =   5565
      TabIndex        =   61
      Top             =   1320
      Width           =   6570
      _Version        =   720898
      _ExtentX        =   11589
      _ExtentY        =   10186
      _StockProps     =   79
      Caption         =   "HISTÓRICO "
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton CmdNovoItem 
         Height          =   300
         Left            =   240
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   5280
         Visible         =   0   'False
         Width           =   1455
         _Version        =   720898
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "&Gerar Fatura"
         ForeColor       =   12582912
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCadCliente.frx":2C4B
      End
      Begin iGrid251_75B4A91C.iGrid GrdSessao 
         Height          =   4695
         Left            =   120
         TabIndex        =   63
         Top             =   555
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   8281
         Appearance      =   0
      End
      Begin XtremeSuiteControls.PushButton CmdExtende 
         Height          =   240
         Left            =   1440
         TabIndex        =   81
         Top             =   315
         Width           =   240
         _Version        =   720898
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "»"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   8
         Appearance      =   2
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.GroupBox GrpValor 
         Height          =   615
         Left            =   120
         TabIndex        =   64
         Top             =   5040
         Width           =   6345
         _Version        =   720898
         _ExtentX        =   11192
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "GroupBox2"
         BackColor       =   -2147483637
         Appearance      =   3
         Begin XtremeSuiteControls.Label LblVALOR 
            Height          =   195
            Left            =   4680
            TabIndex        =   76
            Top             =   240
            Width           =   1530
            _Version        =   720898
            _ExtentX        =   2699
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "R$ 0,00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
         Begin XtremeSuiteControls.Label LblTotal1 
            Height          =   195
            Left            =   4080
            TabIndex        =   75
            Top             =   240
            Width           =   570
            _Version        =   720898
            _ExtentX        =   1005
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   " Total:"
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
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.TabControl TabSessao 
         Height          =   3135
         Left            =   105
         TabIndex        =   62
         Top             =   240
         Width           =   6015
         _Version        =   720898
         _ExtentX        =   10610
         _ExtentY        =   5530
         _StockProps     =   68
         AutoResizeClient=   0   'False
         Appearance      =   2
         Color           =   8
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.MinTabWidth=   70
         ItemCount       =   5
         Item(0).Caption =   "Atendimentos    "
         Item(0).ImageIndex=   0
         Item(0).ControlCount=   0
         Item(1).Caption =   "Agenda"
         Item(1).ImageIndex=   0
         Item(1).ControlCount=   0
         Item(2).Caption =   "Tarefas"
         Item(2).ControlCount=   0
         Item(3).Caption =   "Vendas"
         Item(3).ControlCount=   0
         Item(4).Caption =   "Faturas"
         Item(4).ControlCount=   0
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpProxSessao 
      Height          =   495
      Left            =   9120
      TabIndex        =   85
      Top             =   840
      Width           =   2835
      _Version        =   720898
      _ExtentX        =   5001
      _ExtentY        =   873
      _StockProps     =   79
      Transparent     =   -1  'True
      Appearance      =   1
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton CmdProxSessao 
         Height          =   240
         Left            =   240
         TabIndex        =   86
         Top             =   240
         Width           =   2520
         _Version        =   720898
         _ExtentX        =   4445
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Segunda-Feira 28/12 10:40h"
         ForeColor       =   8421504
         Transparent     =   -1  'True
         TextAlignment   =   1
         Appearance      =   2
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.Label LblProx 
         Height          =   195
         Left            =   720
         TabIndex        =   88
         Top             =   240
         Width           =   1725
         _Version        =   720898
         _ExtentX        =   3043
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Quinta 28/12 10:40h"
         ForeColor       =   8421504
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblLblProx 
         Height          =   195
         Left            =   1200
         TabIndex        =   87
         Top             =   0
         Width           =   1485
         _Version        =   720898
         _ExtentX        =   2619
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Próxima Sessão"
         ForeColor       =   8421504
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList ImgFatura 
      Left            =   7440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCliente.frx":2FE5
            Key             =   "K1"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCliente.frx":313F
            Key             =   "K2"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCliente.frx":3789
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCliente.frx":38E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCadCliente.frx":3C7D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.Label LblDTCADASTRO 
      Height          =   240
      Left            =   8640
      TabIndex        =   90
      Top             =   120
      Width           =   3240
      _Version        =   720898
      _ExtentX        =   5715
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   " Data de Cadastro:"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   330
      Index           =   17
      Left            =   120
      TabIndex        =   59
      Top             =   5720
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Observação:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   330
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "C.P.F:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblSessao 
      Height          =   330
      Left            =   7920
      TabIndex        =   72
      Top             =   1320
      Width           =   960
      _Version        =   720898
      _ExtentX        =   1693
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Sessões"
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   330
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nome:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   510
      _Version        =   720898
      _ExtentX        =   900
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Id.:"
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmCadCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Load()
Event Resize()
Event ChkClasseClick()
Event CmbBairroLostFocus()
Event CmbCidadeLostFocus()
Event CmbProfissaoLostFocus()
Event CmbConhecimentoClick()
Event CmbConhecimentoLostFocus()
Event CmbDSCTPCONHECGotFocus()
Event CmdSairClick()
Event CmdSalvarClick()
Event CmdOkClick()
Event CmdExcluirClick()
Event CmdExpHisClick()
Event CmdNovoClick()
Event CmdNovoItemClick()
Event CmdLovClick()
Event CmdIncluirClasseClick()
Event CmdEditarClasseClick()
Event CmdExtendeClick()
Event GrdSessaoTextEditDblClick(ByVal lRow As Long, ByVal lCol As Long)
Event GrdSessaoDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
Event LblProxDblClick()
Event TabSessaoSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Event TxtDTNASCLostFocus()
Event TxtEmailLostFocus()
Event TxtNomeLostFocus()
Event TxtRegistroGotFocus()
Event TxtRegistroLostFocus()
Event OptATIVOClick(Index As Integer)
Event TrvClasseDblClick()
Event TrvClasseNodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
Event TrvClasseNodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
Private Sub ChkClasse_Click()
   RaiseEvent ChkClasseClick
End Sub

Private Sub CmbBairro_KeyUp(KeyCode As Integer, Shift As Integer)
'   Dim Pos As Integer
'   On Error Resume Next
'   If UCase(Chr(KeyCode)) >= "A" And UCase(Chr(KeyCode)) <= "Z" Then
'      Pos = Len(Me.CmbBairro.Text)
'      If LocalizarCombo(Me.CmbBairro, Me.CmbBairro.Text, False) >= 0 Then
'         Call LocalizarCombo(Me.CmbBairro, Me.CmbBairro.Text)
'         If Pos <= Len(Me.CmbBairro.Text) Then
'            Me.CmbBairro.SelStart = Pos
'            Me.CmbBairro.SelLength = Len(Me.CmbBairro.Text)
'         End If
'      End If
'   End If
End Sub
Private Sub CmbBairro_LostFocus()
   RaiseEvent CmbBairroLostFocus
End Sub
Private Sub CmbCidade_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim Pos As Integer
   On Error Resume Next
   If UCase(Chr(KeyCode)) >= "A" And UCase(Chr(KeyCode)) <= "Z" Then
      Pos = Len(Me.CmbCidade.Text)
      If LocalizarCombo(Me.CmbCidade, Me.CmbCidade.Text, False) >= 0 Then
         Call LocalizarCombo(Me.CmbCidade, Me.CmbCidade.Text)
         If Pos <= Len(Me.CmbCidade.Text) Then
            Me.CmbCidade.SelStart = Pos
            Me.CmbCidade.SelLength = Len(Me.CmbCidade.Text)
         End If
      End If
   End If
End Sub
Private Sub CmbCidade_LostFocus()
   RaiseEvent CmbCidadeLostFocus
End Sub
Private Sub CmbConhecimento_Click()
   RaiseEvent CmbConhecimentoClick
End Sub
Private Sub cmbConhecimento_LostFocus()
   RaiseEvent CmbConhecimentoLostFocus
End Sub
Private Sub CmbDSCTPCONHEC_GotFocus()
   RaiseEvent CmbDSCTPCONHECGotFocus
   Call SelecionarTexto(Me.CmbDSCTPCONHEC)
End Sub
Private Sub cmbEstado_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim Pos As Integer
   On Error Resume Next
   If UCase(Chr(KeyCode)) >= "A" And UCase(Chr(KeyCode)) <= "Z" Then
      Pos = Len(Me.cmbEstado.Text)
      If LocalizarCombo(Me.cmbEstado, Me.cmbEstado.Text, False) >= 0 Then
         Call LocalizarCombo(Me.cmbEstado, Me.cmbEstado.Text)
         If Pos <= Len(Me.cmbEstado.Text) Then
            Me.cmbEstado.SelStart = Pos
            Me.cmbEstado.SelLength = Len(Me.cmbEstado.Text)
         End If
      End If
   End If
End Sub
Private Sub CmbProfissao_KeyUp(KeyCode As Integer, Shift As Integer)
'   Dim Pos As Integer
'   On Error Resume Next
'   If UCase(Chr(KeyCode)) >= "A" And UCase(Chr(KeyCode)) <= "Z" Then
'      Pos = Len(Me.CmbProfissao.Text)
'      If LocalizarCombo(Me.CmbProfissao, Me.CmbProfissao.Text, False) >= 0 Then
'         Call LocalizarCombo(Me.CmbProfissao, Me.CmbProfissao.Text)
'         If Pos <= Len(Me.CmbProfissao.Text) Then
'            Me.CmbProfissao.SelStart = Pos
'            Me.CmbProfissao.SelLength = Len(Me.CmbProfissao.Text)
'         End If
'      End If
'   End If
End Sub
Private Sub CmbProfissao_LostFocus()
   RaiseEvent CmbProfissaoLostFocus
End Sub
Private Sub CmdEditarClasse_Click()
   RaiseEvent CmdEditarClasseClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdExpHis_Click()
   RaiseEvent CmdExpHisClick
End Sub
Private Sub CmdExtende_Click()
   RaiseEvent CmdExtendeClick
End Sub
Private Sub CmdIncluirClasse_Click()
   RaiseEvent CmdIncluirClasseClick
End Sub
Private Sub cmdLov_Click()
   RaiseEvent CmdLovClick
End Sub
Private Sub cmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub
Private Sub CmdNovoItem_Click()
   RaiseEvent CmdNovoItemClick
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub CmdProxSessao_Click()
   RaiseEvent LblProxDblClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub cmdSalvar_Click()
   RaiseEvent CmdSalvarClick
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
Private Sub GrdSessao_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   RaiseEvent GrdSessaoDblClick(lRow, lCol, bRequestEdit)
End Sub
Private Sub GrdSessao_TextEditDblClick(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdSessaoTextEditDblClick(lRow, lCol)
End Sub

Private Sub LblLblProx_DblClick()
   RaiseEvent LblProxDblClick
End Sub
Private Sub LblProx_DblClick()
   RaiseEvent LblProxDblClick
End Sub

Private Sub OptATIVO_Click(Index As Integer)
   RaiseEvent OptATIVOClick(Index)
End Sub
Private Sub TabSessao_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   RaiseEvent TabSessaoSelectedChanged(Item)
End Sub
Private Sub TrvClasse_DblClick()
   RaiseEvent TrvClasseDblClick
End Sub
Private Sub TrvClasse_NodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
   RaiseEvent TrvClasseNodeCheck(Node)
End Sub
Private Sub TrvClasse_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
    RaiseEvent TrvClasseNodeClick(Node)
End Sub

Private Sub TxtDSCTPCONHEC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub TxtDTNASC_LostFocus()
   RaiseEvent TxtDTNASCLostFocus
End Sub
Private Sub TxtEmail_LostFocus()
   RaiseEvent TxtEmailLostFocus
End Sub
Private Sub txtNome_LostFocus()
   RaiseEvent TxtNomeLostFocus
End Sub
Private Sub TxtRegistro_GotFocus()
   RaiseEvent TxtRegistroGotFocus
End Sub
Private Sub TxtRegistro_LostFocus()
   RaiseEvent TxtRegistroLostFocus
End Sub
