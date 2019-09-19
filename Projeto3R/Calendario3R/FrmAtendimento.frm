VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "CODEJO~1.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAtendimento 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atendimento"
   ClientHeight    =   7440
   ClientLeft      =   2565
   ClientTop       =   2070
   ClientWidth     =   11415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton CmdChave 
      Height          =   375
      Left            =   120
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
      _Version        =   720898
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmAtendimento.frx":0000
   End
   Begin XtremeSuiteControls.PushButton CmdVenda 
      Height          =   375
      Left            =   600
      TabIndex        =   32
      Top             =   6960
      Width           =   1695
      _Version        =   720898
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Venda"
      ForeColor       =   32768
      BackColor       =   16777152
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmAtendimento.frx":059A
   End
   Begin XtremeSuiteControls.TabControl TabAtend 
      Height          =   1935
      Left            =   120
      TabIndex        =   26
      Top             =   4920
      Width           =   6375
      _Version        =   720898
      _ExtentX        =   11245
      _ExtentY        =   3413
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.MultiRowJustified=   0   'False
      PaintManager.FixedTabWidth=   80
      ItemCount       =   1
      Item(0).Caption =   "Vendas"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "GrdVenda"
      Item(0).Control(1)=   "CmdEditar"
      Item(0).Control(2)=   "CmdVincular"
      Item(0).Control(3)=   "CmdNovo"
      Begin iGrid251_75B4A91C.iGrid GrdVenda 
         Height          =   1335
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2355
         BorderStyle     =   1
         HighlightBackColorNoFocus=   14737632
      End
      Begin XtremeSuiteControls.PushButton CmdEditar 
         Height          =   350
         Left            =   5160
         TabIndex        =   28
         Top             =   880
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "  &Editar"
         ForeColor       =   8388608
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":0934
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdVincular 
         Height          =   350
         Left            =   5160
         TabIndex        =   39
         Top             =   360
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "&Vincular"
         ForeColor       =   4210816
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":0A8E
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdNovo 
         Height          =   350
         Left            =   5160
         TabIndex        =   50
         Top             =   1380
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "  &Nova"
         ForeColor       =   8388608
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":0E28
         ImageAlignment  =   0
      End
   End
   Begin MSComctlLib.ImageList lstImage 
      Left            =   4320
      Top             =   6720
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
            Picture         =   "FrmAtendimento.frx":18F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtendimento.frx":1C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtendimento.frx":2226
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAtendimento.frx":25C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox GrpTempo 
      Height          =   1455
      Left            =   8760
      TabIndex        =   14
      Top             =   120
      Width           =   2535
      _Version        =   720898
      _ExtentX        =   4471
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Data / Hora"
      Appearance      =   4
      Begin XtremeSuiteControls.FlatEdit TxtHHINI 
         Height          =   315
         Left            =   1230
         TabIndex        =   19
         Top             =   600
         Width           =   1170
         _Version        =   720898
         _ExtentX        =   2064
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "00:00"
         BackColor       =   16777215
         MaxLength       =   5
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHFIM 
         Height          =   315
         Left            =   1230
         TabIndex        =   21
         Top             =   960
         Width           =   1170
         _Version        =   720898
         _ExtentX        =   2064
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "00:00"
         BackColor       =   16777215
         MaxLength       =   5
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker CmbDTATEND 
         Height          =   315
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   1455
         _Version        =   720898
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40356.1843055556
      End
      Begin XtremeSuiteControls.DateTimePicker CmbHHINI 
         Height          =   315
         Left            =   960
         TabIndex        =   18
         Top             =   600
         Width           =   1455
         _Version        =   720898
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         CustomFormat    =   "HH:mm"
         Format          =   3
         UpDown          =   -1  'True
         CurrentDate     =   40356.7954166667
      End
      Begin XtremeSuiteControls.DateTimePicker CmbHHFIM 
         Height          =   315
         Left            =   960
         TabIndex        =   22
         Top             =   960
         Width           =   1455
         _Version        =   720898
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         CustomFormat    =   "HH:mm"
         Format          =   3
         UpDown          =   -1  'True
         CurrentDate     =   40356
      End
      Begin VB.Label LblDuracao 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Duração : "
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   720
         TabIndex        =   44
         Top             =   1260
         Width           =   1725
      End
      Begin VB.Label LblHHFIM 
         AutoSize        =   -1  'True
         Caption         =   "Término"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label LblHHINI 
         AutoSize        =   -1  'True
         Caption         =   "Início"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   405
      End
      Begin VB.Label LblDTATEND 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   435
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpDados 
      Height          =   1455
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      _Version        =   720898
      _ExtentX        =   5741
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Dados"
      Appearance      =   4
      Begin XtremeSuiteControls.ComboBox CmbIDFUNCIONARIO 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1935
         _Version        =   720898
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox CmbIDSALA 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   600
         Width           =   1935
         _Version        =   720898
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox CmbIDMAQUINA 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   960
         Width           =   1935
         _Version        =   720898
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label LblFuncionario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Profissional"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   795
      End
      Begin VB.Label LblSala 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sala"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   315
      End
      Begin VB.Label LblMaquina 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Máquina"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   7560
      TabIndex        =   33
      Top             =   6960
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   9720
      TabIndex        =   35
      Top             =   6960
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpSessao 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _Version        =   720898
      _ExtentX        =   9128
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Cliente"
      Appearance      =   4
      Begin XtremeSuiteControls.FlatEdit TxtTEL1 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   1575
         _Version        =   720898
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   14737632
         Enabled         =   0   'False
         BackColor       =   14737632
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtNOME 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   2985
         _Version        =   720898
         _ExtentX        =   5265
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "Patricia Moreira"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdIDCLIENTE 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   825
         _Version        =   720898
         _ExtentX        =   1455
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Nome"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   10
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.FlatEdit TxtFOTOTIPO 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   1935
         _Version        =   720898
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   14737632
         Enabled         =   0   'False
         Text            =   "I - Branca +Clara"
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdLov 
         Height          =   315
         Left            =   4680
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
         _Version        =   720898
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":295A
      End
      Begin XtremeSuiteControls.FlatEdit TxtIDCLIENTE 
         Height          =   315
         Left            =   1200
         TabIndex        =   46
         Top             =   240
         Width           =   495
         _Version        =   720898
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   -2147483631
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "8888"
         BackColor       =   16777215
         Alignment       =   2
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label LblFototipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fototipo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label LblTel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Celular/Tel."
         Enabled         =   0   'False
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1035
      End
   End
   Begin XtremeSuiteControls.PushButton CmdFechar 
      Height          =   375
      Left            =   840
      TabIndex        =   36
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Fechamento"
      ForeColor       =   32768
      BackColor       =   65280
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmAtendimento.frx":2ADD
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDATENDIMENTO 
      Height          =   345
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   13430493
      BackColor       =   13430493
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1935
      Left            =   6600
      TabIndex        =   29
      Top             =   4920
      Width           =   4695
      _Version        =   720898
      _ExtentX        =   8281
      _ExtentY        =   3413
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      PaintManager.MultiRowJustified=   0   'False
      PaintManager.FixedTabWidth=   80
      ItemCount       =   2
      Item(0).Caption =   "Obs. do Atendimento  "
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "PushButton1"
      Item(0).Control(1)=   "PushButton2"
      Item(0).Control(2)=   "TxtOBS"
      Item(1).Caption =   "Todas Observações"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "TxtObsOutras"
      Item(1).Control(1)=   "CmdMaximized"
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   420
         Left            =   -64840
         TabIndex        =   41
         Top             =   1200
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "  &Editar"
         ForeColor       =   8388608
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":3077
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   420
         Left            =   -64840
         TabIndex        =   42
         Top             =   600
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "&Vincular"
         ForeColor       =   4210816
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":31D1
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit TxtObsOutras 
         Height          =   1335
         Left            =   -69880
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   4455
         _Version        =   720898
         _ExtentX        =   7858
         _ExtentY        =   2355
         _StockProps     =   77
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdMaximized 
         Height          =   225
         Left            =   -66160
         TabIndex        =   43
         Top             =   1700
         Visible         =   0   'False
         Width           =   495
         _Version        =   720898
         _ExtentX        =   873
         _ExtentY        =   397
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":356B
      End
      Begin XtremeSuiteControls.FlatEdit TxtOBS 
         Height          =   1335
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   4455
         _Version        =   720898
         _ExtentX        =   7858
         _ExtentY        =   2355
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   5520
      TabIndex        =   34
      Top             =   6960
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   32768
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl TabControl2 
      Height          =   3075
      Left            =   120
      TabIndex        =   23
      Top             =   1605
      Width           =   11175
      _Version        =   720898
      _ExtentX        =   19711
      _ExtentY        =   5424
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.MultiRowJustified=   0   'False
      PaintManager.FixedTabWidth=   80
      ItemCount       =   2
      Item(0).Caption =   " Serviços"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "PushButton3"
      Item(0).Control(1)=   "PushButton4"
      Item(0).Control(2)=   "GrdSESSAO"
      Item(1).Caption =   " Produtos"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "GrdPRODUTO"
      Item(1).Control(1)=   "LblProd"
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   420
         Left            =   -64840
         TabIndex        =   47
         Top             =   1200
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "  &Editar"
         ForeColor       =   8388608
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":372D
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   420
         Left            =   -64840
         TabIndex        =   48
         Top             =   600
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "&Vincular"
         ForeColor       =   4210816
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmAtendimento.frx":3887
         ImageAlignment  =   0
      End
      Begin iGrid251_75B4A91C.iGrid GrdSESSAO 
         Height          =   2655
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4683
         Appearance      =   0
         BorderStyle     =   1
         HighlightBackColorNoFocus=   14737632
      End
      Begin iGrid251_75B4A91C.iGrid GrdPRODUTO 
         Height          =   2535
         Left            =   -69880
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   4471
         BorderStyle     =   1
         HighlightBackColorNoFocus=   14737632
      End
      Begin VB.Label LblProd 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente levou 0 cemes"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -60520
         TabIndex        =   49
         Top             =   2880
         Visible         =   0   'False
         Width           =   1560
      End
   End
   Begin VB.Label LblUltSessao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Última Sessão: 01/01/2000 (45 dias)"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   45
      Top             =   4680
      Width           =   2610
   End
End
Attribute VB_Name = "FrmAtendimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Unload(Cancel As Integer)
Event Resize()
Event CmbHHINIChange()
Event CmbHHFIMChange()
Event CmbHHFIMValidate(Cancel As Boolean)
Event CmbIDFUNCIONARIOClick()
Event CmbIDSALAClick()
Event CmdIDCLIENTEClick()
Event CmdLovClick()
Event CmdOkClick()
Event CmdCancelClick()
Event CmdVendaClick()
Event CmdVendaDropDown()
Event CmdChaveClick()
Event CmdEditarClick()
Event CmdNovoClick()
Event CmdFecharClick()
Event CmdSalvarClick()
Event CmdVincularClick()
Event CmdMaximizedClick()

Event GrdSESSAOAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdSESSAOBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdSESSAOColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdSESSAOColHeaderDblClick(ByVal lCol As Long)
Event GrdSESSAODblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
Event GrdSESSAOMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdSESSAOLostFocus()
Event GrdSESSAORequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdSESSAOValidate(Cancel As Boolean)

Event GrdPRODUTOAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdPRODUTOBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdPRODUTOColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdPRODUTOMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdPRODUTOLostFocus()
Event GrdPRODUTORequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdPRODUTOValidate(Cancel As Boolean)

Event GrdVendaAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdVendaBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdVendaColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdVendaDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
Event GrdVendaMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdVendaLostFocus()
Event GrdVendaRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdVendaValidate(Cancel As Boolean)

Event TxtHHINILostFocus()
Event TxtHHFIMLostFocus()
Event TxtIDCLIENTELostFocus()
Event TxtNOMEChange()
Event TxtNOMEKeyPress(KeyAscii As Integer)
Event TxtTEL1LostFocus()
Private Sub CmbHHFIM_Change()
   RaiseEvent CmbHHFIMChange
End Sub
Private Sub CmbHHFIM_GotFocus()
   DoEvents
   If Me.TxtHHFIM.Enabled Then
      Me.TxtHHFIM.SetFocus
   End If
End Sub
Private Sub CmbHHFIM_Validate(Cancel As Boolean)
   RaiseEvent CmbHHFIMValidate(Cancel)
End Sub
Private Sub CmbHHINI_Change()
   RaiseEvent CmbHHINIChange
End Sub
Private Sub CmbHHINI_GotFocus()
   On Error Resume Next
   DoEvents
   If Me.ActiveControl Is Me.TxtHHFIM Then Exit Sub
   If Me.TxtHHINI.Enabled And Me.TxtHHINI.Visible Then
      Me.TxtHHINI.SetFocus
   End If
End Sub
Private Sub CmbIDFUNCIONARIO_Click()
   RaiseEvent CmbIDFUNCIONARIOClick
End Sub
Private Sub CmbIDSALA_Click()
   RaiseEvent CmbIDSALAClick
End Sub
Private Sub cmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdChave_Click()
   RaiseEvent CmdChaveClick
End Sub
Private Sub CmdEditar_Click()
   RaiseEvent CmdEditarClick
End Sub
Private Sub CmdFechar_Click()
   RaiseEvent CmdFecharClick
End Sub
Private Sub CmdIDCLIENTE_Click()
   Me.CmdIDCLIENTE.Enabled = False
   RaiseEvent CmdIDCLIENTEClick
   Me.CmdIDCLIENTE.Enabled = True
End Sub
Private Sub CmdLov_Click()
   RaiseEvent CmdLovClick
End Sub
Private Sub CmdMaximized_Click()
   RaiseEvent CmdMaximizedClick
End Sub
Private Sub CmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub
Private Sub cmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub

Private Sub CmdVenda_Click()
   RaiseEvent CmdVendaClick
End Sub
Private Sub CmdVenda_DropDown()
   RaiseEvent CmdVendaDropDown
End Sub

Private Sub CmdVincular_Click()
   RaiseEvent CmdVincularClick
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
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
Private Sub GrdPRODUTO_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdPRODUTOAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdPRODUTO_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdPRODUTOBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdPRODUTO_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdPRODUTOColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdPRODUTO_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   
   With Me.GrdPRODUTO
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
Private Sub GrdPRODUTO_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdPRODUTO.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdPRODUTO_LostFocus()
   RaiseEvent GrdPRODUTOLostFocus
End Sub
Private Sub GrdPRODUTO_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdPRODUTOMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdPRODUTO_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdPRODUTORequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdPRODUTO_Validate(Cancel As Boolean)
   RaiseEvent GrdPRODUTOValidate(Cancel)
End Sub
Private Sub GrdSESSAO_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdSESSAOAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdSESSAO_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdSESSAOBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdSESSAO_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdSESSAOColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdSESSAO_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdSESSAOColHeaderDblClick(lCol)
End Sub
Private Sub GrdSESSAO_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   
   With Me.GrdSESSAO
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
Private Sub GrdSESSAO_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   RaiseEvent GrdSESSAODblClick(lRow, lCol, bRequestEdit)
End Sub
Private Sub GrdSESSAO_LostFocus()
   RaiseEvent GrdSESSAOLostFocus
End Sub
Private Sub GrdSESSAO_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdSESSAOMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdSESSAO_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdSESSAORequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdSESSAO_Validate(Cancel As Boolean)
   RaiseEvent GrdSESSAOValidate(Cancel)
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
   RaiseEvent GrdVendaDblClick(lRow, lCol, bRequestEdit)
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
Private Sub GrdVenda_Validate(Cancel As Boolean)
   RaiseEvent GrdVendaValidate(Cancel)
End Sub

Private Sub TxtHHFIM_GotFocus()
   Me.TxtHHFIM.SelStart = 0
   Me.TxtHHFIM.SelLength = Len(Me.TxtHHFIM.Text)
End Sub
Private Sub TxtHHFIM_LostFocus()
   RaiseEvent TxtHHFIMLostFocus
End Sub
Private Sub TxtHHINI_GotFocus()
   Me.TxtHHINI.SelStart = 0
   Me.TxtHHINI.SelLength = Len(Me.TxtHHFIM.Text)
End Sub
Private Sub TxtHHINI_LostFocus()
   RaiseEvent TxtHHINILostFocus
End Sub
Private Sub TxtIDCLIENTE_LostFocus()
   RaiseEvent TxtIDCLIENTELostFocus
End Sub
Private Sub TxtNOME_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtTEL1_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtTEL1_LostFocus()
   RaiseEvent TxtTEL1LostFocus
End Sub

