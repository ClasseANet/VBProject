VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmCADRPONTO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resgistro de Ponto"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   6135
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _Version        =   720898
      _ExtentX        =   11668
      _ExtentY        =   10821
      _StockProps     =   79
      Caption         =   " Registro "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
      Begin XtremeSuiteControls.CheckBox ChkFLGDIA 
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   1020
         Width           =   975
         _Version        =   720898
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Dia Útil"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit TxtIDMOVHH 
         Height          =   330
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   945
         _Version        =   720898
         _ExtentX        =   1667
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtCHAPA 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   1065
         _Version        =   720898
         _ExtentX        =   1879
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtNOME 
         Height          =   330
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   4065
         _Version        =   720898
         _ExtentX        =   7170
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtDTPONTO 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   1065
         _Version        =   720898
         _ExtentX        =   1879
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "88/88/8888"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHINI 
         Height          =   330
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   1065
         _Version        =   720898
         _ExtentX        =   1879
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "88:88"
         Alignment       =   2
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHFIM 
         Height          =   330
         Left            =   3840
         TabIndex        =   12
         Top             =   1320
         Width           =   945
         _Version        =   720898
         _ExtentX        =   1667
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHREFEICAO 
         Height          =   330
         Left            =   5880
         TabIndex        =   14
         Top             =   1320
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "1,00"
         Alignment       =   1
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHTRAB 
         Height          =   330
         Left            =   1320
         TabIndex        =   16
         Top             =   1680
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHESPERADO 
         Height          =   330
         Left            =   1320
         TabIndex        =   18
         Top             =   2040
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtSaldoParcial 
         Height          =   330
         Left            =   1320
         TabIndex        =   20
         Top             =   2400
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   5160
         Width           =   6255
         _Version        =   720898
         _ExtentX        =   11033
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   " Resultado "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit TxtSALDODIA 
            Height          =   330
            Left            =   2880
            TabIndex        =   30
            Top             =   360
            Width           =   825
            _Version        =   720898
            _ExtentX        =   1455
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Text            =   "3,50"
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtACUMULADO 
            Height          =   330
            Left            =   4800
            TabIndex        =   32
            Top             =   360
            Width           =   1185
            _Version        =   720898
            _ExtentX        =   2090
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "228,50"
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtAcumulado0 
            Height          =   330
            Left            =   960
            TabIndex        =   49
            Top             =   360
            Width           =   825
            _Version        =   720898
            _ExtentX        =   1455
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Text            =   "3,50"
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   330
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   840
            _Version        =   720898
            _ExtentX        =   1482
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Saldo Ant.:"
            Enabled         =   0   'False
         End
         Begin XtremeSuiteControls.Label Label17 
            Height          =   330
            Left            =   3960
            TabIndex        =   31
            Top             =   360
            Width           =   1080
            _Version        =   720898
            _ExtentX        =   1905
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Acumulado:"
         End
         Begin XtremeSuiteControls.Label Label16 
            Height          =   330
            Left            =   2040
            TabIndex        =   29
            Top             =   360
            Width           =   840
            _Version        =   720898
            _ExtentX        =   1482
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Saldo Dia:"
            Enabled         =   0   'False
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   6255
         _Version        =   720898
         _ExtentX        =   11033
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   " Abono "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox CmbIDABONO 
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Top             =   240
            Width           =   4815
            _Version        =   720898
            _ExtentX        =   8493
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            UseVisualStyle  =   -1  'True
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit TxtHHABONO 
            Height          =   330
            Left            =   1440
            TabIndex        =   25
            Top             =   600
            Width           =   585
            _Version        =   720898
            _ExtentX        =   1032
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblHHABONO 
            Height          =   330
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   1080
            _Version        =   720898
            _ExtentX        =   1905
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "HH Abonado:"
         End
         Begin XtremeSuiteControls.Label LblIDABONO 
            Height          =   330
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1080
            _Version        =   720898
            _ExtentX        =   1905
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Tipo de Abono:"
         End
      End
      Begin XtremeSuiteControls.CheckBox ChkFLGFALTA 
         Height          =   255
         Left            =   5400
         TabIndex        =   44
         Top             =   240
         Width           =   975
         _Version        =   720898
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Falta"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   1095
         Left            =   120
         TabIndex        =   26
         Top             =   3960
         Width           =   6255
         _Version        =   720898
         _ExtentX        =   11033
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   " Observação "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit TxtOBS 
            Height          =   690
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   5865
            _Version        =   720898
            _ExtentX        =   10345
            _ExtentY        =   1217
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox ChkFeriado 
         Height          =   255
         Left            =   4200
         TabIndex        =   51
         Top             =   240
         Width           =   975
         _Version        =   720898
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Feriado"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.Label Label14 
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   1080
         _Version        =   720898
         _ExtentX        =   1905
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Saldo Parcial:"
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1200
         _Version        =   720898
         _ExtentX        =   2117
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "HH Trabalhado:"
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   330
         Left            =   5160
         TabIndex        =   13
         Top             =   1320
         Width           =   720
         _Version        =   720898
         _ExtentX        =   1270
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Almoço:"
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1080
         _Version        =   720898
         _ExtentX        =   1905
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Expediente:"
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   330
         Left            =   2760
         TabIndex        =   11
         Top             =   1320
         Width           =   960
         _Version        =   720898
         _ExtentX        =   1693
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Hora Saída:"
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1080
         _Version        =   720898
         _ExtentX        =   1905
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Hora Entrada:"
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1080
         _Version        =   720898
         _ExtentX        =   1905
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Data :"
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1080
         _Version        =   720898
         _ExtentX        =   1905
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Funcionário :"
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   240
         _Version        =   720898
         _ExtentX        =   423
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "ID.:"
         Enabled         =   0   'False
      End
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   9360
      TabIndex        =   43
      Top             =   6480
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sair"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   7680
      TabIndex        =   42
      Top             =   6480
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salvar"
      ForeColor       =   16711680
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   6135
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   3975
      _Version        =   720898
      _ExtentX        =   7011
      _ExtentY        =   10821
      _StockProps     =   79
      Caption         =   " Informações "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
      Begin iGrid251_75B4A91C.iGrid GrdBatidas 
         Height          =   1575
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2778
         Editable        =   0   'False
         RowMode         =   -1  'True
      End
      Begin iGrid251_75B4A91C.iGrid GrdAtend 
         Height          =   2655
         Left            =   120
         TabIndex        =   39
         Top             =   3000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4683
         Editable        =   0   'False
         RowMode         =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHBATIDA0 
         Height          =   240
         Left            =   840
         TabIndex        =   36
         Top             =   2220
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   423
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "88:88"
         Alignment       =   2
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHBATIDA1 
         Height          =   240
         Left            =   3120
         TabIndex        =   37
         Top             =   2220
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   423
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "88:88"
         Alignment       =   2
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHATEND0 
         Height          =   240
         Left            =   840
         TabIndex        =   40
         Top             =   5700
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   423
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "88:88"
         Alignment       =   2
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtHHATEND1 
         Height          =   240
         Left            =   3120
         TabIndex        =   41
         Top             =   5700
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   423
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "88:88"
         Alignment       =   2
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdBatida0 
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2220
         Width           =   735
         _Version        =   720898
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Primeira:"
         FlatStyle       =   -1  'True
         Appearance      =   2
         MultiLine       =   0   'False
         ImageAlignment  =   4
         BorderGap       =   0
         ImageGap        =   0
      End
      Begin XtremeSuiteControls.PushButton CmdBatida1 
         Height          =   255
         Left            =   2400
         TabIndex        =   46
         Top             =   2220
         Width           =   735
         _Version        =   720898
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Última:"
         FlatStyle       =   -1  'True
         Appearance      =   2
         MultiLine       =   0   'False
         ImageAlignment  =   4
         BorderGap       =   0
         ImageGap        =   0
      End
      Begin XtremeSuiteControls.PushButton CmdAtend0 
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   5700
         Width           =   735
         _Version        =   720898
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Primeiro:"
         FlatStyle       =   -1  'True
         Appearance      =   2
         MultiLine       =   0   'False
         ImageAlignment  =   4
         BorderGap       =   0
         ImageGap        =   0
      End
      Begin XtremeSuiteControls.PushButton CmdAtend1 
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         Top             =   5700
         Width           =   735
         _Version        =   720898
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Último:"
         FlatStyle       =   -1  'True
         Appearance      =   2
         MultiLine       =   0   'False
         ImageAlignment  =   4
         BorderGap       =   0
         ImageGap        =   0
      End
      Begin XtremeSuiteControls.Label LblAtends 
         Height          =   330
         Left            =   120
         TabIndex        =   38
         Top             =   2640
         Width           =   1920
         _Version        =   720898
         _ExtentX        =   3387
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Atendimentos do Dia"
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
      Begin XtremeSuiteControls.Label LblBatidas 
         Height          =   330
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1440
         _Version        =   720898
         _ExtentX        =   2540
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Batidas do Dia"
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
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   52
      Top             =   6480
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Excluir"
      ForeColor       =   64
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADRPONTO.frx":0000
   End
End
Attribute VB_Name = "FrmCADRPONTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event ChkFeriadoClick()
Event ChkFeriadoMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event ChkFLGFALTAClick()
Event CmdSairClick()
Event CmdSalvarClick()
Event CmdExcluirClick()
Event CmbSENTIDOKeyPress(KeyAscii As Integer)
Event CmbUNIDADEKeyPress(KeyAscii As Integer)
Event CmdAtend1Click()
Event TxtCHAPALostFocus()
Event TxtCHAPAKeyPress(KeyAscii As Integer)
Event TxtCHAPAGotFocus()
Event TxtSENHALostFocus()
Event TxtSENHAKeyPress(KeyAscii As Integer)
Event TxtHHABONOLostFocus()
Event TxtHHESPERADOLostFocus()
Event TxtHHREFEICAOLostFocus()
Event TxtHHFIMLostFocus()
Event TxtHHINILostFocus()
Event Timer()
Private Sub CmbSENTIDO_KeyPress(KeyAscii As Integer)
   RaiseEvent CmbSENTIDOKeyPress(KeyAscii)
End Sub
Private Sub CmbUnidade_KeyPress(KeyAscii As Integer)
   RaiseEvent CmbSENTIDOKeyPress(KeyAscii)
End Sub

Private Sub ChkFeriado_Click()
   RaiseEvent ChkFeriadoClick
End Sub
Private Sub ChkFeriado_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent ChkFeriadoMouseUp(Button, Shift, x, y)
End Sub
Private Sub ChkFLGFALTA_Click()
   RaiseEvent ChkFLGFALTAClick
End Sub
Private Sub CmbIDABONO_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub CmdAtend0_Click()
   TxtHHINI.Text = TxtHHATEND0.Text
   Call TxtHHINI_LostFocus
End Sub
Private Sub CmdAtend1_Click()
   RaiseEvent CmdAtend1Click
End Sub
Private Sub CmdBatida0_Click()
   TxtHHINI.Text = TxtHHBATIDA0.Text
   Call TxtHHINI_LostFocus
End Sub
Private Sub CmdBatida1_Click()
   TxtHHFIM.Text = TxtHHBATIDA1.Text
   Call TxtHHFIM_LostFocus
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
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
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub TxtACUMULADO_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtACUMULADO_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtACUMULADO_LostFocus()
   Me.TxtACUMULADO.Text = ValBr(xVal(Me.TxtACUMULADO.Text))
End Sub
Private Sub TxtAcumulado0_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtAcumulado0_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtCHAPA_GotFocus()
   RaiseEvent TxtCHAPAGotFocus
End Sub
Private Sub TxtCHAPA_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
   RaiseEvent TxtCHAPAKeyPress(KeyAscii)
End Sub
Private Sub TxtCHAPA_LostFocus()
   RaiseEvent TxtCHAPALostFocus
End Sub
Private Sub TxtSENHA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtSENHA_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtSENHAKeyPress(KeyAscii)
End Sub
Private Sub TxtSENHA_LostFocus()
   RaiseEvent TxtSENHALostFocus
End Sub
Private Sub TxtDTPONTO_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtDTPONTO_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtHHABONO_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHABONO_LostFocus()
   RaiseEvent TxtHHABONOLostFocus
End Sub
Private Sub TxtHHATEND0_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHATEND0_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub

Private Sub TxtHHATEND1_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHBATIDA0_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHBATIDA1_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHESPERADO_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHESPERADO_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtHHESPERADO_LostFocus()
   RaiseEvent TxtHHESPERADOLostFocus
End Sub
Private Sub TxtHHFIM_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHFIM_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtHHFIM_LostFocus()
   RaiseEvent TxtHHFIMLostFocus
End Sub
Private Sub TxtHHINI_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHINI_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtHHINI_LostFocus()
   RaiseEvent TxtHHINILostFocus
End Sub
Private Sub TxtHHREFEICAO_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHREFEICAO_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtHHREFEICAO_LostFocus()
   RaiseEvent TxtHHREFEICAOLostFocus
End Sub
Private Sub TxtHHTRAB_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtHHTRAB_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtIDMOVHH_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtIDMOVHH_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtNOME_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub txtNome_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtSALDODIA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtSALDODIA_KeyPress(KeyAscii As Integer)
   Call SendTab(Me, KeyAscii)
End Sub
Private Sub TxtSaldoParcial_KeyPress(KeyAscii As Integer)
   Call SelecionarTexto(Me.ActiveControl)
End Sub
