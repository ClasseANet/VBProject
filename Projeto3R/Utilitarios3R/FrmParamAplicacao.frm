VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmParamAplicacao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   15540
   ClientTop       =   1305
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GrpSenhaMestre 
      Height          =   2295
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
      _Version        =   720898
      _ExtentX        =   6800
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "Senha Operacional"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton CmdOkSenha 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   615
         _Version        =   720898
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ok"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtSenhaAntiga 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "123"
         PasswordChar    =   "$"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GrpSenhaNova 
         Height          =   1215
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   3255
         _Version        =   720898
         _ExtentX        =   5741
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Entre com a nova Senha"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit TxtSENHAMESTRE1 
            Height          =   315
            Left            =   1320
            TabIndex        =   10
            Top             =   360
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            Text            =   "123"
            PasswordChar    =   "$"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtSENHAMESTRE2 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   720
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            Text            =   "123"
            PasswordChar    =   "$"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblSENHAMESTRE2 
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Confirmação: "
            Enabled         =   0   'False
         End
         Begin XtremeSuiteControls.Label LblSENHAMESTRE1 
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Nova Senha: "
            Enabled         =   0   'False
         End
      End
      Begin XtremeSuiteControls.Label LblSenhaAnt 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         _Version        =   720898
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Senha Atual: "
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpBoxTop 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9375
      _Version        =   720898
      _ExtentX        =   16536
      _ExtentY        =   1508
      _StockProps     =   79
      BackColor       =   16777215
      Begin XtremeSuiteControls.Label LblTitulo 
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1380
         _Version        =   720898
         _ExtentX        =   2434
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Aplicação"
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
      Begin XtremeSuiteControls.Label LblDSCTitulo 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6855
         _Version        =   720898
         _ExtentX        =   12091
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Parâmetros relacionados à operacionalidade da aplicação"
         ForeColor       =   8421504
         UseMnemonic     =   0   'False
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpTela 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   9375
      _Version        =   720898
      _ExtentX        =   16536
      _ExtentY        =   8916
      _StockProps     =   79
      Begin XtremeSuiteControls.GroupBox GrpBackup 
         Height          =   1800
         Left            =   360
         TabIndex        =   13
         Top             =   2895
         Width           =   3855
         _Version        =   720898
         _ExtentX        =   6800
         _ExtentY        =   3175
         _StockProps     =   79
         Caption         =   " Cópia de Segurança"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkBackupDia 
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   2535
            _Version        =   720898
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " Realizar Cópia ao 'Fechar Dia'"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkBackupPergunta 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   3015
            _Version        =   720898
            _ExtentX        =   5318
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " Realizar pergunta antes da cópia"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkBackupEnd 
            Height          =   255
            Left            =   3240
            TabIndex        =   16
            Top             =   720
            Visible         =   0   'False
            Width           =   2535
            _Version        =   720898
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " Desligar máquina após a cópia."
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtPathBackup 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   3375
            _Version        =   720898
            _ExtentX        =   5953
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "C:\Tmp\P3R\Backup"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdPathBackup 
            Height          =   315
            Left            =   3480
            TabIndex        =   18
            Top             =   1320
            Width           =   315
            _Version        =   720898
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CommonDialog CommonDialog1 
            Left            =   3000
            Top             =   960
            _Version        =   720898
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Local"
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpSenhaGerente 
         Height          =   2295
         Left            =   4440
         TabIndex        =   19
         Top             =   480
         Width           =   3855
         _Version        =   720898
         _ExtentX        =   6800
         _ExtentY        =   4048
         _StockProps     =   79
         Caption         =   "Senha Gerencial"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton CmdOkSenhaGer 
            Height          =   375
            Left            =   2760
            TabIndex        =   20
            Top             =   360
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Ok"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtSenhaGerAntiga 
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Top             =   360
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "123"
            PasswordChar    =   "$"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GrpSenhaGerNova 
            Height          =   1215
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Width           =   3255
            _Version        =   720898
            _ExtentX        =   5741
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Entre com a nova Senha"
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.FlatEdit TxtSENHAGER1 
               Height          =   315
               Left            =   1320
               TabIndex        =   23
               Top             =   360
               Width           =   1335
               _Version        =   720898
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Enabled         =   0   'False
               Text            =   "123"
               PasswordChar    =   "$"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit TxtSENHAGER2 
               Height          =   315
               Left            =   1320
               TabIndex        =   24
               Top             =   720
               Width           =   1335
               _Version        =   720898
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Enabled         =   0   'False
               Text            =   "123"
               PasswordChar    =   "$"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.Label LblSENHAGER1 
               Height          =   375
               Left            =   240
               TabIndex        =   26
               Top             =   360
               Width           =   1095
               _Version        =   720898
               _ExtentX        =   1931
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Nova Senha: "
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.Label LblSENHAGER2 
               Height          =   375
               Left            =   240
               TabIndex        =   25
               Top             =   720
               Width           =   1095
               _Version        =   720898
               _ExtentX        =   1931
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Confirmação: "
               Enabled         =   0   'False
            End
         End
         Begin XtremeSuiteControls.Label LblSenhaGerAnt 
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   2175
            _Version        =   720898
            _ExtentX        =   3836
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Senha Atual: "
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpFechaDia 
         Height          =   1815
         Left            =   4440
         TabIndex        =   29
         Top             =   2880
         Width           =   3855
         _Version        =   720898
         _ExtentX        =   6800
         _ExtentY        =   3201
         _StockProps     =   79
         Caption         =   " Fechamento do Dia"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox ChkFechaParcial 
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   3135
            _Version        =   720898
            _ExtentX        =   5530
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Permitir fechamento parcial"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkFechaEnd 
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   3135
            _Version        =   720898
            _ExtentX        =   5530
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   " Desligar máquina após Fechamento."
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkFechaTelaAg 
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1080
            Width           =   3135
            _Version        =   720898
            _ExtentX        =   5530
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anexar tela de agenda"
            UseVisualStyle  =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "FrmParamAplicacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Activate()
Event Resize()
Event Unload()
Event ChkBackupDiaClick()
Event CmdOkSenha()
Event CmdOkSenhaGer()
Event CmdCancelar()
Event CmdPadrao()
Event CmdPathBackupClick()
Event TxtSenhaAntigaKeyPress(KeyAscii As Integer)
Event TxtSenhaAntigaKeyUp(KeyCode As Integer, Shift As Integer)
Event TxtSenhaAntigaLostFocus()
Event TxtSenhaAntigaGotFocus()
Event TxtSenhaGerAntigaKeyPress(KeyAscii As Integer)
Event TxtSenhaGerAntigaKeyUp(KeyCode As Integer, Shift As Integer)
Event TxtSenhaGerAntigaLostFocus()
Event TxtSenhaGerAntigaGotFocus()
Event TxtPathBackupGotFocus()
Event TxtPathBackupLostFocus()
Private Sub CmdCancelar_Click()
   RaiseEvent CmdCancelar
End Sub
Private Sub CheckBox1_Click()

End Sub
Private Sub ChkBackupDia_Click()
   RaiseEvent ChkBackupDiaClick
End Sub
Private Sub CmdOkSenha_Click()
   RaiseEvent CmdOkSenha
End Sub
Private Sub CmdPadrao_Click()
   RaiseEvent CmdPadrao
End Sub
Private Sub CmdOkSenhaGer_Click()
   RaiseEvent CmdOkSenhaGer
End Sub
Private Sub CmdPathBackup_Click()
   RaiseEvent CmdPathBackupClick
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
Private Sub TxtPathBackup_GotFocus()
   RaiseEvent TxtPathBackupGotFocus
End Sub
Private Sub TxtPathBackup_LostFocus()
   RaiseEvent TxtPathBackupLostFocus
End Sub
Private Sub TxtSenhaAntiga_GotFocus()
   RaiseEvent TxtSenhaAntigaGotFocus
End Sub
Private Sub TxtSenhaAntiga_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtSenhaAntigaKeyPress(KeyAscii)
End Sub
Private Sub TxtSenhaAntiga_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtSenhaAntigaKeyUp(KeyCode, Shift)
End Sub
Private Sub TxtSenhaAntiga_LostFocus()
   RaiseEvent TxtSenhaAntigaLostFocus
End Sub
Private Sub TxtSenhaGerAntiga_GotFocus()
   RaiseEvent TxtSenhaGerAntigaGotFocus
End Sub
Private Sub TxtSenhaGerAntiga_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtSenhaGerAntigaKeyPress(KeyAscii)
End Sub
Private Sub TxtSenhaGerAntiga_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtSenhaGerAntigaKeyUp(KeyCode, Shift)
End Sub
Private Sub TxtSenhaGerAntiga_LostFocus()
   RaiseEvent TxtSenhaGerAntigaLostFocus
End Sub

