VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmParamSMS 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Parâmetros de Comunicação"
   ClientHeight    =   9900
   ClientLeft      =   5055
   ClientTop       =   585
   ClientWidth     =   9495
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GrpBoxSocioCom 
      Height          =   2415
      Left            =   120
      TabIndex        =   37
      Top             =   6480
      Width           =   9255
      _Version        =   720898
      _ExtentX        =   16325
      _ExtentY        =   4260
      _StockProps     =   79
      Caption         =   "Serviços"
      UseVisualStyle  =   -1  'True
      Appearance      =   4
      Begin XtremeSuiteControls.FlatEdit TxtLstEMailSocio 
         Height          =   315
         Left            =   1680
         TabIndex        =   39
         Top             =   720
         Width           =   6000
         _Version        =   720898
         _ExtentX        =   10583
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "diogenes72@bol.com.br"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtLstCelSocio 
         Height          =   315
         Left            =   1680
         TabIndex        =   41
         Top             =   1080
         Width           =   6000
         _Version        =   720898
         _ExtentX        =   10583
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "2178344618"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtDPILSuporte 
         Height          =   315
         Left            =   1680
         TabIndex        =   50
         Top             =   1920
         Width           =   2760
         _Version        =   720898
         _ExtentX        =   4868
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "suporte@dpilbrasil.com.br"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox ChkFRANQCOM 
         Height          =   375
         Left            =   240
         TabIndex        =   52
         Top             =   1440
         Width           =   7455
         _Version        =   720898
         _ExtentX        =   13150
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Habilite o serviço de comunicação exclusiva com a Franqueadora Dpil Brasil"
         ForeColor       =   16576
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         Appearance      =   5
      End
      Begin XtremeSuiteControls.CheckBox ChkSOCIOCOM 
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   7455
         _Version        =   720898
         _ExtentX        =   13150
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Habilite o serviço de comunicação exclusiva com os Sócios."
         ForeColor       =   16576
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         Appearance      =   5
      End
      Begin XtremeSuiteControls.CheckBox ChkExibirSenhas 
         Height          =   255
         Left            =   7440
         TabIndex        =   54
         Top             =   2040
         Width           =   1575
         _Version        =   720898
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Exibir Senhas."
         ForeColor       =   4210752
         BackColor       =   16777215
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   5
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblDPILSuporte 
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   1920
         Width           =   1065
         _Version        =   720898
         _ExtentX        =   1879
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "e-Mail Suporte:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblLstCelSocio 
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   1290
         _Version        =   720898
         _ExtentX        =   2275
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Lista de Celulares:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblLstEMailSocio 
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   1140
         _Version        =   720898
         _ExtentX        =   2011
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Lista de e-Mails:"
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpBoxSMS 
      Height          =   1575
      Left            =   120
      TabIndex        =   28
      Top             =   4800
      Width           =   9255
      _Version        =   720898
      _ExtentX        =   16325
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "SMS"
      UseVisualStyle  =   -1  'True
      Appearance      =   4
      Begin XtremeSuiteControls.FlatEdit TxtSMSURLBASE 
         Height          =   315
         Left            =   1080
         TabIndex        =   31
         Top             =   720
         Width           =   6600
         _Version        =   720898
         _ExtentX        =   11642
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "http://www.comtele.com.br/sms/api/api_fuse_connection.php?fuse=send_msg"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtSMSUID 
         Height          =   315
         Left            =   1080
         TabIndex        =   33
         Top             =   1080
         Width           =   1800
         _Version        =   720898
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "MjE0"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtSMSPWD 
         Height          =   315
         Left            =   3840
         TabIndex        =   35
         Top             =   1080
         Width           =   1800
         _Version        =   720898
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "123456"
         BackColor       =   16777215
         PasswordChar    =   "#"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdTesteSMS 
         Height          =   375
         Left            =   7920
         TabIndex        =   36
         Top             =   1080
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Testar SMS"
         ForeColor       =   12582912
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblTitSMS 
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   7455
         _Version        =   720898
         _ExtentX        =   13150
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " Especifique parâmetros básicos do Serviço de Mensagens Instantâneas."
         ForeColor       =   8421504
         BackColor       =   14737632
      End
      Begin XtremeSuiteControls.Label LBlSMSPWD 
         Height          =   375
         Left            =   3120
         TabIndex        =   34
         Top             =   1080
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Senha:"
      End
      Begin XtremeSuiteControls.Label LblUID 
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Usuário:"
      End
      Begin XtremeSuiteControls.Label LblURL 
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "URL Base:"
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpBoxBotton 
      Height          =   975
      Left            =   0
      TabIndex        =   42
      Top             =   9000
      Width           =   9495
      _Version        =   720898
      _ExtentX        =   16748
      _ExtentY        =   1720
      _StockProps     =   79
      BackColor       =   -2147483639
      Begin XtremeSuiteControls.PushButton CmdCancelar 
         Height          =   375
         Left            =   5160
         TabIndex        =   45
         Top             =   360
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Cancelar"
         ForeColor       =   192
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdOk 
         Height          =   375
         Left            =   3600
         TabIndex        =   44
         Top             =   360
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Salvar"
         ForeColor       =   12582912
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdPadrao 
         Height          =   375
         Left            =   6600
         TabIndex        =   46
         Top             =   360
         Width           =   1455
         _Version        =   720898
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Padrão"
         ForeColor       =   4210752
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.TabControlPage TabPgBotton 
         Height          =   855
         Left            =   0
         TabIndex        =   43
         Top             =   120
         Width           =   9375
         _Version        =   720898
         _ExtentX        =   16536
         _ExtentY        =   1508
         _StockProps     =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpBoxInternet 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9255
      _Version        =   720898
      _ExtentX        =   16325
      _ExtentY        =   6800
      _StockProps     =   79
      Caption         =   "INTERNET"
      UseVisualStyle  =   -1  'True
      Appearance      =   4
      Begin XtremeSuiteControls.GroupBox GrpBoxeMail 
         Height          =   1695
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   8775
         _Version        =   720898
         _ExtentX        =   15478
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Conta de e-Mail"
         UseVisualStyle  =   -1  'True
         Appearance      =   4
         Begin XtremeSuiteControls.FlatEdit TxtSMTPPort 
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   960
            Width           =   600
            _Version        =   720898
            _ExtentX        =   1058
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "587"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtSMTPHost 
            Height          =   315
            Left            =   960
            TabIndex        =   6
            Top             =   600
            Width           =   2760
            _Version        =   720898
            _ExtentX        =   4868
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "smtps.bol.com.br"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtPOP3Host 
            Height          =   315
            Left            =   960
            TabIndex        =   4
            Top             =   240
            Width           =   2760
            _Version        =   720898
            _ExtentX        =   4868
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "pop3.bol.com.br"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdTesteMail 
            Height          =   375
            Left            =   7560
            TabIndex        =   17
            Top             =   960
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Testar e-Mail"
            ForeColor       =   12582912
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtMailUID 
            Height          =   315
            Left            =   4680
            TabIndex        =   12
            Top             =   600
            Width           =   2760
            _Version        =   720898
            _ExtentX        =   4868
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "diogenes72@bol.com.br"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtMailPWD 
            Height          =   315
            Left            =   4680
            TabIndex        =   14
            Top             =   960
            Width           =   2760
            _Version        =   720898
            _ExtentX        =   4868
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "dolphin7"
            BackColor       =   16777215
            PasswordChar    =   "#"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtFromDisplayName 
            Height          =   315
            Left            =   4680
            TabIndex        =   10
            Top             =   240
            Width           =   2760
            _Version        =   720898
            _ExtentX        =   4868
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "D'pil Freguesia/RJ"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkUseAuthentication 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   2535
            _Version        =   720898
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Servidor requer autenticação"
            ForeColor       =   16576
            BackColor       =   16777215
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   5
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkUsePopAuthentication 
            Height          =   255
            Left            =   3840
            TabIndex        =   16
            Top             =   1320
            Width           =   1935
            _Version        =   720898
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Autenticação POP"
            ForeColor       =   16576
            BackColor       =   16777215
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   5
            Value           =   1
         End
         Begin XtremeSuiteControls.Label LblSMTPPort 
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Porta:"
         End
         Begin XtremeSuiteControls.Label LblFromDisplayName 
            Height          =   375
            Left            =   3840
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Nome:"
         End
         Begin XtremeSuiteControls.Label LblMailUID 
            Height          =   375
            Left            =   3840
            TabIndex        =   11
            Top             =   600
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Usuário:"
         End
         Begin XtremeSuiteControls.Label LblMailPWD 
            Height          =   375
            Left            =   3840
            TabIndex        =   13
            Top             =   960
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Senha:"
         End
         Begin XtremeSuiteControls.Label LblSMTPHost 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "SMTP:"
         End
         Begin XtremeSuiteControls.Label LblPOP3Host 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "POP:"
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpBoxFTP 
         Height          =   1215
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   8775
         _Version        =   720898
         _ExtentX        =   15478
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Conta de FTP"
         UseVisualStyle  =   -1  'True
         Appearance      =   4
         Begin XtremeSuiteControls.FlatEdit TxtFTP 
            Height          =   315
            Left            =   600
            TabIndex        =   20
            Top             =   360
            Width           =   3240
            _Version        =   720898
            _ExtentX        =   5715
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "ftp.classeanet.com.br"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtFtpUID 
            Height          =   315
            Left            =   4800
            TabIndex        =   24
            Top             =   360
            Width           =   2640
            _Version        =   720898
            _ExtentX        =   4657
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "freguesia"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtFtpPWD 
            Height          =   315
            Left            =   4800
            TabIndex        =   26
            Top             =   720
            Width           =   2640
            _Version        =   720898
            _ExtentX        =   4657
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "tresrios10"
            BackColor       =   16777215
            PasswordChar    =   "#"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtFTPBakPath 
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Top             =   720
            Width           =   2400
            _Version        =   720898
            _ExtentX        =   4233
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   16777215
            Text            =   "/Banco/"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdTesteFTP 
            Height          =   375
            Left            =   7560
            TabIndex        =   27
            Top             =   720
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Testar FTP"
            ForeColor       =   12582912
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblFTPBakPath 
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Pasta de Backup:"
         End
         Begin XtremeSuiteControls.Label LblFtpUID 
            Height          =   375
            Left            =   4080
            TabIndex        =   23
            Top             =   360
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Usuário:"
         End
         Begin XtremeSuiteControls.Label LblFtpPWD 
            Height          =   375
            Left            =   4080
            TabIndex        =   25
            Top             =   720
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Senha:"
         End
         Begin XtremeSuiteControls.Label LblFTP 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "FTP:"
         End
      End
      Begin XtremeSuiteControls.Label LblTitInternet 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   7455
         _Version        =   720898
         _ExtentX        =   13150
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " Especifique as configurações de comunicação pela Internet (e-mail , FTP...)."
         ForeColor       =   8421504
         BackColor       =   14737632
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpBoxTop 
      Height          =   855
      Left            =   0
      TabIndex        =   47
      Top             =   -120
      Width           =   9375
      _Version        =   720898
      _ExtentX        =   16536
      _ExtentY        =   1508
      _StockProps     =   79
      BackColor       =   16777215
      Begin XtremeSuiteControls.Label LblDSCTitulo 
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   6855
         _Version        =   720898
         _ExtentX        =   12091
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Configuração para o serviços de comunicação do Sistema."
         ForeColor       =   8421504
         UseMnemonic     =   0   'False
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblTitulo 
         Height          =   360
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   1920
         _Version        =   720898
         _ExtentX        =   3387
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Comunicação"
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
End
Attribute VB_Name = "FrmParamSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Activate()
Event Resize()
Event Unload()
Event CmdOk()
Event CmdCancelar()
Event CmdPadrao()
Event ChkFRANQCOMClick()
Event ChkSOCIOCOMClick()
Event ChkExibirSenhasClick()
Event CmdTesteFTPClick()
Event CmdTesteMailClick()
Event CmdTesteSMSClick()
Private Sub ChkExibirSenhas_Click()
   RaiseEvent ChkExibirSenhasClick
End Sub
Private Sub ChkFRANQCOM_Click()
   RaiseEvent ChkFRANQCOMClick
End Sub
Private Sub ChkSOCIOCOM_Click()
   RaiseEvent ChkSOCIOCOMClick
End Sub
Private Sub CmdCancelar_Click()
   RaiseEvent CmdCancelar
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOk
End Sub
Private Sub CmdPadrao_Click()
   RaiseEvent CmdPadrao
End Sub
Private Sub CmdTesteFTP_Click()
   RaiseEvent CmdTesteFTPClick
End Sub
Private Sub CmdTesteMail_Click()
   RaiseEvent CmdTesteMailClick
End Sub
Private Sub CmdTesteSMS_Click()
   RaiseEvent CmdTesteSMSClick
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

