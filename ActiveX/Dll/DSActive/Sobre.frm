VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSobre 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre o programa"
   ClientHeight    =   5595
   ClientLeft      =   2625
   ClientTop       =   330
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Sobre"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox SSPanel1 
      Height          =   5292
      Left            =   60
      ScaleHeight     =   5235
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   60
      Width           =   6612
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&OK"
         Height          =   375
         Left            =   5280
         TabIndex        =   1
         Top             =   180
         Width           =   1185
      End
      Begin TabDlg.SSTab TabSobre 
         Height          =   4452
         Left            =   132
         TabIndex        =   2
         Top             =   708
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   7858
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   529
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
         TabCaption(0)   =   " Programa "
         TabPicture(0)   =   "Sobre.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbl(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbl(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "ImgEmpresa"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Timer1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   " Memória "
         TabPicture(1)   =   "Sobre.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label17"
         Tab(1).Control(1)=   "Label16"
         Tab(1).Control(2)=   "Label15"
         Tab(1).Control(3)=   "Label14"
         Tab(1).Control(4)=   "Label13"
         Tab(1).Control(5)=   "Label12"
         Tab(1).Control(6)=   "Label11"
         Tab(1).Control(7)=   "Label10"
         Tab(1).Control(8)=   "Label9"
         Tab(1).Control(9)=   "Label8"
         Tab(1).Control(10)=   "Label7"
         Tab(1).Control(11)=   "Label6"
         Tab(1).Control(12)=   "Label5"
         Tab(1).Control(13)=   "Label4"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "     Rede     "
         TabPicture(2)   =   "Sobre.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label18"
         Tab(2).Control(1)=   "Label19"
         Tab(2).Control(2)=   "Label20"
         Tab(2).Control(3)=   "Label21"
         Tab(2).Control(4)=   "Label22"
         Tab(2).Control(5)=   "Label23"
         Tab(2).Control(6)=   "Label24"
         Tab(2).Control(7)=   "Label25"
         Tab(2).Control(8)=   "Label26"
         Tab(2).Control(9)=   "Label27"
         Tab(2).Control(10)=   "Label28"
         Tab(2).Control(11)=   "Label29"
         Tab(2).ControlCount=   12
         TabCaption(3)   =   "Computador"
         TabPicture(3)   =   "Sobre.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label41"
         Tab(3).Control(1)=   "Label40"
         Tab(3).Control(2)=   "Label39"
         Tab(3).Control(3)=   "Label38"
         Tab(3).Control(4)=   "Label37"
         Tab(3).Control(5)=   "Label36"
         Tab(3).Control(6)=   "Label35"
         Tab(3).Control(7)=   "Label34"
         Tab(3).Control(8)=   "Label33"
         Tab(3).Control(9)=   "Label32"
         Tab(3).Control(10)=   "Label31"
         Tab(3).Control(11)=   "Label30"
         Tab(3).ControlCount=   12
         TabCaption(4)   =   "     Disco     "
         TabPicture(4)   =   "Sobre.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "SSTab2"
         Tab(4).Control(1)=   "CmbDrv"
         Tab(4).ControlCount=   2
         Begin VB.Timer Timer1 
            Interval        =   100
            Left            =   5880
            Top             =   3480
         End
         Begin VB.ComboBox CmbDrv 
            Height          =   315
            Left            =   -74832
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   420
            Width           =   2952
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   3612
            Left            =   -74820
            TabIndex        =   4
            Top             =   780
            Width           =   5964
            _ExtentX        =   10530
            _ExtentY        =   6376
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabHeight       =   529
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
            TabCaption(0)   =   "    Geral    "
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label54"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label53"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label52"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label51"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label50"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label49"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label48"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label47"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label46"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label45"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label44"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Label43"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label42"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).ControlCount=   13
            TabCaption(1)   =   "    Volume    "
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label66"
            Tab(1).Control(1)=   "Label65"
            Tab(1).Control(2)=   "Label64"
            Tab(1).Control(3)=   "Label63"
            Tab(1).Control(4)=   "Label62"
            Tab(1).Control(5)=   "Label61"
            Tab(1).Control(6)=   "Label60"
            Tab(1).Control(7)=   "Label59"
            Tab(1).Control(8)=   "Label58"
            Tab(1).Control(9)=   "Label57"
            Tab(1).Control(10)=   "Label56"
            Tab(1).Control(11)=   "Label55"
            Tab(1).ControlCount=   12
            Begin VB.Label Label42 
               Caption         =   "Label42"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   288
               Left            =   384
               TabIndex        =   29
               Top             =   672
               Width           =   5196
            End
            Begin VB.Label Label43 
               Caption         =   "Setores por cluster:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   384
               TabIndex        =   28
               Top             =   1152
               Width           =   2892
            End
            Begin VB.Label Label44 
               Caption         =   "Label44"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   3360
               TabIndex        =   27
               Top             =   1155
               Width           =   2220
            End
            Begin VB.Label Label45 
               Caption         =   "Bytes por setor:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   384
               TabIndex        =   26
               Top             =   1536
               Width           =   2892
            End
            Begin VB.Label Label46 
               Caption         =   "Label46"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   405
               Left            =   3360
               TabIndex        =   25
               Top             =   1530
               Width           =   2220
            End
            Begin VB.Label Label47 
               Caption         =   "Clusters livres:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   384
               TabIndex        =   24
               Top             =   1920
               Width           =   2892
            End
            Begin VB.Label Label48 
               Caption         =   "Label48"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   420
               Left            =   3360
               TabIndex        =   23
               Top             =   1920
               Width           =   2220
            End
            Begin VB.Label Label49 
               Caption         =   "Total de clusters:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   384
               TabIndex        =   22
               Top             =   2304
               Width           =   2892
            End
            Begin VB.Label Label50 
               Caption         =   "Label50"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   3360
               TabIndex        =   21
               Top             =   2310
               Width           =   2220
            End
            Begin VB.Label Label51 
               Caption         =   "Bytes de espaço total em disco:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   384
               TabIndex        =   20
               Top             =   2688
               Width           =   2892
            End
            Begin VB.Label Label52 
               Caption         =   "Label52"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   420
               Left            =   3360
               TabIndex        =   19
               Top             =   2685
               Width           =   2220
            End
            Begin VB.Label Label53 
               Caption         =   "Bytes de espaço livre em disco:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   384
               TabIndex        =   18
               Top             =   3072
               Width           =   2892
            End
            Begin VB.Label Label54 
               Caption         =   "Label54"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   420
               Left            =   3360
               TabIndex        =   17
               Top             =   3075
               Width           =   2220
            End
            Begin VB.Label Label55 
               Caption         =   "Nome do volume:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -74616
               TabIndex        =   16
               Top             =   672
               Width           =   3180
            End
            Begin VB.Label Label56 
               Caption         =   "Label56"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -71352
               TabIndex        =   15
               Top             =   672
               Width           =   1932
            End
            Begin VB.Label Label57 
               Caption         =   "Tamanho do nome de volume:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -74616
               TabIndex        =   14
               Top             =   1152
               Width           =   3180
            End
            Begin VB.Label Label58 
               Caption         =   "Label58"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -71352
               TabIndex        =   13
               Top             =   1152
               Width           =   1932
            End
            Begin VB.Label Label59 
               Caption         =   "Número serial do volume:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -74616
               TabIndex        =   12
               Top             =   1632
               Width           =   3180
            End
            Begin VB.Label Label60 
               Caption         =   "Label60"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -71352
               TabIndex        =   11
               Top             =   1632
               Width           =   1932
            End
            Begin VB.Label Label61 
               Caption         =   "Número máximo de caracteres nos nomes de arquivos:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -74616
               TabIndex        =   10
               Top             =   2112
               Width           =   3180
            End
            Begin VB.Label Label62 
               Caption         =   "Label62"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -71352
               TabIndex        =   9
               Top             =   2112
               Width           =   1932
            End
            Begin VB.Label Label63 
               Caption         =   "Tipo do sistema:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -74616
               TabIndex        =   8
               Top             =   2592
               Width           =   3180
            End
            Begin VB.Label Label64 
               Caption         =   "Label64"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -71352
               TabIndex        =   7
               Top             =   2592
               Width           =   1932
            End
            Begin VB.Label Label65 
               Caption         =   "Tamanho do tipo do sistema:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -74616
               TabIndex        =   6
               Top             =   3072
               Width           =   3180
            End
            Begin VB.Label Label66 
               Caption         =   "Label66"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   396
               Left            =   -71352
               TabIndex        =   5
               Top             =   3072
               Width           =   1932
            End
         End
         Begin VB.Image ImgEmpresa 
            BorderStyle     =   1  'Fixed Single
            Height          =   960
            Left            =   120
            Picture         =   "Sobre.frx":008C
            Stretch         =   -1  'True
            Top             =   720
            Width           =   1665
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Sistema de ..."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   22.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   615
            Index           =   0
            Left            =   1920
            TabIndex        =   73
            Top             =   480
            Width           =   4335
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Este sistema foi desenvolvido para a MARÍTIMA PETRÓLEO E ENGENHARIA LTDA. por Diogenes Santos Ramos. "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   1
            Left            =   1800
            TabIndex        =   72
            Top             =   1140
            Width           =   4455
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "(c)MARÍTIMA PETRÓLEO ENGENHARIA LTDA - Todos os Direitos reservados."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   435
            Index           =   2
            Left            =   1800
            TabIndex        =   71
            Top             =   1845
            Width           =   4455
         End
         Begin VB.Label Label4 
            Caption         =   "Percentual de memória em uso:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   435
            Left            =   -74670
            TabIndex        =   69
            Top             =   675
            Width           =   2655
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -71895
            TabIndex        =   68
            Top             =   675
            Width           =   2970
         End
         Begin VB.Label Label18 
            Caption         =   "Usuário:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   67
            Top             =   672
            Width           =   2184
         End
         Begin VB.Label Label19 
            Caption         =   "Label19"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72405
            TabIndex        =   66
            Top             =   675
            Width           =   3480
         End
         Begin VB.Label Label20 
            Caption         =   "Nome do computador:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   65
            Top             =   1188
            Width           =   2184
         End
         Begin VB.Label Label21 
            Caption         =   "Label21"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72405
            TabIndex        =   64
            Top             =   1185
            Width           =   3480
         End
         Begin VB.Label Label22 
            Caption         =   "Grupo de trabalho:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   63
            Top             =   1704
            Width           =   2184
         End
         Begin VB.Label Label23 
            Caption         =   "Label23"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72405
            TabIndex        =   62
            Top             =   1710
            Width           =   3480
         End
         Begin VB.Label Label24 
            Caption         =   "Rede primária:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   61
            Top             =   2220
            Width           =   2184
         End
         Begin VB.Label Label25 
            Caption         =   "Label25"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72405
            TabIndex        =   60
            Top             =   2220
            Width           =   3480
         End
         Begin VB.Label Label26 
            Caption         =   "Rede secundária:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   59
            Top             =   2736
            Width           =   2184
         End
         Begin VB.Label Label27 
            Caption         =   "Label27"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72405
            TabIndex        =   58
            Top             =   2730
            Width           =   3480
         End
         Begin VB.Label Label28 
            Caption         =   "Caminho:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   57
            Top             =   3252
            Width           =   2184
         End
         Begin VB.Label Label29 
            Caption         =   "Label29"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   600
            Left            =   -72405
            TabIndex        =   56
            Top             =   3255
            Width           =   3480
         End
         Begin VB.Label Label6 
            Caption         =   "Memória física total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   55
            Top             =   1188
            Width           =   2652
         End
         Begin VB.Label Label7 
            Caption         =   "Label7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -71895
            TabIndex        =   54
            Top             =   1185
            Width           =   2970
         End
         Begin VB.Label Label8 
            Caption         =   "Memória física disponível:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   53
            Top             =   1704
            Width           =   2652
         End
         Begin VB.Label Label9 
            Caption         =   "Label9"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -71895
            TabIndex        =   52
            Top             =   1710
            Width           =   2970
         End
         Begin VB.Label Label10 
            Caption         =   "Arquivo de paginação total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   51
            Top             =   2220
            Width           =   2652
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -71895
            TabIndex        =   50
            Top             =   2220
            Width           =   2970
         End
         Begin VB.Label Label12 
            Caption         =   "Arquivo de paginação disponível:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   435
            Left            =   -74670
            TabIndex        =   49
            Top             =   2730
            Width           =   2655
         End
         Begin VB.Label Label13 
            Caption         =   "Label13"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -71895
            TabIndex        =   48
            Top             =   2730
            Width           =   2970
         End
         Begin VB.Label Label14 
            Caption         =   "Memória virtual total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74640
            TabIndex        =   47
            Top             =   3252
            Width           =   2652
         End
         Begin VB.Label Label15 
            Caption         =   "Label15"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   480
            Left            =   -71895
            TabIndex        =   46
            Top             =   3255
            Width           =   2970
         End
         Begin VB.Label Label16 
            Caption         =   "Memória virtual disponível:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   45
            Top             =   3768
            Width           =   2652
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -71895
            TabIndex        =   44
            Top             =   3765
            Width           =   2970
         End
         Begin VB.Label Label30 
            Caption         =   "Fabricante do processador:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   -74670
            TabIndex        =   43
            Top             =   675
            Width           =   2355
         End
         Begin VB.Label Label31 
            Caption         =   "Label31"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72225
            TabIndex        =   42
            Top             =   675
            Width           =   3480
         End
         Begin VB.Label Label32 
            Caption         =   "Tipo do processador:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   41
            Top             =   1188
            Width           =   2184
         End
         Begin VB.Label Label33 
            Caption         =   "Label33"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72225
            TabIndex        =   40
            Top             =   1185
            Width           =   3480
         End
         Begin VB.Label Label34 
            Caption         =   "Resolução de vídeo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   39
            Top             =   1704
            Width           =   2184
         End
         Begin VB.Label Label35 
            Caption         =   "Label35"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   -72225
            TabIndex        =   38
            Top             =   1710
            Width           =   3480
         End
         Begin VB.Label Label36 
            Caption         =   "Driver de vídeo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   37
            Top             =   2220
            Width           =   2184
         End
         Begin VB.Label Label37 
            Caption         =   "Label37"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72225
            TabIndex        =   36
            Top             =   2220
            Width           =   3480
         End
         Begin VB.Label Label38 
            Caption         =   "Adaptador primário:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   35
            Top             =   2736
            Width           =   2184
         End
         Begin VB.Label Label39 
            Caption         =   "Label39"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   -72225
            TabIndex        =   34
            Top             =   2730
            Width           =   3480
         End
         Begin VB.Label Label40 
            Caption         =   "Adaptador secundário:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Left            =   -74664
            TabIndex        =   33
            Top             =   3252
            Width           =   2196
         End
         Begin VB.Label Label41 
            Caption         =   "Label41"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   600
            Left            =   -72225
            TabIndex        =   32
            Top             =   3255
            Width           =   3480
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1260
            Left            =   360
            TabIndex        =   31
            Top             =   2400
            Width           =   5610
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   420
            TabIndex        =   30
            Top             =   3840
            Width           =   5610
         End
      End
      Begin VB.Line Linha1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         X1              =   120
         X2              =   6468
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label LblAppName 
         Caption         =   "LblAppName"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   120
         TabIndex        =   70
         Top             =   120
         Width           =   5100
      End
   End
   Begin VB.Label LbleMail 
      BackStyle       =   0  'Transparent
      Caption         =   "dramos@osite.com.br"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4320
      TabIndex        =   74
      Top             =   5340
      Width           =   2295
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
' Nome:        SOBRE.FRM
' Propósito:   Formulário de informações sobre o programa,
'              sistema e o computador
' Parâmetros:
' Retorna:
' Criada em:   14/03/97
' Alterada em: 14/03/97
'----------------------------------------------------------------

Option Explicit

Event Load()

Dim MemStat As MemoryStatus
Dim MemoryLoad As Long
Dim TotalPhys As Long
Dim AvailPhys As Long
Dim TotalPageFile As Long
Dim AvailPageFile As Long
Dim TotalVirtual As Long
Dim AvailVirtual As Long
  
Dim disco As DISK_INFO
Dim volume As VOL_INFO

Dim sBuffer As String
Dim sDrives As String
Dim sDriveID As String
Dim lDrive As Long

Const DRIVE_ERROR = 0
Const DRIVE_NOTPRESENT = 1
Const DRIVE_REMOVABLE = 2
Const DRIVE_FIXED = 3
Const DRIVE_REMOTE = 4
Const DRIVE_CDROM = 5
Const DRIVE_RAMDISK = 6
Dim HighLighted As Boolean

Private Function ShowFreeSpace(sRootDir As String, disco As DISK_INFO) As DISK_INFO
  Dim nSectorsPerCluster As Long
  Dim nBytesPerSector As Long
  Dim nFreeClusters As Long
  Dim nTotalClusters As Long
  
  If GetDiskFreeSpace(sRootDir, _
    nSectorsPerCluster, _
    nBytesPerSector, _
    nFreeClusters, _
    nTotalClusters) Then

    disco.dwSectorsPerCluster = Format$(nSectorsPerCluster, "#,##0")
    disco.dwBytesPerSector = Format$(nBytesPerSector, "#,##0")
    disco.dwFreeClusters = Format$(nFreeClusters, "#,##0")
    disco.dwTotalClusters = Format$(nTotalClusters, "#,##0")
    disco.dwDiskTotal = Format$(disco.dwTotalClusters * disco.dwSectorsPerCluster * disco.dwBytesPerSector, "#,##0")
    disco.dwFreeDisk = Format$(disco.dwFreeClusters * disco.dwSectorsPerCluster * disco.dwBytesPerSector, "#,##0")
    
    ShowFreeSpace = disco
  End If

End Function
Private Function VolumeInfo(sRootPath As String, volume As VOL_INFO) As VOL_INFO
  Dim pstrVolName As String
  Dim plngVolNameSize As Long
  Dim plngVolSerialNum As Long
  Dim plngMaxFilenameLen As Long
  Dim plngSysFlags As Long
  Dim pstrSystemType As String
  Dim plngSysTypeSize As Long

  pstrVolName = Space$(256)
  pstrSystemType = Space$(32)
  plngSysTypeSize = CLng(Len(pstrSystemType))
  plngVolNameSize = CLng(Len(pstrVolName))
  
  If GetVolumeInformation(sRootPath, _
    pstrVolName, plngVolNameSize, plngVolSerialNum, _
    plngMaxFilenameLen, plngSysFlags, pstrSystemType, plngSysTypeSize) Then
  
    volume.dwstrRootPath = ClsSobre.ExtractNullTermString(Trim$(sRootPath))
    volume.dwstrVolName = ClsSobre.ExtractNullTermString(Trim$(pstrVolName))
    volume.dwlngVolNameSize = Format$(plngVolNameSize, "#,##0")
    volume.dwlngVolSerialNum = IIf(plngVolSerialNum <> 0, Format$(Hex(plngVolSerialNum), "@@@@-@@@@"), Trim$(0))
    volume.dwlngMaxFilenameLen = Format$(plngMaxFilenameLen, "#,##0")
    volume.dwlngSysFlags = Format$(plngSysFlags, "#,##0")
    volume.dwstrSystemType = ClsSobre.ExtractNullTermString(Trim$(pstrSystemType))
    volume.dwlngSysTypeSize = Format$(plngSysTypeSize, "#,##0")
    
    VolumeInfo = volume
  End If

End Function
Sub CmbDrv_Click()
  Dim sRoot As String, lResult As Long, lTotal As Long, lFree As Long

  On Error GoTo CmbDrv_Click_Error

  sRoot = Left$(CmbDrv.Text, 2)
  ShowFreeSpace sRoot & "\", disco
  VolumeInfo sRoot & "\", volume
  sRoot = Dir(sRoot)
  If IsEmpty(sRoot) Then
    Label44.Caption = "Não disponível."
    Label46.Caption = "Não disponível."
    Label48.Caption = "Não disponível."
    Label50.Caption = "Não disponível."
    Label52.Caption = "Não disponível."
    Label54.Caption = "Não disponível."
  
    Label56.Caption = "Não disponível."
    Label58.Caption = "Não disponível."
    Label60.Caption = "Não disponível."
    Label62.Caption = "Não disponível."
    Label64.Caption = "Não disponível."
    Label66.Caption = "Não disponível."
  Else
    Label44.Caption = Trim$(disco.dwSectorsPerCluster)
    Label46.Caption = Trim$(disco.dwBytesPerSector)
    Label48.Caption = Trim$(disco.dwFreeClusters)
    Label50.Caption = Trim$(disco.dwTotalClusters)
    Label52.Caption = Trim$(disco.dwDiskTotal)
    Label54.Caption = Trim$(disco.dwFreeDisk)
    
    Label56.Caption = volume.dwstrVolName
    Label58.Caption = volume.dwlngVolNameSize
    Label60.Caption = volume.dwlngVolSerialNum
    Label62.Caption = volume.dwlngMaxFilenameLen
    Label64.Caption = volume.dwstrSystemType
    Label66.Caption = volume.dwlngSysTypeSize
  End If
  
CmbDrv_Click_Exit:
  Exit Sub
  
CmbDrv_Click_Error:
  If Err = 52 Then
    Label44.Caption = "Não disponível."
    Label46.Caption = "Não disponível."
    Label48.Caption = "Não disponível."
    Label50.Caption = "Não disponível."
    Label52.Caption = "Não disponível."
    Label54.Caption = "Não disponível."
  
    Label56.Caption = "Não disponível."
    Label58.Caption = "Não disponível."
    Label60.Caption = "Não disponível."
    Label62.Caption = "Não disponível."
    Label64.Caption = "Não disponível."
    Label66.Caption = "Não disponível."

  Else
    MsgBox "Erro " & Format$(Err) & ": " & Error$ & " em CmbDrv_Click"
  End If
  Resume CmbDrv_Click_Exit
End Sub
Private Sub CmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub
Private Sub LbleMail_Click()
   ClsDsr.ExecuteLink "mailto:" & Me.LbleMail.Caption
End Sub

Private Sub LbleMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   Dim pt As PointAPI
   '* See where the cursor is.
   GetCursorPos pt
    
   '* Translate into window coordinates.
   ScreenToClient hWnd, pt

   ' See if we're still within the control.
   With Me.LbleMail
      If (pt.X * Screen.TwipsPerPixelX < .Left) Or _
         (pt.Y * Screen.TwipsPerPixelY < .Top) Or _
         (pt.X * Screen.TwipsPerPixelX > .Left + .Width) Or _
         (pt.Y * Screen.TwipsPerPixelY > .Top + .Height) Then
           HighLighted = False
           .ForeColor = &H8000&
           Timer1.Enabled = False
       Else
           LbleMail.ForeColor = &HC000&
       End If
    End With
End Sub
