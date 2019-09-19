VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form RunSqlScript 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Execução de Tarefas"
   ClientHeight    =   3615
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   7020
   Icon            =   "RunSqlScript.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   6615
      _Version        =   720898
      _ExtentX        =   11668
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483633
      Appearance      =   6
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   7335
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   840
         Width           =   7030
         Begin VB.Label LblStatus 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "(0) instruções executadas "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   2175
         End
      End
      Begin XtremeSuiteControls.PushButton CmdDetalhes 
         Height          =   315
         Left            =   5640
         TabIndex        =   6
         Top             =   360
         Width           =   1245
         _Version        =   720898
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Detalhes >>>"
         BackColor       =   -2147483633
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdOk 
         Height          =   435
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   1245
         _Version        =   720898
         _ExtentX        =   2196
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Atualizar"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.Label LblTitStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Conexão atual :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label LblTitulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Atualizar Base de Dados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3705
   End
   Begin VB.Label LblTexto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Para realizar esta operação você será conectado à Base de Dados, caso tenha problemas, redefina a conexão no menu acima."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label LblStatusBD 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Desconectado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Menu MnuConn 
      Caption         =   "Conexão..."
   End
End
Attribute VB_Name = "RunSqlScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event CmdOkClick()
Event LblStatusBDDblClick()
Event LblTitStatusDblClick()
Event MnuConnClick()
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub LblStatusBD_DblClick()
   RaiseEvent LblStatusBDDblClick
End Sub

Private Sub LblTitStatus_DblClick()
   RaiseEvent LblTitStatusDblClick
End Sub
Private Sub MnuConn_Click()
   RaiseEvent MnuConnClick
End Sub
