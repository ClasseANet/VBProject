VERSION 5.00
Begin VB.Form FrmIdentidade 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Registro Biométrico"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   4575
      Begin VB.Label LblDia 
         Alignment       =   2  'Center
         Caption         =   " Quarta-Feira, 02/04/2012"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
      Begin VB.Label LblHora 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "09:18h"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   240
      Top             =   240
      Width           =   1605
   End
   Begin VB.Label LblNome 
      Alignment       =   2  'Center
      Caption         =   "Diogenes Santos Ramos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label LblSaudacao 
      Caption         =   "Bom dia!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmIdentidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Timer1Timer()
Event CmdOkClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)

Public mvarClFinger As Object
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub Timer1_Timer()
   RaiseEvent Timer1Timer
   Static nVez
   nVez = nVez + 1
   If Me.LblHora.ForeColor = vbBlack Then
         
      If CDate(Replace(Me.LblHora.Caption, "h", "")) < CDate("09:10") Or (CDate(Replace(Me.LblHora.Caption, "h", "")) > CDate("19:00") And CDate(Replace(Me.LblHora.Caption, "h", "")) < CDate("19:10")) Then
         Me.LblHora.ForeColor = vbBlue
      Else
         Me.LblHora.ForeColor = vbRed
      End If
   Else
      Me.LblHora.ForeColor = vbBlack
   End If
   If nVez >= 6 Then
      nVez = 0
      Unload Me
   End If
End Sub
