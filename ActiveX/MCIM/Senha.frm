VERSION 5.00
Begin VB.Form FrmSenha 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   0
   ClientWidth     =   2550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "&Cancela"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox TxtIDUSU 
      Height          =   300
      Left            =   240
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox TxtSENHA 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   10
      PasswordChar    =   "#"
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox TxtSERVIDOR 
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox TxtBANCO 
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label LblError 
      BackStyle       =   0  'Transparent
      Caption         =   "Error"
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape Moldura 
      BackColor       =   &H80000007&
      BorderColor     =   &H80000010&
      BorderWidth     =   3
      Height          =   3255
      Left            =   15
      Top             =   15
      Width           =   105
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "USUÁRIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   500
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "SENHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   495
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVIDOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   495
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   495
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image ImgUsuário 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   240
      Picture         =   "Senha.frx":0000
      Top             =   200
      Width           =   240
   End
   Begin VB.Image ImgBanco 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   240
      Picture         =   "Senha.frx":0442
      Top             =   1995
      Width           =   240
   End
   Begin VB.Image ImgServidor 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   240
      Picture         =   "Senha.frx":0884
      Top             =   1425
      Width           =   240
   End
   Begin VB.Image ImgSenha 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   240
      Picture         =   "Senha.frx":257E
      Top             =   795
      Width           =   240
   End
End
Attribute VB_Name = "FrmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyPress(KeyAscii As Integer)
Event Load()
Event CmdCancelClick()
Event CmdOKClick()
Event TxtBANCOGotFocus()
Event TxtBANCOKeyPress(KeyAscii As Integer)
Event TxtIDUSUGotFocus()
Event TxtSENHAGotFocus()
Event TxtSENHAKeyPress(KeyAscii As Integer)
Event TxtSERVIDORGotFocus()
Private Sub CmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub cmdOK_Click()
   RaiseEvent CmdOKClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub ImgBanco_Click()
   Dim sAux As String
   sAux = UCase(InputBox("Favor entre com o nome do Banco de Dados.", "BANCO DE DADOS", Me.TxtBANCO.Text))
   If Trim(sAux) <> "" Then
      Me.TxtBANCO.Text = sAux
   End If
End Sub

Private Sub ImgServidor_DblClick()
   Dim sAux As String
   sAux = UCase(InputBox("Favor entre com o nome do Servidor.", "SERVIDOR", Me.TxtSERVIDOR.Text))
   If Trim(sAux) <> "" Then
      Me.TxtSERVIDOR.Text = sAux
   End If
End Sub
Private Sub TxtBANCO_GotFocus()
   RaiseEvent TxtBANCOGotFocus
End Sub
Private Sub TxtBANCO_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtBANCOKeyPress(KeyAscii)
End Sub
Private Sub TxtIDUSU_GotFocus()
   RaiseEvent TxtIDUSUGotFocus
End Sub
Private Sub TxtSENHA_GotFocus()
   RaiseEvent TxtSENHAGotFocus
End Sub
Private Sub TxtSENHA_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtSENHAKeyPress(KeyAscii)
End Sub
Private Sub TxtSERVIDOR_GotFocus()
   RaiseEvent TxtSERVIDORGotFocus
End Sub
