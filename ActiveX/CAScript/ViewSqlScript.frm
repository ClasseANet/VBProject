VERSION 5.00
Begin VB.Form ViewSqlScript 
   AutoRedraw      =   -1  'True
   Caption         =   """Script"""
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
   Icon            =   "ViewSqlScript.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox TxtUsuario 
      Height          =   315
      Left            =   7440
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox TxtBanco 
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox TxtServidor 
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdOper 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6840
      TabIndex        =   13
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdOper 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   12
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox TxtScript 
      Height          =   4935
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   960
      Width           =   8775
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   0
      Pattern         =   "*.Sql"
      TabIndex        =   2
      Top             =   4080
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Lbl 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "ViewSqlScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOper_Click(Index As Integer)
   Select Case Index
      Case 0: Call ExecScript
      Case 1: Unload Me
   End Select
End Sub
Private Sub Dir1_Change()
   Me.TxtScript.Text = ""
   Me.File1.Path = Me.Dir1.Path
End Sub
Private Sub Drive1_Change()
   Me.TxtScript.Text = ""
   Me.Dir1.Path = Me.Drive1.Drive & "\"
End Sub
Private Sub TxtBanco_GotFocus()
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
Private Sub TxtSenha_GotFocus()
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
Private Sub TxtServidor_GotFocus()
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
Private Sub TxtUsuario_GotFocus()
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
Private Sub File1_Click()
   Dim lArq As String
   Dim lFile  As Integer
   Dim lLinha As String
   Me.TxtScript.Text = ""
   lArq = Me.File1.Path & "\" & Me.File1.FileName
   lFile = FreeFile()
   Open lArq For Input As #lFile
      Do While Not EOF(lFile)
         Line Input #lFile, lLinha
         Me.TxtScript.Text = Me.TxtScript.Text & vbNewLine & lLinha
      Loop
   Close #lFile
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
   End If
End Sub
Private Sub Form_Load()
   Me.Dir1.Path = App.Path
End Sub
Public Function VerificaCampos() As Boolean
   VerificaCampos = False
   If Trim(Me.TxtServidor.Text) = "" Then
      MsgBox "Servidor Inválido."
      Me.TxtServidor.SetFocus
      Exit Function
   End If
   If Trim(Me.TxtBanco.Text) = "" Then
      MsgBox "Banco de Dados Inválido."
      Me.TxtBanco.SetFocus
      Exit Function
   End If
   If Trim(Me.TxtUsuario.Text) = "" Then
      MsgBox "Usuário Inválido."
      Me.TxtBanco.SetFocus
      Exit Function
   End If
   If Trim(Me.TxtSenha.Text) = "" Then
      MsgBox "Senha Inválida."
      Me.TxtBanco.SetFocus
      Exit Function
   End If
   VerificaCampos = True
End Function
Public Function ExecScript() As Boolean
   Dim CollSql As Collection
   Dim Query  As String
   Dim StrAux As String
   Dim xDb As DS_BANCO
   
   StrAux = Me.TxtScript.Text
   Set CollSql = New Collection
   
   If Not VerificaCampos Then
      Exit Function
   End If
   
   While Trim(StrAux) <> ""
      If InStr(StrAux, vbNewLine) = 0 Then
         Query = Trim(StrAux)
         StrAux = ""
      Else
         Query = Trim(Mid(StrAux, 1, InStr(StrAux, vbNewLine) - 1))
         StrAux = Trim(Mid(StrAux, InStr(StrAux, vbNewLine) + 2))
      End If
      If Trim(Query) <> "" Then
         CollSql.Add Query
      End If
   Wend
   Set xDb = New DS_BANCO
   With xDb
      .isADO = True
      .dbTipo = SQL_SERVER
      .dbVersao = "7.0"
      .Server = Me.TxtServidor.Text
      .dbName = Me.TxtBanco.Text
      .UID = Me.TxtUsuario
      .PWD = Me.TxtSenha.Text
      .SrvConecta
   End With
   Set xDb = Nothing
   ExecScript = True
End Function
