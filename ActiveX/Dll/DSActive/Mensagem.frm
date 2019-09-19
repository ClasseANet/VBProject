VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmMessage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin Threed.SSPanel PnlMessage 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Operação Realizada com Sucesso."
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
   End
End
Attribute VB_Name = "FrmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTime As Integer
Enum COR
   Cinza = &HC0C0C0
   Verde = &H8000&
   Vermelho = &H80&
   Preto = &H0&
   Branco = &H8000000E
End Enum
Public MsgPositiva As Boolean
Public Mensagem As String
Public NumPisca  As Integer
Private Sub Form_Load()
   Me.Left = 60
   Me.Top = Screen.Height - (Me.Height + 800)
   
   If NumPisca = 0 Then
      NumPisca = 2
   End If
   
   Me.PnlMessage.Caption = Mensagem
   DoEvents
End Sub

Private Sub Timer_Timer()
   DoEvents
   nTime = nTime + 1
   If NumPisca = 1 And nTime = 2 Then
      NumPisca = 0
   End If
   If NumPisca < 0 Then
      Me.PnlMessage.BackColor = IIf(MsgPositiva, COR.Verde, COR.Vermelho)
      Me.PnlMessage.ForeColor = COR.Branco
      Me.PnlMessage.Refresh
      Me.Refresh
      nTime = 0
      NumPisca = 0
      UnLoad Me
   End If
   If (nTime Mod 2) = 1 Then '* Cor
      Me.PnlMessage.BackColor = IIf(MsgPositiva, COR.Verde, COR.Vermelho)
      Me.PnlMessage.ForeColor = COR.Branco
   Else '* Neutro
      Me.PnlMessage.BackColor = COR.Cinza
      Me.PnlMessage.ForeColor = COR.Preto
   End If
   If nTime = NumPisca + 2 Then
      nTime = 0
      NumPisca = 0
      UnLoad Me
   End If
End Sub

