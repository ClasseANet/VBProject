VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{8153511F-FE57-47E0-A0A1-DBA712C97332}#1.0#0"; "MCIControl.ocx"
Begin VB.Form FrmMensagem 
   BackColor       =   &H00C0C0C0&
   Caption         =   " Conversa On-Line"
   ClientHeight    =   6045
   ClientLeft      =   570
   ClientTop       =   4410
   ClientWidth     =   9885
   Icon            =   "FrmCliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9885
   Begin VB.Timer Timer01 
      Interval        =   1000
      Left            =   8160
      Top             =   0
   End
   Begin MCIControls.MCIMenu MCIMenu1 
      Left            =   9240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      BackColor       =   -2147483633
      HaveComboBox    =   0   'False
      HaveCheckBox    =   0   'False
      HaveTextBox     =   0   'False
   End
   Begin MCIControls.MCIButton CmdEnviar 
      Height          =   280
      Left            =   6480
      TabIndex        =   1
      Top             =   5565
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Enviar"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.TextBox TxtEnviar 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4560
      Width           =   7575
   End
   Begin VB.TextBox TxtReceber 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   7575
   End
   Begin MSWinsockLib.Winsock WskMensagem 
      Index           =   0
      Left            =   8760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label LblEspera 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ".........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label LblAviso 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Não Conectado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Width           =   1125
   End
   Begin VB.Label LblLocalHost 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<LOCAL HOST>"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2640
      TabIndex        =   9
      Top             =   4320
      Width           =   1185
   End
   Begin VB.Label LblDe 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "De : "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   345
   End
   Begin VB.Label LblRemoteHost 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<REMOTE HOST>"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label LblServidor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   720
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   7590
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Index           =   2
      Left            =   120
      Top             =   1200
      Width           =   7590
   End
   Begin VB.Label LblPara 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Para : "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   465
   End
   Begin VB.Label LblCliente 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   4320
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EBE3&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   7590
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5EBE3&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Index           =   3
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   7590
   End
End
Attribute VB_Name = "FrmMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' &H00D8E9EC&
Option Explicit
Event Activate()
Event Load()
Event Unload(Cancel As Integer)
Event Resize()
Event Timer01()
Event CmdEnviarClick()
Event TxtEnviarKeyDown(KeyCode As Integer, Shift As Integer)
Event WskMensagemDataArrival(Index As Integer, ByVal bytesTotal As Long)
Event WskMensagemConnectionRequest(Index As Integer, ByVal RequestID As Long)
Event WskMensagemError(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Event WskMensagemSendComplete(Index As Integer)
Event WskMensagemSendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Event WskMensagemConnect(Index As Integer)
Event WskMensagemClose(Index As Integer)
Private Sub CmdEnviar_Click()
   RaiseEvent CmdEnviarClick
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
   RaiseEvent Unload(Cancel)
End Sub
Private Sub Timer01_Timer()
   RaiseEvent Timer01
End Sub
Private Sub TxtEnviar_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtEnviarKeyDown(KeyCode, Shift)
End Sub
Private Sub WskMensagem_Close(Index As Integer)
   If WskMensagem(Index).State = StateConstants.sckClosed Then
'MsgBox "Close"
      RaiseEvent WskMensagemClose(Index)
   End If
End Sub
Private Sub WskMensagem_Connect(Index As Integer)
'MsgBox "Conect"
   RaiseEvent WskMensagemConnect(Index)
End Sub
Private Sub WskMensagem_ConnectionRequest(Index As Integer, ByVal RequestID As Long)
 'MsgBox RequestID
   RaiseEvent WskMensagemConnectionRequest(Index, RequestID)
End Sub
Private Sub WskMensagem_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   RaiseEvent WskMensagemDataArrival(Index, bytesTotal)
End Sub
Private Sub WskMensagem_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   MsgBox "Erro [" & Source & "] : " & Number & " - " & Description
   RaiseEvent WskMensagemError(Index, Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub
Private Sub WskMensagem_SendComplete(Index As Integer)
'MsgBox "Sended"
   RaiseEvent WskMensagemSendComplete(Index)
End Sub
Private Sub WskMensagem_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'MsgBox "Send Progress"
   RaiseEvent WskMensagemSendProgress(Index, bytesSent, bytesRemaining)
End Sub
