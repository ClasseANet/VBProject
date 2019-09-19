VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmVerifVersao 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Verificar Versão..."
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6555
   Icon            =   "VerifVersao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox CmdSair 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   415
      Left            =   2640
      MouseIcon       =   "VerifVersao.frx":038A
      MousePointer    =   99  'Custom
      Picture         =   "VerifVersao.frx":0694
      ScaleHeight     =   420
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox CmdVerificar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5160
      MouseIcon       =   "VerifVersao.frx":0C16
      MousePointer    =   99  'Custom
      Picture         =   "VerifVersao.frx":0F20
      ScaleHeight     =   315
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   6615
   End
   Begin InetCtlsObjects.Inet InetFTP 
      Left            =   5760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Label LblStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Situação atual ...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label LblTexto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Para realizar esta operação, você precisará estar conectado à Internet."
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
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label LblTitulo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Atualizar Sistema"
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
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Derrubar Conexões"
      Index           =   0
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Filtro..."
      Index           =   1
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "mPopupSys"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "FrmVerifVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event MnuMainClick(Index As Integer)
Event LblStatusDblClick()
Event InetFTPStateChanged(ByVal State As Integer)
Event Activate()
Event Load()
Event Unload(Cancel As Integer)
Event CmdSairClick()
Event CmdVerificarClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event LblTituloDblClick()
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdVerificar_Click()
   RaiseEvent CmdVerificarClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Form_Resize()
   'this is necessary to assure that the minimized window is hidden
   If Me.WindowState = vbMinimized Then
      Me.Hide
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
Private Sub InetFTP_StateChanged(ByVal State As Integer)
   RaiseEvent InetFTPStateChanged(State)
End Sub
Private Sub LblStatus_DblClick()
   RaiseEvent LblStatusDblClick
End Sub
Private Sub LblTitulo_DblClick()
   RaiseEvent LblTituloDblClick
End Sub
Private Sub MnuMain_Click(Index As Integer)
   RaiseEvent MnuMainClick(Index)
End Sub
Private Sub mPopRestore_Click()
   'called when the user clicks the popup menu Restore command
   Dim Result As Long
   Me.WindowState = vbNormal
   Result = SetForegroundWindow(Me.hwnd)
   Me.Show
End Sub

