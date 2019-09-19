VERSION 5.00
Begin VB.Form FrmFlood 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOper 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frme 
      Caption         =   "Carregando..."
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label LblPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   180
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   405
      End
   End
End
Attribute VB_Name = "FrmFlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancel As Boolean
Public ProgBar As New CProgBar32
Private Sub CmdOper_Click()
   Cancel = True
   'Unload Me
End Sub
Private Sub CmdOper_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Initialize()
   Cancel = False
End Sub

Private Sub Form_Load()
   Cancel = False
   With ProgBar
      Set .Parent = Me
      .Create 90, 250, 200, 15
   End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
   ProgBar.DestroyProgBar
End Sub

