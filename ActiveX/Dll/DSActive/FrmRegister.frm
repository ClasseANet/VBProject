VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRegister 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Register"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   Icon            =   "FrmRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOper 
      Caption         =   "&Sair"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdOper 
      Caption         =   "UnRegister"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdOper 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Register"
      Height          =   375
      Index           =   0
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox TxtArq 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   0
      WhatsThisHelpID =   10244
      Width           =   990
   End
   Begin VB.Image ImgOpen 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   5040
      MouseIcon       =   "FrmRegister.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "FrmRegister.frx":0614
      Top             =   240
      Width           =   405
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Activate()
Event Load()
Event CmdOperClick(index As Integer)
Event ImgOpen()
Private Sub CmdOper_Click(index As Integer)
   RaiseEvent CmdOperClick(index)
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub ImgOpen_Click()
   RaiseEvent ImgOpen
End Sub
Private Sub ImgOpen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.ImgOpen.BorderStyle = 1
End Sub
Private Sub ImgOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.ImgOpen.BorderStyle = 0
End Sub
