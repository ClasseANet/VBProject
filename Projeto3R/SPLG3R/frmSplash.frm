VERSION 5.00
Object = "{E6C4280E-288E-41E1-B348-A0E583B65166}#1.1#0"; "AnimatedGif.ocx"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400040&
   BorderStyle     =   0  'None
   ClientHeight    =   2775
   ClientLeft      =   210
   ClientTop       =   1095
   ClientWidth     =   6330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      Top             =   -95
      Width           =   6345
      Begin AnimatedGif.AniGif AniGif1 
         Height          =   495
         Left            =   2880
         TabIndex        =   3
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
      End
      Begin VB.Label LblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Verificando atualização..."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label LblNMPROJ 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Projeto 3R"
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   6195
      End
      Begin VB.Image imgLogo 
         Height          =   1395
         Left            =   2580
         Picture         =   "frmSplash.frx":000C
         Top             =   1400
         Width           =   3720
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event Resize()
Event KeyPress(KeyAscii As Integer)
Event LblNMPROJClick()

Private Sub Form_Activate()
   Me.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   RaiseEvent Activate
   Me.Visible = True
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Me.MousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Sub

Private Sub LblNMPROJ_Click()
   RaiseEvent LblNMPROJClick
End Sub
