VERSION 5.00
Begin VB.Form FrmView 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image ImageAss 
      BorderStyle     =   1  'Fixed Single
      Height          =   1000
      Left            =   0
      Top             =   0
      Width           =   5300
   End
End
Attribute VB_Name = "FrmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event Resize()
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
