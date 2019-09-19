VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSck 
   Caption         =   "frmSck"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   2250
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmSck.frx":0000
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock WinSck 
      Left            =   840
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Form included as a container for the Winsock control.
'
' The Class module instantiates the Winsock Control
' WithEvents and handles all events locally.
'
' Dynamically loading the Winsock Control at run time
' without utilizing a container Form is supported by
' the Class module, however there are currently unresolved
' deployment issues when using the 'formless' method.
'
' See the notation in the Sub Class_Initialize() in
' the class module for additional information.
'
Private Sub WinSck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub
