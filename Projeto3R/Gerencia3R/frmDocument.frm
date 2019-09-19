VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   HelpContextID   =   100
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmDocument.frx":0000
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fMain As MDI
Public nDoc  As Long

Private Sub Form_Load()
    Top = -200
    Left = -200
    Width = 31000
    Height = 31000
    Form_Resize
    Me.Caption = "Documento " & nDoc
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 0, 0, Me.ScaleWidth - 0, Me.ScaleHeight - 0
    rtfText.RightMargin = rtfText.Width - 400
End Sub

Private Sub rtfText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     
    If (Button = 2) Then
        ZOrder
        
        Dim Popup As CommandBar
        Set Popup = fMain.CommandBars.ContextMenus.Find(202)
    
        Popup.ShowPopup
    End If


End Sub
