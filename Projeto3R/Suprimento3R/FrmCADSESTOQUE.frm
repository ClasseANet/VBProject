VERSION 5.00
Begin VB.Form FrmCADSESTOQUE 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Locais de Estoque"
   ClientHeight    =   7995
   ClientLeft      =   6090
   ClientTop       =   1965
   ClientWidth     =   6585
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "FrmCADSESTOQUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Load()
Event Rezise()
Event Unload(Cancel As Integer)
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Rezise
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
