VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "COF0B3~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "CODEJO~2.OCX"
Begin VB.Form FrmShortBar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   7725
   ClientLeft      =   2535
   ClientTop       =   2760
   ClientWidth     =   4185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3480
      Top             =   720
   End
   Begin XtremeShortcutBar.ShortcutBar ScbMain 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _Version        =   720898
      _ExtentX        =   5741
      _ExtentY        =   11033
      _StockProps     =   64
   End
   Begin XtremeCommandBars.ImageManager ImgToobar 
      Left            =   3600
      Top             =   120
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmShortBar.frx":0000
   End
   Begin XtremeCommandBars.ImageManager ImgShortcutBar 
      Left            =   960
      Top             =   6600
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmShortBar.frx":9D0A
   End
End
Attribute VB_Name = "FrmShortBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event Resize()
Event Terminate()
Event ScbMainClientSizeChanged()
Event ScbMainExpandButtonDown(CancelMenu As Boolean)
Event ScbMainSelectedChanged(ByVal Item As XtremeShortcutBar.IShortcutBarItem)
Event Timer1Timer()
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
'Private Sub Form_KeyPress(KeyAscii As Integer)
'   KeyAscii = KeyAscii
'End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub Form_Terminate()
   RaiseEvent Terminate
End Sub
Private Sub ScbMain_ClientSizeChanged()
   RaiseEvent ScbMainClientSizeChanged
End Sub
Private Sub ScbMain_ExpandButtonDown(CancelMenu As Boolean)
   RaiseEvent ScbMainExpandButtonDown(CancelMenu)
End Sub
Private Sub ScbMain_SelectedChanged(ByVal Item As XtremeShortcutBar.IShortcutBarItem)
   RaiseEvent ScbMainSelectedChanged(Item)
End Sub

Private Sub Timer1_Timer()
   RaiseEvent Timer1Timer
End Sub
