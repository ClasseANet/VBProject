VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "CODEJO~2.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COC287~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPaneParam 
   BorderStyle     =   0  'None
   Caption         =   "Parâmetros"
   ClientHeight    =   7995
   ClientLeft      =   19350
   ClientTop       =   6255
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   0
      ScaleHeight     =   3525
      ScaleWidth      =   3690
      TabIndex        =   0
      Top             =   720
      Width           =   3690
      Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
         Height          =   1215
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _Version        =   720898
         _ExtentX        =   3413
         _ExtentY        =   2143
         _StockProps     =   64
         VisualTheme     =   6
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   1200
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneParam.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneParam.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption SccTit2 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3735
      _Version        =   720898
      _ExtentX        =   6588
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "Parâmetros do Sistema"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption SccTit1 
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3720
      _Version        =   720898
      _ExtentX        =   6562
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Parâmetros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPaneParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Resize()
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Event wndTaskPanelFocusedItemChanged()
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub wndTaskPanel_FocusedItemChanged()
   RaiseEvent wndTaskPanelFocusedItemChanged
End Sub
Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   RaiseEvent WndTaskPanelItemClick(Item)
End Sub

