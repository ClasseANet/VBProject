VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.TaskPanel.v11.2.2.ocx"
Begin VB.Form FrmShortBar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   7725
   ClientLeft      =   2535
   ClientTop       =   2760
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   0
      ScaleHeight     =   3525
      ScaleWidth      =   4650
      TabIndex        =   0
      Top             =   720
      Width           =   4650
      Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
         Height          =   1215
         Left            =   360
         TabIndex        =   3
         Top             =   240
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
      Left            =   1560
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmShortBar.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmShortBar.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmShortBar.frx":07EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption SccContato 
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4680
      _Version        =   720898
      _ExtentX        =   8255
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Cadastro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption SccContato2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      _Version        =   720898
      _ExtentX        =   8281
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "Meu Cadastro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "FrmShortBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Resize()
Event WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Event wndTaskPanelFocusedItemChanged()
Event wndTaskPanelHotItemChanged()
Event CmdPerquisaClick()
Private Sub CmdPerquisar_Click()
   RaiseEvent CmdPerquisaClick
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      RaiseEvent CmdPerquisaClick
   End If
End Sub
Private Sub txtNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      RaiseEvent CmdPerquisaClick
   End If
End Sub
Private Sub txtTel_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      RaiseEvent CmdPerquisaClick
   End If
End Sub
Private Sub wndTaskPanel_FocusedItemChanged()
   RaiseEvent wndTaskPanelFocusedItemChanged
End Sub
Private Sub wndTaskPanel_HotItemChanged()
   RaiseEvent wndTaskPanelHotItemChanged
   
End Sub
Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   RaiseEvent WndTaskPanelItemClick(Item)
End Sub

