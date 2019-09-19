VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.TaskPanel.v11.2.2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Begin VB.Form FrmPaneTarefa 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Contatos"
   ClientHeight    =   6030
   ClientLeft      =   7965
   ClientTop       =   2865
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   0
      ScaleHeight     =   3525
      ScaleWidth      =   4650
      TabIndex        =   1
      Top             =   720
      Width           =   4650
      Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
         Height          =   2415
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   3975
         _Version        =   720898
         _ExtentX        =   7011
         _ExtentY        =   4260
         _StockProps     =   64
         VisualTheme     =   6
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   1200
      Top             =   4680
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
            Picture         =   "FrmPaneTarefa.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneTarefa.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   2640
      Top             =   4800
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   2
      VisualTheme     =   2
   End
   Begin XtremeShortcutBar.ShortcutCaption SccConta2 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   4695
      _Version        =   720898
      _ExtentX        =   8281
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "Conta Caixa"
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
   Begin XtremeShortcutBar.ShortcutCaption SccConta 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _Version        =   720898
      _ExtentX        =   8255
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Financeiro"
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
End
Attribute VB_Name = "FrmPaneTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Resize()
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Event SccConta2MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Event wndTaskPanelFocusedItemChanged()
Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   RaiseEvent CommandBarsExecute(Control)
End Sub
Private Sub Form_Activate()
   Me.CommandBars.DeleteAll
End Sub

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
Private Sub SccConta2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent SccConta2MouseDown(Button, Shift, x, y)
End Sub

Private Sub wndTaskPanel_FocusedItemChanged()
   RaiseEvent wndTaskPanelFocusedItemChanged
End Sub
Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   RaiseEvent WndTaskPanelItemClick(Item)
End Sub
