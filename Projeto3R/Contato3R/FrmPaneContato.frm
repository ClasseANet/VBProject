VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.TaskPanel.v11.2.2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPaneContato 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   15855
   ClientTop       =   2250
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox FraFiltro 
      Height          =   960
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   1693
      _StockProps     =   79
      Caption         =   "GroupBox1"
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox ChkAtivo 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ativos"
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
         Appearance      =   5
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox ChkInativo 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Inativos"
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
         Appearance      =   5
      End
      Begin XtremeSuiteControls.CheckBox ChkEmEspera 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Em Espera"
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
         Appearance      =   5
         Value           =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox fraBuscaDetalhada 
      Height          =   2400
      Left            =   2640
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   4233
      _StockProps     =   79
      Caption         =   "GroupBox1"
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin VB.TextBox TxtBairro 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1590
      End
      Begin XtremeSuiteControls.PushButton CmdPerquisar 
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   1980
         Width           =   1590
         _Version        =   720898
         _ExtentX        =   2805
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox TxtTel 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   945
         Width           =   1590
      End
      Begin VB.TextBox TxtNome 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Width           =   1590
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   1305
         Width           =   915
         _Version        =   720898
         _ExtentX        =   1614
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Bairro:"
         ForeColor       =   8421504
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   720
         Width           =   915
         _Version        =   720898
         _ExtentX        =   1614
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Celular:"
         ForeColor       =   8421504
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   90
         Width           =   915
         _Version        =   720898
         _ExtentX        =   1614
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nome:"
         ForeColor       =   8421504
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
   End
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
         Height          =   2055
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   2415
         _Version        =   720898
         _ExtentX        =   4260
         _ExtentY        =   3625
         _StockProps     =   64
         VisualTheme     =   6
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   5085
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":0275
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":050E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":0691
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":0829
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":09D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":0B73
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":0CFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":0E89
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":0F8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":121C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":1328
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":166A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":180A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":19A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":1A51
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":1EA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPaneContato.frx":22F5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption SccContato2 
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Top             =   360
      Width           =   4695
      _Version        =   720898
      _ExtentX        =   8281
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "Meus Contatos"
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
   Begin XtremeShortcutBar.ShortcutCaption SccContato 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _Version        =   720898
      _ExtentX        =   8255
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Contatos"
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
Attribute VB_Name = "FrmPaneContato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Load()
Event Resize()
Event WndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Event ChkFiltroClick()
Event CmdPerquisaClick()
Private Sub ChkAtivo_Click()
   RaiseEvent ChkFiltroClick
End Sub
Private Sub ChkEmEspera_Click()
   RaiseEvent ChkFiltroClick
End Sub
Private Sub ChkInativo_Click()
   RaiseEvent ChkFiltroClick
End Sub
Private Sub CmdPerquisar_Click()
'Me.wndTaskPanel.VisualTheme = 0
'For i = 0 To 20
'   Me.wndTaskPanel.VisualTheme = Me.wndTaskPanel.VisualTheme + 1
'Next
   RaiseEvent CmdPerquisaClick
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   Me.BackColor = &H80000005
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
Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   RaiseEvent WndTaskPanelItemClick(Item)
End Sub
