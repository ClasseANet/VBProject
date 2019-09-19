VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmOrderBy 
   AutoRedraw      =   -1  'True
   Caption         =   "Order By"
   ClientHeight    =   3780
   ClientLeft      =   2760
   ClientTop       =   510
   ClientWidth     =   4755
   Icon            =   "OrderBy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox ChkSelected 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exibir Itens Selecionados"
      Height          =   195
      Left            =   120
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   6
      Top             =   3600
      Width           =   2063
   End
   Begin Threed.SSCommand CmdOrder 
      Height          =   420
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   390
      _Version        =   65536
      _ExtentX        =   688
      _ExtentY        =   741
      _StockProps     =   78
      AutoSize        =   2
      Picture         =   "OrderBy.frx":0442
   End
   Begin VB.ListBox LstCampos 
      Height          =   3435
      ItemData        =   "OrderBy.frx":0A24
      Left            =   120
      List            =   "OrderBy.frx":0A2B
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin Threed.SSCommand CmdLovOper 
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   3120
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdLovOper 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   2520
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Cancela"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdOrder 
      Height          =   420
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   1200
      Width           =   390
      _Version        =   65536
      _ExtentX        =   688
      _ExtentY        =   741
      _StockProps     =   78
      AutoSize        =   2
      Picture         =   "OrderBy.frx":0A38
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3915
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "FrmOrderBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Load()
Event QueryUnload(Cancel, UnloadMode)
Event UnLoad(Cancel As Integer)
Event Activate()
Event CmdOper(index As Integer)
Event FrmKeyPress(KeyAscii As Integer)
Event FrmKeyUp(KeyCode As Integer, Shift As Integer)
Event ChkSelectedClick()

Public Sist$, Ver$, Cia$
Private Sub ChkSelected_Click()
   RaiseEvent ChkSelectedClick
End Sub
Private Sub CmdLovOper_Click(index As Integer)
   RaiseEvent CmdOper(index)
End Sub

Private Sub CmdOrder_Click(index As Integer)
   Dim StrAux$, VlAux As Boolean, DataAux%
   Dim Ind%, IndAux%, Bool As Boolean
   
   Bool = False
   Ind% = Me.LstCampos.ListIndex
   Select Case index
      Case 0
         If Ind% > 0 Then
            IndAux = Ind - 1
            Bool = True
         End If
      Case 1
         If Ind% < Me.LstCampos.ListCount - 1 Then
            IndAux = Ind + 1
            Bool = True
         End If
   End Select
   If Bool Then
      
      StrAux = Me.LstCampos.List(Ind)
      VlAux = Me.LstCampos.Selected(Ind)
      DataAux = Me.LstCampos.ItemData(Ind)
      
      Me.LstCampos.List(Ind) = Me.LstCampos.List(IndAux)
      Me.LstCampos.Selected(Ind) = Me.LstCampos.Selected(IndAux)
      Me.LstCampos.ItemData(Ind) = Me.LstCampos.ItemData(IndAux)
      
      Me.LstCampos.List(IndAux) = StrAux
      Me.LstCampos.Selected(IndAux) = VlAux
      Me.LstCampos.ItemData(IndAux) = DataAux
      
      Me.LstCampos.ListIndex = IndAux
   End If
End Sub

Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Click()
   Me.LstCampos.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent FrmKeyPress(KeyAscii)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent FrmKeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   RaiseEvent QueryUnload(Cancel, UnloadMode)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent UnLoad(Cancel)
End Sub

