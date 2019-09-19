VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmOrderBy 
   AutoRedraw      =   -1  'True
   Caption         =   "GrdLov"
   ClientHeight    =   6555
   ClientLeft      =   2760
   ClientTop       =   510
   ClientWidth     =   5070
   Icon            =   "OrderBy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6555
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin Crystal.CrystalReport CryRprt 
      Left            =   4560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.TextBox TxtLov 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Data DataLov 
      Caption         =   "LOV"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid GrdLov 
      Bindings        =   "OrderBy.frx":0442
      Height          =   4425
      Left            =   60
      TabIndex        =   0
      Top             =   870
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7805
      _Version        =   65541
      BackColor       =   12648447
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand CmdPrint 
      Height          =   420
      Left            =   4620
      TabIndex        =   1
      Top             =   6030
      Width           =   420
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   78
      Picture         =   "OrderBy.frx":045B
   End
   Begin Threed.SSCommand CmdLovOper 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   6000
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
      Left            =   120
      TabIndex        =   8
      Top             =   6000
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
   Begin Threed.SSCommand CmdLovOper 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Novo"
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
      Index           =   3
      Left            =   1080
      TabIndex        =   10
      Top             =   6000
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Atualizar"
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
      Index           =   4
      Left            =   2340
      TabIndex        =   11
      Top             =   6000
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Excluir"
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
   Begin Threed.SSCommand CmdLovOper 
      Height          =   375
      Index           =   5
      Left            =   3450
      TabIndex        =   12
      Top             =   6000
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   0
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
   Begin VB.Label LblLov 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Localizar :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   1050
   End
   Begin VB.Label LblLov 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "* Use [Esc] para limpar o campo ""Localizar"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   3885
   End
   Begin VB.Label LblLov 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "* Clique no título da lista para ordená-la"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   390
      Width           =   3885
   End
   Begin VB.Label LblLov 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "* Localize rapidamente digitando sua consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   4095
   End
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   0
      Width           =   990
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
Event cmdOper(index As Integer)
Event Excluir()
Event GrdDblClick()
Event FrmKeyPress(KeyAscii As Integer)
Event FrmKeyUp(KeyCode As Integer, Shift As Integer)
Event CmdPrintClick()

Public Sist$, Ver$, Cia$
'Public IdField, Id, Cab
 
'* Variaveis de Impressão
Public vetCab
Const FONTE_TIT = "Times New Roman"
Const FONTE_VAL = "Times New Roman"
Const TAM_FONTE_TIT = 14
Const TAM_FONTE_VAL = 9
Private Sub CmdLovOper_Click(index As Integer)
   RaiseEvent cmdOper(index)
End Sub

Private Sub CmdPrint_Click()
   RaiseEvent CmdPrintClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Click()
   Me.GrdLov.SetFocus
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

Private Sub GrdLov_DblClick()
   RaiseEvent GrdDblClick  '* CmdOper(0)
End Sub
Private Sub GrdLov_KeyPress(KeyAscii As Integer)
   Dim Col%
   Col% = Val(Chr(KeyAscii))
   If InArray(Col%, Array(1, 2, 3, 4, 5, 6, 7, 8, 9)) Then
      Me.TxtLov.Text = ""
      Call OrdenarMSGrid(Me.DataLov, Me.GrdLov, Col% - 1)
   Else
      If Not Between(KeyAscii, vbKeyA, vbKeyZ) Then Exit Sub
      Call MSGrdProcurar(Me.DataLov, Me.GrdLov, Me.TxtLov, KeyAscii)
      Me.TxtLov = Me.TxtLov + UCase(Chr(KeyAscii))
   End If
   Me.GrdLov.SetFocus
End Sub

Private Sub GrdLov_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then Me.TxtLov = Me.TxtLov + " "
End Sub

Private Sub GrdLov_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'   Call  HintMSGrid(Me.GrdLov, X, Y)
End Sub
Private Sub GrdLov_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i%, Tam&
   Me.TxtLov.Text = ""
   If y <= Me.GrdLov.RowHeight(0) Then
      For i% = 0 To Me.GrdLov.Cols - 1
         If x < Tam Then
            Me.GrdLov.Tag = i%
            Exit For
         End If
         Tam = Tam + Me.GrdLov.ColWidth(i%)
      Next
      Me.GrdLov.Visible = False
      Call OrdenarMSGrid(Me.DataLov, Me.GrdLov, i - 1)
      Me.GrdLov.Visible = True
      Me.GrdLov.SetFocus
   End If
End Sub
Private Sub TxtLov_Change()
   Call MSGrdProcurar(Me.DataLov, Me.GrdLov, Me.TxtLov, 0)
End Sub
Private Sub TxtLov_GotFocus()
   TxtLov.SelStart = 0
   TxtLov.SelLength = Len(TxtLov)
End Sub
Private Sub TxtLov_KeyPress(KeyAscii As Integer)
'   Call  MSGrdProcurar(Me.DataLov, Me.GrdLov, Me.TxtLov, KeyAscii)
'   Me.TxtLov.SetFocus
End Sub
