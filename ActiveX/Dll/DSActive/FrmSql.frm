VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmSql 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Executar Consulta"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   2055
   ClientWidth     =   9660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5610
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSRDC.MSRDC DataSql 
      Height          =   330
      Left            =   1440
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Options         =   0
      CursorDriver    =   0
      BOFAction       =   0
      EOFAction       =   0
      RecordsetType   =   1
      LockType        =   3
      QueryType       =   0
      Prompt          =   3
      Appearance      =   1
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   ""
      RecordSource    =   ""
      UserName        =   ""
      Password        =   ""
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "MSRDC1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox CmbOwner 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   0
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid GrdSql 
      Bindings        =   "FrmSql.frx":0000
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   12648447
      AllowUserResizing=   1
   End
   Begin VB.TextBox TxtSql 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   9375
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      WhatsThisHelpID =   10287
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Executar [F5]"
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
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   1
      Left            =   8520
      TabIndex        =   3
      Top             =   5160
      WhatsThisHelpID =   10290
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
      MousePointer    =   5
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Owner : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   7080
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Top             =   4800
      Width           =   3135
   End
End
Attribute VB_Name = "FrmSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event FormLoad()
Event FormActivate()
Event FormQueryUnload(Cancel As Integer, UnloadMode As Integer)
Event FormResize()
Event CmdOperClick(index As Integer)
Event CmbOwnerClick()
Event GrdSqlClick()
Event GrdSqlDblClick()
Event GrdSqlKeyPress(KeyAscii As Integer)
Event TxtSqlKeyPress(KeyAscii As Integer)
Private Sub CmbOwner_Click()
   RaiseEvent CmbOwnerClick
End Sub
Private Sub CmdOper_Click(index As Integer)
   RaiseEvent CmdOperClick(index)
End Sub
Private Sub Form_Activate()
   RaiseEvent FormActivate
End Sub
Private Sub Form_Load()
   RaiseEvent FormLoad
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyEscape: UnLoad Me
      Case Else: KeyAscii = ClsDsr.SendTab(Me, KeyAscii)
   End Select
   Lbl(0).Caption = ""
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   DoEvents
   Select Case KeyCode
      Case vbKeyF5: Call CmdOper_Click(0)
   End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   RaiseEvent FormQueryUnload(Cancel, UnloadMode)
End Sub
Private Sub Form_Resize()
   RaiseEvent FormResize
End Sub
Private Sub Form_Unload(Cancel As Integer)
'   Set MDIFilho = Nothing
   Call ClsDsr.SetDefault(hWnd)
End Sub
Private Sub GrdSql_Click()
   RaiseEvent GrdSqlClick
End Sub
Private Sub GrdSql_DblClick()
   RaiseEvent GrdSqlDblClick
End Sub
Private Sub GrdSql_KeyPress(KeyAscii As Integer)
   RaiseEvent GrdSqlKeyPress(KeyAscii)
End Sub
Private Sub GrdSql_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Shift = 1 Then
      Me.TxtSql = Me.TxtSql & " " & Me.GrdSql.TextMatrix(Me.GrdSql.MouseRow, Me.GrdSql.MouseCol)
  End If
End Sub
Private Sub TxtSql_KeyDown(KeyCode As Integer, Shift As Integer)
   Lbl(0).Caption = ""
End Sub
Private Sub TxtSql_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtSqlKeyPress(KeyAscii)
End Sub

Private Sub TxtSql_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
'         TxtSql = TxtSql & ">"
'         SendKeys "^{END}"
      Case vbKeyBack
         If Right(Trim(TxtSql), 2) = Chr(13) & Chr(10) Then
'            SendKeys "{BS}"
         End If
   End Select
End Sub
