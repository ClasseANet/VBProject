VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Begin VB.Form FrmView 
   AutoRedraw      =   -1  'True
   Caption         =   "View"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   2115
   ClientWidth     =   9930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4650
   ScaleWidth      =   9930
   Begin Crystal.CrystalReport Crpt 
      Bindings        =   "FrmView.frx":0000
      Left            =   240
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      ReportSource    =   3
      WindowState     =   2
   End
   Begin VB.TextBox TxtEVT 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Text            =   "TxtLov"
      Top             =   3120
      Visible         =   0   'False
      WhatsThisHelpID =   10363
      Width           =   645
   End
   Begin MSRDC.MSRDC rDataEVT 
      Height          =   330
      Left            =   240
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   327681
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
      Caption         =   "rEVT"
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
   Begin MSFlexGridLib.MSFlexGrid GrdView 
      Bindings        =   "FrmView.frx":0013
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7011
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   12648447
      AllowUserResizing=   1
      MouseIcon       =   "FrmView.frx":0026
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   0
      Left            =   8880
      TabIndex        =   2
      Top             =   4200
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   714
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Picture         =   "FrmView.frx":0340
   End
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FrmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Acesso$
Public UserDB$, Caption_Ori$
Public Tabela$, OrderBy$
Private Sub CmdOper_Click(Index As Integer)
   Dim Sql$, DscExclusao$, i%, j%
   Select Case Index
      Case 0: Unload Me
   End Select
End Sub


Private Sub Form_Activate()
   Dim Tam%, TamTit%, TotTam&
   Set Sys.MDIFilho = Me
   Me.Caption = Caption_Ori & " [" & UCase(DB.DSN) & _
                              "." & UCase(DB.StrDATABASE) & _
                              "." & UCase(Me.Tabela) & "]"
   If Tabela = "" Then
      Unload Me
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyDelete, vbKeyBack
   End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyEscape: Unload Me
      Case Else: KeyAscii = SendTab(Me, KeyAscii)
   End Select
End Sub
Private Sub Form_Load()
   Dim i%, Pos%
   Screen.MousePointer = vbHourglass
   Caption_Ori = Me.Caption
   With Me.rDataEVT
      .DataSourceName = DB.DSN
      .UserName = DB.UID
      .Password = DB.PWD
      .CursorDriver = rdUseOdbc
'      .Connect = Db.StrConect
   End With
   If Tabela <> "" Then Call MontaGridLocal("select * from " & Tabela$)
   Me.Crpt.ReportSource = Me.rDataEVT ' rcrptDataControl
   Call ConfigForm(Me, SysMdi.Icon, Sys.FundoTela)
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
   With Me
      With .GrdView
         .Move .Left, .Top, IIf(Me.Width < 285, 0, Me.Width - 285), IIf(Me.Height < 1200, 0, Me.Height - 1200)
      End With
      With CmdOper(0)
         .Move Me.Width - (.Width + 240), Me.Height - (.Height + 480)
      End With
   End With
   Call PintarFundo(Me.ImgFundo, Sys.FundoTela)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   With BANCO
      Set .TB_CRG_TERMO = Nothing
      Set .TB_DAI = Nothing
      Set .TB_PCG_IMP = Nothing
   End With
   Set Sys.MDIFilho = Nothing
End Sub

Private Sub GrdView_DblClick()
   With Me.GrdView
      MsgBox .TextMatrix(.Row, .Col), vbOKOnly, "Tab. : " & Tabela & " Campo : " & .TextMatrix(0, .Col)
   End With
End Sub

Private Sub GrdView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i%, Tam&
   If Button = 2 Then
      If y <= Me.GrdView.RowHeight(0) Then
         For i% = 0 To Me.GrdView.Cols - 1
            If x < Tam Then
               Me.GrdView.Tag = i%
               Exit For
            End If
            Tam = Tam + Me.GrdView.ColWidth(i%)
         Next
         SysMdi.MnuMouse01(1).Caption = "&Filtrar Coluna ( " & Me.GrdView.TextMatrix(0, i - 1) & " )"
      Else
         SysMdi.MnuMouse01(1).Caption = "&Ordenação Múltipla"
      End If
      PopupMenu SysMdi.MnuMouse(1)
   End If
End Sub

Private Sub GrdView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If y <= Me.GrdView.RowHeight(0) Then
      Me.GrdView.MousePointer = flexCustom
   Else
      Me.GrdView.MousePointer = flexDefault
   End If
End Sub

Private Sub GrdView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i%, Tam&
   Dim Pos%, Desc$
   If Button = 1 Then
      If y <= Me.GrdView.RowHeight(0) Then
         For i% = 0 To Me.GrdView.Cols - 1
            If x < Tam Then
               Me.GrdView.Tag = i%
               Exit For
            End If
            Tam = Tam + Me.GrdView.ColWidth(i%)
         Next
         Pos = InStr(UCase(Me.rDataEVT.Sql), "ORDER BY")
         Sql = Me.rDataEVT.Sql
         If Pos > 0 Then
            Sql = Mid(Me.rDataEVT.Sql, 1, Pos - 1)
            If Me.GrdView.TextMatrix(0, i - 1) = UCase(Trim(Mid(Me.rDataEVT.Sql, Pos + 8))) Then
              Desc = " DESC "
            Else
               Desc = ""
            End If
         End If
         OrderBy = " Order by " + CStr(Me.GrdView.TextMatrix(0, i - 1)) & Desc
         Sql = Sql & OrderBy
         Call MontaGridLocal(Sql)
      End If
   End If
End Sub
Public Sub MontaGridLocal(Sql$)
   Dim Tam&, TamW&, i%
   Dim TamTit%, TotTam&
   On Error GoTo Fim
   TamW& = Me.GrdView.Width
   Me.GrdView.Visible = False
   Me.rDataEVT.Sql = Sql$
   Me.rDataEVT.Refresh
   Me.GrdView.Cols = Me.rDataEVT.Resultset.rdoColumns.Count
   For i = 0 To Me.rDataEVT.Resultset.rdoColumns.Count - 1
      Tam = Me.rDataEVT.Resultset.rdoColumns(i).SIZE
      TamTit = Len(Me.rDataEVT.Resultset.rdoColumns(i).Name)
      Tam = IIf(Tam >= TamTit, Tam, TamTit)
      Tam = IIf(Tam >= 50, 50, Tam)
      Me.GrdView.ColWidth(i) = 120 * Tam
      TotTam = TotTam + Me.GrdView.ColWidth(i)
   Next
   If Me.WindowState = vbMaximized Then
      Me.GrdView.Width = TamW&
   Else
      If TotTam > Me.Width Then
         Me.GrdView.Width = Me.Width - 240
      Else
         Me.GrdView.Width = TotTam + 383
         If Me.Width < Screen.Width Then
            Me.Width = Me.GrdView.Width + 240
         End If
      End If
   End If
   Me.GrdView.Left = 60
   Me.Refresh
   Me.GrdView.Visible = True
Exit Sub
Fim:
   Call ShowError
End Sub
Public Sub F_REFRESH()
   Call MontaGridLocal(Me.rDataEVT.Sql)
End Sub
