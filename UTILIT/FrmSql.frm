VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Begin VB.Form FrmSql 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Executar ""Query"""
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   9660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSRDC.MSRDC DataSql 
      Height          =   330
      Left            =   6960
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "DataSql"
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
   Begin MSFlexGridLib.MSFlexGrid GrdSql 
      Bindings        =   "FrmSql.frx":0000
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3836
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
      TabIndex        =   0
      Text            =   "FrmSql.frx":0012
      Top             =   120
      Width           =   9375
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4440
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
      Left            =   3840
      TabIndex        =   2
      Top             =   4440
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
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FrmSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Suja As Boolean, PrimeiraVez As Boolean
Public Grd As MSFlexGrid
Public Function ValidaCampos()
'   If Me.MskIdCusto = "" Then
'      Call ExibirAviso(LoadMsg(27) & vbCrLf & Me.LblId(1), LoadMsg(1))
'      Me.MskId.SetFocus
'      Exit Function
'   End If
   ValidaCampos = True
End Function
Public Sub Popula_()
   Dim Id$
'   If Me.MskId = "" Then Exit Sub
'   With BANCO.TB_
'      Select Case .GetSelect(Me.MskId, Me.Msk)
'         Case ALTERACAO
         '* Popula Tela
'            Me.MskDt = DToMask(.DT, Me.MskDt)
'            Me.Txt = .CAMPO
'            Me.Suja = False
'         Case INCLUSAO
'            Id$ = StrZero(Me.MskId, Me.MskId.MaxLength)
'            Call LimparTela(Me)
'            Me.MskId = Id$
'         Case ERRO
'      End Select
'   End With
End Sub
Public Sub F_INCLUIR()
'   If Not VerificaAcesso(Me.Acesso, INCLUSAO) Then Exit Sub
'   Call F_SALVAR
'   Call LimparTela(Me)
'   Me.MskIdReq.SetFocus
End Sub
Public Function F_SALVAR() As Boolean
'   If Not VerificaAcesso(Me.Acesso, ALTERACAO) Then Exit Sub
'   If Not ValidaCampos() Then Exit Sub
'   With BANCO.TB_
'      Call .GetSelect(Me.MskId, Me.Msk)
'      .ID = Me.MskId
'      If .EXISTE = ALTERACAO Then
'         .ALTERA
'      ElseIf .EXISTE = INCLUSAO Then
'         .INCLUI
'      End If
'   End With
   F_SALVAR = True
End Function
Public Function F_EXCLUIR() As Boolean
'   Dim Arr(2)
'   If Not VerificaAcesso(Me.Acesso, EXCLUSAO) Then Exit Sub
'   Arr(0) = BANCO.TB_.QryDelete(Me.MskId, Me.Msk)
'   If DB.Executa(Arr) Then
'      Call LimparTela(Me)
'      DoEvents
'      Me.MskId.SetFocus
'   End If
   
   F_EXCLUIR = True
End Function
Public Sub F_REFRESH()
'   Call Popula_
End Sub
Public Sub F_PROCURAR(Optional Index = 0)
   Dim Arrid
   Select Case Index
'      Case 0: Arrid = F_LOV("TB_")
'      Case 1: Arrid = F_LOV("TB_")
   End Select
   '=======================
   If IsEmpty(Arrid) Then Exit Sub
   If UBound(Arrid) < 0 Then Exit Sub
   '=======================
   Select Case Index
      Case 0
'         Me.MskId = Arrid(0)
'         Me.Msk = Arrid(1)
'         Call Popula_
      Case 1
'         Me.MskId = Arrid(0)
'         Call MskId_LostFocus
'         Me.TxtDsc.SetFocus
   End Select
End Sub
Private Sub CmdOper_Click(Index As Integer)

   Select Case Index
      Case 0: Call ExecSql
      Case 1: Unload Me
   End Select
End Sub
Private Sub Form_Activate()
   Call SetHourglass(hWnd)
   Set Sys.MDIFilho = Me
'   Call Popula_
'   If PrimeiraVez Then
'      Me.MskId.SetFocus
'      PrimeiraVez = False
'   End If
   Call SetDefault(hWnd)
'   If Not VerificaAcesso(Me.Acesso, LEITURA) Then
'      Unload Me
'   End If
End Sub

Private Sub Form_Load()
   Dim i%, Pos%
   Call SetHourglass(hWnd)
   
   'Me.DataEVT.DatabaseName = Db.dBase.Name
   With Me.DataSql
      .DataSourceName = DB.DSN
      .UserName = DB.UID
      .Password = DB.PWD
      .CursorDriver = rdUseOdbc
   End With
'   Me.GrdSql.Rows = 0
'   Me.GrdSql.Cols = 0
   
   Call ConfigForm(Me, SysMdi.Icon, Sys.FundoTela)
   Me.TxtSql.SetFocus
   SendKeys "^{END}"
   Call SetDefault(hWnd)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyEscape: Suja = False: Unload Me
      Case Else: KeyAscii = SendTab(Me, KeyAscii)
   End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyDelete, vbKeyBack: Me.Suja = True
      Case vbKeyF2
         '* Executar Lista de Valores ao teclar [F2]
'         Select Case Me.ActiveControl.Name
'            Case Me.MskId.Name: Call LblId_Click(0)
'         End Select
      Case vbKeyF5: Call CmdOper_Click(0)
   End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '=============
   '=  Se nenhum campo foi alterado -> SAIR
   '=============
'   If Not Me.Suja Then Exit Sub
   '=============
   '=   Se não deseja salvar -> SAIR
   '=============
'   If ExibirPergunta(LoadMsg(54), Me.Caption) = vbNo Then
'      Exit Sub
'   End If
   '=============
   '=   Verificar e validar campos
   '=============
'   If ValidaCampos Then F_SALVAR
End Sub
Private Sub Form_Resize()
   Call PintarFundo(Me.ImgFundo, Sys.FundoTela)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set MDIFilho = Nothing
'   Set BANCO.TB_ = Nothing
   Call SetDefault(hWnd)
End Sub
Private Sub MskId_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub MskId_LostFocus()
   Call Popula_
   If Trim(MskId) = "" And Me.ActiveControl <> Me.CmdOper(3) Then
      Call LimparTela(Me)
   End If
End Sub
Private Sub Txt_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub GrdSql_DblClick()
   Me.TxtSql = Trim(Me.TxtSql)
   Me.TxtSql = Me.TxtSql & " " & Me.GrdSql.TextMatrix(Me.GrdSql.MouseRow, Me.GrdSql.MouseCol)
End Sub

Private Sub GrdSql_KeyPress(KeyAscii As Integer)
   Me.TxtSql = Me.TxtSql & Chr(KeyAscii)
End Sub

Private Sub GrdSql_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Shift = 1 Then
      Me.TxtSql = Me.TxtSql & " " & Me.GrdSql.TextMatrix(Me.GrdSql.MouseRow, Me.GrdSql.MouseCol)
  End If

End Sub
Public Sub ExecSql()
   Dim Sql$
   Dim Cab, Pos%
   Screen.MousePointer = vbHourglass
   Sql = StrReplace(Me.TxtSql, Chr(13) & Chr(10), " ")
   Sql = UCase(Trim(StrReplace(Sql, """", "'")))
   Pos = InStr(Sql, ";")
   While Pos <> 0
      If Pos = Len(Sql) Then
         Sql = Trim(Mid(Sql, 1, Pos - 1))
         Pos = 0
      Else
         Sql = Trim(Mid(Sql, Pos + 1))
         Pos = InStr(Sql, ";")
      End If
   Wend
   Pos = InStr(Trim(Sql), " ")
   If Pos = 0 Then Exit Sub
   Select Case Mid(Sql, 1, Pos - 1)
      Case "SELECT"
         Call MontarMSGrid(Me.DataSql, Me.GrdSql, Cab, Sql)
         Me.GrdSql.SetFocus
      Case "INSERT", "UPDATE", "DELETE"
         Call DB.Executa(Sql)
         Me.GrdSql.SetFocus
      Case "DESC"
         Tabela = Mid(Sql, Pos + 1)
         With DB.dBase.TableDefs("teca." & Tabela)
            Me.GrdSql.Rows = .Fields.Count + 1
            Me.GrdSql.Cols = 3
            Me.GrdSql.TextMatrix(0, 0) = "CAMPO"
            Me.GrdSql.TextMatrix(0, 1) = "TIPO"
            Me.GrdSql.TextMatrix(0, 2) = "TAMANHO"

            For i = 0 To .Fields.Count - 1
               Me.GrdSql.TextMatrix(i + 1, 0) = .Fields(i).Name
               Select Case GrpTipoCampo(.Fields(i).Type)
                  Case 1: Me.GrdSql.TextMatrix(i + 1, 1) = "NUMÉRICO"
                  Case 2: Me.GrdSql.TextMatrix(i + 1, 1) = "DATA"
                  Case 3: Me.GrdSql.TextMatrix(i + 1, 1) = "CARACTER"
               End Select
               Me.GrdSql.TextMatrix(i + 1, 2) = .Fields(i).Size
            Next
         End With
         Me.TxtSql
   End Select
   Me.GrdSql.FixedCols = 0
   Screen.MousePointer = vbDefault
End Sub

Private Sub TxtSql_KeyPress(KeyAscii As Integer)
   Dim Sql$
   If KeyAscii = vbKeyReturn Then
      Sql = StrReplace(Me.TxtSql, Chr(13) & Chr(10), " ")
      If Right(Trim(Sql), 1) = ";" Then
         Call ExecSql
      End If
   End If
End Sub

Private Sub TxtSql_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         TxtSql = TxtSql & ">"
         SendKeys "^{END}"
      Case vbkey
   End Select

End Sub
