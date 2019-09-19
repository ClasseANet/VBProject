VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ComboBox CmbUnid 
      Height          =   315
      ItemData        =   "FrmSetup.frx":0000
      Left            =   6000
      List            =   "FrmSetup.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox TxtDscItem 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      MaxLength       =   50
      TabIndex        =   11
      Top             =   360
      Width           =   3120
   End
   Begin VB.ListBox LstTabela 
      Height          =   5235
      ItemData        =   "FrmSetup.frx":002F
      Left            =   120
      List            =   "FrmSetup.frx":0036
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      WhatsThisHelpID =   10287
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
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   5640
      WhatsThisHelpID =   10288
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Salvar"
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
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   5640
      WhatsThisHelpID =   10289
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
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   6
      Top             =   5640
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
   Begin MSMask.MaskEdBox MskId 
      Height          =   330
      Left            =   6480
      TabIndex        =   0
      Top             =   5520
      WhatsThisHelpID =   10247
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "9999"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCampos 
      Bindings        =   "FrmSetup.frx":0043
      Height          =   4635
      Left            =   2880
      TabIndex        =   10
      Top             =   720
      WhatsThisHelpID =   10506
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   8176
      _Version        =   393216
      BackColor       =   12648447
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "FrmSetup.frx":0055
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unidade"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   6000
      TabIndex        =   14
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Descrição"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   7
      Left            =   2880
      TabIndex        =   13
      Top             =   120
      Width           =   825
   End
   Begin VB.Label LblId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LblId Padrão"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   5400
      MouseIcon       =   "FrmSetup.frx":036F
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5520
      WhatsThisHelpID =   10246
      Width           =   1035
   End
   Begin VB.Label LblFrme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      WhatsThisHelpID =   10299
      Width           =   5175
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LBL Padrão"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   5400
      TabIndex        =   1
      Top             =   5880
      Width           =   990
   End
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FrmSetup"
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
Public Sub MontaLstTabela()
'* Montar Combo de Tabelas
   With DB.dBase
      Me.LstTabela.Clear
      
      For i = 0 To .TableDefs.Count - 1
         If (.TableDefs(i).Attributes And dbSystemObject) = 0 Then
            Pos = InStr(DB.dBase.TableDefs(i).Name, ".")
            If Pos > 0 Then
               If Mid(DB.dBase.TableDefs(i).Name, 1, Pos - 1) = "TECA" Then
                  UserDB = Mid(DB.dBase.TableDefs(i).Name, 1, Pos - 1)
                  Me.LstTabela.AddItem Mid(DB.dBase.TableDefs(i).Name, Pos + 1)
               End If
            Else
               Me.LstTabela.AddItem DB.dBase.TableDefs(i).Name
            End If
         End If
      Next
   End With
End Sub

Private Sub CmdOper_Click(Index As Integer)
   Select Case Index
      Case 0: Call F_INCLUIR
      Case 1: Call F_SALVAR
      Case 2: Call F_EXCLUIR
      Case 3: Unload Me
   End Select
End Sub
Private Sub Form_Activate()
   Call SetHourglass(hWnd)
   Set MDIFilho = Me
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
   
'* Montar Combo de Tabelas
   Call MontaLstTabela
   Me.GrdCampos.Cols = 3
   Me.GrdCampos.ColWidth(0) = 200
   Me.GrdCampos.ColWidth(1) = 1600
   Me.GrdCampos.ColWidth(2) = 1600
   Call LstTabela_Click
   
   Call ConfigForm(Me, SysMdi.Icon, FundoTela)
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
   End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '=============
   '=  Se nenhum campo foi alterado -> SAIR
   '=============
   If Not Me.Suja Then Exit Sub
   '=============
   '=   Se não deseja salvar -> SAIR
   '=============
   If ExibirPergunta(LoadMsg(54), Me.Caption) = vbNo Then
      Exit Sub
   End If
   '=============
   '=   Verificar e validar campos
   '=============
   If ValidaCampos Then F_SALVAR
End Sub
Private Sub Form_Resize()
   Call PintarFundo(Me.ImgFundo, FundoTela)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set MDIFilho = Nothing
'   Set BANCO.TB_ = Nothing
   Call SetDefault(hWnd)
End Sub
Private Sub GrdCampos_DblClick()
   If Me.GrdCampos.TextMatrix(Me.GrdCampos.Row, 0) = "X" Then
      Me.GrdCampos.TextMatrix(Me.GrdCampos.Row, 0) = ""
   Else
      Me.GrdCampos.TextMatrix(Me.GrdCampos.Row, 0) = "X"
   End If
End Sub

Private Sub GrdCampos_EnterCell()
   Call LocalizarCombo(Me.CmbUnid, Me.GrdCampos.TextMatrix(Me.GrdCampos.Row, 2))
End Sub

Private Sub GrdCampos_LeaveCell()
   Me.GrdCampos.TextMatrix(Me.GrdCampos.Row, 2) = Me.CmbUnid.Text
End Sub

Private Sub GrdCampos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i%, AllChk As Boolean
   If x <= Me.GrdCampos.ColWidth(0) And y <= Me.GrdCampos.RowHeight(0) Then
      AllChk = True
      For i = 1 To Me.GrdCampos.Rows - 1
         If Me.GrdCampos.TextMatrix(i, 0) = "" Then
            AllChk = False
            Exit For
         End If
      Next
      For i = 1 To Me.GrdCampos.Rows - 1
         Me.GrdCampos.TextMatrix(i, 0) = IIf(AllChk, "", "X")
      Next
   End If
End Sub

Private Sub GrdCampos_RowColChange()
  Me.CmbUnid.Width = 1570
  Me.GrdCampos.Width = 3700
  'Me.GrdCampos.RowHeight(Me.GrdCampos.Row) = 300
  Me.CmbUnid.Visible = True
  Me.CmbUnid.Move 4720, Me.GrdCampos.Top + Me.GrdCampos.RowPos(Me.GrdCampos.Row)
End Sub

Private Sub GrdCampos_Scroll()
   If Me.GrdCampos.RowPos(Me.GrdCampos.Row) <= 0 Or _
      Me.GrdCampos.RowPos(Me.GrdCampos.Row) >= (Me.GrdCampos.Width + Me.GrdCampos.Top) Then
      Me.CmbUnid.Visible = False
   Else
      Me.CmbUnid.Visible = True
   End If
   Me.CmbUnid.Move 4720, Me.GrdCampos.Top + Me.GrdCampos.RowPos(Me.GrdCampos.Row)
  
End Sub

Private Sub LblId_Click(Index As Integer)
   Call F_PROCURAR(Index)
End Sub
Private Sub LstTabela_Click()
   Dim Tabela$
'* Montar Grid de Campos
   Me.GrdCampos.Rows = 1

   If UserDB = "" Then
      Tabela = Me.LstTabela
   Else
      Tabela = UserDB & "." & Me.LstTabela
   End If
   With DB.dBase.TableDefs(Tabela)
      For i = 0 To .Fields.Count - 1
         If .Fields(i).Attributes < dbSystemField Then
            Me.GrdCampos.Rows = Me.GrdCampos.Rows + 1
            Me.GrdCampos.TextMatrix(Me.GrdCampos.Rows - 1, 1) = .Fields(i).Name
         End If
      Next
   End With
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
