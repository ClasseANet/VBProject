VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Begin VB.Form FrmSeekCrg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localizar Carga"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6555
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ComboBox CmbAER_COD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4080
      Width           =   2355
   End
   Begin VB.CheckBox ChkSelected 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xibir Itens Selecionados"
      Height          =   195
      Index           =   0
      Left            =   120
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   22
      Top             =   6320
      Width           =   2063
   End
   Begin MSRDC.MSRDC rDataEVT 
      Height          =   330
      Left            =   6600
      Top             =   3360
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
   Begin VB.TextBox TxtEVT 
      Height          =   285
      Left            =   8640
      TabIndex        =   19
      Text            =   "TxtLov"
      Top             =   3360
      Visible         =   0   'False
      WhatsThisHelpID =   10363
      Width           =   645
   End
   Begin VB.ComboBox CmbDAI_NUM 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7800
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   5880
      Width           =   2355
   End
   Begin VB.ComboBox CmbPcg_Ninf 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7800
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   4440
      Width           =   2355
   End
   Begin VB.ComboBox CmbHCrg_Num 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7800
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   5160
      Width           =   2355
   End
   Begin VB.ComboBox CmbNum_Termo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7800
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   5520
      Width           =   2355
   End
   Begin VB.TextBox TxtCrg_Num 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   7
      Top             =   4800
      WhatsThisHelpID =   10215
      Width           =   2355
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCampo 
      Height          =   5895
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   10398
      _Version        =   393216
      BackColor       =   12648447
      AllowUserResizing=   1
   End
   Begin VB.ListBox LstTabela 
      Height          =   5910
      ItemData        =   "SeekCrg.frx":0000
      Left            =   120
      List            =   "SeekCrg.frx":0007
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   450
      Index           =   0
      Left            =   10200
      TabIndex        =   15
      ToolTipText     =   "Criar Classe"
      Top             =   5280
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "&Limpar Tela"
      ForeColor       =   -2147483635
      Font3D          =   3
      Picture         =   "SeekCrg.frx":0014
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   1
      Left            =   10200
      TabIndex        =   14
      Top             =   5880
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
      Picture         =   "SeekCrg.frx":0030
   End
   Begin MSFlexGridLib.MSFlexGrid GridHist 
      Bindings        =   "SeekCrg.frx":004C
      Height          =   3495
      Left            =   6480
      TabIndex        =   18
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6165
      _Version        =   393216
      BackColor       =   12648447
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   450
      Index           =   2
      Left            =   10200
      TabIndex        =   21
      ToolTipText     =   "Criar Classe"
      Top             =   3960
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "&Pesquisar"
      ForeColor       =   -2147483635
      Font3D          =   3
      Picture         =   "SeekCrg.frx":005F
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AER_COD"
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
      Index           =   8
      Left            =   6600
      TabIndex        =   2
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Eventos"
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
      Left            =   6480
      TabIndex        =   20
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DAI_NUM"
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
      Index           =   6
      Left            =   6600
      TabIndex        =   12
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUM_TERMO"
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
      Left            =   6600
      TabIndex        =   10
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HCRG_NUM"
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
      Index           =   4
      Left            =   6600
      TabIndex        =   8
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CRG_NUM"
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
      Index           =   3
      Left            =   6600
      TabIndex        =   6
      Top             =   4920
      Width           =   930
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PCG_NINF"
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
      Index           =   2
      Left            =   6600
      TabIndex        =   4
      Top             =   4560
      Width           =   930
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tabelas "
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
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo / &Descrição"
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
      Index           =   1
      Left            =   2640
      TabIndex        =   17
      Top             =   0
      Width           =   1530
   End
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FrmSeekCrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Suja%
Public UserDB$, Caption_Ori$
Public VetTabela As Collection
Public Sub F_REFRESH()
   Dim VlChk%, Sql$
   Call SetHourglass(Me.hWnd)
   
   Me.Caption = Caption_Ori & " [" & UCase(DB.DSN) & "." & UCase(DB.StrDATABASE) & "]"
   Me.Refresh
   With Me.rDataEVT
      .DataSourceName = DB.DSN
      .UserName = DB.UID
      .Password = DB.PWD
      .CursorDriver = rdUseOdbc
   End With
   
'* Formatar Grid de Eventos
   Call Popula_GridHist

'* Montar Combo de AER_COD
   Sql = "select distinct PRM_VLR "
   Sql = Sql & "from PRM_TECA"
   Sql = Sql & " where PRM_COD='AER_COD'"
   Call MontarDBCombo(DB, Me.CmbAER_COD, Sql$, "PRM_VLR")
   
'* Montar Combo de Tabelas
   Me.LstTabela.Visible = False
   Me.LstTabela.Refresh
   Call MontarLstTabela
   Call ChkSelected_Click(0)

'* Formata Grid
   Me.GrdCampo.TextMatrix(0, 0) = "Campo"
   Me.GrdCampo.TextMatrix(0, 1) = "Valor"
   Me.GrdCampo.ColWidth(0) = 15 * 120
   Me.GrdCampo.ColWidth(1) = 20 * 120
      
   Call LstTabela_Click
   
   Call PopulaCampos

   Me.LstTabela.Visible = True
   Me.LstTabela.Refresh
   Call SetDefault(Me.hWnd)
End Sub
Public Sub LimpaTela()
   Dim ValChk%, i%
   On Error Resume Next
   ValChk% = Me.ChkSelected(0).Value
   Call LimparTela(Me)
   Me.CmbDAI_NUM.Clear
   Me.CmbHCrg_Num.Clear
   Me.CmbNum_Termo.Clear
   Me.CmbPcg_Ninf.Clear
   With BANCO
      Set .TB_CRG_TERMO = Nothing
      Set .TB_DAI = Nothing
      Set .TB_PCG_IMP = Nothing
   End With
   For i = 1 To Me.GrdCampo.Rows - 1
      Me.GrdCampo.TextMatrix(i, 1) = ""
   Next
   Me.GridHist.Rows = 1
   Me.ChkSelected(0).Value = ValChk%
   'Me.CmbPcg_Ninf.SetFocus
End Sub
Public Sub PopulaCampos()
   With BANCO
      .TB_CRG_TERMO.CRG_NUM = Me.TxtCrg_Num
      .TB_CRG_TERMO.HCRG_NUM = Me.CmbHCrg_Num
      .TB_CRG_TERMO.NUM_TERMO = Me.CmbNum_Termo
      .TB_CRG_TERMO.AER_COD = Me.CmbAER_COD
   End With
   If Me.TxtCrg_Num <> "" Then Call Popula_GridHist
   Call LstTabela_Click
'   Call PesquisaGridCampos(Me.LstTabela.ItemData(Me.LstTabela.ListIndex), UserDB & "." & Me.LstTabela)
End Sub
Public Sub PopularChaves()
   Dim lPCG$, lDAI$, n As Variant
   Dim i%
   Dim MyPCG As New TB_PCG_IMP, MyDAI As New TB_DAI
   With BANCO
      With .TB_PCG_IMP
         If .PCG_NINF <> Me.CmbPcg_Ninf Then
            Call .GetSelect(Me.CmbPcg_Ninf, Me.CmbAER_COD)
            If .EXISTE = ALTERACAO Then
               Me.TxtCrg_Num = .CRG_NUM
               Me.CmbHCrg_Num = .HCRG_NUM
               Me.CmbNum_Termo = .NUM_TERMO
               Me.CmbDAI_NUM.Clear
            Else
               Call ExibirAviso("PCG não Existe", LoadMsg(1))
               lPCG = Me.CmbPcg_Ninf
               Call LimpaTela
               Me.CmbPcg_Ninf = lPCG
               Exit Sub
            End If
         End If
      End With
      With .TB_CRG_TERMO
         If .CRG_NUM <> Me.TxtCrg_Num Or .HCRG_NUM <> Me.CmbHCrg_Num Or _
         .NUM_TERMO <> Me.CmbNum_Termo Or CmbPcg_Ninf.ListCount = 0 Or _
         CmbDAI_NUM.ListCount = 0 Then
            Call .GetSelect(Me.TxtCrg_Num, Me.CmbHCrg_Num, Me.CmbNum_Termo, Me.CmbAER_COD)
            If .EXISTE = ALTERACAO Then
               lPCG = Me.CmbPcg_Ninf
               Me.CmbPcg_Ninf.Clear
               Set .PCGs = Nothing
               For Each n In .PCGs
                  Set MyPCG = n
                  Me.CmbPcg_Ninf.AddItem MyPCG.PCG_NINF
                  Set MyPCG = Nothing
               Next
               If lPCG <> "" Then
                  Call LocalizarCombo(Me.CmbPcg_Ninf, lPCG)
               Else
                  If Me.CmbPcg_Ninf.ListCount > 0 Then Me.CmbPcg_Ninf.ListIndex = 0
               End If
                              
               Me.CmbDAI_NUM.Clear
               Set .DAIs = Nothing
               For Each n In .DAIs
                  Set MyDAI = n
                  Me.CmbDAI_NUM.AddItem MyDAI.DAI_NUM & "." & MyDAI.DAI_SEQ & "." & MyDAI.DAI_DV
                  Set MyPCG = Nothing
               Next
               If .DAIs.Count > 0 Then Me.CmbDAI_NUM = .DAIs(1).DAI_NUM & "." & .DAIs(1).DAI_SEQ & "." & .DAIs(1).DAI_DV
            Else
'               Call ExibirAviso("Carga não Existe", loadmsg(1))
            End If
         End If
      End With
   End With
   Me.GridHist.Rows = 1
   For i = 1 To Me.GrdCampo.Rows - 1
      Me.GrdCampo.TextMatrix(i, 1) = ""
   Next
End Sub

Private Sub ChkSelected_Click(Index As Integer)
   Select Case Index
      Case 0
         Dim i%
         If ChkSelected(Index).Value = vbChecked Then
            Me.LstTabela.Visible = False
            Me.LstTabela.Refresh
            Set VetTabela = New Collection
            For i = Me.LstTabela.ListCount - 1 To 0 Step -1
               If Me.LstTabela.Selected(i) Then
                  VetTabela.Add i
               Else
                  Me.LstTabela.RemoveItem i
              End If
            Next
            Me.LstTabela.Visible = True
            Me.LstTabela.Refresh
         Else
            Call MontarLstTabela
         End If
   End Select
End Sub

Private Sub CmbAER_COD_LostFocus()
  Screen.MousePointer = vbHourglass
  Call PopularChaves
  Screen.MousePointer = vbDefault
End Sub

Private Sub CmbDAI_NUM_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub
Private Sub CmbDAI_NUM_LostFocus()
  Call SetHourglass(hWnd)
  Call PopularChaves
  Call SetDefault(hWnd)
End Sub
Private Sub CmbHCrg_Num_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub

Private Sub CmbHCrg_Num_LostFocus()
   Dim lTERMO$
   lTERMO = Me.CmbNum_Termo
   If Not PreencheTermo(CStr(Me.TxtCrg_Num), CStr(Me.CmbHCrg_Num), Me.CmbAER_COD, Me.CmbNum_Termo) Then
      Me.CmbHCrg_Num.SetFocus
   Else
      Call SetHourglass(hWnd)
      Call PopularChaves
      If lTERMO <> "" Then Call LocalizarCombo(Me.CmbNum_Termo, lTERMO)
      Call SetDefault(hWnd)
   End If
End Sub
Private Sub CmbNum_Termo_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub

Private Sub CmbNum_Termo_LostFocus()
  Call SetHourglass(hWnd)
  Call PopularChaves
  Call SetDefault(hWnd)
End Sub
Private Sub CmbPcg_Ninf_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub
Private Sub CmbPcg_Ninf_LostFocus()
  Screen.MousePointer = vbHourglass
  Call PopularChaves
  Screen.MousePointer = vbDefault
End Sub
Private Sub CmdOper_Click(Index As Integer)
   Dim Sql$, DscExclusao$, i%, j%
   
   Select Case Index
      Case 0: Call LimpaTela
      Case 1: Unload Me
      Case 2: Call PopulaCampos
   End Select
End Sub
Private Sub Form_Activate()
   Screen.MousePointer = vbHourglass
   Set Sys.MDIFilho = Me
   'Me.Caption = Caption_Ori & " [" & UCase(Db.DSN) & "." & UCase(Db.StrDATABASE) & "]"
   Call F_REFRESH
   Screen.MousePointer = vbDefault
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
'If False Then
   'Me.DataEVT.DatabaseName = Db.dBase.Name
   With Me.rDataEVT
      .DataSourceName = DB.DSN
      .UserName = DB.UID
      .Password = DB.PWD
      .CursorDriver = rdUseOdbc
   End With
   
'* Formatar Grid de Eventos
   Call Popula_GridHist
'* Montar Combo de Tabelas
   Call MontarLstTabela
'* Formata Grid
   Me.GrdCampo.TextMatrix(0, 0) = "Campo"
   Me.GrdCampo.TextMatrix(0, 1) = "Valor"
   Me.GrdCampo.ColWidth(0) = 15 * 120
   Me.GrdCampo.ColWidth(1) = 20 * 120
      
   Call LstTabela_Click
'End If
   Call ConfigForm(Me, SysMdi.Icon, Sys.FundoTela)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Suja = False
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
'   If ValidaCampos Then  F_SALVAR
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set Sys.MDIFilho = Nothing
End Sub

Private Sub LstTabela_Click()
   Dim Tabela$, i%, j%
   Screen.MousePointer = vbHourglass

'* Montar Combo de Campo de Descrição
   If UserDB = "" Then
      Tabela = Me.LstTabela
   Else
      Tabela = UserDB & "." & Me.LstTabela
   End If
   j = 0
   Me.GrdCampo.Rows = 1
   'With Db.dBase.TableDefs(Tabela)
   Dim TabInd%
   TabInd = Me.LstTabela.ItemData(Me.LstTabela.ListIndex)
   With DB.dBase.TableDefs(Me.LstTabela.ItemData(Me.LstTabela.ListIndex))
      Me.GrdCampo.Rows = .Fields.Count + 1
      For i = 0 To .Fields.Count - 1
         If .Fields(i).Attributes < dbSystemField Then
            j = j + 1
            Me.GrdCampo.TextMatrix(j, 0) = .Fields(i).Name
         End If
      Next
   End With
   Call PesquisaGridCampos(TabInd, Tabela$)
   Screen.MousePointer = vbDefault
End Sub
Private Sub LstTabela_ItemCheck(Item As Integer)
   Dim Tabela$
   If UserDB = "" Then
      Tabela = Me.LstTabela
   Else
      Tabela = UserDB & "." & Me.LstTabela
   End If
End Sub
Public Sub PesquisaGridCampos(Tabela%, TabName$)
   Dim Sql$, j%, Pos%, Pos2%
   Dim bCRG_NUM%, bHCRG_NUM%, bNUM_TERMO%
   Dim bDAI_NUM%, bDAI_SEQ%, bDAI_DV%
   Dim bCRG_TERMO%, bPCG%, bDAI%
   Dim bPCG_NINF%
   Dim i%
   On Error GoTo Fim
   Me.GrdCampo.BackColorFixed = -2147483633
   With DB.dBase.TableDefs(Tabela)
      bCRG_TERMO = False
      bDAI = False
      bPCG = False
      For i = 0 To .Fields.Count - 1
         If .Fields(i).Attributes < dbSystemField Then
            If Not bDAI Then
               For j = 0 To DB.dBase.TableDefs("TECA.DAI").Fields.Count - 1
                  If .Fields(i).Name = "DAI_NUM" Then bDAI_NUM = True
                  If .Fields(i).Name = "DAI_SEQ" Then bDAI_SEQ = True
                  If .Fields(i).Name = "DAI_DV" Then bDAI_DV = True
                  If bDAI_NUM And bDAI_SEQ And bDAI_DV Then
                     bDAI = True
                     Exit For
                  End If
               Next
            End If
         End If
         If bDAI Then Exit For
      Next
      For i = 0 To .Fields.Count - 1
         If bDAI Then Exit For
         If .Fields(i).Attributes < dbSystemField Then
            If Not bPCG Then
               For j = 0 To DB.dBase.TableDefs("TECA.PCG_IMP").Fields.Count - 1
                  If .Fields(i).Name = "PCG_NINF" Then bPCG_NINF = True
                  If bPCG_NINF Then
                     bPCG = True
                     Exit For
                  End If
               Next
            End If
         End If
         If bPCG Then Exit For
      Next
      For i = 0 To .Fields.Count - 1
         If bDAI Or bPCG Then Exit For
         If .Fields(i).Attributes < dbSystemField Then
            If Not bDAI Then
               For j = 0 To DB.dBase.TableDefs("TECA.CRG_TERMO").Fields.Count - 1
                  If .Fields(i).Name = "CRG_NUM" Then bCRG_NUM = True
                  If .Fields(i).Name = "HCRG_NUM" Then bHCRG_NUM = True
                  If .Fields(i).Name = "NUM_TERMO" Then bNUM_TERMO = True
                  If bCRG_NUM And bHCRG_NUM And bNUM_TERMO Then
                     bCRG_TERMO = True
                     Exit For
                  End If
               Next
            End If
         End If
         If bCRG_TERMO Then Exit For
      Next
      If BANCO.TB_CRG_TERMO.CRG_NUM <> "" Then
         If bCRG_TERMO Then
            Sql = "select * from " & Mid(TabName, InStr(TabName, ".") + 1)
            Sql = Sql & " where CRG_NUM=" & Aspas(Me.TxtCrg_Num)    ' BANCO.TB_CRG_TERMO.CRG_NUM)
            Sql = Sql & " and  HCRG_NUM=" & Aspas(Me.CmbHCrg_Num)   '  BANCO.TB_CRG_TERMO.HCRG_NUM)
            Sql = Sql & " and  NUM_TERMO=" & Aspas(Me.CmbNum_Termo) '  BANCO.TB_CRG_TERMO.NUM_TERMO)
            Sql = Sql & " and  AER_COD = " & Aspas(Me.CmbAER_COD)   ' BANCO.TB_CRG_TERMO.AER_COD)
            DB.AbreTabela (Sql)
            If DB.CodeSql Then
               Call PopulaGridCampos(DB.Dys)
            End If
            Exit Sub
         End If
         If bPCG Then
            Sql = "select * from " & Mid(TabName, InStr(TabName, ".") + 1)
            Sql = Sql & " where PCG_NINF=" & Aspas(Me.CmbPcg_Ninf)  ' BANCO.TB_PCG_IMP.PCG_NINF)
            Sql = Sql & " and AER_COD = " & Aspas(Me.CmbAER_COD)    '  BANCO.TB_PCG_IMP.AER_COD)
            DB.AbreTabela (Sql)
            If DB.CodeSql Then
               Call PopulaGridCampos(DB.Dys)
            End If
            Exit Sub
         End If
         If bDAI Then
            Pos = InStr(Me.CmbDAI_NUM, ".")
            Pos2 = InStr(Pos + 1, Me.CmbDAI_NUM, ".")
            If Pos <> 0 And Pos2 <> 0 Then
               Sql = "select * from " & Mid(TabName, InStr(TabName, ".") + 1)
               Sql = Sql & " where DAI_NUM=" & Aspas(Mid(Me.CmbDAI_NUM, 1, Pos - 1))
               Sql = Sql & " and   DAI_SEQ=" & Aspas(Mid(Me.CmbDAI_NUM, Pos + 1, Pos2 - Pos - 1))
               Sql = Sql & " and   DAI_DV=" & Aspas(Mid(Me.CmbDAI_NUM, Pos2 + 1))
               Sql = Sql & " and   AER_COD = " & Aspas(Me.CmbAER_COD)
               DB.AbreTabela (Sql)
               If DB.CodeSql Then
                  Call PopulaGridCampos(DB.Dys)
               End If
               Exit Sub
            End If
         End If
      End If
   End With
'    Else
'    Exit Sub
'    End If
   If Not (bCRG_TERMO Or bPCG Or bDAI) Then
      If BANCO.TB_CRG_TERMO.CRG_NUM = "" Then
         Me.GrdCampo.BackColorFixed = vbWhite
      Else
         Call JoinTabelas(TabName)
      End If
   End If
   Exit Sub
Fim:
   Call ShowError("FrmSeekCrg.PesquisaGridCampos")
End Sub

Private Sub LstTabela_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call MdiPrincipal.MnuMouse00_Click(0)
   End If
End Sub

Private Sub LstTabela_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then PopupMenu MdiPrincipal.MnuMouse(0)
End Sub

Private Sub TxtCrg_Num_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub

Private Sub TxtCrg_Num_LostFocus()
   Dim lHCRG$
   lHCRG = Me.CmbHCrg_Num
   If Not PreencheHouse(CStr(Me.TxtCrg_Num), Me.CmbAER_COD, Me.CmbHCrg_Num) Then
      Me.TxtCrg_Num.SetFocus
   Else
      If lHCRG <> "" Then Call LocalizarCombo(Me.CmbHCrg_Num, lHCRG)
   End If
End Sub
Function PreencheHouse(AWB$, AER_COD$, cmbHouse As ComboBox) As Integer
   Dim Sql$
   
   On Error GoTo Fim
   Screen.MousePointer = vbHourglass
   cmbHouse.Clear
   
   Sql = "SELECT DISTINCT HCRG_NUM,CRTE_TIP "
   Sql = Sql & " FROM CRG_TERMO "
   Sql = Sql & " WHERE CRG_NUM = " & Aspas(AWB)
   Sql = Sql & " AND AER_COD = " & Aspas(AER_COD)
   Sql = Sql & " ORDER BY HCRG_NUM"
   Call MontarDBCombo(DB, Me.CmbHCrg_Num, Sql, "HCRG_NUM")
   If cmbHouse.ListCount >= 1 Then
      cmbHouse.ListIndex = 0
      cmbHouse.Enabled = True
   End If
   PreencheHouse = True
   Screen.MousePointer = vbDefault
   Exit Function
Fim:
  Call ShowError("PreencheHouse")
  PreencheHouse = False
End Function
Function PreencheTermo(AWB$, HOUSE$, AERCOD$, cmbTermo As ComboBox) As Integer
    Dim Sql$
    On Error GoTo Fim
    Screen.MousePointer = vbHourglass
    'Carrega combo de Termo
    cmbTermo.Clear
    Sql = "select distinct NUM_TERMO "
    Sql = Sql & " from CRG_TERMO "
    Sql = Sql & " where CRG_NUM = " & Aspas(AWB)
    Sql = Sql & " and HCRG_NUM = " & IIf(HOUSE = "", "' '", Aspas(HOUSE))
    Sql = Sql & " and AER_COD = " & Aspas(AERCOD)
    Call MontarDBCombo(DB, Me.CmbNum_Termo, Sql, "NUM_TERMO")
    If cmbTermo.ListCount >= 1 Then
       cmbTermo.ListIndex = 0
    End If
    Screen.MousePointer = vbDefault
    PreencheTermo = True
    Exit Function
Fim:
   Call ShowError("PreencheTermo")
   PreencheTermo = False
End Function
Public Sub Popula_GridHist()
   Dim Sql$, Cab
   On Error GoTo Fim
   Screen.MousePointer = vbHourglass
  Cab = Array("PCG", "PCG_NINF", 12, flexAlignLeftCenter, _
              "COD.", "EVT_CRG_COD", 4, flexAlignLeftCenter, _
              "EVENTO", "SIT_CRG_DES", 22, flexAlignLeftCenter, _
              "MOTIVO", "EVT_CRG_MOT", 15, flexAlignLeftCenter, _
              "CAMPO", "EVT_CRG_CAMPO", 0, flexAlignLeftCenter, _
              "ANTERIOR", "EVT_CRG_ANT", 0, flexAlignLeftCenter, _
              "ATUAL", "EVT_CRG_ATU", 0, flexAlignLeftCenter, _
              "PESS", "PESS_COD", 9, flexAlignLeftCenter, _
              "AFTN", "AFTN_ID", 0, flexAlignLeftCenter, _
              "USUARIO", "EVT_USR_ID", 7, flexAlignLeftCenter, _
              "DATA", "EVT_CRG_DAT", 8, flexAlignLeftCenter, _
              "HORA", "EVT_CRG_HOR", 8, flexAlignLeftCenter)
                 
   Sql = "select E.NUM_TERMO, E.HCRG_NUM, E.PCG_NINF, E.EVT_CRG_COD, S.SIT_CRG_DES "
   Sql = Sql & ", E.EVT_CRG_MOT, E.EVT_CRG_CAMPO, E.EVT_CRG_ANT, E.EVT_CRG_ATU, E.PESS_COD"
   Sql = Sql & ", E.AFTN_ID, E.EVT_USR_ID, E.EVT_CRG_DAT, E.EVT_CRG_HOR"
   Sql = Sql & " from EVT_CRG_IMP E, SITUACAO_CRG S"
   Sql = Sql & " Where E.CRG_NUM = " & Aspas(Me.TxtCrg_Num)
   Sql = Sql & " and   E.AER_COD = " & Aspas(Me.CmbAER_COD)
   If Trim(CmbHCrg_Num.Text) <> "" Then
       Sql = Sql & " AND E.HCRG_NUM = " & Aspas(Me.CmbHCrg_Num)
   End If
   If Trim(CmbNum_Termo.Text) <> "" Then
       Sql = Sql & " AND E.NUM_TERMO = " & Aspas(Me.CmbNum_Termo.Text)
   End If
   Sql = Sql & " AND S.SIT_CRG (+)= E.EVT_CRG_COD "
   Sql = Sql & " ORDER BY TRUNC(E.EVT_CRG_DAT) ASC, E.EVT_CRG_HOR ASC"
   If Me.TxtCrg_Num = "" Then
      Call MontarCabGrid(Me.GridHist, Cab, 4695)
      Me.GridHist.Rows = 1
   Else
      Call MontarMSGrid(Me.rDataEVT, Me.GridHist, Cab, Sql, 4695)
   End If
   Screen.MousePointer = vbDefault
   Me.GridHist.FixedCols = 3
   Exit Sub
Fim:
   ShowError ("Popula_GridHist")
End Sub
Public Sub MontarLstTabela(Optional TabSys = False)
   Dim i%, j%, Pos%, n
   Dim Bool As Boolean
   Me.LstTabela.Visible = False
   Me.LstTabela.Refresh
   TabSys = SysMdi.MnuMouse00(1).Checked
   With DB.dBase
      Me.LstTabela.Clear
      j = 0
      For i = 0 To .TableDefs.Count - 1
          Bool = ((.TableDefs(i).Attributes And dbSystemObject) = 0)
          Bool = IIf(TabSys, Not Bool, Bool)
          If Bool Then
            Pos = InStr(DB.dBase.TableDefs(i).Name, ".")
            If Pos > 0 Then
               If Mid(DB.dBase.TableDefs(i).Name, 1, Pos - 1) = "TECA" And _
                  InStr(DB.dBase.TableDefs(i).Name, "EXP") = 0 And _
                  InStr(DB.dBase.TableDefs(i).Name, "TIRET") = 0 And _
                  InStr(DB.dBase.TableDefs(i).Name, "TICOT") = 0 And _
                  InStr(DB.dBase.TableDefs(i).Name, "TILIT") = 0 And _
                  InStr(DB.dBase.TableDefs(i).Name, "TERET") = 0 And _
                  InStr(DB.dBase.TableDefs(i).Name, "TECOT") = 0 And _
                  InStr(DB.dBase.TableDefs(i).Name, "TEART") = 0 And _
                  InStr(DB.dBase.TableDefs(i).Name, "TEE") = 0 Then
                  UserDB = Mid(DB.dBase.TableDefs(i).Name, 1, Pos - 1)
                  Me.LstTabela.AddItem Mid(DB.dBase.TableDefs(i).Name, Pos + 1)
                  Me.LstTabela.ItemData(j) = i
                  j = j + 1
               End If
            Else
               Me.LstTabela.AddItem DB.dBase.TableDefs(i).Name
               Me.LstTabela.ItemData(j) = i
               j = j + 1
            End If
         End If
      Next
   End With
   If Not VetTabela Is Nothing Then
      For Each n In VetTabela
         Me.LstTabela.Selected(n) = True
      Next
   End If
   Me.LstTabela.Visible = True
   Me.LstTabela.Refresh
End Sub
Public Sub JoinTabelas(Tabela$)
   Dim Pos%
   Dim lTAB$, Sql$
   Pos = InStr(Tabela, "TECA.")
   If Pos > 0 Then lTAB = Mid(Tabela, Pos + 5)
   Select Case lTAB
      Case "MRC_PERD"
         Aspas (Trim(BANCO.TB_CRG_TERMO.HCRG_NUM))
         Sql = "select M.* from MRC_PERD M, CRG_PERD C"
         Sql = Sql & " where C.CRG_NUM=" & Aspas(BANCO.TB_CRG_TERMO.CRG_NUM)
         Sql = Sql & " and   C.HCRG_NUM=" & Aspas(BANCO.TB_CRG_TERMO.HCRG_NUM)
         Sql = Sql & " and   C.NUM_TERMO=" & Aspas(BANCO.TB_CRG_TERMO.NUM_TERMO)
         Sql = Sql & " and   C.AER_COD = " & Aspas(BANCO.TB_CRG_TERMO.AER_COD)
         Sql = Sql & " and   C.PERD_SEQ= M.PERD_SEQ"
         Sql = Sql & " and   C.AER_COD = M.AER_COD)"
         Call DB.AbreTabela(Sql)
         Call PopulaGridCampos(DB.Dys)
      Case Else: Me.GrdCampo.BackColorFixed = vbWhite 'Call ExibirAviso("Tabela Sem Relacionamento.", LoadMsg(1))
      
   End Select
End Sub
Public Sub PopulaGridCampos(Dys As Recordset)
   Dim k%

   For k = 1 To Me.GrdCampo.Rows - 1
      Me.GrdCampo.TextMatrix(k, 1) = Trim(Dys(Me.GrdCampo.TextMatrix(k, 0)) & "")
      Me.GrdCampo.ColAlignment(1) = 1 'flexAlignLeft
   Next
End Sub
