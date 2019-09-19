VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmPadrão 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Padrão"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   5445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3045
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2400
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
      Left            =   1200
      TabIndex        =   4
      Top             =   2400
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
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
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
      Left            =   3840
      TabIndex        =   6
      Top             =   2400
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
      Left            =   2160
      TabIndex        =   0
      Top             =   120
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
      Left            =   1080
      MouseIcon       =   "FrmPadrão.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   120
      WhatsThisHelpID =   10246
      Width           =   1035
   End
   Begin VB.Label LblFrme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   2280
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
      Left            =   1080
      TabIndex        =   1
      Top             =   480
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
Attribute VB_Name = "FrmPadrão"
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
Private Sub LblId_Click(Index As Integer)
   Call F_PROCURAR(Index)
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
