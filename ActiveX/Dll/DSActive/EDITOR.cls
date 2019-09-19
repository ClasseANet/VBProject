VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDITOR"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private WithEvents mvarMe As FrmEditor
Attribute mvarMe.VB_VarHelpID = -1

Private mvarAplic       As Object
Private mvarFundoTela   As String
Private mvarIdioma      As Integer
Private mvarNomeArq     As String
Private mvarTexto       As String
Private mvarTextoOrigem As String
Private mvarEdicao      As Boolean
Private mvarCaption     As String
Private mvarWidth       As Long
Private mvarHeight      As Long

Private mvarFontName    As String
Private mvarFontSize    As Double
Private Resizing        As Boolean
Private mvarBtnPrint    As Boolean
Private mvarCabDetalhe  As String
Private mvarTituloRel   As String
Public Property Let TituloRel(ByVal vData As String)
   mvarTituloRel = vData
End Property
Public Property Get TituloRel() As String
    TituloRel = mvarTituloRel
End Property
Public Property Let CabDetalhe(ByVal vData As String)
   mvarCabDetalhe = vData
End Property
Public Property Get CabDetalhe() As String
    CabDetalhe = mvarCabDetalhe
End Property
Public Property Let BtnPrint(ByVal vData As Boolean)
   mvarBtnPrint = vData
End Property
Public Property Get BtnPrint() As Boolean
    BtnPrint = mvarBtnPrint
End Property
Public Property Let FontSize(ByVal vData As Double)
   mvarFontSize = vData
End Property
Public Property Get FontSize() As Double
    FontSize = mvarFontSize
End Property
Public Property Let FontName(ByVal vData As String)
   mvarFontName = vData
End Property
Public Property Get FontName() As String
    FontName = mvarFontName
End Property
Public Property Let Width(ByVal vData As Long)
   mvarMe.Width = vData
   mvarWidth = vData
End Property
Public Property Get Width() As Long
    Width = mvarWidth
End Property
Public Property Let Height(ByVal vData As Long)
    mvarMe.Height = vData
    mvarHeight = vData
End Property
Public Property Get Height() As Long
    Height = mvarHeight
End Property
Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property
Public Property Get Caption() As String
    Caption = mvarCaption
End Property
Public Property Let Edicao(ByVal vData As Boolean)
    mvarEdicao = vData
End Property
Public Property Get Edicao() As Boolean
    Edicao = mvarEdicao
End Property
Public Property Let Texto(ByVal vData As String)
    mvarTexto = vData
    mvarTextoOrigem = vData
End Property
Public Property Get Texto() As String
    Texto = mvarTexto
End Property
Public Property Let FundoTela(ByVal vData As String)
    mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
    FundoTela = mvarFundoTela
End Property
Public Property Let NomeArq(ByVal vData As String)
    mvarNomeArq = vData
End Property
Public Property Get NomeArq() As String
    NomeArq = mvarNomeArq
End Property
Public Sub Show(Optional Mode As Integer = vbModal)
   mvarMe.Show Mode
End Sub
Public Property Let Aplic(ByVal vData As Object)
    Set mvarAplic = vData
End Property
Public Property Get Aplic() As Object
    Set Aplic = mvarAplic
End Property

Private Sub Class_Initialize()
   Set mvarMe = New FrmEditor
   mvarCaption = "Erros"
   mvarFontName = "MS Sans Serif"
   mvarFontSize = 8.25
   mvarFundoTela = "FUNDO"
   mvarIdioma = 5000
'   mvarWidth = mvarMe.Width
'   mvarHeight = mvarMe.Height
End Sub
Public Sub SetIcon(frm As Object)
    mvarMe.ICON = Nothing
    mvarMe.ICON = frm.ICON
End Sub
Private Sub mvarMe_Active()
   mvarMe.Caption = mvarCaption
   If mvarTexto = "" And UCase(Me.Caption) = "ERROS" Then
      mvarNomeArq = ClsDos.DrvTmp & "error.log"
   End If
   mvarMe.TxtTexto.Font.Name = mvarFontName
   mvarMe.TxtTexto.Font.SIZE = mvarFontSize
   mvarMe.TxtTexto.Text = ""
   If Trim(mvarNomeArq) = "" Then
      mvarTexto = mvarTextoOrigem
      mvarMe.TxtTexto.Text = mvarTexto
   Else
      mvarMe.TxtTexto.FileName = mvarNomeArq
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_CmdOperClick(index As Integer)
   Select Case index
      Case 0: Unload mvarMe
      Case 1:
         mvarMe.TxtTexto.Text = mvarTextoOrigem
         'SendKeys "{End}"
      Case 2: mvarMe.TxtTexto.SaveFile mvarNomeArq '* Salvar
   End Select
End Sub
Private Sub mvarMe_CmdImprimirClick()
   Dim MyPrint As IMPRESSAO
   Screen.MousePointer = vbHourglass
   mvarMe.CmdImprimir.MousePointer = vbHourglass
   
   Set MyPrint = New IMPRESSAO
   With MyPrint
      .CryRpt = mvarMe.CryRprt
      .Rpt_Drive = mvarAplic.Path & "\" 'dbdrive  ''"C:\TMP\"
      If mvarTituloRel = "" Then
         .Titulo = UCase(ClsMsg.LoadOriMsg(mvarIdioma + 62) & ClsMsg.LoadOriMsg(mvarIdioma + 42)) & mvarCaption
      Else
         .Titulo = mvarTituloRel
      End If
      .Aplic = mvarAplic
      .Idioma = mvarIdioma
      Call .ImprimeText(mvarMe.TxtTexto, mvarCabDetalhe)
   End With
   Set MyPrint = Nothing
      
   Screen.MousePointer = vbDefault
   mvarMe.CmdImprimir.MousePointer = 99
End Sub

Private Sub mvarMe_Load()
   Call ClsDsr.SetHourglass(mvarMe.hWnd)
   
'   mvarMe.TxtTexto.Enabled = mvarEdicao
   
   If mvarWidth = 0 Then mvarWidth = mvarMe.Width
   If mvarHeight = 0 Then mvarHeight = mvarMe.Height
   mvarMe.Width = mvarWidth
   mvarMe.Height = mvarHeight
   
   mvarMe.CmdOper(0).Move mvarMe.Width - mvarMe.CmdOper(0).Width - 360, mvarMe.Height - mvarMe.CmdOper(0).Height - 520
   mvarMe.CmdOper(1).Move mvarMe.CmdOper(0).Left - mvarMe.CmdOper(0).Width - 380, mvarMe.CmdOper(0).Top
   mvarMe.TxtTexto.Move 120, 120, mvarMe.Width - 360, mvarMe.Height - 1140
   mvarMe.CmdImprimir.Top = mvarBtnPrint
    
   Call ClsCtrl.ConfigForm(mvarMe, , mvarFundoTela)
   Call ClsDsr.SetDefault(mvarMe.hWnd)
End Sub

Private Sub mvarMe_Resize()
   On Error Resume Next
   If Resizing Then Exit Sub
   Resizing = True
   If Not mvarMe.PrimeiraVez Then
      mvarMe.CmdOper(0).Move mvarMe.Width - mvarMe.CmdOper(0).Width - 360, mvarMe.Height - mvarMe.CmdOper(0).Height - 520
      mvarMe.CmdOper(1).Move mvarMe.CmdOper(0).Left - mvarMe.CmdOper(0).Width - 380, mvarMe.CmdOper(0).Top
      mvarMe.TxtTexto.Move 120, 120, mvarMe.Width - 360, mvarMe.Height - 1140
      mvarMe.CmdImprimir.Top = mvarMe.CmdOper(0).Top
      Call ClsCtrl.PintarFundo(mvarMe, mvarFundoTela)
   End If
   Resizing = False
End Sub

Private Sub mvarMe_TxtTextoChange()
   mvarTexto = mvarMe.TxtTexto.Text
   If mvarMe.CmdOper(1).Visible Then
      mvarMe.CmdOper(1).Enabled = (mvarTexto <> mvarTextoOrigem)
   End If
End Sub

Private Sub mvarMe_TxtTextoKeyDown(KeyCode As Integer, Shift As Integer)
   If Not mvarEdicao Then
      If ClsDsr.InArray(KeyCode, Array(vbKeyDelete, vbKeyBack)) Then
         KeyCode = 0
      End If
   End If
End Sub
Private Sub mvarMe_TxtTextoKeyPress(KeyAscii As Integer)
   If Not mvarEdicao Then
      KeyAscii = 0
   End If
End Sub
Public Sub ExibirMsg(Optional MsgPositiva As Boolean = True, Optional msg, Optional NumPisca = 2)
   On Error Resume Next
   With FrmMessage
      If IsMissing(msg) Then
         msg = IIf(MsgPositiva, "Operação Realizada com Sucesso.", "Operação Falhou !")
      End If
      .Mensagem = msg
      .MsgPositiva = MsgPositiva
      .NumPisca = NumPisca
      .Show vbModal
   End With
End Sub
