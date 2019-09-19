VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SenhaEsp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe As FrmSenhaEsp
Attribute mvarMe.VB_VarHelpID = -1

Private mvarxDb         As Object
Private mvarPermissao   As Boolean

Private mvarIDPROJ      As String
Private mvarSETOR       As String
Private mvarIDREQ       As String
Private mvarANOREQ      As String

Private mvarEVENTO      As String

Private mvarMOTIVO      As String
Private mvarSENHA       As String
Private mvarICON        As Variant
Private mvarFundoTela   As String

Private mvarHabilitaSenha     As Boolean
Private mvarHabilitaMotivo    As Boolean
Private mvarNecessitaMotivo   As Boolean
Public Property Let FundoTela(ByVal vData As Double)
   mvarFundoTela = vData
End Property
Public Property Get FundoTela() As Double
   FundoTela = mvarFundoTela
End Property
Public Property Let ICON(ByVal vData As Variant)
   mvarICON = vData
End Property
Public Property Get ICON() As Variant
   Set ICON = mvarICON
End Property
Public Property Let MOTIVO(ByVal vData As String)
   mvarMOTIVO = vData
End Property
Public Property Get MOTIVO() As String
   MOTIVO = mvarMOTIVO
End Property
Public Property Let SENHA(ByVal vData As String)
   mvarSENHA = vData
End Property
Public Property Get SENHA() As String
   SENHA = mvarSENHA
End Property
Public Property Let EVENTO(ByVal vData As String)
   mvarEVENTO = vData
End Property
Public Property Get EVENTO() As String
   EVENTO = mvarEVENTO
End Property
Public Property Let Permissao(ByVal vData As Boolean)
   mvarPermissao = vData
End Property
Public Property Get Permissao() As Boolean
   Permissao = mvarPermissao
End Property
Public Property Let NecessitaMotivo(ByVal vData As Boolean)
   mvarNecessitaMotivo = vData
End Property
Public Property Get NecessitaMotivo() As Boolean
   NecessitaMotivo = mvarNecessitaMotivo
End Property
Public Property Let HabilitaMotivo(ByVal vData As Boolean)
   mvarHabilitaMotivo = vData
End Property
Public Property Get HabilitaMotivo() As Boolean
   HabilitaMotivo = mvarHabilitaMotivo
End Property
Public Property Let HabilitaSenha(ByVal vData As Boolean)
   mvarHabilitaSenha = vData
End Property
Public Property Get HabilitaSenha() As Boolean
   HabilitaSenha = mvarHabilitaSenha
End Property
Public Property Let xDb(ByVal vData As Object)
   Set mvarxDb = vData
End Property
Public Property Get xDb() As Object
   Set xDb = mvarxDb
End Property
Public Property Set xDb(ByVal vData As Object)
   Set mvarxDb = vData
End Property
Public Property Let IDPROJ(ByVal vData As String)
   mvarIDPROJ = vData
End Property
Public Property Get IDPROJ() As String
   IDPROJ = mvarIDPROJ
End Property
Public Property Let SETOR(ByVal vData As String)
   mvarSETOR = vData
End Property
Public Property Get SETOR() As String
   SETOR = mvarSETOR
End Property
Public Property Let IDREQ(ByVal vData As String)
   mvarIDREQ = vData
End Property
Public Property Get IDREQ() As String
   IDREQ = mvarIDREQ
End Property
Public Property Let ANOREQ(ByVal vData As String)
   mvarANOREQ = vData
End Property
Public Property Get ANOREQ() As String
   ANOREQ = mvarANOREQ
End Property
Public Sub Show(Optional Pesquisa = True)
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmSenhaEsp
   mvarHabilitaSenha = True
   mvarHabilitaMotivo = True
   mvarNecessitaMotivo = True
   mvarPermissao = False
   mvarFundoTela = "FUNDO"
End Sub
Private Sub Class_Terminate()
   Set mvarxDb = Nothing
End Sub
Private Sub mvarMe_Active()
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdOperClick(Index As Integer)
   Select Case Index
      Case 0
      Case 1
         Call VerificarPermissao
         If mvarNecessitaMotivo And Trim(mvarMOTIVO) = "" Then
            Call ClsMsg.ExibirAviso("Campo de Motivo Inválido", ClsMsg.LoadMsg(1))
            mvarMe.TxtMOTIVO.SetFocus
            Exit Sub
         End If
   End Select
   Unload mvarMe
End Sub
Private Sub mvarMe_KeyPress(KeyAscii As Integer)
   KeyAscii = ClsDsr.SendTab(mvarMe, KeyAscii, vbString, , False)
End Sub
Private Sub mvarMe_Load()
   Call MontarTela
   Call ClsCtrl.ConfigForm(mvarMe, mvarICON, mvarFundoTela)
End Sub
Private Sub mvarMe_Resize()
   Dim nAux As Double
   On Error Resume Next
   nAux = mvarMe.TxtSENHA.FontSize
   mvarMe.TxtSENHA.FontSize = 10
   
   mvarMe.TxtSENHA.Move mvarMe.TxtSENHA.Left, mvarMe.TxtSENHA.Top, mvarMe.Width - 465, 735
   mvarMe.TxtMOTIVO.Move mvarMe.TxtMOTIVO.Left, mvarMe.TxtMOTIVO.Top, mvarMe.Width - 465, mvarMe.Height - mvarMe.TxtMOTIVO.Top - mvarMe.CmdOper(0).Height - 720
   
   mvarMe.CmdOper(0).Move mvarMe.Width - mvarMe.CmdOper(0).Width - 360, mvarMe.Height - mvarMe.CmdOper(0).Height - 600
   mvarMe.CmdOper(1).Move mvarMe.CmdOper(0).Left - mvarMe.CmdOper(1).Width - 360, mvarMe.CmdOper(0).Top
   
   mvarMe.TxtSENHA.FontSize = nAux
   
   Call ClsCtrl.PintarFundo(mvarMe, mvarFundoTela)
End Sub
Public Sub VerificarPermissao()
   Dim ChaveNula As Boolean
   Dim Existe As Boolean
     
   mvarPermissao = True
     
   ChaveNula = (mvarIDPROJ = "" Or mvarSETOR = "" Or mvarIDREQ = "" Or mvarANOREQ = "")
   If Not ChaveNula Then
   End If
   
   mvarSENHA = mvarMe.TxtSENHA.Text
   mvarMOTIVO = mvarMe.TxtMOTIVO.Text
   
   If mvarHabilitaSenha And mvarPermissao Then
      
   End If
   If mvarHabilitaMotivo And mvarPermissao And mvarNecessitaMotivo Then
      mvarPermissao = Not (Trim(mvarMOTIVO) = "")
   End If
   
End Sub
Private Sub MontarTela()
   mvarMe.Lbl(0).Enabled = mvarHabilitaSenha
   mvarMe.TxtSENHA.Enabled = mvarHabilitaSenha
   mvarMe.Lbl(1).Enabled = mvarHabilitaMotivo
   mvarMe.TxtMOTIVO.Enabled = mvarHabilitaMotivo
   
   mvarMe.TxtSENHA.BackColor = IIf(mvarMe.TxtSENHA.Enabled, vbWhite, vbInactiveCaptionText)
   mvarMe.TxtMOTIVO.BackColor = IIf(mvarMe.TxtMOTIVO.Enabled, vbWhite, vbInactiveCaptionText)
   
   mvarMe.TxtMOTIVO.Text = IIf(mvarMe.TxtMOTIVO.Enabled, mvarMOTIVO, "")
End Sub

