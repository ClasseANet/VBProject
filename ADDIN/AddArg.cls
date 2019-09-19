VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADDARG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents mvarMe As FrmAddArg
Attribute mvarMe.VB_VarHelpID = -1
Private mvarNOME As String
Private mvarTIPO As String
Private mvarPorValor As Boolean
Private mvarOpcional As Boolean
Private mvarValorPadrao As String
Private mvarDclImplicit As Boolean
Private mvarPorRefImplicito As Boolean
Private mvarisSaved As Boolean
Public Property Let isSaved(ByVal vData As Boolean)
   mvarisSaved = vData
End Property
Public Property Get isSaved() As Boolean
   isSaved = mvarisSaved
End Property
Public Property Let DclImplicit(ByVal vData As Boolean)
   mvarDclImplicit = vData
End Property
Public Property Get DclImplicit() As Boolean
   DclImplicit = mvarDclImplicit
End Property
Public Property Let PorRefImplicito(ByVal vData As Boolean)
   mvarPorValor = Not vData
   mvarPorRefImplicito = vData
End Property
Public Property Get PorRefImplicito() As Boolean
   PorRefImplicito = mvarPorRefImplicito
End Property
Public Property Let NOME(ByVal vData As String)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Public Property Let TIPO(ByVal vData As String)
   mvarTIPO = vData
End Property
Public Property Get TIPO() As String
   TIPO = mvarTIPO
End Property
Public Property Let PorValor(ByVal vData As Boolean)
   mvarPorValor = vData
End Property
Public Property Get PorValor() As Boolean
   PorValor = mvarPorValor
End Property
Public Property Let Opcional(ByVal vData As Boolean)
   If vData Then mvarTIPO = "Variant"
   mvarDclImplicit = vData
   mvarOpcional = vData
End Property
Public Property Get Opcional() As Boolean
   Opcional = mvarOpcional
End Property
Public Property Let ValorPadrao(ByVal vData As String)
   mvarValorPadrao = vData
End Property
Public Property Get ValorPadrao() As String
   ValorPadrao = mvarValorPadrao
End Property

Public Sub Show()
   Dim i%
   With mvarMe
      .LstTipo.Selected(9) = True '* Variant
      If Trim(mvarNOME) <> "" Then
         .TxtNome = mvarNOME
         For i = 0 To .LstTipo.ListCount - 1
            If .LstTipo.List(i) = mvarTIPO Then
               .LstTipo.Selected(i) = True
               Exit For
            End If
         Next
         .ChkByVal.Value = IIf(mvarPorValor, vbChecked, vbUnchecked)
         If mvarOpcional Then
            .ChkOptional.Value = vbChecked
            .TxtDefaultValue = mvarValorPadrao
         End If
         If Trim(mvarTIPO) = "" Then mvarTIPO = "Variant"
         If VarWithDclImplicit(mvarNOME) Then mvarTIPO = NameDclImplicit(mvarNOME)
         Call LocalizarList(.LstTipo, mvarTIPO)
      End If
      .Show vbModal
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmAddArg
   mvarPorRefImplicito = True
   mvarPorValor = False
End Sub

Private Sub mvarMe_ChkOptionalClick()
   With mvarMe
      .Lbl(4).Visible = (.ChkOptional.Value = vbChecked)
      .TxtDefaultValue.Visible = (.ChkOptional.Value = vbChecked)
   End With
End Sub

Private Sub mvarMe_CmdOperClick(Index As Integer)
   
   If Index = 0 Then  '* OK
      mvarisSaved = True
      mvarNOME = mvarMe.TxtNome
      
      If mvarMe.ChkOptional.Value Then
         mvarTIPO = ""
      Else
          mvarDclImplicit = VarWithDclImplicit(mvarNOME)
          If mvarDclImplicit Then
             mvarTIPO = NameDclImplicit(mvarNOME)
          Else
             mvarTIPO = mvarMe.LstTipo.Text
          End If
      End If
      mvarPorValor = (mvarMe.ChkByVal.Value = vbChecked)
      mvarOpcional = (mvarMe.ChkOptional.Value = vbChecked)
      mvarValorPadrao = mvarMe.TxtDefaultValue
   Else
      mvarisSaved = False
   End If
   Unload mvarMe
End Sub

Private Sub mvarMe_Load()
   Call SetHourglass(mvarMe.hwnd)
   Call ConfigForm(mvarMe, SysMdi.Icon, Sys.Proj.FundoTela)
   Call SetDefault(mvarMe.hwnd)
End Sub

Private Sub mvarMe_TxtNomeKeyPress(KeyAscii As Integer)
   If mvarMe.ChkOptional.Value = vbChecked Then
      If InArray(Right(Chr(KeyAscii), 1), Array("!", "#", "$", "%", "&")) Then
         Call ExibirAviso("Parâmetro opcional deve ser do tipo 'Variant'.", mvarMe.Caption)
         KeyAscii = 0
      End If
   Else
      Select Case KeyAscii
         Case Asc("!"): Call LocalizarList(mvarMe.LstTipo, "Single")
         Case Asc("#"): Call LocalizarList(mvarMe.LstTipo, "Double")
         Case Asc("$"): Call LocalizarList(mvarMe.LstTipo, "String")
         Case Asc("%"): Call LocalizarList(mvarMe.LstTipo, "Integer")
         Case Asc("&"): Call LocalizarList(mvarMe.LstTipo, "Long")
      End Select
   End If
End Sub
