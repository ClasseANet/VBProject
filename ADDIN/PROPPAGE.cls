VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PROPPAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum PageType
   tpNull
   tpClasse
   tpPropriedade
   tpMetodo
   tpEvento
   tpColecao
   tpForm
   tpModulo
   tpVariavel
End Enum

Private WithEvents mvarMe As FrmPropPage
Attribute mvarMe.VB_VarHelpID = -1

Private mvarTipoPagina As PageType
Private mvarNOME As String
Private mvarArgumentos As Collection
Private mvarRetorno As String
'Private mvarAsFriend As Boolean
'Private mvarMetodoPadrao As Boolean
Private mvarDescricao As String
Private mvarHelpFile As String
Private mvarHelpID As String
Private mvarEscopo As String
Private mvarTpDeclaracao As Integer
Private mvarPropriedadePadrao As Boolean
Private mvarClasse1 As Boolean
Private mvarColecaoClasseNova As Boolean
Private mvarClasseNova As Object
Private mvarisSaved As Boolean
Private mvarTIPO As String
Private mvarisConst As Boolean
Private mvarValorPadrao As String
Private mvarisEnum As Boolean
Private mvarisType As Boolean

Private mvarProjeto As clsPROJETO
Public Property Let ValorPadrao(ByVal vData As String)
   mvarValorPadrao = vData
End Property
Public Property Get ValorPadrao() As String
   ValorPadrao = mvarValorPadrao
End Property
Public Property Let isEnum(ByVal vData As Boolean)
   mvarisEnum = vData
End Property
Public Property Get isEnum() As Boolean
   isEnum = mvarisEnum
End Property
Public Property Let isType(ByVal vData As Boolean)
   mvarisType = vData
End Property
Public Property Get isType() As Boolean
   isType = mvarisType
End Property
Public Property Let isConst(ByVal vData As Boolean)
   mvarisConst = vData
End Property
Public Property Get isConst() As Boolean
   isConst = mvarisConst
End Property
Public Property Let TIPO(ByVal vData As String)
   mvarTIPO = vData
End Property
Public Property Get TIPO() As String
   TIPO = mvarTIPO
End Property
Public Property Set PROJETO(vData As clsPROJETO)
    Set mvarProjeto = vData
End Property
Public Property Let isSaved(ByVal vData As Boolean)
   mvarisSaved = vData
End Property
Public Property Get isSaved() As Boolean
   isSaved = mvarisSaved
End Property
Public Property Let NOME(ByVal vData As String)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Public Property Let Argumentos(ByVal vData As Collection)
   Set mvarArgumentos = vData
End Property
Public Property Get Argumentos() As Collection
   Set Argumentos = mvarArgumentos
End Property
Public Property Let Retorno(ByVal vData As String)
   mvarRetorno = vData
End Property
Public Property Get Retorno() As String
   Retorno = mvarRetorno
End Property
'Public Property Let AsFriend(ByVal vData As Boolean)
'   mvarAsFriend = vData
'   mvarMetodoPadrao = Not vData
'End Property
'Public Property Get AsFriend() As Boolean
'   AsFriend = mvarAsFriend
'End Property
'Public Property Let MetodoPadrao(ByVal vData As Boolean)
'   mvarMetodoPadrao = vData
'   mvarAsFriend = Not vData
'End Property
'Public Property Get MetodoPadrao() As Boolean
'   MetodoPadrao = mvarMetodoPadrao
'End Property
Public Property Let Descricao(ByVal vData As String)
   mvarDescricao = vData
End Property
Public Property Get Descricao() As String
   Descricao = mvarDescricao
End Property
Public Property Let HelpFile(ByVal vData As String)
   mvarHelpFile = vData
End Property
Public Property Get HelpFile() As String
   HelpFile = mvarHelpFile
End Property
Public Property Let HelpID(ByVal vData As String)
   mvarHelpID = vData
End Property
Public Property Get HelpID() As String
   HelpID = mvarHelpID
End Property
Public Property Let Escopo(ByVal vData As String)
   If Trim(vData) = "" Then vData = "Private"
   Select Case LCase(vData)
      Case "friend": TpDeclaracao = 1
      Case "public": TpDeclaracao = IIf(mvarTipoPagina = tpPropriedade, 0, 2)
      Case "global": TpDeclaracao = 3
      Case "private": TpDeclaracao = 4
      Case "static": TpDeclaracao = 5
      Case "dim": TpDeclaracao = 6
      Case Else:
         TpDeclaracao = IIf(mvarTipoPagina = tpPropriedade, 0, 6)
         vData = IIf(mvarTipoPagina = tpPropriedade, "Public", "Dim")
   End Select
   mvarEscopo = Trim(vData)
End Property
Public Property Get Escopo() As String
   Escopo = mvarEscopo
End Property
Public Property Let TpDeclaracao(ByVal vData As Integer)
   mvarTpDeclaracao = vData
End Property
Public Property Get TpDeclaracao() As Integer
   TpDeclaracao = mvarTpDeclaracao
End Property
'Public Property Let PropriedadePadrao(ByVal vData As Boolean)
'   mvarPropriedadePadrao = vData
'End Property
Public Property Get PropriedadePadrao() As Boolean
   PropriedadePadrao = mvarPropriedadePadrao
End Property
Public Property Let Classe1(ByVal vData As Boolean)
   mvarClasse1 = vData
End Property
Public Property Get Classe1() As Boolean
   Classe1 = mvarClasse1
End Property
Public Property Let ColecaoClasseNova(ByVal vData As Boolean)
   mvarColecaoClasseNova = vData
End Property
Public Property Get ColecaoClasseNova() As Boolean
   ColecaoClasseNova = mvarColecaoClasseNova
End Property
Public Property Let ClasseNova(ByVal vData As Object)
   mvarClasseNova = vData
End Property
Public Property Get ClasseNova() As Object
   Set ClasseNova = mvarClasseNova
End Property
Public Property Let TipoPagina(ByVal vData As PageType)
   mvarTipoPagina = vData
End Property
Public Property Get TipoPagina() As PageType
   TipoPagina = mvarTipoPagina
End Property
Public Sub Show()
   Dim n, MyAddIn As New CONSTRUTOR
   Dim MyArg As ADDARG
   Dim StrArg As String
   Dim i As Integer
      
   mvarisSaved = False
   If mvarTipoPagina = tpNull Then
      Exit Sub
   End If

   For Each n In mvarProjeto.CLASSES
      mvarMe.CmbReturn.AddItem n.NOME
   Next
   mvarMe.CmbReturn.ListIndex = 0
   
   If Trim(mvarNOME) <> "" Then
      '* Atualização de Membro
      With mvarMe
         .TxtNome = mvarNOME
         If mvarTipoPagina = tpVariavel Then
            If mvarisEnum Or mvarisType Then
               .GrdEnumType.RowHeightMin = 300
               .GrdEnumType.Rows = mvarArgumentos.Count + 1
               .GrdEnumType.TextMatrix(0, 0) = "Itens"
               .TxtItem.Width = IIf(.GrdEnumType.Rows <= 5, 4120, 3880)
               For i = 1 To mvarArgumentos.Count
                  .GrdEnumType.TextMatrix(i, 0) = mvarArgumentos(i)
               Next
               .OptDeclaração(0).Visible = False
               .OptDeclaração(1).Visible = False
               .OptDeclaração(2).Top = .OptDeclaração(0).Top
               .OptDeclaração(3).Top = .OptDeclaração(1).Top
               .OptDeclaração(4).Left = 1800
               .OptDeclaração(5).Left = 1800
               .OptDeclaração(6).Move 3300, .OptDeclaração(0).Top
               .GrdEnumType.Move 240, 900
               .TxtItem.Move .GrdEnumType.Left + 20
               .GrdEnumType.Row = 0
               .GrdEnumType.Row = 1
            End If
            If LocalizarCombo(.CmbBasedOn_DataType, mvarTIPO) = -1 Then mvarMe.CmbBasedOn_DataType.ListIndex = 0
         Else
            For Each n In mvarArgumentos
               Set MyArg = n
               StrArg = ""
               With MyArg
                  If .PorValor Then
                     StrArg = StrArg & "ByVal "
                  Else
                     If Not .PorRefImplicito Then StrArg = StrArg & "ByRef "
                  End If
                  StrArg = StrArg & .NOME & " "
                  If .Opcional Then
                     StrArg = "Optional " & StrArg
                     If .ValorPadrao <> "" Then StrArg = StrArg & "= " & .ValorPadrao
                  Else
                     If Not .DclImplicit Then StrArg = StrArg & " As " & .TIPO
                  End If
               End With
               .LstArg.AddItem Trim(StrArg)
            Next
         End If
         If LocalizarCombo(.CmbReturn, mvarRetorno) = -1 Then mvarMe.CmbReturn.ListIndex = 0
         If LocalizarCombo(.CmbEscopo, mvarEscopo) = -1 Then mvarMe.CmbEscopo.ListIndex = 0

         .ChkisConst.Value = IIf(mvarisConst, vbChecked, vbUnchecked)
         .TxtValorPadrao = mvarValorPadrao
         
         '.OptDefaultMetodo(0).Value = mvarAsFriend
         '.OptDefaultMetodo(1).Value = mvarMetodoPadrao
         .OptDeclaração(mvarTpDeclaracao).Value = True
         .TxtDesc = mvarDescricao
         .LblHelpFile = mvarHelpFile
         .TxtHelpID = mvarHelpID
         .ChkDefaultProp.Value = IIf(mvarPropriedadePadrao, vbChecked, vbUnchecked)
         .ChkTopLevel.Value = IIf(mvarClasse1, vbChecked, vbUnchecked)
         .OptCollOf(1).Value = mvarColecaoClasseNova
         If Not ClasseNova Is Nothing Then
            .TxtClassName = ClasseNova.NOME
        End If
      End With
   End If
   mvarMe.Show vbModal
End Sub
Public Function LinhaDeArgumentos() As String
   Dim MyArg As ADDARG, n As Variant
   Dim StrAux$
   For Each n In mvarArgumentos
      Set MyArg = n
      With MyArg
         If .Opcional Then StrAux$ = StrAux$ & "Optional "
         If .PorValor Then
            StrAux$ = StrAux$ & "ByVal "
         Else
            If Not .PorRefImplicito Then StrAux$ = StrAux$ & "ByRef "
         End If
         StrAux$ = StrAux$ & .NOME & " "
         If .Opcional Then
            If Trim(.ValorPadrao) <> "" Then StrAux$ = StrAux$ & "= " & .ValorPadrao
         Else
            If Not .DclImplicit Then StrAux$ = StrAux$ & "As " & .TIPO
         End If
         StrAux$ = Trim(StrAux$) & ", "
         Set MyArg = Nothing
      End With
   Next
   If Len(StrAux$) > 2 Then
      StrAux$ = Trim(Mid(StrAux$, 1, Len(StrAux$) - 2))
   End If
   LinhaDeArgumentos = Trim(StrAux$)
End Function
Public Sub SetArgumentos(StrFrase As String)
   Dim Pos As Integer
   Dim MyArg As ADDARG, n As Variant
   Dim StrWord As String
   Dim Palavras As Collection
   
   Set mvarArgumentos = Nothing
   Set mvarArgumentos = New Collection
   While Trim(StrFrase) <> ""
      Pos = InStr(StrFrase, ",")
      If Pos = 0 Then
         Pos = Len(StrFrase) + 1
      ElseIf InStr(StrFrase, "(") <> 0 Then
         Pos = InStr(StrFrase, ",")
         If Pos > InStr(StrFrase, "(") Then
         End If
      End If
      Set MyArg = New ADDARG
      StrWord = Trim(Mid(StrFrase, 1, Pos - 1))
      Set Palavras = GetPalavras(StrFrase)
      With MyArg
         For Each n In Palavras
            If Not PalavraReservada(CStr(n)) Then
               .NOME = n
               Exit For
            End If
         Next
         .Opcional = (InStr(UCase(StrWord), "OPTIONAL") <> 0)
         .PorValor = (InStr(UCase(StrWord), "BYVAL") <> 0)
         
         If Not .PorValor Then
            .PorRefImplicito = (InStr(UCase(StrWord), "BYREF") = 0)
         End If
         
         .DclImplicit = True
         If InStr(UCase(StrWord), " AS ") <> 0 Then
            .DclImplicit = False
            .TIPO = RichWordOver(StrWord, 0, 0, InStr(UCase(StrWord), " AS ") + 4)
         ElseIf VarWithDclImplicit(.NOME) Then
            .DclImplicit = True
            .TIPO = NameDclImplicit(.NOME)
         End If
         
         If InStr(UCase(StrWord), "=") <> 0 Then
            .ValorPadrao = Trim(Mid(StrWord, InStr(UCase(StrWord), "=") + 1))
         End If
      End With
     
      mvarArgumentos.Add MyArg
      Set MyArg = Nothing
      StrFrase = Trim(Mid(StrFrase, Pos + 1))
   Wend
End Sub

Private Sub Class_Initialize()
   Set mvarMe = New FrmPropPage
   Set mvarArgumentos = New Collection
   mvarTIPO = "Variant"
 '  Set mvarClasseNova = New PROPPAGE
End Sub

Private Sub mvarMe_Active()
   Call SetHourglass(mvarMe.hwnd)

'   Call Popula_
   Call SetDefault(mvarMe.hwnd)
End Sub

Private Sub mvarMe_ChkisConstClick()
   mvarMe.CmbBasedOn_DataType.Enabled = (mvarMe.ChkisConst.Value = vbUnchecked)
   mvarMe.TxtValorPadrao.Visible = (mvarMe.ChkisConst.Value = vbChecked)
End Sub
Private Sub mvarMe_CmdArgClick(Index As Integer)
   Dim MyArg As New ADDARG
   Dim i%, Existe As Boolean, StrNome$
   Dim n As Variant
   
   Select Case Index
      Case 0
         MyArg.Show
         If Trim(MyArg.NOME) <> "" Then
            For Each n In mvarArgumentos
               StrNome$ = n.NOME
               If n.DclImplicit And Not n.Opcional Then
                  StrNome$ = Mid(StrNome$, 1, Len(StrNome$) - 1)
               End If
               If UCase(StrNome$) = UCase(MyArg.NOME) Then
                  Existe = True
                  Exit For
               End If
            Next
            If Not Existe Then
               StrNome$ = ""
               If MyArg.PorValor Then StrNome$ = StrNome$ & "ByVal "
               StrNome$ = StrNome$ & MyArg.NOME & " "
               If MyArg.Opcional Then
                  StrNome$ = "Optional " & StrNome$
                  If Trim(MyArg.ValorPadrao) <> "" Then StrNome$ = StrNome$ & "= " & MyArg.ValorPadrao
               Else
                  If Not MyArg.DclImplicit Then StrNome$ = StrNome$ & " As " & MyArg.TIPO
               End If
               If Not MyArg.Opcional Then
                  i = 0
                  For Each n In mvarArgumentos
                     If n.Opcional Then Exit For
                     i = i + 1
                  Next
               Else
                 i% = mvarMe.LstArg.ListCount
               End If
               mvarMe.LstArg.AddItem StrNome$, i%
               If i% = 0 Then
                  mvarArgumentos.Add MyArg
               Else
                  mvarArgumentos.Add MyArg, , , i%
               End If
            End If
         End If
         
      Case 1
         i% = mvarMe.LstArg.ListIndex
         If i% >= 0 Then
            mvarMe.LstArg.RemoveItem i%
            mvarArgumentos.Remove i + 1
         End If
      
      Case 2
         i% = mvarMe.LstArg.ListIndex
         If i > 0 Then
            StrNome$ = mvarMe.LstArg.List(i - 1)
            mvarMe.LstArg.List(i - 1) = mvarMe.LstArg.List(i)
            mvarMe.LstArg.List(i) = StrNome
            mvarMe.LstArg.Selected(i - 1) = True
            
            Set MyArg = mvarArgumentos(i + 1)
            mvarArgumentos.Remove i + 1
            mvarArgumentos.Add MyArg, , i
         End If
      
      Case 3
         i% = mvarMe.LstArg.ListIndex
         If i + 1 < mvarMe.LstArg.ListCount Then
            StrNome$ = mvarMe.LstArg.List(i + 1)
            mvarMe.LstArg.List(i + 1) = mvarMe.LstArg.List(i)
            mvarMe.LstArg.List(i) = StrNome
            mvarMe.LstArg.Selected(i + 1) = True
            
            Set MyArg = mvarArgumentos(i + 1)
            mvarArgumentos.Remove i + 1
            mvarArgumentos.Add MyArg, , , i + 1
         End If
   End Select
   mvarMe.LstArg.SetFocus
   Set MyArg = Nothing
   
End Sub

Private Sub mvarMe_CmdOperClick(Index As Integer)
   Dim i As Integer, n As Variant
   Dim bSair As Boolean, Opc1 As Boolean

   mvarisSaved = False
   If Index = 0 Then '* OK
      mvarNOME = mvarMe.TxtNome
      mvarDescricao = mvarMe.TxtDesc
      mvarHelpFile = mvarMe.LblHelpFile
      mvarHelpID = mvarMe.TxtHelpID
      
      If Me.TipoPagina = tpVariavel Then
         mvarEscopo = RichWordOver(mvarMe.OptDeclaração(mvarTpDeclaracao).Caption, 0, 0, 1)
      ElseIf Me.TipoPagina = tpMetodo Then
         With mvarMe.CmbEscopo
            mvarEscopo = IIf(.ListIndex = 0, "Private", .Text)
         End With
      End If
      With mvarMe.CmbReturn
         mvarRetorno = IIf(.ListIndex = 0, "", .Text)
      End With
      If mvarMe.CmbBasedOn_DataType.ListIndex = 0 Then
         mvarTIPO = "Variant"
      Else
         mvarTIPO = mvarMe.CmbBasedOn_DataType
      End If
      mvarisConst = (mvarMe.ChkisConst.Value = vbChecked)
      mvarValorPadrao = IIf(mvarisConst, mvarMe.TxtValorPadrao.Text, "")
      mvarisSaved = True
   End If
   '* Verifica se toas as declarações do tipo "Optional"
   '* está no final doas argumentos
   i = 0
   bSair = True
   Opc1 = False
   If mvarTipoPagina = tpVariavel Then
   
   Else
      For Each n In Argumentos
         If Not Opc1 Then
            Opc1 = n.Opcional
         End If
         If Opc1 And Not n.Opcional Then
            Call ExibirAviso("Argumento em Ordem inválida.", "Atenção !")
            mvarMe.LstArg.Selected(i) = True
            bSair = False
         End If
        i = i + 1
      Next
   End If
   If bSair Then Unload mvarMe
End Sub
Private Sub mvarMe_Load()
   Dim i%, Pos%
   Call SetHourglass(mvarMe.hwnd)
   
   With mvarMe
      .FrmeColl.Visible = False
      .FrmeMetoEve.Visible = False
      .FrmeMetodo.Visible = False
      .FrmeProp.Visible = False
      .ChkTopLevel.Visible = False
      .CmbBasedOn_DataType.Visible = False
      .Lbl(1).Visible = False
      .GrdEnumType.ColWidth(0) = .GrdEnumType.Width
      .TabProp.TabVisible(2) = False
      .TabProp.TabsPerRow = 2
      Select Case mvarTipoPagina
         Case tpClasse
            .Caption = "Propriedades - CLASSE"
            .ChkTopLevel.Visible = True
            .Lbl(1).Visible = True
            .Lbl(1) = "Baseado em :"
            .CmbBasedOn_DataType.Visible = True
            .TabProp.TabVisible(2) = True
            .TabProp.TabsPerRow = 3
         Case tpColecao
            .Caption = "Propriedades - COLEÇÃO"
            .FrmeColl.Visible = True
            .Lbl(1).Visible = True
            .Lbl(1) = "Baseado em :"
            .CmbBasedOn_DataType.Visible = True
            .TabProp.TabVisible(2) = True
            .TabProp.TabsPerRow = 3
         Case tpEvento
            .Caption = "Propriedades - EVENTO"
            .FrmeMetoEve.Visible = True
         Case tpForm
            .Caption = "Propriedades - FORMULÁRIO"
         Case tpMetodo
            .Caption = "Propriedades - MÉTODO"
            .FrmeMetoEve.Visible = True
            .FrmeMetodo.Visible = True
         Case tpModulo:
            .Caption = "Propriedades - MÓDULO"
         Case tpPropriedade, tpVariavel
            .Caption = "Propriedades - VARIÁVEL, CONSTANTE" & IIf(mvarTipoPagina = tpPropriedade, ", PROPRIEDADE", "")
            .FrmeProp.Visible = True
            .Lbl(1).Visible = True
            .Lbl(1) = "Tipo de Dado : "
            .ChkDefaultProp.Visible = (mvarTipoPagina = tpPropriedade)
            .ChkisConst.Visible = (mvarTipoPagina = tpVariavel)
            .CmbBasedOn_DataType.Enabled = Not (mvarisEnum Or mvarisType)
            .GrdEnumType.Visible = (mvarisEnum Or mvarisType)
            .TxtItem.Visible = (mvarisEnum Or mvarisType)
            For i = 0 To 6
               If i <= 1 Then
                  .OptDeclaração(i).Enabled = (mvarTipoPagina = tpPropriedade)
               Else
                  .OptDeclaração(i).Enabled = (mvarTipoPagina = tpVariavel)
               End If
            Next
            .OptDeclaração(2).Enabled = True
            With .CmbBasedOn_DataType
               .Visible = True
               .AddItem "<Nenhum>"
               .AddItem "Byte"
               .AddItem "Boolean"
               .AddItem "Enum"
               .AddItem "Integer"
               .AddItem "Long"
               .AddItem "Single"
               .AddItem "Double"
               .AddItem "Currency"
               .AddItem "Date"
               .AddItem "String"
               .AddItem "Type"
               .AddItem "Variant"
               .AddItem "Object"
            End With
      End Select
      .Move .Left, .Top, 5865, 6255
   End With
   
   Call ConfigForm(mvarMe, SysMdi.Icon, Sys.Proj.FundoTela)
   Call SetDefault(mvarMe.hwnd)
End Sub

Private Sub mvarMe_LstArgDblClick()
   Dim ClsAdd As New ADDARG
   Dim Pos%, Palavras As New Collection, n
   Dim StrArg$, i%
   
   With mvarArgumentos(mvarMe.LstArg.ListIndex + 1)
      .Show
      If Trim(.NOME) <> "" And .isSaved Then

         StrArg = IIf(.Opcional, "Optional ", "")
         StrArg = StrArg & IIf(.PorValor, "ByVal ", "")
         If Not .PorValor Then .PorRefImplicito = True
         StrArg = StrArg & .NOME
         If .Opcional Then
            StrArg = StrArg & IIf(Trim(.ValorPadrao) <> "", " = " & .ValorPadrao, "")
         End If
         If Not VarWithDclImplicit(.NOME) Then
            StrArg = StrArg & IIf(Trim(.TIPO) = "", "", " As " & .TIPO)
         End If
         mvarMe.LstArg.List(mvarMe.LstArg.ListIndex) = StrArg
      End If

   End With
   mvarMe.LstArg.SetFocus
End Sub

Private Sub mvarMe_OptDeclaraçãoClick(Index As Integer)
   mvarTpDeclaracao = Index
End Sub

Private Sub mvarMe_Resize()
   With mvarMe
      If .Width > 2200 And .Height > 2200 Then
         '* Ajustar Frames
         .FrmeColl.Move 240, 1680
         .FrmeMetoEve.Move 240, 1080
         .FrmeProp.Move 240, 1680
         '* Ajustar outros objetos
         .TabProp.Move 240, 240, .Width - 600, .Height - (1000 + .CmdOper(1).Height)
         .CmdOper(0).Move .Width - (2 * (.CmdOper(0).Width + 360)), .Height - (.CmdOper(0).Height + 500)
         .CmdOper(1).Move .Width - (.CmdOper(1).Width + 360), .Height - (.CmdOper(1).Height + 500)
      End If
      Call PintarFundo(mvarMe, Sys.Proj.FundoTela)
   End With
End Sub
