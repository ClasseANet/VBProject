VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CONSTRUTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"AreaTemp"
Attribute VB_Ext_KEY = "Member1" ,"clsPROJETO"
Option Explicit
Private WithEvents mvarMe As FrmAddIn
Attribute mvarMe.VB_VarHelpID = -1
Private mvarPROJETOS As Collection
Private mvarPROJETO As PROJETO
Private mvarClsConnect As Connect
Private mvarVbProj As VBProject

Private mvarVbComp As VBComponent
Private mvarVbMember As MEMBER
Private mvarVbCode As CodeModule

Private Icon As TypeIcon
Private mvarAreaTemp As AreaTemp

Private cCodigoAntes As String
Private cLinhaAntes As Long
Private cCompAtual As String

Public Enum ShortcutTypes
    vbwFile = 0
    vbwFolder = 1
    vbwURL = 2
    vbwEMail = 3
End Enum
Private Type TypeIcon
   VBPImage As String
   FRMImage As String
   CLSImage As String
   BASImage As String
   RESImage As String
   OpenImage As String
   CloseImage As String
   CollImage As String
   EvtImage As String
   MetImage As String
   PropImage As String
   VarImage As String
   ConstImage As String
End Type
Public Enum IndMnuMouse
   mm_ADICIONAR = 0
   mm_ADICIONAR_OBJ
   mm_SALVAR
   mm_Salvar_COMO
   mm_SEP1
   mm_RECORTAR
   mm_COPIAR
   mm_COLAR
   mm_SEP2
   mm_EXCLUIR
   mm_RENOMEAR
   mm_SEP3
   mm_PROPRIEDADES
   mm_DEFINIR_COMO
   mm_SEP4
   mm_IR_PROJETO
   mm_RECARREGAR_OBJ
End Enum
Public Property Get AreaTemp() As AreaTemp
    If mvarAreaTemp Is Nothing Then
        Set mvarAreaTemp = New AreaTemp
    End If
    Set AreaTemp = mvarAreaTemp
End Property
Public Property Set AreaTemp(vData As AreaTemp)
    Set mvarAreaTemp = vData
End Property
Public Property Get VbCode() As CodeModule
   Set VbCode = mvarVbCode
End Property
Public Property Set VbCode(vData As CodeModule)
   Set mvarVbCode = vData
End Property
Public Property Get VbMember() As MEMBER
   Set VbMember = mvarVbMember
End Property
Public Property Set VbMember(vData As MEMBER)
   Set mvarVbMember = vData
End Property
Public Property Get VbComp() As VBComponent
   Set VbComp = mvarVbComp
End Property
Public Property Set VbComp(vData As VBComponent)
   Set mvarVbComp = vData
End Property
Public Property Get VbProj() As VBProject
   Set VbProj = mvarVbProj
End Property
Public Property Set VbProj(vData As VBProject)
   Dim n, Existe As Boolean
   Set mvarVbProj = vData
   For Each n In PROJETOS
      Existe = (n.NOME = vData.Name)
      If Existe Then Exit For
   Next
   If Not Existe Then
      Call CarregaProjeto
   End If
   Set mvarVbProj = vData
End Property
Public Property Get ClsConnect() As Connect
   Set ClsConnect = mvarClsConnect
End Property
Public Property Set ClsConnect(vData As Connect)
   Set mvarClsConnect = vData
End Property
Public Property Get PROJETO(NOME$) As PROJETO
   If mvarPROJETO Is Nothing Then
       Set mvarPROJETO = New PROJETO
   End If
   Set mvarPROJETO = PROJETOS(NOME$)
   Set PROJETO = mvarPROJETO
End Property
'Public Property Set PROJETO(vData As PROJETO)
'    Set mvarPROJETO = vData
'End Property
Public Property Get PROJETOS() As Collection
   If mvarPROJETOS Is Nothing Then
       Set mvarPROJETOS = New Collection
   End If
   Set PROJETOS = mvarPROJETOS
End Property
Public Property Set PROJETOS(vData As Collection)
   Set mvarPROJETOS = vData
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmAddIn
   With Icon
      .VBPImage = "VBP"
      .FRMImage = "FRM"
      .CLSImage = "CLS"
      .BASImage = "BAS"
      .RESImage = "RES"
      .OpenImage = "OPEN"
      .CloseImage = "CLOSED"
      
      .CollImage = "COLECAO"
      .EvtImage = "EVENTO"
      .MetImage = "METODO"
      .PropImage = "PROPRIEDADE"
      .VarImage = "VARIAVEL"
      .ConstImage = "CONSTANTE"
   End With
   
End Sub
Private Sub Class_Terminate()
  Set mvarAreaTemp = Nothing
  Set mvarPROJETO = Nothing
  'Set mvarPROJ = Nothing
  Set mvarPROJETOS = Nothing
End Sub
Sub Show()
   On Error Resume Next
   If mvarMe Is Nothing Then
       Set mvarMe = New FrmAddIn
   End If
   mvarClsConnect.FormDisplayed = True
   
   'Set mvarMe.VBInstance = VBInstance
   'Set mvarMe.Connect = mvarClsConnect
   ClsConnect.FormDisplayed = True
   mvarMe.Visible = True
   mvarMe.Show
   mvarMe.ZOrder
End Sub
Sub Hide()
   On Error Resume Next
   ClsConnect.FormDisplayed = False
   ClsConnect.Hide
   'mvarMe.Hide
   Unload mvarMe
End Sub
Private Sub CarregaProjeto(Optional vProj)
   Set mvarPROJETO = New PROJETO
   With mvarPROJETO
      .NOME = mvarVbProj.Name
      .DESCRIÇÃO = mvarVbProj.Description
      .FILENAME = mvarVbProj.FILENAME
      .NOME = mvarVbProj.Name
      .PATH = mvarVbProj.FILENAME
   End With
   mvarPROJETOS.Add mvarPROJETO, mvarPROJETO.NOME
   Call CarregaComponentes
End Sub
Private Sub CarregaComponentes(Optional vProj)
   Dim MyComp As VBComponent
   Dim MyObj As Variant, n As Variant
   With mvarPROJETO
      For Each n In mvarVbProj.VBComponents
         Set MyComp = n
         Select Case n.Type
            Case vbext_ct_VBMDIForm, vbext_ct_VBForm
               Set MyObj = New FORMULARIO
               With MyObj
                  .FILENAME = GetNameFromPath(MyComp.FileNames(1))
                  .NOME = MyComp.Name
                  .PATH = GetNameFromPath(MyComp.FileNames(1), 1)
                  .MDI = (vbext_ct_VBMDIForm = n.Type)
               End With
               mvarPROJETO.FORMS.Add MyObj, MyObj.NOME
            Case vbext_ct_StdModule
               Set MyObj = New MODULO
               With MyObj
                  .FILENAME = GetNameFromPath(MyComp.FileNames(1))
                  .NOME = MyComp.Name
                  .PATH = GetNameFromPath(MyComp.FileNames(1), 1)
                  '.PagProp.TipoPagina = tpMetodo
                  'Call GetPropPage(.PagProp, MyComp, MyComp.Name)
               End With
               mvarPROJETO.MODULOS.Add MyObj, MyObj.NOME
            Case vbext_ct_ClassModule
               Set MyObj = New CLASSE
               With MyObj
                  .FILENAME = GetNameFromPath(MyComp.FileNames(1))
                  .NOME = MyComp.Name
                  .PATH = GetNameFromPath(MyComp.FileNames(1), 1)
                  '.PagProp.TipoPagina = tpClasse
                  'Call GetPropPage(.PagProp, MyComp, MyComp.Name)
               End With
               mvarPROJETO.CLASSES.Add MyObj, MyObj.NOME
            Case vbext_ct_ResFile
               .RESFILE = GetNameFromPath(MyComp.FileNames(1))
            Case Else
               Dim i%
               i = i
         End Select
      Next
   End With
   Set MyObj = Nothing
End Sub
Private Sub CarregaMembers0(ByRef vComp As VBComponent, ByRef vObj As Object)
   Dim MyMember As New Members
   Dim MyMetodo As New METODO
   Dim MyPropriedade As New Propriedade
   Dim MyVar As New VARIAVEL
   Dim MyEvento As New EVENTO
   Dim MyFuncao As New FUNCAO
   Dim i%, NodX As Node, Mbr As Variant
   Dim lImage$, lComp$, lPai$

'on error resume next
   For Each Mbr In vComp.CodeModule.Members
      Select Case Mbr.Type
         '********************************
         Case vbext_mt_Method '* Métodos *
         '********************************
            Set MyMetodo = New METODO
            With MyMetodo
               .NOME = Mbr.Name
               '.PagProp.TipoPagina = tpMetodo
               'Call GetPropPage(.PagProp, vComp, .NOME)
               vObj.MEMBER.METODOS.Add MyMetodo, .NOME
            End With
            Set MyMetodo = Nothing
         '****************************************
         Case vbext_mt_Property  '* Propriedades *
         '****************************************
            Set MyPropriedade = New Propriedade
            MyPropriedade.NOME = Mbr.Name
            vObj.MEMBER.PROPRIEDADES.Add MyPropriedade, MyPropriedade.NOME
            Set MyPropriedade = Nothing
            
         '************************************
         Case vbext_mt_Variable, vbext_mt_Const '* Variáveis e Constantes*
         '************************************
            Set MyVar = New VARIAVEL
            MyVar.ESCOPO = Mbr.Scope
            MyVar.NOME = Mbr.Name
            MyVar.PassArg = Mbr.Category
            MyVar.TIPO = Mbr.Type
            MyVar.isConst = (vbext_mt_Const = Mbr.Type)
            vObj.MEMBER.VARIAVEIS.Add MyVar, MyVar.NOME
            Set MyVar = Nothing
            
         '*******************************
         Case vbext_mt_Event '* Eventos *
         '*******************************
            Set MyEvento = New EVENTO
            MyEvento.NOME = Mbr.Name
            vObj.MEMBER.EVENTOS.Add MyEvento, MyEvento.NOME
            Set MyEvento = Nothing
          
          '**************************************
          Case Else 'vbext_mt_EventSink '* Event Sink *
          '**************************************
             lImage$ = "CLOSED"
             lPai$ = lComp$ & "SINK"
             If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Outros
                Set NodX = mvarMe.TreProj.Nodes.Add(lComp$, tvwChild, lPai$, "Outros", "CLOSED")
                NodX.ExpandedImage = "OPEN"
                NodX.Sorted = True
             End If
      End Select
   Next
   
   vObj.MEMBER.Carregado = True
End Sub
Private Sub CarregaMembers(strFileName As String, ByRef vObj As Object)
   Dim MyMember As New Members
   Dim MyMetodo As New METODO
   Dim MyPropriedade As New Propriedade
   Dim MyVar As New VARIAVEL
   Dim MyEvento As New EVENTO
   Dim MyFuncao As New FUNCAO
   Dim i%, NodX As Node, Mbr As Variant
   Dim lImage$, lComp$, lPai$
   Dim Pos%, PosAux%, VarDeclarado As Boolean
   Dim LinhaCorrente%

   Dim nArq%, Textline$
   Set vObj.MEMBER = Nothing
   nArq = FreeFile
   Open strFileName For Input As #nArq
   While Not EOF(nArq)
      Line Input #nArq, Textline
      Textline = Trim(Textline)
'      If UCase(Mid(Textline & " ", 1, 6)) = "BEGIN " Then
'         While UCase(Mid(Textline, 1, 3)) <> "END" And Not EOF(nArq)
'            Line Input #nArq, Textline
'            Textline = Trim(Textline)
'         Wend
'         Line Input #nArq, Textline
'      End If
'      Textline = Trim(Textline)
'      Call SomarLinha(Textline, LinhaCorrente%)
'      vObj.LINHAS = LinhaCorrente%
'      Call SetAttribute(Textline, vObj)
      'Call AnalisarLinha_Bas(TextLine, pBAS)
      
      If Textline <> "" Then
         '**********************
         '* Variáveis de Sistema
         '**********************
         If Mid(LCase(Trim(Textline) & " "), 1, 10) = LCase("Attribute ") Then
            Pos = InStr(Textline, ".") + 1
            PosAux = InStr(Textline, "=") - 1
            Pos = IIf(Pos = 1 Or Pos > PosAux, 11, Pos)
            PosAux = IIf(PosAux < 0, 1, PosAux)
            Select Case Mid(Textline, Pos, PosAux - Pos)
               Case "VB_Name": Textline = Textline
               Case "VB_GlobalNameSpace ": Textline = Textline
               Case "VB_Creatable": Textline = Textline
               Case "VB_PredeclaredId:Textline = Textline"
               Case "VB_Exposed": Textline = Textline
               Case "VB_Ext_KEY"
                  Pos = InStr(Textline, "=") + 1
                  PosAux = InStr(Pos, Textline, ",")
'                  Select case EliminarString(trim(mid(textline, pos+1 ,posaux - Pos-1)), """")
'                     case "SavedWithClassBuilder" :Textline = Textline
'                     case "Top_Level"             :Textline = Textline
'                  end select
            End Select
         '**********************
         '* Linhas úteis de Programa.
         '**********************
         ElseIf Mid(Textline, 1, 1) <> "'" Then
            '**********************
            '* Definir Sessão de Varáveis
            '**********************
            If InStr(LCase(Textline), LCase("Function ")) + _
               InStr(LCase(Textline), LCase("Sub ")) + _
               InStr(LCase(Textline), LCase("Property ")) <> 0 And _
               Not VarDeclarado Then
               VarDeclarado = True
            End If
            If VarDeclarado Then
               If InStr(LCase(Textline), LCase("Function ")) <> 0 Then
                  Set MyMetodo = New METODO
                  Pos = InStr(LCase(Textline), LCase("Function "))
                  If InStr(Textline, "Declare ") <> 0 Then
                     VarDeclarado = False
                     With MyMetodo
                        .NOME = Mid(Textline, Pos + 9, InStr(Textline, "Lib") - (Pos + 9))
                        '.PagProp.TipoPagina = tpMetodo
                        'Call GetPropPage(.PagProp, vComp, .NOME)
                        vObj.MEMBER.METODOS.Add MyMetodo, .NOME
                     End With
                     Set MyMetodo = Nothing
                     
                  Else
                  
                     With MyMetodo
                        .NOME = Mid(Textline, Pos + 9, InStr(Textline, "(") - (Pos + 9))
'                     If vObj.LINHAS > 470 And vObj.LINHAS < 520 Then
'                        i = i
'                    End If
                        .StartLine = mvarVbComp.CodeModule.ProcStartLine(.NOME, vbext_ProcKind.vbext_pk_Proc)
                        '.PagProp.TipoPagina = tpMetodo
                        'Call GetPropPage(.PagProp, vComp, .NOME)
                        While Textline <> "End Function" And Not EOF(nArq)
                           Line Input #nArq, Textline
                           Textline = Trim(Textline)
'                           Call SomarLinha(Textline, LinhaCorrente%)
'                           vObj.LINHAS = LinhaCorrente%
'                           Call SetAttribute(Textline, vObj)
                        Wend
                        vObj.MEMBER.METODOS.Add MyMetodo, .NOME
                     End With
                     Set MyMetodo = Nothing
               
                  End If
               ElseIf InStr(LCase(Textline), LCase("Sub ")) <> 0 Then
                  Set MyMetodo = New METODO
                  Pos = InStr(LCase(Textline), LCase("Sub "))
                  With MyMetodo
                     .NOME = Mid(Textline, Pos + 4, InStr(Textline, "(") - (Pos + 4))
'                     If vObj.LINHAS > 400 And vObj.LINHAS < 520 Then
'                        i = i
'                    End If
                     .StartLine = mvarVbComp.CodeModule.ProcStartLine(.NOME, vbext_ProcKind.vbext_pk_Proc)
                     '.PagProp.TipoPagina = tpMetodo
                     'Call GetPropPage(.PagProp, vComp, .NOME)
                     While Textline <> "End Sub" And Not EOF(nArq)
                        Line Input #nArq, Textline
                        Textline = Trim(Textline)
'                        Call SomarLinha(Textline, LinhaCorrente%)
'                        vObj.LINHAS = LinhaCorrente%
'                        Call SetAttribute(Textline, vObj)
                     Wend
                     vObj.MEMBER.METODOS.Add MyMetodo, .NOME
                  End With
                  
                  Set MyMetodo = Nothing
                                    
               ElseIf InStr(LCase(Textline), (" property ")) <> 0 Or Mid(LCase(Textline), 1, 9) = "property " Then
                  Set MyPropriedade = New Propriedade
                  Pos = InStr(LCase(Textline), LCase("Property "))
                  MyPropriedade.NOME = Mid(Textline, Pos + 13, InStr(Textline, "(") - (Pos + 13))
                  On Error Resume Next
                  vObj.MEMBER.PROPRIEDADES.Add MyPropriedade, MyPropriedade.NOME
                  If Err = 457 Then '* This key is already associated with an element of this collection
                     Err = 0
                  Else
                     While Textline <> "End Property" And Not EOF(nArq)
                        Line Input #nArq, Textline
                        Textline = Trim(Textline)
'                        Call SomarLinha(Textline, LinhaCorrente%)
'                        vObj.LINHAS = LinhaCorrente%
                     Wend
                     Set MyPropriedade = Nothing
                  End If
               End If
            Else
               
               '**********************
               '* Sessão de variáveis.
               '**********************
               Pos = InStr(LCase(Textline), LCase("Dim ")) + _
                     InStr(LCase(Textline), LCase("Friend ")) + _
                     InStr(LCase(Textline), LCase("Global ")) + _
                     InStr(LCase(Textline), LCase("Private ")) + _
                     InStr(LCase(Textline), LCase("Public ")) + _
                     InStr(LCase(Textline), LCase("Static "))
               If Pos <> 0 Then
                  If InStr(LCase(Textline), LCase("Option ")) <> 0 Then
                     '* Option Explicity
                  Else
                     Pos = InStr(LCase(Textline), LCase("WithEvents "))

                     Set MyVar = New VARIAVEL
                     MyVar.isConst = (InStr(LCase(Textline), " const ") <> 0)
                     If MyVar.isConst Then
                        '**********************
                        '* Constantes.
                        '**********************
                        Pos = InStr(LCase(Textline), " const ") + 7
                        MyVar.ESCOPO = Mid(Textline, 1, InStr(Textline, " ") - 1)
                        PosAux = InStr(Pos, Textline, " ")
                        PosAux = IIf(PosAux = 0, Len(Textline) + 1, PosAux)
                        PosAux = PosAux - Pos
                        
                        If InArray(Right(Mid(Textline, Pos, PosAux), 1), Array("!", "#", "$", "%", "&", "(")) Then
                           MyVar.NOME = Trim(Mid(Textline, Pos, PosAux - 1))
                        Else
                           MyVar.NOME = Trim(Mid(Textline, Pos, PosAux + 1))
                        End If
                        PosAux = InStr(Textline, " = ") + 3
                        If InStr(PosAux, Textline, Chr(34)) <> 0 Then
                           MyVar.VALOR = Trim(Mid(Textline, PosAux, InStr(PosAux + 1, Textline, Chr(34)) - PosAux))
                        Else
                           If InStr(PosAux, Textline, " ") <> 0 Then
                              MyVar.VALOR = Trim(Mid(Textline, PosAux, InStr(PosAux, Textline, " ") - PosAux))
                           Else
                              MyVar.VALOR = Trim(Mid(Textline, PosAux))
                           End If
                        End If
                        If Mid(MyVar.VALOR, 1, 1) = Chr(34) Then
                           MyVar.VALOR = Mid(MyVar.VALOR, 2)
                        End If
                        If Right(MyVar.VALOR, 1) = Chr(34) Then
                           MyVar.VALOR = Mid(MyVar.VALOR, 1, Len(MyVar.VALOR) - 1)
                        End If
                        
                        
                        'MyVar.PassArg = Mbr.Category
                        Pos = InStr(Pos, Textline, " As ")
                        If Pos <> 0 Then
                           MyVar.TIPO = Trim(Mid(Textline, Pos + 4))
                        Else
                           Select Case Mid(Textline, InStr(Textline, MyVar.NOME) + Len(MyVar.NOME), 1)
                              Case "!":      MyVar.TIPO = "Single"
                              Case "#":      MyVar.TIPO = "Double"
                              Case "$":      MyVar.TIPO = "String"
                              Case "%":      MyVar.TIPO = "Integer"
                              Case "&":      MyVar.TIPO = "Long"
                              Case "(", ")": MyVar.TIPO = "Array"
                              Case Else:     MyVar.TIPO = "Variant"
                           End Select
                        End If
                        
                     ElseIf InStr(LCase(Textline), LCase("Type ")) <> 0 Then
                        MyVar.NOME = Mid(Textline, InStr(LCase(Textline), LCase("Type ")) + 5)
                        MyVar.TIPO = "Type"
                        While Textline <> "End Type" And Not EOF(nArq)
                           Line Input #nArq, Textline
'                           Textline = Trim(Textline)
'                           Call SomarLinha(Textline, LinhaCorrente%)
'                           vObj.LINHAS = LinhaCorrente%
                        Wend
                     ElseIf InStr(LCase(Textline), LCase("Enum ")) <> 0 Then
                        MyVar.NOME = Mid(Textline, InStr(LCase(Textline), LCase("Enum ")) + 5)
                        MyVar.TIPO = "Enum"
                        While Textline <> "End Enum" And Not EOF(nArq)
                           Line Input #nArq, Textline
'                           Textline = Trim(Textline)
'                           Call SomarLinha(Textline, LinhaCorrente%)
'                           vObj.LINHAS = LinhaCorrente%
                        Wend
                        
                     Else
                        If Pos = 0 Then
                           Pos = InStr(Textline, " ") + 1
                           MyVar.ESCOPO = Mid(Textline, 1, InStr(Textline, " ") - 1)
                        Else
                           Pos = Pos + Len("WithEvents ")
                           MyVar.ESCOPO = Mid(Textline, 1, InStr(Textline, " ") - 1)
                        End If
                        
                        PosAux = InStr(Pos, Textline, " ")
                        PosAux = IIf(PosAux = 0, Len(Textline) + 1, PosAux)
                        PosAux = PosAux - Pos
                                                
                        If InArray(Right(Mid(Textline, Pos, PosAux), 1), Array("!", "#", "$", "%", "&", "(")) Then
                           MyVar.NOME = Trim(Mid(Textline, Pos, PosAux - 1))
                        Else
                           MyVar.NOME = Trim(Mid(Textline, Pos, PosAux))
                        End If
                        
                        'MyVar.PassArg = Mbr.Category
                        
                        Pos = InStr(Pos, Textline, " As ")
                        If Pos <> 0 Then
                           MyVar.TIPO = Trim(Mid(Textline, Pos + 4))
                        Else
                           Select Case Mid(Textline, InStr(Textline, MyVar.NOME) + Len(MyVar.NOME), 1)
                              Case "!":      MyVar.TIPO = "Single"
                              Case "#":      MyVar.TIPO = "Double"
                              Case "$":      MyVar.TIPO = "String"
                              Case "%":      MyVar.TIPO = "Integer"
                              Case "&":      MyVar.TIPO = "Long"
                              Case "(", ")": MyVar.TIPO = "Array"
                              Case Else:     MyVar.TIPO = "Variant"
                           End Select
                        End If
                     End If
                     If Trim(MyVar.NOME) <> "" Then
                        vObj.MEMBER.VARIAVEIS.Add MyVar, MyVar.NOME
                     End If
                     Set MyVar = Nothing
                  End If
               ElseIf InStr(LCase(Textline), LCase("Const ")) <> 0 Then
                  
                  Set MyVar = New VARIAVEL
                  MyVar.ESCOPO = Mbr.Scope
                  MyVar.NOME = Mbr.Name
                  MyVar.PassArg = Mbr.Category
                  MyVar.TIPO = Mbr.Type
                  MyVar.isConst = (vbext_mt_Const = Mbr.Type)
                  vObj.MEMBER.VARIAVEIS.Add MyVar, MyVar.NOME
                  Set MyVar = Nothing
               End If
               
            End If
         End If
      End If
   Wend
   Close #nArq
   vObj.MEMBER.Carregado = True
End Sub

Public Sub DefineMembro(Optional NodeX As Variant)
   Dim lKey$, Pos%
   Dim eClasse As Boolean, eFilhoClasse As Boolean
   Dim Nox As Node
   On Error GoTo Fim
   If IsMissing(NodeX) Then
      Set Nox = mvarMe.TreProj.SelectedItem
   Else
     Set Nox = NodeX
   End If
   If Not mvarVbMember Is Nothing Then
      If mvarVbMember.Name = Nox.Text Then GoTo Fim
   End If
   If Nox.Image = Icon.VBPImage Then GoTo Fim
   eClasse = isClass(Nox)
   If eClasse Then
      If Nox.Key <> "VBClass" Then
         If Nox.Image = Icon.CLSImage Then
            lKey$ = Nox.Key
'            Pos = InStr(mvarMe.TreProj.SelectedItem.Key, Chr(255))
'            Pos = IIf(Pos <= 2, Len(mvarMe.TreProj.SelectedItem.Key) + 1, Pos)
'            lKey$ = Mid(mvarMe.TreProj.SelectedItem.Key, 1, Pos - 1)
''            Set mvarVbComp = mvarVbProj.VBComponents.Item(lKey$)
''            Set mvarVbCode = mvarVbComp.CodeModule
         Else
            lKey$ = Trim(Nox.Text)
         End If
''         Call MontaSubItens(NoX.Key)
''         Set mvarVbMember = mvarVbCode.Members.Item(lKey$
''         Call Rel_Tab_Tree
''         Call MontaListItens
      End If
''      GoTo Fim
   Else
      eFilhoClasse = isClassChild(Nox)
      If eFilhoClasse Then
         lKey$ = GetParentClass(Nox).Key
      Else
         lKey$ = GetParent(Nox, 2).Key
      End If
   End If
   Pos = InStr(lKey$, Chr(255))
   Pos = IIf(Pos <= 2, Len(lKey$) + 1, Pos)
   lKey$ = Trim(Mid(lKey$, 1, Pos - 1))
   
   If Not mvarVbComp Is Nothing And mvarMe.TabComp.Tab <> 5 Then
      '* Se o componente já foi carregado e é o mesmo item da árvore
      '* então não precisa caregá-lo de novo
      If Nox.Children <> 0 And lKey$ = mvarVbComp.Name Then GoTo Fim
   End If
   If Trim(lKey$) = "" Then GoTo Fim
   If UCase(Right(lKey$, 3)) = "RES" Then GoTo Fim
   
   
'   For Pos = 1 To mvarVbProj.VBComponents.Count
'      Debug.Print mvarVbProj.VBComponents(Pos).Name
'   Next
   Set mvarVbComp = mvarVbProj.VBComponents.Item(lKey$)
   Set mvarVbCode = mvarVbComp.CodeModule
'   If Not eFilhoClasse Then
'      lKey$ = GetParent(NoX, 4).Text
'   Else
'      lKey$ = NoX.Key
'   End If
   Pos% = GetNivel(Nox)
   If Pos% < 3 Or (eFilhoClasse And eClasse) Then
   
   Else
      If lKey$ <> Trim(Nox.Text) Then
         If Pos% = 3 Then
            lKey$ = GetParent(Nox, 4)
         Else
            lKey$ = Nox.Text
         End If
         Pos = InStr(lKey$, Chr(255))
         Pos = IIf(Pos <= 2, Len(lKey$) + 1, Pos)
         lKey$ = Trim(Mid(lKey$, 1, Pos - 1))
         Set mvarVbMember = mvarVbCode.Members.Item(lKey$)
      End If
   End If
Fim:
   Call ShowError
End Sub
Public Sub Refresh_Tela()
   Dim n As Variant, ItemExcluido As New Collection
   If Sys.ExibeSubPasta Then
      Call mvarMe_TreProjNodeClick(mvarMe.TreProj.SelectedItem)
   Else
      For Each n In mvarMe.TreProj.Nodes
         If n.Image = Icon.CloseImage Then
            If GetNivel(mvarMe.TreProj.Nodes(n.Key)) > 1 Then
               ItemExcluido.Add n.Key, n.Key
            End If
         End If
      Next
      For Each n In ItemExcluido
         mvarMe.TreProj.Nodes.Remove mvarMe.TreProj.Nodes(n).Index
      Next
   End If
   mvarMe.TabComp.TabVisible(5) = Sys.ExibeSubPasta
   Call PintarFundo(mvarMe.ImgFundo, Sys.FundoTela, mvarMe)
   mvarMe.Refresh
End Sub
Private Sub mvarMe_Active()
   Dim n As Variant, NmComp$
   Dim i%, Tot&
   Screen.MousePointer = vbHourglass
   mvarMe.Lbl = " Projeto : " & mvarVbProj.Description
   
   Tot = mvarMe.TreProj.Nodes.Count
   If mvarMe.PrimeiraVez And Sys.LoadIni Then
      mvarMe.PrimeiraVez = False
      'mvarMe.TreProj.Visible = False
      For i = 1 To Tot
         NmComp$ = mvarMe.TreProj.Nodes(i).Key
         Call MontaSubItens(NmComp$)
         If Not AtuFlood(Int(i / Tot * 100)) Then
            Exit For
         End If
      Next
      Call FimFlood
   End If
   
   mvarMe.PrimeiraVez = False
   Call Refresh_Tela
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_Load()
   Screen.MousePointer = vbHourglass
   Set SysMdi = FrmAddIn
   
   
   '* Variáveis do Registro do Sistema
   Call GetConfig
   
  '*********
   '* Montar e configurar 'Lay-Out' de Componentes
   With mvarMe
      .PrimeiraVez = True
      .MnuMouse_Main.Visible = False
      .TabComp.Tab = 0
      .TreProj.ImageList = .ImgList
      .LstItens.SmallIcons = .ImgList
      .LstItens.Icons = .ImgList
      
      .Toolbar.ImageList = .ImgListToolbar
   
   
      Call MontaArvoreProjeto
      Call ConfigLstItens(0)
      '************
      '* Canfigurar Tamanho
      .Spl.SplitPercent = 35
      Call .Spl.SizeControls(.Spl.Left, .TreProj, .TabComp, .Lbl)
      .LstItens.Move .TabComp.Left + 60, .LstItens.Top, .TabComp.Width - 120, .TabComp.Height - 480
      Call ConfigForm(mvarMe, , Sys.FundoTela, False)
   
      .WindowState = vbMaximized
            
   End With
End Sub
Private Sub mvarMe_LstItensClick()
   Dim lKey$
   If mvarMe.LstItens.SelectedItem Is Nothing Then Exit Sub
   lKey$ = Trim(mvarMe.LstItens.SelectedItem.Text)
   lKey$ = EliminarString(lKey$, "()")
'   If GetNivel(mvarMe.TreProj.SelectedItem) < 4 Then
     Set mvarVbMember = mvarVbCode.Members.Item(lKey$)
'   Else
'      Set mvarVbMember = mvarVbCode.Members.Item(Trim(mvarMe.TreProj.SelectedItem.Text))
'   End If

End Sub
Private Sub mvarMe_LstItensDblClick()
   Call mvarMe_MnuMouseClick("00", IndMnuMouse.mm_PROPRIEDADES)
End Sub

Private Sub mvarMe_LstItensMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim i%, lKey$
   If mvarMe.LstItens.SelectedItem Is Nothing Then Exit Sub
   lKey$ = Trim(mvarMe.LstItens.SelectedItem.Text)
   lKey$ = EliminarString(lKey$, "()")
   Set mvarVbMember = mvarVbCode.Members.Item(lKey$)
   Select Case Button
      Case 1
      Case 2
         '* Objeto
         On Error Resume Next
         For i = mvarMe.MnuMouse00.LBound To mvarMe.MnuMouse00.UBound
            mvarMe.MnuMouse00(i).Visible = True
         Next
         If mvarMe.TabComp.Tab = 4 Then
            mvarMe.MnuMouse00(IndMnuMouse.mm_ADICIONAR).Visible = False
         Else
            mvarMe.MnuMouse00(IndMnuMouse.mm_ADICIONAR).Caption = "Adicionar " & mvarMe.TabComp.TabCaption(mvarMe.TabComp.Tab)
            mvarMe.MnuMouse00(IndMnuMouse.mm_ADICIONAR_OBJ).Visible = False
         End If
         If Not (isClass(mvarMe.TreProj.SelectedItem) Or isClassChild(mvarMe.TreProj.SelectedItem)) Then
            mvarMe.MnuMouse00(IndMnuMouse.mm_PROPRIEDADES).Visible = False
            mvarMe.MnuMouse00(IndMnuMouse.mm_DEFINIR_COMO).Visible = False
            mvarMe.MnuMouse00(IndMnuMouse.mm_SEP4).Visible = False
         End If
         mvarMe.MnuMouse00(IndMnuMouse.mm_RECARREGAR_OBJ).Visible = (GetNivel(mvarMe.TreProj.SelectedItem) = 2)
         Call mvarMe.PopupMenu(mvarMe.MnuMouse(0))
   End Select
End Sub

Private Sub mvarMe_MenuClick(Menu As String, Index As Integer)
   Select Case Menu
      Case "0000"
         Select Case Index
            Case 0 '* Abrir Projeto
 
            Case 1 '* Separador
            Case 2
               FrmConfig.Show vbModal
              
            Case 3 '* Separador
            Case 4
               Call Hide
         End Select
   End Select
End Sub
Private Sub mvarMe_MnuMouseClick(Menu As String, Index As Integer)
   On Error GoTo Fim
   Select Case Menu
      Case "00"
         Select Case Index
            Case IndMnuMouse.mm_ADICIONAR
            
            Case IndMnuMouse.mm_ADICIONAR_OBJ
            Case IndMnuMouse.mm_SALVAR
               Select Case mvarMe.ActiveControl.Name
                  Case mvarMe.TxtCode.Name
                     Call TrocarLinhasCodigo(False)
                     Call mvarVbComp.SaveAs(mvarVbComp.FileNames(1))
                  Case mvarMe.TreProj.Name
                  
                  Case mvarMe.LstItens.Name
               End Select
            Case IndMnuMouse.mm_Salvar_COMO
               Dim NovoNome$
               Select Case mvarMe.ActiveControl.Name
                  Case mvarMe.TxtCode.Name
                     Call TrocarLinhasCodigo(False)
                     NovoNome = InputBox("Enter the new folder name")
                     If Trim(NovoNome) <> "" Then Call mvarVbComp.SaveAs(NovoNome)
                  Case mvarMe.TreProj.Name
                  
                  Case mvarMe.LstItens.Name
               End Select
            
            Case IndMnuMouse.mm_SEP1          '* 2 - Separador
            Case IndMnuMouse.mm_RECORTAR, IndMnuMouse.mm_COPIAR  '* 3, 4 - Recortar / Copiar
               With mvarAreaTemp
                  .NomeObj = ""
                  .TipoObj = ""
                  .TextoObj = ""
                  .TipoOper = IIf(Index = 3, Recortar, Copiar)
               End With
               mvarMe.MnuMouse00(IndMnuMouse.mm_COLAR).Enabled = True
            Case IndMnuMouse.mm_COLAR         '* 5 - Colar
            Case IndMnuMouse.mm_SEP2          '* 6 - Separador
            Case IndMnuMouse.mm_EXCLUIR       '* 7 - Excluir
            Case IndMnuMouse.mm_RENOMEAR      '* 8 - Renomear
               Select Case mvarMe.ActiveControl.Name
                  Case mvarMe.TreProj.Name
                     If Not InArray(mvarMe.TreProj.SelectedItem.Image, Array(Icon.VBPImage, Icon.CloseImage)) Then
                        mvarMe.TreProj.StartLabelEdit
                     End If
                  Case mvarMe.LstItens.Name
                     mvarMe.LstItens.StartLabelEdit
               End Select
            Case IndMnuMouse.mm_SEP3           '* 9 - Separador
            Case IndMnuMouse.mm_PROPRIEDADES   '* 10 - Propriedades
               Dim MyProp As New PROPPAGE
               Dim Img$, StrAux$
               'StrAux = Me.PROJETO(mvarPROJETO.NOME).CLASSES(mvarVbComp.Name).NOME
               StrAux = mvarVbComp.Name
               With MyProp
                  
                  Set .PROJETO = mvarPROJETO
                 
                  Select Case mvarMe.ActiveControl.Name
                     Case "LstItens", "TabComp"
                        Img$ = mvarMe.LstItens.SelectedItem.Icon
                     Case "TreProj"
                        Img$ = mvarMe.TreProj.SelectedItem.Image
                  End Select
                  Select Case Img$
                     Case Icon.CLSImage:                  .TipoPagina = tpClasse
                     Case Icon.CollImage:                 .TipoPagina = tpColecao
                     Case Icon.EvtImage:                  .TipoPagina = tpEvento
                     Case Icon.FRMImage:                  .TipoPagina = tpForm
                     Case Icon.MetImage:                  .TipoPagina = tpMetodo
                     Case Icon.BASImage:                  .TipoPagina = tpModulo
                     Case Icon.PropImage:                 .TipoPagina = tpPropriedade
                     Case Icon.VarImage, Icon.ConstImage: .TipoPagina = tpPropriedade
                     Case Else: .TipoPagina = tpNull
                  End Select
                  On Error Resume Next
                  'On Error GoTo 0
                  Dim LinhaInicial%, PosRetorno%, PosIniArg%, PosAux%
                  Dim StrFunc$, Palavras As Collection
                  Dim MyArg As ADDARG, n
                  Dim Parenteses%, i%, Pos%, isFunc As Boolean
                  '* Definir Retorno e Argumentos
                  If .TipoPagina = tpMetodo Then
                     LinhaInicial = mvarVbComp.CodeModule.ProcStartLine(mvarVbMember.Name, vbext_ProcKind.vbext_pk_Proc)
                     
                     If Err = 35 Or LinhaInicial = 0 Then
                        LinhaInicial = 1
                     Else
                        While InStr(mvarVbCode.Lines(LinhaInicial, 1), "Function " & mvarVbMember.Name) + InStr(mvarVbCode.Lines(LinhaInicial, 1), "Sub " & mvarVbMember.Name) = 0
                           LinhaInicial = LinhaInicial + 1
                        Wend
                     End If
                  Else
                     Exit Sub
                  End If
                  StrFunc$ = mvarVbCode.Lines(LinhaInicial, 1)
                  StrAux$ = StrFunc$
                  isFunc = (InStr(UCase(StrFunc), "FUNCTION") <> 0)
                  If isFunc Then .Retorno = "Variant"
                  
'******
                  PosAux = InStr(UCase(StrAux$), ")")
                  While PosAux <> 0
                     PosRetorno = PosAux
                     PosAux = InStr(UCase(StrAux$), ")")
                     StrAux$ = Trim(Mid(StrAux$, PosAux + 1))
                  Wend
                  PosIniArg% = InStr(UCase(StrFunc$), "(")
                  PosRetorno = IIf(PosRetorno = PosIniArg% + 1, 0, PosRetorno)
                  If PosRetorno > 1 Then StrAux$ = Mid(StrFunc$, PosIniArg% + 1, PosRetorno - 1)
                  Set Palavras = GetPalavras(StrAux$)
                  Parenteses = 1
                  StrAux$ = ""
                  For i = 1 To Palavras.Count
                     If Parenteses <= 0 Then '* Retorno
                        If PosRetorno = True Then
                           .Retorno = Palavras(i)
                        End If
                        If UCase(Palavras(i)) = "AS" Then PosRetorno = True
                     Else
                        If Mid(Palavras(i), 1, 1) = "(" Then
                           If Len(Palavras(i)) > 1 Then Palavras(i) = Mid(Palavras(i), 2)
                           Parenteses = Parenteses + 1
                        End If
                        If Right(Palavras(i), 1) = ")" Then
                           If Len(Palavras(i)) > 1 Then
                              Palavras.Add Mid(Palavras(i), 1, Len(Palavras(i)) - 1), , , i
                              Palavras.Remove i
                           End If
                           Parenteses = Parenteses - 1
                           'If Parenteses = 0 Then
                           '   StrAux$ = StrAux$ & IIf(Trim(StrAux$) = "", "", " ") & Palavras(i)
                           'End If
                        End If
                     
                        If Parenteses >= 0 Then
                           If Palavras(i) = "," Or Parenteses = 0 Then
                              If Trim(StrAux$) = "" Then StrAux$ = Palavras(i)
                              If InStr(StrAux$, "ByVal ") = 0 Then
                                 If InStr(StrAux$, "ByRef ") = 0 Then
                                    Pos = InStr(StrAux$, "Optional ")
                                    If Pos = 0 Then
                                       StrAux$ = "ByRef " & StrAux$
                                    Else
                                       StrAux$ = Mid(StrAux$, 1, Pos + 8) & "ByRef" & Mid(StrAux$, Pos + 8)
                                    End If
                                 End If
                              End If
                              .Argumentos.Add StrAux$
                              StrAux$ = ""
                           Else
                              StrAux$ = StrAux$ & IIf(Trim(StrAux$) = "", "", " ") & Palavras(i)
                           End If
                        End If
                     End If
                  Next

'******
'                  StrAux$ = Mid(StrFunc$, PosRetorno + 1)
'                  If StrAux$ <> "" Then
'                     PosRetorno = InStr(UCase(StrAux$), "AS ")
'                     If PosRetorno <> 0 Then
'                        .Retorno = Trim(Mid(StrAux$, PosRetorno + 3))
'                     End If
'                  End If
                  .NOME = mvarVbMember.Name
                  .Descricao = mvarVbMember.Description
                  .HelpID = mvarVbMember.HelpContextID
                  .Show
                  mvarVbMember.Description = .Descricao
                  mvarVbMember.HelpContextID = .HelpID
                  If .isSaved Then
                     Call mvarVbComp.SaveAs(mvarVbComp.FileNames(1))
                  End If
               End With
            Case IndMnuMouse.mm_DEFINIR_COMO '* 11 - Definir como Coleção ou Classe
            Case IndMnuMouse.mm_SEP4         '* 12 - Separador
            Case IndMnuMouse.mm_IR_PROJETO   '* 13 - Ir para o Projeto
               Dim LinIni&
               Call TrocarLinhasCodigo
               mvarVbComp.CodeModule.CodePane.Window.WindowState = 2 '* Maximized
               LinIni = Me.Get_LinIni_Codigo(mvarVbMember)
               mvarVbComp.CodeModule.CodePane.TopLine = LinIni
               Call mvarVbComp.CodeModule.CodePane.SetSelection(LinIni, 1, LinIni, 1)
               mvarVbComp.CodeModule.CodePane.Show
            Case IndMnuMouse.mm_RECARREGAR_OBJ '* 14 - Recarregar Objeto
                GetMember(mvarMe.TreProj.SelectedItem.Key).MEMBER.Carregado = False
                Call MontaSubItens(mvarMe.TreProj.SelectedItem.Key)
            End Select
         
      Case "0000" '* Adicionar Objeto
         Select Case Index
            Case 0 '* Classe de Banco de Dados
               If Not DB.Conectado Then Call ConectarBanco
               If DB.Conectado Then FrmMontaCls.Show vbModal
            Case 1 '* Clase
            Case 2 '* Método
            Case 3 '* Propriedade, Variável
         End Select
   End Select
Fim:
   Call ShowError
End Sub

Private Sub mvarMe_Resize()
   Dim HHStatus&
   On Error Resume Next
   With mvarMe
      If .Width >= 3000 And .Height >= 3000 Then
      If .Width - (.TabComp.Left + 180) < 0 Then Exit Sub
         HHStatus& = 960
         .TreProj.Left = 60
         .TabComp.Move .TabComp.Left, .TabComp.Top, .Width - (.TabComp.Left + 180), .Height - .TabComp.Top - HHStatus
         .LstItens.Move .TabComp.Left + .TreProj.Left, .TabComp.Top + 400, .TabComp.Width - 120, .TabComp.Height - 480
         .TxtCode.Move 60, 400, .TabComp.Width - 120, .TabComp.Height - 480
         .TreProj.Move .TreProj.Left, .TreProj.Top, .TabComp.Left - (.TreProj.Left + .Spl.SplitterWidth), .Height - .TreProj.Top - HHStatus
         .Lbl.Move .TreProj.Left, .Lbl.Top, .TreProj.Width
         .Spl.Move .Spl.Left, .Spl.Top, .Spl.Width, .TreProj.Height
      End If
      Call PintarFundo(mvarMe.ImgFundo, Sys.FundoTela)
   End With
End Sub

Private Sub mvarMe_TabCompClick(PreviousTab As Integer)
   Screen.MousePointer = vbHourglass
   Call ConfigLstItens(mvarMe.TabComp.Tab)
   Call MontaListItens(mvarMe.TabComp.Tab)
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_ToolbarButtonClick(ByVal Button As ComctlLib.Button)
   Select Case Button.Key
      Case "SAIR": Call Hide
      Case "MONTAR CLASSE"
         If Not DB.Conectado Then Call ConectarBanco
         If DB.Conectado Then FrmMontaCls.Show vbModal
   End Select
End Sub
Public Sub AnalisarLinha(StrLinha$, PosTexto&)
   Dim Palavras As New Collection
   Dim n, Pos&, PosLin&, PosIni&
   Dim AchouAspas As Boolean
   Dim AchouAbreAspas As Boolean, AchouFechaAspas As Boolean
   Dim StrEspecial As Boolean
   Dim StrEspaco$
   
   Set Palavras = New Collection
   Set Palavras = GetPalavras(StrLinha$)
   StrLinha$ = StrLinha$ & " "
   PosLin& = 1
   For Each n In Palavras
      StrEspecial = InArray(UCase(Palavras(1)), Array("'", "#", "DEBUG"))
      StrEspaco = IIf(StrEspecial, "", " ")
      PosIni& = InStr(StrLinha$, Palavras(1) & StrEspaco)
         
      StrEspecial = InArray(UCase(CStr(n)), Array("'", "#", "DEBUG"))
      StrEspaco = IIf(StrEspecial, "", " ")
      If InStr(PosLin&, StrLinha$, CStr(n) & StrEspaco) <> 0 Or CStr(n) = "#" Then
         PosLin& = InStr(PosLin&, StrLinha$, CStr(n) & StrEspaco)
         'PosLin& = IIf(PosLin& = 0, 1, PosLin&)
         Pos& = PosTexto& + PosLin& - PosIni&
         
         AchouAbreAspas = (Mid(CStr(n), 1, 1) = """")
         AchouFechaAspas = (Right(CStr(n), 1) = """")
         If AchouAbreAspas And Not AchouFechaAspas Then AchouAspas = True
         If AchouFechaAspas Then AchouAspas = False
         If Not AchouAspas Then
            If InStr(CStr(n), "'") <> 0 Then
               Call MudarCor(mvarMe.TxtCode, Pos&, Len(StrLinha$) - (PosLin& - 1), &H8000&)
               Exit For
            Else
               If PalavraReservada(CStr(n)) Then
                  Call MudarCor(mvarMe.TxtCode, Pos&, Len(CStr(n)), &H800000)
               End If
            End If
         End If
      End If
   Next
End Sub
Public Sub ConectarBanco()
   Dim Arq$
   With DB
      FrmOpBanco.Show vbModal
      If Not Sys.dbODBC Then
         mvarMe.CmDialog.InitDir = "C:\DSR\"
         Arq$ = ProcurarArquivo(mvarMe.CmDialog, "Abrir Banco de Dados Access", , "Microsoft Access MDBs (*.mdb)|*.mdb")
         If Arq$ <> "" Then
            .dbODBC = False
            Sys.dbDrive = mvarMe.CmDialog.Tag
            Sys.dbName = Arq$
            Call .SrvConecta(Sys.dbDrive, Sys.dbName, "", "", "", "")
         End If
      Else
         .dbODBC = True
         frmODBCLog.Show vbModal
      End If
   End With
End Sub
Public Sub ConfigLstItens(TabAtual As Integer)
   Dim i%
   With mvarMe.LstItens
      .Visible = False
      .ColumnHeaders.Clear
      For i = 1 To 4
         .ColumnHeaders.Add i
         If i = 1 Then
            .ColumnHeaders(i).Width = 2500
         Else
            .ColumnHeaders(i).Width = 1250
         End If
      Next
      .ColumnHeaders(1).Text = "Nome"
      Select Case TabAtual
         Case 0 '* Propriedades
            .ColumnHeaders(2).Text = "Tipo"
            .ColumnHeaders(3).Text = "Declaração"
            .ColumnHeaders.Remove .ColumnHeaders.Count
         Case 1 '* Métodos
            .ColumnHeaders(2).Text = "Tipo"
            .ColumnHeaders(3).Text = "Argumento"
            .ColumnHeaders.Remove .ColumnHeaders.Count
         Case 2 '* Eventos
            .ColumnHeaders(2).Text = "Argumento"
            .ColumnHeaders.Remove .ColumnHeaders.Count
            .ColumnHeaders.Remove .ColumnHeaders.Count
         Case 3 '* Variáveis e Constantes
            .ColumnHeaders(2).Text = "Tipo"
            .ColumnHeaders(3).Text = "Declaração"
            .ColumnHeaders.Remove .ColumnHeaders.Count
         Case 4 '* Tudo
            .ColumnHeaders(2).Text = "Tipo"
            .ColumnHeaders(3).Text = "Argumento"
            .ColumnHeaders(4).Text = "Declaração"
         Case 5 '* Código
            .Visible = False
            mvarMe.Refresh
      End Select
      If TabAtual <> 5 Then .Visible = True
   End With
   
End Sub
Public Sub Rel_Tab_Tree()
   Dim NodX As Node, NodObj As Node, NodFilho As Node
   Dim i%, j%, TabAntes%
   Dim TabFalse As New Collection, n
   Dim Tab_Visible()
'   Set NodX = mvarMe.TreProj.SelectedItem
'   If GetNivel(NodX) >= 2 Then
'      If isClass(NodX) Then
'         Set NodObj = NodX
'      Else
'         i% = GetNivel(NodX)
'         i% = IIf(i% <= 3, 4, i%)
'         Set NodObj = GetParent(NodX, i% - 2)
'      End If
'      If NodObj.Children > 0 Then
'         Set NodFilho = NodObj.Child
'      Else
'         GoTo Fim
'      End If
      TabAntes% = mvarMe.TabComp.Tab

   '* Definir Chave do Component
   If mvarVbComp Is Nothing Then GoTo Fim
'   lComp$ = mvarVbComp.Name
   Dim MyObj As Object

   Set MyObj = GetMember(mvarVbComp.Name)
   With mvarMe.TabComp
      .TabVisible(0) = (MyObj.MEMBER.PROPRIEDADES.Count > 0)
      .TabVisible(1) = (MyObj.MEMBER.METODOS.Count > 0)
      .TabVisible(2) = (MyObj.MEMBER.EVENTOS.Count > 0)
      .TabVisible(3) = (MyObj.MEMBER.VARIAVEIS.Count > 0)
      .TabVisible(4) = (.TabVisible(0) Or .TabVisible(1) Or .TabVisible(2) Or .TabVisible(3))
   End With
'*********************************************************************
      
      
'      With mvarMe.TabComp
'         ReDim Tab_Visible(mvarMe.TabComp.Tabs - 1)
'         For i = LBound(Tab_Visible) To UBound(Tab_Visible)
'            .TabVisible(i) = True
'            Tab_Visible(i) = False
'         Next
'         For i = 1 To NodObj.Children
'            Select Case Trim(NodFilho.Text)
'               Case "Properties", "Propriedades": j = 0
'               Case "Metodies", "Métodos": j = 1
'               Case "Events", "Eventos": j = 2
'               Case "Variables", "Variáveis": j = 3
'            End Select
'            Tab_Visible(j) = True
'            Set NodFilho = NodFilho.Next
'         Next
'         '* Habilita Tab
'         Tab_Visible(4) = True
'         Tab_Visible(5) = True
'         For i = LBound(Tab_Visible) To UBound(Tab_Visible)
'            .TabVisible(i) = Tab_Visible(i)
'         Next
'      End With
'   End If

   If mvarMe.TabComp.TabVisible(TabAntes%) Then
      mvarMe.TabComp.Tab = TabAntes
   Else
      For i = 0 To mvarMe.TabComp.Tabs - 1
         If mvarMe.TabComp.TabVisible(i) Then
            mvarMe.TabComp.Tab = i
            Exit For
         End If
      Next
   End If
Fim:
End Sub
Public Sub TrocarLinhasCodigo(Optional PerguntaAntes = True)
   Dim lKey$
   Dim QtdLin&, LinIni&
   Dim Lin&, Tam&, Ini&, StrLinha$, i&
   Dim StrAux$
   Dim Dif_Lin%
   'On Error Resume Next
   If mvarMe.ActiveControl.Name = mvarMe.TabComp.Name Then
      If mvarMe.TabComp.Tab <> 5 Then Exit Sub
   Else
      If mvarMe.ActiveControl.Name <> mvarMe.TxtCode.Name Then Exit Sub
   End If
   lKey$ = GetParent(mvarMe.TreProj.SelectedItem, 2).Key
   Lin = 0
   Tam = 1
   Ini = 1

   LinIni = Me.Get_LinIni_Codigo(mvarVbMember)   'mvarVbMember.CodeLocation
   QtdLin = Me.Get_QtdLin_Codigo(mvarVbMember, LinIni) 'mvarVbCode.ProcCountLines(mvarVbMember.Name, vbext_ProcKind.vbext_pk_Proc)
   
   '* Se Houver necessidade de Troca
   If Trim(mvarMe.TxtCode.Text) <> Trim(mvarVbCode.Lines(LinIni, QtdLin)) Then
      If PerguntaAntes Then
         If Not ExibirPergunta("Salvar código do Projeto ?", LoadMsg(1)) = vbYes Then
            Exit Sub
         End If
      End If
      '* Percorre todos os caracteres do texto
      For i = 1 To Len(Trim(mvarMe.TxtCode.Text))
         '* Verificar linha a linha
         If Lin <> mvarMe.TxtCode.GetLineFromChar(i) Then
            StrAux = Mid(mvarMe.TxtCode.Text, Ini, Tam)
            If InStr(StrAux, Chr(13) & Chr(10)) = 0 Then
               StrLinha = Mid(mvarMe.TxtCode.Text, Ini, Tam)
            Else
               StrLinha = Mid(mvarMe.TxtCode.Text, Ini, Tam - 2)
               '* Se linha for diferente
               If StrLinha <> mvarVbCode.Lines(LinIni + Lin, 1) Then
                  'Call AnalisarLinha(StrLinha, mvarMe.TxtCode.Find(Trim(StrLinha$)))
                  Call mvarVbCode.ReplaceLine(LinIni + Lin, StrLinha)
                  '* Se o resto do texto for igual sai da verificação de troca.
                  If QtdLin > (Lin + 1) Then '* Se não é Fim de função
                     If Mid(mvarMe.TxtCode.Text, Ini + Len(StrLinha) + 2, Len(Trim(mvarMe.TxtCode.Text))) = mvarVbCode.Lines(LinIni + Lin + 1, QtdLin - (Lin + 1)) Then
                        Exit For
                     End If
                  End If
               End If
               Lin = Lin + 1 'mvarMe.TxtCode.GetLineFromChar(i)
               Tam = 0
               Ini = Ini + Len(StrLinha) + 2
                
            End If
          End If

         Tam = Tam + 1
      Next
   End If
End Sub
Public Sub MudarCor(Control As Object, Pos&, Tam&, Cor&)
'&H00008000& verde &H0000C000&
'&H00800000& azul
'&H000000C0& vermelho
   Dim PosAnt&
   On Error GoTo Fim
   With Control
      PosAnt = .SelStart
      .SelStart = Pos
      .SelLength = Tam
      .SelColor = Cor&
      .SelStart = PosAnt
      .SelLength = 0
   End With
Fim:
   Control.SelStart = PosAnt
End Sub
Public Sub ExibirCodigo()
   Dim lKey$, lPai$
   Dim QtdLin&, LinIni&, PosAnt&, PosTexto&, i%
   Dim StrLinha$, eClasse As Boolean
   lPai$ = lKey$ & "CODIGO"
   mvarMe.LstItens.Visible = False
   On Error Resume Next
'Call mvarMe.TxtCode.LoadFile(Me.VbComp.FileNames(1), 1)
'Exit Sub
   eClasse = isClass(mvarMe.TreProj.SelectedItem)
   If eClasse Then
      lKey$ = Trim(mvarMe.TreProj.SelectedItem.Text)
      If mvarMe.TreProj.SelectedItem.Image = Icon.CloseImage Then
         lKey$ = Trim(mvarMe.TreProj.SelectedItem.Child.Text)
      End If
      Exit Sub
   Else
      If GetNivel(mvarMe.TreProj.SelectedItem) < 4 Then
         If LCase(Trim(mvarMe.TreProj.SelectedItem.Text)) = "variáveis" Then
            lKey$ = Mid(mvarMe.TreProj.SelectedItem.Key, 1, Len(mvarMe.TreProj.SelectedItem.Key) - Len("VARIAVEL"))
         Else
            lKey$ = Trim(GetParent(mvarMe.TreProj.SelectedItem, 4).Text)
            If lKey$ <> mvarMe.TreProj.SelectedItem.Text Then
               Set mvarVbMember = mvarVbCode.Members.Item(lKey$)
            End If
         End If
      Else
         lKey$ = Trim(mvarMe.TreProj.SelectedItem.Text)
         If mvarMe.TreProj.SelectedItem.Image = Icon.CloseImage Then
            lKey$ = Trim(mvarMe.TreProj.SelectedItem.Child.Text)
         End If
         Set mvarVbMember = mvarVbCode.Members.Item(lKey$)
      End If
   End If
   If Err = 0 Then
      
      mvarMe.TxtCode.Visible = False
      PosAnt = mvarMe.TxtCode.SelStart
      
      mvarMe.TxtCode.SelStart = 0
      mvarMe.TxtCode.SelLength = Len(mvarMe.TxtCode.Text)
      mvarMe.TxtCode.SelColor = vbBlack
      If eClasse Then
         LinIni = 1
         QtdLin = mvarVbCode.CodePane.CodeModule.CountOfLines
      Else
         If LCase(Trim(mvarMe.TreProj.SelectedItem.Text)) = "variáveis" Then
            LinIni = 1
            QtdLin = mvarVbCode.CountOfDeclarationLines
         Else
            LinIni = Get_LinIni_Codigo(mvarVbMember)
            QtdLin = Get_QtdLin_Codigo(mvarVbMember, LinIni)
         End If
      End If

      mvarMe.TxtCode.Text = mvarVbCode.Lines(LinIni, QtdLin)
      
   If True Then
      Call ColorizeVbWords(mvarMe.TxtCode)
   Else
      PosTexto& = -1
      For i = 0 To QtdLin - 1
         StrLinha$ = mvarVbCode.Lines(LinIni + i, 1)
         If Trim(StrLinha$) <> "" Then
            PosTexto& = mvarMe.TxtCode.Find(Trim(StrLinha$) + Chr(13), PosTexto& + 1)
            If PosTexto& >= 0 And Trim(StrLinha$) <> "" Then
               Call AnalisarLinha(StrLinha$, PosTexto&)
            End If
         End If
      Next
   End If
      mvarMe.TxtCode.SelStart = PosAnt
      mvarMe.TxtCode.SelLength = 0
      mvarMe.TxtCode.Visible = True
   Else
      If eClasse Then
         LinIni = 1
         QtdLin = mvarVbCode.CodePane.CodeModule.CountOfLines
         mvarMe.TxtCode.Text = mvarVbCode.Lines(LinIni, QtdLin)
      End If
   End If
   If Err <> 0 Then mvarMe.TxtCode = ""
   mvarMe.TxtCode.SelColor = vbBlack
   mvarMe.LstItens.Visible = False
   cCodigoAntes = mvarMe.TxtCode.Text
   cLinhaAntes = 0 'mvarMe.TxtCode.GetLineFromChar(mvarMe.TxtCode.SelStart)
End Sub
Public Function isClass(pNo As Node) As Boolean
   Dim i%, Nv%, eClasse As Boolean
   If InStr(LCase(pNo.Key), ".vbp") <> 0 Then Exit Function
   eClasse = ("VBClass" = GetParent(pNo, 1).Key)
   If eClasse Then
      Nv% = GetNivel(pNo)
      For i% = 1 To Nv%
         If Right(GetParent(pNo, i%).Key, 6) = "METODO" Then
            eClasse = False
            Exit For
         End If
         If Right(GetParent(pNo, i%).Key, 8) = "VARIAVEL" Then
            eClasse = False
            Exit For
         End If
      Next
   End If
   isClass = eClasse
End Function
Public Function GetParentClass(pNo As Node) As Node
   Dim i%, Nv%
   Dim NodX As Node
   Set GetParentClass = Nothing
   Nv% = GetNivel(pNo)
   If Nv% > 2 Then
      For i% = Nv% - 1 To 1 Step -1
         Set NodX = GetParent(pNo, i%)
         If isClass(NodX) Then
            Set GetParentClass = NodX
            Exit For
         End If
      Next
   Else
      Set GetParentClass = Nothing
   End If

End Function
Public Function Get_QtdLin_Codigo(pVbMember As MEMBER, Optional LinIni = 0) As Long
   Dim QtdLin&
   If pVbMember.Type = vbext_mt_Method Then
      QtdLin = mvarVbCode.ProcCountLines(pVbMember.Name, vbext_ProcKind.vbext_pk_Proc)
   ElseIf pVbMember.Type = vbext_mt_Property Then
      QtdLin = mvarVbCode.ProcCountLines(pVbMember.Name, vbext_ProcKind.vbext_pk_Get)
   Else
      QtdLin = 2
   End If
   If LinIni = 0 Then
      LinIni = Get_LinIni_Codigo(pVbMember)
   End If
   Do While mvarVbCode.ProcOfLine(LinIni + QtdLin - 1, vbext_pk_Proc) <> pVbMember.Name
      QtdLin = QtdLin - 1
      If QtdLin <= 1 Then Exit Do
   Loop
   
   Get_QtdLin_Codigo = QtdLin
End Function
Public Function Get_LinIni_Codigo(pVbMember As MEMBER) As Long
   Dim LinIni&
   LinIni = pVbMember.CodeLocation
   Do While InStr(mvarVbCode.Lines(LinIni, 1), pVbMember.Name) = 0
      LinIni = LinIni + 1
      If mvarVbCode.Lines(LinIni, 1) = "" Then Exit Do
   Loop
   Get_LinIni_Codigo = LinIni
End Function

Public Function isClassChild(pNo As Node) As Boolean
   Dim i%, Nv%, eClasse As Boolean
   eClasse = ("VBClass" = GetParent(pNo, 1).Key)
   If eClasse Then
      Nv% = GetNivel(pNo)
      If Nv% > 1 Then
         For i% = Nv% - 1 To 1 Step -1
            If isClass(GetParent(pNo, i%)) Then
               eClasse = True
               Exit For
            End If
         Next
      Else
         eClasse = False
      End If
   End If
   isClassChild = eClasse
End Function
Public Function ImportVBProject(ByVal strFile As String) As Boolean
'************
'* Objetivo : Importar um Projeto à árvore de estrutura do programa atraves do seu nome físico.
'************
   Dim Title$, fNum%, sLine$
   Dim lKey$
   ' On Error GoTo ErrHandler
   If InStr(LCase(strFile), ".vbp") <> 0 Then
      lKey$ = GetCaption(strFile)
      Title = LerPropriedade(strFile, "Name") & " (" & lKey$ & ")"
   ElseIf InStr(LCase(strFile), ".vbg") <> 0 Then
      Title = LerPropriedade(strFile, "StartupProject")
   Else
      Title = GetCaption(strFile)
   End If

   Call AddVBProjectItems(strFile, NewFolder(Title, lKey$).Key, True)
   Call MontaSubClass
   ImportVBProject = True
   Exit Function
ErrHandler:
'    ErrHandler Err, Error, "Project.Import"
End Function
Public Function AddVBProjectItems(ByVal sProjectFile As String, ByVal sParent As String, Optional bStartUpProject As Boolean) As Boolean
'************
'* Objetivo : Adicionar um Projeto à árvore de estrutura do programa.
'************
    'On Error GoTo ErrHandler
    Dim nFileNum As Integer
    Dim sLine As String
    Dim sFileTitle As String
    Dim lngEqualPos As Long
    Dim sFileName As String
    Dim nodNode As Node
    Dim sText As String
    Dim sKey As String
    Dim sParentNode As String
    Dim strVBProjectText As String
    nFileNum = FreeFile
    Open sProjectFile For Input As nFileNum
    
    mvarMe.TreProj.Nodes(sParent).Selected = True
    mvarMe.TreProj.Nodes(sParent).Expanded = True
    sParentNode = NewFolder("Related Documents", "VBRelDoc", True, sParent).Key
    sParentNode = NewFolder("Class Modules", "VBClass", True, sParent).Key
    sParentNode = NewFolder("Modules", "VBModule", True, sParent).Key
    sParentNode = NewFolder("Forms", "VBForm", True, sParent).Key
    Do While Not EOF(nFileNum)
        Line Input #nFileNum, sLine
        lngEqualPos = InStr(1, sLine, "=")
        If lngEqualPos > 1 Then
            sText = ""
            Select Case UCase$(Mid(sLine, 1, lngEqualPos - 1))
                Case "CLASS"
                   sText = "Class Modules"
                   sKey = "VBClass"
                Case "FORM"
                   sText = Mid(sLine, 1, lngEqualPos - 1) & "s"
                   sKey = "VBForm"
                Case "MODULE"
                   sText = Mid(sLine, 1, lngEqualPos - 1) & "s"
                   sKey = "VBModule"
                Case "RESFILE32"
                   sText = "Related Documents"
                   sKey = "VBRelDoc"
                Case "USERCONTROL"
                   sText = Mid(sLine, 1, lngEqualPos - 1) & "s"
                   sKey = "VBRelDoc"
                Case "PROPERTYPAGE"
                   sText = Mid(sLine, 1, lngEqualPos - 1) & "s"
                   sKey = "VBRelDoc"
                Case "DESIGNER"
                   sText = Mid(sLine, 1, lngEqualPos - 1) & "s"
                   sKey = "VBRelDoc"
                Case Else
                    If strVBProjectText = "" Then
                        strVBProjectText = sLine
                    Else
                        strVBProjectText = strVBProjectText & vbCrLf & sLine
                    End If
            End Select
            If Trim(sText) <> "" Then
               sParentNode = NewFolder(sText, sKey, True, sParent).Key
               Call ParseFileName(sProjectFile, sLine, sFileName, sFileTitle)
               'sParent.Selected = True
               Call AddItem(sParentNode, sFileTitle, sFileName, vbwFile, False)
            End If
        End If
    Loop
    Close nFileNum
    If bStartUpProject Then
       mvarMe.TreProj.Nodes(sParent).Tag = "1" & sProjectFile & vbCrLf & strVBProjectText
       Call BoldTreeNode(mvarMe.TreProj, mvarMe.TreProj.Nodes(sParent))
    Else
       mvarMe.TreProj.Nodes(sParent).Tag = "0" & sProjectFile & vbCrLf & strVBProjectText
    End If
    
    If mvarMe.TreProj.Nodes("VBRelDoc").Children = 0 Then
       mvarMe.TreProj.Nodes.Remove "VBRelDoc"
    End If
    If mvarMe.TreProj.Nodes("VBClass").Children = 0 Then
       mvarMe.TreProj.Nodes.Remove "VBClass"
    End If
    If mvarMe.TreProj.Nodes("VBModule").Children = 0 Then
       mvarMe.TreProj.Nodes.Remove "VBModule"
    End If
    If mvarMe.TreProj.Nodes("VBForm").Children = 0 Then
       mvarMe.TreProj.Nodes.Remove "VBForm"
    End If
    
    AddVBProjectItems = True
    Exit Function
ErrHandler:
    Close nFileNum
'    ErrHandler Err, Error, "Project.AddVBProjectItems"
End Function
Public Function NewFolder(Optional StrNewFolder As String = "", Optional StrKeyNewFolder As String = "", Optional blnNoError As Boolean, Optional nodNode As String) As Node
    Dim FolderName As String
    Dim nodParent As Node
    Dim strTest As String
    Dim Img$
    '// reset the file pos
    '// add the file, and the rest of the string
    If StrNewFolder = Empty Then StrNewFolder = InputBox("Enter the new folder name")

    If StrNewFolder = Empty Then
        '// still empty!
        Set NewFolder = Nothing
        Exit Function
    End If
    If StrKeyNewFolder = Empty Then StrKeyNewFolder = StrNewFolder
    Dim frmD As Node
    If nodNode = Empty Then
        FolderName = GetParent0
    Else
        FolderName = nodNode
    End If
    '// Set the object
    'Set frmD = mvarMe.TreProj.Nodes.Add(FolderName, tvwChild, , strNewFolder, "FolderClosed")
    If InStr(LCase(FolderName), ".vbp") = 0 Then
       Set frmD = mvarMe.TreProj.Nodes.Add(, , , StrNewFolder, Icon.VBPImage)
    Else
       Set frmD = mvarMe.TreProj.Nodes.Add(FolderName, tvwChild, , StrNewFolder)
    End If
    
    '* Get the Parent
    On Error GoTo foldererror
    '* Set the Parent
    If Trim(FolderName) <> "" Then
       Set frmD.Parent = mvarMe.TreProj.Nodes(FolderName)
    End If
    
    '* Set the Key
    frmD.Key = StrKeyNewFolder
    
    '* Set the Image
    Img$ = GetImage(vbwFile, StrNewFolder)
    frmD.Image = IIf(Img$ = "Misc", Icon.CloseImage, Img$)
    frmD.ExpandedImage = IIf(frmD.Image = Icon.CloseImage, Icon.OpenImage, frmD.ExpandedImage)
    'frmD.ExpandedImage = "FolderOpen"
    
    '* Set the Order
    If Trim(FolderName) <> "" Then
       frmD.Sorted = True
    End If
    
    '* Remove Variables
    Set NewFolder = frmD
    Set frmD = Nothing
    Exit Function
foldererror:
'    If Not (blnNoError) Then ErrHandler vbObjectError + 1010, "That folder already exists in the directory", "Project.NewFolder"
    Call mvarMe.TreProj.Nodes.Remove(frmD.Index)
    '// return existing node
    On Error Resume Next
    Set NewFolder = mvarMe.TreProj.Nodes(StrKeyNewFolder)
End Function
Public Function GetParent0() As String
    On Error Resume Next
    If mvarMe.TreProj.Nodes.Count = 0 Then Exit Function
'    mvarMe.TreProj.SelectedItem.Bold = False
    If Err Then
        mvarMe.TreProj.Nodes("Project").Selected = True
    End If
    On Error GoTo 0
    If mvarMe.TreProj.SelectedItem.Image = Icon.CloseImage Then
        On Error Resume Next
        GetParent0 = mvarMe.TreProj.SelectedItem.FullPath
        On Error GoTo 0
    ElseIf mvarMe.TreProj.SelectedItem.Key = "Project" Then
        GetParent0 = "Project"
    Else
        GetParent0 = mvarMe.TreProj.SelectedItem.Parent.FullPath
    End If
End Function
Public Sub AddItem(ByVal strParent As String, ByVal strText As String, ByVal strPath As String, Optional ByVal ShortcutType As ShortcutTypes, Optional blnOverwrite As Boolean = False)
    Dim nodNew As Node
    '// Set the object
    Set nodNew = mvarMe.TreProj.Nodes.Add(strParent, tvwChild, , strText)
    With nodNew
        '// set the parent
        'Set .Parent = mvarMe.TreProj.Nodes(strParent)
        If IsMissing(ShortcutType) Then
            .Image = Icon.CloseImage

            .Tag = strPath
        Else
            '// set the tag, with the actual path/url etc
            .Tag = ShortcutType & strPath
            '// set the image
            .Image = GetImage(ShortcutType, strPath)
        End If
        .Sorted = True
        '// set the key
        On Error Resume Next
        If InStr(Trim(strText), " ") <> 0 Then
           .Key = Mid(strText, 1, InStr(Trim(strText), " ") - 1) '.FullPath
        Else
           .Key = strText '.FullPath
        End If
        
        If Err Then
            If blnOverwrite Then
                mvarMe.TreProj.Nodes.Remove (.FullPath)
                .Key = .FullPath
            Else
                mvarMe.TreProj.Nodes.Remove (nodNew.Index)
                
            End If
        End If
    End With
End Sub
Public Function LerPropriedade(Arquivo As String, Propriedade As String, Optional RetiraAspas = True) As String
   Dim lNum%, lLine$
   lNum = FreeFile
   Open Arquivo For Input As lNum
   Do While Not EOF(lNum)
      Line Input #lNum, lLine$
      If InStr(LCase(Trim(lLine$)), LCase(Propriedade)) = 1 Then
         lLine$ = Trim(Mid(lLine$, InStr(lLine$, "=") + 1))
         Exit Do
      End If
   Loop
   Close lNum
   If RetiraAspas Then lLine$ = EliminarString(lLine$, """")
   LerPropriedade = lLine$
End Function
Public Sub MontaArvoreProjeto()
'************
'* Objetivo : Importar Projeto através do seu nome físico.
'************
   mvarMe.TreProj.Nodes.Clear
   If Not ImportVBProject(mvarVbProj.FILENAME) Then
      Call ExibirAviso("Projeto não foi carregado.", "Atenção")
   End If
End Sub
Public Sub MontaArvoreProjeto0()
   Dim n, lNome$, lDESC$, lKey$, lTipo$, i%
   Dim lPai$
   Dim NodX As Node, MyClass As CLASSE
   
   mvarMe.TreProj.Nodes.Clear
   Set NodX = mvarMe.TreProj.Nodes.Add(, , mvarVbProj.Name, mvarVbProj, Icon.VBPImage)
   NodX.ExpandedImage = Icon.VBPImage
   NodX.Expanded = True
'   NodX.Sorted = True
   '* Construir árvore do Projeto
   
   '********
   '* Inclui Pasta de Forms
   If mvarPROJETO.FORMS.Count > 0 Then  'Not ExisteNo(mvarMe.TreProj, "VBForm")
      Set NodX = mvarMe.TreProj.Nodes.Add(mvarVbProj.Name, tvwChild, "VBForm", "Forms", "CLOSED")
      NodX.ExpandedImage = Icon.OpenImage
      NodX.Sorted = True
   End If
   For Each n In mvarPROJETO.FORMS
      Set NodX = mvarMe.TreProj.Nodes.Add("VBForm", tvwChild, n.NOME, n.NOME, IIf(n.MDI, "MDIFORM", "FORM"))
      NodX.Sorted = True
   Next
   
   '********
   '* Inclui Pasta de Módulos
   If mvarPROJETO.MODULOS.Count > 0 Then  'Not ExisteNo(mvarMe.TreProj, "VBModule")
      Set NodX = mvarMe.TreProj.Nodes.Add(mvarVbProj.Name, tvwChild, "VBModule", "Modules", "CLOSED")
      NodX.ExpandedImage = Icon.OpenImage
      NodX.Sorted = True
   End If
   For Each n In mvarPROJETO.MODULOS
      Set NodX = mvarMe.TreProj.Nodes.Add("VBModule", tvwChild, n.NOME, n.NOME, "MODULO")
      NodX.Sorted = True
   Next
   
   '********
   '* Inclui Pasta de Classes
   If mvarPROJETO.CLASSES.Count > 0 Then  'Not ExisteNo(mvarMe.TreProj, "VBClass")
      Set NodX = mvarMe.TreProj.Nodes.Add(mvarVbProj.Name, tvwChild, "VBClass", "Class Modules", "CLOSED")
      NodX.ExpandedImage = Icon.OpenImage
      NodX.Sorted = True
   End If
   
   For Each n In mvarPROJETO.CLASSES
      Set NodX = mvarMe.TreProj.Nodes.Add("VBClass", tvwChild, n.NOME, n.NOME, Icon.CLSImage)
      NodX.Sorted = True
   Next
   '********
   '* Monta Estrutura de Sub-Classes
   For Each n In mvarPROJETO.CLASSES
      Set MyClass = n
      i = 0
      lPai$ = mvarVbProj.VBComponents(MyClass.NOME).ReadProperty("Member0")
      Do While lPai$ <> ""
         If ExisteNo(mvarMe.TreProj, lPai$) Then
            MyClass.SubClasses.Add mvarPROJETO.CLASSES(lPai$)
         End If
         i = i + 1
         lPai$ = mvarVbProj.VBComponents(n.NOME).ReadProperty("Member" & Trim(CStr(i)))
      Loop
   Next
'   mvarMe.TreProj.Nodes("VBClass").Expanded = True
'   On Error Resume Next
   For Each n In mvarPROJETO.CLASSES
      Set MyClass = n
      With MyClass
          
         If mvarVbProj.VBComponents(.NOME).ReadProperty("Top_Level") = "Yes" Then
            Call MontaSubClasse(mvarMe.TreProj, MyClass)
         End If
      End With
   Next
   
   '********
   '* Inclui Pasta de ResFile
   If mvarPROJETO.RESFILE <> "" Then  'Not ExisteNo(mvarMe.TreProj, "VBRelDoc")
      Set NodX = mvarMe.TreProj.Nodes.Add(mvarVbProj.Name, tvwChild, "VBRelDoc", "Related Documents", "CLOSED")
      NodX.ExpandedImage = Icon.OpenImage
      NodX.Sorted = True
      
      lNome = mvarPROJETO.RESFILE  '"(" & GetNameFromPath(n.FileNames(1)) & ")"
      Set NodX = mvarMe.TreProj.Nodes.Add("VBRelDoc", tvwChild, lNome, lNome, "RES")
      NodX.Sorted = True

   End If
   'For Each n In mvarPROJETO.RELDOCS
   '   Set NodX = mvarMe.TreProj.Nodes.Add("VBClass", tvwChild, n.Name, n.Name, "RELDOC")
   '   NodX.Sorted = True
   'Next
End Sub
Public Sub MontaSubClass()
   Dim n, i%, lPai$
   Dim MyClass As CLASSE

   '********
   '* Monta Estrutura de Sub-Classes
   For Each n In mvarPROJETO.CLASSES
      Set MyClass = n
      i = 0
      lPai$ = mvarVbProj.VBComponents(MyClass.NOME).ReadProperty("Member0")
      Do While lPai$ <> ""
         If ExisteNo(mvarMe.TreProj, lPai$) Then
            MyClass.SubClasses.Add mvarPROJETO.CLASSES(lPai$)
         End If
         i = i + 1
         lPai$ = mvarVbProj.VBComponents(n.NOME).ReadProperty("Member" & Trim(CStr(i)))
      Loop
   Next
   For Each n In mvarPROJETO.CLASSES
      Set MyClass = n
      With MyClass
         If mvarVbProj.VBComponents(.NOME).ReadProperty("Top_Level") = "Yes" Then
            Call MontaSubClasse(mvarMe.TreProj, MyClass)
         End If
      End With
   Next
End Sub
Public Sub MontaSubItens(NomeComponente As String)
   Dim MyComp As VBComponent
   Dim MyObj As Object
   Dim NodX As Node
   Dim lComp0$, lComp$, lNome$, lTipo$, lKey$, lPai$, lImage$
   Dim i%, Pos%, eClasse As Boolean, eFilhoClasse As Boolean
   Dim n As Variant
'   On Error GoTo Fim
'   Screen.MousePointer = vbArrowHourglass
'   isClass = ("VBClass" = GetParent(mvarMe.TreProj.SelectedItem, 1).Key)
   If UCase(Right(NomeComponente, 3)) = "RES" Then GoTo Fim
   If Trim(NomeComponente) <> "" Then
      eClasse = isClass(mvarMe.TreProj.Nodes(NomeComponente))
      eFilhoClasse = isClassChild(mvarMe.TreProj.Nodes(NomeComponente))
      If Not eClasse And Not eFilhoClasse Then
         If mvarMe.TreProj.Nodes(NomeComponente).Children > 0 Then GoTo Fim
         If GetNivel(mvarMe.TreProj.Nodes(NomeComponente)) = 4 Then GoTo Fim
         lComp$ = GetParent(mvarMe.TreProj.Nodes(NomeComponente), 2).Key
         lComp0$ = lComp$
      Else
         If eFilhoClasse And Not eClasse Then
            GoTo Fim
         End If
         Pos = InStr(mvarMe.TreProj.Nodes(NomeComponente).Key, Chr(255))
         Pos = IIf(Pos = 0, Len(mvarMe.TreProj.Nodes(NomeComponente).Key) + 1, Pos)
         lComp0$ = mvarMe.TreProj.Nodes(NomeComponente).Key
         lComp$ = Mid(mvarMe.TreProj.Nodes(NomeComponente).Key, 1, Pos - 1)
         'lComp$ = mvarMe.TreProj.SelectedItem.Key
      End If
      
      '***********
      '* Carregar Classes com Membros do Projeto
      Set MyObj = GetMember(lComp$)
      If MyObj Is Nothing Then GoTo Fim
      lComp$ = MyObj.NOME
      Set MyComp = mvarVbProj.VBComponents.Item(lComp$)
      If Not MyObj.MEMBER.Carregado Then
         Call CarregaMembers(MyObj.PATH & MyObj.FILENAME, MyObj)
         While mvarMe.TreProj.Nodes(lComp$).Children > 0
             mvarMe.TreProj.Nodes.Remove mvarMe.TreProj.Nodes(lComp$).Child.Key
         Wend
'         Call CarregaMembers(MyComp, MyObj)
      End If
   
      If Not Sys.ExibeSubPasta Then GoTo Fim

      If MyObj.MEMBER.Carregado And mvarMe.TreProj.Nodes(lComp0$).Children <> 0 Then
         Set NodX = mvarMe.TreProj.Nodes(lComp0$).Child
         For i = 1 To mvarMe.TreProj.Nodes(lComp0$).Children
            If NodX.Image = Icon.CloseImage Or NodX.Image = Icon.OpenImage Then
               GoTo Fim
            End If
            Set NodX = NodX.Next
         Next
         'GoTo Fim
      End If
      
      With MyObj.MEMBER
         
         '* EVENTOS
         If .EVENTOS.Count > 0 Then
            'If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Eventos
            lPai$ = lComp0$ & "EVENTO"
            Set NodX = mvarMe.TreProj.Nodes.Add(lComp0$, tvwChild, lPai$, " Eventos", "CLOSED")
            NodX.ExpandedImage = Icon.OpenImage
            NodX.Sorted = True
            'End If
            For Each n In .EVENTOS
               lImage$ = "EVENTO"
               lKey = lPai$ & CStr(mvarMe.TreProj.Nodes(lPai).Children + 1)
               Set NodX = mvarMe.TreProj.Nodes.Add(lPai$, tvwChild, lKey, n.NOME, lImage$)
               NodX.Sorted = True
            Next
         End If
         
         '* FUNCOES
         If .FUNCOES.Count > 0 Then
            'If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Eventos
            lPai$ = lComp0$ & "METODO"
            Set NodX = mvarMe.TreProj.Nodes.Add(lComp0$, tvwChild, lPai$, " Métodos", "CLOSED")
            NodX.ExpandedImage = Icon.OpenImage
            NodX.Sorted = True
            'End If
            For Each n In .FUNCOES
               lImage$ = "METODO"
               lKey = lPai$ & CStr(mvarMe.TreProj.Nodes(lPai).Children + 1)
               Set NodX = mvarMe.TreProj.Nodes.Add(lPai$, tvwChild, lKey, n.NOME, lImage$)
               NodX.Sorted = True
            Next
         End If
         
         '* METODOS
         
         If .METODOS.Count > 0 Then
            'If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Métodos
            lPai$ = lComp0$ & "METODO"
            Set NodX = mvarMe.TreProj.Nodes.Add(lComp0$, tvwChild, lPai$, " Métodos", "CLOSED")
            NodX.ExpandedImage = Icon.OpenImage
            NodX.Sorted = True
            'End If
            For Each n In .METODOS
               lImage$ = "METODO"
               lKey = lPai$ & CStr(mvarMe.TreProj.Nodes(lPai).Children + 1)
               Set NodX = mvarMe.TreProj.Nodes.Add(lPai$, tvwChild, lKey, n.NOME, lImage$)
               NodX.Sorted = True
            Next
         End If
         
         '* VARIAVEL
         If .VARIAVEIS.Count > 0 Then
            'If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Eventos
            lPai$ = lComp0$ & "VARIAVEL"
            Set NodX = mvarMe.TreProj.Nodes.Add(lComp0$, tvwChild, lPai$, " Variáveis", "CLOSED")
            NodX.ExpandedImage = "OPEN"
            NodX.Sorted = True
            'End If
            For Each n In .VARIAVEIS
               If LCase(n.TIPO) = "enum" Then
                  lImage$ = "ENUM"
               ElseIf LCase(n.TIPO) = "type" Then
                  lImage$ = "ENUM"
               Else
                 lImage$ = IIf(n.isConst, "CONSTANTE", "VARIAVEL")
               End If
               'lImage$ = "VARIAVEL"
               lKey = lPai$ & CStr(mvarMe.TreProj.Nodes(lPai).Children + 1)
               Set NodX = mvarMe.TreProj.Nodes.Add(lPai$, tvwChild, lKey, n.NOME, lImage$)
               NodX.Sorted = True
            Next
         End If
      End With
   End If
   Exit Sub
Fim:
  ShowError
End Sub
Public Sub MontaSubClasse(Tree As Object, pClass As CLASSE, Optional SubClasse)
   Dim i%, lPai$, Txt$
   Dim NodX As Node
   For i = 1 To pClass.SubClasses.Count
      If pClass.SubClasses(i).PAI = "" Then
         pClass.SubClasses(i).PAI = pClass.NOME
         pClass.SubClasses(i).OCCUR = "1"
         Set Tree.Nodes(pClass.SubClasses(i).NOME).Parent = Tree.Nodes(pClass.NOME)
      Else
         Txt$ = pClass.SubClasses(i).NOME & Chr(255) & pClass.SubClasses(i).OCCUR
         If IsEmpty(SubClasse) Or IsMissing(SubClasse) Then
            lPai$ = pClass.NOME
            'lPai = pClass.NOME & IIf(Val(pClass.OCCUR) <= 1, "", Chr(255) & Trim(CStr(Val(pClass.OCCUR) - 1)))
         Else
            lPai$ = SubClasse
         End If
         'lPai$ = pClass.NOME & IIf(Val(pClass.OCCUR) <= 1, "", Chr(255) & Trim(CStr(Val(pClass.OCCUR) - 1)))
         Set NodX = mvarMe.TreProj.Nodes.Add(lPai$, tvwChild, Txt, pClass.SubClasses(i).NOME, Icon.CLSImage)
         NodX.Sorted = True
         pClass.SubClasses(i).OCCUR = Trim(CStr(Val(pClass.SubClasses(i).OCCUR) + 1))
      End If
      Tree.Nodes(pClass.NOME).Sorted = True
      If pClass.SubClasses(i).SubClasses.Count > 0 Then
         lPai$ = pClass.SubClasses(i).NOME & IIf(Val(pClass.SubClasses(i).OCCUR) <= 1, "", Chr(255) & Trim(CStr(Val(pClass.SubClasses(i).OCCUR) - 1)))
         Call MontaSubClasse(Tree, pClass.SubClasses(i), lPai$)
      End If
   Next
End Sub
Public Sub MontaListItens(Optional NumTab = "")
   Dim MyObj As Object, n As Variant, i%
   Dim lComp$, lImage$
   Dim ItemX As ListItem
   On Error GoTo Fim

   '* Definir Chave do Component
   If mvarVbComp Is Nothing Then GoTo Fim
   lComp$ = mvarVbComp.Name
   'cCompAtual = IIf(cCompAtual = "", lComp, "")
   If cCompAtual = lComp And Trim(cCompAtual) <> "" And _
      mvarMe.TabComp.Tab <> 5 And NumTab = "" Then
      'CInt(NumTab) <> mvarMe.TabComp.Tab Then
      Exit Sub
   Else
      cCompAtual = lComp
   End If
   Set MyObj = GetMember(lComp)
   
   If NumTab = "" Then NumTab = mvarMe.TabComp.Tab
   If mvarMe.TabComp.Tab <> 4 Then mvarMe.LstItens.ListItems.Clear
   Select Case NumTab
      Case 0 '* CompType = vbext_mt_Property
         For Each n In MyObj.MEMBER.PROPRIEDADES
            Set ItemX = mvarMe.LstItens.ListItems.Add(, , n.NOME, "PROPRIEDADE", "PROPRIEDADE")
         Next
      
      Case 1 '* CompType = vbext_mt_Method
         i = 0
         Dim MyMetodo As METODO
         Dim StrAux$
         Dim Pos%
         For Each n In MyObj.MEMBER.METODOS
            Set MyMetodo = n
            i = i + 1
            With MyMetodo
               Set ItemX = mvarMe.LstItens.ListItems.Add(, , .NOME, "METODO", "METODO")
                                 
               Set .PagProp = Nothing
               Call GetPropPage(.PagProp, mvarVbComp, .NOME)
               With .PagProp
                  ItemX.SubItems(1) = .Retorno
                  For i = 1 To .Argumentos.Count
                     StrAux$ = EliminarString(.Argumentos(i), "BYREF", False)
                     ItemX.SubItems(2) = ItemX.SubItems(2) & IIf(i = 1, "", ", ")
                     ItemX.SubItems(2) = ItemX.SubItems(2) & Trim(StrAux$)
                  Next
               End With
            End With
            Set MyMetodo = Nothing
         Next
      
      Case 2 '* CompType = vbext_mt_Event
         For Each n In MyObj.MEMBER.EVENTOS
            Set ItemX = mvarMe.LstItens.ListItems.Add(, , n.NOME, "EVENTO", "EVENTO")
         Next
      
      Case 3 '* CompType = vbext_mt_Variable
         For Each n In MyObj.MEMBER.VARIAVEIS
            If LCase(n.TIPO) = "enum" Then
               lImage$ = "ENUM"
            ElseIf LCase(n.TIPO) = "type" Then
               lImage$ = "ENUM"
            Else
              lImage$ = IIf(n.isConst, "CONSTANTE", "VARIAVEL")
            End If
            Set ItemX = mvarMe.LstItens.ListItems.Add(, , n.NOME, lImage$, lImage$)
         Next
      
      Case 4
         mvarMe.LstItens.ListItems.Clear
         Call MontaListItens(0)
         Call MontaListItens(1)
         Call MontaListItens(2)
         Call MontaListItens(3)
      Case 5 '* Código
         Call ExibirCodigo
         Exit Sub
   End Select
Exit Sub
Fim:
   ShowError
End Sub
Public Function DefineImage(lTipo As Integer) As String
   Select Case CInt(lTipo)
      Case vbext_mt_Method:   DefineImage = "METODO"
      Case vbext_mt_Property: DefineImage = "PROPRIEDADE"
      Case vbext_mt_Variable: DefineImage = "VARIAVEL"
      Case vbext_mt_Event:    DefineImage = "EVENTO"
      Case vbext_mt_Const:    DefineImage = "CONSTANTE"
      Case Else:              DefineImage = "CLOSED"
   End Select
End Function
Public Function DefinePai(lTipo) As String
   Dim lComp$
   If Not mvarVbComp Is Nothing Then lComp$ = mvarVbComp.Name
   Select Case CInt(lTipo)
      Case vbext_mt_Method:   DefinePai = lComp$ & "METODO"
      Case vbext_mt_Property: DefinePai = lComp$ & "PROPRIEDADE"
      Case vbext_mt_Variable: DefinePai = lComp$ & "VARIAVEL"
      Case vbext_mt_Event:    DefinePai = lComp$ & "EVENTO"
      Case vbext_mt_Const:    DefinePai = lComp$ & "VARIAVEL"
      Case Else:              DefinePai = lComp$ & "SINK"
   End Select
End Function
Private Sub mvarMe_TreProjBeforeLabelEdit(Cancel As Integer)
   Dim No$
   No = mvarMe.TreProj.SelectedItem.Text
   If InArray(No, Array("Métodos", "Eventos", "Propriedades", "Variáveis")) Then Cancel = True
   If InArray(No, Array("Formulários", "Módulos", "Classes", "Documentos")) Then Cancel = True
   If InArray(No, Array("Forms", "Modules", "Class Modules", "Related Documents")) Then Cancel = True
   If No = mvarMe.TreProj.SelectedItem.Root.Text Then Cancel = True
End Sub
Private Sub mvarMe_TreProjDblClick()
   Call mvarMe_MnuMouseClick("00", IndMnuMouse.mm_PROPRIEDADES)
End Sub
Private Sub mvarMe_TreProjExpand(ByVal Node As ComctlLib.Node)
   If mvarMe.TreProj.SelectedItem Is Nothing Then GoTo Fim
   If Node.Key = "VBClass" Then GoTo Fim
   Screen.MousePointer = vbHourglass
   Call MontaSubItens(Node.Key)
Fim:
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_TreProjMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim i%, lKey$
   If mvarMe.TreProj.Nodes.Count = 0 Then Exit Sub
   lKey$ = mvarMe.TreProj.SelectedItem.Text
   Call DefineMembro(mvarMe.TreProj.SelectedItem)
'   Set mvarVbMember = mvarVbCode.Members.Item(lKey$)
   Select Case Button
      Case 1
      Case 2:
         '* Objeto TreProj
         For i = mvarMe.MnuMouse00.LBound To mvarMe.MnuMouse00.UBound
            mvarMe.MnuMouse00(i).Visible = True
         Next
         If UCase(GetParent(mvarMe.TreProj.SelectedItem, 1).Key) = "VBCLASS" And UCase(mvarMe.TreProj.SelectedItem.Key) <> "VBCLASS" Then
            If mvarMe.TreProj.SelectedItem.Image = Icon.CLSImage Then
               mvarMe.MnuMouse00(IndMnuMouse.mm_DEFINIR_COMO).Caption = "Definir como Coleção"
            Else
               mvarMe.MnuMouse00(IndMnuMouse.mm_DEFINIR_COMO).Caption = "Definir como Classe"
            End If
            mvarMe.MnuMouse00(IndMnuMouse.mm_DEFINIR_COMO).Visible = True
         Else
            mvarMe.MnuMouse00(IndMnuMouse.mm_DEFINIR_COMO).Visible = False
         End If
         mvarMe.MnuMouse00(IndMnuMouse.mm_ADICIONAR).Visible = False
         mvarMe.MnuMouse00(IndMnuMouse.mm_RECARREGAR_OBJ).Visible = (GetNivel(mvarMe.TreProj.SelectedItem) = 2)
         Call mvarMe.PopupMenu(mvarMe.MnuMouse(0))
   End Select
End Sub

Private Sub mvarMe_TreProjNodeClick(Node As Variant)
   Screen.MousePointer = vbHourglass
   If Node.Image = Icon.VBPImage Then GoTo Fim
   Call DefineMembro
   Call MontaSubItens(Node.Key)
   Call Rel_Tab_Tree
   Call MontaListItens
   Screen.MousePointer = vbDefault
   Exit Sub
Fim:
   ShowError
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_TxtCodeChange()
   'Call TrocarLinhasCodigo
End Sub
Private Function GetMember(ByVal NomeComponente$) As Object
   Dim MyObj As Object
   Dim Pos%
   On Error GoTo Fim
   If Not isClass(mvarMe.TreProj.Nodes(NomeComponente)) Then
      Pos% = GetNivel(mvarMe.TreProj.Nodes(NomeComponente))
      Pos = IIf(Pos <= 3, 4, Pos)
      NomeComponente = GetParent(mvarMe.TreProj.Nodes(NomeComponente), Pos% - 2).Key
   End If
   '***********
   '* Carregar Classes com Membros do Projeto
   Select Case GetParent(mvarMe.TreProj.Nodes(NomeComponente), 1).Key
      Case "VBForm": Set MyObj = mvarPROJETO.DeFORMULARIO(NomeComponente)
      Case "VBModule":  Set MyObj = mvarPROJETO.DeMODULO(NomeComponente)
      Case "VBClass"
         If NomeComponente <> "VBClass" Then
            Pos% = InStr(NomeComponente, Chr(255))
            Pos = IIf(Pos = 0, Len(NomeComponente) + 1, Pos)
            NomeComponente = Mid(NomeComponente, 1, Pos - 1)
            Set MyObj = mvarPROJETO.DeClasse(NomeComponente)
         End If
      Case "VB":  Set MyObj = mvarPROJETO.DeFORMULARIO(NomeComponente)
   End Select
   Set GetMember = MyObj
Exit Function
Fim:
   ShowError
End Function
Private Sub mvarMe_TxtCodeClick()
   Dim LinhaAtual As Long
   Dim PosIni As Long
   Dim PosFim As Long
      
   LinhaAtual = mvarMe.TxtCode.GetLineFromChar(mvarMe.TxtCode.SelStart)
   If cLinhaAntes <> LinhaAtual Then
      If mvarMe.TxtCode.Text <> cCodigoAntes Then
         PosIni = 1
         While cLinhaAntes <> mvarMe.TxtCode.GetLineFromChar(PosIni)
           PosIni = PosIni + 1
         Wend
         
         PosFim = PosIni
         While cLinhaAntes = mvarMe.TxtCode.GetLineFromChar(PosFim)
           PosFim = PosFim + 1
         Wend
         
         Call ColorizeVbWords(mvarMe.TxtCode, PosIni, PosFim)
      End If
   End If
   cLinhaAntes = mvarMe.TxtCode.GetLineFromChar(mvarMe.TxtCode.SelStart)
   cCodigoAntes = mvarMe.TxtCode.Text
End Sub
Private Sub mvarMe_TxtCodeKeyUp(KeyCode As Integer, Shift As Integer)
   mvarMe.TxtCode.SelColor = RGB(0, 0, 0)
   Select Case KeyCode
      Case vbKeyReturn, vbKeyDown, vbKeyUp, vbKeyEnd, vbKeyHome
         Call mvarMe_TxtCodeClick
   End Select
End Sub

Private Sub mvarMe_TxtCodeMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim i%
   Select Case Button
      Case 1
      Case 2
         '* O menu não pode ter todos os seu itens invisiveis
         '* Logo temos que definir pelo menos um item visivel
         '* apenas leitura.
'         For i = Me.MnuMouse00.LBound To Me.MnuMouse00.UBound
'            Me.MnuMouse00(i).Visible = True
'         Next
         mvarMe.MnuMouse00(IndMnuMouse.mm_SALVAR).Visible = True
         For i = mvarMe.MnuMouse00.LBound To mvarMe.MnuMouse00.UBound
            If i <> IndMnuMouse.mm_SALVAR Then
               mvarMe.MnuMouse00(i).Visible = False
            End If
         Next
         mvarMe.MnuMouse00(IndMnuMouse.mm_IR_PROJETO).Visible = True
         Call mvarMe.PopupMenu(mvarMe.MnuMouse(IndMnuMouse.mm_ADICIONAR))
   End Select
End Sub
