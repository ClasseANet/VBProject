VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CONSTRUTOR"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"AreaTemp"
Attribute VB_Ext_KEY = "Member1" ,"clsPROJETO"
Option Explicit
Public Enum ShortcutTypes
    vbwFile = 0
    vbwFolder = 1
    vbwURL = 2
    vbwEMail = 3
End Enum
Public Enum IndMnuMouse
   mm_ADICIONAR = 0
   mm_ADICIONAR_OBJ
   mm_SALVAR
   mm_Salvar_COMO
   mm_RECARREGAR
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
   EnumImage As String
   TypeImage As String
   LibImage As String
   MDIImage As String
   
   ExitImage As String
   BDWizard As String
End Type

Private Icon As TypeIcon

Private WithEvents mvarMe As FrmAddIn
Attribute mvarMe.VB_VarHelpID = -1
Private mvarPROJETOS As Collection
Private mvarProjeto As clsPROJETO
Private mvarClsConnect As Connect
Private mvarVbProj As VBProject

Private mvarVbComp As VBComponent
Private mvarVbMember As MEMBER
Private mvarVbCode As CodeModule

Private mvarAreaTemp As AreaTemp

Private cCodigoAntes As String
Private cLinhaAntes As Long
Private cCompAtual As String

Private mvarProjetoAtivo As String

Private TreIndAnterior As Integer

Public Property Get AreaTemp() As AreaTemp
    If mvarAreaTemp Is Nothing Then
        Set mvarAreaTemp = New AreaTemp
    End If
    Set AreaTemp = mvarAreaTemp
End Property
Public Property Set AreaTemp(vData As AreaTemp)
    Set mvarAreaTemp = vData
End Property
Public Property Get ProjetoAtivo() As String
   Dim Pos As Integer
   mvarProjetoAtivo = mvarMe.TreProj.SelectedItem.Root.Text
   Pos = InStr(mvarProjetoAtivo, " ")
   If Pos <> 0 Then mvarProjetoAtivo = Trim(Mid(mvarProjetoAtivo, 1, Pos - 1))
   ProjetoAtivo = mvarProjetoAtivo
End Property
Public Property Let ProjetoAtivo(vData As String)
   mvarProjetoAtivo = vData
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
Public Property Get PROJETO(NOME$) As clsPROJETO
   If mvarProjeto Is Nothing Then
       Set mvarProjeto = New clsPROJETO
   End If
   If vbKey0 <= Asc(Mid(NOME$, 1, 1)) And Asc(Mid(NOME$, 1, 1)) <= vbKey9 Then
      Set mvarProjeto = PROJETOS(CInt(NOME$))
   Else
      Set mvarProjeto = PROJETOS(NOME$)
   End If
   Set PROJETO = mvarProjeto
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
   
      .VBPImage = "ICO_PROJ"         ' "VBP"
      .FRMImage = "ICO_FORM"         ' "FRM"
      .CLSImage = "ICO_CLASSE"       ' "CLS"
      .BASImage = "ICO_MODULO"       ' "BAS"
      .RESImage = "ICO_RES"          ' "RES"
      .OpenImage = "ICO_OPEN"        ' "OPEN"
      .CloseImage = "ICO_CLOSED"     ' "CLOSED"
      
      .CollImage = "COLECAO"
      .EvtImage = "ICO_EVENTO"       ' "EVENTO"
      .MetImage = "ICO_METODO"       ' "METODO"
      .PropImage = "ICO_PROPRIEDADE" ' "PROPRIEDADE"
      .VarImage = "ICO_VARIAVEL"     ' "VARIAVEL"
      .ConstImage = "ICO_CONSTANTE"  ' "CONSTANTE"
      .EnumImage = "ICO_ENUM"        ' "ENUM"
      .TypeImage = "ICO_TYPE"        ' "TYPE"
      .LibImage = "ICO_DLL"          ' "DLL"
      .MDIImage = "ICO_MDI"          ' "MDI"
      
      .BDWizard = "ICO_BDWIZARD"
      .ExitImage = "BMP_EXIT"
   End With
End Sub
Private Sub Class_Terminate()
  Set mvarAreaTemp = Nothing
  Set mvarProjeto = Nothing
  'Set mvarPROJ = Nothing
  Set mvarPROJETOS = Nothing
End Sub
Sub Show()
   On Error Resume Next
   If mvarMe Is Nothing Then
      Set mvarMe = New FrmAddIn
   End If
   'Set mvarMe.VBInstance = VBInstance
   'Set mvarMe.Connect = mvarClsConnect
   
   mvarMe.Visible = mvarClsConnect.FormDisplayed
   If mvarMe.Visible Then
      mvarMe.Show
   Else
'      Call CarregaSubItens
   End If
   mvarMe.ZOrder
End Sub
Sub Hide()
   On Error Resume Next
   ClsConnect.FormDisplayed = False
   ClsConnect.Hide
   'mvarMe.Hide
   Unload mvarMe
End Sub
Public Sub SalvarCodigo(Optional FileName, Optional Texto)
   Dim l_FileName As String
   l_FileName = mvarVbComp.FileNames(1)
   
   If Not IsMissing(FileName) Then l_FileName = IIf(Trim(FileName) = "", l_FileName, FileName)
   If IsMissing(Texto) Then Texto = ""

   
   Call TrocarLinhasCodigo(False)
   Call mvarVbComp.SaveAs(l_FileName)
End Sub
Private Sub CarregaProjeto(Optional vProj)
   Set mvarProjeto = New clsPROJETO
   With mvarProjeto
      .NOME = mvarVbProj.Name
      .DESCRIÇÃO = mvarVbProj.Description
      .FileName = mvarVbProj.FileName
      .NOME = mvarVbProj.Name
      .PATH = mvarVbProj.FileName
   End With
   mvarPROJETOS.Add mvarProjeto, mvarProjeto.NOME
   Call CarregaComponentes
End Sub
Private Sub CarregaComponentes(Optional vProj)
   Dim MyComp As VBComponent
   Dim MyObj As Variant, n As Variant
   With mvarProjeto
      For Each n In mvarVbProj.VBComponents
         Set MyComp = n
         Select Case n.Type
            Case vbext_ct_VBMDIForm, vbext_ct_VBForm
               Set MyObj = New clsFORMULARIO
               With MyObj
                  .FileName = GetNameFromPath(MyComp.FileNames(1))
                  .NOME = MyComp.Name
                  .PATH = GetNameFromPath(MyComp.FileNames(1), 1)
                  .MDI = (vbext_ct_VBMDIForm = n.Type)
               End With
               mvarProjeto.FORMS.Add MyObj, MyObj.NOME
            Case vbext_ct_StdModule
               Set MyObj = New clsMODULO
               With MyObj
                  .FileName = GetNameFromPath(MyComp.FileNames(1))
                  .NOME = MyComp.Name
                  .PATH = GetNameFromPath(MyComp.FileNames(1), 1)
                  '.PagProp.TipoPagina = tpMetodo
                  'Call GetPropPage(.PagProp, MyComp, MyComp.Name)
               End With
               mvarProjeto.MODULOS.Add MyObj, MyObj.NOME
            Case vbext_ct_ClassModule
               Set MyObj = New clsCLASSE
               With MyObj
                  .FileName = GetNameFromPath(MyComp.FileNames(1))
                  .NOME = MyComp.Name
                  .PATH = GetNameFromPath(MyComp.FileNames(1), 1)
                  '.PagProp.TipoPagina = tpClasse
                  'Call GetPropPage(.PagProp, MyComp, MyComp.Name)
               End With
               mvarProjeto.CLASSES.Add MyObj, MyObj.NOME
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
   Dim MyMetodo As New clsMETODO
   Dim MyPropriedade As New clsPROPRIEDADE
   Dim MyVar As New clsVARIAVEL
   Dim MyEvento As New clsEVENTO
   Dim MyFuncao As New clsFUNCAO
   Dim i%, NodX As Node, Mbr As Variant
   Dim lImage$, lComp$, lPai$

'on error resume next
   For Each Mbr In vComp.CodeModule.Members
      Select Case Mbr.Type
         '********************************
         Case vbext_mt_Method '* Métodos *
         '********************************
            Set MyMetodo = New clsMETODO
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
            Set MyPropriedade = New clsPROPRIEDADE
            MyPropriedade.NOME = Mbr.Name
            vObj.MEMBER.PROPRIEDADES.Add MyPropriedade, MyPropriedade.NOME
            Set MyPropriedade = Nothing
            
         '************************************
         Case vbext_mt_Variable, vbext_mt_Const '* Variáveis e Constantes*
         '************************************
            Set MyVar = New clsVARIAVEL
            MyVar.Escopo = Mbr.Scope
            MyVar.NOME = Mbr.Name
            MyVar.PassArg = Mbr.Category
            MyVar.TIPO = Mbr.Type
            MyVar.isConst = (vbext_mt_Const = Mbr.Type)
            vObj.MEMBER.VARIAVEIS.Add MyVar, MyVar.NOME
            Set MyVar = Nothing
            
         '*******************************
         Case vbext_mt_Event '* Eventos *
         '*******************************
            Set MyEvento = New clsEVENTO
            MyEvento.NOME = Mbr.Name
            vObj.MEMBER.EVENTOS.Add MyEvento, MyEvento.NOME
            Set MyEvento = Nothing
          
          '**************************************
          Case Else 'vbext_mt_EventSink '* Event Sink *
          '**************************************
             lImage$ = Icon.CloseImage
             lPai$ = lComp$ & "SINK"
             If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Outros
                Set NodX = mvarMe.TreProj.Nodes.Add(lComp$, tvwChild, lPai$, "Outros", Icon.CloseImage)
                NodX.ExpandedImage = "OPEN"
                NodX.Sorted = True
             End If
      End Select
   Next
   
   vObj.MEMBER.Carregado = True
End Sub
Private Sub CarregaMembers(strFileName As String, ByRef vObj As Object)
   Dim MyMember As New Members
   Dim MyMetodo As New clsMETODO
   Dim MyPropriedade As New clsPROPRIEDADE
   Dim MyVar As New clsVARIAVEL
   Dim MyEvento As New clsEVENTO
   Dim MyFuncao As New clsFUNCAO
   Dim MyLib As New clsLIB
   Dim MyControle As New clsCONTROLE

   Dim i As Integer
   Dim NodX As Node, Mbr As Variant, n As Variant
   Dim lImage$, lComp$, lPai$
   Dim Pos%, PosAux%, VarDeclarado As Boolean
   Dim LinhaCorrente As Long, CodeLine As Boolean, Arr As Variant

   Dim nArq%, Textline$
   Dim StrEscopo As String
   Dim isType As Boolean, isEnum As Boolean
   Dim isConst As Boolean, isObjWithEvents As Boolean
   Dim ExisteItem As Boolean

   Set vObj.MEMBER = Nothing
   nArq = FreeFile
   On Error Resume Next
'   Close #nArq
   Open strFileName For Input As #nArq
   LinhaCorrente = 1
   CodeLine = False
   While Not EOF(nArq)
      Line Input #nArq, Textline
      Textline = Trim(Textline)
      '      If UCase(Mid(Textline & " ", 1, 6)) = "BEGIN " Then
      '         While UCase(Mid(Textline, 1, 3)) <> "END" And Not EOF(nArq)
      '            Line Input #nArq, Textline
      '            Textline = Trim(Textline)
      '            LinhaCorrente= LinhaCorrente+1
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
            If CodeLine Then LinhaCorrente = LinhaCorrente - 1
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
         ElseIf Mid(Trim(Textline), 1, Len("Begin ")) = "Begin " Then
            Set MyControle = New clsCONTROLE
            'MyControle.FileName = strFileName
            Pos = InStr(Textline, ".")
            With MyControle
               .TIPO = RichWordOver(Textline, 0, 0, Pos + 1)
               Pos = InStr(Pos, Textline, " ")
               If .TIPO <> "Form" Then
                  .NOME = Mid(Textline, Pos + 1)
                  .Flag = True
                  .NotNull = True
                  vObj.MEMBER.CONTROLES.Add MyControle, .NOME
               End If
            End With
            Set MyControle = Nothing
            '**********************
            '* Linhas úteis de Programa.
            '**********************
         ElseIf Mid(Textline, 1, 1) <> "'" Then
            '**********************
            '* Definir Sessão de Varáveis
            '**********************
            If InStr(" " & LCase(Textline), LCase(" Function ")) + _
               InStr(" " & LCase(Textline), LCase(" Sub ")) + _
               InStr(" " & LCase(Textline), LCase(" Property ")) <> 0 And _
               Not VarDeclarado Then
               VarDeclarado = True
            End If
            '*********************************
            '* Verificar Sessão de Código VB *
            '*********************************
            If Not CodeLine Then
               Arr = Array("Option", "Public", "Global", "Private", "Dim", "Static", _
                           "Function", "Sub", "Property")
                           
               CodeLine = InArray(RichWordOver(Textline, 0, 0, 1), Arr)
               CodeLine = CodeLine Or (Mid(Textline, 1, 1) = "'")
            End If
            If VarDeclarado Then
               If InStr(Textline, "Declare ") <> 0 Then
                  VarDeclarado = False

                  Set MyLib = New clsLIB
                  'MyLib.FileName = strFileName
                  Pos = InStr(Textline, "Lib ")
                  Pos = InStr(Pos, Textline, """") + 1
                  MyLib.NOME = RichWordOver(Textline, 0, 0, Pos, False)

                  ExisteItem = False
                  For Each n In vObj.MEMBER.LIBS
                     If n.NOME = MyLib.NOME Then
                        ExisteItem = True
                     End If
                  Next
                  With MyLib
                     .StartLine = LinhaCorrente
                     Pos = InStr(LCase(Textline), LCase("Function ")) + 9
                     Set MyFuncao = New clsFUNCAO
                     MyFuncao.FileName = strFileName
                     If Pos > 9 Then
                        MyFuncao.NOME = RichWordOver(Textline, 0, 0, Pos, False)
                     Else
                        Pos = InStr(LCase(Textline), LCase("Sub ")) + 4
                        If Pos > 4 Then
                           MyFuncao.NOME = Trim(Mid(Textline, Pos, InStr(Textline, "Lib ") - Pos))
                        End If
                     End If
                     If MyFuncao.NOME <> "" Then
                        If ExisteItem Then
                           vObj.MEMBER.LIBS(MyLib.NOME).FUNCOES.Add MyFuncao, MyFuncao.NOME
                        Else
                           .FUNCOES.Add MyFuncao, MyFuncao.NOME
                        End If
                     End If
                  End With
                  If Not Err = 35 And Not ExisteItem Then '* Sub or Function not defined
                     Err = 0
                     vObj.MEMBER.LIBS.Add MyLib, MyLib.NOME
                  End If
                  
                  Set MyLib = Nothing
                  Set MyFuncao = Nothing
               ElseIf InStr(LCase(Textline), LCase("Function ")) <> 0 Then
                  Set MyMetodo = New clsMETODO
                  MyMetodo.FileName = strFileName
                  Pos = InStr(LCase(Textline), LCase("Function "))
                  If InStr(Textline, "Declare ") = 0 Then
                     With MyMetodo
                        .NOME = RichWordOver(Textline, 0, 0, Pos + Len("Function "), False)
                        '                     If vObj.LINHAS > 470 And vObj.LINHAS < 520 Then
                        '                        i = i
                        '                    End If
                        .StartLine = mvarVbComp.CodeModule.ProcBodyLine(.NOME, vbext_ProcKind.vbext_pk_Proc)
                        If LinhaCorrente <> MyMetodo.StartLine Then
                           MyMetodo.StartLine = MyMetodo.StartLine
                        End If
                        .StartLine = LinhaCorrente
                        If .PagProp Is Nothing Then
                           Set .PagProp = New PROPPAGE
                           .PagProp.TipoPagina = tpMetodo
                           Call GetPropPage(.PagProp, mvarVbComp, .NOME)
                        End If
                        While Textline <> "End Function" And Not EOF(nArq)
                           Line Input #nArq, Textline
                           Textline = Trim(Textline)
                           '                           Call SomarLinha(Textline, LinhaCorrente%)
                           '                           vObj.LINHAS = LinhaCorrente%
                           '                           Call SetAttribute(Textline, vObj)
                           LinhaCorrente = LinhaCorrente + 1
                        Wend
                        If Not Err = 35 Then '* Sub or Function not defined
                           vObj.MEMBER.METODOS.Add MyMetodo, .NOME
                        End If
                     End With
                     Set MyMetodo = Nothing

                  End If
               ElseIf InStr(LCase(Textline), LCase("Sub ")) <> 0 Then
                  Set MyMetodo = New clsMETODO
                  MyMetodo.FileName = strFileName
                  Pos = InStr(LCase(Textline), LCase("Sub "))
                  With MyMetodo
                     .NOME = RichWordOver(Textline, 0, 0, Pos + Len("Sub "), False)
                     .StartLine = mvarVbComp.CodeModule.ProcBodyLine(.NOME, vbext_ProcKind.vbext_pk_Proc)
                     If LinhaCorrente <> MyMetodo.StartLine Then
                        'MyMetodo.StartLine = MyMetodo.StartLine
                     End If
                     .StartLine = LinhaCorrente
                     If .PagProp Is Nothing Then
                        Set .PagProp = New PROPPAGE
                        .PagProp.TipoPagina = tpMetodo
                        Call GetPropPage(.PagProp, mvarVbComp, .NOME)
                     End If

                     While Textline <> "End Sub" And Not EOF(nArq)
                        Line Input #nArq, Textline
                        Textline = Trim(Textline)
                        '                        Call SomarLinha(Textline, LinhaCorrente%)
                        '                        vObj.LINHAS = LinhaCorrente%
                        '                        Call SetAttribute(Textline, vObj)
                        LinhaCorrente = LinhaCorrente + 1
                     Wend
                     If Not Err = 35 Then '* Sub or Function not defined
                        vObj.MEMBER.METODOS.Add MyMetodo, .NOME
                     End If
                  End With

                  Set MyMetodo = Nothing

               ElseIf InStr(LCase(Textline), (" property ")) <> 0 Or Mid(LCase(Textline), 1, 9) = "property " Then
                  Set MyPropriedade = New clsPROPRIEDADE
                  With MyPropriedade
                     Pos = InStr(LCase(Textline), LCase("Property "))
                     .FileName = strFileName
                     .NOME = RichWordOver(Textline, 0, 0, Pos + 13, False)
                     If .PagProp Is Nothing Then
                              Set .PagProp = New PROPPAGE
                              .PagProp.TipoPagina = tpMetodo
                              Call GetPropPage(.PagProp, mvarVbComp, .NOME)
                     End If
                     Set .PagProp.PROJETO = mvarPROJETOS(ProjetoAtivo)
                     Call GetPropPage(.PagProp, mvarVbComp, .NOME)
                     'MyPropriedade.PagProp.Escopo = ""
                     'MyPropriedade.PagProp.NOME = ""
                     'MyPropriedade.PagProp.TIPO = ""
                     'MyPropriedade.PagProp.TipoPagina = tpPropriedade
                     'MyPropriedade.PagProp.TpDeclaracao = ""
                     'MyPropriedade.PagProp.ValorPadrao = ""
                     
                     On Error Resume Next
                     vObj.MEMBER.PROPRIEDADES.Add MyPropriedade, .NOME
                     If Err = 457 Then '* This key is already associated with an element of this collection
                        Err = 0
                     Else
                        While Textline <> "End Property" And Not EOF(nArq)
                           Line Input #nArq, Textline
                           Textline = Trim(Textline)
                           '                        Call SomarLinha(Textline, LinhaCorrente%)
                           '                        vObj.LINHAS = LinhaCorrente%
                           LinhaCorrente = LinhaCorrente + 1
                        Wend
                        Set MyPropriedade = Nothing
                     End If
                  End With
               End If
            Else

               '************************
               '* Sessão de variáveis. *
               '************************
               Pos = InStr(" " & LCase(Textline), LCase(" Dim ")) + _
                  InStr(" " & LCase(Textline), LCase(" Friend ")) + _
                  InStr(" " & LCase(Textline), LCase(" Global ")) + _
                  InStr(" " & LCase(Textline), LCase(" Private ")) + _
                  InStr(" " & LCase(Textline), LCase(" Public ")) + _
                  InStr(" " & LCase(Textline), LCase(" Static ")) + _
                  InStr(" " & LCase(Textline), LCase(" Const ")) + InStr(" " & LCase(Textline), LCase(" Enum "))
               If Pos <> 0 Then
                  isObjWithEvents = False
                  isConst = False
                  isEnum = False
                  isType = False
                  '**********
                  '** TYPE **
                  '**********
                  If InStr(" " & LCase(Textline), LCase(" Type ")) <> 0 Then
                     isType = True
                     PosAux = InStr(" " & LCase(Textline), LCase(" Type "))
                     If PosAux = 1 Then
                        If LCase(Right(strFileName, 3)) = "bas" Then
                           StrEscopo = "Global"
                        Else
                           StrEscopo = "Private"
                        End If
                        Textline = StrEscopo & " " & Textline
                     Else
                        StrEscopo = RichWordOver(Trim(Textline), 0, 0, 1)
                     End If
                  '**************
                  '** Constant **
                  '**************
                  ElseIf InStr(" " & LCase(Textline), LCase(" Const ")) <> 0 Then
                     PosAux = InStr(" " & LCase(Textline), LCase(" Const "))
                     isConst = True
                     If PosAux = 1 Then
                        If LCase(LeftB(strFileName, 3)) = "bas" Then
                           StrEscopo = "Global"
                        Else
                           StrEscopo = "Private"
                        End If
                        Textline = StrEscopo & " " & Textline
                     Else
                        StrEscopo = RichWordOver(Trim(Textline), 0, 0, 1)
                     End If
                  '**********
                  '** Enum **
                  '**********
                  ElseIf InStr(LCase(Textline), LCase("Enum ")) <> 0 Then
                     isEnum = True
                     PosAux = InStr(" " & LCase(Textline), LCase(" Enum "))
                     If PosAux = 1 Then
                        If LCase(Right(strFileName, 3)) = "bas" Then
                           StrEscopo = "Global"
                        Else
                           StrEscopo = "Private"
                        End If
                        Textline = StrEscopo & " " & Textline
                     Else
                        StrEscopo = RichWordOver(Trim(Textline), 0, 0, 1)
                     End If
                  '****************
                  '** WithEvents **
                  '****************
                  ElseIf InStr(Textline, "WithEvents ") <> 0 Then
                     StrEscopo = Mid(Textline, 1, InStr(Textline, " ") - 1)
                     PosAux = InStr(Textline, "WithEvents ")
                     Textline = Mid(Textline, 1, PosAux - 1) & Mid(Textline, PosAux + Len("WithEvents"))
                     Pos = Pos + Len("WithEvents ")
                     isObjWithEvents = True
                  
                  '************
                  '** Option **
                  '************
                  ElseIf InStr(LCase(Textline), LCase("Option ")) <> 0 Then
                     '* Option Explicity
                  Else
                     StrEscopo = RichWordOver(Textline, 0, 0, 1)
                  End If
                  Set MyVar = New clsVARIAVEL
                  MyVar.FileName = strFileName
                  
                  Textline = Trim(Mid(Trim(Textline), Len(StrEscopo) + 1) & ",")
                  While InStr(Textline, ",") <> 0
                     Set MyVar = New clsVARIAVEL
                     MyVar.FileName = strFileName
                     MyVar.isConst = isConst
                     MyVar.isObjWithEvents = isObjWithEvents
                     MyVar.StartLine = LinhaCorrente
                     If isConst Then
                        '**********************
                        '* Constantes.
                        '**********************
                        Pos = InStr(LCase(Textline), " const ") + 7
                        MyVar.Escopo = StrEscopo
                        MyVar.NOME = RichWordOver(Textline, 0, 0, Pos, False)
                        PosAux = InStr(Textline, " = ") + 3
                        If InStr(PosAux, Textline, Chr(34)) <> 0 Then '* Aspas
                           PosAux = InStr(PosAux, Textline, Chr(34)) + 1
                           MyVar.VALOR = Trim(Mid(Textline, PosAux, InStr(PosAux, Textline, Chr(34)) - PosAux))
                        Else
                           If InStr(PosAux, Textline, " ") <> 0 Then
                              PosAux = InStr(PosAux, Textline, " ")
                              MyVar.VALOR = RichWordOver(Textline, 0, 0, PosAux)
                              'MyVar.VALOR = Trim(Mid(Textline, PosAux, InStr(PosAux, Textline, " ") - PosAux))
                           Else
                              MyVar.VALOR = RichWordOver(Textline, 0, 0, PosAux + 1)
                           End If
                        End If
'                        If Mid(MyVar.VALOR, 1, 1) = Chr(34) Then
'                           MyVar.VALOR = Mid(MyVar.VALOR, 2)
'                        End If
'                        If Right(MyVar.VALOR, 1) = Chr(34) Then
'                           MyVar.VALOR = Mid(MyVar.VALOR, 1, Len(MyVar.VALOR) - 1)
'                        End If

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

                     ElseIf isType Then
                        MyVar.Escopo = StrEscopo
                        MyVar.NOME = RichWordOver(Textline, 0, 0, Len("Type "), False)
                        MyVar.TIPO = "Type"
                        MyVar.PagProp.TipoPagina = tpVariavel
                        Set MyVar.PagProp.PROJETO = mvarPROJETOS(ProjetoAtivo)
                        While Textline <> "End Type" And Not EOF(nArq)
                           Line Input #nArq, Textline
                           If Textline <> "End Type" Then
                              MyVar.PagProp.Argumentos.Add Textline
                           End If
                           '                           Textline = Trim(Textline)
                           '                           Call SomarLinha(Textline, LinhaCorrente%)
                           '                           vObj.LINHAS = LinhaCorrente%
                           LinhaCorrente = LinhaCorrente + 1
                        Wend
                     
                     ElseIf isEnum Then
                        MyVar.Escopo = StrEscopo
                        MyVar.NOME = RichWordOver(Textline, 0, 0, Len("Enum "), False)
                        MyVar.TIPO = "Enum"
                        Set MyVar.PagProp.PROJETO = mvarPROJETOS(ProjetoAtivo)
                        
                        While Textline <> "End Enum" And Not EOF(nArq)
                           Line Input #nArq, Textline
                           If Textline <> "End Enum" Then
                              MyVar.PagProp.Argumentos.Add Textline
                           End If
                           '                           Textline = Trim(Textline)
                           '                           Call SomarLinha(Textline, LinhaCorrente%)
                           '                           vObj.LINHAS = LinhaCorrente%
                           LinhaCorrente = LinhaCorrente + 1
                        Wend

                     Else
                        MyVar.Escopo = StrEscopo
                        MyVar.NOME = RichWordOver(Trim(Textline), 0, 0, 1, False)
                        'MyVar.PassArg = Mbr.Category

                        Pos = InStr(Len(MyVar.NOME), Textline, " As ")
                        If Pos <> 0 Then
                           MyVar.TIPO = RichWordOver(Textline, 0, 0, Pos + 4)
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
                     Textline = Trim(Mid(Textline, InStr(Textline, ",") + 1))
                     If Trim(MyVar.NOME) <> "" Then
                        vObj.MEMBER.VARIAVEIS.Add MyVar, MyVar.NOME
                     End If
                     Set MyVar = Nothing
                  Wend
               ElseIf InStr(LCase(Textline), LCase("Const ")) <> 0 Then

                  Set MyVar = New clsVARIAVEL
                  MyVar.FileName = strFileName
                  MyVar.Escopo = Mbr.Scope
                  MyVar.NOME = Mbr.Name
                  MyVar.PassArg = Mbr.Category
                  MyVar.TIPO = Mbr.Type
                  MyVar.isConst = (vbext_mt_Const = Mbr.Type)
                  MyVar.StartLine = LinhaCorrente
                  vObj.MEMBER.VARIAVEIS.Add MyVar, MyVar.NOME
                  Set MyVar = Nothing
               End If
            End If
         ElseIf Mid(Textline, 1, 1) = "'" And Not CodeLine Then
            CodeLine = True
         End If
      End If
      Err = 0
      If CodeLine Then LinhaCorrente = LinhaCorrente + 1
   Wend
   Close #nArq
   vObj.MEMBER.Carregado = True
End Sub
Public Sub RecarregarMembro(lKey As String)
   GetMember(lKey$).MEMBER.Carregado = False
   Call SetTag(mvarMe.TreProj.Nodes(lKey$), "Text", "")
   Call SetTag(mvarMe.TreProj.Nodes(lKey$), "RtfText", "")
   cCompAtual = ""
   Call mvarMe_TreProjNodeClick(mvarMe.TreProj.Nodes(lKey$))
End Sub
Public Sub DefineMembro(Optional NodeX As Variant)
   Dim lKey$, Pos%, i%
   Dim eClasse As Boolean, eFilhoClasse As Boolean
   Dim Nox As Node
   Dim MesmoCampo As Boolean
   
   On Error GoTo Fim
   If IsMissing(NodeX) Then
      Set Nox = mvarMe.TreProj.SelectedItem
   Else
     Set Nox = NodeX
   End If
   If InArray(Nox.Key, Array("VBForm", "VBModule", "VBClass", "VBLibrary", "VBRelDoc")) Then
      'If Nox.Children > 0 And Nox.Key <> "VBLibrary" Then
      '   Call DefineMembro(Nox.Child)
      'End If
      GoTo Saida
   End If
   MesmoCampo = (TreIndAnterior = mvarMe.TreProj.SelectedItem.Index)
   If Not mvarVbMember Is Nothing Then
      If MesmoCampo And mvarVbMember.Name = Nox.Text Then GoTo Saida
   End If
   If Nox.Image = Icon.VBPImage Then GoTo Saida
   eClasse = isClass(Nox)
   If eClasse Then
      If Nox.Key <> "VBClass" Then
         If Nox.Image = Icon.CLSImage Then
            lKey$ = Nox.Key
         Else
            lKey$ = Trim(Nox.Text)
         End If
      End If
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
   
   If Not GetMember(lKey$).MEMBER.Carregado Then
      i = 0
      While Nox.Children <> i
         If eClasse Then
            If Nox.Child.Image <> Icon.CLSImage Then
               mvarMe.TreProj.Nodes.Remove Nox.Child.Key
            Else
               i = i + 1
            End If
         Else
            mvarMe.TreProj.Nodes.Remove Nox.Child.Key
         End If
      Wend
   End If
   If Not mvarVbComp Is Nothing And mvarMe.TabComp.Tab <> 5 Then
      '* Se o componente já foi carregado e é o mesmo item da árvore
      '* então não precisa caregá-lo de novo
      If Nox.Children <> 0 And lKey$ = mvarVbComp.Name Then GoTo Saida
   End If
   If Trim(lKey$) = "" Then GoTo Saida
   If UCase(Right(lKey$, 3)) = "RES" Then GoTo Saida
   
   
   Set mvarVbComp = mvarVbProj.VBComponents.Item(lKey$)
   Set mvarVbCode = mvarVbComp.CodeModule
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
Saida:
   Exit Sub
Fim:
   If Err = 9 Then
   ElseIf Err <> 0 Then
      mvarMe.TreProj.Nodes(1).Selected = True
      Call ShowError
   End If
End Sub
Public Function PropertyPage(Optional pTipo, Optional NmMember) As Boolean
   Dim MyProp As New PROPPAGE
   Dim MyMet As New clsMETODO
   Dim MyPropriedade As New clsPROPRIEDADE
   Dim MyVar As New clsVARIAVEL
   Dim MyArg As ADDARG
   Dim MyObj As Object, n
   
   Dim Img$, StrAux$, StrArg$
   
   Dim LinhaInicial%, PosRetorno%, PosIniArg%, PosAux%
   Dim StrFunc$, nResp As Boolean
   
   Dim Parenteses%, i%, Pos%, isFunc As Boolean
   Dim StrWord$, StrFrase$, PosW%
   Dim AddNewObj As Boolean
   Dim Palavras As New Collection
   Dim sAntes As String, sDepois As String
   
   '* Definir nome do "Membro"
   AddNewObj = IsMissing(NmMember)
   If AddNewObj Then NmMember = ""
   If Trim(NmMember) = "" Then
      If Not mvarVbMember Is Nothing Then
         NmMember = mvarVbMember.Name
      End If
   End If

   Select Case mvarMe.ActiveControl.Name
      Case "LstItens": Img$ = mvarMe.LstItens.SelectedItem.Icon
      Case "TabComp":  Img$ = mvarMe.LstItens.SelectedItem.Icon
      Case "TreProj":  Img$ = mvarMe.TreProj.SelectedItem.Image
   End Select
   If IsMissing(pTipo) Then
      Select Case Img$
         Case Icon.CLSImage:   pTipo = tpClasse
         Case Icon.CollImage:  pTipo = tpColecao
         Case Icon.EvtImage:   pTipo = tpEvento
         Case Icon.FRMImage:   pTipo = tpForm
         Case Icon.MetImage:   pTipo = tpMetodo
         Case Icon.BASImage:   pTipo = tpModulo
         Case Icon.PropImage:  pTipo = tpPropriedade
         Case Icon.VarImage:   pTipo = tpVariavel
         Case Icon.ConstImage: pTipo = tpVariavel
         Case Else:            pTipo = tpNull
      End Select
   End If

   'StrAux = Me.PROJETO(mvarPROJETO.NOME).CLASSES(mvarVbComp.Name).NOME
   'StrAux = mvarVbComp.Name
   Set MyObj = GetMember(mvarVbComp.Name)
   Select Case pTipo
      Case tpClasse
      Case tpColecao
      Case tpEvento
      Case tpForm
      Case tpMetodo
         If AddNewObj Then
            For i = 1 To MyObj.MEMBER.METODOS.Count
               nResp = False
               NmMember = "Metodo" & Trim(CStr(i))
               For Each n In MyObj.MEMBER.METODOS
                  If n.NOME = NmMember Then
                     nResp = True
                     Exit For
                  End If
               Next
               If Not nResp Then Exit For
            Next
            MyMet.NOME = NmMember
            MyObj.MEMBER.METODOS.Add MyMet, MyMet.NOME
         End If
         Set MyProp = MyObj.MEMBER.METODOS(NmMember).PagProp
         
      Case tpModulo
      Case tpPropriedade
         If AddNewObj Then
            For i = 1 To MyObj.MEMBER.METODOS.Count
               nResp = False
               For Each n In MyObj.MEMBER.METODOS
                  NmMember = "Metodo" & Trim(CStr(i))
                  If n.NOME = NmMember Then
                     nResp = True
                     Exit For
                  End If
               Next
               If Not nResp Then Exit For
            Next
            MyMet.NOME = NmMember
            MyObj.MEMBER.METODOS.Add MyMet, MyMet.NOME
         End If
         Set MyProp = MyObj.MEMBER.PROPRIEDADES(NmMember).PagProp
         
      
      Case tpVariavel
         If AddNewObj Then
            For i = 1 To MyObj.MEMBER.METODOS.Count
               nResp = False
               For Each n In MyObj.MEMBER.METODOS
                  NmMember = "Metodo" & Trim(CStr(i))
                  If n.NOME = NmMember Then
                     nResp = True
                     Exit For
                  End If
               Next
               If Not nResp Then Exit For
            Next
            MyMet.NOME = NmMember
            MyObj.MEMBER.METODOS.Add MyMet, MyMet.NOME
         End If
         Set MyProp = MyObj.MEMBER.VARIAVEIS(NmMember).PagProp
      
      Case Else:
         '.pTipoPagina = tpNul
   End Select
   If (Img$ = Icon.EnumImage) Or (Img$ = Icon.TypeImage) Then
       'Set MyProp = mvarPROJETOS(ProjetoAtivo).MODULO(mvarVbComp.Name).MEMBER.VARIAVEIS(NmMember).PagProp
       Call mvarPROJETOS(ProjetoAtivo).MODULO(mvarVbComp.Name).MEMBER.VARIAVEIS(NmMember).PagProp.Show
       'MyProp.Show
       Exit Function
   End If
   
   
   With MyProp
      
      Set .PROJETO = mvarProjeto
     
      On Error Resume Next
      'On Error GoTo 0
      .TipoPagina = pTipo
      
      '* Definir Retorno e Argumentos
      
      If .TipoPagina = tpMetodo Then
         If AddNewObj Then
            LinhaInicial = mvarVbComp.CodeModule.CountOfLines + 1
            StrFunc$ = StrFunc$ & "Sub " & NmMember & vbNewLine & vbNewLine & "End Sub "
         Else
            LinhaInicial = mvarVbComp.CodeModule.ProcBodyLine(NmMember, vbext_ProcKind.vbext_pk_Proc)
            If Err = 35 Or LinhaInicial = 0 Then
               LinhaInicial = 1
            End If
            StrFunc$ = mvarVbCode.Lines(LinhaInicial, 1)
         End If
         StrFrase = StrFunc$
         isFunc = (InStr(UCase(StrFunc), "FUNCTION") <> 0)
         If isFunc Then
            .Escopo = Mid(StrFunc, 1, InStr(UCase(StrFunc), "FUNCTION") - 1)
         Else
            .Escopo = Mid(StrFunc, 1, InStr(UCase(StrFunc), "SUB") - 1)
         End If
         If isFunc Then
            .Retorno = "Variant"
         
      '******
            PosRetorno = InStr(UCase(StrFrase), ")")
            While PosAux <> 0
               StrFrase = Trim(Mid(StrFrase, PosRetorno + 1))
               PosRetorno = InStr(UCase(StrFrase), ")")
            Wend
            PosIniArg% = InStr(UCase(StrFunc$), "(") + 1
            PosRetorno = IIf(PosRetorno = PosIniArg%, 0, PosRetorno)
            If PosRetorno > 1 Then
               StrFrase = Mid(StrFunc$, PosIniArg%, PosRetorno - PosIniArg%)
            Else
               StrFrase = ""
            End If
            
            Parenteses = 1
            StrAux$ = Mid(StrFunc$, PosRetorno + 1)
            If StrAux$ <> "" Then
               PosRetorno = InStr(UCase(StrAux$), "AS ")
               If PosRetorno <> 0 Then
                  .Retorno = Trim(Mid(StrAux$, PosRetorno + 3))
               End If
            End If
            Call .SetArgumentos(StrFrase)
         End If
         
      ElseIf .TipoPagina = tpVariavel Then
         .isEnum = (Img$ = Icon.EnumImage)
         .isType = (Img$ = Icon.TypeImage)
         If AddNewObj Then
            LinhaInicial = mvarVbComp.CodeModule.CountOfDeclarationLines + 1
         Else
            LinhaInicial = mvarVbComp.CodeModule.Members(NmMember).CodeLocation
            If Err = 35 Or LinhaInicial = 0 Then
               If .isEnum Or .isType Then
                  Set MyProp = mvarPROJETOS(ProjetoAtivo).MODULO(mvarVbComp.Name).MEMBER.VARIAVEIS(NmMember).PROPPAGE
               End If
               LinhaInicial = 1
            End If
            StrFrase$ = mvarVbCode.Lines(LinhaInicial, 1)
         End If
         .Escopo = RichWordOver(StrFrase$, 0, 0, 1)
         Pos = InStr(LCase(StrFrase$), "as")
         .isConst = (InStr(" " & Trim(LCase(StrFrase$)), " const ") <> 0)
         Pos = InStr(Trim(StrFrase$), " = ")
         If .isConst And Pos <> 0 Then
            .ValorPadrao = Mid(StrFrase$, Pos + 3)
         End If
                     
         Pos = InStr(LCase(StrFrase$), LCase(NmMember))
         sAntes = Trim(Mid(StrFrase$, 1, Pos - 1))
         sDepois = Trim(Mid(StrFrase$, Pos))
         If InStr(sDepois, ",") = 0 Then
            sDepois = ""
         Else
            sDepois = Trim(Mid(sDepois, InStr(sDepois, ",") + 1))
         End If

         If .isConst Then
            .TIPO = ""
         ElseIf .isType Then
            .TIPO = "Type"
         ElseIf .isEnum Then
            .TIPO = "Enum"
         Else
            StrFrase$ = Trim(Mid(StrFrase$, Pos))
            Pos = InStr(LCase(StrFrase$), "as")
            If Pos <> 0 Then
               .TIPO = RichWordOver(StrFrase$, 0, 0, Pos + 2)
            Else
               .TIPO = NameDclImplicit(RichWordOver(StrFrase$, 0, 0, 1))
            End If
            '.TpDeclaracao = RichWordOver(StrFrase$, 0, 0, 1)
         End If
         
      ElseIf .TipoPagina = tpPropriedade Then
         If AddNewObj Then
            LinhaInicial = mvarVbComp.CodeModule.CountOfDeclarationLines + 1
         Else
            LinhaInicial = mvarVbComp.CodeModule.Members(NmMember).CodeLocation
            If Err = 35 Or LinhaInicial = 0 Then LinhaInicial = 1
         End If
      End If
      
      .NOME = NmMember
      If Not AddNewObj Then
         .Descricao = mvarVbMember.Description
         .HelpID = mvarVbMember.HelpContextID
      End If
      
      '************************************
      '**** Exibe Tela de Propriedades ****
      '************************************
      .Show
      
      PropertyPage = .isSaved
      If .isSaved Then
         Select Case pTipo
            Case tpMetodo
               StrArg$ = ""
               StrAux$ = .Escopo
               StrArg$ = .LinhaDeArgumentos
               If isFunc Then
                  If Trim(.Retorno) = "" Then
                     isFunc = (vbNo = ExibirPergunta("Converter 'Function' para 'Sub' ?", "Atenção !"))
                  End If
               Else
                  isFunc = (Trim(.Retorno) <> "")
               End If
            
               nResp = (isFunc And Trim(.Retorno) <> "")
               StrAux$ = .Escopo & " " & IIf(isFunc, "Function ", "Sub ")
               StrAux$ = StrAux$ & .NOME & "(" & StrArg$ & ")"
               StrAux$ = StrAux$ & IIf(nResp, " As " & .Retorno, "")
               If AddNewObj Then
                  StrAux = StrAux & vbNewLine & vbNewLine & IIf(isFunc, "End Function", "End Sub")
               End If
            Case tpVariavel
               sAntes = Trim(StrReplace(" " & sAntes & " ", " Const", " "))
               If .isConst Then
                  If InStr(sAntes, " ") = 0 Then
                     If Trim(sDepois) = "" Then
                        sAntes = .Escopo & " Const"
                     Else
                        sAntes = sAntes & " " & sDepois & IIf(Trim(sDepois) <> "", vbNewLine, "") & .Escopo & " Const"
                     End If
                  Else
                     If Right(sAntes, 1) = "," Then sAntes = Mid(sAntes, 1, Len(sAntes) - 1)
                     sAntes = sAntes & vbNewLine
                  End If
               End If
               Pos = InStr(sAntes, Chr(13))
               Pos = IIf(Pos <= 0, 1, Pos + 2)
               If LCase(RichWordOver(sAntes, 0, 0, Pos)) = LCase(.Escopo) Then
                  StrAux = sAntes & " " & .NOME
               Else
                  If Right(sAntes, 1) = "," Then sAntes = Mid(sAntes, 1, Len(sAntes) - 1)
                  nResp = (Trim(sDepois) <> "" And InStr(sAntes, " ") <> 0)
                  StrAux = sAntes & IIf(nResp, ", " & sDepois, "")
                  StrAux = StrAux & IIf(.isConst, "", " " & sDepois & vbNewLine)
                  StrAux = StrAux & .Escopo & IIf(.isConst, " Const ", " ") & .NOME
               End If
               If .isConst Then
                  StrAux = StrAux & " = " & .ValorPadrao
               Else
                  StrAux = StrAux & " As " & .TIPO
               End If
               If .isEnum Or .isType Then StrAux = ""
            Case Else
               StrAux = ""
         End Select
         If Trim(StrAux) <> "" Then
            If AddNewObj Then
               Call mvarVbComp.CodeModule.InsertLines(LinhaInicial, StrAux)
            Else
               Call mvarVbComp.CodeModule.ReplaceLine(LinhaInicial, StrAux)
            End If
         End If
         Call mvarVbComp.SaveAs(mvarVbComp.FileNames(1))
         Set mvarVbMember = mvarVbComp.CodeModule.Members(NmMember)
         mvarVbMember.Description = .Descricao
         mvarVbMember.HelpContextID = .HelpID
         Call mvarVbComp.SaveAs(mvarVbComp.FileNames(1))
      Else
         If AddNewObj Then
            If .TipoPagina = tpMetodo Then
               MyObj.MEMBER.METODOS.Remove NmMember
            ElseIf .TipoPagina = tpVariavel Then
               MyObj.MEMBER.VARIAVEIS.Remove NmMember
            ElseIf .TipoPagina = tpPropriedade Then
               MyObj.MEMBER.PROPRIEDADES.Remove NmMember
            End If
         End If
      End If
      
   End With
   Set MyObj = Nothing
Saida:

End Function
Public Sub Goto_Projeto()
   Dim LinIni&

   mvarMe.TxtCode.Enabled = False
   DoEvents
   If mvarVbComp.IsDirty Then
      Call TrocarLinhasCodigo
   End If
   mvarVbComp.CodeModule.CodePane.Window.WindowState = 2 '* Maximized
   LinIni = Me.Get_LinIni_Code(mvarVbMember)
   LinIni = IIf(LinIni <= 0, 1, LinIni)
   mvarVbComp.CodeModule.CodePane.TopLine = LinIni
   Call mvarVbComp.CodeModule.CodePane.SetSelection(LinIni, 1, LinIni, 1)
   mvarVbComp.CodeModule.CodePane.Show
   mvarMe.Timer(0).Enabled = True
End Sub
Public Sub Refresh_Tela()
   Dim n As Variant, ItemExcluido As New Collection
   If Sys.Constru.ExibeSubPasta Then
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
   mvarMe.TabComp.TabVisible(5) = True 'Sys.Constru.ExibeSubPasta
   Call PintarFundo(mvarMe, Sys.Proj.FundoTela, mvarMe)
   mvarMe.Refresh
End Sub
Private Sub mvarMe_Active()

   Screen.MousePointer = vbHourglass
   mvarMe.Lbl = " Projeto : " & mvarVbProj.Description
   
   Call CarregaSubItens
  
   mvarMe.PrimeiraVez = False
   Call Refresh_Tela
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_CmbControlClick()
   Dim n As Variant
   Dim lKey As String
   Dim Pos As Integer, PosCmb As Integer
   Dim Tam As Integer
   'lKey = GetParent(mvarMe.TreProj.SelectedItem, 2).Key
   lKey = RichWordOver(mvarMe.TreProj.SelectedItem.Key, 0, 0, 1)
   Tam = Len(lKey)
   If Trim(mvarMe.TreProj.SelectedItem.Text) = "Métodos" Then
      If InStr(lKey, Icon.MetImage) <> 0 Then lKey = Mid(lKey, 1, Tam - Len(Icon.MetImage))
   ElseIf Trim(mvarMe.TreProj.SelectedItem.Text) = "Variáveis" Then
      If InStr(lKey, Icon.VarImage) <> 0 Then lKey = Mid(lKey, 1, Tam - Len(Icon.VarImage))
   ElseIf InStr(lKey, Icon.MetImage) <> 0 Then
      lKey = Mid(lKey, 1, InStr(lKey, Icon.MetImage) - 1)
   ElseIf InStr(lKey, Icon.VarImage) <> 0 Then
      lKey = Mid(lKey, 1, InStr(lKey, Icon.VarImage) - 1)
   End If

'      lKey
'   End If
   
   With mvarMe
      .CmbMember.Clear
      If Trim(mvarMe.TreProj.SelectedItem.Text) <> "Métodos" Then
         .CmbMember.AddItem "(Declarations)"
      End If
      Select Case GetParent(mvarMe.TreProj.SelectedItem, 1).Key
         
         Case "VBForm"
            Select Case .CmbControl.ListIndex
               
               Case 0
                  For Each n In Me.PROJETO(Me.ProjetoAtivo).FORMULARIO(lKey).MEMBER.METODOS
                     Pos = InStr(n.NOME, "_")
                     If Pos <> 0 Then
                        PosCmb = LocalizarCombo(.CmbControl, Mid(n.NOME, 1, Pos - 1), False)
                        If PosCmb >= 0 Then
                          If .CmbControl.List(PosCmb) <> Mid(n.NOME, 1, Pos - 1) Then
                             .CmbMember.AddItem n.NOME
                          End If
                        Else
                           .CmbMember.AddItem n.NOME
                        End If
                     Else
                        .CmbMember.AddItem n.NOME
                     End If
                  Next
               Case Else
                  For Each n In Me.PROJETO(Me.ProjetoAtivo).FORMULARIO(lKey).MEMBER.METODOS
                     Pos = InStr(n.NOME, .CmbControl & "_")
                     If Pos <> 0 Then
                        .CmbMember.AddItem Mid(n.NOME, Len(.CmbControl & "_ "))
                     End If
                  Next
                  If .CmbMember.ListCount > 1 Then .CmbMember.RemoveItem 0
            End Select
         Case "VBModule"
            For Each n In Me.PROJETO(Me.ProjetoAtivo).MODULO(lKey).MEMBER.METODOS
               .CmbMember.AddItem n.NOME
            Next
         Case "VBClass"
            For Each n In Me.PROJETO(Me.ProjetoAtivo).CLASSE(lKey).MEMBER.METODOS
               Pos = InStr(n.NOME, "_")
               PosCmb = -1
               If Pos <> 0 Then
                  PosCmb = LocalizarCombo(.CmbControl, Mid(n.NOME, 1, Pos - 1), False)
               End If
               If .CmbControl.ListIndex = 0 Then
                  If PosCmb < 0 Then
                     .CmbMember.AddItem n.NOME
                  Else
                     If .CmbControl.List(PosCmb) <> Mid(n.NOME, 1, Pos - 1) Then
                        .CmbMember.AddItem n.NOME
                     End If
                  End If
               Else
                  If PosCmb >= 0 Then
                     If .CmbControl.List(.CmbControl.ListIndex) = Mid(n.NOME, 1, Pos - 1) Then
                          .CmbMember.AddItem Mid(n.NOME, Len(.CmbControl & "_ "))
                     End If
                  End If
               End If
            Next
      End Select
      If .CmbMember.ListCount > 0 Then
         .CmbMember.ListIndex = 0
      End If
'      Call ListView_SpecialEffects(.LstItens.hwnd)
      Call ListView_ActiveMove(mvarMe.LstItens.hwnd)
      Call ListView_TrackSelect(mvarMe.LstItens.hwnd)
      Call ListView_FullRowSelect(mvarMe.LstItens.hwnd)
      

      .LstItens.Refresh
   End With
   
End Sub

Private Sub mvarMe_CmbMemberClick()
   Dim Pos As Integer, Linha As Integer
   Dim Func As String
   With mvarMe
      Pos = .TxtCode.Find("Sub " & mvarMe.CmbMember.Text)
      If Pos < 0 Then Pos = .TxtCode.Find("Function " & mvarMe.CmbMember.Text)
      If Pos < 0 Then Pos = .TxtCode.Find("Sub " & .CmbControl & "_" & .CmbMember.Text)
      Pos = IIf(Pos < 0, 0, Pos)
      .TxtCode.SelStart = Len(.TxtCode.Text)
      .TxtCode.SelStart = Pos
      If Pos > 0 Then
         Pos = InStr(Pos, .TxtCode.Text, vbCrLf) + 1
         .TxtCode.SelStart = Pos
      End If
      .TxtCode.Enabled = True
      .TxtCode.Visible = True
      .TxtCode.SetFocus
   End With
End Sub

Public Function SaveResItemToDisk(ByVal iResourceNum As Variant, ByVal sResourceType As String, ByVal sDestFileName As String) As Boolean
    '=============================================
    'Saves a resource item to disk
    'Returns 0 on success, error number on failure
    '=============================================
    
    'Example Call:
    ' iRetVal = SaveResItemToDisk(101, "CUSTOM", "C:\myImage.gif")
    
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    
   
    On Error GoTo SaveResItemToDisk_err
    
    'Retrieve the resource contents (data) into a byte array
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    
    'Get Free File Handle
    iFileNumOut = FreeFile
    
    'Open the output file
    Open sDestFileName For Binary Access Write As #iFileNumOut
        
        'Write the resource to the file
        Put #iFileNumOut, , bytResourceData
    
    'Close the file
    Close #iFileNumOut
    
    'Return True for success
    SaveResItemToDisk = True
    
    Exit Function
SaveResItemToDisk_err:
    'Return error number
    'MsgBox CStr(Err) & " - " & Error
    SaveResItemToDisk = False
End Function
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then
      Call mvarMe_MnuMouseClick("00", mm_EXCLUIR)
   ElseIf KeyCode = vbKeyInsert Then
      Call mvarMe_MnuMouseClick("00", IndMnuMouse.mm_ADICIONAR)
   ElseIf KeyCode = vbKeyReturn Then
      Call mvarMe_MnuMouseClick("00", mm_PROPRIEDADES)
   End If
End Sub
Private Sub mvarMe_Load()
   Dim ArrIco As Variant
   Dim i As Integer, j As Integer
   Screen.MousePointer = vbHourglass
   
   
   Set SysMdi = FrmAddIn
   
   '* Variáveis do Registro do Sistema
   Call GetConfig
   
   '* Configurar controles
   Call Toolbar_MakeFlat(mvarMe.Toolbar)
   Call ResizeCombo(mvarMe.CmbControl, 11, mvarMe.TabComp)
   Call ResizeCombo(mvarMe.CmbMember, 11, mvarMe.TabComp)
   Call SetTVBackColour(mvarMe.TreProj.hwnd, mvarMe.TxtCode.BackColor)
'   Call ListView_SpecialEffects(mvarMe.LstItens.hwnd)
   Call ListView_ActiveMove(mvarMe.LstItens.hwnd)
'   Call ListView_TrackSelect(mvarMe.LstItens.hwnd)
   Call ListView_FullRowSelect(mvarMe.LstItens.hwnd)

   
   mvarMe.LstItens.Refresh
   'If SaveResItemToDisk("FILE_FUNDO_PAREDE", "CUSTOM", App.PATH & "\FUNDOLIST.bmp") Then
   '   Call ListView_SetBackGroung(mvarMe.LstItens.hWnd, App.PATH & "\FUNDOLIST.bmp")
   'End If
   
  '*********
   '* Montar e configurar 'Lay-Out' de Componentes
   With mvarMe
      .PrimeiraVez = True
      .MnuMouse_Main.Visible = False
      .TabComp.Tab = 0
      
      '.Toolbar.ImageList = .ImgListToolbar
      
      '* Montar ImgLstIcons
      With Icon
         ArrIco = Array(.VBPImage, .CloseImage, .OpenImage, .FRMImage, .BASImage, _
                        .CLSImage, .RESImage, .MetImage, .EvtImage, .VarImage, _
                        .ConstImage, .EnumImage, .TypeImage, .PropImage, .MDIImage, _
                        .LibImage) '.CollImage
      End With
      For i = LBound(ArrIco) To UBound(ArrIco)
         .ImgLstIcons.ListImages.Add i + 1, ArrIco(i), LoadResPicture(ArrIco(i), IIf(Mid(ArrIco(i), 1, 3) = "BMP", vbResBitmap, vbResIcon))
      Next
      .TreProj.ImageList = .ImgLstIcons
      .LstItens.SmallIcons = .ImgLstIcons
      .LstItens.Icons = .ImgLstIcons
      
      
      '* ToolBar
'      .Toolbar.ImageList = .ImgLstIcons
'      ArrIco = Array(Icon.ExitImage, Icon.BDWizard, Icon.PropImage)
'      j = 0
'      For i = 0 To UBound(ArrIco)
'         j = j + 1
'         If j <= .Toolbar.Buttons.Count Then
'            While .Toolbar.Buttons(j).Style <> tbrDefault
'               j = j + 1
'            Wend
'            .Toolbar.Buttons(j).Image = ArrIco(i)
'         End If
'      Next
   
      Call MontaArvoreProjeto
      Call ConfigLstItens(0)
      '************
      '* Canfigurar Tamanho
'      .Spl.SplitPercent = 35
'      Call .Spl.SizeControls(.Spl.Left, .TreProj, .TabComp, .Lbl)
      .LstItens.Move .TabComp.Left + 60, .LstItens.Top, .TabComp.Width - 120, .TabComp.Height - 480
      Call ConfigForm(mvarMe, , Sys.Proj.FundoTela, False)
   
      .WindowState = vbMaximized
            
            
   End With
End Sub
Private Sub mvarMe_LstItensClick()
   Dim lKey$
   On ERRO GoTo VerifKey
   If mvarMe.LstItens.SelectedItem Is Nothing Then Exit Sub
   lKey$ = Trim(mvarMe.LstItens.SelectedItem.Text)
   lKey$ = EliminarString(lKey$, "()")
   Set mvarVbMember = mvarVbCode.Members.Item(lKey$)
Saida:
   Exit Sub
VerifKey:
   lKey = RichWordOver(lKey, 0, 0, 1)
   If VarWithDclImplicit(lKey) Then
     lKey = Mid(lKey, 1, Len(lKey) - 1)
     Resume
   Else
      GoTo Fim
   End If
   
Fim:
   Call ShowError
   GoTo Saida
End Sub
Private Sub mvarMe_LstItensDblClick()
   Call mvarMe_MnuMouseClick("00", IndMnuMouse.mm_PROPRIEDADES)
End Sub

Private Sub mvarMe_LstItensMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim i%, lKey$
   If mvarMe.LstItens.SelectedItem Is Nothing Then Exit Sub
   lKey$ = Trim(mvarMe.LstItens.SelectedItem.Text)
   lKey$ = EliminarString(lKey$, "()")
   On Error GoTo VerifKey
   Set mvarVbMember = mvarVbCode.Members.Item(lKey$)
   On Error GoTo Fim
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
            'mvarMe.MnuMouse00(IndMnuMouse.mm_PROPRIEDADES).Visible = False
            mvarMe.MnuMouse00(IndMnuMouse.mm_DEFINIR_COMO).Visible = False
            mvarMe.MnuMouse00(IndMnuMouse.mm_SEP4).Visible = False
         End If
         mvarMe.MnuMouse00(IndMnuMouse.mm_RECARREGAR).Visible = (GetNivel(mvarMe.TreProj.SelectedItem) = 2)
         Call mvarMe.PopupMenu(mvarMe.MnuMouse(0))
   End Select
Saida:
   Exit Sub
VerifKey:
   lKey = RichWordOver(lKey, 0, 0, 1)
   If VarWithDclImplicit(lKey) Then
     lKey = Mid(lKey, 1, Len(lKey) - 1)
     Resume
   Else
      GoTo Fim
   End If
   
Fim:
   Call ShowError
   GoTo Saida
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
   Dim lKey As String, TpOP As String
   Dim sAux As String
   Dim nAux As Integer, nTipo As Integer, nTab As Integer
   Dim bResult As Boolean, Ctrl As Control
   Dim Img As String, NovoNome As String
   
   On Error GoTo Fim
   Select Case mvarMe.ActiveControl.Name
      Case mvarMe.TreProj.Name
         Set Ctrl = mvarMe.TreProj
         lKey$ = mvarMe.TreProj.SelectedItem.Key
         TpOP = "TREE"
         Img$ = mvarMe.TreProj.SelectedItem.Image
         
      Case mvarMe.LstItens.Name
         Set Ctrl = mvarMe.LstItens
         lKey$ = mvarMe.LstItens.SelectedItem.Key
         TpOP = "LIST"
         Img$ = mvarMe.LstItens.SelectedItem.Icon
      
      Case mvarMe.TxtCode.Name
         Set Ctrl = mvarMe.TxtCode
         lKey$ = ""
         TpOP = "TEXT"
         Img$ = ""
   End Select
   Select Case Img
      Case Icon.BASImage:   nTipo = tpModulo
      Case Icon.CLSImage:   nTipo = tpClasse
      Case Icon.CollImage:  nTipo = tpColecao
      Case Icon.ConstImage: nTipo = tpVariavel
      Case Icon.EnumImage:  nTipo = tpVariavel
      Case Icon.EvtImage:   nTipo = tpEvento
      Case Icon.FRMImage:   nTipo = tpForm
      Case Icon.LibImage:
      Case Icon.MDIImage:
      Case Icon.MetImage:   nTipo = tpMetodo
      Case Icon.PropImage:  nTipo = tpPropriedade
      Case Icon.RESImage:
      Case Icon.TypeImage:  nTipo = tpVariavel
      Case Icon.VarImage:   nTipo = tpVariavel
      Case Icon.VBPImage
      Case Else:            nTipo = tpNull
   End Select

   Select Case Menu
      Case "00"
         Select Case Index
            Case mm_ADICIONAR
               If PropertyPage(nTipo) Then
                  Call RecarregarMembro(mvarVbComp.Name)
               End If
            Case mm_ADICIONAR_OBJ
               'select case mvarme.Mnu0000.i
'               If PropertyPage(nTipo) Then
'                  Call RecarregarMembro(mvarVbComp.Name)
'               End If

            Case mm_SALVAR
               If TpOP = "TEXT" Then Call SalvarCodigo
            
            Case mm_Salvar_COMO
               NovoNome = InputBox("Enter the new folder name")
               If TpOP = "TEXT" Then
                   Call SalvarCodigo(NovoNome)
               End If
            
            Case mm_RECARREGAR
               Call RecarregarMembro(lKey$)
            
            Case mm_SEP1          '* 2 - Separador
            
            Case mm_RECORTAR, IndMnuMouse.mm_COPIAR  '* 3, 4 - Recortar / Copiar
               With mvarAreaTemp
                  .NomeObj = ""
                  .TipoObj = ""
                  .TextoObj = ""
                  .TipoOper = IIf(Index = 3, Recortar, Copiar)
               End With
               mvarMe.MnuMouse00(IndMnuMouse.mm_COLAR).Enabled = True
            
            Case mm_COLAR         '* 5 - Colar
            
            Case mm_SEP2          '* 6 - Separador
            
            Case mm_EXCLUIR       '* 7 - Excluir
               If ExcluirObj(Img, lKey) Then
                  Call RecarregarMembro(mvarVbComp.Name)
               End If
               
            Case mm_RENOMEAR      '* 8 - Renomear
               'If Not InArray(mvarMe.TreProj.SelectedItem.Image, Array(Icon.VBPImage, Icon.CloseImage)) Then
               If TpOP = "TREE" Or TpOP = "LIST" Then Ctrl.StartLabelEdit
            
            Case mm_SEP3           '* 9 - Separador
            
            Case mm_PROPRIEDADES   '* 10 - Propriedades
               nAux = mvarMe.LstItens.SelectedItem.Index
               'nItem1 = mvarMe.LstItens.ListItems(CStr(mvarMe.LstItens.GetFirstVisible)).Index
               If Ctrl.Name = mvarMe.TreProj.Name Then
                  lKey$ = mvarMe.TreProj.SelectedItem.Text
               End If
               If PropertyPage(nTipo, lKey) Then
                  If nTipo = tpPropriedade Then nTab = 0
                  If nTipo = tpMetodo Then nTab = 1
                  If nTipo = tpEvento Then nTab = 2
                  If nTipo = tpPropriedade And Icon.PropImage <> Img$ Then nTab = 3
                  
                  If mvarMe.TabComp.Tab = 4 Or mvarMe.TabComp.Tab = nTab Then
                     Call MontaListItens(mvarMe.TabComp.Tab)
                     mvarMe.LstItens.ListItems(nAux).Selected = True
                     mvarMe.LstItens.ListItems(nAux).EnsureVisible
                  End If
               End If

            Case mm_DEFINIR_COMO '* 11 - Definir como Coleção ou Classe
            
            Case mm_SEP4         '* 12 - Separador
            
            Case mm_IR_PROJETO   '* 13 - Ir para o Projeto
               Call Goto_Projeto
         
         End Select
         
      Case "0000" '* Adicionar Objeto
         Select Case Index
            Case 0 '* Classe de Banco de Dados
               If Not DB.Conectado Then Call ConectarBanco
               If DB.Conectado Then FrmMontaCls.Show vbModal
            Case 1 '*Classe
               If PropertyPage(tpClasse) Then
                  Call RecarregarMembro(mvarVbComp.Name)
               End If
            Case 2 '* Coleção
               If PropertyPage(tpColecao) Then
                  Call RecarregarMembro(mvarVbComp.Name)
               End If
            Case 3 '* Propriedade, Variável
               If PropertyPage(tpPropriedade) Then
                  Call RecarregarMembro(mvarVbComp.Name)
               End If
            Case 4 '* Método
               If PropertyPage(tpMetodo) Then
                  Call RecarregarMembro(mvarVbComp.Name)
               End If
            Case 5 '* Evento
               If PropertyPage(tpEvento) Then
                  Call RecarregarMembro(mvarVbComp.Name)
               End If
         End Select
   End Select
Saida:
   Screen.MousePointer = vbDefault
   Exit Sub
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
         .TxtCode.Move 60, .TxtCode.Top, .TabComp.Width - 120, .TabComp.Height - 480 - (.TxtCode.Top - 400)
         .CmbControl.Move 60, .CmbControl.Top, (.TxtCode.Width / 2) - 60
         .CmbMember.Move (.TxtCode.Width / 2) + 60, .CmbMember.Top, (.TxtCode.Width / 2) - 150
'         .TreProj.Move .TreProj.Left, .TreProj.Top, .TabComp.Left - (.TreProj.Left + .Spl.SplitterWidth), .Height - .TreProj.Top - HHStatus
         .Lbl.Move .TreProj.Left, .Lbl.Top, .TreProj.Width
         .Spl.Move .Spl.Left, .Spl.Top, .Spl.Width, .TreProj.Height
      End If
      Call PintarFundo(mvarMe, Sys.Proj.FundoTela)
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
      
         If DB Is Nothing Then Set DB = New DS_BANCO
         If Not DB.Conectado Then Call ConectarBanco
         Screen.MousePointer = vbHourglass
         If DB.Conectado Then FrmMontaCls.Show vbModal
         Screen.MousePointer = vbDefault
      Case "PROPRIEDADES"
         Call mvarMe_MenuClick("0000", 2)
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
   Dim Arq$, i As Integer
   With DB
      .StrConect = ""
      .Alias = ""
      FrmOpBanco.Show vbModal
      If .CodeSql = 1 Then
         .CodeSql = 0
         Exit Sub
      End If
      If .isODBC Then
         If DB.StrConect = "" Then
            frmODBCLog.Show vbModal
         Else
            Call .SrvConecta("", "", .DSN, .UID, .PWD, "")
         End If
      Else
         If Dir("C:\DSR\", vbDirectory) <> "" Then
            mvarMe.CmDialog.InitDir = "C:\DSR\"
         End If
         Arq$ = .Alias
         If .Alias = "" Then
            Arq$ = ProcurarArquivo(SysMdi.CmDialog, "Abrir Banco de Dados Access", , "Microsoft Access MDBs (*.mdb)|*.mdb")
            .dbDrive = SysMdi.CmDialog.Tag
            .dbName = Arq$
         Else
            .dbDrive = GetNameFromPath(Arq, 1)
            .dbName = GetNameFromPath(Arq, 2)
         End If
         If .dbDrive <> "" And .dbName <> "" Then
            Call .SrvConecta(.dbDrive, .dbName, "", "", "", "")
         End If
      End If
      For i = 2 To 5
         Arq$ = Trim(GetSetting(Sys.Constru.AppName, "Outros", "BDRecente" & CStr(i - 1), ""))
         If DB.Alias <> Arq$ And Arq$ <> "" Then
            Call SaveSetting(Sys.Constru.AppName, "Outros", "BDRecente" & CStr(i), Arq$)
         End If
'         Call SaveSetting(Sys.Constru.AppName, "Outros", "BDRecente" & CStr(i), "")
      Next
      If DB.Alias <> "" Then
         Call SaveSetting(Sys.Constru.AppName, "Outros", "BDRecente1", DB.Alias)
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
      .Visible = True
      
      .ColumnHeaders(1).Width = 2100
      .ColumnHeaders(2).Width = 1200
      If .ColumnHeaders.Count = 3 Then
         .ColumnHeaders(3).Width = 3450
      ElseIf .ColumnHeaders.Count = 4 Then
         .ColumnHeaders(2).Width = 1200
         .ColumnHeaders(3).Width = 2400
         .ColumnHeaders(4).Width = 800
      End If
      
      
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
   Dim QtdLin&, LinIni&
   Dim Pos As Integer

   If mvarMe.ActiveControl.Name = mvarMe.TabComp.Name Then
      If mvarMe.TabComp.Tab <> 5 Then Exit Sub
   End If
   Pos = GetNivel(mvarMe.TreProj.SelectedItem)
   If Pos <= 3 Then
      If Pos = 3 Then
         If Trim(mvarMe.TreProj.SelectedItem.Text) = "Variáveis" Then
            LinIni = 1
            QtdLin = mvarVbCode.CountOfDeclarationLines
         ElseIf Trim(mvarMe.TreProj.SelectedItem.Text) = "Métodos" Then
            LinIni = mvarVbCode.CountOfDeclarationLines + 1
            QtdLin = mvarVbCode.CountOfLines - mvarVbCode.CountOfDeclarationLines
         End If
      Else
         LinIni = 1
         QtdLin = mvarVbCode.CountOfLines
      End If
   Else
      LinIni = Get_Lin0_Member(mvarVbMember)
      QtdLin = Get_QtdLin_Code(mvarVbMember) + (Get_LinIni_Code(mvarVbMember) - LinIni)
   End If
   
   '* Se Houver necessidade de Troca
   If Trim(mvarMe.TxtCode.Text) <> Trim(mvarVbCode.Lines(LinIni, QtdLin)) Then
      If PerguntaAntes And Not Sys.Constru.SalvarOnLine Then
         If Not ExibirPergunta("Salvar código do Projeto ?", LoadMsg(1)) = vbYes Then
            Exit Sub
         End If
      End If
      Call mvarVbCode.DeleteLines(LinIni, QtdLin)
      Call mvarVbCode.InsertLines(LinIni, mvarMe.TxtCode.Text)
       
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
Public Function ExcluirObj(pImg As String, pKey As String) As Boolean
   Dim sMsg As String, sLinha As String
   Dim LinIni As Long, QtdLin As Long
   Dim Pos As Integer, sAux As String
   Dim MyObj As Object
   
   ExcluirObj = False
   Select Case pImg
      Case Icon.BASImage:   sMsg = " o Módulo "
      Case Icon.CLSImage:   sMsg = " a Classe "
      Case Icon.CollImage:  sMsg = " a Coleção "
      Case Icon.ConstImage: sMsg = " a Constante "
      Case Icon.EnumImage:  sMsg = " o Enumerado "
      Case Icon.EvtImage:   sMsg = " o Evento "
      Case Icon.FRMImage:   sMsg = " o Formulário "
      Case Icon.LibImage:   sMsg = " a Biblioteca "
      Case Icon.MDIImage:   sMsg = " o Formulário 'MDI' "
      Case Icon.MetImage:   sMsg = " o Método "
      Case Icon.PropImage:  sMsg = " a Propriedade "
      Case Icon.RESImage:   sMsg = " o Arquivo '.Res' "
      Case Icon.TypeImage:  sMsg = " o Tipo "
      Case Icon.VarImage:   sMsg = " a Variável "
      Case Icon.VBPImage:   sMsg = " o Projeto "
   End Select
   sAux = mvarMe.ActiveControl.SelectedItem.Text
   sMsg = "Deseja realmente excluir" & sMsg & "'" & sAux & "' ?"
   If vbYes = ExibirPergunta(sMsg, "Exclusão de " & sAux) Then
      pKey = RichWordOver(sAux, 0, 0)
      ExcluirObj = True
      Set MyObj = GetMember(mvarVbComp.Name)
      Select Case pImg
         Case Icon.BASImage
         Case Icon.CLSImage
         Case Icon.CollImage
'         Case Icon.ConstImage
         Case Icon.EnumImage
         Case Icon.EvtImage
         Case Icon.FRMImage
         Case Icon.LibImage
         Case Icon.MDIImage
         Case Icon.MetImage
            LinIni = mvarVbComp.CodeModule.ProcStartLine(pKey, vbext_ProcKind.vbext_pk_Proc)
            QtdLin = mvarVbComp.CodeModule.ProcCountLines(pKey, vbext_ProcKind.vbext_pk_Proc)
            Call mvarVbComp.CodeModule.DeleteLines(LinIni, QtdLin)
            MyObj.MEMBER.METODOS.Remove pKey
            
         Case Icon.PropImage
         Case Icon.RESImage
         Case Icon.TypeImage
         Case Icon.VarImage, Icon.ConstImage
            LinIni = mvarVbComp.CodeModule.Members(pKey).CodeLocation
            QtdLin = 1
            sLinha = mvarVbComp.CodeModule.Lines(LinIni, QtdLin)
            If InStr(sLinha, ",") <> 0 Then
               Pos = InStr(sLinha, " " & Trim(pKey) & " ")
               sAux = Mid(sLinha, Pos)
               sLinha = Mid(sLinha, 1, Pos)
               Pos = InStr(sAux, ",")
               If Pos <> 0 Then
                  sAux = Trim(Mid(sAux, Pos))
               Else
                  sAux = ""
                  If Right(Trim(sLinha), 1) = "," Then
                     sLinha = Mid(Trim(sLinha), 1, Len(Trim(sLinha)) - 1)
                  End If
               End If
               sLinha = sLinha & Mid(Trim(sAux), 2)
               Call mvarVbComp.CodeModule.ReplaceLine(LinIni, sLinha)
            Else
               Call mvarVbComp.CodeModule.DeleteLines(LinIni, QtdLin)
            End If
            MyObj.MEMBER.VARIAVEIS.Remove pKey
         
         Case Icon.VBPImage
      End Select
      Call mvarVbComp.SaveAs(mvarVbComp.FileNames(1))
   End If
   Set MyObj = Nothing
End Function
Public Function ExibirCodigo() As Integer
   Dim lKey$, VbTexto$, StrLinha$
   Dim Pos&, Tam&, LeftFLood&
   Dim QtdLin&, LinIni&, PosAnt&, PosTexto&
   Dim i As Integer, nAux As Long, sAux As String
   Dim ExibeDiv  As Boolean, ExibeTudo As Boolean
   Dim isVar As Boolean, isMetodo As Boolean
   Dim isVariaveis As Boolean, isMetodos As Boolean
   Dim NodX As Node, n As Variant
   
'   Dim eClasse As Boolean, eFilhoClasse As Boolean
   
   
   On Error GoTo Fim

   mvarMe.LstItens.Visible = False
   
'   eClasse = isClass(mvarMe.TreProj.SelectedItem)
'   eFilhoClasse = isClassChild(mvarMe.TreProj.SelectedItem)
   ExibeDiv = (GetNivel(mvarMe.TreProj.SelectedItem) < 4)
   If GetNivel(mvarMe.TreProj.SelectedItem) = 0 Then
      Exit Function
   Else
      lKey = RichWordOver(mvarMe.TreProj.SelectedItem.Key, 0, 0, 1)
      If InArray(lKey, Array("VBClass", "VBForm", "VBModule", "VBLibrary", "VBRelDoc")) Then
         Exit Function
      End If
   End If
   Tam = Len(lKey)
   If Trim(mvarMe.TreProj.SelectedItem.Text) = "Métodos" Then
      lKey = Mid(lKey, 1, Tam - Len(Icon.MetImage))
      isMetodos = True
   ElseIf Trim(mvarMe.TreProj.SelectedItem.Text) = "Variáveis" Then
      lKey = Mid(lKey, 1, Tam - Len(Icon.VarImage))
      isVariaveis = True
   End If
   
   Select Case mvarMe.TreProj.SelectedItem.Image
      Case Icon.MetImage: Pos = InStr(lKey, Icon.MetImage) - 1
      Case Icon.VarImage: Pos = InStr(lKey, Icon.VarImage) - 1
   End Select
   sAux = Mid(lKey, 1, IIf(Pos <= 0, Len(lKey), Pos))
   Set NodX = mvarMe.TreProj.Nodes(sAux)
   
   If ExibeDiv Then
      If Not isVariaveis And Not isMetodos Then
         On Error GoTo Verif_Key
         sAux = RichWordOver(Trim(mvarMe.TreProj.SelectedItem.Text), 0, 0, 1)
         If sAux <> mvarMe.TreProj.SelectedItem.Text Then
            If sAux = RichWordOver(NodX, 0, 0, 1) Then
               Set mvarVbMember = Nothing
            Else
               Set mvarVbMember = mvarVbCode.Members.Item(sAux)
            End If
         End If
         On Error GoTo Fim
      End If
         
   Else
      sAux = Trim(mvarMe.TreProj.SelectedItem.Text)
      If mvarMe.TreProj.SelectedItem.Image = Icon.CloseImage Then
         sAux = Trim(mvarMe.TreProj.SelectedItem.Child.Text)
      End If
      Set mvarVbMember = mvarVbCode.Members.Item(sAux)
   End If

   If Not mvarVbMember Is Nothing Then
      isVar = (Not isVariaveis And InArray(mvarVbMember.Type, Array(vbext_mt_Const, vbext_mt_Variable)))
      isMetodo = InArray(mvarVbMember.Type, Array(vbext_mt_Event, vbext_mt_Method, vbext_mt_Property))
   End If
   If Err = 0 Then
      PosAnt = mvarMe.TxtCode.SelStart
      
      If LCase(Trim(mvarMe.TreProj.SelectedItem.Text)) = "variáveis" Then
         LinIni = 1
         QtdLin = mvarVbCode.CountOfDeclarationLines
      ElseIf LCase(Trim(mvarMe.TreProj.SelectedItem.Text)) = "métodos" Then
         LinIni = mvarVbCode.CountOfDeclarationLines + 1
         QtdLin = mvarVbCode.CountOfLines - mvarVbCode.CountOfDeclarationLines
      ElseIf Not ExibeTudo Then
         If Not mvarVbMember Is Nothing Then
            LinIni = Get_Lin0_Member(mvarVbMember)
            QtdLin = Get_QtdLin_Code(mvarVbMember) + (Get_LinIni_Code(mvarVbMember) - LinIni)
         End If
      End If
      LinIni = IIf(LinIni = 0, 1, LinIni)
      QtdLin = IIf(QtdLin = 0, mvarVbCode.CountOfLines, QtdLin)
      
      ExibeTudo = (QtdLin >= (mvarVbCode.CountOfLines - mvarVbCode.CountOfDeclarationLines))
      With mvarMe
         If ExibeTudo Then
            Call MontarCmbControl(lKey)
            i = 400
         Else
            i = 0
         End If
         .TxtCode.Move 60, 400 + i, (.TabComp.Width - 120), (.TabComp.Height - (480 + i))
         .CmbControl.Visible = ExibeTudo
         .CmbMember.Visible = ExibeTudo
      End With
      
      '**********
      '* Verifcar se o nó do módulo foi carregado
      If Trim(GetTag(NodX, "RtfText")) = "" Then
         LeftFLood = mvarMe.TabComp.Left + mvarMe.TxtCode.Left
         mvarMe.TxtCode.Visible = False
         mvarMe.TxtCode.Text = mvarVbCode.Lines(1, mvarVbCode.CountOfLines)
         Call ColorizeVbWords(mvarMe.TxtCode, , , , LeftFLood)
         mvarMe.TxtCode.Visible = False
         nAux& = mvarVbCode.CountOfLines - mvarVbCode.CountOfDeclarationLines
         
         nAux& = IIf(nAux& = 0, mvarVbCode.CountOfLines, nAux&)
         VbTexto = mvarVbCode.Lines(mvarVbCode.CountOfDeclarationLines + 1, nAux&)
         Call FindDeadConstants(mvarMe.TxtCode, VbTexto, False)

         Call SetTag(NodX, "RtfText", mvarMe.TxtCode.TextRTF)
         Call SetTag(NodX, "Text", mvarMe.TxtCode.Text)
         mvarMe.TxtCode.Text = ""
         
      End If
      
      lKey = mvarMe.TreProj.SelectedItem.Text
      If Trim(GetTag(mvarMe.TreProj.SelectedItem, "Text")) = Trim(mvarVbCode.Lines(LinIni, QtdLin)) Then
         mvarMe.TxtCode.TextRTF = GetTag(mvarMe.TreProj.SelectedItem, "RtfText")
         GoTo Saida
      ElseIf Trim(GetTag(NodX, "RtfText")) <> "" Then
         mvarMe.RtfTemp.TextRTF = GetTag(NodX, "RtfText")
         Pos = 0
         Tam = 0
         If isVariaveis Then
            Tam = -1
            nAux = 0
         ElseIf isMetodos Then
            nAux = mvarVbCode.CountOfDeclarationLines
         ElseIf isVar Then
            nAux = LinIni - 1
         ElseIf mvarVbMember Is Nothing Then
            nAux = 0
            Tam = InStr(mvarMe.RtfTemp.Text, "Type " & lKey)
            If Tam = 0 Then
               Tam = InStr(mvarMe.RtfTemp.Text, "Enum " & lKey)
            End If
            Do
              Tam = Tam - 1
              If Tam <= 0 Then
                 Exit Do
              End If
            Loop Until (Mid(mvarMe.RtfTemp.Text, Tam, 1) = vbCr) Or Tam = 1
         Else
            nAux = mvarVbCode.ProcStartLine(lKey, vbext_ProcKind.vbext_pk_Proc) - 1
         End If
         
         For i = 1 To nAux
             Tam = InStr(Tam + 1, mvarMe.RtfTemp.Text, vbCrLf)
         Next
         
         mvarMe.RtfTemp.SelStart = 0
         mvarMe.RtfTemp.SelLength = Tam + 1
         mvarMe.RtfTemp.SelText = ""
         
         Tam = 0
         If isVariaveis Then
            Pos = mvarVbCode.CountOfDeclarationLines
            Tam = 0 ' mvarVbCode.CountOfLines - mvarVbCode.CountOfDeclarationLines
         ElseIf isMetodos Then
            Tam = Len(mvarMe.RtfTemp.Text)
            Pos = 0 'mvarVbCode.CountOfLines - mvarVbCode.CountOfDeclarationLines
         ElseIf isVar Then
            Pos = QtdLin
         ElseIf mvarVbMember Is Nothing Then
            Tam = InStr(mvarMe.RtfTemp.Text, "End Enum")
            If Tam > InStr(mvarMe.RtfTemp.Text, "End Type") Or Tam = 0 Then
               Tam = InStr(mvarMe.RtfTemp.Text, "End Type")
            End If
            Tam = IIf(Tam = 0, 0, Tam + 8)
            Pos = 0
         Else
            Pos = mvarVbCode.ProcBodyLine(lKey, vbext_ProcKind.vbext_pk_Proc)
            Pos = Pos - mvarVbCode.ProcStartLine(lKey, vbext_ProcKind.vbext_pk_Proc)
            Pos = Pos + mvarVbCode.ProcCountLines(lKey, vbext_ProcKind.vbext_pk_Proc)
         End If
         
         For i = 1 To Pos
            Tam = InStr(Tam + 1, mvarMe.RtfTemp.Text, vbCrLf)
            If Tam = 0 Then i = Pos
         Next
         '* Se Fim de Arquivo
         If Tam = 0 Then Tam = Len(mvarMe.RtfTemp.Text)

         mvarMe.RtfTemp.SelStart = Tam
         mvarMe.RtfTemp.SelLength = Len(mvarMe.RtfTemp.Text) - (Tam - 1)
         mvarMe.RtfTemp.SelText = ""
         mvarMe.TxtCode.TextRTF = mvarMe.RtfTemp.TextRTF
         If Trim(mvarVbCode.Lines(LinIni, QtdLin)) = Trim(mvarMe.TxtCode.Text) Then
            GoTo Saida
         ElseIf mvarVbMember Is Nothing Then
            GoTo Saida
         Else
            '* Renovar
            mvarMe.TxtCode.Text = mvarVbCode.Lines(LinIni, QtdLin)
         End If
      Else
          mvarMe.TxtCode.Text = mvarVbCode.Lines(LinIni, QtdLin)
      End If
         
      DoEvents
      Screen.MousePointer = vbHourglass
      
'      If ExibeTudo Then mvarMe.TxtCode.Visible = False
      LeftFLood = mvarMe.TabComp.Left + mvarMe.TxtCode.Left
      Call ColorizeVbWords(mvarMe.TxtCode, , , , LeftFLood)
'      If ExibeTudo Then mvarMe.TxtCode.Visible = False
      
      If LinIni <= mvarVbCode.CountOfDeclarationLines Then
         QtdLin = mvarVbCode.CountOfLines - mvarVbCode.CountOfDeclarationLines
         LinIni = mvarVbCode.CountOfDeclarationLines + 1
         QtdLin = IIf(QtdLin = 0, mvarVbCode.CountOfLines, QtdLin)
         VbTexto = mvarVbCode.Lines(LinIni, QtdLin)
         Call FindDeadConstants(mvarMe.TxtCode, VbTexto, False)
      Else
         Call FindDeadConstants(mvarMe.TxtCode, mvarMe.TxtCode.Text)
      End If
'      If ExibeDiv Then
'         ExibeDiv = ExibeDiv
'         For i = 1 To mvarVbCode.Members.Count
'            If mvarVbCode.Members(i).Type = vbext_mt_Method Then
'               Pos = InStr(mvarMe.TxtCode.Text, mvarVbCode.Members(i).Name)
'            End If
'         Next
'      End If
      
      mvarMe.TxtCode.SelStart = PosAnt
      mvarMe.TxtCode.SelLength = 0
      mvarMe.TxtCode.Visible = True
      Screen.MousePointer = vbDefault
   End If
   
Saida:
   If Err <> 0 Then mvarMe.TxtCode = ""
   Call SetTag(mvarMe.TreProj.SelectedItem, "RtfText", mvarMe.TxtCode.TextRTF)
   Call SetTag(mvarMe.TreProj.SelectedItem, "Text", mvarMe.TxtCode.Text)
   mvarMe.TxtCode.SelColor = vbBlack
   mvarMe.LstItens.Visible = False
   mvarMe.TxtCode.Visible = True
   cCodigoAntes = mvarMe.TxtCode.Text
   cLinhaAntes = 0 'mvarMe.TxtCode.GetLineFromChar(mvarMe.TxtCode.SelStart)
   ExibirCodigo = IIf(ExibeTudo, 1, 0)
   Exit Function
Verif_Key:
   If Err = 9 Then '* Subscript out of range
      If GetNivel(mvarMe.TreProj.Nodes(lKey$)) <> 4 Then
         ExibeTudo = True
         'lKey = ""
      End If
      Resume Next
   End If
   GoTo Fim
Fim:
   If Err = 9 Then '*Subscript out of range
      Set mvarVbMember = Nothing
      Resume Next
   ElseIf Err <> 0 Then
      Call ShowError
   End If
   GoTo Saida
End Function
Public Sub MontarCmbControl(pKey As String)
   Dim n As Variant
   With mvarMe
   
      .CmbControl.Clear
      .CmbControl.AddItem "(General)"
      Select Case GetParent(mvarMe.TreProj.SelectedItem, 1).Key
         
         Case "VBForm"
            .CmbControl.AddItem "Form"
            For Each n In Me.PROJETO(Me.ProjetoAtivo).FORMULARIO(pKey).MEMBER.CONTROLES
               .CmbControl.AddItem n.NOME
            Next
             
         Case "VBModule"
'            For Each n In Me.PROJETO(Me.ProjetoAtivo).MODULO(pKey).MEMBER.METODOS
'               .CmbControl.AddItem n.NOME
'            Next
         
         Case "VBClass"
            .CmbControl.AddItem "Class"
            For Each n In Me.PROJETO(Me.ProjetoAtivo).CLASSE(pKey).MEMBER.VARIAVEIS
               If n.isObjWithEvents Then
                  .CmbControl.AddItem n.NOME
               End If
            Next
            
       End Select
      If .CmbControl.ListCount > 0 Then
         .CmbControl.ListIndex = 0
      End If
   End With
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
Public Function Get_QtdLin_Code(pVbMember As MEMBER, Optional Codigo_Puro = True) As Long
   Dim LinIni As Long, QtdLin As Long
   Dim IniProc As Long, Kind As Integer
   On Error GoTo Verif_Proc
   
   If pVbMember.Type = 3 Or pVbMember.Type = 5 Then '* Variável ou Constante
      QtdLin = 1
   Else
      Kind% = vbext_pk_Proc
      If Codigo_Puro Then
         LinIni = mvarVbCode.ProcBodyLine(pVbMember.Name, Kind%)
         IniProc = mvarVbCode.ProcStartLine(pVbMember.Name, Kind%)
      End If
      QtdLin = mvarVbCode.ProcCountLines(pVbMember.Name, Kind%)
   End If
   
   Get_QtdLin_Code = QtdLin - (LinIni - IniProc)
   
Fim:
   Exit Function
Verif_Proc:
   If Err = 35 Then '* Sub or Function not defined
      Kind% = Kind% + 1
      If Kind% > 3 Then GoTo Fim
      Resume
   Else
      GoTo Fim
   End If
End Function
Public Function Get_LinIni_Code(pVbMember As MEMBER) As Long
   Dim Kind%
   On Error GoTo Verif_Proc
   If pVbMember.Type = 3 Or pVbMember.Type = 5 Then '* Variável ou Constante
      Get_LinIni_Code = pVbMember.CodeLocation
   Else
      Kind% = vbext_ProcKind.vbext_pk_Proc
      Get_LinIni_Code = mvarVbCode.ProcBodyLine(pVbMember.Name, Kind%)
   End If
Fim:
   Exit Function
Verif_Proc:
   If Err = 35 Then '* Sub or Function not defined
      Kind% = Kind% + 1
      If Kind% > 3 Then GoTo Fim
      Resume
   Else
      GoTo Fim
   End If
End Function
Public Function Get_Lin0_Member(pVbMember As MEMBER) As Long
   Dim Kind%
   On Error GoTo Verif_Proc
   If pVbMember.Type = 3 Or pVbMember.Type = 5 Then '* Variável ou Constante
      Get_Lin0_Member = pVbMember.CodeLocation
   Else
      Kind% = vbext_ProcKind.vbext_pk_Proc
      Get_Lin0_Member = mvarVbCode.ProcStartLine(pVbMember.Name, Kind%)
   End If
Fim:
   Exit Function
Verif_Proc:
   If Err = 35 Then '* Sub or Function not defined
      Kind% = Kind% + 1
      If Kind% > 3 Then GoTo Fim
      Resume
   Else
      GoTo Fim
   End If
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
       Call SetTag(mvarMe.TreProj.Nodes(sParent), "StartUp", "1")
       Call BoldTreeNode(mvarMe.TreProj, mvarMe.TreProj.Nodes(sParent))
    Else
       Call SetTag(mvarMe.TreProj.Nodes(sParent), "StartUp", "0")
       mvarMe.TreProj.Nodes(sParent).Tag = "|StartUp=0|File=" & sProjectFile '& vbCrLf & strVBProjectText
    End If
    Call SetTag(mvarMe.TreProj.Nodes(sParent), "File", sProjectFile)
    'Call SetTag(mvarMe.TreProj.Nodes(sParent), "FileText", strVBProjectText)
    
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
    Img$ = GetImage(ShortcutTypes.vbwFile, StrNewFolder)
    'Img$ = GetImage(ShortcutTypes.vbwFile, StrNewFolder)
    If Img$ <> "" Then
       frmD.Image = IIf(Img$ = "Misc", Icon.CloseImage, Img$)
    End If
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
    Dim Img As String
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
            Img = GetImage(ShortcutType, strPath)
            If Trim(Img) <> "" Then
               .Image = GetImage(ShortcutType, strPath)
            End If
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
   If Not ImportVBProject(mvarVbProj.FileName) Then
      Call ExibirAviso("Projeto não foi carregado.", "Atenção")
   End If
End Sub
Public Sub MontaArvoreProjeto0()
   Dim n, lNome$, lDESC$, lKey$, lTipo$, i%
   Dim lPai$
   Dim NodX As Node, MyClass As clsCLASSE
   
   mvarMe.TreProj.Nodes.Clear
   Set NodX = mvarMe.TreProj.Nodes.Add(, , mvarVbProj.Name, mvarVbProj, Icon.VBPImage)
   NodX.ExpandedImage = Icon.VBPImage
   NodX.Expanded = True
'   NodX.Sorted = True
   '* Construir árvore do Projeto
   
   '********
   '* Inclui Pasta de Forms
   If mvarProjeto.FORMS.Count > 0 Then  'Not ExisteNo(mvarMe.TreProj, "VBForm")
      Set NodX = mvarMe.TreProj.Nodes.Add(mvarVbProj.Name, tvwChild, "VBForm", "Forms", Icon.CloseImage)
      NodX.ExpandedImage = Icon.OpenImage
      NodX.Sorted = True
   End If
   For Each n In mvarProjeto.FORMS
      Set NodX = mvarMe.TreProj.Nodes.Add("VBForm", tvwChild, n.NOME, n.NOME, IIf(n.MDI, Icon.MDIImage, Icon.FRMImage))
      NodX.Sorted = True
   Next
   
   '********
   '* Inclui Pasta de Módulos
   If mvarProjeto.MODULOS.Count > 0 Then  'Not ExisteNo(mvarMe.TreProj, "VBModule")
      Set NodX = mvarMe.TreProj.Nodes.Add(mvarVbProj.Name, tvwChild, "VBModule", "Modules", Icon.CloseImage)
      NodX.ExpandedImage = Icon.OpenImage
      NodX.Sorted = True
   End If
   For Each n In mvarProjeto.MODULOS
      Set NodX = mvarMe.TreProj.Nodes.Add("VBModule", tvwChild, n.NOME, n.NOME, Icon.BASImage)
      NodX.Sorted = True
   Next
   
   '********
   '* Inclui Pasta de Classes
   If mvarProjeto.CLASSES.Count > 0 Then  'Not ExisteNo(mvarMe.TreProj, "VBClass")
      Set NodX = mvarMe.TreProj.Nodes.Add(mvarVbProj.Name, tvwChild, "VBClass", "Class Modules", Icon.CloseImage)
      NodX.ExpandedImage = Icon.OpenImage
      NodX.Sorted = True
   End If
   
   For Each n In mvarProjeto.CLASSES
      Set NodX = mvarMe.TreProj.Nodes.Add("VBClass", tvwChild, n.NOME, n.NOME, Icon.CLSImage)
      NodX.Sorted = True
   Next
   '********
   '* Monta Estrutura de Sub-Classes
   For Each n In mvarProjeto.CLASSES
      Set MyClass = n
      i = 0
      lPai$ = mvarVbProj.VBComponents(MyClass.NOME).ReadProperty("Member0")
      Do While lPai$ <> ""
         If ExisteNo(mvarMe.TreProj, lPai$) Then
            MyClass.SubClasses.Add mvarProjeto.CLASSES(lPai$)
         End If
         i = i + 1
         lPai$ = mvarVbProj.VBComponents(n.NOME).ReadProperty("Member" & Trim(CStr(i)))
      Loop
   Next
'   mvarMe.TreProj.Nodes("VBClass").Expanded = True
'   On Error Resume Next
   For Each n In mvarProjeto.CLASSES
      Set MyClass = n
      With MyClass
          
         If mvarVbProj.VBComponents(.NOME).ReadProperty("Top_Level") = "Yes" Then
            Call MontaSubClasse(mvarMe.TreProj, MyClass)
         End If
      End With
   Next
   
   '********
   '* Inclui Pasta de ResFile
   If mvarProjeto.RESFILE <> "" Then  'Not ExisteNo(mvarMe.TreProj, "VBRelDoc")
      Set NodX = mvarMe.TreProj.Nodes.Add(mvarVbProj.Name, tvwChild, "VBRelDoc", "Related Documents", Icon.CloseImage)
      NodX.ExpandedImage = Icon.OpenImage
      NodX.Sorted = True
      
      lNome = mvarProjeto.RESFILE  '"(" & GetNameFromPath(n.FileNames(1)) & ")"
      Set NodX = mvarMe.TreProj.Nodes.Add("VBRelDoc", tvwChild, lNome, lNome, Icon.RESImage)
      NodX.Sorted = True

   End If
   'For Each n In mvarPROJETO.RELDOCS
   '   Set NodX = mvarMe.TreProj.Nodes.Add("VBClass", tvwChild, n.Name, n.Name, "RELDOC")
   '   NodX.Sorted = True
   'Next
End Sub
Public Sub MontaSubClass()
   Dim n, i%, lPai$
   Dim MyClass As clsCLASSE

   '********
   '* Monta Estrutura de Sub-Classes
   For Each n In mvarProjeto.CLASSES
      Set MyClass = n
      i = 0
      lPai$ = mvarVbProj.VBComponents(MyClass.NOME).ReadProperty("Member0")
      Do While lPai$ <> ""
         If ExisteNo(mvarMe.TreProj, lPai$) Then
            MyClass.SubClasses.Add mvarProjeto.CLASSES(lPai$)
         End If
         i = i + 1
         lPai$ = mvarVbProj.VBComponents(n.NOME).ReadProperty("Member" & Trim(CStr(i)))
      Loop
   Next
   For Each n In mvarProjeto.CLASSES
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
   Dim eClasse As Boolean, eFilhoClasse As Boolean
   Dim i%, Pos%
   Dim n As Variant, n1 As Variant
   On Error GoTo Fim
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
      Set mvarVbComp = mvarVbProj.VBComponents.Item(lComp$)
      If Not MyObj.MEMBER.Carregado Then
         MyObj.LINHAS = mvarVbComp.CodeModule.CountOfLines
         Call CarregaMembers(MyObj.PATH & MyObj.FileName, MyObj)
         i = 0
         While mvarMe.TreProj.Nodes(lComp$).Children > i
             If mvarMe.TreProj.Nodes(lComp$).Child.Image <> Icon.CLSImage Then
                mvarMe.TreProj.Nodes.Remove mvarMe.TreProj.Nodes(lComp$).Child.Key
             Else
                i = i + 1
             End If
         Wend
      End If
   
      If Not Sys.Constru.ExibeSubPasta Then GoTo Fim
      If MyObj.MEMBER.Carregado And mvarMe.TreProj.Nodes(lComp0$).Children <> 0 Then
         Set NodX = mvarMe.TreProj.Nodes(lComp0$).Child
         For i = 1 To mvarMe.TreProj.Nodes(lComp0$).Children
            If (NodX.Image = Icon.CloseImage Or NodX.Image = Icon.OpenImage) Then
               GoTo Fim
            End If
            Set NodX = NodX.Next
         Next
      End If
      
      With MyObj.MEMBER
         
         '* EVENTOS
         If .EVENTOS.Count > 0 Then
            'If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Eventos
            lPai$ = lComp0$ & Icon.EvtImage
            Set NodX = mvarMe.TreProj.Nodes.Add(lComp0$, tvwChild, lPai$, " Eventos", Icon.CloseImage)
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
            lPai$ = lComp0$ & Icon.MetImage
            Set NodX = mvarMe.TreProj.Nodes.Add(lComp0$, tvwChild, lPai$, " Métodos", Icon.CloseImage)
            NodX.ExpandedImage = Icon.OpenImage
            NodX.Sorted = True
            'End If
            For Each n In .FUNCOES
               lImage$ = Icon.MetImage
               lKey = lPai$ & CStr(mvarMe.TreProj.Nodes(lPai).Children + 1)
               Set NodX = mvarMe.TreProj.Nodes.Add(lPai$, tvwChild, lKey, n.NOME, lImage$)
               NodX.Sorted = True
            Next
         End If
         
         '* METODOS
         
         If .METODOS.Count > 0 Then
            'If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Métodos
            lPai$ = lComp0$ & Icon.MetImage
            Set NodX = mvarMe.TreProj.Nodes.Add(lComp0$, tvwChild, lPai$, " Métodos", Icon.CloseImage)
            NodX.ExpandedImage = Icon.OpenImage
            NodX.Sorted = True
            'End If
            For Each n In .METODOS
               lImage$ = Icon.MetImage
               lKey = lPai$ & CStr(mvarMe.TreProj.Nodes(lPai).Children + 1)
               Set NodX = mvarMe.TreProj.Nodes.Add(lPai$, tvwChild, lKey, n.NOME, lImage$)
               NodX.Sorted = True
            Next
         End If
         
         '* VARIAVEL
         If .VARIAVEIS.Count > 0 Then
            'If Not ExisteNo(mvarMe.TreProj, lPai$) Then  '* Inclui Pasta de Eventos
            lPai$ = lComp0$ & Icon.VarImage
            Set NodX = mvarMe.TreProj.Nodes.Add(lComp0$, tvwChild, lPai$, " Variáveis", Icon.CloseImage)
            NodX.ExpandedImage = Icon.OpenImage
            NodX.Sorted = True
            'End If
            For Each n In .VARIAVEIS
               If LCase(n.TIPO) = "enum" Then
                  lImage$ = Icon.EnumImage
               ElseIf LCase(n.TIPO) = "type" Then
                  lImage$ = Icon.TypeImage
               Else
                 lImage$ = IIf(n.isConst, Icon.ConstImage, Icon.VarImage)
               End If
               lKey = lPai$ & CStr(mvarMe.TreProj.Nodes(lPai).Children + 1)
               Set NodX = mvarMe.TreProj.Nodes.Add(lPai$, tvwChild, lKey, n.NOME, lImage$)
               NodX.Sorted = True
            Next
         End If
         '* Bibliotecas
         If .LIBS.Count > 0 Then
            If Not ExisteNo(mvarMe.TreProj, "VBLibrary") Then
               Set NodX = mvarMe.TreProj.Nodes.Add(mvarMe.TreProj.Nodes(1).Key, tvwChild, "VBLibrary", "Library", Icon.CloseImage)
            End If
            For Each n In .LIBS
               If Not ExisteNo(mvarMe.TreProj, n.NOME) Then
                  Set NodX = mvarMe.TreProj.Nodes.Add("VBLibrary", tvwChild, n.NOME, n.NOME, Icon.LibImage)
               End If
               'NodX.ExpandedImage = Icon.OpenImage
               NodX.Sorted = True
               For Each n1 In n.FUNCOES
                  lImage$ = Icon.MetImage
                  lKey = n1.NOME
                  Set NodX = mvarMe.TreProj.Nodes.Add(n.NOME, tvwChild, lKey, lKey, lImage$)
                  NodX.Sorted = True
               Next
            Next
         End If
         
      End With
   End If
   Exit Sub
Fim:
   If Err = 35602 Then '* Key is not unique in collection
      Err = 0
      Resume Next
   ElseIf Err <> 0 Then
      ShowError
   End If
End Sub
Public Sub MontaSubClasse(Tree As Object, pClass As clsCLASSE, Optional SubClasse)
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
   Dim lKey As String
   Dim MesmoCampo As Boolean
      
   On Error GoTo Fim

   '* Definir Chave do Component
   If mvarVbComp Is Nothing Then GoTo Saida
   lComp$ = mvarVbComp.Name
   MesmoCampo = (TreIndAnterior = mvarMe.TreProj.SelectedItem.Index)
   MesmoCampo = ((MesmoCampo) And (cCompAtual = lComp))
   
   If MesmoCampo And Trim(cCompAtual) <> "" And _
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
            lKey = n.NOME
            Set ItemX = mvarMe.LstItens.ListItems.Add(, lKey, n.NOME, Icon.PropImage, Icon.PropImage)
         Next
      
      Case 1 '* CompType = vbext_mt_Method
         i = 0
         Dim MyMetodo As clsMETODO
         Dim StrAux$
         Dim Pos%
         For Each n In MyObj.MEMBER.METODOS
            Set MyMetodo = n
            i = i + 1
            With MyMetodo
               lKey = .NOME
               Set ItemX = mvarMe.LstItens.ListItems.Add(, lKey, .NOME, Icon.MetImage, Icon.MetImage)
               If .PagProp Is Nothing Then
                  Set .PagProp = New PROPPAGE
                  Call GetPropPage(.PagProp, mvarVbComp, .NOME)
               End If
               With .PagProp
                  ItemX.SubItems(1) = .Retorno
                  ItemX.SubItems(2) = .LinhaDeArgumentos
               End With
            End With
            Set MyMetodo = Nothing
         Next
      
      Case 2 '* CompType = vbext_mt_Event
         For Each n In MyObj.MEMBER.EVENTOS
            lKey = n.NOME
            Set ItemX = mvarMe.LstItens.ListItems.Add(, lKey, n.NOME, Icon.EvtImage, Icon.EvtImage)
         Next
      
      Case 3 '* CompType = vbext_mt_Variable
         For Each n In MyObj.MEMBER.VARIAVEIS
            If LCase(n.TIPO) = "enum" Then
               lImage$ = Icon.EnumImage
            ElseIf LCase(n.TIPO) = "type" Then
               lImage$ = Icon.TypeImage
            Else
              lImage$ = IIf(n.isConst, Icon.ConstImage, Icon.VarImage)
            End If
            lKey = n.NOME
            Set ItemX = mvarMe.LstItens.ListItems.Add(, lKey, n.NOME, lImage$, lImage$)
            ItemX.SubItems(1) = n.TIPO
            ItemX.SubItems(IIf(mvarMe.TabComp.Tab = 4, 3, 2)) = n.Escopo
         Next
      
      Case 4
         mvarMe.LstItens.ListItems.Clear
         Call MontaListItens(0)
         Call MontaListItens(1)
         Call MontaListItens(2)
         Call MontaListItens(3)
      Case 5 '* Código
         Call ExibirCodigo
         mvarMe.TxtCode.SelStart = 0
         mvarMe.TreProj.SetFocus
         Exit Sub
   End Select
'   Call ListView_SpecialEffects(mvarMe.LstItens.hwnd)
   Call ListView_ActiveMove(mvarMe.LstItens.hwnd)
'   Call ListView_TrackSelect(mvarMe.LstItens.hwnd)
   Call ListView_FullRowSelect(mvarMe.LstItens.hwnd)
   mvarMe.LstItens.Refresh
Saida:
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
      Case Else:              DefineImage = Icon.CloseImage
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
         mvarMe.MnuMouse00(IndMnuMouse.mm_RECARREGAR).Visible = (GetNivel(mvarMe.TreProj.SelectedItem) = 2)
         Call mvarMe.PopupMenu(mvarMe.MnuMouse(0))
   End Select
   DoEvents
End Sub

Private Sub mvarMe_TreProjNodeClick(Node As Variant)
   Screen.MousePointer = vbHourglass
   If Node Is Nothing Then GoTo Fim
   If Node.Image = Icon.VBPImage Then GoTo Fim
   Call DefineMembro
   Call MontaSubItens(Node.Key)
   Call Rel_Tab_Tree
   Call MontaListItens
   TreIndAnterior = mvarMe.TreProj.SelectedItem.Index
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
'   If InArray(mvarMe.TreProj.SelectedItem.Key, Array("VBForm", "VBModule", "VBClass", "VB")) Then
'      GoTo Saida
'   End If
   Select Case GetParent(mvarMe.TreProj.Nodes(NomeComponente), 1).Key
      Case "VBForm": Set MyObj = mvarProjeto.FORMULARIO(NomeComponente)
      Case "VBModule":  Set MyObj = mvarProjeto.MODULO(NomeComponente)
      Case "VBClass"
         If NomeComponente <> "VBClass" Then
            Pos% = InStr(NomeComponente, Chr(255))
            Pos = IIf(Pos = 0, Len(NomeComponente) + 1, Pos)
            NomeComponente = Mid(NomeComponente, 1, Pos - 1)
            Set MyObj = mvarProjeto.CLASSE(NomeComponente)
         End If
      Case "VB":  Set MyObj = mvarProjeto.FORMULARIO(NomeComponente)
   End Select
   Set GetMember = MyObj
Saida:
   Exit Function
Fim:
   ShowError
End Function

Private Sub mvarMe_TxtCodeClick(IsMouseClick As Boolean)
   Dim LinhaAtual As Long
   Dim PosIni As Long
   Dim PosFim As Long
   If IsMouseClick Then
      Dim LinIni&
'      Call TrocarLinhasCodigo
'      mvarVbComp.CodeModule.CodePane.Window.WindowState = 2 '* Maximized
'      LinIni = Me.Get_LinIni_Code(mvarVbMember)
'      mvarVbComp.CodeModule.CodePane.TopLine = LinIni
'      Call mvarVbComp.CodeModule.CodePane.SetSelection(LinIni, 1, LinIni, 1)
'      mvarVbComp.CodeModule.CodePane.Show
'      Exit Sub
   End If
   LinhaAtual = mvarMe.TxtCode.GetLineFromChar(mvarMe.TxtCode.SelStart)
   If cLinhaAntes <> LinhaAtual Then
      If mvarMe.TxtCode.Text <> cCodigoAntes Then
         PosIni = 1
         While cLinhaAntes <> mvarMe.TxtCode.GetLineFromChar(PosIni)
           PosIni = PosIni + 1
         Wend
         
         PosFim = PosIni
         While cLinhaAntes = mvarMe.TxtCode.GetLineFromChar(PosFim) And PosFim < Len(mvarMe.TxtCode.Text)
           PosFim = PosFim + 1
         Wend
         Call ColorizeVbWords(mvarMe.TxtCode, PosIni, PosFim)
'         Call FindDeadConstants(mvarMe.TxtCode, mvarMe.TxtCode.Text)
      End If
   End If
   cLinhaAntes = mvarMe.TxtCode.GetLineFromChar(mvarMe.TxtCode.SelStart)
   cCodigoAntes = mvarMe.TxtCode.Text
End Sub

Private Sub mvarMe_TxtCodeKeyUp(KeyCode As Integer, Shift As Integer)
   If mvarMe.TxtCode.SelLength < 1 Then
      mvarMe.TxtCode.SelColor = RGB(0, 0, 0)
   End If
   Select Case KeyCode
      Case vbKeyReturn, vbKeyDown, vbKeyUp, vbKeyEnd, vbKeyHome
         Call mvarMe_TxtCodeClick(False)
   End Select
End Sub

Private Sub mvarMe_TxtCodeMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim i%
   Dim Pos As Long, Linha As Integer
   Dim StrFunc As String
   
   Select Case Button
      Case 1
      Case 2
         '* O menu não pode ter todos os seu itens invisiveis
         '* Logo temos que definir pelo menos um item visivel
         '* apenas leitura.
         Linha = 0
         Pos = GetNivel(mvarMe.TreProj.SelectedItem)
         If Pos <= 3 Then
            If Pos = 3 Then
               If Trim(mvarMe.TreProj.SelectedItem.Text) = "Variáveis" Then
                  Linha = 0
               ElseIf Trim(mvarMe.TreProj.SelectedItem.Text) = "Métodos" Then
                  Linha = mvarVbCode.CountOfDeclarationLines
              End If
            End If
            
            Pos = RichMousePos(mvarMe.TxtCode.hwnd, x, Y)
            StrFunc = mvarVbComp.CodeModule.ProcOfLine(mvarMe.TxtCode.GetLineFromChar(Pos) + Linha, vbext_ProcKind.vbext_pk_Proc)
            If Trim(StrFunc) = "" Then
               Set mvarVbMember = Nothing
            Else
               Set mvarVbMember = mvarVbCode.Members.Item(StrFunc)
            End If
         End If
         
         mvarMe.MnuMouse00(IndMnuMouse.mm_IR_PROJETO).Visible = True
         For i = mvarMe.MnuMouse00.LBound To mvarMe.MnuMouse00.UBound
            If i <> IndMnuMouse.mm_IR_PROJETO Then mvarMe.MnuMouse00(i).Visible = False
         Next
         mvarMe.MnuMouse00(IndMnuMouse.mm_SALVAR).Visible = Not Sys.Constru.SalvarOnLine
         Call mvarMe.PopupMenu(mvarMe.MnuMouse(IndMnuMouse.mm_ADICIONAR))
   End Select
End Sub
Public Function GetImage(ByVal ShortcutType As ShortcutTypes, ByVal strPath As String) As String
   Select Case ShortcutType
      Case 0 'vbwFile
         '// set the image
         GetImage = SetIcon(strPath)
      Case 1 'vbwFolder
         If IsDrive(strPath) Then
            GetImage = "Drive"
         Else
            GetImage = "FolderShortcut"
         End If
      Case 2 'vbwURL
         GetImage = "URL"
      Case 3 'vbwEMail
         GetImage = "Mail"
   End Select
End Function
Public Function SetIcon(ByVal strFileName As String) As String
   Dim strExtension As String
   '// Get the extension
   strExtension = GetExtension(strFileName)
   '// What extension?
   Select Case LCase$(strExtension)
      Case "vbp"
         SetIcon = Icon.VBPImage
      Case "bas"
         SetIcon = Icon.BASImage
      Case "cls"
         SetIcon = Icon.CLSImage
      Case "frm"
         SetIcon = Icon.FRMImage
      Case "res"
         SetIcon = Icon.RESImage
      
      
      Case "htm", "html"
         SetIcon = "Web"
      Case "vbg"
         SetIcon = "Root"
         SetIcon = UCase$(strExtension)
      Case "ctl", "txt", "dob", "pag", "ini", "bat" ', "exe", "c", "cpp", "cxx", "tli", "h", "tlh", "inl", "rc"
         SetIcon = UCase$(strExtension)
         SetIcon = ""
      Case Else
         SetIcon = Icon.CloseImage
   End Select
End Function
Public Sub CarregaSubItens()
   Dim i%, Tot&
   Dim n As Variant, NmComp$
   
   Tot = mvarMe.TreProj.Nodes.Count
   If mvarMe.PrimeiraVez And Sys.Constru.LoadIni Then
      mvarMe.PrimeiraVez = False
      'mvarMe.TreProj.Visible = False
      For i = 1 To Tot
         NmComp$ = mvarMe.TreProj.Nodes(i).Key
         Call MontaSubItens(NmComp$)
         If Not AtuFlood(Int(i / Tot * 100), , "Carregando Projeto...") Then
            Exit For
         End If
      Next
      Call FimFlood
   End If
End Sub

