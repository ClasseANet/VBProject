VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DESCPROJ"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"PROJETO"
Private mvarPROJETOS As New Collection 'local copy
Private mvarPathProj As String 'local copy
Private mvarFileProj As String 'local copy
Private mvarPROJETO As New PROJETO
Private mvarfMe As Object 'local copy
Public Property Set fMe(ByVal vData As Object)
    Set mvarfMe = vData
End Property
Public Property Get fMe() As Object
    Set fMe = mvarfMe
End Property
Public Property Get PROJETO() As PROJETO
    If mvarPROJETO Is Nothing Then
        Set mvarPROJETO = New PROJETO
    End If
    Set PROJETO = mvarPROJETO
End Property
Public Property Set PROJETO(vData As PROJETO)
    Set mvarPROJETO = vData
End Property
Private Sub Class_Terminate()
    Set mvarPROJETO = Nothing
End Sub
Public Property Let FileProj(ByVal vData As String)
    mvarFileProj = vData
End Property
Public Property Get FileProj() As String
    FileProj = mvarFileProj
End Property
Public Property Let PathProj(ByVal vData As String)
    mvarPathProj = vData
End Property
Public Property Get PathProj() As String
    PathProj = mvarPathProj
End Property
Public Property Set PROJETOS(ByVal vData As Object)
    Set mvarPROJETOS = vData
End Property
Public Property Get PROJETOS() As Collection
    Set PROJETOS = mvarPROJETOS
End Property
Public Sub CarregaForm(ByRef lForm As FORMULARIO, TxtLin$)
   Dim ColArq  As New Collection
   Set ColArq = DefineFullPath(TxtLin$)
   lForm.FILENAME = ColArq("FILE")
   lForm.NOME = ColArq("FILE")
   'lForm.TAMANHO = FileLen(ColArq("PATH") & ColArq("FILE"))
End Sub
Public Function CarregaFuncao(ByRef pFunc As FUNCAO, TxtLin$) As Boolean
   Dim Txt$, Pos%, nArq%
   Dim m, n
   Dim ColArq  As New Collection
   Dim MyFunc As New FUNCAO
   Dim MyProj As New PROJETO
   FuncSub$ = "SUB "
   Pos% = InStr(TxtLin$, FuncSub$)
   pFunc.isFuncao = (Pos = 0)
   If pFunc.isFuncao Then
      FuncSub$ = "FUNCTION "
      Pos% = InStr(TxtLin$, FuncSub$)
   End If
   '* Recuperar Nome
   pFunc.NOME = GetWord(Mid(TxtLin$, Pos + Len(FuncSub$)))
   pFunc.Tipo = UCase(GetTipoSimbol(pFunc.NOME))
   If pFunc.Tipo <> "" Then
      pFunc.NOME = Mid(pFunc.NOME, 1, InStr(pFunc.NOME, RightB$(pFunc.NOME, 2)) - 1)
   End If
   '* Verificar se o Objeto já foi carregado
   For Each m In mvarPROJETOS
      Set MyProj = m
      If MyProj.FUNCOES Is Nothing Then Exit For
      For Each n In MyProj.FUNCOES
        Set MyFunc = n
        If pFunc.NOME = MyFunc.NOME Then
           Set pFunc = MyFunc
           Set MyFunc = Nothing
           Set MyProj = Nothing
           Exit Function
        End If
        Set MyFunc = Nothing
     Next
     Set MyProj = Nothing
   Next
   CarregaFuncao = True
   pFunc.ESCOPO = GetWord(TxtLin$)
   If pFunc.Tipo = "" And pFunc.isFuncao Then
      Txt = UCase(TxtLin$)
      Pos = InStr(Txt, " AS ")
      While Pos <> 0
         Txt = Trim(Mid(Txt, Pos + 4))
         Pos = InStr(Txt, " AS ")
      Wend
      If Txt <> "" And (InStr(Txt, ")") = 0) Then
        pFunc.Tipo = Txt
      End If
   End If
   pFunc.LINHAS = 1
   If pFunc.ESCOPO = pFunc.NOME Then pFunc.ESCOPO = "Private"
   '* Definir Parametros
   If pFunc.isFuncao Then
      Call GetParam(pFunc, GetWordInPar(TxtLin$))
   End If

   If InStr(TxtLin$, "DECLARE") <> 0 Then
      'mvarPROJETO.FUNCOES.Add pFunc, pFunc.NOME
      Exit Function
   End If
   
   nArq = FreeFile - 1
   Do While UCase(TxtLin) <> Trim("END " & FuncSub$)
      Line Input #nArq, TxtLin$
      Call AnalisarLinha_Fun(TxtLin$, pFunc)
      If TxtLin$ <> "" Then pFunc.LINHAS = CStr(Val(pFunc.LINHAS) + 1)
   Loop
   mvarPROJETO.FUNCOES.Add pFunc
End Function
Public Function CarregaModulo(ByRef pBAS As Modulo, TxtLin$)
   Dim TextLine$, Cont%, nArq%
   Dim ColArq  As New Collection
   Dim MyModulo As New Modulo
   Dim MyProj As New PROJETO
   Dim MyBas As New Modulo
   Dim m, n
   '* Recuperar Nome
   Pos% = InStr(TxtLin$, ";")
   Set ColArq = DefineVarValor(UCase(Mid(TxtLin$, 1, Pos - 1)))
   pBAS.NOME = ColArq("VALOR")
   '* Verificar se o Objeto já foi carregado
   For Each m In mvarPROJETOS
     Set MyProj = m
     For Each n In MyProj.MODULOS
        Set MyBas = n
        If pBAS.NOME = MyBas.NOME Then
           Set pBAS = MyBas
           Set MyBas = Nothing
           Set MyProj = Nothing
           Exit Function
        End If
        Set MyBas = Nothing
     Next
     Set MyProj = Nothing
   Next
   Set ColArq = DefineFullPath(Mid(TxtLin$, InStr(TxtLin$, ";") + 1))
   pBAS.PATH = ColArq("PATH")
   nArq = FreeFile
   Open ColArq("PATH") & ColArq("FILE") For Input As #nArq
   Do While Not EOF(nArq)
      Line Input #nArq, TextLine
      Call AnalisarLinha_Bas(TextLine, pBAS)
      If TextLine <> "" Then pBAS.LINHAS = CStr(Val(pBAS.LINHAS) + 1)
   Loop
   Close #nArq ' Close file.
   CarregaModulo = True
End Function
Public Sub CarregaProjeto(pPATH$, pARQ$)
   Dim TextLine$, Cont%, nArq%
   Dim PathAux$, JaExiste As Boolean, n
   Dim ColVar As New Collection
   Dim pProj$
   pProj$ = pPATH$ & pARQ$
   nArq = FreeFile
   
   Open pProj$ For Input As #nArq
   Do While Not EOF(nArq)
      Line Input #nArq, TextLine
      If UCase(Mid(pProj$, InStr(pProj$, "."))) = ".VBG" Then
         PathAux = mvarPathProj
         Call AnalisarLinha_Vbg(TextLine)
         mvarPathProj = PathAux
      Else
         If TextLine <> "" Then
            If Not JaExiste Then
               Set mvarPROJETO = Nothing
               mvarPROJETO.FILENAME = pARQ$
               mvarPROJETOS.Add mvarPROJETO, mvarPROJETO.FILENAME
               JaExiste = True
            End If
            Call AnalisarLinha_Vbp(TextLine, mvarPROJETO)
            mvarPROJETO.LINHAS = CStr(Val(mvarPROJETO.LINHAS) + 1)
         End If
      End If
   Loop
'   If mvarPROJETO.NOME <> "" Then
'      For Each n In mvarPROJETOS
'         JaExiste = (n.NOME = mvarPROJETO.NOME)
'      Next
'       If Not JaExiste Then
'         mvarPROJETOS.Add mvarPROJETO, mvarPROJETO.NOME
'         Set mvarPROJETO = Nothing
'      End If
'   End If
   Close #nArq ' Close file.
End Sub
Public Sub MontaTreeProjeto()
   Dim Tree As TreeView, nodX As Object
   Dim IdRoot$, DscRoot$
   Dim Cod$, Dsc$
   Dim MyProj As New PROJETO
   Dim Cont%, PROJ, n
   
   Set Tree = fMe.TreProj
   Set nodX = Tree.Nodes.Add(, , "Root", "SISTEMA", "PROJ") ' Root
   nodX.Expanded = True
   nodX.Sorted = True
   For Each PROJ In mvarPROJETOS
      Cont% = Cont% + 1
      Set MyProj = PROJ
      IdRoot = MyProj.FILENAME
      DscRoot = MyProj.NOME
      On Error Resume Next
      Set nodX = Tree.Nodes.Add("Root", 4, IdRoot, DscRoot, "PROJ")  ' Root
      'nodX.Sorted = True
      '* se ja existe deve-se excluir o projeto e carregálo de novo
      If Err = 35602 Then 'Key is not unique in collection
         Tree.Nodes.Remove Tree.Nodes(IdRoot).Index
         Err = 0
         Set nodX = Tree.Nodes.Add("Root", 4, IdRoot, DscRoot, "PROJ")  ' Root
         'nodX.Sorted = True
      End If
      Call TreeInSubClass(Tree, IdRoot, "Forms", CStr(Cont%), MyProj.FORMS, "FORM")
      Call TreeInSubClass(Tree, IdRoot, "Modulos", CStr(Cont%), MyProj.MODULOS, "MODULO")
      Call TreeInSubClass(Tree, IdRoot, "Classes", CStr(Cont%), MyProj.CLASSES, "CLASSE")
      If MyProj.RESFILE <> "" Then
         Set nodX = Tree.Nodes.Add(IdRoot$, 4, "ResFile" & CStr(Cont%), "Resource File", "PASTA")
         Set nodX = Tree.Nodes.Add("ResFile" & CStr(Cont%), 4, MyProj.RESFILE & CStr(Cont%), MyProj.RESFILE, "RES")
      End If
      Set MyProj = Nothing
   Next
End Sub
Private Sub TreeInSubClass(Tree As TreeView, IdRoot$, CodRoot$, Item$, Coll As Collection, Optional Image = "")
   Dim n, CodColl$, DscColl$, Cod$, Dsc$
   Dim lFunc As New FUNCAO
   On Error Resume Next
   If Coll.Count = 0 Then Exit Sub
   Set nodX = Tree.Nodes.Add(IdRoot$, 4, CodRoot$ & Item$, CodRoot$, "PASTA")
   nodX.Sorted = True
   If Err = 0 Then
      Err = 0
   Else
      Call ShowError("TreeInSubClass")
      Exit Sub
   End If
   For Each n In Coll
      CodColl$ = UCase(n.FILENAME)
      DscColl$ = UCase(n.NOME)
      If Image = "" Then
         Set nodX = Tree.Nodes.Add(CodRoot$ & Item$, 4, CodColl$ & Item$, DscColl$)
         nodX.Sorted = True
      Else
         Set nodX = Tree.Nodes.Add(CodRoot$ & Item$, 4, CodColl$ & Item$, DscColl$, Image)
         nodX.Sorted = True
      End If
      Select Case UCase(CodRoot)
         Case "MODULOS"
            For Each m In n.FUNCOES
               Set lFunc = m
               Cod = UCase(lFunc.NOME)
               Dsc = Cod
               Set nodX = Tree.Nodes.Add(CodColl$ & Item$, 4, Cod$ & Item$, Dsc$, "PASTA")
               Set lFunc = Nothing
            Next
      End Select
   Next
End Sub
Public Sub AnalisarLinha_Bas(TxtLin$, ByRef pBAS As Modulo)
   Dim Pos%, Txt$
   Dim ColVar As New Collection
   Dim lFunc As New FUNCAO
'   On Error Resume Next
   If TxtLin = "" Then Exit Sub
   TxtLin$ = UCase(TxtLin$)
   If InStr(TxtLin$, "ATTRIBUTE VB_NAME") <> 0 Then Exit Sub
   
   Pos% = InStr(TxtLin$, "SUB ")
   If InStr(TxtLin$, "SUB ") <> 0 Or InStr(TxtLin$, "FUNCTION ") <> 0 Then
      Set lFunc = Nothing
      If CarregaFuncao(lFunc, TxtLin$) Then
         pBAS.FUNCOES.Add lFunc, lFunc.NOME
         mvarPROJETO.FUNCOES.Add lFunc, lFunc.NOME
      End If
      pBAS.LINHAS = CStr(Val(pBAS.LINHAS) + Val(lFunc.LINHAS) - 1)
   End If
End Sub
Public Sub AnalisarLinha_Fun(TxtLin$, ByRef pBAS As FUNCAO)
   Dim Pos%, Txt$
   Dim ColVar As New Collection
   Dim lFunc As New FUNCAO
   Exit Sub
   If TxtLin = "" Then Exit Sub
   TxtLin$ = UCase(TxtLin$)
   If InStr(TxtLin$, "ATTRIBUTE VB_NAME") <> 0 Then Exit Sub
   
   Pos% = InStr(TxtLin$, " SUB ")
   If Pos% <> 0 Then
      Call CarregaFuncao(lFunc, TxtLin$)
      pBAS.FUNCOES.Add Txt$, Txt$
      pBAS.FUNCOES(Txt$).isFuncao = False
      'pBAS.ESCOPO = GetWord(TxtLin$)
   End If
   Pos% = InStr(TxtLin$, " SUB ")
   If Pos% <> 0 Then
      Txt = Trim(Mid(TxtLin$, Pos + 5))
      Txt = Trim(Mid(Txt, 1, InStr(Txt, " ") - 1))
      pBAS.FUNCOES.Add Txt$, Txt$
      pBAS.FUNCOES(Txt$).isFuncao = False
   End If
End Sub
Private Sub AnalisarLinha_Vbg(TextLine$)
   Dim ColArq As New Collection
   If TextLine = "" Then Exit Sub
   If UCase(Mid(TextLine, 1, 7)) = "PROJECT" Or UCase(Mid(TextLine, 1, 14)) = "STARTUPPROJECT" Then
      Set ColArq = DefineFullPath(TextLine)
      mvarPathProj = ColArq("PATH")
      Call CarregaProjeto(ColArq("PATH"), ColArq("FILE"))
   End If
End Sub
Public Sub AnalisarLinha_Vbp(TxtLin$, ByRef pProj As PROJETO)
   Dim Pos%, Arq$
   Dim lForm As New FORMULARIO
   Dim lBas As New Modulo
   Dim lCls As New CLASSE
   Dim ColVar As New Collection
   
   If TxtLin = "" Then Exit Sub
   
   Set ColVar = DefineVarValor(TxtLin$)
   Select Case Trim(UCase(ColVar("VAR")))
      '* Nome do Projeto
      Case "NAME"
         pProj.NOME = Mid(ColVar("VALOR"), 2, Len(ColVar("VALOR")) - 2)
      '* Formulários do Projeto
      Case "FORM"
         Call CarregaForm(lForm, TxtLin$)
         pProj.FORMS.Add lForm, lForm.FILENAME
      '* Módulos do Projeto
      Case "MODULE"
         Arq$ = ColVar("VALOR")
         '* Recuperar Nome
         Pos% = InStr(Arq, ";")
         lBas.NOME = Mid(Arq, 1, Pos - 1)
         '* Recuperar Nome do Arquivo
         Arq = Mid(TxtLin$, InStr(TxtLin$, ";") + 1)
         Pos% = InStr(Arq$, "\")
         While Pos <> 0
            Arq = Mid(Arq, Pos + 1)
            Pos% = InStr(Arq$, "\")
         Wend
         lBas.FILENAME = Arq
         Call CarregaModulo(lBas, TxtLin$)
         pProj.MODULOS.Add lBas, lBas.FILENAME
         pProj.LINHAS = CStr(Val(pProj.LINHAS) + Val(lBas.LINHAS))
      '* Classes do Projeto
      Case "CLASS"
         Arq$ = ColVar("VALOR")
         '* Recuperar Nome
         Pos% = InStr(Arq, ";")
         lCls.NOME = Mid(Arq, 1, Pos - 1)
         '* Recuperar Nome do Arquivo
         Arq = Mid(TxtLin$, InStr(TxtLin$, ";") + 1)
         Pos% = InStr(Arq$, "\")
         While Pos <> 0
            Arq = Mid(Arq, Pos + 1)
            Pos% = InStr(Arq$, "\")
         Wend
         lCls.FILENAME = Arq
         pProj.CLASSES.Add lCls, lCls.FILENAME
      '* ResFile do Projeto
      Case "RESFILE32"
         pProj.RESFILE = Mid(ColVar("VALOR"), 2, Len(ColVar("VALOR")) - 2)
      '* Path do Projeto
      Case "PATH32"
         pProj.PATH = ColVar("VALOR")
      '* Descrição do Projeto
      Case "DESCRIPTION"
         pProj.DESCRIÇÃO = Mid(ColVar("VALOR"), 2, Len(ColVar("VALOR")) - 2)
   End Select
   Set lForm = Nothing
   Set lBas = Nothing
   Set lCls = Nothing
End Sub
Private Function DefineFullPath(TextLine$) As Collection
   Dim Txt$, ArqAux$, PathAux$, IniPath$
   Dim ColArq As New Collection
   Dim Cont%
   IniPath$ = ""
   PathAux = ""
   ArqAux = ""
   Txt = TextLine
   '* Define quantos sub-diretórios devemos percorrer
   Cont = 0
   While InStr(Txt, "..") <> 0
      Txt = Mid(Txt, InStr(Txt, "..") + 2)
      Cont = Cont + 1
   Wend
   While InStr(Txt, "\") <> 0
      PathAux = PathAux & Mid(Txt, 1, InStr(Txt, "\"))
      Txt = Mid(Txt, InStr(Txt, "\") + 1)
   Wend
   '* Se não está no mesmo diretório
   If Cont > 0 Then
      ArqAux = Txt
      nBarra = 0
      PathAux = Mid(PathAux, 2)
      
      Txt = mvarPathProj
      '* Nível do diretório do Projeto
      While InStr(Txt, "\") <> 0
         Txt = Mid(Txt, InStr(Txt, "\") + 1)
         nBarra = nBarra + 1
      Wend
      Txt = mvarPathProj
      '* Obter a partir do diretório do Projeto
      '* o diretório do arquivo.
      For i = 1 To nBarra - Cont
         Txt = Mid(Txt, 1, InStr(Txt, "\"))
         IniPath$ = IniPath$ + Txt
         Txt = Mid(mvarPathProj, Len(IniPath$) + 1)
'         Txt = Mid(mvarPathProj, InStr(mvarPathProj, "\") + 1)
      Next
   Else
      PathAux$ = mvarPathProj
      ArqAux$ = Mid(TextLine, InStr(TextLine, "=") + 1)
   End If
   ColArq.Add Trim(IniPath$) & Trim(PathAux$), "PATH"
   ColArq.Add Trim(ArqAux$), "FILE"
   Set DefineFullPath = ColArq
End Function
Private Function DefineVarValor(TxtLin$) As Collection
   Dim ColVar As New Collection, Pos%
   Pos% = InStr(TxtLin$, "=")
   If Pos% > 0 Then
      ColVar.Add Mid(TxtLin$, 1, Pos% - 1), "VAR"
      ColVar.Add Mid(TxtLin$, Pos% + 1), "VALOR"
   Else
      ColVar.Add "", "VAR"
      ColVar.Add "", "VALOR"
   End If
   Set DefineVarValor = ColVar
End Function
Private Function GetTipoSimbol(pTxt$) As String
   Select Case RightB(Trim(pTxt), 2)
      Case "$": GetTipoSimbol = "String"
      Case "%": GetTipoSimbol = "Integer"
      Case "&": GetTipoSimbol = "Long"
      Case "#": GetTipoSimbol = "Double"
      Case "@": GetTipoSimbol = "Data"
      Case Else: GetTipoSimbol = ""
   End Select
End Function
Private Function GetWord(Texto$) As String
   Dim Pos%, PosPar%
   Texto$ = Trim(Texto$)
   Pos = InStr(Trim(Texto$), " ")
   PosPar = InStr(Texto$, "(")
   If (PosPar <= Pos And PosPar <> 0) Or Pos = 0 Then Pos = PosPar
   If Pos = 0 Then
      GetWord = Trim(Mid(Texto$, 1))
   Else
      GetWord = Trim(Mid(Texto$, 1, Pos - 1))
   End If
End Function
Public Function GetWordInPar(TxtLin$)
   Dim PosAbre%, PosFecha%
   Dim Txt$
   PosAbre% = InStr(TxtLin$, "(")
   Txt$ = Mid(TxtLin$, PosAbre% + 1)
   While PosAbre% <> 0
      PosAbre% = InStr(Txt, "(")
      PosFecha% = InStr(Txt, ")")
      If PosAbre% <> 0 Then
'         Txt = Mid(Txt, IIf(PosAbre% = 1, 2, 1), PosAbre%) & _
'               GetWordInPar(Mid(Txt, IIf(PosAbre% = 1, 2, PosAbre%))) & _
'               Mid(Txt, PosFecha%)
         Txt = Mid(Txt, 1, PosAbre%) & GetWordInPar(Mid(Txt, PosAbre%)) & Mid(Txt, PosFecha%)
      Else
         If PosFecha% <> 0 Then Txt$ = Mid(Txt, 1, PosFecha%)
      End If
      PosAbre% = InStr(Mid(Txt, PosFecha% + 1), "(")
   Wend
   GetWordInPar = Mid(Txt, 1, Len(Txt) - 1)
End Function
Public Sub GetParam(pFunc As FUNCAO, ByVal TxtParam$)
   Dim Txt$, Pos%, Aux$, DepoisIgual As Boolean
   
   Dim MyParam As New VARIAVEL

   Pos = 1
   While Pos <> 0
      Pos = InStr(TxtParam, ",")
      Txt$ = Mid(TxtParam$, 1, IIf(Pos = 0, Len(TxtParam$), Pos - 1))
      TxtParam$ = Mid(TxtParam$, Pos + 1)
      While Len(Txt) <> 0
         Aux$ = UCase(GetWord(Txt))
         Aux$ = IIf(Len(Aux$) = 0, " ", Aux$)
         Txt = Trim(Mid(Txt, Len(Aux) + 1))
         Aux$ = IIf(Len(Aux) = 0, Txt, Aux$)
         Select Case Aux$
            Case "BYVAL"
               MyParam.PassArg = "Valor"
            Case "OPTIONAL"
               MyParam.ESCOPO = "Opcional"
            Case "AS"
               Aux$ = UCase(GetWord(Txt))
               Txt = Trim(Mid(Txt, Len(Aux) + 1))
               MyParam.Tipo = Aux$
            Case "=": DepoisIgual = True
            Case Else
               If DepoisIgual Then
                  MyParam.VALOR = Aux
               Else
                  MyParam.Tipo = GetTipoSimbol(Aux)
                  If MyParam.Tipo = "" Then
                     MyParam.NOME = Aux
                  Else
                     MyParam.NOME = Mid(Aux, 1, Len(Aux) - 1)
                  End If
               End If
         End Select
      Wend
      If MyParam.PassArg = "" Then MyParam.PassArg = "Referência"
      If MyParam.Tipo = "" Then MyParam.PassArg = "Variant"
      pFunc.Param.Add MyParam, MyParam.NOME
      Set MyParam = Nothing
   Wend
End Sub
