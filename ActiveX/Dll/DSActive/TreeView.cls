VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dsTreeView"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum cBoolean
  cFalse = 0
  CTrue = 1
End Enum
Private mvarTree As TreeView 'Object
Private mvarLstView As ListView
Private mvarItem As TreeView
Private mvarChild As Collection
Private mvarSql As String

Private mvarSqlList As String
Private mvarCabList As Variant

Private mvarParent As dsTreeView
Private mvarRoot As dsTreeView
Private mvarSorted As Boolean
Private mvarKey As String
Private mvarText As String
Private mvarSqlNivel As String
Private mvarDB As Variant
Private mvarNivel As String
Private mvarImagem As String
Private mvarExpandedImage As String
Private mvarItens As Collection
Private mvarFlood As Object
Private mvarLblFlood As Object
Private mvarPrikey As String
Private mvarChaveConcatenada As Boolean
Private mvarFldRoot As String
Private mvarFldDesc As String

'*************************
'* ScrollBar Definitions *
'*************************
Public Enum ScrollDirectionFlags
  sdLeft = &H1
  sdUp = &H2
  sdRight = &H4
  sdDown = &H8
End Enum
Public Enum RectFlags
  rfLeft = &H1
  rfTop = &H2
  rfRight = &H4
  rfBottom = &H8
End Enum
Public Enum SB_Type
  SB_HORZ = 0
  SB_VERT = 1
  SB_CTL = 2
  SB_BOTH = 3
End Enum

'****************************
'* TVItem State, stateMask
'****************************
'Public Type TVHITTESTINFO   ' was TV_HITTESTINFO
'  Pt As POINTAPI
'  Flags As TVHITTESTINFO_Flags
'  hItem As Long
'End Type
Public Property Let CabList(ByVal vData As Variant)
    mvarCabList = vData
End Property
Public Property Set CabList(ByVal vData As Object)
    Set mvarCabList = vData
End Property
Public Property Get CabList() As Variant
    If IsObject(mvarCabList) Then
        Set CabList = mvarCabList
    Else
        CabList = mvarCabList
    End If
End Property
Public Property Let FldDesc(ByVal vData As String)
    mvarFldDesc = vData
End Property
Public Property Get FldDesc() As String
    FldDesc = mvarFldDesc
End Property
Public Property Let fldROOT(ByVal vData As String)
    mvarFldRoot = vData
End Property
Public Property Get fldROOT() As String
    fldROOT = mvarFldRoot
End Property
Public Property Let Prikey(ByVal vData As String)
    mvarPrikey = vData
End Property
Public Property Get Prikey() As String
    Prikey = mvarPrikey
End Property
Public Property Let ChaveConcatenada(ByVal vData As Boolean)
    mvarChaveConcatenada = vData
End Property
Public Property Get ChaveConcatenada() As Boolean
    ChaveConcatenada = mvarChaveConcatenada
End Property
Public Property Set LblFlood(ByVal vData As Object)
    Set mvarLblFlood = vData
End Property
Public Property Get LblFlood() As Object
    Set LblFlood = mvarLblFlood
End Property
Public Property Set Flood(ByVal vData As Object)
    Set mvarFlood = vData
End Property
Public Property Get Flood() As Object
    Set Flood = mvarFlood
End Property
Public Property Set Itens(ByVal vData As Object)
    Set mvarItens = vData
End Property
Public Property Get Itens() As Collection
   If mvarItens Is Nothing Then
      Set mvarItens = New Collection
   End If
    Set Itens = mvarItens
End Property
Public Property Let Imagem(ByVal vData As String)
    mvarImagem = vData
End Property
Public Property Get Imagem() As String
    Imagem = mvarImagem
End Property
Public Property Let ExpandedImage(ByVal vData As String)
    mvarExpandedImage = vData
End Property
Public Property Get ExpandedImage() As String
    ExpandedImage = mvarExpandedImage
End Property
Public Property Let Nivel(ByVal vData As String)
    mvarNivel = vData
End Property
Public Property Get Nivel() As String
    Nivel = mvarNivel
End Property
Public Property Let DB(ByVal vData As Variant)
    mvarDB = vData
End Property
Public Property Set DB(ByVal vData As Object)
    Set mvarDB = vData
End Property
Public Property Get DB() As Variant
    If IsObject(mvarDB) Then
        Set DB = mvarDB
    Else
        DB = mvarDB
    End If
End Property
Public Property Let SqlNivel(ByVal vData As String)
    mvarSqlNivel = vData
End Property
Public Property Get SqlNivel() As String
    SqlNivel = mvarSqlNivel
End Property
Public Property Set LstView(ByVal vData As Object)
    Set mvarLstView = vData
End Property
Public Property Get LstView() As Object
    Set LstView = mvarLstView
End Property
Public Property Set Tree(ByVal vData As Object)
    Set mvarTree = vData
End Property
Public Property Get Tree() As Object
    Set Tree = mvarTree
End Property
Public Property Let Text(ByVal vData As String)
    mvarText = vData
End Property
Public Property Get Text() As String
    Text = mvarText
End Property
Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property
Public Property Get Key() As String
    Key = mvarKey
End Property
Public Property Let Sorted(ByVal vData As Boolean)
    mvarSorted = vData
End Property
Public Property Get Sorted() As Boolean
    Sorted = mvarSorted
End Property
Public Property Set Root(ByVal vData As Object)
    Set mvarRoot = vData
End Property
Public Property Get Root() As dsTreeView
    Set Root = mvarRoot
End Property
Public Property Set Parent(ByVal vData As Object)
    Set mvarParent = vData
End Property
Public Property Get Parent() As dsTreeView
    Set Parent = mvarParent
End Property
Public Property Let Sql(ByVal vData As String)
    mvarSql = vData
End Property
Public Property Get Sql() As String
    Sql = mvarSql
End Property
Public Property Let SqlList(ByVal vData As String)
    mvarSqlList = vData
End Property
Public Property Get SqlList() As String
    SqlList = mvarSqlList
End Property
Public Property Set Child(ByVal vData As Object)
    Set mvarChild = vData
End Property
Public Property Get Child() As Collection
   If mvarChild Is Nothing Then
      Set mvarChild = New Collection
   End If
   Set Child = mvarChild
End Property
Public Property Set Item(ByVal vData As Object)
    Set mvarItem = vData
End Property
Public Property Get Item() As dsTreeView
    Set Item = mvarItem
End Property
Private Sub Class_Initialize()
   mvarSorted = True
   mvarFldRoot = "IDROOT"
End Sub
Public Function ExisteNo(Tree As Object, Key As String) As Boolean
   Dim Nome As String
   On Error Resume Next
   Nome = Tree.Nodes(Key)
   ExisteNo = (Err = 0)
   Err = 0
End Function
Public Function GetParent(NodX As Node, Nivel As Integer) As Node
   Dim lNo As Node
   Dim i As Integer
   Set lNo = NodX
   i = 0
   While NodX.Root.Key <> lNo.Key
      i = i + 1
      Set lNo = lNo.Parent
   Wend
   i = i - Nivel
   Set lNo = NodX
   If i < 0 Then
      While i < 0
         i = i + 1
         If lNo.Children > 0 Then
            Set lNo = lNo.Child
         Else
'            If i <> 0 Then Set lNo = Nothing
            i = 0
         End If
      Wend
   Else
      While i > 0
         i = i - 1
         Set lNo = lNo.Parent
      Wend
   End If
   Set GetParent = lNo
End Function
Public Function GetNivel(NodX As Object) As Integer
   Dim lNo As Object
   Dim i As Integer
   On Error Resume Next
   Set lNo = NodX
   i = 0
   Do While NodX.Root.Key <> lNo.Key
      If lNo.Parent Is Nothing Then
         Exit Do
      Else
         i = i + 1
         Set lNo = lNo.Parent
      End If
   Loop
   GetNivel = i
End Function
Public Function LocalizarNo_0%(ObjRecordSet As Recordset, ByRef Tree As Object, ArrID_DSC_PAI, Tabela As String, IdFind$)
   Dim MySql$, MyCol As New Collection
   Dim MyRoot$, FimArq As Boolean

   Dim i%, Ind%
   
   Dim fldID As String
   Dim fldROOT As String
   
   On Error Resume Next

   Set MyCol = Nothing
   Set MyCol = New Collection

   fldID = ArrID_DSC_PAI(0)
   fldROOT = ArrID_DSC_PAI(2)
   
   MySql$ = "select " & fldID & ", " & fldROOT
   MySql$ = MySql$ & " from " & Tabela
   MySql$ = MySql$ & " where " & fldID & "="

   MyRoot$ = "'" & IdFind$ & "'"
   MyRoot$ = IdFind$
   Ind = 1
   MyCol.Add MyRoot$

   FimArq = False
   While Not FimArq
      FimArq = True
'      Call Dbase.AbreTabela(MySql & MyRoot$)
'      If Dbase.CodeSql Then
'         MyRoot$ = "'" & Dbase.Dys(fldROOT) & "" & "'"
'         FimArq = (Dbase.Dys(fldROOT) = "")
         If Not FimArq Then
            Ind = Ind + 1
            MyCol.Add MyRoot$
         End If
'      End If
      If Ind > 100 Then
         Err = 1
         Call ClsMsg.ShowError("M_LocalizarNo( " & Tabela & "." & mvarFldRoot & " = " & MyRoot$ & " )")
         LocalizarNo_0 = 0
         Exit Function
      End If
   Wend

   For i = Ind To 1 Step -1
      Tree.Nodes(MyCol(i)).Selected = True
      Call MontarNode(ObjRecordSet, Tree, MyCol(i), ArrID_DSC_PAI)
      Tree.Nodes(MyCol(i)).Expanded = True
   Next
   IdFind$ = "'" & IdFind$ & "'"
   LocalizarNo_0 = Tree.Nodes(IdFind$).Index
'  If Selecionar Then
     Tree.Nodes(IdFind$).EnsureVisible
     Tree.SelectedItem = Tree.Nodes(IdFind$)
'   End If
Exit Function
End Function
Public Sub MontarNode(ObjRecordSet As Object, TView As Object, pIDPAI As String, Arr_ID_DSC_PAI, Optional MDIStatus)
   Dim NodX As Object
   Dim QtdReg As Long
   Dim pos As Long
   Dim PosID As Integer
   Dim PosDSC As Integer
   Dim PosPAI As Integer

   On Error GoTo Fim

   If ObjRecordSet.RecordCount = 0 Then Exit Sub
   
   PosID = Arr_ID_DSC_PAI(0)
   PosDSC = Arr_ID_DSC_PAI(1)
   PosPAI = Arr_ID_DSC_PAI(2)
   
   QtdReg = ObjRecordSet.RecordCount
   
   On Error Resume Next
   While Not ObjRecordSet.EOF
      On Error Resume Next
      'Set NodX = TView.Nodes.Add(TView.Nodes(pIDPAI).Key, tvwChild, "k" & ObjRecordSet(PosID), Mid(ObjRecordSet(PosDSC), 1, 30), "CLOSED")
      Set NodX = TView.Nodes.Add(, , "k" & ObjRecordSet(PosID), Mid(ObjRecordSet(PosDSC), 1, 30), "CLOSED")
      NodX.ExpandedImage = "OPEN"
      NodX.Sorted = True
      
      NodX.Tag = ObjRecordSet(PosPAI)
      If Not IsMissing(MDIStatus) And pos <> 100 Then
         pos = CInt((ObjRecordSet.AbsolutePosition / QtdReg) * 100)
         pos = IIf(pos = 100, 99, pos)
         Call ClsCtrl.ShowProgress(MDIStatus, CInt(pos / 2))
      End If
      ObjRecordSet.MoveNext
   Wend
   ObjRecordSet.MoveFirst
   While Not ObjRecordSet.EOF
      Set TView.Nodes("k" & ObjRecordSet(PosID)).Parent = TView.Nodes(pIDPAI)
      If Not IsMissing(MDIStatus) And pos <> 100 Then
         pos = CInt((ObjRecordSet.AbsolutePosition / QtdReg) * 100)
         pos = IIf(pos = 100, 99, pos)
         Call ClsCtrl.ShowProgress(MDIStatus, CInt(pos / 2) + 50)
      End If
      ObjRecordSet.MoveNext
   Wend
   If Not IsMissing(MDIStatus) Then
      Call ClsCtrl.ShowProgress(MDIStatus, 100)
   End If
   Set NodX = Nothing
Exit Sub
Fim:
  ClsMsg.ShowError
End Sub
Public Sub MontarTView(ObjRecordSet As Object, TView As Object, Array_ID_DSC_PAI, Optional MDIStatus, Optional ExibirKey = False, Optional isCheckBox = False, Optional RootText = "Sistema", Optional ClosedImage, Optional OpenImage, Optional ExibirCOD As Boolean = False, Optional IniExpand As Boolean = False)
   Dim Key    As String
   Dim Texto  As String
   Dim Root   As String
   
   Dim PosID  As Variant
   Dim PosDSC As Variant
   Dim PosPAI As Variant
   Dim PosCOD As Variant
   
   Dim NodX   As Object
   Dim QtdReg As Long
   Dim pos    As Long
   Dim Cont   As Long
   
   Dim n      As Variant
   
   On Error Resume Next
   
   TView.Nodes.Clear
   
   If Not TView.ImageList Is Nothing Then
      If IsMissing(ClosedImage) Then
         ClosedImage = "CLOSED"
      End If
      If IsMissing(ClosedImage) Then
         OpenImage = "OPEN"
      End If
   End If
   
   Root$ = "k0"
   Set NodX = TView.Nodes.Add(, , Root$, RootText, ClosedImage)
   
   If Not TView.ImageList Is Nothing Then
      NodX.ExpandedImage = OpenImage
   End If
   NodX.Sorted = True
   Call BoldTreeNode(TView, TView.Nodes(Root$))
   
   QtdReg = ObjRecordSet.RecordCount
   
   If QtdReg < 1 Then Exit Sub
   
   PosID = Array_ID_DSC_PAI(0)
   PosDSC = Array_ID_DSC_PAI(1)
   PosPAI = Array_ID_DSC_PAI(2)
   PosCOD = Array_ID_DSC_PAI(3)
   
   ObjRecordSet.MoveFirst
   While Not ObjRecordSet.EOF
      Key = "k" & ObjRecordSet(PosID) & ""
      Texto = IIf(ExibirKey, Mid(Key, 2) & " - ", "") & ObjRecordSet(PosDSC) & ""
      Texto = IIf(ExibirCOD, ObjRecordSet(PosCOD) & " - ", "") & ObjRecordSet(PosDSC) & ""
      
      If Key = "k149" Then
         Key = Key
      End If
      Set NodX = TView.Nodes.Add(, , Key, Texto, IIf(isCheckBox, "UNCHECKED", ClosedImage))
      
      Key = Root$
      Key = IIf(ObjRecordSet(PosPAI) & "" = "", Root$, "k" & ObjRecordSet(PosPAI))
      
      'Set NodX.Parent = TView.Nodes(Key)
      NodX.Tag = Key
      If Not isCheckBox And Not TView.ImageList Is Nothing Then
         NodX.ExpandedImage = OpenImage
      End If
      NodX.Sorted = True
'      NodX.Expanded = True
      
'      If Not IsMissing(MDIStatus) And Pos <> 100 Then
'          Pos = CInt((ObjRecordSet.AbsolutePosition / QtdReg) * 100)
'          Call ClsCtrl.ShowProgress(MDIStatus, Pos)
'      End If
      
      ObjRecordSet.MoveNext
   Wend
   
   For Each n In TView.Nodes
'      If ExisteNo(TView, n.Tag) Then
         Set n.Parent = TView.Nodes(n.Tag)
'         TView.Nodes(n.Tag).Sorted = True
'         TView.Nodes(n.Tag).Expanded = IniExpand
'      End If
'      n.Expanded = IniExpand
'      n.Sorted = True
   Next
   For Each n In TView.Nodes
      n.Expanded = IniExpand
      n.Sorted = True
   Next
   TView.Nodes(1).Expanded = True
   TView.HideSelection = False
   Set NodX = Nothing
   
'   If Not IsMissing(MDIStatus) Then
'      Call ClsCtrl.ShowProgress(MDIStatus, 100)
'   End If

End Sub

Public Sub M_MontarTreeView(Optional Dbase, Optional Tree, Optional Sql = "", Optional Root = "", Optional IdRoot = "R", Optional DscRoot = "Sistema", Optional LenKEy = 3, Optional Imagem = "", Optional cTAG, Optional Flood = "", Optional LblFlood = "")
   Dim n, MyItem As New dsTreeView
   Dim Flood_Ini&, Flood_Fim&, Cont&
   Dim dyTre As Object
   
   Dim NodX As Object
   Dim Cod$, CodRoot$, Txt$, msg$
   Dim k%, Vet() As String
   Dim NvZero%, Tam%
   
   On Error GoTo Fim
   '* Redefinir Parâmetros
   Set Dbase = IIf(IsMissing(Dbase), mvarDB, Dbase)
   Set Tree = IIf(IsMissing(Tree), mvarTree, Tree)
   Sql = IIf(Sql = "", mvarSql, Sql)
   Root = IIf(Root = "", Not mvarParent Is Nothing, Root)
   Imagem = IIf(Imagem = "", mvarImagem, Imagem)
   ExpandedImage = IIf(ExpandedImage = "", mvarExpandedImage, ExpandedImage)
   'If Root Then IdRoot = IIf(dBase = "", , IdRoot)
   'cTAG = IIf(cTAG = "", Me.KEYmvarDB, cTAG)
   'Flood = IIf(Flood = "", mvarDB, Flood)
   'FLblFlood = IIf(LblFlood = "", mvarDB, LblFlood)
   Call ClsDsr.SetHourglass(Tree.Parent.hWnd)
   
   ReDim Vet(0)
   '* Montar Query
   Sql = IIf(Trim(Sql) = "", Tree.Tag, Sql)
   If Me.Tree.Nodes.Count = 0 Then
      Sql = Sql & " where " & mvarFldRoot & "=''"
      Sql = Sql & " or isnull(" & mvarFldRoot & ")"
   Else
      Sql = Sql & " where " & mvarFldRoot & "='" & ClsDsr.EliminarString(mvarKey, "'") & "'"
   End If
   Call Dbase.AbreTabela(Sql, dyTre)
'   mvarTree.Visible = False
'   mvarTree.Refresh

   If Dbase.CodeSql = -1 Then  '* FOUND = -1
      dyTre.MoveFirst
      While Not dyTre.EOF
         With MyItem
            .Key = "'" & dyTre(mvarPrikey) & "'"
            .Text = dyTre(mvarFldDesc)
            .FldDesc = mvarFldDesc
            .fldROOT = mvarFldRoot
'            Set .DB = mvarDB
'            Set .Tree = mvarTree
            Set .Parent = Me
            '* Configuração inicial de Flood
            If Not mvarFlood Is Nothing Then
               Set .Flood = mvarFlood
               Set .LblFlood = mvarLblFlood
               mvarFlood.Value = 0
            End If
         End With
         '* se já existe o Item ignora o erro.
         On Error Resume Next
         Child.Add MyItem, MyItem.Key
         If Err = 457 Or Err = 0 Then
            On Error GoTo 0
            On Error GoTo Fim
         Else
            GoTo Fim
         End If
         Set MyItem = Nothing
         dyTre.MoveNext
      Wend
   End If
   
   
'   If Root Then If DscRoot = "Sistema" Then DscRoot = App.Title
   Cont = 1
   If Not mvarChild Is Nothing Then
      For Each n In mvarChild
         '* Configurações de Flood
         If Not mvarFlood Is Nothing Then
            Flood_Ini = mvarFlood.Value
            Flood_Fim = CDbl(IIf(mvarFlood.Tag = "", 100, mvarFlood.Tag))
            mvarFlood.Enabled = True
            mvarFlood.Visible = True
            If Not mvarLblFlood Is Nothing Then
               mvarLblFlood.Enabled = True
               mvarLblFlood.Caption = Trim(CStr(Flood_Ini)) & " %"
               mvarLblFlood.Visible = True
               mvarLblFlood.Refresh
            End If
         End If
         Set MyItem = n
         With MyItem
            Set .DB = mvarDB
            Set .Tree = mvarTree
            Set .Parent = Me
            .Imagem = IIf(.Imagem = "", mvarImagem, .Imagem)
            .ExpandedImage = IIf(.ExpandedImage = "", mvarExpandedImage, .ExpandedImage)
'* Nivel
'            .Nivel = IIf(.Nivel = "", CInt(mvarNivel) + 1, .Nivel)
            .ChaveConcatenada = (mvarChaveConcatenada Or .ChaveConcatenada)
            .Prikey = mvarPrikey
            
            If mvarParent Is Nothing Then
               Set .Root = .Parent
               'Set .Root = IIf(.Parent.Key = "", MyItem, .Parent)
            Else
               Set .Root = IIf(mvarParent.Root Is Nothing, mvarParent, mvarParent.Root)
            End If
            .Sql = IIf(.Sql = "", mvarSql, .Sql)
            .SqlNivel = IIf(.SqlNivel = "", mvarSqlNivel, .SqlNivel)
         End With
         If mvarParent Is Nothing Then
            Set NodX = Tree.Nodes.Add(, , MyItem.Key, MyItem.Text)       ' Root
         Else
            Set NodX = Tree.Nodes.Add(mvarKey, 4, MyItem.Key, MyItem.Text) ' Root
         End If
         Tree.Nodes(Tree.Nodes.Count).Sorted = MyItem.Sorted
         'Tree.Nodes(Tree.Nodes.Count).Tag = MyItem.Key
         If MyItem.Imagem <> "" Then NodX.Image = MyItem.Imagem
         If MyItem.ExpandedImage <> "" Then NodX.ExpandedImage = MyItem.ExpandedImage

         Set NodX = Tree.Nodes.Add(MyItem.Key, 4, "", "")
         If mvarParent Is Nothing Then
            '* Se elemento já existir ignora erro.
            On Error Resume Next
            Itens.Add MyItem, MyItem.Key
'            Set MyItem.Root = IIf(MyItem.Root Is Nothing, Me, MyItem.Root)
            If Err = 457 Or Err = 0 Then '* err=457->This key is already associated with an element of this collection
               On Error GoTo 0
               On Error GoTo Fim
            Else
               GoTo Fim
            End If
         Else
            On Error Resume Next
            mvarRoot.Itens.Add MyItem, MyItem.Key
            If Err = 457 Or Err = 0 Then '* err=457->This key is already associated with an element of this collection
               On Error GoTo 0
               On Error GoTo Fim
            Else
               GoTo Fim
            End If
         End If
         Set MyItem = Nothing
         If Not mvarFlood Is Nothing Then
            mvarFlood.Value = CInt(Flood_Ini + ((Cont / mvarChild.Count) * (Flood_Fim - Flood_Ini)))
            If Not mvarLblFlood Is Nothing Then
               If Trim(CStr(mvarFlood.Value)) & " %" <> mvarLblFlood.Caption Then
                  mvarLblFlood.Caption = Trim(CStr(mvarFlood.Value)) & " %"
                  mvarLblFlood.Refresh
               End If
            End If
         End If
         Cont = Cont + 1
      Next
      '* Montar List
      If Not (mvarLstView Is Nothing) Then
         If mvarTree.Nodes.Count > 0 Then
            Call Me.M_PopulaListView(mvarTree.Nodes(1).Key)  '.SelectedItem.Key)
         End If
      End If
   End If
   If Not mvarFlood Is Nothing Then
      mvarFlood.Visible = False
      mvarLblFlood.Visible = False
      mvarLblFlood.Refresh
   End If
   Tree.HideSelection = False
'   If Root Then Tree.Nodes(2).EnsureVisible
'   Tree.SelectedItem = Tree.Nodes(1)
 '  Tree.Visible = True
 '  Tree.Refresh
   Call ClsDsr.SetDefault(Tree.Parent.hWnd)
Exit Sub
Fim:
   Tree.Visible = True
   Tree.Refresh
   Call ClsDsr.SetDefault(Tree.Parent.hWnd)
   ReDim Vet(0)
   ClsMsg.ShowError
End Sub

Public Sub M_MontarTreeView0(Optional Dbase, Optional Tree, Optional Sql = "", Optional Root = "", Optional IdRoot = "R", Optional DscRoot = "Sistema", Optional LenKEy = 3, Optional Imagem = "", Optional cTAG, Optional Flood = "", Optional LblFlood = "")
   Dim n, MyItem As New dsTreeView
   Dim Flood_Ini&, Flood_Fim&, Cont&
   Dim dyTre As Recordset
   
   Dim NodX As Object
   Dim Cod$, CodRoot$, Txt$, msg$
   Dim k%, Vet() As String
   Dim NvZero%, Tam%
   
   On Error GoTo Fim
   '* Redefinir Parâmetros
   Set Dbase = IIf(IsMissing(Dbase), mvarDB, Dbase)
   Set Tree = IIf(IsMissing(Tree), mvarTree, Tree)
   Sql = IIf(Sql = "", mvarSql, Sql)
   Root = IIf(Root = "", Not mvarParent Is Nothing, Root)
   Imagem = IIf(Imagem = "", mvarImagem, Imagem)
   ExpandedImage = IIf(ExpandedImage = "", mvarExpandedImage, ExpandedImage)
   'If Root Then IdRoot = IIf(dBase = "", , IdRoot)
   'cTAG = IIf(cTAG = "", Me.KEYmvarDB, cTAG)
   'Flood = IIf(Flood = "", mvarDB, Flood)
   'FLblFlood = IIf(LblFlood = "", mvarDB, LblFlood)
   Call ClsDsr.SetHourglass(Tree.Parent.hWnd)
   
   ReDim Vet(0)
   '* Montar Query
   Sql = Mid(Me.Sql, 1, InStr(UCase(Me.Sql), "WHERE") + 5) & Me.SqlNivel & mvarNivel
   If mvarChaveConcatenada And Me.Key <> "" Then
      '********************************************
      '********************************************
      If Len(Me.Key) - 2 >= 2 Then
         Sql = Sql & " and left(" & Prikey & "," & CStr(Len(Me.Key) - 2) & ")=" & Me.Key
      End If
   End If
   Sql = Sql & " and " & Mid(Me.Sql, InStr(UCase(Me.Sql), "WHERE") + 5)
   If Trim(Sql) = "" Then
      Sql = Tree.Tag
   Else
      Tree.Tag = Sql
   End If
   Call Dbase.AbreTabela(Sql, dyTre)
   mvarTree.Visible = False
   mvarTree.Refresh

   If Dbase.CodeSql = -1 Then  '* FOUND = -1
      dyTre.MoveFirst
      While Not dyTre.EOF
         With MyItem
            .Key = "'" & dyTre(0) & "'"
            .Text = dyTre(1)
'            Set .DB = mvarDB
'            Set .Tree = mvarTree
            Set .Parent = Me
            '* Configuração inicial de Flood
            If Not mvarFlood Is Nothing Then
               Set .Flood = mvarFlood
               Set .LblFlood = mvarLblFlood
               mvarFlood.Value = 0
            End If
         End With
         '* se já existe o Item ignora o erro.
         On Error Resume Next
         Child.Add MyItem, MyItem.Key
         If Err = 457 Or Err = 0 Then
            On Error GoTo 0
            On Error GoTo Fim
         Else
            GoTo Fim
         End If
         Set MyItem = Nothing
         dyTre.MoveNext
      Wend
   End If
   
   
'   If Root Then If DscRoot = "Sistema" Then DscRoot = App.Title
   Cont = 1
   If Not mvarChild Is Nothing Then
      For Each n In mvarChild
         '* Configurações de Flood
         If Not mvarFlood Is Nothing Then
            Flood_Ini = mvarFlood.Value
            Flood_Fim = CDbl(IIf(mvarFlood.Tag = "", 100, mvarFlood.Tag))
            mvarFlood.Enabled = True
            mvarFlood.Visible = True
            If Not mvarLblFlood Is Nothing Then
               mvarLblFlood.Enabled = True
               mvarLblFlood.Caption = Trim(CStr(Flood_Ini)) & " %"
               mvarLblFlood.Visible = True
               mvarLblFlood.Refresh
            End If
         End If
         Set MyItem = n
         With MyItem
            Set .DB = mvarDB
            Set .Tree = mvarTree
            Set .Parent = Me
            .Imagem = IIf(.Imagem = "", mvarImagem, .Imagem)
            .ExpandedImage = IIf(.ExpandedImage = "", mvarExpandedImage, .ExpandedImage)
            .Nivel = IIf(.Nivel = "", CInt(mvarNivel) + 1, .Nivel)
            .ChaveConcatenada = (mvarChaveConcatenada Or .ChaveConcatenada)
            .Prikey = mvarPrikey
            
            If mvarParent Is Nothing Then
               Set .Root = .Parent
            Else
               Set .Root = IIf(mvarParent.Root Is Nothing, mvarParent, mvarParent.Root)
            End If
            .Sql = IIf(.Sql = "", mvarSql, .Sql)
            .SqlNivel = IIf(.SqlNivel = "", mvarSqlNivel, .SqlNivel)
         End With
         If mvarParent Is Nothing Then
            Set NodX = Tree.Nodes.Add(, , MyItem.Key, MyItem.Text)       ' Root
         Else
            Set NodX = Tree.Nodes.Add(mvarKey, 4, MyItem.Key, MyItem.Text) ' Root
         End If
         Tree.Nodes(Tree.Nodes.Count).Sorted = MyItem.Sorted
         Tree.Nodes(Tree.Nodes.Count).Tag = MyItem.Key
         If MyItem.Imagem <> "" Then NodX.Image = MyItem.Imagem
         If MyItem.ExpandedImage <> "" Then NodX.ExpandedImage = MyItem.ExpandedImage

         Set NodX = Tree.Nodes.Add(MyItem.Key, 4, "", "")
         If mvarParent Is Nothing Then
            '* Se elemento já existir ignora erro.
            On Error Resume Next
            Itens.Add MyItem, MyItem.Key
            If Err = 457 Or Err = 0 Then '* err=457->This key is already associated with an element of this collection
               On Error GoTo 0
               On Error GoTo Fim
            Else
               GoTo Fim
            End If
         Else
            On Error Resume Next
            mvarRoot.Itens.Add MyItem, MyItem.Key
            If Err = 457 Or Err = 0 Then '* err=457->This key is already associated with an element of this collection
               On Error GoTo 0
               On Error GoTo Fim
            Else
               GoTo Fim
            End If
         End If
         Set MyItem = Nothing
         If Not mvarFlood Is Nothing Then
            mvarFlood.Value = CInt(Flood_Ini + ((Cont / mvarChild.Count) * (Flood_Fim - Flood_Ini)))
            If Not mvarLblFlood Is Nothing Then
               If Trim(CStr(mvarFlood.Value)) & " %" <> mvarLblFlood.Caption Then
                  mvarLblFlood.Caption = Trim(CStr(mvarFlood.Value)) & " %"
                  mvarLblFlood.Refresh
               End If
            End If
         End If
         Cont = Cont + 1
      Next
   End If
   If Not mvarFlood Is Nothing Then
      mvarFlood.Visible = False
      mvarLblFlood.Visible = False
      mvarLblFlood.Refresh
   End If
   Tree.HideSelection = False
'   If Root Then Tree.Nodes(2).EnsureVisible
   Tree.SelectedItem = Tree.Nodes(1)
   Tree.Visible = True
   Tree.Refresh
   Call ClsDsr.SetDefault(Tree.Parent.hWnd)
Exit Sub
Fim:
   Tree.Visible = True
   Tree.Refresh
   Call ClsDsr.SetDefault(Tree.Parent.hWnd)
   ReDim Vet(0)
   ClsMsg.ShowError
End Sub
Public Sub M_AddItem(pIDROOT$, pKEY$, pText$)
   Dim MyItem As New dsTreeView
   Dim NodX As Object
   
   mvarTree.Visible = False
   mvarTree.Refresh
   With MyItem
      .Key = "'" & pKEY$ & "'"
      .Text = pText$
      .FldDesc = mvarFldDesc
      .fldROOT = mvarFldRoot
      On Error Resume Next
      Child.Add MyItem, MyItem.Key
      If Err = 457 Or Err = 0 Then
         On Error GoTo 0
         On Error GoTo Fim
      Else
         GoTo Fim
      End If
      Set .DB = mvarDB
      Set .Tree = mvarTree
      Set .Parent = mvarItens("'" & pIDROOT & "'")
      
      .Imagem = IIf(.Imagem = "", mvarImagem, .Imagem)
      .ExpandedImage = IIf(.ExpandedImage = "", mvarExpandedImage, .ExpandedImage)
      .ChaveConcatenada = (mvarChaveConcatenada Or .ChaveConcatenada)
      .Prikey = mvarPrikey
            
      If mvarParent Is Nothing Then
          Set .Root = .Parent
      Else
         Set .Root = IIf(mvarParent.Root Is Nothing, mvarParent, mvarParent.Root)
      End If
      .Sql = IIf(.Sql = "", mvarSql, .Sql)
      .SqlNivel = IIf(.SqlNivel = "", mvarSqlNivel, .SqlNivel)

      If .Parent Is Nothing Then
         Set NodX = Tree.Nodes.Add(, , MyItem.Key, MyItem.Text)       ' Root
      Else
         Set NodX = Tree.Nodes.Add(.Parent.Key, 4, MyItem.Key, MyItem.Text) ' Root
      End If
      mvarTree.Nodes(mvarTree.Nodes.Count).Sorted = MyItem.Sorted
      'mvarTree.Nodes(mvarTree.Nodes.Count).Tag = MyItem.Key
      If MyItem.Imagem <> "" Then NodX.Image = MyItem.Imagem
      If MyItem.ExpandedImage <> "" Then NodX.ExpandedImage = MyItem.ExpandedImage
   
      Set NodX = Tree.Nodes.Add(MyItem.Key, 4, "", "")
      On Error Resume Next
      Itens.Add MyItem, MyItem.Key
      If Err = 457 Or Err = 0 Then '* err=457->This key is already associated with an element of this collection
         On Error GoTo 0
         On Error GoTo Fim
      Else
         GoTo Fim
      End If
   End With
   Set MyItem = Nothing
   mvarTree.HideSelection = False
   mvarTree.Visible = True
   mvarTree.Refresh
   Call ClsDsr.SetDefault(Tree.Parent.hWnd)
Exit Sub
Fim:
   mvarTree.Visible = True
   mvarTree.Refresh
   Call ClsDsr.SetDefault(Tree.Parent.hWnd)
   ReDim Vet(0)
   ClsMsg.ShowError
End Sub
Public Function CheckNode(ByVal Node As ComctlLib.Node) As Collection
   Dim MyColl As New Collection
   Dim sAux As String
   Dim i As Integer
   Dim n As Variant
   
   On Error Resume Next
   
   sAux = Node.Image
   
   Node.Image = IIf(Node.Image = "CHECKED", "UNCHECKED", "CHECKED")
   
   MyColl.Add Node
   For i = 1 To Node.Children
      If i = 1 Then
         Set Node = Node.Child
      Else
         Set Node = Node.Next
      End If
      Node.Image = sAux
      For Each n In CheckNode(Node)
         MyColl.Add n
      Next
   Next
   Set CheckNode = MyColl
End Function

Public Sub M_ExpandNode(ByRef Node As ComctlLib.Node)
   Dim MyItem As New dsTreeView
   Dim lIndex%
   On Error GoTo Fim
'   lIndex = Node.Index
   If mvarRoot Is Nothing Then
      Set MyItem = mvarItens(Node.Key)
   Else
      Set MyItem = mvarRoot.Itens(Node.Key)
   End If
   If Node.Child.Text = "" Then
      If Node.Child.Index <= mvarTree.Nodes.Count Then
         mvarTree.Nodes.Remove Node.Child.Index
      End If
      Set MyItem.Tree = mvarTree
      MyItem.M_MontarTreeView
      If Node.Children = 0 Then
         mvarTree.Nodes(Node.Key).ExpandedImage = Node.Image
      Else
         mvarTree.Nodes(Node.Key).ExpandedImage = Node.ExpandedImage
      End If
   End If
'   mvarTree.SelectedItem = mvarTree.Nodes(lIndex)
   DoEvents
'   mvarTree.Refresh
Fim:
  
End Sub
Public Sub PopulaListView(LstView As Object, NodX As Object, pIDROOT$)
   Dim i&, pos%, Sql$
   Dim MyNode As ComctlLib.Node
   
   LstView.ListItems.Clear
   pIDROOT = ClsDsr.EliminarString(CStr(pIDROOT), "'")
   Call M_ExpandNode(mvarTree.Nodes("'" & pIDROOT & "'"))
   
   pos = InStr(UCase(mvarSqlList), "WHERE")
   If pos > 0 Then
      Sql = Mid(mvarSqlList, 1, pos - 1)
      Sql = Sql & " where " & Me.fldROOT & "='" & CStr(pIDROOT) & "'"
      Sql = Sql & " and " & Mid(mvarSqlList, pos + 5)
   Else
      Sql = mvarSqlList & " where " & Me.fldROOT & "='" & CStr(pIDROOT) & "'"
   End If
   If pIDROOT <> "" Then
      Call ClsCtrl.MontarDbList(DB, mvarLstView, Sql, mvarCabList, mvarPrikey)
   Else
      'Call RefreshList(DB, Me.LstPiece, Sql, Cab, "IDPIECE")
   End If

End Sub
Public Sub M_PopulaListView(pIDROOT$)
   Dim i&, pos%, Sql$
   Dim MyNode As ComctlLib.Node
   
   mvarLstView.ListItems.Clear
   pIDROOT = ClsDsr.EliminarString(CStr(pIDROOT), "'")
   Call M_ExpandNode(mvarTree.Nodes("'" & pIDROOT & "'"))
   
   pos = InStr(UCase(mvarSqlList), "WHERE")
   If pos > 0 Then
      Sql = Mid(mvarSqlList, 1, pos - 1)
      Sql = Sql & " where " & Me.fldROOT & "='" & CStr(pIDROOT) & "'"
      Sql = Sql & " and " & Mid(mvarSqlList, pos + 5)
   Else
      Sql = mvarSqlList & " where " & Me.fldROOT & "='" & CStr(pIDROOT) & "'"
   End If
   Call DB.AbreTabela(Sql)
   If pIDROOT <> "" Then
'      Call ClsCtrl.MontarDbList(DB.RS, mvarLstView, Sql, mvarCabList, mvarPrikey)
   Else
      'Call RefreshList(DB, Me.LstPiece, Sql, Cab, "IDPIECE")
   End If

End Sub
Public Sub M_PopulaListView0(ByRef Node As ComctlLib.Node)
   Dim i&
   Dim MyNode As ComctlLib.Node
   
   Call M_ExpandNode(Node)
   mvarLstView.ListItems.Clear
   
   If Node.Children = 0 Then
      Set MyNode = Node
      mvarLstView.ListItems.Add 1, ClsDsr.Aspas(MyNode.Index), MyNode.Text, MyNode.Image, MyNode.Image
      mvarTree.Nodes(MyNode.Index).Selected = True
   Else
      mvarLstView.ListItems.Clear
      For i = 1 To Node.Children
         If i = 1 Then
            Set MyNode = Node.Child
         Else
            Set MyNode = MyNode.Next
         End If
         mvarLstView.ListItems.Add i, ClsDsr.Aspas(MyNode.Index), MyNode.Text, MyNode.Image, MyNode.Image
      Next
   '   mvarTree.Nodes(Node.Child.Index).Selected = True
   End If

End Sub
Public Function M_LocalizarNo%(Id$, Optional Selecionar = True)
   Dim i%, Ind%, Cont&
   Dim Tabela$
   On Error Resume Next
   M_LocalizarNo% = mvarTree.SelectedItem.Index
   Cont = mvarTree.Nodes.Count
   
   Dim MySql$, MyCol As New Collection
   Dim MyTable$, FimArq As Boolean
   Dim MyRoot$
Set MyCol = Nothing
Set MyCol = New Collection
   MyRoot$ = "'" & Id & "'"
   Ind = 1
   MyCol.Add MyRoot$
   
   i = InStr(UCase(mvarSql), " FROM ") + 6
   Tabela = Mid(mvarSql, i)
   i = InStr(UCase(Tabela), " ")
      
   Tabela = Trim(IIf(i = 0, Tabela, Mid(Tabela, 1, i)))
   MySql = "select " & mvarPrikey & ", " & mvarFldRoot
   MySql = MySql & " From  " & Tabela
   MySql = MySql & " Where " & mvarPrikey & "="
   FimArq = False
   While Not FimArq
      FimArq = True
      Call mvarDB.AbreTabela(MySql & MyRoot$)
      If mvarDB.CodeSql Then
         MyRoot$ = "'" & mvarDB.Dys(1) & "" & "'"
         FimArq = (mvarDB.Dys(1) = "")
         If Not FimArq Then
            Ind = Ind + 1
            MyCol.Add MyRoot$
         End If
      End If
      If Ind > 100 Then
         Err = 1
         Call ClsMsg.ShowError("M_LocalizarNo( " & Tabela & "." & mvarFldRoot & " = " & MyRoot$ & " )")
         M_LocalizarNo% = 0
         Exit Function
      End If
   Wend
   For i = Ind To 1 Step -1
      Call Me.M_ExpandNode(mvarTree.Nodes(MyCol(i)))
   Next
   Id$ = "'" & Id$ & "'"
   M_LocalizarNo% = mvarTree.Nodes(Id$).Index
   If Selecionar Then
      mvarTree.Nodes(Id$).EnsureVisible
      mvarTree.SelectedItem = mvarTree.Nodes(Id$)
   End If
'Exit Function
'   i = 1
'   While i <= Cont
'      If mvarTree.Nodes(i).Key <> "" Then
'         If mvarTree.Nodes(i).Key = ClsDsr.Aspas(Mid(Trim(Id), 1, Len(mvarTree.Nodes(i).Key) - 2)) Then
'            Call Me.M_ExpandNode(mvarTree.Nodes(i))
'            Cont = mvarTree.Nodes.Count
'            'mvarTree.Nodes(i).Expanded
'         End If
'         If mvarTree.Nodes(i).Key = "'" + Trim(Id) + "'" Then
'            M_LocalizarNo% = i
'            i = Cont
'         End If
'      End If
'      i = i + 1
'   Wend
End Function
Public Sub BoldTreeNode(TView As Object, pNode As Object)
'
' Make a tree node bold
'
' Many thanks to VBNet for this code
'
   On Error GoTo Fim

   Dim tvi As TVItem
   Dim lRet As Long, hItemTV As Long, lHwnd As Long
   
   Set TView.SelectedItem = pNode
   
   lHwnd = TView.hWnd
   hItemTV = SendMessageLong(lHwnd, TVM_GETNEXTItem, TVGN_CARET, 0&)
   
   If hItemTV > 0 Then
       With tvi
           .hItem = hItemTV
           .Mask = TVIF_State
           .stateMask = TVIS_BOLD
           lRet = SendMessageAny(lHwnd, TVM_GETItem, 0&, tvi)
           .State = TVIS_BOLD
       End With
       lRet = SendMessageAny(lHwnd, TVM_SETItem, 0&, tvi)
   End If
   
   Exit Sub
Fim:
   ClsMsg.ShowError
End Sub
'Public Function BeginDrag(TreeView As Object, ItemDrag As Long, NodeDrag As Object, PosDown As POINTAPI) As Long
Public Function BeginDrag(TreeView As Object, ItemDrag As Long, NodeDrag As Object, x As Long, y As Long) As Long
  Dim PosDown As PointAPI
  Dim pt As PointAPI
  Dim TvHti As TVHITTESTINFO
  Dim RcItem As RECT
  Dim SzIcon As SIZE
  Dim PtImage As PointAPI

Dim hImageDrag   As Long
Dim m_szDrag As SIZE         ' x and y distance cursor moves before dragging begins, in pixels
Dim m_ptHotSpot As PointAPI  ' x and y position of the cursor relative to the drag image origin, in pixels
Dim TreehWnd As Long
PosDown.x = x
PosDown.y = y

TreehWnd = TreeView.hWnd
' Store thet distance the cursor moves to initiate dragging.
m_szDrag.cx = GetSystemMetrics(SM_CXDRAG)
m_szDrag.cy = GetSystemMetrics(SM_CYDRAG)
  
  
#If DBG Then
Debug.Print "BeginDrag"
#End If
  
  ' If the cursor position has exceeded the non-drag slop...
  Call GetCursorPos(pt)
  If (Abs(pt.x - PosDown.x) >= m_szDrag.cx) Or _
     (Abs(pt.y - PosDown.y) >= m_szDrag.cy) Then
    
    ' We're dragging, get the TreeView client coords at MouseDown
    pt.x = PosDown.x
    pt.y = PosDown.y
    Call ScreenToClient(TreehWnd, pt)
    
    ' Get the handle of the Item under the cursor at MouseDown.
    TvHti.pt = pt
    'Call TreeView_HitTest(TreehWnd, TvHti)
    Call SendMessage(TreehWnd, TVM_HITTEST, 0, TvHti)
    
    If (TvHti.Flags And TVHT_ONItem) Then
      
      ' Allow dragging of only non-root Items.
      If TreeView_GetParent(TreehWnd, TvHti.hItem) Then
        
        ' Store the handle of the Item being dragged.
        ItemDrag = TvHti.hItem
        
        ' The TreeView_CreateDragImage call fails if the TreeView does not have
        ' an imagelist associated with it, the call succeds but the returned imagelist
        ' contains no drag image if the specified Item is not assigned an icon, and
        ' for some reason the call also fails if the TreeView Item's iImage member is
        ' using callbacks (I_IMAGECALLBACK, but is not the case for real treeviews...?)
        Call SetNodeCallbacks(TreehWnd, ItemDrag, NodeDrag, False)
        hImageDrag = TreeView_CreateDragImage(TreehWnd, ItemDrag)
        Call SetNodeCallbacks(TreehWnd, ItemDrag, NodeDrag, True)
        If hImageDrag Then
          
'Debug.Print ImageList_GetImageCount(hImageDrag)
'Call ImageList_Draw(hImageDrag, 0, Picture1.hDC, 0, 0, ILD_NORMAL)
          
          ' Get the position of the cursor relative to the image's origin,
          ' for ImageList_BeginDrag, is used to postion the drag image
          ' from the window coords passed to ImageList_DragEnter and
          ' ImageList_DragMove (szIcon.cx is the width of the actual
          ' drag image [the Item's icon and label])

         '*Call TreeView_GetItemRect(TreehWnd, ItemDrag, RcItem, CTrue)
          RcItem.Left = ItemDrag
          Call SendMessage(TreehWnd, TVM_GETItemRECT, ByVal CTrue, RcItem)
          
          Call ImageList_GetIconSize(hImageDrag, SzIcon.cx, SzIcon.cy)
          m_ptHotSpot.x = (pt.x - RcItem.Left) + (SzIcon.cx - (RcItem.Right - RcItem.Left))
          m_ptHotSpot.y = pt.y - RcItem.Top
          
          ' Convert the Item label's origin, which is relative to the
          ' Treeview's client area, to coods relative to the TreeView's
          ' window rect origin, for ImageList_DragEnter
          
'Call ClientToWindow(TreehWnd, PtImage)
'Public Function ClientToWindow(hwnd As Long, Pt As Variant) As Boolean
  Dim fRtn As Boolean
  Dim rcClient As RECT
  Dim rcWindow As RECT
  
  If IsWindow(TreehWnd) Then
    fRtn = CBool(GetClientRect(TreehWnd, rcClient))
    fRtn = fRtn And CBool(ClientToScreen(TreehWnd, rcClient))
    fRtn = fRtn And CBool(GetWindowRect(TreehWnd, rcWindow))
    If fRtn Then
      PtImage.x = PtImage.x + (rcClient.Left - rcWindow.Left)
      PtImage.y = PtImage.y + (rcClient.Top - rcWindow.Top)
'      ClientToWindow = True
    End If
  End If
          
          
          ' Set capture so we get MouseMoves when dragging
          ' outside the TreeView,
          Call SetCapture(TreehWnd)
          
          ' Select the drop target
          Call TreeView_SelectDropTarget(TreehWnd, ItemDrag)
          
          ' Set the drag imagelist and the image hotspot
          Call ImageList_BeginDrag(hImageDrag, 0, m_ptHotSpot.x, m_ptHotSpot.y)
          
          ' Lock the screen and draw the first drag image
          Call ImageList_DragEnter(TreehWnd, PtImage.x, PtImage.y)
        
        End If   ' hImageDrag
      End If   ' TreeView_GetParent
    End If   ' TreeView_HitTest
    
  End If   ' begin drag
  BeginDrag = hImageDrag
End Function

Public Sub DoDrag(TreehWnd As Long, Optional NoDrop = False)
  Dim pt As PointAPI
  Dim TvHti As TVHITTESTINFO
  Dim hItem As Long
  Dim rcClient As RECT
  Static hItemPrev As Long
  
#If DBG Then
Debug.Print "DoDrag"
#End If

  ' Get the cursor postion in TreeView client coords
  Call GetCursorPos(pt)
  Call ScreenToClient(TreehWnd, pt)
  
  ' Unlock the treeview's painting
  Call ImageList_DragLeave(TreehWnd)
  
  ' Highlights the new drop target if the cursor is over an Item,
  ' removes the drop highlight otherwise...
  TvHti.pt = pt
  'hItem = TreeView_HitTest(TreehWnd, TvHti)
  hItem = SendMessage(TreehWnd, TVM_HITTEST, 0, TvHti)
  Call TreeView_SelectDropTarget(TreehWnd, hItem)
  
  ' If the cursor is over an Item and TreeView has a horizontal scrollbar, assume
  ' that the Item's tooltip is shown, and update the TreeView on each drag move
  ' (for some reason the TreeView's tooltip does not erase correctly...).
  If CBool(TvHti.Flags And TVHT_ONItemLINE) And _
        (GetWindowLong(TreehWnd, GWL_Style) And WS_HSCROLL) Then
    Call UpdateWindow(TreehWnd)
  End If
  
  ' Convert the Item label's origin, which is relative to the
  ' Treeview's client area, to coods relative to the TreeView's
  ' window rect origin, for ImageList_DragEnter
  Call ClientToWindow(TreehWnd, pt)
  
  ' Erase the old drag image and draw a new one.
  Call ImageList_DragMove(pt.x, pt.y)
  
  ' Lock the treeview's painting again
  Call ImageList_DragEnter(TreehWnd, pt.x, pt.y)

  ' Modify the cursor to provide visual feedback to the user.
  ' Note: It's important to do this AFTER the call to DragMove. (so says Jeff...??)
  If hItem Then
    Screen.MousePointer = vbDefault
  Else

    Screen.MousePointer = vbNoDrop
  End If
  
  ' If the cursor is still over same Item as it was on the previous call,
  ' the cursor is over button, label, or icon of a collapsed parent Item,
  ' start the auto expand timer, disable the timer otherwise.
'''''''''  If (hItem = hItemPrev) And (TvHti.flags And (TVHT_ONItemBUTTON Or TVHT_ONItem)) And _
      IsTVItemCollapsedParent(TreehWnd, hItem) Then
'''''''''    tmrAutoExpand.Enabled = True
'''''''''  Else
'''''''''    tmrAutoExpand.Enabled = False
'''''''''  End If
  
  ' cache the current Item's handle as the previous Item's handle
  hItemPrev = hItem
  
  ' If the window is scrollable, and the cursor is within that auto scroll
  ' distance, start the auto scroll timer, disable the timer otherwise.
  Call GetClientRect(TreehWnd, rcClient)
'Debug.Print "scroll: &H" & Hex(IsWindowScrollable(hwndTV)) & _
                    " , region &H" & Hex(PtInRectRegion(rc, m_cxyAutoScroll, tvhti.pt))
'''''''''  If (IsWindowScrollable(TreehWnd) And _
          PtInRectRegion(rcClient, m_cxyAutoScroll, TvHti.Pt)) Then
'''''''''    tmrAutoScroll.Enabled = True
'''''''''  Else
'''''''''    tmrAutoScroll.Enabled = False
'''''''''  End If
  
End Sub
'*************************
'*************************
'*************************
'*************************
'


' Returns a set of bit flags indicating whether the specified
' window can be scrolled in any given direction.

'Public Function IsWindowScrollable(hwnd As Long) As ScrollDirectionFlags
Public Function IsWindowScrollable(hWnd As Long) As Variant
  Dim si As SCROLLINFO
  Dim dwScrollFlags As ScrollDirectionFlags
  
  si.cbSize = Len(si)
  si.fMask = SIF_ALL
  
  ' Get the horizontal scrollbar's info (GetScrollInfo returns
  ' TRUE after a scrollbar has been added to a window,
  ' even if the respective Style bit is not set...)
  If (GetWindowLong(hWnd, GWL_Style) And WS_HSCROLL) Then
    If GetScrollInfo(hWnd, SB_HORZ, si) Then
      dwScrollFlags = (sdLeft And (si.nPos > 0))
      dwScrollFlags = dwScrollFlags Or (sdRight And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
    End If
  End If
  
  ' Get the vertical scrollbar's info.
  If (GetWindowLong(hWnd, GWL_Style) And WS_VSCROLL) Then
    If GetScrollInfo(hWnd, SB_VERT, si) Then
      dwScrollFlags = dwScrollFlags Or (sdUp And (si.nPos > 0))
      dwScrollFlags = dwScrollFlags Or (sdDown And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
    End If
  End If
  
  IsWindowScrollable = dwScrollFlags

End Function
' Toggles the callback attributes (text, images, button) for the specific
' Item (Node), as specified by the fSet flag. The TreeView normally
' does callback for these Items, but can be explicitly set and overridden
' (and the VB TreeView has no idea what's happening).

Public Function SetNodeCallbacks(hWndTV As Long, _
                                                        hItem As Long, _
                                                        nod As Node, _
                                                        fSet As Boolean) As Boolean
  Dim tvi As TVItem
  
  tvi.Mask = TVIF_IMAGE  ' Or TVIF_SELECTEDIMAGE Or TVIF_TEXT  ' Or TVIF_CHILDREN
  tvi.hItem = hItem

  If fSet Then
    ' Set the callbacks
    tvi.iImage = I_IMAGECALLBACK
'    tvi.iSelectedImage = I_IMAGECALLBACK
'    tvi.pszText = LPSTR_TEXTCALLBACK
'' this causes problems, something's going on with the
'' VB TreeView here that has yet to be understood...
'    tvi.cChildren = I_CHILDRENCALLBACK
    
  Else
    ' Get the Node from the hItem, and remove the Item's
    ' callback attributes by explicitly setting them.
    If ((nod Is Nothing) = False) Then
      ' real imagelist indices are zero-based
'      tvi.iImage = nod.Image - 1
      tvi.iSelectedImage = CLng(nod.SelectedImage) - 1
'      ' Store the Node's Text in an allocated pointer'
'      tvi.pszText = StrPtr(String$(MAX_Item, 0))
'      Call lstrcpyA(ByVal tvi.pszText, ByVal nod.Text)
'' see above...
'      tvi.cChildren = Abs(CBool(TreeView_GetChild(m_hwndTV, hItem)))
    End If
  
  End If   ' fSet
  
  SetNodeCallbacks = TreeView_SetItem(hWndTV, tvi)
    
End Function

' Returns True if the specified treeview Item is a collapsed parent
' (has a buttom and is collapsed). returns False otherwise

Public Function IsTVItemCollapsedParent(hWndTV As Long, hItem As Long) As Boolean
  Dim tvi As TVItem
  
  tvi.hItem = hItem
  tvi.Mask = TVIF_State Or TVIF_CHILDREN
  
  If TreeView_GetItem(hWndTV, tvi) Then
    IsTVItemCollapsedParent = ((tvi.State And TVIS_EXPANDED) = False) And (tvi.cChildren <> 0)
  End If

End Function

' Returns True if hItemParent is an ancestor of hItemChild,
' returns False otherwise

Public Function AreTVItemsParentChild(hWndTV As Long, hItemParent As Long, ByVal hItemChild As Long) As Boolean
  
  Do While hItemChild
    hItemChild = TreeView_GetParent(hWndTV, hItemChild)
    If (hItemChild = hItemParent) Then
      AreTVItemsParentChild = True
      Exit Function
    End If
  Loop

End Function

' ===========================================================================
' Treeview macros defined in Commctrl.h

' Expands or collapses the list of child Items, if any, associated with the specified parent Item.
' Returns TRUE if successful or FALSE otherwise.
' (docs say TVM_EXPAND does not send the TVN_ItemEXPANDING and
' TVN_ItemEXPANDED notification messages to the parent window...?)

Public Function TreeView_Expand(hWnd As Long, hItem As Long, flag As Long) As Boolean
  TreeView_Expand = SendMessage(hWnd, TVM_EXPAND, ByVal flag, ByVal hItem)
End Function

' TreeView_GetNextItem

' Retrieves the tree-view Item that bears the specified relationship to a specified Item.
' Returns the handle to the Item if successful or 0 otherwise.

Public Function TreeView_GetNextItem(hWnd As Long, hItem As Long, flag As Long) As Long
  TreeView_GetNextItem = SendMessage(hWnd, TVM_GETNEXTItem, ByVal flag, ByVal hItem)
End Function
'
'' Retrieves the first child Item. The hItem parameter must be NULL.
'' Returns the handle to the Item if successful or 0 otherwise.
'
Public Function TreeView_GetChild(hWnd As Long, hItem As Long) As Long
  TreeView_GetChild = TreeView_GetNextItem(hWnd, hItem, TVGN_CHILD)
End Function

' Retrieves the parent of the specified Item.
' Returns the handle to the Item if successful or 0 otherwise.

Public Function TreeView_GetParent(hWnd As Long, hItem As Long) As Long
  TreeView_GetParent = TreeView_GetNextItem(hWnd, hItem, TVGN_PARENT)
End Function

' Retrieves the Item that is the target of a drag-and-drop operation.
' Returns the handle to the Item if successful or 0 otherwise.

Public Function TreeView_GetDropHilight(hWnd As Long) As Long
  TreeView_GetDropHilight = TreeView_GetNextItem(hWnd, 0, TVGN_DROPHILITE)
End Function

' Retrieves the topmost or very first Item of the tree-view control.
' Returns the handle to the Item if successful or 0 otherwise.

Public Function TreeView_GetRoot(hWnd As Long) As Long
  TreeView_GetRoot = TreeView_GetNextItem(hWnd, 0, TVGN_ROOT)
End Function

' TreeView_Select

' Selects the specified tree-view Item, scrolls the Item into view, or redraws the Item
' in the Style used to indicate the target of a drag-and-drop operation.
' If hItem is NULL, the selection is removed from the currently selected Item, if any.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_Select(hWnd As Long, hItem As Long, code As Long) As Boolean
  TreeView_Select = SendMessage(hWnd, TVM_SELECTItem, ByVal code, ByVal hItem)
End Function

' Redraws the given Item in the Style used to indicate the target of a drag and drop operation.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_SelectDropTarget(hWnd As Long, hItem As Long) As Boolean
  TreeView_SelectDropTarget = TreeView_Select(hWnd, hItem, TVGN_DROPHILITE)
End Function

' Retrieves some or all of a tree-view Item's attributes.
' Returns TRUE if successful or FALSE otherwise.

Private Function TreeView_GetItem(hWnd As Long, pItem As TVItem) As Boolean
'Public Function TreeView_GetItem(hwnd As Long, pItem As Variant) As Boolean
  TreeView_GetItem = SendMessage(hWnd, TVM_GETItem, 0, pItem)
End Function

' Sets some or all of a tree-view Item's attributes.
' Old docs say returns zero if successful or - 1 otherwise.
' New docs say returns TRUE if successful, or FALSE otherwise

Private Function TreeView_SetItem(hWnd As Long, pItem As TVItem) As Boolean
'Public Function TreeView_SetItem(hwnd As Long, pItem As Variant) As Boolean
  TreeView_SetItem = SendMessage(hWnd, TVM_SETItem, 0, pItem)
End Function

' Determines the location of the specified point relative to the client area of a tree-view control.
' Returns the handle to the tree-view Item that occupies the specified point or NULL if no Item
' occupies the point.

'Public Function TreeView_HitTest(hwnd As Long, lpht As TVHITTESTINFO) As Long
Public Function TreeView_HitTest(hWnd As Long, x As Long, y As Long) 'lpht lpht As Variant) As Long
  Dim pt As PointAPI
  Dim lpht As TVHITTESTINFO
  lpht.pt.x = x
  lpht.pt.y = y
  TreeView_HitTest = SendMessage(hWnd, TVM_HITTEST, 0, lpht)
  x = lpht.pt.x
  y = lpht.pt.y
End Function

' Creates a dragging bitmap for the specified Item in a tree-view control, creates an image list
' for the bitmap, and adds the bitmap to the image list. An application can display the image
' when dragging the Item by using the image list functions.
' Returns the handle of the image list to which the dragging bitmap was added if successful or
' NULL otherwise.

Public Function TreeView_CreateDragImage(hWnd As Long, hItem As Long) As Long
  TreeView_CreateDragImage = SendMessage(hWnd, TVM_CREATEDRAGIMAGE, 0, ByVal hItem)
End Function
Private Function GetTVBackColour(hWndTV As Long) As Long
   Dim ClrRef As Long
  
  'try for the treeview backcolor
   ClrRef = SendMessage(hWndTV, TVM_GETBKCOLOR, 0, ByVal 0)
   
  'if clrref = -1, then the color is a system color.
  'In theory, system colors need to be Or'd with &HFFFFFF
  'to retrieve the actual RGB value, but not Or'ing
  'seems to work for me. The default system colour for
  'a treeview background is COLOR_WINDOW.
   If ClrRef = -1 Then
      ClrRef = GetSysColor(COLOR_WINDOW)  ' Or &HFFFFFF
   End If
   
  'one way or another, pass it back
   GetTVBackColour = ClrRef
  
End Function
Public Function GetTVForeColour(hWndTV As Long) As Long
   Dim ClrRef As Long
   
  'try for the treeview text colour
   ClrRef = SendMessage(hWndTV, TVM_GETTEXTCOLOR, 0, ByVal 0)
   
  'if clrref = -1, then the color is a system color.
  'In theory, system colors need to be Or'd with &HFFFFFF
  'to retrieve the actual RGB value, but not Or'ing
  'seems to work for me. The default system colour for
  'treeview text is COLOR_WINDOWTEXT.
   If ClrRef = -1 Then
      ClrRef = GetSysColor(COLOR_WINDOWTEXT) ' Or &HFFFFFF
   End If
   
  'one way or another, pass it back
   GetTVForeColour = ClrRef
   
End Function
Public Sub SetTVBackColour(hWndTV As Long, ClrRef As Long)
   Dim Style As Long
  
  'Change the background
   Call SendMessage(hWndTV, TVM_SETBKCOLOR, 0, ByVal ClrRef)
   
  'reset the treeview Style so the
  'tree lines appear properly
   Style = GetWindowLong(hWndTV, GWL_Style)
   
  'if the treeview has lines, temporarily
  'remove them so the back repaints to the
  'selected colour, then restore
   If Style And TVS_HASLINES Then
      Call SetWindowLong(hWndTV, GWL_Style, Style Xor TVS_HASLINES)
      Call SetWindowLong(hWndTV, GWL_Style, Style)
   End If
 
End Sub
Public Sub SetTVForeColour(hWndTV As Long, ClrRef As Long)
   Dim Style As Long
   
  'Change the background
   Call SendMessage(hWndTV, TVM_SETTEXTCOLOR, 0, ByVal ClrRef)
   
  'reset the treeview Style so the
  'tree lines appear properly
   Style = GetWindowLong(hWndTV, GWL_Style)
   
  'if the treeview has lines, temporarily
  'remove them so the back repaints to the
  'selected colour, then restore
   If Style And TVS_HASLINES Then
      Call SetWindowLong(hWndTV, GWL_Style, Style Xor TVS_HASLINES)
      Call SetWindowLong(hWndTV, GWL_Style, Style)
   End If
End Sub
