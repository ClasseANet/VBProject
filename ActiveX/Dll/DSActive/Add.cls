VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DsAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private WithEvents mvarMe As FrmADD
Attribute mvarMe.VB_VarHelpID = -1

Private mvarXDb         As DS_BANCO

Private mvarPreSelecao  As Collection
Private mvarSelecao     As Collection
Private mvarRS          As Object
Private mvarMaxTamId    As Integer
Private mvarCheckBox    As Boolean
Private mvarCaption     As String
Private mvarQuery       As String

Private mvarMask        As String
Private mvarItemMask    As String

Private mvarTitulo      As Collection
Private mvarChaves      As Collection
Public Property Let Mask(ByVal vData As String)
    mvarMask = vData
End Property
Public Property Get Mask() As String
    Mask = mvarMask
End Property
Public Property Let ItemMask(ByVal vData As String)
    mvarItemMask = vData
End Property
Public Property Get ItemMask() As String
    ItemMask = mvarItemMask
End Property
Public Property Get Titulo(Optional ByVal Index) As Titulo
   Dim MyTitulo As Titulo
   If IsMissing(Index) Then
      If mvarTitulo.Count = 0 Then
         Index = 1
      Else
         Set MyTitulo = New Titulo
         Set MyTitulo = mvarTitulo(1)
         Set mvarTitulo = MyTitulo.Item
         mvarTitulo(1).Item = MyTitulo.Item
         Set Titulo = mvarTitulo(1)
         Set MyTitulo = Nothing
         Exit Property
      End If
   End If
   If mvarTitulo Is Nothing Then
      Set mvarTitulo = New Collection
   End If
   
   If Index <= mvarTitulo.Count + 1 Then
      On Error Resume Next
      Set Titulo = mvarTitulo(Index)
      If Err <> 0 Then
         Err = 0
         Set MyTitulo = New Titulo
         If mvarTitulo.Count > 0 Then
            MyTitulo.Item = mvarTitulo(1).Item
            MyTitulo.Item.Add MyTitulo, CStr(mvarTitulo.Count + 1)
            Set mvarTitulo = MyTitulo.Item
         Else
            MyTitulo.Item.Add MyTitulo, CStr(mvarTitulo.Count + 1)
            mvarTitulo.Add MyTitulo, CStr(mvarTitulo.Count + 1)
         End If
         
         
         Set MyTitulo = Nothing
         
         Set Titulo = mvarTitulo(Index)
      End If
   Else
      MsgBox "Item inválido." & vbNewLine & "Próximo Item deve ser " & CStr(mvarTitulo.Count + 1)
      Set Titulo = Nothing
   End If
End Property
Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property
Public Property Get Caption() As String
    Caption = mvarCaption
End Property
Public Property Let Query(ByVal vData As String)
    mvarQuery = vData
End Property
Public Property Get Query() As String
    Query = mvarQuery
End Property
Public Property Let CheckBox(ByVal vData As Integer)
    mvarCheckBox = vData
End Property
Public Property Get CheckBox() As Integer
    CheckBox = mvarCheckBox
End Property
Public Property Let MaxTamId(ByVal vData As Integer)
    mvarMaxTamId = vData
End Property
Public Property Get MaxTamId() As Integer
    MaxTamId = mvarMaxTamId
End Property
Public Property Set Rs(ByVal vData As Object)
    Set mvarRS = vData
End Property
Public Property Get Rs() As Object
    Set Rs = mvarRS
End Property
Public Property Set PreSelecao(ByVal vData As Collection)
    Set mvarPreSelecao = vData
End Property
Public Property Get PreSelecao() As Collection
    Set PreSelecao = mvarPreSelecao
End Property
Public Property Set Selecao(ByVal vData As Collection)
    Set mvarSelecao = vData
End Property
Public Property Get Selecao() As Collection
    Set Selecao = mvarSelecao
End Property
Public Property Set Chaves(ByVal vData As Collection)
    Set mvarChaves = vData
End Property
Public Property Get Chaves() As Collection
    Set Chaves = mvarChaves
End Property
Public Property Set XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Let XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Get XDb() As Object
   Set XDb = mvarXDb
End Property
Public Sub Show()
   If mvarRS Is Nothing Then
      If Not mvarXDb Is Nothing And mvarQuery <> "" Then
         If mvarXDb.AbreTabela(mvarQuery, mvarRS) Then
            mvarMe.Show vbModal
         Else
            Call ClsMsg.ExibirAviso("Não existem itens a serem exibidos.", mvarCaption)
         End If
      Else
         Call ClsMsg.ExibirAviso("Não existem itens a serem exibidos.", mvarCaption)
      End If
   Else
      mvarMe.Show vbModal
   End If
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmADD
      
   Set mvarTitulo = New Collection
   Set mvarChaves = New Collection
   Set mvarPreSelecao = New Collection
   Set mvarSelecao = New Collection
End Sub
Private Sub Class_Terminate()
   Set Selecao = Nothing
   Set mvarTitulo = Nothing
   Set mvarChaves = Nothing
   Set mvarPreSelecao = Nothing
End Sub
Private Sub mvarMe_CmdChkClick(Index As Integer)
   Dim i As Long
  
   For i = 1 To mvarMe.LstItens.ListItems.Count
      Select Case Index
         Case 0
            If mvarMe.LstItens.Checkboxes Then
               mvarMe.LstItens.ListItems(i).Checked = True
            End If
            mvarMe.LstItens.ListItems(i).Selected = True
         Case 1
            If mvarMe.LstItens.Checkboxes Then
               mvarMe.LstItens.ListItems(i).Checked = False
            End If
            mvarMe.LstItens.ListItems(i).Selected = False
         Case 2
            mvarMe.LstItens.ListItems(i).Checked = Not mvarMe.LstItens.ListItems(i).Checked
            mvarMe.LstItens.ListItems(i).Selected = mvarMe.LstItens.ListItems(i).Checked
         Case 3
            If mvarMe.LstItens.ListItems(i).Selected Then
               mvarMe.LstItens.ListItems(i).Checked = True
            End If
      End Select
   Next
End Sub

Private Sub mvarMe_CmdOperClick(Index As Integer)
   Dim i As Integer
   Dim j As Integer
   Dim Selecionado   As Boolean
   
   Select Case Index
      Case 0
         Set Selecao = New Collection
         For i = 1 To mvarMe.LstItens.ListItems.Count
            Selecionado = False
            If mvarCheckBox Then
               Selecionado = mvarMe.LstItens.ListItems(i).Checked
            Else
               Selecionado = mvarMe.LstItens.ListItems(i).Selected
            End If
            If Selecionado Then
               Dim MyColl As Collection
               Dim sCod    As String
               
               Set MyColl = New Collection
               sCod = Mid(mvarMe.LstItens.ListItems(i).Key, 2)
               j = 0
               While InStr(sCod, ".") <> 0
                  j = j + 1
                  MyColl.Add Mid(sCod, 1, InStr(sCod, ".") - 1), mvarChaves(j)
                  sCod = Mid(sCod, InStr(sCod, ".") + 1)
               Wend
               j = j + 1
               If mvarChaves.Count >= j Then
                  MyColl.Add sCod, mvarChaves(j)
               End If
               
               Selecao.Add MyColl
            End If
         Next
      
      Case 1
         Set Selecao = New Collection
   End Select
   Unload mvarMe
End Sub
Private Sub mvarMe_Activate()
'   Call MontaTela
End Sub
Private Sub mvarMe_Load()
   If mvarCaption = "" Then mvarCaption = Me.Caption
   Call MontaTela
   
   Dim bAux As Boolean
   bAux = mvarMe.LstItens.MultiSelect
   mvarMe.LstItens.MultiSelect = False
   mvarMe.LstItens.ListItems(1).Selected = True
   mvarMe.LstItens.MultiSelect = bAux
   
   mvarMe.CmdChk(0).Visible = mvarMe.LstItens.MultiSelect
   mvarMe.CmdChk(1).Visible = mvarMe.LstItens.MultiSelect
   mvarMe.CmdChk(2).Visible = mvarMe.LstItens.MultiSelect
   mvarMe.CmdChk(3).Visible = mvarMe.LstItens.MultiSelect
   mvarMe.CmdChk(3).Visible = mvarMe.LstItens.MultiSelect And mvarMe.LstItens.Checkboxes
End Sub
Private Sub MontaTela()
   Dim MaxTam  As Integer
   Dim sCod    As String
   Dim sDesc   As String
   Dim n       As Variant
   Dim ExibeCab As Boolean
   
   
   Dim b As ADODB.Recordset
   
   On Error GoTo TrataErro
   
   If mvarMask <> "" Then
      mvarMask = ClsDsr.StrReplace(mvarMask, "9", "&")
      mvarMask = ClsDsr.StrReplace(mvarMask, "A", "&")
      mvarMask = ClsDsr.StrReplace(mvarMask, "X", "&")
   End If
   
   mvarMe.Caption = mvarCaption
   
   mvarMe.CmbCampo.Clear
   mvarMe.CmbCampo.Text = ""
   mvarMe.TxtChave.Text = ""
   mvarMe.CmbCampo.Visible = False
   mvarMe.TxtChave.Visible = False
   mvarMe.CmdFiltrar.Visible = False
   mvarMe.CmdLocalizar.Visible = False
   
   
   With mvarMe.LstItens
      .ListItems.Clear
      .FullRowSelect = True
      .GridLines = True
      .LabelEdit = lvwManual
      .HideSelection = False
      .MultiSelect = True
      .View = lvwReport
      .Checkboxes = mvarCheckBox
      '.Sorted = True
   End With
   
   If Titulo.Count = 0 Then
      If mvarMaxTamId = 0 And mvarRS.Fields.Count > 1 Then
          mvarRS.MoveFirst
         While Not mvarRS.EOF
            MaxTam = IIf(MaxTam >= mvarRS(0).ActualSize, MaxTam, mvarRS(0).ActualSize)
            mvarRS.MoveNext
         Wend
      Else
         MaxTam = mvarMaxTamId
      End If
   
      mvarMe.LstItens.ColumnHeaders.Clear
      mvarMe.LstItens.ColumnHeaders.Add , , "", mvarMe.LstItens.Width - 360
      mvarMe.LstItens.HideColumnHeaders = True
      mvarRS.MoveFirst
      While Not mvarRS.EOF
         If mvarRS.Fields.Count = 1 Then
            sCod = mvarRS(0)
            mvarMe.LstItens.ListItems.Add , "k" & sCod, sCod
         Else
            sCod = Mid(mvarRS(0) & Space(MaxTam), 1, MaxTam)
            sDesc = mvarRS(1)
            mvarMe.LstItens.ListItems.Add , "k" & sCod, sCod & " - " & sDesc
         End If
         mvarRS.MoveNext
      Wend
   Else
      mvarMe.LstItens.Left = 120
      mvarMe.LstItens.ColumnHeaders.Clear
      mvarMe.LstItens.HideColumnHeaders = False
      mvarMe.LstItens.Width = 360
      ExibeCab = False
      For Each n In mvarTitulo
         If Trim(n.Campo) <> "" Then
            mvarMe.LstItens.ColumnHeaders.Add , n.Campo, n.Texto, n.Tamanho * 120, n.Alinhamento
            mvarMe.LstItens.Width = mvarMe.LstItens.Width + (n.Tamanho * 120) '((n.Tamanho * 120) / Screen.TwipsPerPixelX)
            mvarMe.CmbCampo.AddItem n.Texto
         End If
      Next
      If mvarMe.CmbCampo.ListCount > 0 Then mvarMe.CmbCampo.ListIndex = 0
      
      mvarMe.LstItens.Width = IIf(mvarMe.LstItens.Width < 6375, 6375, mvarMe.LstItens.Width)
      mvarMe.Width = mvarMe.LstItens.Width + 360
      If mvarMe.Width < mvarMe.CmdOper(0).Left + mvarMe.CmdOper(1).Width + 240 Then
         mvarMe.Width = mvarMe.CmdOper(0).Left + mvarMe.CmdOper(1).Width + 240
      End If

      If mvarMe.LstItens.ColumnHeaders.Count = 0 Then
         mvarMe.LstItens.ColumnHeaders.Add , ""
         mvarMe.LstItens.ColumnHeaders(1).Width = mvarMe.LstItens.Width - 360
      End If
      
      mvarRS.MoveFirst
      While Not mvarRS.EOF
         If mvarRS.Fields.Count = 1 Then
            sCod = mvarRS(0)
            mvarMe.LstItens.ListItems.Add , sCod, sCod
         Else
            sCod = ""
            If mvarChaves Is Nothing Then
               Set mvarChaves = New Collection
            End If
            If mvarChaves.Count = 0 Then
               sCod = mvarRS(0) & ""
               sCod = sCod & "."
            Else
               For Each n In mvarChaves
                  sCod = sCod & mvarRS(n) & "."
               Next
            End If
            If Len(sCod) > 1 Then
               sCod = Mid(sCod, 1, Len(sCod) - 1)
            End If
            sCod = "k" & sCod
            For i = 1 To mvarTitulo.Count
               If mvarTitulo(i).Campo = "" Then
                  If i = 1 Then
                     mvarMe.LstItens.ListItems.Add , sCod, mvarRS(i)
                  Else
                     mvarMe.LstItens.ListItems(sCod).SubItems(i - 1) = mvarRS(i - 1)
                  End If
               Else
                  sValor = mvarRS(mvarTitulo(i).Campo)
                  If mvarTitulo(i).Campo = ItemMask And mvarMask <> "" Then
                     sValor = Mid(sValor & Space(Len(ClsDsr.UnFormat(mvarMask))), 1, Len(ClsDsr.UnFormat(mvarMask)) + 1)
                     sValor = Format(sValor, mvarMask)
                  End If

                  If i = 1 Then
                     mvarMe.LstItens.ListItems.Add , sCod, sValor
                  Else
                     mvarMe.LstItens.ListItems(sCod).SubItems(i - 1) = sValor
                  End If
               End If
            Next
            If mvarCheckBox Then
               If ClsDsr.ExisteItem(mvarPreSelecao, sCod) Then
                  mvarMe.LstItens.ListItems(sCod).Checked = True
               End If
            End If
         End If
         mvarRS.MoveNext
      Wend
   End If
   If mvarMe.LstItens.ListItems.Count > 0 Then
      mvarMe.LstItens.ListItems(1).Selected = True
   End If
   Exit Sub
TrataErro:
   Call ClsMsg.ShowError
End Sub
