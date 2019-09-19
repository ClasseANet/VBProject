VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DS_ItensRelacao"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe As FrmItensRelacionados
Attribute mvarMe.VB_VarHelpID = -1

Private mvarXDb         As Object
Private mvarQuery       As String
Private mvarQueryAdd    As String
Private mvarLblItens    As String
Private mvarICON        As Variant
Private mvarFundoTela   As String
Private mvarCaption     As String

Private mvarCollItens   As Collection

Private mvarMask                 As String
Private mvarItemMask             As String
Private mvarMultiSelectItensRel  As Boolean

Private mvarTitulo      As Collection
Private mvarChaves      As Collection

Private mvarisDirt      As Boolean
Public Property Let isDirt(ByVal vData As Boolean)
    mvarisDirt = vData
End Property
Public Property Get isDirt() As Boolean
    isDirt = mvarisDirt
End Property
Public Property Let ItemMask(ByVal vData As String)
    mvarItemMask = vData
End Property
Public Property Get ItemMask() As String
    ItemMask = mvarItemMask
End Property
Public Property Let MultiSelectItensRel(ByVal vData As Boolean)
    mvarMultiSelectItensRel = vData
End Property
Public Property Get MultiSelectItensRel() As Boolean
     MultiSelectItensRel = mvarMultiSelectItensRel
End Property
Public Property Get Titulo(Optional ByVal index) As Titulo
   Dim MyTitulo As Titulo
   
   If IsMissing(index) Then
      If mvarTitulo.Count = 0 Then
         index = 1
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
   
   If index <= mvarTitulo.Count + 1 Then
      On Error Resume Next
      Set Titulo = mvarTitulo(index)
         
      If Err <> 0 Or Titulo Is Nothing Then
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
         
         Set Titulo = mvarTitulo(index)
      End If
   Else
      MsgBox "Item inválido." & vbNewLine & "Próximo Item deve ser " & CStr(mvarTitulo.Count + 1)
      Set Titulo = Nothing
   End If
End Property
Public Property Set Chaves(ByVal vData As Collection)
    Set mvarChaves = vData
End Property
Public Property Get Chaves() As Collection
    Set Chaves = mvarChaves
End Property
Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property
Public Property Get Caption() As String
    Caption = mvarCaption
End Property
Public Property Let FundoTela(ByVal vData As String)
    mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
    FundoTela = mvarFundoTela
End Property
Public Property Let Icon(ByVal vData As Variant)
    mvarICON = vData
End Property
Public Property Get Icon() As Variant
    Icon = mvarICON
End Property
Public Property Set XDb(ByVal vData As Object)
    Set mvarXDb = vData
End Property
Public Property Get XDb() As Object
    Set XDb = mvarXDb
End Property
Public Property Set CollItens(ByVal vData As Collection)
   Set mvarCollItens = vData
End Property
Public Property Get CollItens() As Collection
   Set CollItens = mvarCollItens
End Property
Public Property Let LblItens(ByVal vData As String)
   mvarLblItens = vData
End Property
Public Property Get LblItens() As String
   LblItens = mvarLblItens
End Property
Public Property Let Mask(ByVal vData As String)
   mvarMask = vData
End Property
Public Property Get Mask() As String
   Mask = mvarMask
End Property
Public Property Let QueryAdd(ByVal vData As String)
   mvarQueryAdd = vData
End Property
Public Property Get QueryAdd() As String
   QueryAdd = mvarQueryAdd
End Property
Public Property Let Query(ByVal vData As String)
   mvarQuery = vData
End Property
Public Property Get Query() As String
   Query = mvarQuery
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmItensRelacionados
   Set mvarChaves = New Collection
   Set mvarTitulo = New Collection
   mvarFundoTela = "FUNDO"
   mvarCaption = "Itens Relacionados"
   mvarisDirt = False
End Sub
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Terminate()
   Set mvarTitulo = Nothing
   Set mvarChaves = Nothing
End Sub
Private Sub mvarMe_Activate()
'   CarregarItens
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdADDClick()
   Dim MyAdd   As New DsAdd
   Dim n       As Variant
   Dim Myrs    As Object
   Dim sCod    As String
   Dim sKey    As String
   Dim m       As Variant
   
   
   mvarMe.CmdADD.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   With MyAdd
      If mvarXDb.AbreTabela(mvarQueryAdd, Myrs) Then
         Set .Rs = Myrs
         Set .XDb = mvarXDb
         .CheckBox = True
         .Caption = mvarCaption
         .Query = mvarQueryAdd
         .Mask = mvarMask
         .ItemMask = mvarItemMask
         Set .Chaves = mvarChaves
         For Each n In mvarTitulo
            With .Titulo
               Call .Add(n.Campo, n.Texto, n.Tamanho, n.Alinhamento)
            End With
         Next
         .Show
         
         On Error Resume Next
         For Each n In .Selecao
            sKey = ""
            For Each m In mvarChaves
               sCod = CStr(n(m))
               sKey = sKey & CStr(m) & "="
               Select Case MyAdd.Rs(CStr(m)).Type
                  Case adVarChar: sKey = sKey & SqlStr(CStr(n(m)))
                  Case Else: sKey = sKey & CStr(n(m))
               End Select
               sKey = sKey & " And "
            Next
            sKey = Mid(sKey, 1, Len(sKey) - Len(" And "))
            .Rs.MoveFirst
            .Rs.Find sKey
            If .Rs.AbsolutePosition > 0 Then
               Call AddItem(.Rs)
               mvarisDirt = True
               mvarMe.LstItens.ListItems("k" & sCod).Selected = True
               mvarMe.LstItens.ListItems("k" & sCod).EnsureVisible
            End If
         Next
      End If
   End With
   mvarMe.CmdADD.MousePointer = 99
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdOperClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdREMOVEClick()
   Dim i As Integer, j As Integer
   For i = mvarMe.LstItens.ListItems.Count To 1 Step -1
      If mvarMe.LstItens.ListItems(i).Selected Then
         mvarCollItens.Remove Mid(mvarMe.LstItens.ListItems(i).Key, 2)
         mvarMe.LstItens.ListItems.Remove i
         mvarisDirt = True
      End If
   Next
End Sub
Private Sub mvarMe_Load()
   Call ClsCtrl.ListView_FullRowSelect(mvarMe.LstItens.hWnd)
   
   If Trim(mvarLblItens) <> "" Then mvarMe.Lbl.Caption = mvarLblItens
   mvarMe.Caption = mvarCaption
   Call MontaTela
   Call CarregarItens
   Call ClsCtrl.ConfigForm(mvarMe, mvarICON, mvarFundoTela)
End Sub
Private Sub MontaTela()
   Dim n As Variant
   If Titulo.Count = 0 Then
      mvarMe.LstItens.ColumnHeaders.Clear
      mvarMe.LstItens.ColumnHeaders.Add , , "", mvarMe.LstItens.Width - 360
      mvarMe.LstItens.HideColumnHeaders = True
   Else
      mvarMe.LstItens.Left = 120
      mvarMe.LstItens.ColumnHeaders.Clear
      mvarMe.LstItens.HideColumnHeaders = False
      mvarMe.LstItens.Width = 360
      For Each n In mvarTitulo
         If Trim(n.Campo) <> "" Then
            mvarMe.LstItens.ColumnHeaders.Add , n.Campo, n.Texto, n.Tamanho * 120, n.Alinhamento
            mvarMe.LstItens.Width = mvarMe.LstItens.Width + (n.Tamanho * 120) '((n.Tamanho * 120) / Screen.TwipsPerPixelX)
'            mvarMe.CmbCampo.AddItem n.Texto
         End If
      Next
'      If mvarMe.CmbCampo.ListCount > 0 Then
'         mvarMe.CmbCampo.ListIndex = 0
'      End If
      
      mvarMe.LstItens.Width = IIf(mvarMe.LstItens.Width < 6375, 6375, mvarMe.LstItens.Width)
      mvarMe.Width = mvarMe.LstItens.Width + 360

      If mvarMe.LstItens.ColumnHeaders.Count = 0 Then
         mvarMe.LstItens.ColumnHeaders.Add , ""
         mvarMe.LstItens.ColumnHeaders(1).Width = mvarMe.LstItens.Width - 360
      End If
   End If
End Sub
Private Sub CarregarItens()
   Dim Myrs As Object
   Dim i    As Integer
   
   mvarMe.LstItens.ListItems.Clear
   Set mvarCollItens = New Collection
   If mvarMask <> "" Then
      mvarMask = ClsDsr.StrReplace(mvarMask, "9", "&")
      mvarMask = ClsDsr.StrReplace(mvarMask, "A", "&")
      mvarMask = ClsDsr.StrReplace(mvarMask, "X", "&")
   End If
   With mvarMe.LstItens
      .ListItems.Clear
      .FullRowSelect = True
      .GridLines = False
      .LabelEdit = lvwManual
      .HideSelection = False
      .MultiSelect = mvarMultiSelectItensRel
      .View = lvwReport
      .Checkboxes = False
      .Sorted = True
   End With

   If mvarXDb.AbreTabela(mvarQuery, Myrs) Then
      Myrs.MoveFirst
      While Not Myrs.EOF
         Call AddItem(Myrs)
         Myrs.MoveNext
      Wend
   End If
   
   If mvarMe.LstItens.ListItems.Count > 0 Then
      mvarMe.LstItens.ListItems(1).Selected = True
      Call ClsCtrl.ListView_FullRowSelect(mvarMe.LstItens.hWnd)
   End If
   mvarMe.LstItens.Width = 360
   For i = 1 To mvarMe.LstItens.ColumnHeaders.Count
      mvarMe.LstItens.Width = mvarMe.LstItens.Width + mvarMe.LstItens.ColumnHeaders(i).Width
   Next
   mvarMe.CmdOper.Left = mvarMe.LstItens.Left + mvarMe.LstItens.Width + 120
   mvarMe.CmdADD.Left = mvarMe.CmdOper.Left
   mvarMe.CmdREMOVE.Left = mvarMe.CmdOper.Left
   mvarMe.Width = mvarMe.CmdOper.Left + mvarMe.CmdOper.Width + 120
   mvarMe.Width = IIf(mvarMe.Width > Screen.Width, Screen.Width, mvarMe.Width)
End Sub
Private Sub AddItem(ByRef Myrs As ADODB.Recordset)
   Dim sKey    As String
   Dim sCod    As String
   Dim sValor  As String
   Dim i       As Integer
   Dim n       As Variant
   Dim nAux    As Long
   
   For i = 1 To mvarTitulo.Count
      sCod = mvarTitulo(i).Campo
      If sCod = "" Then sCod = Myrs(i - 1).Name
      sKey = ""
      For Each n In mvarChaves
         sKey = sKey & Myrs(CStr(n)) & "."
      Next
      sKey = "k" & Mid(sKey, 1, Len(sKey) - 1)
      sValor = Myrs(sCod)

      If sCod = ItemMask Then
         sValor = Mid(sValor & Space(Len(ClsDsr.UnFormat(mvarMask))), 1, Len(ClsDsr.UnFormat(mvarMask)) + 1)
         sValor = Format(sValor, mvarMask)
      End If
      
      If i = 1 Then
         mvarMe.LstItens.ListItems.Add , sKey, sValor
         nAux = IIf(Len(sValor) > 8, Len(sValor), 8)
      Else
         mvarMe.LstItens.ListItems(sKey).SubItems(i - 1) = sValor
         nAux = IIf(Len(sValor) > 22, Len(sValor), 22)
      End If
      
      If mvarMe.LstItens.ColumnHeaders.Count >= i Then
         If nAux * 120 > mvarMe.LstItens.ColumnHeaders(i).Width Then
            mvarMe.LstItens.ColumnHeaders(i).Width = nAux * 120
         End If
      End If
   Next
   If Not ClsDsr.ExisteItem(mvarCollItens, Myrs(sCod)) Then
      mvarCollItens.Add Mid(sKey, 2), Mid(sKey, 2)
   End If
End Sub
