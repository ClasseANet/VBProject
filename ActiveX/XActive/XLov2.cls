VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XLOV2"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
Private Const CAD = "CAD"
Private Const LOV = "LOV"

Private WithEvents mvarMe  As FrmLov
Attribute mvarMe.VB_VarHelpID = -1
Private mvarId             As Variant
Private mvarIdField        As Variant
Private mvarCaption        As String
Private mvarTituloRel      As String

Private mvarWs             As Workspace
Private mvarDBase          As Object


Private mvarTable          As String
Private mvarCab            As Variant
Private mvarQuery          As Variant
Private mvarFundoTela      As String

Private mvarFrmCad         As Object
Private mvarFormOwner      As Object

Private mvarTipo           As String

Private mvarIdioma         As Long
Private mvarAplic          As Object

Private mvarMultRows       As Boolean
Private mvarMerge          As Boolean
Private mvarMergeCells     As MergeCellsSettings
Private mvarMergeRow       As New Collection
Private mvarMergeCol       As New Collection

Private mvarSistema        As String
Private mvarVersao         As String
Private mvarEmpresa        As String

Private mvarSortCol        As Long
Private mvarSortAscending  As Boolean
Private mvarisTree         As Boolean
Private mvarTreeDefault    As Boolean
Private mvarCAMPO_ID       As String
Private mvarCAMPO_CODIGO   As String
Private mvarCAMPO_NOME     As String
Private mvarCAMPO_PAI      As String
Private mvarExibeCodigo    As Boolean
Private mvarWidthScr       As Double

'* Locais
Private RowSelAntes        As Long
Private RowLoop            As Boolean
Private Rs                 As Object
Private RSLov              As Object
Private ListaFiltrada      As Boolean
Public Property Let TituloRel(ByVal vData As String)
    mvarTituloRel = vData
End Property
Public Property Get TituloRel() As String
    TituloRel = mvarTituloRel
End Property
Public Property Let WidthScr(ByVal vData As Double)
    mvarWidthScr = vData
End Property
Public Property Get WidthScr() As Double
    WidthScr = mvarWidthScr
End Property
Public Property Let ExibeCodigo(ByVal vData As Boolean)
    mvarExibeCodigo = vData
End Property
Public Property Get ExibeCodigo() As Boolean
    ExibeCodigo = mvarExibeCodigo
End Property
Public Property Let CAMPO_ID(ByVal vData As String)
    mvarCAMPO_ID = vData
End Property
Public Property Get CAMPO_ID() As String
    CAMPO_ID = mvarCAMPO_ID
End Property
Public Property Let CAMPO_CODIGO(ByVal vData As String)
    mvarCAMPO_CODIGO = vData
End Property
Public Property Get CAMPO_CODIGO() As String
    CAMPO_CODIGO = mvarCAMPO_CODIGO
End Property
Public Property Let CAMPO_NOME(ByVal vData As String)
    mvarCAMPO_NOME = vData
End Property
Public Property Get CAMPO_NOME() As String
    CAMPO_NOME = mvarCAMPO_NOME
End Property
Public Property Let CAMPO_PAI(ByVal vData As String)
    mvarCAMPO_PAI = vData
End Property
Public Property Get CAMPO_PAI() As String
    CAMPO_PAI = mvarCAMPO_PAI
End Property
Public Property Let TreeDefault(ByVal vData As Boolean)
    mvarTreeDefault = vData
End Property
Public Property Get TreeDefault() As Boolean
    TreeDefault = mvarTreeDefault
End Property
Public Property Let isTree(ByVal vData As Boolean)
    mvarisTree = vData
End Property
Public Property Get isTree() As Boolean
    isTree = mvarisTree
End Property
Public Property Let Sistema(ByVal vData As String)
    mvarSistema = vData
End Property
Public Property Get Sistema() As String
    Sistema = mvarSistema
End Property
Public Property Let Versao(ByVal vData As String)
    mvarVersao = vData
End Property
Public Property Get Versao() As String
    Versao = mvarVersao
End Property
Public Property Let Empresa(ByVal vData As String)
    mvarEmpresa = vData
End Property
Public Property Get Empresa() As String
    Empresa = mvarEmpresa
End Property
Public Property Let SortAscending(ByVal vData As Boolean)
    mvarSortAscending = vData
End Property
Public Property Get SortAscending() As Boolean
    SortAscending = mvarSortAscending
End Property
Public Property Let SortCol(ByVal vData As Long)
    mvarSortCol = vData
End Property
Public Property Get SortCol() As Long
    SortCol = mvarSortCol
End Property
Public Property Set MergeCol(ByVal vData As Object)
    Set mvarMergeCol = vData
End Property
Public Property Get MergeCol() As Collection
    Set MergeCol = mvarMergeCol
End Property
Public Property Set MergeRow(ByVal vData As Object)
    Set mvarMergeRow = vData
End Property
Public Property Get MergeRow() As Collection
    Set MergeRow = mvarMergeRow
End Property
Public Property Let Idioma(ByVal vData As Long)
    mvarIdioma = vData
End Property
Public Property Get Idioma() As Long
    Idioma = mvarIdioma
End Property
Public Sub Show()
   Dim Sql$, Campos$, i%
   Campos = ""
'* Monta Query se a propiedade Query for um inteiro que
'* indica a coluna de ordenação
   If TypeName(mvarQuery) = "Integer" Then
      For i = 1 To (UBound(mvarCab) + 1) / 4
         If i <> 1 Then Campos = Campos + ","
         Campos = Campos + mvarCab((4 * i) - 3)
      Next
      If mvarisTree Then
         If InStr(Campos, mvarCAMPO_PAI) = 0 Then
            Campos = Campos & "," & mvarCAMPO_PAI
         End If
      End If
      Sql = "select " & Campos & " from " & mvarTable
      Sql = Sql + " order by " & CStr(mvarQuery)
      mvarQuery = Sql
   End If
   With mvarMe
         .Sist = mvarSistema
         .Ver = mvarVersao
         .Cia = mvarEmpresa
         
'         .IdField = mvarIdField
'         .DataLov.DatabaseName = mvarDbase.Name
'         Call  MontarMSGrid(.DataLov, .GrdLov, mvarCab, CStr(mvarQuery))
   End With
   mvarMe.Show vbModal
End Sub
Public Property Let MultRows(ByVal vData As Boolean)
    mvarMultRows = vData
End Property
Public Property Get MultRows() As Boolean
    MultRows = mvarMultRows
End Property
Public Property Let Merge(ByVal vData As Boolean)
    mvarMerge = vData
End Property
Public Property Get Merge() As Boolean
    Merge = mvarMerge
End Property
Public Property Let FundoTela(ByVal vData As String)
    mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
    FundoTela = mvarFundoTela
End Property
Public Property Let Tipo(ByVal vData As String)
    mvarTipo = vData
End Property
Public Property Get Tipo() As String
    mvarTipo = IIf(Trim(mvarTipo) = "", "LOV", mvarTipo)
    Tipo = mvarTipo
End Property
Public Property Get FrmCad() As Object
    Set FrmCad = mvarFrmCad
End Property
Public Property Let FrmCad(frm As Object)
    Set mvarFrmCad = frm
End Property
Public Property Get FormOwner() As Object
    Set FormOwner = mvarFormOwner
End Property
Public Property Let FormOwner(frm As Object)
    Set mvarFormOwner = frm
End Property
Public Property Let Dbase(ByVal vData As Object)
   Set mvarDBase = vData
End Property
Public Property Get Dbase() As Object
   Set Dbase = mvarDBase
End Property
Public Property Let Aplic(ByVal vData As Object)
   Set mvarAplic = vData
End Property
Public Property Get Aplic() As Object
   Set Aplic = mvarAplic
End Property
Public Property Set WS(ByVal vData As Workspace)
    Set mvarWs = vData
End Property
Public Property Get WS() As Variant
   Set WS = mvarWs
End Property
Public Property Get IdField() As Variant
    IdField = mvarIdField
End Property
Public Property Let IdField(Vet As Variant)
    mvarIdField = Vet
End Property
Public Property Get Id() As Variant
    Id = mvarId
End Property
Public Property Let Id(Vet As Variant)
    mvarId = Vet
End Property
Public Property Get Cab() As Variant
    Cab = mvarCab
End Property
Public Property Let Cab(Vet As Variant)
    mvarCab = Vet
End Property
Public Property Get Caption() As String
    Caption = mvarCaption
End Property
Public Property Let Caption(Txt$)
    mvarCaption = Txt$
End Property
Public Property Get Table() As String
    Table = mvarTable
End Property
Public Property Let Table(Txt$)
    mvarTable = Txt$
End Property
Public Property Get Query() As Variant
    Query = mvarQuery
End Property
Public Property Let Query(Txt As Variant)
   If IsEmpty(Txt) Then Txt = 2
   mvarQuery = Txt
End Property
Private Sub mvarMe_Activate()
   Dim i    As Long
   Dim bAux As Boolean
   '* Definir o Textos Form
   If bDebug Then mvarMe_Load
      
   Call FormataTela
   bAux = (mvarMe.GrdLov.Rows.Count = mvarMe.GrdLov.HeaderRows.Count)
   If bAux Then
      Call ClsMsg.ExibirInformacao("A Lista de " & mvarCaption & " " & ClsMsg.LoadMsg(mvarIdioma + 67), ClsMsg.LoadMsg(mvarIdioma + 1))
      mvarId = Array()
      Unload mvarMe
      Exit Sub
   End If
   mvarMe.Visible = True

   If mvarisTree Then
      mvarMe.TreLOV.SetFocus
   Else
      Call mvarMe_GrdLovSelChange
   End If

'   Call ClsDsTView.MontarTView(RSLov, mvarMe.TreLOV, Array(mvarCAMPO_ID, mvarCAMPO_NOME, mvarCAMPO_PAI, mvarCAMPO_CODIGO), ExibirCOD:=mvarExibeCodigo, IniExpand:=False)
End Sub
Private Sub mvarMe_CmdFiltrarClick()
   Screen.MousePointer = vbHourglass
   mvarMe.CmdFiltrar.MousePointer = vbHourglass
   Call FiltrarTree
   mvarMe.CmdFiltrar.MousePointer = 99
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdToExcelClick()
   mvarMe.CmdToExcel.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   Call ClsOffice.GridToExcel(mvarMe.GrdLov, mvarMe.Caption)
   Screen.MousePointer = vbDefault
   mvarMe.CmdToExcel.MousePointer = 99
End Sub
Private Sub mvarMe_CmdTreeViewClick()
   Screen.MousePointer = vbHourglass
   mvarMe.CmdTreeView.MousePointer = vbHourglass
   
   If Not mvarisTree Then
      mvarMe_CmdGridClick
   End If
   If Not mvarMe.TreLOV.Visible Then
      mvarMe.GrdLov.Visible = False
      mvarMe.TreLOV.Visible = True
      mvarMe.TreLOV.SetFocus
   End If
   Screen.MousePointer = vbDefault
   mvarMe.CmdTreeView.MousePointer = 99
End Sub
Private Sub mvarMe_CmdGridClick()
   Screen.MousePointer = vbHourglass
   mvarMe.CmdGrid.MousePointer = vbHourglass
   
   If Not mvarMe.GrdLov.Visible Then
      If mvarMe.GrdLov.Rows.Count - 1 <> RSLov.RecordCount Then
         Call ClsMSGrid.MontarMSGrid(RSLov, mvarMe.GrdLov, mvarCab, CStr(mvarQuery))
      End If
      mvarMe.TreLOV.Visible = False
      mvarMe.GrdLov.Visible = True
      mvarMe.GrdLov.SetFocus
   End If
   Screen.MousePointer = vbDefault
   mvarMe.CmdGrid.MousePointer = 99
End Sub

Private Sub mvarMe_FrmKeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyReturn
         If mvarisTree Then
            If Not mvarMe.ActiveControl Is mvarMe.TxtLov Then
               Call mvarMe_CmdOperClick(0)
            End If
         Else
            Call mvarMe_CmdOperClick(0)
         End If
      Case vbKeyEscape
         mvarMe.TxtLov.Text = ""
       Case 199, 231: KeyAscii = 199  '* Ç
       Case 195, 227: KeyAscii = 195  '* Ã
       Case 213, 245: KeyAscii = 213  '* Õ
      Case Else: KeyAscii = ClsDsr.SendTab(mvarMe, KeyAscii)
   End Select
   
   DoEvents
End Sub
Private Sub mvarMe_FrmKeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn:
      Case vbKeyInsert:
      Case vbKeyEscape
         mvarMe.TxtLov = ""
   End Select
   If mvarMe.ActiveControl Is Nothing Then Exit Sub
   If Not mvarMe.ActiveControl Is mvarMe.GrdLov Then Exit Sub
'* SE o Foco estiver no Grid
   Select Case KeyCode
      Case vbKeyDelete:
      Case vbKeyEnd
         mvarMe.GrdLov.TopRow = IIf(mvarMe.GrdLov.Rows.Count - 20 > 0, mvarMe.GrdLov.Rows.Count - 20, mvarMe.GrdLov.Rows.Count - 1)
         Call ClsMSGrid.SelRowMSGrid(mvarMe.GrdLov, mvarMe.GrdLov.Rows.Count - 1)
      Case vbKeyHome
         mvarMe.GrdLov.TopRow = 1
         Call ClsMSGrid.SelRowMSGrid(mvarMe.GrdLov, 1)
      Case vbKeyBack: mvarMe.TxtLov.Text = Mid(mvarMe.TxtLov, 1, Len(mvarMe.TxtLov) - IIf(Len(mvarMe.TxtLov) = 0, 0, 1))
   End Select
End Sub

Private Sub mvarMe_GrdClick()
   Dim iCol As Long, iRow As Long
   Dim DateAux As Date, i As Long
   Dim bSortAscending As Boolean
   Dim Pos As Long
   Dim nKeyCode As Long
   
   With mvarMe.GrdLov
      iCol = .MouseCol
      iRow = .MouseRow
      If ClsCtrl.GetTag(mvarMe.TxtLov, "FILTRO") <> "TRUE" Then
         mvarMe.TxtLov.Text = ""
      End If
      If iRow = 0 Then
         If mvarSortCol = iCol Then
            mvarSortAscending = Not mvarSortAscending
         Else
            mvarSortAscending = True
         End If
         mvarSortCol = iCol
         
         iRow = .Row
         .Row = .FixedRows
         .Col = iCol
         .RowSel = .Rows.Count - .Row
         
         If Trim(.TextMatrix(iRow, iCol)) = "" Then
            .Sort = 2 + mvarSortAscending '*1= flexSortGenericAscending
         Else
            nKeyCode = Asc(Mid(Trim(.TextMatrix(iRow, iCol)), 1, 1))
            On Error Resume Next
                     
            i = 0
            Pos = InStr(.TextMatrix(iRow, iCol), "/")
            If Pos = 2 Or Pos = 3 Then
               i = InStr(Pos + 1, .TextMatrix(iRow, iCol), "/") - Pos
            End If
            DateAux = CDate(.TextMatrix(iRow, iCol))
            If Err = 0 And (i = 3 Or i = 2) Then
               .Sort = 9  '* flexSortCustom
            ElseIf Asc("0") <= nKeyCode And nKeyCode <= Asc("9") Then
               .Sort = 4 + mvarSortAscending '* 3=flexSortNumericAscending
            Else
               .Sort = 8 + mvarSortAscending '*7= flexSortStringAscending
            End If
         End If
         
         Call ClsMSGrid.SetPctOrder(mvarMe.GrdLov, CStr(mvarSortCol), mvarSortAscending)
         
         If .Rows.Count > 1 Then .Row = 1
      
         Call ClsMSGrid.SelRowMSGrid(mvarMe.GrdLov, .Row)
'         mvarMe.GrdLov.Redraw = True
      End If
   End With
End Sub

Private Sub mvarMe_GrdCompare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    On Error Resume Next
    Dim dtmRow1 As Date, dtmRow2 As Date
    
    With mvarMe.GrdLov
        dtmRow1 = CDate(.TextMatrix(Row1, mvarSortCol))
        dtmRow2 = CDate(.TextMatrix(Row2, mvarSortCol))
        If dtmRow1 > dtmRow2 Then
            Cmp = IIf(mvarSortAscending, 1, -1)
        ElseIf dtmRow1 = dtmRow2 Then
            Cmp = 0
        Else
            Cmp = IIf(mvarSortAscending, -1, 1)
        End If
    End With
End Sub
Private Sub mvarMe_GrdDblClick()
   Call mvarMe_CmdOperClick(0)
End Sub
Private Sub mvarMe_GrdLovLeaveCell()
  If Not RowLoop Then
     If mvarMerge And RowSelAntes <> mvarMe.GrdLov.Row And RowSelAntes <> 0 Then
        Call mvarMe_GrdLovSelChange
      End If
   End If
End Sub
Private Sub mvarMe_GrdLovSelChange()
   Dim i&, Lin&, LinAntes&
   RowLoop = True
   Lin = mvarMe.GrdLov.Rows.Count - 1
   If mvarMerge And RowSelAntes <> Lin Then
      If RowSelAntes > 0 And RowSelAntes <> Lin Then
         mvarMe.GrdLov.Row = RowSelAntes
        For i = 0 To mvarMe.GrdLov.Cols - 1
           If i > mvarMe.GrdLov.FixedCols - 1 Then
               mvarMe.GrdLov.Col = i
               mvarMe.GrdLov.CellBackColor = mvarMe.GrdLov.BackColor
               mvarMe.GrdLov.CellForeColor = vbBlack
            End If
        Next
     End If
     If RowSelAntes <> Lin Then
         mvarMe.GrdLov.Row = Lin
         For i = 0 To mvarMe.GrdLov.Cols - 1
            If i > mvarMe.GrdLov.FixedCols - 1 Then
               mvarMe.GrdLov.Col = i
               mvarMe.GrdLov.CellBackColor = &H8000000D '* HighLight
               mvarMe.GrdLov.CellForeColor = vbWhite
            End If
         Next
      End If
      RowSelAntes = mvarMe.GrdLov.Row
      mvarMe.GrdLov.Col = IIf(mvarMe.GrdLov.Col = 1, 0, 1)
   End If
   RowLoop = False
End Sub
Private Sub mvarMe_Load()
   Dim i As Long
   Dim n As Variant
   
   Screen.MousePointer = vbHourglass
   
   Call MontaGrdLov
   Call mvarDBase.AbreTabela(CStr(mvarQuery), RSLov)
   With mvarMe
      
      Call FillRCFromRS(RSLov, .GrdLov)
      'Call ClsMSGrid.MontarMSGrid(RSLov, .GrdLov, mvarCab, CStr(mvarQuery))
      
      If IsEmpty(mvarIdField) Then
         ReDim mvarIdField(RSLov.FIELDS.Count - 1)
         For i = 0 To (RSLov.FIELDS.Count - 1)
            If RSLov.FIELDS(i).Name = "" Then '* Nome do Campo
               mvarIdField(i) = i
            Else
               mvarIdField(i) = RSLov.FIELDS(i).Name
            End If
         Next
      End If
      
      If mvarisTree Then
         Call ClsDsTView.MontarTView(RSLov, mvarMe.TreLOV, Array(mvarCAMPO_ID, mvarCAMPO_NOME, mvarCAMPO_PAI, mvarCAMPO_CODIGO), ExibirCOD:=mvarExibeCodigo, IniExpand:=False)
      End If
      
      If Not .GrdLov.FocusedColumn Is Nothing Then
         mvarSortCol = .GrdLov.FocusedColumn
         'mvarSortCol = .GrdLov.Columns.Column.s .Index
      End If
      
      '* Determina Propriedades Relacionadas ao "MERGE" do Grid
      If mvarMerge Then
         With .GrdLov
            .MergeCells = mvarMergeCells
            For i = 0 To .Cols - 1
               .MergeCol(i) = True
            Next
            For Each n In mvarMergeCol
               .MergeCol(n.Key) = n.Item
            Next
            For i = 0 To .Rows.Count - 1
               .MergeRow(i) = True
            Next
            For Each n In mvarMergeRow
               .MergeRow(n.Key) = n.Item
            Next
         End With
      End If
   End With
   
   Call ConfigLov
   Call ClsCtrl.ConfigForm(mvarMe, , mvarFundoTela, PintarTitulo:=True)
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdOperClick(Index As Integer)
   Dim lOper   As Long
   Dim Min     As Long
   Dim Max     As Long
   Dim i       As Long
   Dim j       As Long
   Dim k       As Long
   Dim Col     As Long
   Dim MyId    As Variant
   Dim iMouse  As Integer
         
   iMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   mvarMe.CmdLovOper(Index).MousePointer = vbHourglass
         
   Select Case Index
   '* LOV
      Case 0
         If mvarMultRows Then
            Min = mvarMe.GrdLov.Row
            Max = mvarMe.GrdLov.RowSel
            If Min > Max Then
               Min = mvarMe.GrdLov.RowSel
               Max = mvarMe.GrdLov.Row
            End If
            ReDim MyId(UBound(mvarIdField))
            ReDim mvarId((Max - Min) + 1)
            
            k = 0
            For i = Min To Max
               For j = 0 To UBound(MyId)
                  Col = DefCol(j)
                  MyId(j) = mvarMe.GrdLov.TextMatrix(i, Col)
               Next
               mvarId((i - Min)) = MyId
               ReDim MyId(UBound(mvarIdField))
            Next
            
'            For i = 0 To UBound(mvarIdField)
'               Col = DefCol(i)
'               For J = Min To Max
'                  MyId(k) = mvarMe.GrdLov.TextMatrix(J, Col)
'                  k = k + 1
'                  If Min = Max Then 'Acrescentar um item duplicado no vetor setor for apenas um elemento selecionado
'                     MyId(k) = mvarMe.GrdLov.TextMatrix(J, Col)
'                  End If
'               Next
'               mvarId(i) = MyId(k)
'            Next
         Else
            If isTree And mvarMe.TreLOV.Visible And mvarMe.TreLOV.Enabled Then
               If mvarMe.TreLOV.SelectedItem Is mvarMe.TreLOV.Nodes(1) Then
                  ClsMsg.ExibirAviso ("Item inválido")
                  Exit Sub
               End If
               For i = 0 To UBound(mvarIdField)
                  ReDim Preserve mvarId(i)
                  Select Case mvarIdField(i)
                     Case mvarCAMPO_ID
                        mvarId(i) = Mid(mvarMe.TreLOV.SelectedItem.Key, 2)
                        
                     Case mvarCAMPO_CODIGO
                        If mvarExibeCodigo Then
                           mvarId(i) = ClsDsr.UnFormat(Trim(Mid(mvarMe.TreLOV.SelectedItem.Text, 1, InStr(mvarMe.TreLOV.SelectedItem.Text, "-") - 1)))
                        End If
                        
                     Case mvarCAMPO_NOME
                        If mvarExibeCodigo Then
                           mvarId(i) = Trim(Mid(mvarMe.TreLOV.SelectedItem.Text, InStr(mvarMe.TreLOV.SelectedItem.Text, "-") + 1))
                        Else
                           mvarId(i) = Trim(mvarMe.TreLOV.SelectedItem.Text)
                        End If
                        
                     Case mvarCAMPO_PAI
                        On Error Resume Next
                        mvarId(i) = Mid(mvarMe.TreLOV.SelectedItem.Parent.Key, 2)
                        
                     Case Else
                        On Error Resume Next
                        For j = 0 To mvarMe.GrdLov.Cols - 1
                           If mvarCab((j * 4) + 1) = mvarIdField(i) Then
                              mvarId(i) = mvarMe.GrdLov.TextMatrix(mvarMe.GrdLov.Row, j)
                              Exit For
                           End If
                        Next
                  End Select
               Next
            Else
               For i = 0 To UBound(mvarIdField)
                  ReDim Preserve mvarId(i)
                  If DefCol(i) <= mvarMe.GrdLov.Cols - 1 Then
                     mvarId(i) = mvarMe.GrdLov.TextMatrix(mvarMe.GrdLov.Row, DefCol(i))
                  Else
                     On Error Resume Next
                     For j = 0 To mvarMe.GrdLov.Cols - 1
                        If mvarCab((j * 4) + 1) = mvarIdField(i) Then
                           mvarId(i) = mvarMe.GrdLov.TextMatrix(mvarMe.GrdLov.Row, j)
                           Exit For
                        End If
                     Next
                  End If
               Next
            End If
         End If
      Case 1: mvarId = Array()
   End Select
   
   mvarMe.CmdLovOper(Index).MousePointer = 99
   Screen.MousePointer = iMouse
   Unload mvarMe
End Sub
Private Function DefCol%(ByVal Ind%)
   Dim Sql  As String
   Dim nAux As Long
   Dim nCol As Long
   Dim i    As Long
   
   
   Sql = mvarQuery   'mvarMe.DataLov.RecordSource
   
   nAux = 0
   nCol = 0
   nAux = InStr(UCase(Sql), UCase(" AS " & mvarIdField(Ind%)))
   nCol = IIf(nCol = 0, nAux, nCol)
   nCol = IIf(nAux < nCol And nAux <> 0, nAux, nCol)
   
   nAux = InStr(UCase(Sql), UCase(" AS [" & mvarIdField(Ind%) & "]"))
   nCol = IIf(nCol = 0, nAux, nCol)
   nCol = IIf(nAux < nCol And nAux <> 0, nAux, nCol)
   
   nAux = InStr(Sql, "." & mvarIdField(Ind%) & ",")
   nCol = IIf(nCol = 0, nAux, nCol)
   nCol = IIf(nAux < nCol And nAux <> 0, nAux, nCol)
   
   nAux = InStr(Sql, "." & mvarIdField(Ind%) & " ")
   nCol = IIf(nCol = 0, nAux, nCol)
   nCol = IIf(nAux < nCol And nAux <> 0, nAux, nCol)
   
   nAux = InStr(Sql, "," & mvarIdField(Ind%) & ",")
   nCol = IIf(nCol = 0, nAux, nCol)
   nCol = IIf(nAux < nCol And nAux <> 0, nAux, nCol)
   
   nAux = InStr(Sql, "," & mvarIdField(Ind%) & " ")
   nCol = IIf(nCol = 0, nAux, nCol)
   nCol = IIf(nAux < nCol And nAux <> 0, nAux, nCol)
   
   nAux = InStr(Sql, " " & mvarIdField(Ind%) & " ")
   nCol = IIf(nCol = 0, nAux, nCol)
   nCol = IIf(nAux < nCol And nAux <> 0, nAux, nCol)
   
   nAux = InStr(Sql, " " & mvarIdField(Ind%) & ",")
   nCol = IIf(nCol = 0, nAux, nCol)
   nCol = IIf(nAux < nCol And nAux <> 0, nAux, nCol)
   
         
   If nCol <> 0 Then Sql = Mid(Sql, 1, nCol)
   
   '*************
   '* Eliminar Conteúdo entre parenteses
   If InStr(Sql, "(") <> 0 Then
      nAux = InStr(Sql, "(")
      i = 0
      Do While (InStr(Sql, "(") <> 0)
         nAux = InStr(Sql, "(")
         If nAux + i > 0 Then
            If Len(Sql) <> Len(Mid(Sql, 1, nAux + i) & Mid(Sql, nAux + 2 + i)) Then
               Sql = Mid(Sql, 1, nAux + i) & Mid(Sql, nAux + 2 + i)
               If Mid(Sql, nAux + i + 1, 1) = "(" Then
                  i = i + 1
               End If
               
               If InStr(Sql, "()") <> 0 Then
                  i = i - 1
                  nAux = InStr(Sql, "()")
                  Sql = Mid(Sql, 1, nAux - 1) & Mid(Sql, nAux + 2)
               End If
            Else
               Exit Do
            End If
         End If
      Loop
   End If
   
   DefCol% = 0
   While InStr(Sql, ",") <> 0
      DefCol% = DefCol% + 1
      Sql = Mid(Sql, InStr(Sql, ",") + 1)
   Wend
End Function
Private Sub FormataTela()
   mvarMe.Caption = mvarCaption
   mvarMe.LblLov(3) = ClsMsg.LoadMsg(mvarIdioma + 5014) 'Localizar
   mvarMe.CmdLovOper(1).Caption = ClsMsg.LoadMsg(mvarIdioma + 5015)  'Cancela
   mvarMe.CmdLovOper(0).Caption = ClsMsg.LoadMsg(mvarIdioma + 5018) 'OK
      
   '* Escrerve Tela.
   mvarMe.LblLov(3).ToolTipText = "Localize rapidamente digitando sua consulta"
   mvarMe.TxtLov.ToolTipText = "Use [Esc] para limpar o campo"
   mvarMe.GrdLov.ToolTipText = "Clique no título da lista para ordená-la"
      
      
   On Error Resume Next
   'DSACTIVE.Idioma  = mvarIdioma
End Sub
Private Sub ConfigLov()
'================================================================
'= Última Alteração : 10/07/98                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Configurar Form de Lista de valores.            **
'**                                                            **
'** Recebe: F - Form a ser configurado                         **
'**                                                            **
'** Retorna:                                                   **
'**                                                            **
'****************************************************************
   Dim t&, L&
   On Error GoTo Fim
   With mvarMe
      .CmdLovOper(0).Visible = True
      .CmdLovOper(1).Visible = True
      
      .CmdTreeView.Visible = mvarisTree
      .CmdGrid.Visible = mvarisTree
      .GrdLov.Visible = Not mvarisTree Or Not mvarTreeDefault
      .TreLOV.Visible = mvarisTree And mvarTreeDefault
      
      .CmdFiltrar.Visible = True
      .CmdFiltrar.Left = .TxtLov.Left + .TxtLov.Width
      
      '* Definir Tamanho mínimo do Objeto
      If .GrdLov.Width < 4290 Then
         .GrdLov.Width = 4290
      End If
      '* Definir tamanho do Form
      .Width = .GrdLov.Width + 200  '160
      '* Centralir grid no Form
      Call ClsCtrl.CentrarObj(mvarMe, .GrdLov, "H")
      Call ClsMSGrid.SelRowMSGrid(.GrdLov, 1)
      
      If mvarisTree Then
         .TreLOV.Indentation = 350
         '* Definir tamanho do Form
         If .Width < .TreLOV.Width + .TreLOV.Left + 120 Then
            .Width = .TreLOV.Width + .TreLOV.Left + 120
         End If
         If mvarWidthScr > 0 And .Width < mvarWidthScr Then
            .Width = mvarWidthScr
         End If
         If .TreLOV.Width < (.Width - 390) Then
            .TreLOV.Width = (.Width - 390)
         End If
      End If
      
      '* Definir Características do "CmdImprimir" e "CmdFiltrar"
      .CmdFiltrar.ToolTipText = "Filtrar" & " "
      .CmdImprimir.ToolTipText = ClsMsg.LoadMsg(61) & " "
      .CmdImprimir.ToolTipText = .CmdImprimir.ToolTipText & Trim(mvarCaption)
      .CmdFiltrar.ToolTipText = .CmdFiltrar.ToolTipText & Trim(mvarCaption)
   End With
   Exit Sub
Fim:
   ClsMsg.ShowError
End Sub

Private Sub mvarMe_QueryUnload(Cancel As Variant, UnloadMode As Variant)
   If UnloadMode = 0 Then mvarId = Array()
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
      Call ClsCtrl.PintarFundo(mvarMe, bPintarTitulo:=True)
      .TreLOV.Width = .Width - .TreLOV.Left - 120
      .GrdLov.Width = .Width - .TreLOV.Left - 120
   End With
End Sub
Private Sub mvarMe_TreLOVDblClick()
   If mvarisTree Then
      If mvarMe.TreLOV.SelectedItem Is mvarMe.TreLOV.Nodes(1) Then
         Exit Sub
      End If
   End If
   Call mvarMe_CmdOperClick(0)
End Sub

Private Sub mvarMe_TreLOVExpand(ByVal Node As MSComctlLib.Node)
   If mvarMe.TreLOV.FullRowSelect Then
      Set Node = Node
   End If
End Sub

Private Sub mvarMe_TxtLovChange()
   Dim iRow As Long
   Screen.MousePointer = vbHourglass
   If ClsCtrl.GetTag(mvarMe.TxtLov, "FILTRO") = "TRUE" And ((mvarMe.TxtLov.Text = "") Or (mvarMe.TxtLov.Text <> "" And mvarMe.GrdLov.Visible)) Then
      Call mvarDBase.AbreTabela(mvarQuery, RSLov)
      Call ClsMSGrid.MontarMSGrid(RSLov, mvarMe.GrdLov, mvarCab, CStr(mvarQuery))
      If mvarisTree Then
         Call ClsDsTView.MontarTView(RSLov, mvarMe.TreLOV, Array(mvarCAMPO_ID, mvarCAMPO_NOME, mvarCAMPO_PAI, mvarCAMPO_CODIGO), ExibirCOD:=mvarExibeCodigo, IniExpand:=False)
      End If
      Call ClsCtrl.SetTag(mvarMe.TxtLov, "FILTRO", "FALSE")
   End If
   
   If mvarMe.TxtLov.Text <> "" Then
      If mvarMe.GrdLov.Visible Then
         iRow = ClsMSGrid.PesquisarMSGrid(mvarMe.GrdLov, mvarMe.TxtLov.Text, mvarSortCol)
         mvarMe.GrdLov.TopRow = iRow
         Call ClsMSGrid.SelRowMSGrid(mvarMe.GrdLov, iRow)
      Else
'         iRow = ClsMSGrid.PesquisarMSGrid(mvarMe.GrdLov, mvarMe.TxtLov.Text, mvarSortCol)
'         mvarMe.GrdLov.TopRow = iRow
'         Call ClsMSGrid.SelRowMSGrid(mvarMe.GrdLov, iRow)
      End If
      mvarMe.TxtLov.SetFocus
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_TxtLovKeyPress(KeyAscii As Integer)
   Dim Sql  As String
   Dim sAux As String
   If KeyAscii = vbKeyReturn Then
      Call FiltrarTree
   End If
End Sub
Public Sub FiltrarTree()
   Dim Sql        As String
   Dim sAux       As String
   Dim nAux       As Long
   Dim nVir       As Long
   Dim nWhere     As Long
   Dim nFim       As Long
   Dim sAlias     As String
   Dim CampoAux   As String
   
   Screen.MousePointer = vbHourglass
   On Error Resume Next
   
   CampoAux = mvarCAMPO_NOME
   
   If mvarMe.GrdLov.Visible Then CampoAux = Cab((SortCol * 4) + 1)
   If CampoAux = "" Then CampoAux = Cab((SortCol * 4) + 1)
   
   nAux = InStr(UCase(mvarQuery), "FROM")
   If nAux > 0 Then
      nAux = InStr(nAux + 5, UCase(mvarQuery), UCase(mvarTable))
      If nAux > 0 Then
         nAux = nAux + Len(mvarTable)
         
         nVir = InStr(nAux + 1, UCase(mvarQuery), ",")
         nWhere = InStr(nAux + 1, UCase(mvarQuery), "WHERE")
         If nWhere = 0 Then
            Dim Palavras As Collection
            Set Palavras = ClsDsr.GetWords(Mid(mvarQuery, nAux))
            If ClsDsr.InArray(UCase(Palavras(1)), Array("ORDER", "GROUP", "HAVING")) Then
               nWhere = InStr(nAux, UCase(mvarQuery), UCase(Palavras(1)))
            Else
               nWhere = nAux + Len(Palavras(1)) + 1
            End If
         End If
         If nVir = 0 Then
            nFim = nWhere
         Else
            If nVir < nWhere Then
               nFim = nVir
            Else
               nFim = nWhere
            End If
         End If
         If nFim > nAux Then
            sAlias = Trim(Mid(mvarQuery, nAux, nFim - nAux))
         End If
      End If
   End If
   If Trim(sAlias) <> "" Then
      CampoAux = sAlias & "." & CampoAux
   End If

   If Not Trim(mvarMe.TxtLov.Text) = "" Then
      Call ClsCtrl.SetTag(mvarMe.TxtLov, "FILTRO", "FALSE")
      Sql = UCase(CStr(mvarQuery))
      If InStr(Sql, " WHERE ") <> 0 Then
         sAux = Mid(Sql, 1, InStr(Sql, " WHERE ") + 6)
         sAux = sAux & CampoAux & " Like '%" & mvarMe.TxtLov.Text & "%'"
         sAux = sAux & " And " & Mid(Sql, InStr(Sql, " WHERE ") + 7)
         
         Sql = sAux
      Else
         If InStr(Sql, " GROUP BY ") <> 0 Then
            sAux = Mid(Sql, 1, InStr(Sql, " GROUP BY ") - 1)
            sAux = sAux & " WHERE "
            sAux = sAux & CampoAux & "Like '%" & mvarMe.TxtLov.Text & "%'"
            sAux = sAux & Mid(Sql, InStr(Sql, " GROUP BY "))
            
            Sql = sAux
         Else
            If InStr(Sql, " ORDER BY ") <> 0 Then
               sAux = Mid(Sql, 1, InStr(Sql, " ORDER BY ") - 1)
               sAux = sAux & " WHERE "
               sAux = sAux & CampoAux & " Like '%" & mvarMe.TxtLov.Text & "%'"
               sAux = sAux & Mid(Sql, InStr(Sql, " ORDER BY "))
               
               Sql = sAux
            End If
         End If
      End If
      If Sql <> UCase(CStr(mvarQuery)) Then
         Call ClsCtrl.SetTag(mvarMe.TxtLov, "FILTRO", "TRUE")
      End If
      On Error Resume Next
      If Not mvarDBase.AbreTabela(Sql, RSLov, False) Then
         Call mvarDBase.AbreTabela(mvarQuery, RSLov)
      End If
      If mvarMe.GrdLov.Visible Then
         Call ClsMSGrid.MontarMSGrid(RSLov, mvarMe.GrdLov, mvarCab, CStr(mvarQuery))
      End If
      If mvarisTree Then
         Call ClsDsTView.MontarTView(RSLov, mvarMe.TreLOV, Array(mvarCAMPO_ID, mvarCAMPO_NOME, mvarCAMPO_PAI, mvarCAMPO_CODIGO), ExibirCOD:=mvarExibeCodigo, IniExpand:=False)
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_Unload(Cancel As Integer)
   DoEvents
'   Set mvarMe = Nothing
'   Set mvarFrmCad = Nothing
End Sub
Private Sub Class_Initialize()
   mvarQuery = 2
   mvarCab = Array("")
   mvarId = Array("")
   mvarIdField = Array("")
   mvarIdioma = 5000 'Português
   mvarMultRows = False
   mvarMerge = False
   mvarMergeCells = flexMergeRestrictRows
   mvarisTree = False
   mvarTreeDefault = True
   mvarCAMPO_PAI = "IDPAI"
   Set mvarMe = New FrmLov
   
   Set mvarFrmCad = Nothing
   
   Set mvarWs = DBEngine.Workspaces(0)
'   Set mvarDbase = DBEngine.Workspaces(0).Databases(0)
   
   '* Locais
   RowLoop = False
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Set mvarMe = Nothing
   Set mvarMergeRow = Nothing
   Set mvarMergeCol = Nothing
   
   'Set mvarFrmCad = Nothing
End Sub
Private Sub MontaGrdLov()
   With mvarMe.GrdLov
      .AutoColumnSizing = True
      .ShowFooter = False
      .ShowFooterRows = False
      .ShowGroupBox = True
      .ShowHeader = True
      .ShowHeaderRows = False
      .ShowItemsInGroups = False
      .ShowRowFocus = True
      .PaintManager.ColumnStyle = xtpColumnOffice2003
   End With
End Sub

