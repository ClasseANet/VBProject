VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADOTRATAMENTO_PROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmCADOTRATAMENTO_PROD
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar As TL_ShortBar
Private mvarRptRec   As ReportRow
Private TbCad        As Object
Private mvarSys      As Object

Private mvarID       As Integer
Private mvarIDLOJA   As Integer
Public Property Set TLCadastro(ByRef vData As Object)
   Dim Sql As String
   Dim sAux As String
   With vData
      .Titulo = "Funcionários"
      
      Sql = "Select TP.IDPROD, TP.IDTPSERVICO, TP.IDTPTRATAMENTO, TP.IDAREA" & vbNewLine
      'Sql = Sql & ", P.NMPROD, S.DSCSERVICO, T.DSCTRATAMENTO, A.DSCAREA, P.VLVENDA" & vbNewLine
      Sql = Sql & ", P.NMPROD [Produto]" & vbNewLine
      Sql = Sql & ", 'R$ '+Cast(P.VLVENDA as varchar) [Valor]" & vbNewLine
      Sql = Sql & ", [Serviço]=case IsNull(S.IDTPSERVICO,0) When 0  then '[Todos]' Else S.DSCSERVICO End" & vbNewLine
      Sql = Sql & ", [Tratamento]=case IsNull(T.IDTPTRATAMENTO,0) When 0  then '[Todos]' Else T.DSCTRATAMENTO End" & vbNewLine
      Sql = Sql & ", [Área]=case IsNull(A.IDAREA,0) When 0  then '[Todas]' Else A.DSCAREA End" & vbNewLine
      Sql = Sql & " From OTRATAMENTO_PROD TP" & vbNewLine
      Sql = Sql & " Left Join OTPSERVICO S    On S.IDLOJA=TP.IDLOJA And S.IDTPSERVICO=TP.IDTPSERVICO" & vbNewLine
      Sql = Sql & " Left Join OTPTRATAMENTO T On T.IDLOJA=TP.IDLOJA And T.IDTPTRATAMENTO=TP.IDTPTRATAMENTO" & vbNewLine
      Sql = Sql & " Left Join OAREA A On A.IDLOJA=TP.IDLOJA And A.IDAREA=TP.IDAREA" & vbNewLine
      Sql = Sql & " Left Join SPRODUTO P  On P.IDLOJA=TP.IDLOJA And P.IDPROD=TP.IDPROD" & vbNewLine
      Sql = Sql & " Where TP.IDLOJA=" & mvarIDLOJA & vbNewLine
      Sql = Sql & " Order By S.DSCSERVICO, T.DSCTRATAMENTO, A.DSCAREA, P.NMPROD"
      'Sql = Sql & " Order By A.DSCAREA Desc, T.DSCTRATAMENTO Desc, S.DSCSERVICO Desc, P.NMPROD"
      
      'Sql = ""
      .Consulta = Sql
      
      sAux = ""
      sAux = sAux & "|CAMPO=IDPROD        |CHAVE=S|GRUPO=0|VISIBLE=N:"
      sAux = sAux & "|CAMPO=IDTPSERVICO   |CHAVE=S|GRUPO=0|VISIBLE=N:"
      sAux = sAux & "|CAMPO=IDTPTRATAMENTO|CHAVE=S|GRUPO=0|VISIBLE=N:"
      sAux = sAux & "|CAMPO=IDAREA        |CHAVE=S|GRUPO=0|VISIBLE=N:"
   
      .TAGCAMPOS = sAux
   End With
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
'   mvarIDLOJA = RptRec.Record.Item(0).Value
   If Not mvarRptRec Is Nothing Then mvarID = mvarRptRec.Record.Item(0).Value
End Property
Public Property Get RptRec() As ReportRow
   Set RptRec = mvarRptRec
End Property
Public Property Set ShortBar(ByVal vData As TL_ShortBar)
   Set mvarShortBar = vData
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = Val(mvarSys.Propriedades("IDLOJA"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Let ID(ByVal vData As String)
   mvarID = vData
End Property
Public Property Get ID() As String
   ID = mvarID
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Sub Excluir()
   Call Show
End Sub
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADOTRATAMENTO_PROD
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      Call MontaTela
      Call PopulaTela
      Screen.MousePointer = vbDefault
   End If
End Sub
Private Sub mvarMe_CmdProdutoClick()
   Call ShowProduto
End Sub

Private Sub mvarMe_CmdTratamentoClick()
   Call ShowTratamento
End Sub
Private Sub ShowProduto()
   Dim MyProd As New TL_CADSPRODUTO
   With MyProd
      Set .Sys = mvarSys
      .Show
   End With
   Set MyProd = Nothing
   Call MontaComboIDPROD
   'Call MontaTela
   'Call PopulaTela
End Sub

Private Sub ShowTratamento(Optional sDSCTrat As String = "")
   Dim MyTrat As New TL_CADOTPTRATAMENTO
   With MyTrat
      Set .Sys = mvarSys
      .DSCTRAT = sDSCTrat
      .Show
   End With
   Set MyTrat = Nothing
   Call MontaComboIDTPTRATAMENTO
   'Call MontaTela
   'Call PopulaTela
End Sub
Private Sub mvarMe_GrdProdAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim Sql As String
   With mvarMe.GrdProd
      If .ColKey(lCol) = "IDPROD" Then
         If xVal(.CellValue(lRow, "IDPROD")) <> 0 Then
            If xVal(.CellValue(lRow, "IDPROD")) = "9999" Then
            '   Call ShowProduto
            Else
               Sql = "Select VLVENDA"
               Sql = Sql & " From SPRODUTO"
               Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
               Sql = Sql & " And IDPROD=" & SqlNum(.CellValue(lRow, "IDPROD"))
               If mvarSys.Xdb.AbreTabela(Sql) Then
                  .CellValue(lRow, "VLVENDA") = "R$ " & ValBr(mvarSys.Xdb.RsAux("VLVENDA"))
               End If
            End If
         Else
            .CellValue(lRow, "VLVENDA") = ""
         End If
      ElseIf .ColKey(lCol) = "IDTPTRATAMENTO" Then
         'If xVal(.CellValue(lRow, "IDTPTRATAMENTO")) = "9999" Then
         '   Call ShowTratamento(mvarMe.GrdProd.CellText(lRow, "IDPROD"))
         'End If
      End If
   End With
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lRow As Long
   
   If KeyCode = vbKeyDelete Then
      With mvarMe
         If .ActiveControl Is .GrdProd Then
            With .ActiveControl
               lRow = .CurRow
               If lRow <> .RowCount And lRow > 0 Then
                  .RemoveRow (.CurRow)
                  .Tag = SetTag(.Tag, "ISDIRT", 1)
                  .SetFocus
                  If lRow < .RowCount Then
                     .SetCurCell lRow, 1
                  ElseIf lRow > 1 Then
                     .SetCurCell lRow - 1, 1
                  End If
               End If
               
            End With
         End If
      End With
   End If
End Sub

Private Sub mvarMe_Load()
   DoEvents
'   Call MontaTela
'   Call PopulaTela
   Screen.MousePointer = vbHourglass
End Sub
Private Sub PopulaTela()
   Dim Sql  As String
   Dim MyRs As Object
   Dim i As Integer
   
   'Sql = "Select TP.IDLOJA, TP.ID, TP.IDTPSERVICO, TP.IDTPTRATAMENTO, TP.IDAREA, TP.IDPROD"
   'Sql = Sql & ", P.NMPROD [Produto], 'R$'+Replace(Cast(P.VLVENDA as Varchar),'.', ',') [VLVENDA]"
   'Sql = Sql & ", S.DSCSERVICO [Serviço],T.DSCTRATAMENTO [Tratamento], A.DSCAREA [Area]"
   
   Sql = "Select TP.IDPROD, P.VLVENDA, TP.IDTPSERVICO, TP.IDTPTRATAMENTO, TP.IDAREA"
   Sql = Sql & ", A.DSCAREA, T.DSCTRATAMENTO, S.DSCSERVICO, P.NMPROD"
   Sql = Sql & " From OTRATAMENTO_PROD TP"
   Sql = Sql & " Left Join OTPSERVICO S    On S.IDLOJA=TP.IDLOJA And S.IDTPSERVICO=TP.IDTPSERVICO"
   Sql = Sql & " Left Join OTPTRATAMENTO T On T.IDLOJA=TP.IDLOJA And T.IDTPTRATAMENTO=TP.IDTPTRATAMENTO"
   Sql = Sql & " Left Join OAREA A On A.IDLOJA=TP.IDLOJA And A.IDAREA=TP.IDAREA"
   Sql = Sql & " Left Join SPRODUTO P  On P.IDLOJA=TP.IDLOJA And P.IDPROD=TP.IDPROD"
   Sql = Sql & " Where TP.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " Order By S.DSCSERVICO, T.DSCTRATAMENTO, A.DSCAREA, P.NMPROD"
   With mvarMe.GrdProd
      .Clear False
      If mvarSys.Xdb.AbreTabela(Sql, MyRs) Then
         MyRs.MoveFirst
         While Not MyRs.EOF
            .AddRow
             i = MyRs.AbsolutePosition
            .CellValue(i, "IDPROD") = CStr(MyRs("IDPROD") & "")
            .CellValue(i, "VLVENDA") = "R$ " & ValBr(MyRs("VLVENDA") & "")
            .CellValue(i, "IDTPSERVICO") = CStr(MyRs("IDTPSERVICO") & "")
            .CellValue(i, "IDTPTRATAMENTO") = CStr(xVal(MyRs("IDTPTRATAMENTO") & ""))
            .CellValue(i, "IDAREA") = CStr(xVal(MyRs("IDAREA") & ""))
            
            .CellForeColor(i, "VLVENDA") = vbGrayText
            MyRs.MoveNext
         Wend
         '.FillFromRS MyRs
      End If
      For i = 1 To .RowCount
         .CellForeColor(i, "VLVENDA") = vbGrayText
      Next
   End With
   Call IncluiGrdLinha(False)
End Sub
Private Sub LimpaTela()
   Dim n As Object
   On Error Resume Next
   With mvarMe
      For Each n In mvarMe.Controls
         n.Text = ""
      Next
   End With
End Sub
Private Sub MontaTela()
   Dim TbServ  As Object
'   Dim TbTrat  As Object
   Dim TbArea  As Object
   Dim i       As Integer
'   Dim SQL     As String
   Dim sAux    As String
'   Dim nMax    As Integer
      
   Set TbServ = CriarObjeto("BANCO_3R.TB_OTPSERVICO", False)
   Set TbServ.Xdb = mvarSys.Xdb
   Set TbArea = CriarObjeto("BANCO_3R.TB_OAREA", False)
   Set TbArea.Xdb = mvarSys.Xdb
   
   
   With mvarMe.GrdProd
      .Clear True
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
      For i = .Combos.Count To 1 Step -1
         .Combos.Remove i
      Next
      
      With .Combos
         .Add "IDPROD"
         Call MontaComboIDPROD
         
         With .Add("IDTPSERVICO")
'            .AddItem " [Todos]", CStr(0)
            If TbServ.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1", Ch_OrderBy:="DSCSERVICO") Then
               While Not TbServ.Rs.EOF
                  TbServ.Popula
                  sAux = TbServ.DSCSERVICO
                  sAux = sAux & Space(20 - Len(sAux))
                  .AddItem sAux, CStr(TbServ.IDTPSERVICO)

                  TbServ.Rs.MoveNext
               Wend
            End If
            Set TbServ = Nothing
            .AutoAdjustWidth
         End With
         
         .Add "IDTPTRATAMENTO"
         Call MontaComboIDTPTRATAMENTO
         
         With .Add("IDAREA")
            .AddItem " [Todas]", CStr(0)
            If TbArea.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1", Ch_OrderBy:="DSCAREA") Then
               While Not TbArea.Rs.EOF
                  TbArea.Popula
                  sAux = TbArea.DSCAREA
                  sAux = sAux & Space(20 - Len(sAux))
                  .AddItem sAux, CStr(TbArea.IDAREA)

                  TbArea.Rs.MoveNext
               Wend
            End If
            Set TbArea = Nothing
            .AutoAdjustWidth
         End With
      End With
      
      With .AddCol(sKey:="IDPROD", sHeader:="Produto", lWidth:=120, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDPROD"
      End With
      With .AddCol(sKey:="VLVENDA", sHeader:="Valor", lWidth:=70, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "VLVENDA"
      End With
      With .AddCol(sKey:="IDTPSERVICO", sHeader:="Serviço", lWidth:=115, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDTPSERVICO"
      End With
      With .AddCol(sKey:="IDTPTRATAMENTO", sHeader:="Tratamento", lWidth:=115, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDTPTRATAMENTO"
      End With
      With .AddCol(sKey:="IDAREA", sHeader:="Área", lWidth:=115, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDAREA"
      End With
      
      
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
      .ColTag("VLVENDA") = SetTag(.ColTag("VLVENDA"), "EDIT", False)
      '.ColTag("VLPGTO") = SetTag(.ColTag("VLPGTO"), "EDITOPT", igTextEditNumberOnly)
      IncluiGrdLinha
'      .Width = 7400
   End With
'   mvarMe.Width = 7900
   
   Set TbServ = Nothing
   Set TbArea = Nothing
End Sub
Private Sub MontaComboIDPROD()
   Dim TbProd  As Object
   Dim sAux    As String
   Dim nMax    As Integer
   
   Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO", False)
   Set TbProd.Xdb = mvarSys.Xdb
   
   With mvarMe.GrdProd.Combos("IDPROD")
      .Clear
      If TbProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ESERVICO=1 And EVENDA=1", Ch_OrderBy:="NMPROD") Then
         While Not TbProd.Rs.EOF
            TbProd.Popula
            nMax = IIf(nMax >= Len(TbProd.NMPROD), nMax, Len(TbProd.NMPROD))
            TbProd.Rs.MoveNext
         Wend
         TbProd.Rs.MoveFirst
         While Not TbProd.Rs.EOF
            TbProd.Popula
            sAux = TbProd.NMPROD
            sAux = sAux & Space(nMax - Len(sAux))
            .AddItem sAux, CStr(TbProd.IDPROD)
            
            TbProd.Rs.MoveNext
         Wend
      End If
      .AddItem " [Novo...]", "9999"
      Set TbProd = Nothing
      .AutoAdjustWidth
   End With
   Set TbProd = Nothing
End Sub
Private Sub MontaComboIDTPTRATAMENTO()
   Dim TbTrat  As Object
   Dim sAux    As String
   Dim nMax    As Integer
   
   Set TbTrat = CriarObjeto("BANCO_3R.TB_OTPTRATAMENTO", False)
   Set TbTrat.Xdb = mvarSys.Xdb
   
   With mvarMe.GrdProd.Combos("IDTPTRATAMENTO")
      .Clear
      .AddItem " [Todos]", CStr(0)
      If TbTrat.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1", Ch_OrderBy:="DSCTRATAMENTO") Then
         While Not TbTrat.Rs.EOF
            TbTrat.Popula
            sAux = TbTrat.DSCTRATAMENTO
            sAux = sAux & Space(30 - Len(sAux))
            .AddItem sAux, CStr(TbTrat.IDTPTRATAMENTO)

            TbTrat.Rs.MoveNext
         Wend
      End If
      .AddItem " [Novo...]", "9999"
      Set TbTrat = Nothing
      .AutoAdjustWidth
   End With
   Set TbTrat = Nothing
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim Queries As Collection
   Dim TbProdTrat As Object
   Dim i As Integer
   
   Set Queries = New Collection
   Set TbProdTrat = CriarObjeto("BANCO_3R.TB_OTRATAMENTO_PROD", False)
   Set TbProdTrat.Xdb = mvarSys.Xdb
   With mvarMe.GrdProd
      For i = 1 To .RowCount - 1
         If xVal(.CellValue(i, "IDPROD")) <> 0 Then
            TbProdTrat.ID = i
            TbProdTrat.IDLOJA = mvarIDLOJA
            TbProdTrat.IDPROD = .CellValue(i, "IDPROD")
            TbProdTrat.IDTPSERVICO = .CellValue(i, "IDTPSERVICO")
            TbProdTrat.IDTPTRATAMENTO = .CellValue(i, "IDTPTRATAMENTO")
            TbProdTrat.IDAREA = .CellValue(i, "IDAREA")
            Queries.Add TbProdTrat.QrySave
         End If
      Next
      Queries.Add TbProdTrat.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ID>=" & .RowCount)
   End With
   If Queries.Count > 0 Then
      bOk = mvarSys.Xdb.Executa(Queries)
   End If
   Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   
   Salvar = bOk
End Function
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdSalvarClick()
   Screen.MousePointer = vbHourglass
   If Salvar Then
   
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_GrdProdBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   With mvarMe.GrdProd
      If .ColKey(lCol) = "IDPROD" Then
         If vNewValue = "9999" Then
            Call ShowProduto
            eResult = igEditResCancel
         End If
      ElseIf .ColKey(lCol) = "IDTPTRATAMENTO" Then
         If vNewValue = "9999" Then
            Call ShowTratamento(mvarMe.GrdProd.CellText(lRow, "IDPROD"))
            eResult = igEditResCancel
         End If
      End If
   End With
   If lConvErr <> 0 Then
      ExibirAviso "Valor inválido."
      eResult = igEditResProceed
   End If
   
End Sub
Private Sub mvarMe_GrdProdColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdProdLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer

   With mvarMe.GrdProd
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then .Tag = SetTag(.Tag, "ISDIRT", 1)
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub
Private Sub mvarMe_GrdProdMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   Dim Sql As String
   With mvarMe.GrdProd
      If lRow = .RowCount Then
         Call IncluiGrdLinha
         
         .Tag = SetTag(.Tag, "ISDIRT", 1)
         
         If .RowCount > 0 And lRow > 0 Then
            If xVal(.CellValue(lRow, "IDPROD")) = 0 Then
               Call .SetCurCell(lRow, "IDPROD")
            End If
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdProdRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   With mvarMe.GrdProd
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub
Private Sub mvarMe_GrdProdValidate(Cancel As Boolean)
   mvarMe.GrdProd.CommitEdit
End Sub
Private Sub IncluiGrdLinha(Optional pNew As Boolean = True)
   Dim lRow As Long
   Dim i As Integer
   Dim nColDefault As Integer
   Dim nVlAPagar  As Currency
   Dim nVlPgto    As Currency
   Dim Sql As String
     
   nColDefault = 1
   With mvarMe.GrdProd
      '.Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      'If pNew Then
      '   If .Combos("IDPROD").ListCount <= .RowCount - 1 Then
      '      Call ExibirAviso("Não existem mais produto de venda para ser associado.")
      '      Exit Sub
      '   End If
      'End If
      
      lRow = .RowCount
      If lRow > 0 Then
         If .CellValue(.RowCount, 1) = "Clique para incluir nova linha." Then
            .RemoveRow lRow
         End If
      End If
      If lRow > 1 Then .CellForeColor(lRow - 1, "IDPROD") = vbBlack
      
      'If Not (mvarIDVENDA = 0 And lRow = 0) Then
      If lRow > 0 And pNew Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = 17
   
         .CellType(lRow, "IDPROD") = igCellCombo
         If .Combos("IDPROD").ListCount = 1 Then
            .CellValue(lRow, "IDPROD") = .Combos("IDPROD").ItemValue(0)
         End If
         
         .CellType(lRow, "VLVENDA") = igCellText
         .CellForeColor(lRow, "VLVENDA") = vbGrayText
         If xVal(.CellValue(lRow, "IDPROD")) <> 0 Then
            Sql = "Select VLVENDA From SPRODUTO Where IDPROD=" & SqlNum(.CellValue(lRow, "IDPROD"))
            If mvarSys.Xdb.AbreTabela(Sql) Then
               .CellValue(lRow, "VLVENDA") = "R$ " & ValBr(mvarSys.Xdb.RsAux("VLVENDA"))
            End If
         Else
            .CellValue(lRow, "VLVENDA") = ""
         End If
         .CellForeColor(lRow, "VLVENDA") = vbGrayText
         
         .CellType(lRow, "IDTPSERVICO") = igCellCombo
         If .Combos("IDTPSERVICO").ListCount = 1 Then
            .CellValue(lRow, "IDTPSERVICO") = .Combos("IDTPSERVICO").ItemValue(0)
         End If
   
         .CellType(lRow, "IDTPTRATAMENTO") = igCellCombo
         If .Combos("IDTPTRATAMENTO").ListCount = 1 Then
            .CellValue(lRow, "IDTPTRATAMENTO") = .Combos("IDTPTRATAMENTO").ItemValue(0)
         End If
   
         .CellType(lRow, "IDAREA") = igCellCombo
         If .Combos("IDAREA").ListCount = 1 Then
            .CellValue(lRow, "IDAREA") = .Combos("IDAREA").ItemValue(0)
         End If
           
         For i = 1 To .ColCount
            If GetTag(.ColTag(i), "EDIT", True) = False Then
               .CellForeColor(lRow, i) = vbGrayText
            End If
         Next
   
      End If
      .AddRow
     
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
      Next
      For i = 1 To .ColCount
         If .ColVisible(i) Then
            .CellValue(.RowCount, i) = "Clique para incluir nova linha."
            .CellForeColor(.RowCount, i) = vbGrayText
            .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
            Exit For
         End If
      Next
      .RowMode = (.RowCount = 1)
      If lRow <> 0 Then
         'Call mvarMe_GrdProdAfterCommitEdit(lRow, .ColIndex("IDPROD"))
      
         .SetCurCell .RowCount, "IDPROD"
         If pNew Then
            .SetCurCell lRow, nColDefault
         End If
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub

