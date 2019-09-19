VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_VendaPacote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmVendaPacote
Attribute mvarMe.VB_VarHelpID = -1

Dim bLoad As Boolean

Private mvarSys         As Object
Private mvarIDLOJA      As Integer
Private mvarIDVENDA     As Long
Private mvarIDCLIENTE   As Long
Private mvarDTVENDA     As String
Private mvarPacotes     As Collection
Public Property Set Pacotes(ByVal vData As Collection)
   Set mvarPacotes = vData
End Property
Public Property Get Pacotes() As Collection
   Set Pacotes = mvarPacotes
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDVENDA(ByVal vData As Integer)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Integer
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDCLIENTE(ByVal vData As Integer)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Integer
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let DTVENDA(ByVal vData As String)
   mvarDTVENDA = vData
End Property
Public Property Get DTVENDA() As String
   DTVENDA = mvarDTVENDA
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmVendaPacote
   Set mvarPacotes = New Collection
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
   Set mvarPacotes = Nothing
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdOkClick()
   Call Salvar
   Unload mvarMe
End Sub
Private Sub Salvar()
   Dim i As Integer
   Dim MyPac As Object
   
   With mvarMe.GrdPacote
      Set mvarPacotes = New Collection
      For i = 1 To .RowCount
         If xVal(.CellValue(i, "IDPACOTE")) > 0 Then
            Set MyPac = CriarObjeto("BANCO_3R.TB_CPACOTE")
            Set MyPac.xDb = mvarSys.xDb
            MyPac.IDLOJA = .CellValue(i, "IDLOJA")
            MyPac.IDVENDA = .CellValue(i, "IDVENDA")
            MyPac.IDCLIENTE = .CellValue(i, "IDCLIENTE")
            MyPac.IDPACOTE = .CellValue(i, "IDPACOTE")
            MyPac.DSCPACOTE = Trim(.CellText(i, "IDPROMO"))
            MyPac.DTEMISSAO = .CellValue(i, "DTEMISSAO")
            MyPac.IDPROMO = .CellValue(i, "IDPROMO")
            MyPac.VlTotal = .CellValue(i, "VLTOTAL")
            MyPac.VLDESC = .CellValue(i, "VLDESC")
            MyPac.Valor = .CellValue(i, "VALOR")
            
            mvarPacotes.Add MyPac, CStr(MyPac.IDPACOTE)
            Set MyPac = Nothing
         End If
      Next
   End With
End Sub
Private Sub mvarMe_GrdPacoteAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim MyPromo As Object 'TB_CPROMOCAO
   
   With mvarMe.GrdPacote
      If .ColKey(lCol) = "IDPROMO" Then
         Set MyPromo = CriarObjeto("BANCO_3R.TB_CPROMOCAO")
         Set MyPromo.xDb = mvarSys.xDb
         If MyPromo.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROMO:=.CellValue(lRow, "IDPROMO")) Then
            .CellValue(lRow, "VLTOTAL") = ValBr(MyPromo.VlTotal)
            .CellValue(lRow, "VLDESC") = ValBr(MyPromo.VLDESC)
            .CellValue(lRow, "VALOR") = ValBr(MyPromo.Valor)
            Call CalculaPromo
         End If
         Set MyPromo = Nothing
      End If
   End With
End Sub
Private Sub mvarMe_GrdPacoteBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   If lConvErr <> 0 Then
      ExibirAviso "Valor inválido."
      eResult = igEditResProceed
   End If
End Sub
Private Sub mvarMe_GrdPacoteColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdPacoteLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer
   
   With mvarMe.GrdPacote
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

Private Sub mvarMe_GrdPacoteMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   With mvarMe.GrdPacote
      If lRow = .RowCount Then
         'If .RowCount > .Combos("IDPROMO").ListCount Then
         '   Call ExibirAviso("Não existem mais produtos de venda. (" + CStr(.Combos("IDPROMO").ListCount) + ")")
         '   Exit Sub
         'End If
         
         If lRow > 1 Then
            If xVal(.CellValue(lRow - 1, "IDPROMO")) = 0 Then Exit Sub
         End If
         Call IncluiGrdLinha
         
         If Val(.CellValue(lRow, "IDPROMO")) = 0 Then
            Call .SetCurCell(lRow, "IDPROMO")
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdPacoteRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   With mvarMe.GrdPacote
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub
Private Sub mvarMe_GrdPacoteValidate(Cancel As Boolean)
   mvarMe.GrdPacote.CommitEdit
End Sub
Private Sub MontaGrdPacote()
   Dim TbPromo As Object 'TB_SPRODUTO
   Dim Sql     As String
   Dim sAux    As String
   Dim i       As Integer
   
   Call LimpaTela

   Set TbPromo = CriarObjeto("BANCO_3R.TB_CPROMOCAO")
   Set TbPromo.xDb = mvarSys.xDb
   Call LimpaTela
   
   With mvarMe.GrdPacote
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
      With .Combos
         For i = .Count To 1 Step -1
            .Remove i
         Next
         With .Add("IDPROMO")
            Sql = "ATIVO=1"
            Sql = Sql & " And DTINIV<=" & SqlDate(mvarSys.xDb.SysDate())
            Sql = Sql & " And DTFIMV>=" & SqlDate(mvarSys.xDb.SysDate())
            If TbPromo.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql, Ch_ORDERBY:="IDPROMO") Then
               While Not TbPromo.Rs.EOF
                  TbPromo.Popula
                  sAux = TbPromo.DSCPROMO
                  If Len(sAux) < 20 Then
                     sAux = sAux & Space(20 - Len(sAux))
                  End If
                  .AddItem sAux, TbPromo.IDPROMO
                  
                  TbPromo.Rs.MoveNext
               Wend
            End If
            Set TbPromo = Nothing
            .AutoAdjustWidth
         End With
      End With
      
      .Clear True
      With .AddCol(sKey:="IDLOJA", sHeader:="IDLOJA", lWidth:=25, bVisible:=False)
         .eType = igCellText
         .sCtrlKey = "IDLOJA"
      End With
      With .AddCol(sKey:="IDVENDA", sHeader:="IDVENDA", lWidth:=25, bVisible:=False)
         .eType = igCellText
         .sCtrlKey = "IDVENDA"
      End With
      With .AddCol(sKey:="IDCLIENTE", sHeader:="IDCLIENTE", lWidth:=25, bVisible:=False)
         .eType = igCellText
         .sCtrlKey = "IDCLIENTE"
      End With
      
      With .AddCol(sKey:="IDPACOTE", sHeader:="#", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "IDPACOTE"
      End With
      With .AddCol(sKey:="DTEMISSAO", sHeader:="Emissão", lWidth:=65, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "DTEMISSAO"
      End With
      With .AddCol(sKey:="IDPROMO", sHeader:="Promoção", lWidth:=145, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDPROMO"
      End With
      
      With .AddCol(sKey:="VLTOTAL", sHeader:="Total", lWidth:=50, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "VLTOTAL"
      End With
      With .AddCol(sKey:="VLDESC", sHeader:="Desc.", lWidth:=50, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         '.sFmtString = "R$ #,###.00"
         .sCtrlKey = "VLDESC"
      End With
      With .AddCol(sKey:="VALOR", sHeader:="Valor", lWidth:=50, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "VALOR"
      End With
            
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
      
      .ColTag("IDPACOTE") = SetTag(.ColTag("IDPACOTE"), "EDIT", False)
      .ColTag("DTEMISSAO") = SetTag(.ColTag("DTEMISSAO"), "EDIT", False)
      .ColTag("VLTOTAL") = SetTag(.ColTag("VLTOTAL"), "EDIT", False)
      .ColTag("VLDESC") = SetTag(.ColTag("VLDESC"), "EDIT", False)
      .ColTag("VALOR") = SetTag(.ColTag("VALOR"), "EDIT", False)
            
      IncluiGrdLinha

   End With
End Sub
Private Sub PopulaGrdPacote()
   Dim sTag  As String
   Dim i     As Integer
   Dim vPac  As Variant

   'DoEvents
   With mvarMe.GrdPacote
      .Clear
      If bLoad Then Call IncluiGrdLinha
      'If mvarIDVENDA = 0 Then
         
         For Each vPac In mvarPacotes
            Call IncluiGrdLinha
            
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
            
            Call SetTag(sTag, "IDLOJA", vPac.IDLOJA)
            Call SetTag(sTag, "IDVENDA", vPac.IDVENDA)
            Call SetTag(sTag, "IDCLIENTE", vPac.IDCLIENTE)
            Call SetTag(sTag, "IDPROMO", vPac.IDPROMO)
            Call SetTag(sTag, "IDPACOTE", vPac.IDPACOTE)
            Call SetTag(sTag, "DTEMISSAO", vPac.DTEMISSAO)
            Call SetTag(sTag, "VLTOTAL", ValBr(vPac.VlTotal))
            Call SetTag(sTag, "VLDESC", ValBr(vPac.VLDESC))
            Call SetTag(sTag, "VALOR", ValBr(vPac.Valor))
            
            .RowTag(i) = sTag
            
            .CellValue(i, "IDLOJA") = vPac.IDLOJA
            .CellValue(i, "IDVENDA") = vPac.IDVENDA
            .CellValue(i, "IDCLIENTE") = vPac.IDCLIENTE
            .CellValue(i, "IDPROMO") = vPac.IDPROMO
            .CellValue(i, "IDPACOTE") = vPac.IDPACOTE
            .CellValue(i, "DTEMISSAO") = vPac.DTEMISSAO
            .CellValue(i, "VLTOTAL") = ValBr(vPac.VlTotal)
            .CellValue(i, "VLDESC") = ValBr(vPac.VLDESC)
            .CellValue(i, "VALOR") = ValBr(vPac.Valor)
         Next
      'Else
      '   Call Popula
      'End If
      If .RowCount > 0 Then
         .RowVisible(1) = True
      End If
      '.CellSelected(1, "FOTOTIPO") = True
      .Tag = SetTag(.Tag, "ISDIRT", 0)
   End With
End Sub
Public Sub Popula()
   Dim TbPacote As Object 'TB_CPACOTE
   Dim sTag  As String
   Dim i     As Integer
'   Dim Sql   As String
'   Dim vPac  As Variant
   Dim MyPac As Object
   
   Set mvarPacotes = New Collection
   Set TbPacote = CriarObjeto("BANCO_3R.TB_CPACOTE")
   Set TbPacote.xDb = mvarSys.xDb
   If TbPacote.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_ORDERBY:="IDPACOTE") Then
      If bLoad Then Call IncluiGrdLinha
      While Not TbPacote.Rs.EOF
         TbPacote.Popula
         
         Call SetTag(sTag, "IDLOJA", TbPacote.IDLOJA)
         Call SetTag(sTag, "IDVENDA", TbPacote.IDVENDA)
         Call SetTag(sTag, "IDCLIENTE", TbPacote.IDCLIENTE)
         Call SetTag(sTag, "IDPROMO", TbPacote.IDPROMO)
         Call SetTag(sTag, "IDPACOTE", TbPacote.IDPACOTE)
         Call SetTag(sTag, "DSCPACOTE", TbPacote.DSCPACOTE)
         Call SetTag(sTag, "DTEMISSAO", TbPacote.DTEMISSAO)
         Call SetTag(sTag, "VLTOTAL", ValBr(TbPacote.VlTotal))
         Call SetTag(sTag, "VLDESC", ValBr(TbPacote.VLDESC))
         Call SetTag(sTag, "VALOR", ValBr(TbPacote.Valor))
         
         If bLoad Then
            With mvarMe.GrdPacote
               Call IncluiGrdLinha
               i = .RowCount - 1
               If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
      
               .RowTag(i) = sTag
               
               .CellValue(i, "IDLOJA") = TbPacote.IDLOJA
               .CellValue(i, "IDVENDA") = TbPacote.IDVENDA
               .CellValue(i, "IDCLIENTE") = TbPacote.IDCLIENTE
               .CellValue(i, "IDPROMO") = TbPacote.IDPROMO
               .CellValue(i, "IDPACOTE") = TbPacote.IDPACOTE
               .CellValue(i, "DTEMISSAO") = TbPacote.DTEMISSAO
               .CellValue(i, "VLTOTAL") = ValBr(TbPacote.VlTotal)
               .CellValue(i, "VLDESC") = ValBr(TbPacote.VLDESC)
               .CellValue(i, "VALOR") = ValBr(TbPacote.Valor)
            End With
         End If
         Set MyPac = CriarObjeto("BANCO_3R.TB_CPACOTE")
         Set MyPac.xDb = mvarSys.xDb
         MyPac.IDLOJA = TbPacote.IDLOJA
         MyPac.IDVENDA = TbPacote.IDVENDA
         MyPac.IDCLIENTE = TbPacote.IDCLIENTE
         MyPac.IDPACOTE = TbPacote.IDPACOTE
         MyPac.DSCPACOTE = TbPacote.DSCPACOTE
         MyPac.DTEMISSAO = TbPacote.DTEMISSAO
         MyPac.IDPROMO = TbPacote.IDPROMO
         MyPac.VlTotal = TbPacote.VlTotal
         MyPac.VLDESC = TbPacote.VLDESC
         MyPac.Valor = TbPacote.Valor
         If ExisteItem(mvarPacotes, CStr(MyPac.IDPACOTE)) Then
            mvarPacotes(CStr(MyPac.IDPACOTE)).IDLOJA = TbPacote.IDLOJA
            mvarPacotes(CStr(MyPac.IDPACOTE)).IDVENDA = TbPacote.IDVENDA
            mvarPacotes(CStr(MyPac.IDPACOTE)).IDCLIENTE = TbPacote.IDCLIENTE
            mvarPacotes(CStr(MyPac.IDPACOTE)).IDPACOTE = TbPacote.IDPACOTE
            mvarPacotes(CStr(MyPac.IDPACOTE)).DSCPACOTE = TbPacote.DSCPACOTE
            mvarPacotes(CStr(MyPac.IDPACOTE)).DTEMISSAO = TbPacote.DTEMISSAO
            mvarPacotes(CStr(MyPac.IDPACOTE)).IDPROMO = TbPacote.IDPROMO
            mvarPacotes(CStr(MyPac.IDPACOTE)).VlTotal = TbPacote.VlTotal
            mvarPacotes(CStr(MyPac.IDPACOTE)).VLDESC = TbPacote.VLDESC
            mvarPacotes(CStr(MyPac.IDPACOTE)).Valor = TbPacote.Valor
         Else
            mvarPacotes.Add MyPac, CStr(MyPac.IDPACOTE)
         End If
         Set MyPac = Nothing
                  
         TbPacote.Rs.MoveNext
      Wend
   End If
   Set TbPacote = Nothing
End Sub
Private Sub IncluiGrdLinha()
   Dim lRow As Long
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim bAchou As Boolean
   Dim nColDefault As Integer
      
   nColDefault = 2
   With mvarMe.GrdPacote
      '.Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then .CellForeColor(lRow - 1, "IDPROMO") = vbBlack
      
      'If Not (mvarIDATENDIMENTO = 0 And lRow = 0) Then
      If lRow > 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = 20
   
         For i = 1 To .ColCount
            If .CellType(lRow, .ColKey(i)) = igCellCombo Then
               If .Combos(.ColKey(i)).ListCount = 1 Then
                  .CellValue(lRow, .ColKey(i)) = .Combos(.ColKey(i)).ItemValue(0)
               ElseIf .Combos(.ColKey(i)).ListCount > 1 Then
                  If GetTag(.ColTag(.ColKey(i)), "DEFAULT", "") <> "" Then
                     If IsNumeric(GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")) Then
                        .CellValue(lRow, .ColKey(i)) = Val(GetTag(.ColTag(.ColKey(i)), "DEFAULT", "0"))
                     Else
                        .CellValue(lRow, .ColKey(i)) = Val(GetTag(.ColTag(.ColKey(i)), "DEFAULT", ""))
                     End If
                     If lRow > 1 Then
                        For j = 0 To .Combos(.ColKey(i)).ListCount - 1
                           bAchou = False
                           If Not bAchou Then
                              .CellValue(lRow, .ColKey(i)) = .Combos(.ColKey(i)).ItemValue(j)
                           End If
                           For k = 1 To (lRow - 1)
                              bAchou = (.CellValue(lRow, .ColKey(i)) = .CellValue(k, .ColKey(i)))
                              If bAchou Then Exit For
                           Next
                        Next
                     End If
                  Else
                     If lRow > 1 Then
                        .CellValue(lRow, .ColKey(i)) = .CellValue(lRow - 1, .ColKey(i))
                     End If
                  End If
               End If
            Else
               If GetTag(.ColTag(.ColKey(i)), "DEFAULT", "") = "IDENTITY" Then
                  .CellValue(lRow, .ColKey(i)) = lRow
               Else
                  .CellValue(lRow, i) = GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")
               End If
            End If

            If GetTag(.ColTag(i), "EDIT", True) = False Then
               .CellForeColor(lRow, i) = vbGrayText
            End If
         Next
         If Val(.CellValue(lRow, "IDPROMO")) <> 0 Then
            Call mvarMe_GrdPacoteAfterCommitEdit(lRow, .ColIndex("IDPROMO"))
         End If
         .CellValue(lRow, "IDLOJA") = mvarIDLOJA
         .CellValue(lRow, "IDVENDA") = mvarIDVENDA
         .CellValue(lRow, "IDCLIENTE") = mvarIDCLIENTE
         .CellValue(lRow, "IDPACOTE") = lRow
         .CellValue(lRow, "DTEMISSAO") = mvarDTVENDA
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
         .SetCurCell .RowCount, "IDPROMO"
         .SetCurCell lRow, nColDefault
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub
Private Sub CalculaPromo()
   Dim i As Integer
   Dim nValor As Currency
   Dim nVltotal As Currency
   Dim nVlDesc As Currency
   
   nValor = 0
   nVltotal = 0
   nVlDesc = 0
   For i = 1 To mvarMe.GrdPacote.RowCount
      nVltotal = nVltotal + xVal(mvarMe.GrdPacote.CellValue(i, "VLTOTAL"))
      nVlDesc = nVlDesc + xVal(mvarMe.GrdPacote.CellValue(i, "VLDESC"))
      nValor = nValor + xVal(mvarMe.GrdPacote.CellValue(i, "VALOR"))
   Next
   mvarMe.TxtVLTOTAL.Text = ValBr(nVltotal)
   mvarMe.TxtVLDESC.Text = ValBr(nVlDesc)
   mvarMe.TxtVALOR.Text = ValBr(nValor)
   If xVal(mvarMe.TxtVLTOTAL.Text) > 0 Then
      mvarMe.LblDesc.Caption = ValBr(100 * (xVal(mvarMe.TxtVLDESC.Text) / xVal(mvarMe.TxtVLTOTAL.Text))) & "%"
   Else
      mvarMe.LblDesc.Caption = "0%"
   End If
End Sub
Public Sub Show()
   Dim Sql As String
   Dim bOk As Boolean
   bOk = True
   If mvarIDVENDA = 0 Then
      Sql = Sql & " Select *"
      Sql = Sql & " From CPROMOCAO"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And ATIVO=1"
      Sql = Sql & " And DTINIV<=" & SqlDate(mvarSys.xDb.SysDate() & " 00:00")
      Sql = Sql & " And DTFIMV>=" & SqlDate(mvarSys.xDb.SysDate() & " 23:59")
      bOk = mvarSys.xDb.ExisteReg(Sql)
   End If
   If bOk Then
      mvarMe.Show vbModal
   Else
      Call ExibirAviso("Não existem promoções cadastradas para venda de pacotes")
   End If
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lRow As Integer
   If KeyCode = vbKeyDelete Then
      With mvarMe
         If .ActiveControl Is .GrdPacote Then
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
               For lRow = 1 To .RowCount - 1
                  .CellValue(lRow, 1) = lRow
               Next
            End With
            Call CalculaPromo
         End If
      End With
   End If
End Sub

Private Sub mvarMe_Load()
   bLoad = True
   Call MontaGrdPacote
   Call PopulaGrdPacote
   Screen.MousePointer = vbDefault
End Sub
Private Sub LimpaTela()
   mvarMe.GrdPacote.Clear False
   mvarMe.TxtVLTOTAL.Text = "0,00"
   mvarMe.TxtVLDESC.Text = "0,00"
   mvarMe.TxtVALOR.Text = "0,00"
   mvarMe.LblDesc.Caption = "0%"
End Sub
