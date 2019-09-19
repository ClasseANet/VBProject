VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CupomVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmCUPOMVENDA
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDVENDA        As Long
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
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   If mvarIDLOJA = 0 Then mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmCUPOMVENDA
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_Activate()
   Call MontaTela
   'Call PopulaTela
   If GetTag(mvarMe, "1VEZ", "1") = "1" Then
      Call SetTag(mvarMe, "1VEZ", "0")
   End If
End Sub
Private Sub mvarMe_Load()
   'Call MontaTela
   'Call PopulaTela
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdOkClick()
   If SalvarTela Then
      Unload mvarMe
   End If
End Sub
Private Function SalvarTela() As Boolean

   SalvarTela = True
End Function
Private Sub MontaTela()
   Dim TbPromo    As Object 'TB_CPROMOCAO
   Dim TbCupom    As Object 'TB_CCUPOM
   
   Dim sAux       As String
   Dim i          As Integer
   
   Set TbPromo = CriarObjeto("BANCO_3R.TB_CPROMOCAO")
   Set TbPromo.xDb = mvarSys.xDb
   Set TbCupom = CriarObjeto("BANCO_3R.TB_CCUPOM")
   Set TbCupom.xDb = mvarSys.xDb
   
   With mvarMe.GrdCupom
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
            
      With .Combos
         With .Add("IDPROMO")
            If TbPromo.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1", Ch_ORDERBY:="IDPROMO") Then
               While Not TbPromo.Rs.EOF
                  TbPromo.Popula
                  .AddItem TbPromo.DSCPROMO, TbPromo.IDPROMO
                  TbPromo.Rs.MoveNext
               Wend
            End If
            Set TbPromo = Nothing
            .AutoAdjustWidth
         End With
         With .Add("IDCUPOM")
            'If TbCupom.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPOPROMO:=1, Ch_ORDERBY:="CODCUPOM") Then
            '   While Not TbCupom.RS.EOF
            '      TbCupom.Popula
            '      .AddItem TbCupom.CODCUPOM, TbCupom.IDCUPOM
            '      TbCupom.RS.MoveNext
            '   Wend
            'End If
            'Set TbCupom = Nothing
            '.AutoAdjustWidth
            .Width = 10
         End With
      
      End With
      
      With .AddCol(sKey:="ITEM", sHeader:="#", lWidth:=25, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "ITEM"
         .bSelected = False
      End With
      With .AddCol(sKey:="IDLOJA", sHeader:="Loja", lWidth:=85, bVisible:=False)
         .eType = igCellCombo
         .eTypeFlags = igTextBoxHScrollBar
         .sCtrlKey = "IDLOJA"
      End With
      With .AddCol(sKey:="IDVENDA", sHeader:="Loja", lWidth:=85, bVisible:=False)
         .eType = igCellCombo
         .eTypeFlags = igTextBoxHScrollBar
         .sCtrlKey = "IDVENDA"
      End With
      With .AddCol(sKey:="IDPROMO", sHeader:="Desc/Promoção", lWidth:=90, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDPROMO"
      End With
      With .AddCol(sKey:="IDCUPOM", sHeader:="Cupom", lWidth:=90, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDCUPOM"
      End With
      With .AddCol(sKey:="VLDESC", sHeader:="Valor", lWidth:=60, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "VLDESC"
         .eTextFlags = igTextCenter
      End With
      With .AddCol(sKey:="QTD", sHeader:="Qtd.", lWidth:=60, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "QTD"
         .eTextFlags = igTextCenter
      End With
      With .AddCol(sKey:="TOTAL", sHeader:="Total", lWidth:=60, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "TOTAL"
         .eTextFlags = igTextCenter
      End With
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
                       
      .ColTag("ITEM") = SetTag(.ColTag("ITEM"), "EDIT", False)
      .ColTag("IDLOJA") = SetTag(.ColTag("IDLOJA"), "EDIT", False)
      .ColTag("IDVENDA") = SetTag(.ColTag("IDVENDA"), "EDIT", False)
      .ColTag("TOTAL") = SetTag(.ColTag("TOTAL"), "EDIT", False)
            
      .ColTag("QTD") = SetTag(.ColTag("QTD"), "EDITOPT", igTextEditNumberOnly)
      .ColTag("VLDESC") = SetTag(.ColTag("VLDESC"), "EDITOPT", igTextEditNumberOnly)
      .ColTag("TOTAL") = SetTag(.ColTag("TOTAL"), "EDITOPT", igTextEditNumberOnly)
                                   
      IncluiGrdLinha
   End With
End Sub
Private Sub IncluiGrdLinha()
   Dim lRow As Long
   Dim i As Integer
   Dim nColDefault As Integer
   
'   On Error Resume Next
   nColDefault = 4
   With mvarMe.GrdCupom
      .Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then .CellForeColor(lRow - 1, "ITEM") = vbBlack
      
      .AddRow
      lRow = .RowCount
      .RowHeight(.RowCount) = 20

      .CellType(lRow, "ITEM") = igCellText
      .CellValue(lRow, "ITEM") = lRow
      
      .CellType(lRow, "IDPROMO") = igCellCombo
      .CellValue(lRow, "IDPROMO") = 2
      If lRow > 1 Then .CellValue(lRow, "IDPROMO") = .CellValue(lRow - 1, "IDPROMO")
      

      .CellType(lRow, "IDCUPOM") = igCellCombo
      .CellValue(lRow, "IDCUPOM") = 1
      If lRow > 1 Then .CellValue(lRow, "IDCUPOM") = .CellValue(lRow - 1, "IDCUPOM")

      .CellType(lRow, "VLDESC") = igCellText
      .CellValue(lRow, "VLDESC") = "0"
      
      .CellType(lRow, "QTD") = igCellText
      .CellValue(lRow, "QTD") = "0"
      
      .CellType(lRow, "TOTAL") = igCellText
      .CellValue(lRow, "TOTAL") = "0"
            
      .CellForeColor(lRow, "ITEM") = vbBlack
      
      For i = 1 To .ColCount
         If GetTag(.ColTag(i), "EDIT", True) = False Then
            .CellForeColor(lRow, i) = vbGrayText
         End If
      Next
      
            
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
      .SetCurCell .RowCount, 1
      .SetCurCell lRow, nColDefault
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub

Private Sub PopulaTela()

End Sub

