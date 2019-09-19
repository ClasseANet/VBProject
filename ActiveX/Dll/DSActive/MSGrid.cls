VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSGrid"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum eCheck
   Checked
   Unchecked
   Deny
End Enum

Public WithEvents Grd  As MSFlexGrid
Attribute Grd.VB_VarHelpID = -1

Public WithEvents mvarCtrl As VBControlExtender
Attribute mvarCtrl.VB_VarHelpID = -1
Public WithEvents mvarCtrl1 As VBControlExtender
Attribute mvarCtrl1.VB_VarHelpID = -1
Public WithEvents mvarCtrl2 As VBControlExtender
Attribute mvarCtrl2.VB_VarHelpID = -1
Public WithEvents mvarCtrl3 As VBControlExtender
Attribute mvarCtrl3.VB_VarHelpID = -1
Public WithEvents mvarCtrl4 As VBControlExtender
Attribute mvarCtrl4.VB_VarHelpID = -1
Public WithEvents mvarCtrl5 As VBControlExtender
Attribute mvarCtrl5.VB_VarHelpID = -1
Public WithEvents mvarCtrl6 As VBControlExtender
Attribute mvarCtrl6.VB_VarHelpID = -1
Public WithEvents mvarCtrl7 As VBControlExtender
Attribute mvarCtrl7.VB_VarHelpID = -1
Public WithEvents mvarCtrl8 As VBControlExtender
Attribute mvarCtrl8.VB_VarHelpID = -1
Public WithEvents mvarCtrl9 As VBControlExtender
Attribute mvarCtrl9.VB_VarHelpID = -1
Public WithEvents mvarCtrl10 As VBControlExtender
Attribute mvarCtrl10.VB_VarHelpID = -1


Private mvarCollLin        As New Collection

Private mvarMinLin         As Integer
Private mvarMaxLin         As Integer
Private mvarSorted         As Boolean
Private mvarSortCol        As Integer
Private mvarSortAscending  As Boolean
Private mvarCabecalho      As Variant
Private mvarModoEdicao     As Boolean

Public SetColorLine        As Boolean

Public RunTopRow           As Boolean
Public RunColorLine        As Boolean
Private mvarEditandoLinha  As Boolean
Public Property Let EditandoLinha(ByVal vData As Boolean)
    mvarEditandoLinha = vData
End Property
Public Property Get EditandoLinha() As Boolean
    EditandoLinha = mvarEditandoLinha
End Property
Public Property Let Sorted(ByVal vData As Boolean)
    mvarSorted = vData
End Property
Public Property Get Sorted() As Boolean
    Sorted = mvarSorted
End Property
Public Property Let ModoEdicao(ByVal vData As Boolean)
   Dim n As Variant
    mvarModoEdicao = vData
    
    On Error Resume Next
    If Not vData Then
       For Each n In mvarCollLin
          n.Visible = False
       Next
    End If
End Property
Public Property Get ModoEdicao() As Boolean
    ModoEdicao = mvarModoEdicao
End Property
Public Property Let Cabecalho(ByVal vData As Variant)
    mvarCabecalho = vData
End Property
Public Property Get Cabecalho() As Variant
    Cabecalho = mvarCabecalho
End Property
'Public Property Set Grd(ByVal vData As Object)
'    Set Grd = vData
'End Property
'Public Property Get Grd() As Object
'    Set Grd = Grd
'End Property
Public Property Let MaxLin(ByVal vData As Integer)
    mvarMaxLin = vData
End Property
Public Property Get MaxLin() As Integer
    MaxLin = mvarMaxLin
End Property
Public Property Let MinLin(ByVal vData As Integer)
    vData = IIf(vData < 2, 2, vData)
    mvarMinLin = vData
End Property
Public Property Get MinLin() As Integer
    MinLin = mvarMinLin
End Property
Public Property Let SortAscending(ByVal vData As Boolean)
    mvarSortAscending = vData
End Property
Public Property Get SortAscending() As Boolean
    SortAscending = mvarSortAscending
End Property
Public Property Let SortCol(ByVal vData As Integer)
    mvarSortCol = vData
End Property
Public Property Get SortCol() As Integer
    SortCol = mvarSortCol
End Property
Public Property Set CollLin(ByVal vData As Object)
    Set mvarCollLin = vData
End Property
Public Property Get CollLin() As Collection
    If mvarCollLin Is Nothing Then
       Set mvarCollLin = New Collection
    End If
    Set CollLin = mvarCollLin
End Property
Public Property Let CellValue(Linha As Integer, Coluna As Integer, vData As Variant)
   On Error Resume Next
   Select Case TypeName(mvarCollLin(CStr(Coluna)))
      Case "CheckBox"
         If vData = vbChecked Or vbChecked = True Then
            Call ClsCtrl.SetTag(Grd, "CEL(" & CStr(Linha) & "," & CStr(Coluna) & ")", "True")
         ElseIf vData = vbUnchecked Or vbChecked = False Then
            Call ClsCtrl.SetTag(Grd, "CEL(" & CStr(Linha) & "," & CStr(Coluna) & ")", "False")
         End If
      Case "String"
         Grd.TextMatrix(Linha, Coluna) = vData
      Case Else
         mvarCollLin(Coluna) = vData
   End Select
   If Err <> 0 Then
      Call ClsCtrl.SetTag(Grd, "CEL(" & CStr(Linha) & "," & CStr(Coluna) & ")", vData)
   End If
End Property
Public Property Get CellValue(Linha As Integer, Coluna As Integer) As Variant
   On Error Resume Next
   Select Case TypeName(mvarCollLin(CStr(Coluna)))
      Case "CheckBox"
         If ClsCtrl.GetTag(Grd, "CEL(" & CStr(Linha) & "," & CStr(Coluna) & ")") = "True" Then
            CellValue = vbChecked
         Else
            CellValue = vbUnchecked
         End If
      Case "String"
         CellValue = Grd.TextMatrix(Linha, Coluna)
      Case Else
         CellValue = mvarCollLin(Coluna)
   End Select
   If Err <> 0 Then
      CellValue = ClsCtrl.GetTag(Grd, "CEL(" & CStr(Linha) & "," & CStr(Coluna) & ")")
   End If
End Property
'************************************************************************************
'   Função  :  Deseleciona Linha do Grd. Esta função é utilizado para
'           grids que possuam na primeira coluna Figuras que identificam
'           o status da linha.
'
'   Entrada :  mfgGrid = Grd
'              nLinha = Linha a ser deselecionada
'              cValorFlagCancel = texto a ser colocado na coluna onde
'               será carregado a figura. Para melhor identificar o status da linha
'              além da figura é atribuido um texto à coluna.
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Public Sub DeselecionaItemFG(mfgGrid As Object, nLinha As Byte, cValorFlagCancel$)
        
'Área de declarações
Dim nMarcacao As Byte
Dim nFoco As Byte

'Área de manipulação de dados
    nMarcacao = mfgGrid.HighLight
    nFoco = mfgGrid.FocusRect
    mfgGrid.HighLight = flexHighlightNever
    mfgGrid.FocusRect = flexFocusNone
        
    mfgGrid.Row = nLinha: mfgGrid.Col = 0
    mfgGrid.CellPictureAlignment = flexAlignCenterCenter
    Set mfgGrid.CellPicture = LoadResPicture("CANCEL", vbResBitmap)
    mfgGrid.TextMatrix(nLinha, mfgGrid.Cols - 1) = cValorFlagCancel
        
    mfgGrid.HighLight = nMarcacao
    mfgGrid.FocusRect = nFoco
        
End Sub
Public Sub CheckGrid(Optional Value, Optional iRow, Optional iCol, Optional Tipo = 3, Optional AlignImage As AlignmentSettings = flexAlignCenterCenter)
   Dim sImg As String
   Dim bAux As Boolean
   
   If IsMissing(iRow) Then iRow = Grd.Row
   If IsMissing(iCol) Then iCol = Grd.Col
   
   bAux = RunColorLine
   RunColorLine = True

   If IsMissing(Value) Then
      If Tipo = 3 Then
         Select Case CellValue(CInt(iRow), CInt(iCol))
            Case eCheck.Checked:    Value = eCheck.Deny
            Case eCheck.Unchecked:  Value = eCheck.Checked
            Case eCheck.Deny:       Value = eCheck.Unchecked
            Case Else:              Value = eCheck.Unchecked
         End Select
      Else
         Select Case CellValue(CInt(iRow), CInt(iCol))
            Case eCheck.Checked:    Value = eCheck.Unchecked
            Case eCheck.Unchecked:  Value = eCheck.Checked
            Case Else:              Value = eCheck.Unchecked
         End Select
      End If
   End If
   
   Select Case Value
      Case eCheck.Checked:    sImg = "CHECKED"
      Case eCheck.Unchecked:  sImg = "UNCHECKED"
      Case eCheck.Deny:       sImg = "DENY"
   End Select

   With Grd
      .HighLight = flexHighlightNever
      .FocusRect = flexFocusNone
      CellValue(CInt(iRow), CInt(iCol)) = Value
      If Not IsMissing(iRow) Then .Row = iRow
      If Not IsMissing(iCol) Then .Col = iCol
      
      If sImg <> "" Then
         Set .CellPicture = LoadResPicture(sImg, vbResBitmap)
      End If
      .CellForeColor = .CellBackColor
      If .CellBackColor = 0 Then
         .CellForeColor = vbWhite
      End If
      
      .CellPictureAlignment = AlignImage 'flexAlignCenterCenter
      .FocusRect = flexFocusLight
      .HighLight = flexHighlightAlways
   End With
   If bAux Then
      Call PintarLinha(iRow)
   End If
   RunColorLine = False
End Sub '************************************************************************************
'   Função  : Carrega uma imagem numa linha do Grd ao contrário da atual,
'           ou seja, se a linha estiver OK passrá a estar cancelada e vice-versa.
'           OBS: Usada no Click do Grid
'
'   Entrada :  mfgGrid = Grd
'              nLinha = Linha onde a imagem será carregada
'              cValorFlagOk = texto que indica que a linha está OK.
'              cValorFlagCancel = texto que indica que a linha está CANCEL.
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Public Sub EstabeleceImagemItemFG(mfgGrid As Object, nLinha As Byte, cValorFlagOk$, _
    cValorFlagCancel$)
    
'Área de declarações
Dim nMarcacao As Byte
Dim nFoco As Byte

'Área de manipulação de dados
    nMarcacao = mfgGrid.HighLight
    nFoco = mfgGrid.FocusRect
    mfgGrid.HighLight = flexHighlightNever
    mfgGrid.FocusRect = flexFocusNone
    
    mfgGrid.Col = 0: mfgGrid.Row = nLinha
    mfgGrid.CellPictureAlignment = flexAlignCenterCenter
    If mfgGrid.TextMatrix(nLinha, mfgGrid.Cols - 1) = cValorFlagOk Then
        Set mfgGrid.CellPicture = LoadResPicture("CANCEL", vbResBitmap)
        mfgGrid.TextMatrix(nLinha, mfgGrid.Cols - 1) = cValorFlagCancel
    ElseIf mfgGrid.TextMatrix(nLinha, mfgGrid.Cols - 1) = cValorFlagCancel Then
        Set mfgGrid.CellPicture = LoadResPicture("OK", vbResBitmap)
        mfgGrid.TextMatrix(nLinha, mfgGrid.Cols - 1) = cValorFlagOk
    End If

    mfgGrid.HighLight = nMarcacao
    mfgGrid.FocusRect = nFoco

End Sub
'************************************************************************************
'   Função  :   Alinha uma coluna do Grid
'
'   Entrada :  pGrd = pGrd
'              col% = Coluna a ser alinhada
'              alinha$ = Tipo de alinhamento
'                       "E" - esquerda
'                       "D" - direita
'                       "C" - centrelizado
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Public Sub GridAlign(pGrd As Object, ByVal Col%, ByVal Alinha$)
    Select Case Alinha
    Case "E": pGrd.ColAlignment(Col) = 0
    Case "D": pGrd.ColAlignment(Col) = 1
    Case "C": pGrd.ColAlignment(Col) = 2
    End Select
End Sub
Public Sub GrdClear(pGrd As Object)
'----------------------------------------------------------------------------
'
'   Objetivo: limpar a pGrd.
'   Entrada : pGrd - pGrd a qual se deseja limpar.
'   Saída   : --
'
'   Autor   : Diogenes                  Data de criação:28/05/99
'   Versão  : 1.0
'
'----------------------------------------------------------------------------
   pGrd.Rows = pGrd.FixedRows
End Sub
'************************************************************************************
'   Função  :  Copia o conteúdo de um Grid para outro.
'
'   Entrada :  GrdOrigem = Grid de Origem
'              grdDestino =  Grid de Destino
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Public Sub GridCopia(GrdOrigem As Object, grdDestino As Object)
   Dim i%, j%, LcItm$
   'carregando a Grd origem para destino
   If GrdOrigem.Rows > 1 Then
      Call GrdClear(grdDestino)
      For i = 2 To GrdOrigem.Rows - 1
         LcItm = ValorCel(GrdOrigem, i, 0)
         For j = 1 To GrdOrigem.Cols - 1
            LcItm = LcItm & Chr(9) & ValorCel(GrdOrigem, i, j)
         Next j
         grdDestino.AddItem LcItm, i
      Next i
   End If
End Sub
'************************************************************************************
'   Função  :  Retorna a linha onde um texto específico foi encontrado.
'
'   Entrada :  pGrd = pGrd
'              icol% = coluna onde será realizada a pesquisa
'              Txt$ = texto a ser pesquisado
'
'   Saída   :   Linha onde o texto foi encontrado
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Function GridIndex(pGrd As Object, ByVal iCol%, ByVal Txt$) As Integer

    Dim i%, AuxRow%, AuxCol%

    Txt$ = Trim(Txt$)
    GridIndex = -1
    AuxRow% = pGrd.Row
    AuxCol% = pGrd.Col
    pGrd.Col = iCol%
    For i = 0 To pGrd.Rows - 1
        pGrd.Row = i
        If Trim(pGrd.Text) = Txt$ Then
            GridIndex = i
            Exit For
        End If
    Next
    pGrd.Row = AuxRow%
    pGrd.Col = AuxCol%
End Function
Public Function GrdKey(ByRef pGrd As Object, Key%, Tipo%, Optional NewLine = False)
   Dim Tam%, TopLin%
'* Verificar teclas de função
   Select Case Key
      Case vbKeySnapshot: pGrd.Text = pGrd.Text & Chr(Key)
      Case vbKeyDelete
         Select Case Tipo
            Case vbInteger: pGrd.Text = ""
            Case vbCurrency: pGrd.Text = ""
            Case vbDate: pGrd.Text = "  /  /  "
            Case vbString: pGrd.Text = ""
         End Select
         GrdKey = pGrd.Text
         pGrd.SetFocus
         GoTo Fim
      Case vbKeyBack
         If Tipo = vbDate Then pGrd.Text = ClsDsr.EliminarString(pGrd.Text, "/")
         Tam = Len(Trim(pGrd.Text))
         pGrd.Text = Left(pGrd.Text, IIf(Tam <= 0, 1, Tam) - 1)
      Case vbKeyDown: pGrd.Row = IIf(pGrd.Row = pGrd.Rows - 1, pGrd.Row, pGrd.Row + 1)
      Case vbKeyReturn, vbKeyTab
         If pGrd.Col = pGrd.Cols - 1 Then
            If pGrd.Row = pGrd.Rows - 1 Then
               If pGrd.TextMatrix(pGrd.Row, 0) <> "" Then
                  If NewLine Then
                     pGrd.Rows = pGrd.Rows + 1
                     pGrd.Row = pGrd.Rows - 1
                  End If
               End If
            Else
               pGrd.Row = pGrd.Row + 1
               pGrd.Col = 0
               'MoveCol
               While pGrd.ColWidth(pGrd.Col) = 0 Or pGrd.FixedCols >= pGrd.Col + 1
                  If pGrd.Col = pGrd.Cols Then
                     If pGrd.Row = pGrd.Rows Then
                        If NewLine Then
                           pGrd.Rows = pGrd.Rows + 1
                           pGrd.Row = pGrd.Rows
                        End If
                     Else
                        pGrd.Row = pGrd.Row + 1
                     End If
                  Else
                     pGrd.Col = pGrd.Col + 1
                  End If
               Wend
            End If
         Else
            pGrd.Col = pGrd.Col + 1
            'MoveCol
            While pGrd.ColWidth(pGrd.Col) = 0
               If pGrd.Col = pGrd.Cols Then
                  If pGrd.Row = pGrd.Rows Then
                     If NewLine Then
                        pGrd.Rows = pGrd.Rows + 1
                        pGrd.Row = pGrd.Rows
                     End If
                  Else
                     pGrd.Row = pGrd.Row + 1
                  End If
               Else
                  pGrd.Col = pGrd.Col + 1
               End If
            Wend
         End If
         GrdKey = pGrd.Text
         pGrd.SetFocus
         GoTo Fim
         'pGrd.Row = IIf(pGrd.Row = pGrd.Rows - 1, pGrd.Row, pGrd.Row + 1)
      Case vbKeyUp: If pGrd.Row > 1 Then pGrd.Row = pGrd.Row - 1
   End Select
   
'* Inicializar Celula
   Select Case Tipo
      Case vbInteger, vbCurrency
         pGrd.Text = IIf(pGrd.Text = "0", "", pGrd.Text)
         Tam% = Len(pGrd.Text)
      Case vbDate
         pGrd.Text = IIf(pGrd.Text = "", "  /  /  ", pGrd.Text)
         pGrd.Text = Trim(ClsDsr.EliminarString(pGrd.Text, "/"))
         Tam% = Len(pGrd.Text)
      Case vbString
         pGrd.Text = Trim(pGrd.Text)
         Tam% = Len(pGrd.Text)
   End Select
'* Trata a tecla Pressionada
   If ClsDsr.Between(Key, vbKey0, vbKey9) Then
      pGrd.Text = pGrd.Text & Chr(Key)
   End If
   If ClsDsr.Between(Key, vbKeyNumpad0, vbKeyNumpad9) Then
      Key = Key - vbKey0
      pGrd.Text = pGrd.Text & Chr(Key)
   End If
   If Tipo = vbCurrency Then
      If ClsDsr.InArray(Key, Array(vbKeyDecimal, 188, 190)) Then
         pGrd.Text = pGrd.Text & ","
      End If
   End If
'* Finaliza
   Select Case Tipo
      Case vbInteger
         pGrd.Text = IIf(pGrd.Text = "", "0", pGrd.Text)
      Case vbDate
         pGrd.Text = Left(Left(pGrd, 2) & Space(2), 2) & "/" + Left(Mid(pGrd, 3, 2) & Space(2), 2) & "/" & Left(Mid(pGrd, 5) & Space(4), 4)
         pGrd.Text = IIf(pGrd.Text = "", "  /  /  ", pGrd.Text)
   End Select
   GrdKey = pGrd.Text
Fim:
   TopLin = pGrd.RowSel - (Int(pGrd.Height / pGrd.CellHeight) - (pGrd.FixedRows + 2))
   If TopLin > 0 Then
      If pGrd.TopRow < TopLin Then
         pGrd.TopRow = TopLin
      End If
   End If
End Function
Public Function GrdMousePos(pGrd As Object, X As Single, Y As Single, ByRef Pos As Collection)
'*****************************************
'* Determina Linha e Coluna do Mouse     *
'* no evento Mouse_Click ou Mouse_Down   *
'*                                       *
'* Obs : Com o controle MsFlexGrid esta  *
'*     função se torna inútil, pois ele  *
'*     já contém as propriedades abaixo  *
'*     pGrd.MouseCol e pGrd.MouseRow.      *
'*                                       *
'*****************************************
   Dim Soma&, Lin%, Col%, i%
   Set Pos = New Collection
   Lin% = -1
   Col% = -1
   For i = 0 To pGrd.Rows - 1
      Soma = Soma + pGrd.RowHeight(i)
      If Y <= Soma Then
         Lin% = i
         Soma = 0
         Exit For
      End If
   Next
   'Pos.Add Lin%, "LIN"
   Pos.Add i, "LIN"
   For i = 0 To pGrd.Cols - 1
       Soma = Soma + pGrd.ColWidth(i)
      If X <= Soma Then
         Col% = i
         Soma = 0
         Exit For
      End If
   Next
   'Pos.Add Col%, "COL"
   Pos.Add i, "COL"
End Function
'************************************************************************************
'   Função  :  Atribui a uma célula do pGrd um texto especifico.
'
'   Entrada : pGrd - pGrd
'             irow% - linha a qual se deseja escrever o texto
'             icol% - linha a qual se deseja escrever o texto
'             Txt - text que se deseja atribuir à célula
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Sub GridSetText(pGrd As Object, ByVal iRow%, ByVal iCol%, ByVal Txt)
    
    Dim i%, AuxRow%, AuxCol%
    
    AuxRow% = pGrd.Row
    AuxCol% = pGrd.Col
    pGrd.Row = iRow
    pGrd.Col = iCol
    pGrd.Text = Txt
    pGrd.Row = AuxRow%
    pGrd.Col = AuxCol%

End Sub
'************************************************************************************
'   Função  :  Soma os valores de uma coluna do pGrd
'
'   Entrada : pGrd - pGrd
'             col% - coluna a qual se deseja realizar a soma
'
'   Saída   : Valor do somatório da coluna
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Function GridSumCol(pGrd As Object, ByVal Col%) As Variant
    Dim i%, sum
    sum = 0
    For i = 0 To pGrd.Rows - 1
'dioxxx        sum = sum + Val(strReplace(ValorCel(pGrd, i, Col), ",", "."))
    Next i
    GridSumCol = sum
End Function
'************************************************************************************
'   Função  :  Recupera o texto de uma célula.
'
'   Entrada : Grd - Grd
'             irow% - linha a qual se deseja recuperar o texto
'             icol% - coluna a qual se deseja recuperar o texto
'
'   Saída   : Texto da célula
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Function ValorCel(Grd As Object, ByVal iRow%, ByVal iCol%) As String
   ValorCel = Grd.TextMatrix(iRow, iCol)
End Function
'************************************************************************************
'   Função  :  colocar um título na pGrd e dimensionar as colunas da pGrd.
'
'   Entrada : pGrd       - pGrd a qual se deseja colocar o título e
'                          dimensionar as colunas.
'             headers$   - "string" contendo os títulos das colunas
'                          separados por ",".
'             colswidth$ - "string" contendo as larguras de cada coluna
'                          separadas por ",".
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Sub GridTitle(pGrd As Object, ByVal headers$, ByVal colswidth$)
'----------------------------------------------------------------------------
'
'   Objetivo: colocar um título na pGrd e dimensionar as colunas da pGrd.
'   Entrada : pGrd       - pGrd a qual se deseja colocar o título e
'                          dimensionar as colunas.
'             headers$   - "string" contendo os títulos das colunas
'                          separados por ",".
'             colswidth$ - "string" contendo as larguras de cada coluna
'                          separadas por ",".
'   Saída   : --
'
'   Autor   : Tsai                  Data de criação:07/02/94
'   Versão  : 1.0
'
'   Histórico de modificações:
'   Data    Autor       Versão  Modificação
'
'----------------------------------------------------------------------------
    Dim i%, Pos%, PosIni%, posfim%, nCol%

'   troca de "," por "TAB"

    Pos% = InStr(headers$, ",")
    nCol = 1
    While Pos% <> 0
        Mid(headers$, Pos%, 1) = Chr(9)
        Pos% = InStr(headers$, ",")
        nCol = nCol + 1
    Wend
    pGrd.Cols = nCol
'   coloca títulos nas colunas da pGrd

    If pGrd.FixedRows > 0 Then
        pGrd.FixedRows = 0
        pGrd.RemoveItem 0
    End If
    If Left$(headers$, 1) <> Chr(9) Then
        pGrd.AddItem headers$, 0
        pGrd.FixedRows = 1
    End If

'   dimensiona as colunas da pGrd

    On Error GoTo Fim
    PosIni% = 1
    For i% = 0 To pGrd.Cols - 1
        posfim% = InStr(PosIni%, colswidth$, ",")
        If posfim% = 0 Then posfim% = Len(colswidth$) + 1
        pGrd.ColWidth(i%) = Val(Mid(colswidth$, PosIni%, posfim% - PosIni%))
        PosIni% = posfim% + 1
    Next
    Exit Sub

Fim:

'   evita erro quando houver mais colunas do que larguras na "string"

    On Error GoTo 0
    Exit Sub
End Sub
Public Sub HintMSGrid(pGrd As Object, X As Single, Y As Single)
   Dim C%, L%, Cont%, Soma As Double
   Dim Txt$
'   cont% = pGrd.Cols - 1
'   C% = 0
'   Soma = pGrd.ColWidth(C)
'   While X > Soma
'      C% = IIf(C% >= cont%, cont%, C% + 1)
'      Soma = Soma + pGrd.ColWidth(C)
'   Wend
'   L% = Int(Y / pGrd.RowHeight(L%)) - 1
'   If L < 0 Or L >= pGrd.Rows Then Exit Sub
'   If C < 0 Or C > pGrd.Rows Then Exit Sub
'   Txt$ = ValorCel(pGrd, L%, C%)
'   If L < 0 Or (Len(Txt) * 120) < pGrd.ColWidth(C) Then
'      Txt$ = ""
'   End If
'   pGrd.ToolTipText = Txt$
End Sub
Public Sub Inicializar()
   Dim n As Variant
   
   If Grd Is Nothing Then Exit Sub
   If mvarCollLin Is Nothing Then Exit Sub
      
'   Call LoadControl
'Set mvarCtrl = Grd.Parent.Controls(28)

   Call MontarCabGrid(Grd, mvarCabecalho)
'   Call EditaLinha(0, False)
   Call MoverLinha
   If mvarMinLin > 0 Then Grd.Rows = mvarMinLin
   
   SetColorLine = False
   
   Dim i As Integer
   i = 0
'   For Each n In mvarCollLin
'      Licenses.Add "Project1.Control1", "ewrinvcmcoe"
'      objExt.Visible = True ' The control is invisible by default.
'      Load mvarCtrl(i)
'      Set mvarCtrl = n
'      i = i + 1
'   Next
   For Each n In mvarCollLin
      If TypeName(n) = "CheckBox" Then
         SetColorLine = True
         Exit For
      End If
      
'      If TypeName(n) <> "String" And TypeName(n) <> "TextBox" Then
'      'If VarType(n) = vbObject Then
'         If mvarCtrl1 Is Nothing Then
'            Set mvarCtrl1 = n
'         ElseIf mvarCtrl2 Is Nothing Then
'            Set mvarCtrl2 = n
'         ElseIf mvarCtrl3 Is Nothing Then
'            Set mvarCtrl3 = n
'         ElseIf mvarCtrl4 Is Nothing Then
'            Set mvarCtrl4 = n
'         ElseIf mvarCtrl5 Is Nothing Then
'            Set mvarCtrl5 = n
'         ElseIf mvarCtrl6 Is Nothing Then
'            Set mvarCtrl6 = n
'         ElseIf mvarCtrl7 Is Nothing Then
'            Set mvarCtrl7 = n
'         ElseIf mvarCtrl8 Is Nothing Then
'            Set mvarCtrl8 = n
'         ElseIf mvarCtrl9 Is Nothing Then
'            Set mvarCtrl9 = n
'         ElseIf mvarCtrl10 Is Nothing Then
'            Set mvarCtrl10 = n
'         End If
'      End If
   Next
   
   Grd.Row = 0
   
End Sub
Private Sub LoadControl()
   Licenses.Add "Project1.Control1", "ewrinvcmcoe"
   Set mvarCtrl = Grd.Parent.Controls.Add("Project1.Control1", "myCtl")
   mvarCtrl.Visible = True ' The control is invisible by default.
   mvarCtrl.Left = 60
   mvarCtrl.Top = 60
End Sub

'************************************************************************************
'   Função  :  Seta propriedades basicas do pGrd, se for um pGrd
'              que conterá imagens para indicar status de linha já
'              atribui imagem título.
'
'   Entrada : frm - form onde está localizado o pGrd, será atualizado
'           suas propriedades de fontes.
'             mfgGrid  - pGrd que será inicializado
'             bImagem - indica se o pGrd conterá imagem
'             cTipoSelecao - tipo de seleçào do pGrd, por "L"inha ou "C"oluna
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Public Sub InicializarFG(frm As Object, pGrd As Object, _
    Optional bImagem As Variant, Optional cTipoSelecao As Variant)

'Área de configuração
    With pGrd
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .Rows = pGrd.FixedRows + 1
        .AllowUserResizing = flexResizeNone
        .AllowBigSelection = False
        .FixedRows = 1
        .ScrollBars = 3
        .ScrollTrack = True
        .WordWrap = True
        
        If Not IsMissing(bImagem) = True Then
            If bImagem = True Then
                .FixedCols = 0
                .RowHeightMin = 480
                .Col = 0: .Row = 0
                .CellPictureAlignment = flexAlignCenterCenter
                Set .CellPicture = LoadResPicture("OK_CANCEL", _
                    vbResBitmap)
                .ColWidth(0) = 600
                .ColWidth(pGrd.Cols - 1) = 0
                '.ToolTipText = _
                '"<BARRA DE ESPAÇO> - Seleciona ítem"
            End If
        End If
        .SelectionMode = flexSelectionFree
        If IsMissing(cTipoSelecao) = True Then
            .HighLight = flexHighlightWithFocus
            .FocusRect = flexFocusHeavy
        Else
            Select Case UCase(Trim(cTipoSelecao))
            Case "L"
                .SelectionMode = flexSelectionByRow
                .HighLight = flexHighlightAlways
                .FocusRect = flexFocusNone
            Case "C"
                .HighLight = flexHighlightWithFocus
                .FocusRect = flexFocusHeavy
            End Select
        End If
    
        .FontName = "Times News Roman"
        If .FontSize < 10 Then .FontSize = 10
'        FrClsMsg.FontName = .FontName
'        FrClsMsg.FontSize = .FontSize
'        FrClsMsg.FontBold = .FontBold
'        FrClsMsg.FontItalic = .FontItalic
'        FrClsMsg.KeyPreview = True
    End With
End Sub
Public Sub LimparLinhaGrd(Optional ObjGrd, Optional Linha)
   Dim i As Integer, n As Variant
   Dim nAux As Integer
   
   On Error Resume Next
   
   If IsMissing(ObjGrd) And Grd Is Nothing Then Exit Sub
   
   If IsMissing(ObjGrd) And Not Grd Is Nothing Then
      Set ObjGrd = Grd

      For Each n In mvarCollLin
         If TypeName(n) = "ComboBox" Then
            n.ListIndex = -1
         ElseIf TypeName(n) = "CheckBox" Then
            n.Value = vbUnchecked
            
         ElseIf TypeName(n) <> "String" Then
            n.Text = ""
            
         End If
         
      Next
   End If
   
   nAux = ObjGrd.Row
   If IsMissing(Linha) Then Linha = nAux

   If Linha <= ObjGrd.FixedRows - 1 Then Exit Sub
   For i = 0 To ObjGrd.Cols - 1
      ObjGrd.TextMatrix(Linha, i) = ""
   Next

End Sub
'************************************************************************************
'   Função  :  Limpa o Grd, inclusive as fixedrows.
'
'   Entrada : mfgGrid  - Grd
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Public Sub LimpaFG(mfgGrid As Object)
    
'Área de declarações
Dim nLaco As Byte

'Área de manipulação de dados
    If mfgGrid.FixedRows > 0 Then
        mfgGrid.Rows = mfgGrid.FixedRows
        mfgGrid.Rows = mfgGrid.FixedRows + 1
        mfgGrid.Row = mfgGrid.FixedRows
    Else
        mfgGrid.Row = 1
    End If
    For nLaco = 0 To mfgGrid.Cols - 1 Step 1
        mfgGrid.Col = nLaco
        Set mfgGrid.CellPicture = LoadPicture("")
        mfgGrid.TextMatrix(1, nLaco) = ""
    Next nLaco
    If mfgGrid.RowHeightMin <> 0 Then
        mfgGrid.RowHeight(mfgGrid.FixedRows) = mfgGrid.RowHeightMin
    Else
        mfgGrid.RowHeight(mfgGrid.FixedRows) = mfgGrid.RowHeight(0)
    End If
End Sub
Public Sub RefreshMSGrid(ByRef DataControl As Object, ByRef pGrd As Object)
   Dim i%, j%, Max%, Fixed%, LinTop%
   Dim Lin As Long
   Dim PosCampo%, Pos&
   Dim Campo$, Letra$
   Dim Cab As Variant
   On Error GoTo Fim
'* Salvar número de coluna fixas e a linha corrente
   Fixed% = pGrd.FixedCols
   pGrd.FixedCols = 0
   Lin = pGrd.Row
   LinTop% = pGrd.TopRow

'* Montar Cabeçalho do Grid Existente
   ReDim Cab((4 * pGrd.Cols) - 1)
   For i = 0 To (pGrd.Cols - 1)
      j = i + 1
      Cab((j% * 4) - 4) = pGrd.TextMatrix(0, i)
      Cab((j% * 4) - 2) = pGrd.ColWidth(i) / 120
      Cab((j% * 4) - 1) = pGrd.ColAlignment(i)
   Next

'* Atualizar Dados Exibidos no Grid
   DataControl.Refresh
   For i = 0 To (pGrd.Cols - 1)
      j = i + 1
      Cab((j% * 4) - 3) = pGrd.TextMatrix(0, i)
   Next
   Max% = (UBound(Cab) - LBound(Cab) + 1) / 4
   
'* Esconder todas a colunas
   pGrd.Visible = False
   pGrd.Refresh
'* Mostrar e Formatar as colunas selecionadas
'   pGrd.Width = 260
   For j = 1 To Max%
      For i = 0 To pGrd.Cols - 1
         If pGrd.TextMatrix(0, i) = Cab((j% * 4) - 3) Then
            pGrd.TextMatrix(0, i) = Cab((j% * 4) - 4)      'Título
            pGrd.ColWidth(i) = Cab((j% * 4) - 2) * 120   'Tamanho
            pGrd.ColAlignment(i) = Cab((j% * 4) - 1)     'Alinhamento
'            If pGrd.Width >= 9440 Then pGrd.Width = 9440
'            pGrd.Width = pGrd.Width + pGrd.ColWidth(i)
            pGrd.Parent.Width = IIf(pGrd.Parent.Width > pGrd.Width, pGrd.Parent.Width, pGrd.Width + 240)
         End If
      Next
   Next
   If pGrd.Width >= 9295 Then
      pGrd.Left = 60
      pGrd.Width = 9440
   End If
'* Restaurar Cores do Cabeçalho
   pGrd.Row = 0
   pGrd.Col = Val(pGrd.Tag)
   If Not SetPctOrder(pGrd, CStr(pGrd.Tag)) Then
      pGrd.CellBackColor = vbBlue
      pGrd.CellForeColor = vbWhite
   End If
'* Salvar número de coluna fixas e a linha corrente
   pGrd.FixedCols = Fixed%
   If pGrd.Rows > LinTop% Then
      Lin = IIf(Lin = 0, 1, Lin)
      pGrd.Row = IIf(Lin >= pGrd.Rows, pGrd.Rows - 1, Lin)
      pGrd.TopRow = LinTop%
   End If
   Call SelRowMSGrid(pGrd, Lin)
   pGrd.Visible = True
   pGrd.Refresh
   pGrd.SetFocus
   Exit Sub
Fim:
   If Err <> 5 And Err <> 438 Then ClsMsg.ShowError
End Sub
Public Sub MontarCabGrid(ByRef pGrd As Object, Cab As Variant, Optional ByVal GrdTam = 0, Optional Redimensiona As Boolean = True)
   Dim Lin     As Integer
   Dim Col     As Integer
   Dim Ord     As Integer
   Dim Pos     As Integer
   Dim wFactor As Single
   Dim hFactor As Single
   Dim MaxTamGrid As Long
   Dim sCtrl   As String
   Dim i       As Integer
   Dim nLeft   As Long
   
   
   Call ClsCtrl.GetParents(pGrd)
   MaxTamGrid = pGrd.Parent.Width - pGrd.Left - 120
   pGrd.Visible = False
   
   wFactor = (Screen.Width / Screen.TwipsPerPixelX) / 800
   hFactor = (Screen.Height / Screen.TwipsPerPixelY) / 600
   
   If Not Redimensiona Then
      wFactor = 1
      hFactor = 1
   End If
   
   If GrdTam = 0 Then
      pGrd.Width = (260 + 80) * wFactor
   Else
      pGrd.Width = (GrdTam + 80) * wFactor
   End If
   Ord% = IIf(VarType(Cab(1)) = vbString, 4, 3)
   Lin = 0
   pGrd.Cols = (UBound(Cab) + 1) / Ord
   For Col% = 0 To pGrd.Cols - 1
      Pos% = Ord * (Col% + 1)
      pGrd.TextMatrix(Lin%, Col%) = Cab(Pos% - Ord%)
      pGrd.ColWidth(Col%) = Cab(Pos% - 2) * 120 * wFactor  'Tamanho
      pGrd.ColAlignment(Col%) = Cab(Pos% - 1)     'Alinhamento
      If pGrd.Width >= MaxTamGrid * wFactor Then
         pGrd.Width = MaxTamGrid * wFactor
      End If
      If GrdTam = 0 Then
         pGrd.Width = pGrd.Width + pGrd.ColWidth(Col%) + 9
      End If
      pGrd.Rows = 18
   Next
   On Error Resume Next
   If GrdTam = 0 Then
      i = 2
      nLeft = pGrd.Left
      While ClsCtrl.GetTag(pGrd, "PARENT" & i) <> ""
         sCtrl = ClsCtrl.GetTag(pGrd, "PARENT" & i, "")
         pGrd.Parent.Controls(sCtrl).Width = IIf(pGrd.Parent.Controls(sCtrl).Width > pGrd.Width, pGrd.Parent.Controls(sCtrl).Width, pGrd.Width + 120)
         nLeft = nLeft + pGrd.Parent.Controls(sCtrl).Left
         i = i + 1
      Wend
      If nLeft >= 0 Then
         pGrd.Parent.Width = IIf(pGrd.Parent.Width > pGrd.Width + pGrd.Left, pGrd.Parent.Width, pGrd.Width + pGrd.Left + 240)
      Else
         pGrd.Parent.Width = IIf(pGrd.Parent.Width > pGrd.Width, pGrd.Parent.Width, pGrd.Width + 240)
      End If
   End If
   pGrd.Col = 0
   pGrd.Visible = True
End Sub
Public Function MontarMSGrid(ByRef DataControl As Object, ByRef pGrd As Object, CabArr As Variant, Sql As String, Optional ByVal GrdTam = 0, Optional Redimensiona As Boolean = True) As Boolean
   Dim i%, Cols%, Tam%, ColVal%, Fixed%
   
   Dim PosCampo As Integer
   Dim Pos As Long
   Dim Sql_Ori As String
   Dim Sql_Rest As String
   Dim Campo As String
   Dim Letra As String
   Dim Aux As String
   Dim wFactor As Single
   Dim hFactor As Single
   Dim lEOF As Boolean
   Dim vLinha As String
      
   Dim ColNomeCampo As Collection
   Dim ColNome  As Collection
   Dim ColTam As Collection
   Dim ColAlinha As Collection
   
   Dim nKey As String
   Dim vLinhaCab  As String
   
   Dim cAux As String
   
   Sql = UCase(Sql)
   
   If Trim(Sql) = "" Then Exit Function
   
   wFactor = (Screen.Width / Screen.TwipsPerPixelX) / 800
   hFactor = (Screen.Height / Screen.TwipsPerPixelY) / 600
   
   If Not Redimensiona Then
      wFactor = 1
      hFactor = 1
   End If

   
   On Error GoTo TrataErro
   
   pGrd.Visible = False
   pGrd.Redraw = True
   pGrd.Rows = 1
   
   If IsEmpty(CabArr) And UCase(TypeName(DataControl)) = "RECORDSET" Then
      ReDim CabArr(4 * (DataControl.Fields.Count - 1))
      For i = 0 To (DataControl.Fields.Count - 1) / 4
         CabArr((i * 4)) = DataControl.Fields(i).Name     '* Nome do Campo
         CabArr((i * 4) + 1) = DataControl.Fields(i).Name '* Título
         CabArr((i * 4) + 2) = 10 'DataControl.Fields(i).Lenght  '* Tamanho
         CabArr((i * 4) + 3) = 0                          '* Alinhamento
      Next
   End If
   
'***************************
'* Salvar número de coluna fixas
   Fixed% = pGrd.FixedCols
   If pGrd.Cols = 0 Then
      pGrd.Cols = 1
   End If
   pGrd.FixedCols = 0
   
   If IsEmpty(CabArr) Then
      Cols = 0
   Else
      Cols = (UBound(CabArr) - LBound(CabArr) + 1) / 4
   End If

   pGrd.Cols = Cols
   
   '***************************
   '* Mostrar e Formatar as colunas selecionadas
   pGrd.Width = (IIf(GrdTam = 0, 260, GrdTam) + 80) * wFactor
   
   Set ColNomeCampo = New Collection
   Set ColNome = New Collection
   Set ColTam = New Collection
   Set ColAlinha = New Collection
   
   For i = LBound(CabArr) To ((UBound(CabArr) - LBound(CabArr) + 1) / 4) - 1
      vLinhaCab = vLinhaCab & CabArr((i * 4)) & vbTab
      
      nKey = CabArr((i * 4) + 1)
      ColNomeCampo.Add nKey, nKey                       '* Nome do Campo
      ColNome.Add CabArr((i * 4)), nKey                    '* Título
      ColTam.Add CabArr((i * 4) + 2) * 120 * wFactor, nKey '* Tamanho
      ColAlinha.Add CabArr((i * 4) + 3), nKey              '* Alinhamento
      
      pGrd.TextMatrix(0, i) = CabArr((i * 4))                '* Título
      pGrd.ColWidth(i) = CabArr((i * 4) + 2) * 120 * wFactor '* Tamanho
      If UCase(TypeName(pGrd)) = "MSFLEXGRID" Then
         If CabArr((i * 4) + 3) = vbLeftJustify Then
            pGrd.ColAlignment(i) = flexAlignLeftCenter
         ElseIf CabArr((i * 4) + 3) = vbCenter Then
            pGrd.ColAlignment(i) = flexAlignCenterCenter
         ElseIf CabArr((i * 4) + 3) = vbRightJustify Then
            pGrd.ColAlignment(i) = flexAlignRightCenter
         End If
      Else
         pGrd.ColAlignment(i) = CabArr((i * 4) + 3)             '* Alinhamento
      End If
      
      If GrdTam = 0 Then
         pGrd.Width = pGrd.Width + pGrd.ColWidth(i)
      End If
   Next
   
   '********
   '* Ignora se tela pde ser redimensionada ou não
   On Error Resume Next
   pGrd.Parent.Width = IIf(pGrd.Parent.Width > pGrd.Width, pGrd.Parent.Width, pGrd.Width + 240)
   
   On Error GoTo TrataErro
   If pGrd.Width >= pGrd.Parent.Width * wFactor Then
      pGrd.Left = 60 * wFactor
      pGrd.Width = pGrd.Parent.Width * wFactor
   End If
   
'*   Call ClsCtrl.CentrarObj(pGrd.Parent, pGrd, "H", True)
   
   pGrd.Visible = True
   pGrd.Redraw = False

'***************************
'* Tratar Query para manter o índice
'* se o Grid já estiver ordenado.
   If pGrd.Tag <> "" And pGrd.Rows > 1 Then
      Sql_Ori$ = Sql
      Pos = InStr(Sql, "ORDER BY")

      If Pos <> 0 Then
         Pos = Pos + 8
         Campo$ = ""
         Letra = Mid(Sql, Pos, 1)
         While Letra = " "
            Pos = Pos + 1
            Letra = Mid(Sql, Pos, 1)
         Wend
         Do While (Letra <> " " And Letra <> ";")
            Campo$ = Campo + Letra
            Pos = Pos + 1
            Letra = Mid(Sql, Pos, 1)
            If Pos > Len(Sql) Then Exit Do
         Loop
      
      '***************************
      '* Define Coluna do Índice
         If Trim(str(Val(Campo))) <> Campo And Campo <> "0" Then
            i = 0
            Sql = Mid(Sql, 1, InStr(Sql, Campo))
            While InStr(Sql, ",") <> 0
               i = i + 1
               Sql = Mid(Sql, InStr(Sql, ",") + 1)
            Wend
            Campo = str(i + 1)
         End If
         If pGrd.Tag <> 0 Then
            Sql = Mid(Sql_Ori, 1, InStr(Sql_Ori, " ORDER BY")) + " ORDER BY " + CStr(IIf(pGrd.Tag < 1, 1, pGrd.Tag + 1))
         Else
            Sql = Sql_Ori
         End If
      End If
   End If

'***************************
'* Gerar Dados Exibidos no Grid
   Dim isODBC  As Boolean
   If UCase(TypeName(DataControl)) = "RECORDSET" Then
      On Error GoTo TrataADO
      isODBC = Not (DataControl.State = 1)
      If Not isODBC Then
         If DataControl.State = 1 Then
            DataControl.Close
         End If
         DataControl.Open
      End If
      If DataControl.State = 0 Then
         DataControl.Open
      End If
      DataControl.MoveFirst
      On Error Resume Next
      Do While Not DataControl.EOF
         vLinha = ""
         'For i = 0 To DataControl.Fields.Count - 1
         For i = LBound(CabArr) To ((UBound(CabArr) - LBound(CabArr) + 1) / 4) - 1
            'vLinha = vLinha & DataControl.Fields(i).Value & vbTab
            cAux = ""
            cAux = DataControl.Fields(ColNomeCampo(i + 1)).Value & ""
            If Err <> 0 Then
               cAux = DataControl.Fields(ColNomeCampo(i + 1)).Value & ""
               Err = 0
            End If
            If cAux = Chr(255) Then cAux = ""
            vLinha = vLinha & cAux & vbTab
            If UCase(TypeName(DataControl)) = "RECORDSET" Then
               If DataControl.State = 0 Then
                  Exit Do
               End If
            End If
         Next
         pGrd.AddItem vLinha
         DataControl.MoveNext
      Loop
      On Error GoTo TrataErro
      
      If pGrd.Rows < 2 Then pGrd.Rows = 2
      pGrd.FixedRows = 1
      
   ElseIf TypeName(DataControl) = "MSRDC" Then
      DataControl.Sql = Sql
      DataControl.Refresh
   ElseIf TypeName(DataControl) = "Data" Then
      DataControl.RecordSource = Sql
      DataControl.Refresh
   End If

'***************************
'* Esconder todas a colunas

   If UCase(TypeName(DataControl)) <> "RECORDSET" Then
      If pGrd.Cols > 0 Then '* Número de Colunas
         On Error Resume Next
         For i = 0 To pGrd.Cols - 1
            pGrd.TextMatrix(0, i) = ColNome(pGrd.TextMatrix(0, i))
            If Err <> 0 Then
              pGrd.ColWidth(i) = 0
            End If
         Next
         On Error GoTo TrataErro
      Else
         If TypeName(DataControl) = "MSRDC" Then
            For i = 0 To DataControl.Resultset.rdoColumns.Count - 1
               Tam = DataControl.Resultset.rdoColumns(i).SIZE
               Tam = IIf(DataControl.Resultset.rdoColumns(i).SIZE < Len(DataControl.Resultset.rdoColumns(i).Name), Len(DataControl.Resultset.rdoColumns(i).Name), Tam)
               Tam = IIf(Tam > 30, 30, Tam)
               pGrd.ColWidth(i) = Tam * 120 * wFactor
            Next
         ElseIf TypeName(DataControl) = "Data" Then
            'DataControl.RecordSource = Sql
         End If
      End If
   
   End If
   
   If UCase(TypeName(DataControl)) = "RECORDSET" Then
      lEOF = DataControl.EOF
   ElseIf TypeName(DataControl) = "MSRDC" Then
      lEOF = DataControl.Resultset.EOF
   ElseIf TypeName(DataControl) = "Data" Then
      lEOF = DataControl.Recordset.EOF
   End If
    
   If lEOF Then
      GoTo Saida
   End If
   
   Pos = InStr(Sql, "ORDER BY")
   pGrd.Visible = True
   If Pos <> 0 Then
      Aux = Mid(Sql, Pos)
      If InStr(Aux, ",") = 0 Then '* Só Existe um campo no ORDER BY
         Pos = Pos + 8
         Campo$ = ""
         Letra = Mid(Sql, Pos, 1)
         While Letra = " "
            Pos = Pos + 1
            Letra = Mid(Sql, Pos, 1)
         Wend
         Do While (Letra <> " " And Letra <> ";")
            Campo$ = Campo + Letra
            Pos = Pos + 1
            Letra = Mid(Sql, Pos, 1)
            If Pos > Len(Sql) Then Exit Do
         Loop
      '* Define Coluna do Índice
         If Trim(str(Val(Campo))) <> Campo And Campo <> "0" Then
            i = 0
            Sql = Mid(Sql, 1, InStr(Sql, Campo) - 1)
            While InStr(Sql, ",") <> 0
               i = i + 1
               Sql = Mid(Sql, InStr(Sql, ",") + 1)
            Wend
            Campo = str(i + 1)
         End If
         pGrd.Tag = Val(Campo) - 1
         If pGrd.Tag = CStr(-1) Then
            Sql = Mid(Sql, 1, InStr(Sql, Campo))
            pGrd.Tag = 0
            While InStr(Sql, ",") <> 0
               pGrd.Tag = pGrd.Tag + 1
               Sql = Mid(Sql, 1, InStr(Sql, ",") - 1)
            Wend
         End If
         pGrd.Col = IIf(Val(pGrd.Tag) >= pGrd.Cols, pGrd.Cols - 1, Val(pGrd.Tag))
      End If
   End If

TrataErro:
   MontarMSGrid = False
   ClsMsg.ShowError
   GoTo Saida
TrataGrd:
   Resume Next
TrataADO:
   isODBC = True
   Resume Next

Saida:
   While Fixed >= pGrd.Cols
      Fixed = Fixed% - 1
   Wend

   If pGrd.Rows = 0 Then pGrd.Rows = 1
   
   Set ColNomeCampo = Nothing
   Set ColNome = Nothing
   Set ColTam = Nothing
   Set ColAlinha = Nothing
   
   '***************
   '* Configurar Picture de Ordenação no CabeçaLho do Grid
   '***************
   If Fixed% > 0 Then
      pGrd.FixedCols = Fixed%
   End If
   
   Call SetPctOrder(pGrd, Sql)
   
   pGrd.Visible = True
   pGrd.Redraw = True
   pGrd.Refresh
   MontarMSGrid = True

End Function
Public Sub OrdenarGrid(pGrd As Object, Optional pCol)
   Dim DateAux             As Date
   Dim i                   As Long
   Dim Pos                 As Long
   Dim iCol                As Long
   Dim iRow                As Long
   Dim nKeyCode            As Long
   Dim bSortAscending      As Boolean
   
   Static lSortCol         As Long
   Static lSortAscending   As Boolean
      
   If lSortCol = 0 Then lSortCol = 1
   
   With pGrd
      If IsMissing(pCol) Then pCol = .MouseCol
      iCol = pCol
      iRow = 1
      
      If lSortCol = iCol Then
         lSortAscending = Not lSortAscending
      Else
         lSortAscending = True
      End If
      lSortCol = iCol
      
      .Row = .FixedRows
      .Col = iCol
      .RowSel = .Rows - .Row
      
      If Trim(.TextMatrix(iRow, iCol)) = "" Then
         .Sort = 2 + lSortAscending '*1= flexSortGenericAscending
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
            .Sort = 4 + lSortAscending '* 3=flexSortNumericAscending
         Else
            .Sort = 8 + lSortAscending '*7= flexSortStringAscending
         End If
      End If
      
      Call SetPctOrder(pGrd, CStr(lSortCol), lSortAscending)
      
      If .Rows > 1 Then .Row = 1
   
      Call SelRowMSGrid(pGrd, .Row)
      .Redraw = True
   End With
End Sub
Public Function SetPctOrder(pGrd As Object, Sql As String, Optional OrderAscending) As Boolean
   Dim Pos As Integer
   Dim ObjImg As IPictureDisp
   Dim lSortCol As Integer
   Dim i As Integer
   
   On Error Resume Next
   Sql = UCase(Sql)
   
   If IsMissing(OrderAscending) Then
      OrderAscending = (InStr(Sql, " DESC ") = 0)
   End If
   
   Pos = InStr(Sql, "ORDER BY")
   If Pos <> 0 Then
      Sql = Mid(Sql, Pos + 9)
      If OrderAscending Then
         Sql = Trim(Mid(Sql, 1, InStr(Sql, " DESC ") - 1))
      End If
      While InStr(Sql, ",") <> 0
         Sql = Trim(Mid(Sql, InStr(Sql, ",") + 1))
      Wend
   Else
      Sql = CStr(Val(Sql) + 1)
      If IsMissing(OrderAscending) Then
         OrderAscending = True
      End If
   End If
   
   If Val(Sql) <> 0 Then
     lSortCol = Val(Sql) - 1
   End If
   
   With pGrd
      Pos = .Row
      .Row = 0
      
      For i = 0 To pGrd.Cols - 1
         .Col = i
         Set .CellPicture = Nothing
      Next
      
      .Col = lSortCol
      
      Set ObjImg = LoadResPicture(IIf(OrderAscending, "ORDER_DOWN", "ORDER_UP"), vbResBitmap)
      'Call ClsCtrl.SetPicture(ObjImg, IIf(OrderAscending, "ORDER_UP", "ORDER_DOWN"))
      Set .CellPicture = ObjImg
      .CellPictureAlignment = 7
      SetPctOrder = True
   End With
   pGrd.Row = Pos
   Err = 0
End Function
Public Sub SetFontCell(iRow As Integer, iCol As Integer, Optional iColor, Optional iBold)
   Dim MyRow As Integer
   Dim MyCol As Integer
   Dim bModoAnterior As Boolean
   
   If RunColorLine Then Exit Sub
   RunColorLine = True
   bModoAnterior = mvarModoEdicao
   mvarModoEdicao = False
   MyRow = Grd.Row
   MyCol = Grd.Col
   
   Grd.Row = iRow
   Grd.Col = iCol
   If Not IsMissing(iColor) Then Grd.CellForeColor = iColor
   If Not IsMissing(iBold) Then Grd.CellFontBold = iBold
   
   Grd.Row = MyRow
   Grd.Col = MyCol
   mvarModoEdicao = bModoAnterior
   RunColorLine = False
End Sub
'************************************************************************************
'   Função  :  Seleciona uma linha do pGrd.
'
'   Entrada : mfgGrid  - pGrd
'             nLinha - llinha a ser selecionada
'
'   Saída   :
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Public Sub SelRowMSGrid(pGrd As Object, Lin As Long)
'Área de manipulação de dados
    If pGrd.Rows <= Lin Then Exit Sub
    pGrd.Row = Lin
    pGrd.Col = pGrd.FixedCols
    pGrd.ColSel = pGrd.Cols - 1
End Sub
Public Function MSGrdProcurar(ByRef pGrd As Object, ByRef Pnl As Object, Key) As String
'================================================================
'= Última Alteração : 10/07/98                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Criticar Texto de um "Panel" para executar um   **
'**            PesquisarDbGrid%                                **
'**                                                            **
'** Recebe:  pGrd     - pGrd de Pesquisa                         **
'**          Pnl     - Panel que contem a texto de pesquisa    **
'**          Chave   - Chave de pesquisa                       **
'**                                                            **
'** Retorna: Linha do Grid selecionada.                        **
'**                                                            **
'**                                                            **
'****************************************************************
   Dim Chave  As String
   Dim RowTop As Long
   Dim Coluna As Long
   On Error GoTo Fim
   
   If pGrd.Rows <= 1 Then Exit Function
   MSGrdProcurar = LTrim(Pnl + IIf(Key > 0, UCase$(Chr$(Key)), ""))
   If Len(MSGrdProcurar) > 0 Then
      Chave = MSGrdProcurar
      Coluna = Val(pGrd.Tag)
      RowTop = PesquisarMSGrid(pGrd, Chave, Coluna)
      pGrd.TopRow = RowTop
      Call SelRowMSGrid(pGrd, RowTop)
   End If
   Exit Function
Fim:
   ClsMsg.ShowError
End Function
Public Sub OrdenarMSGrid(DTC As Object, pGrd As Object, Ind%)
   Dim Pos%, Sql$, i%, Fixed%, Ord As Boolean
   Dim MyIndex As index, MyField As Field
   Dim Cab()
   Dim Rc1 As Object, Rc2 As Object

   On Error GoTo Fim
   Screen.MousePointer = vbHourglass
   If Ind < 0 Then Exit Sub
   If pGrd.Cols <= Ind Then Exit Sub
   pGrd.Tag = Ind%
'* Salvar Estrutura do Grid
   Fixed% = pGrd.FixedCols
   pGrd.FixedCols = 0
   ReDim Cab(pGrd.Cols - 1, 2)
   pGrd.Row = 0
   For i = 0 To pGrd.Cols - 1
      pGrd.Col = i
      Cab(i, 0) = pGrd                 'Título
      Cab(i, 1) = pGrd.ColWidth(i)     'Tamanho
      Cab(i, 2) = pGrd.ColAlignment(i) 'Alinhamento
   Next
   If Ind < 0 Then Exit Sub
   If True Then
      If TypeName(DTC) = "MSRDC" Then
         Sql = DTC.Sql
      Else
         Sql = DTC.RecordSource
      End If
      Pos = InStr(UCase(Sql), "ORDER BY")
      If Pos <> 0 Then
         Ord = (InStr(Mid(Sql, Pos), str(Ind + 1)) <> 0)
         Ord = (Ord And (InStr(Mid(Sql, Pos), " DESC ") = 0))
      Else
         Ord = False
         Pos = Len(Sql) + 1
      End If
      If Pos <> 0 Then
         Sql$ = Trim(Mid(Sql, 1, Pos - 1))
         Sql$ = Sql$ & " order by " & Trim(str(Ind + 1))
         Sql = Sql & IIf(Ord, " DESC ", "")
         If TypeName(DTC) = "MSRDC" Then
            DTC.Sql = Sql
         Else
            DTC.RecordSource = Sql$
         End If
         DTC.Refresh
      End If
   Else
      On Error Resume Next
      Set Rc1 = DTC.Recordset
      Rc1.Sort = DTC.Recordset.Fields(Ind).Name
      Set Rc2 = Rc1.OpenRecordset(Rc1.Type)
      Set DTC.Recordset = Rc2
   End If
   Set Rc1 = Nothing
   Set Rc2 = Nothing
    
   On Error GoTo Fim
   'Recuperar Estrutura do Grid
   pGrd.Row = 0
   For i = 0 To pGrd.Cols - 1
      pGrd.Col = i%
      pGrd = Cab(i, 0)                 'Título
      pGrd.ColWidth(i) = Cab(i, 1)     'Tamanho
      pGrd.ColAlignment(i) = Cab(i, 2) 'Alinhamento
   Next
   pGrd.FixedCols = Fixed%
   pGrd.Col = Ind%
   
   '***************
   '* Configurar Picture de Ordenação no CabeçaLho do Grid
   '***************
   If Not SetPctOrder(pGrd, Sql) Then
      pGrd.CellBackColor = vbBlue
      pGrd.CellForeColor = vbWhite
   End If
   On Error GoTo Fim
   Call SelRowMSGrid(pGrd, 1)
   Screen.MousePointer = vbDefault
   Exit Sub
Fim:
   ClsMsg.ShowError
   Screen.MousePointer = vbDefault
End Sub
Public Sub OrderGrid(pGrd As Object, Ind%)
   Exit Sub
Fim:
   ClsMsg.ShowError
   Screen.MousePointer = vbDefault
End Sub
Public Function PesquisarMSGrid(pGrd As Object, ByVal Chave As String, Coluna As Long, Optional MatchCase = False) As Long
'================================================================
'= Última Alteração : 13/01/98                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Pesquisar Item no pGrd                          **
'**                                                            **
'** Recebe:  pGrd    - pGrd de Pesquisa                        **
'**          chave$  - Chave de Pesquisa                       **
'**          coluna  - Coluna de Pesquisa                      **
'**                                                            **
'** Retorna: linha do pGrd                                     **
'**                                                            **
'****************************************************************
   Dim Min           As Long     '* Primeira linha do pGrd com dados
   Dim Max           As Long     '* Última linha do pGrd com dados
   Dim rowi          As Long
   Dim Lin           As Long     '* Linha inicial e atual do pGrd
   Dim Tam           As Long     '* Tamanho da chave
   Dim Vl$                       '* Valor da celula a ser comparada
   Dim isNum         As Boolean  '* Comparação numérica
   Dim Ordem         As Boolean  '* Ordem da Coluna True = Crescente
   Dim Ch_Vl         As Boolean
   Dim Ch_Vl2        As Boolean
   Dim TxtAux        As String
   
   If pGrd.Rows <= 1 Then Exit Function
   rowi = pGrd.Row
   
   Tam = Len(Chave$)
   If Tam = 0 Then
      PesquisarMSGrid = pGrd.TopRow
      Exit Function
   End If
   
   Min = 1
   Max = IIf(pGrd.Rows < 1, 0, pGrd.Rows - 1)
   If isNum Then
      Vl$ = CDbl(pGrd.TextMatrix(Max, Coluna))
   Else
      Vl$ = ClsDsr.String_Sem_Acento(Left$(pGrd.TextMatrix(Max, Coluna), Tam))
   End If
   
   If Not MatchCase Then
      Vl$ = LCase(Vl$)
      Chave$ = LCase(Chave)
      Chave$ = ClsDsr.String_Sem_Acento(Chave)
   End If
   '* Determina se a palavra inicai com caracter numérico
   TxtAux = Mid(pGrd.TextMatrix(Max, Coluna), 1, 1)
   If TxtAux <> "" Then
      isNum = ClsDsr.Between(Asc(TxtAux), vbKey0, vbKey9)
      TxtAux = Mid(pGrd.TextMatrix(1, Coluna), 1, 1)
      If TxtAux <> "" And isNum Then
         isNum = ClsDsr.Between(Asc(TxtAux), vbKey0, vbKey9)
      End If
   End If
   On Error Resume Next
   If isNum Then
      Ordem = (CDbl(pGrd.TextMatrix(Max, Coluna)) > CDbl(pGrd.TextMatrix(1, Coluna)))
   Else
      Ordem = (LCase(pGrd.TextMatrix(Max, Coluna)) > LCase(pGrd.TextMatrix(1, Coluna)))
   End If
   If Err <> 0 Then
      isNum = Not isNum
      If isNum Then
         Ordem = (CDbl(pGrd.TextMatrix(Max, Coluna)) > CDbl(pGrd.TextMatrix(1, Coluna)))
      Else
         Ordem = (LCase(pGrd.TextMatrix(Max, Coluna)) > LCase(pGrd.TextMatrix(1, Coluna)))
      End If
      Err = 0
   End If
   Chave = ClsDsr.EliminarString(Chave, Chr(8))
   If isNum Then
      Ch_Vl = IIf(Ordem, CDbl(Chave) > CDbl(Vl), CDbl(Chave) < CDbl(Vl))
   Else
      Ch_Vl = IIf(Ordem, Chave > Vl, Chave < Vl)
   End If
   If Err <> 0 Then
      isNum = Not isNum
      If isNum Then
         Ch_Vl = IIf(Ordem, CDbl(Chave) > CDbl(Vl), CDbl(Chave) < CDbl(Vl))
      Else
         Ch_Vl = IIf(Ordem, Chave > Vl, Chave < Vl)
      End If
      Err = 0
   End If
   

   If Ch_Vl Then
      Lin = Max
   Else
      If isNum Then
         Ch_Vl = IIf(Ordem, CDbl(Chave$) <= CDbl(pGrd.TextMatrix(Min, Coluna)), CDbl(Chave$) >= CDbl(pGrd.TextMatrix(Min, Coluna)))
      Else
         Ch_Vl = IIf(Ordem, Chave$ <= Left$(LCase(pGrd.TextMatrix(Min, Coluna)), Tam), Chave$ >= Left$(LCase(pGrd.TextMatrix(Min, Coluna)), Tam))
      End If

      If Ch_Vl Then
         Lin = Min
      Else
         Do While Min <= Max
            Lin = (Max + Min) / 2
            If isNum Then
               Vl$ = pGrd.TextMatrix(Lin, Coluna)
               Ch_Vl = IIf(Ordem, CDbl(Chave$) > CDbl(Vl$), CDbl(Chave$) < CDbl(Vl$))
               Ch_Vl2 = IIf(Ordem, CDbl(Chave$) < CDbl(Vl$), CDbl(Chave$) > CDbl(Vl$))
            Else
               Vl$ = ClsDsr.String_Sem_Acento(Left$(pGrd.TextMatrix(Lin, Coluna), Tam))
               If Not MatchCase Then Vl$ = LCase(Vl$)
               Ch_Vl = IIf(Ordem, Chave$ > Vl$, Chave$ < Vl$)
               Ch_Vl2 = IIf(Ordem, Chave$ < Vl$, Chave$ > Vl$)
            End If
            If Ch_Vl Then
               Min = Lin + 1
            ElseIf Ch_Vl2 Then
               Max = Lin - 1
            Else
               Do While Chave$ = Vl$
                  Lin = Lin - 1
                  Vl$ = ClsDsr.String_Sem_Acento(Left$(pGrd.TextMatrix(Lin, Coluna), Tam))
                  If Not MatchCase Then Vl$ = LCase(Vl$)
               Loop
               Lin = Lin + 1
               Exit Do
            End If
         Loop
      End If
   End If
   If Min > Max Then
      Lin = Max   'rowi
   End If
   
   Dim MAX_LINHAS_GRID
   MAX_LINHAS_GRID = 0
   
   If Lin >= pGrd.TopRow + MAX_LINHAS_GRID - 2 Then
      If Lin >= pGrd.Rows - MAX_LINHAS_GRID + 2 Then
         rowi = pGrd.Rows - MAX_LINHAS_GRID - 1
         rowi = IIf(rowi < 0, 1, rowi)
         pGrd.TopRow = rowi
      Else
         pGrd.TopRow = Lin
      End If
   ElseIf Lin < pGrd.TopRow Then
      pGrd.TopRow = IIf(Lin < 1, 1, Lin)
   End If
   
   PesquisarMSGrid = IIf(Lin < 1, 0, Lin)
End Function
Public Sub PintarLinha(Optional iLin, Optional iColor)
   Dim i As Integer
   Dim bAux As Boolean
   Dim bAux2 As Boolean
   Dim iColAntes As Integer
   
   If RunColorLine Then Exit Sub
   If Not mvarModoEdicao Then Exit Sub
   If Grd.Rows <= Grd.FixedRows Then Exit Sub
   If Not Grd.Visible Then Exit Sub
   If IsMissing(iLin) Then iLin = Grd.Row
   If IsMissing(iColor) Then iColor = Grd.BackColorSel
   
   
   bAux = RunColorLine
   bAux2 = Grd.Redraw
   
   Grd.Redraw = False
   RunColorLine = True
   
   If iLin = 0 Then GoTo Saida
   iColAntes = Grd.Col
   
   For i = Grd.FixedCols To Grd.Cols - 1
      iLin = IIf(iLin > Grd.Rows - 1, Grd.Rows - 1, iLin)
      i = IIf(i > Grd.Cols - 1, Grd.Cols - 1, i)
      Grd.Row = iLin
      Grd.Col = i
      Grd.CellBackColor = iColor
      If Grd.CellBackColor <> iColor Then Exit For
   Next
Saida:
   Grd.Col = iColAntes
   
   RunColorLine = bAux
   Grd.Redraw = bAux2
End Sub

'************************************************************************************
'   Função  : Verifica se existe uma linha selecionada no Grd.
'
'   Entrada : mfgGrid  - Grd
'             cValorOk - Texto que indica se a linha está selecionada
'
'   Saída   : True/False
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
Public Function ProcuraLinhaSelecionadaFG(mfgGrid As Object, _
    cValorOk As String) As Boolean
    
'Área de declarações
Dim nLaco As Byte

'Área de manipulação de dados
    ProcuraLinhaSelecionadaFG = False
    For nLaco = mfgGrid.FixedRows _
        To mfgGrid.Rows - 1 Step 1
        If UCase(Trim(mfgGrid.TextMatrix(nLaco, mfgGrid.Cols - 1))) = UCase(Trim(cValorOk)) Then
            ProcuraLinhaSelecionadaFG = True
            Exit For
        End If
    Next nLaco

End Function
'***************************************
'***************************************
'**                                   **
'** Seção Referente ao Grid de Edição **
'**                                   **
'***************************************
'***************************************
Public Sub LeaveCell()
   Dim n, i%
   Dim bAux As Boolean
   Dim bAux2 As Boolean
   Dim isFixedRow As Boolean
   
'   If Not Grd.Visible Then Exit Sub
   If RunColorLine Then Exit Sub
   If Not mvarModoEdicao Then Exit Sub
   
   With Grd
      Call ClsCtrl.SetTag(Grd, "LINHA_ANTES", CStr(.Row))
      
'      If .Row < .FixedRows Then Exit Sub
      
      i = 0
      For Each n In mvarCollLin
         If UCase(TypeName(n)) <> "STRING" Then
            isFixedRow = (.Row < .FixedRows)
            n.Visible = False

            Select Case TypeName(n)
               Case "CheckBox"
                  'DoEvents
'            Grd.Redraw = True
'            n.Visible = True
'            n.Left = n.Left - 120
                  If .Row > .FixedRows - 1 Then
                     bAux = mvarModoEdicao
                     bAux2 = RunColorLine
                     mvarModoEdicao = False
                     RunColorLine = True
                     .Col = i
                     RunColorLine = bAux2
                     mvarModoEdicao = bAux
                     'bSel = (ClsCtrl.GetTag(Grd, "CEL(" & CStr(.Row) & "," & CStr(i) & ")") = "True")
                     Set .CellPicture = LoadResPicture(IIf(n.Value = vbChecked, "CHECKED", "UNCHECKED"), vbResBitmap)
                     'Set .CellPicture = LoadResPicture(IIf(bSel, "CHECKED", "UNCHECKED"), vbResBitmap)
                     Call ClsCtrl.SetTag(Grd, "CEL(" & CStr(.Row) & "," & CStr(i) & ")", IIf(n.Value = vbChecked, "True", "False"))
                     .CellPictureAlignment = flexAlignCenterCenter ' flexAlignCenterCenter
                     'DoEvents
                  End If
               Case "ListBox"
                  If n.ListIndex = -1 Then
                     .TextMatrix(.Row, i) = ""
                  Else
                     If n.Style = 1 Then
                        If n.Selected(n.ListIndex) Then
                           .TextMatrix(.Row, i) = n.Text
                        Else
                           .TextMatrix(.Row, i) = ""
                        End If
                     ElseIf n.Style = 0 Then
                        .TextMatrix(.Row, i) = n.Text
                     End If
                  End If
               Case "MaskEdBox"
                  If Not isFixedRow Then
                     .TextMatrix(.Row, i) = n.Text
                  End If
               Case "RichTextBox"
                  If Not isFixedRow Then
                     .TextMatrix(.Row, i) = n.Text
                  End If
               
               Case "SSCommand", "CommandButton"
                  If .Row > 0 Then .TextMatrix(.Row, i) = "..."
               Case Else
                  If Not isFixedRow Then
                     .TextMatrix(.Row, IIf(i >= .Cols, .Cols - 1, i)) = n
                  End If
            End Select
         End If
         i = i + 1
      Next
      Call PintarLinha(.Row, .BackColor)
   End With
   
End Sub
Public Sub EditaLinha(iLinha As Integer, Optional bLeaveCell = True)
   Dim n As Variant
   DoEvents
   If Grd Is Nothing Then Exit Sub
   If mvarEditandoLinha Then Exit Sub
   If RunColorLine Then Exit Sub
   If Not mvarModoEdicao Then Exit Sub
  
   mvarEditandoLinha = True
   Call PintarLinha
   If bLeaveCell Then
      LeaveCell
   End If
   Call Me.EnterCell
   Call MoverLinha
   
   If Grd.Row <> iLinha Then
      Grd.Row = 0
      If Grd.Rows > iLinha Then
         Grd.Row = iLinha
      End If
   End If
   mvarEditandoLinha = False
End Sub
Public Sub EnterCell()
   Dim n       As Variant
   Dim i       As Integer
   Dim j       As Integer
   Dim bAux    As Boolean
   Dim bAux2   As Boolean
   Dim MyObj   As Object

' Grid ser populada invisivel (CheckBox)
'   If Not Grd.Visible Then Exit Sub
   If RunColorLine Then Exit Sub
   If Not mvarModoEdicao Then Exit Sub
   
   With Grd
     If .Row < .FixedRows Then Exit Sub
      For Each n In mvarCollLin
         'DoEvents
         Select Case TypeName(n)
         
            Case "CheckBox"
               If .Row > .FixedRows - 1 Then
                  bAux = mvarModoEdicao
                  bAux2 = RunColorLine
                  RunColorLine = True
                  mvarModoEdicao = False
                  .Col = i
                  RunColorLine = bAux2
                  mvarModoEdicao = bAux
                  Set .CellPicture = LoadResPicture(IIf(n.Value = vbChecked, "CHECKED", "UNCHECKED"), vbResBitmap)
                  'If n.Visible Then
                    ' n.Top = .CellTop + ((.CellHeight - n.Height) / 2) + 140
                  'End If
                  n.Value = IIf((ClsCtrl.GetTag(Grd, "CEL(" & CStr(.Row) & "," & CStr(i) & ")") = "True"), vbChecked, vbUnchecked)
                  Call ClsCtrl.SetTag(Grd, "CEL(" & CStr(.Row) & "," & CStr(i) & ")", IIf(n.Value = vbChecked, "True", "False"))
                  .CellPictureAlignment = flexAlignCenterCenter ' flexAlignCenterCenter
               End If
            Case "ComboBox"
               If n <> .TextMatrix(.Row, i) Then
                  Call ClsCtrl.LocalizarCombo(n, .TextMatrix(.Row, i))
               End If
            
            Case "ListBox"
               If n.Style = 1 Then
                  'For j = 0 To n.ListCount - 1
                  '   n.Selected(j) = False
                  'Next
               End If
               If .TextMatrix(.Row, i) = "" Then
                  n.ListIndex = -1
               Else
                  Set MyObj = n
                  If -1 <> ClsCtrl.LocalizarList(MyObj, .TextMatrix(.Row, i)) Then
                     If n.Style = 1 Then
                        n.Selected(n.ListIndex) = True
                     End If
                  End If
               End If
            Case "MaskEdBox"
              ' n.FormattedText = .TextMatrix(.Row, i)
              n.Text = .TextMatrix(.Row, i)
            Case "SSCommand", "CommandButton"
            Case "String"
            Case Else
               n.Text = .TextMatrix(.Row, i)
         End Select
         i = i + 1
      Next
   End With
   Call MoverLinha
End Sub
Public Sub Set_Linha_Item()
   Dim n, LL%, tt%, i%
   
   With Grd
      LL = .Left + 40
      tt = .Top + .RowHeight(i)
      For Each n In mvarCollLin
         If UCase(TypeName(n)) = "STRING" Then
            '* célula sem Objeto
         ElseIf TypeName(n) = "CheckBox" Then
            n.Move LL + 50, tt, .ColWidth(i) - IIf(i = 0, 20, 0)
         Else
            n.Move LL, tt, .ColWidth(i) - IIf(i = 0, 20, 0)
         End If
         LL = LL + .ColWidth(i)
         i = i + 1
      Next
      .Height = mvarMaxLin * .RowHeight(0) + (mvarMaxLin * 5)
      
   End With
End Sub
Public Sub MoverLinha(Optional isScroll = False)
   Dim bVisible As Boolean
   Dim ExisteScrool_H As Boolean
   Dim ExisteScrool_V As Boolean
   Dim TotWidth As Long
   Dim i As Integer, n As Variant
   
   Dim bAux As Boolean
   Dim bAux2 As Boolean
   Dim bAuxModoExib As Boolean
   Dim MergeAux As Integer
   
   Dim nAux As Integer
   Dim bAuxRedraw As Boolean
   
   Dim nCellTop As Long
   Dim nCellLeft As Long
   Dim nCellHeight As Long
   Dim nCellWidth As Long
         
   Dim MergeAntes As Integer
   Dim RedrawAntes  As Boolean
   
   On Error Resume Next
   
   '* Formatar os Itens do Grid
   bAuxModoExib = mvarModoEdicao
   RedrawAntes = Grd.Redraw
      
   If Not Grd.Visible Then GoTo Saida
   If Not Grd.Redraw Then GoTo Saida
   If RunTopRow Then GoTo Saida
   
'   Dim MyGrd As MSFlexGrid
'   Set MyGrd = Grd
   
   With Grd
      .Parent.ScaleMode = 1
      
      '* Habilitado para moverlinha corretamente
      MergeAntes = .MergeCells
'       MergeAux = .MergeCells
       .MergeCells = flexMergeNever
'      If .Row < .FixedRows Then GoTo Saida
      i = 0
      
      '* Esconder todos os Itens do Grid
      For Each n In mvarCollLin
         If UCase(TypeName(n)) <> "STRING" Then
            n.Visible = False
            n.ZOrder 0
         End If
      Next
      If Not mvarModoEdicao Then GoTo Saida
      If .Row = 0 Then GoTo Saida
      
      ExisteScrool_H = Not (.RowIsVisible(.Rows - 1) And .ColIsVisible(.FixedCols))
      ExisteScrool_V = Not (.ColIsVisible(.Cols - 1) And .ColIsVisible(.FixedRows))
     '* Formatar os Itens do Grid
      mvarModoEdicao = False
      
'      bAuxRedraw = .Redraw
'      .Redraw = False
'      nAux = .Height
'      .Height = .RowHeight(.Row) * (.Rows + 1)
'      nCellTop = .CellTop
'      nCellLeft = .CellLeft
'      nCellHeight = .CellHeight
'      nCellWidth = .CellWidth
'      .Height = nAux
'      .Redraw = bAuxRedraw
      
      For Each n In mvarCollLin
         bVisible = True
         If UCase(TypeName(n)) <> "STRING" Then
            If TypeName(n) = "CheckBox" Then
               n.Height = 200
               'If Not RunTopRow Then
               If .RowIsVisible(.Row) Then
                  bAuxRedraw = .Redraw
                  .Redraw = False
                  nAux = .Height
                  .Visible = False
                  .Height = .RowHeight(.Row) * (mvarMaxLin + 1)
                  If (.Top + .CellTop + ((.CellHeight - n.Height) / 2) + 10) < .Top Then
                     n.Top = .Top + .CellHeight
                  ElseIf (.Top + .CellTop + ((.CellHeight - n.Height) / 2) + 10) > .Top + .Height Then
                     n.Top = .Top + .Height - .CellHeight
                  Else
                     n.Top = .Top + .CellTop + ((.CellHeight - n.Height) / 2) + 10
                  End If
                  
                  .Height = nAux
                  .Visible = True
                  .Redraw = bAuxRedraw
               End If
            ElseIf TypeName(n) = "ComboBox" Then
               If (.Top + .RowPos(.Row) + 40) < .Top Then
                  n.Top = .Top + .CellHeight
               ElseIf (.Top + .RowPos(.Row) + 40) > .Top + .Height Then
                  n.Top = .Top + .Height - .CellHeight
               Else
                  n.Top = .Top + .RowPos(.Row) + 40
               End If
               If n.Style <> 2 Then
                  n.Height = .RowHeight(.Row)
               End If
            Else
               If (.Top + .RowPos(.Row) + 40) < .Top Then
                  n.Top = .Top + .CellHeight
               ElseIf (.Top + .RowPos(.Row) + 40) > .Top + .Height Then
                   n.Top = .Top + .Height - .CellHeight
               Else
                  n.Top = .Top + .RowPos(.Row) + 40
               End If

               n.Height = .RowHeight(.Row)
            End If
            
            If i < .Cols Then
               If .ColIsVisible(i) Then
                  If TypeName(n) = "CheckBox" Then
                     n.Width = 200
                     n.Height = 200
                     bAux = mvarModoEdicao
                     bAux2 = RunColorLine
                     RunColorLine = True
                     mvarModoEdicao = False
                     .Col = i
                     mvarModoEdicao = bAux
                     RunColorLine = bAux2
                     If .RowIsVisible(.Row) Then

                        bAuxRedraw = .Redraw
                        .Redraw = False
                        nAux = .Height
                        .Visible = False
                        .Height = .RowHeight(.Row) * (mvarMaxLin + 1)
                        'n.Left = .Left + .ColPos(i) + ((nCellWidth - n.Width) / 2) + 50
                        n.Left = .Left + .CellLeft + ((.CellWidth - n.Width) / 2)
                        .Height = nAux
                        .Visible = True
                        .Redraw = bAuxRedraw
                        
                     End If
                  Else
                     n.Left = .Left + .ColPos(i) + 20
                     n.Width = .ColWidth(i) + 10
                  End If
                  If TypeName(n) <> "CheckBox" Then
                     bVisible = (.Col >= .FixedCols)
                     bVisible = (i >= .FixedCols)
                  End If
               Else
                  bVisible = False
               End If
            End If
            
            '* Verificar se Item está dentro da área
            '* Horizontal de visualização do Grid
      
            If .RowIsVisible(.Row) And bVisible Then
               bVisible = (.Row >= .FixedRows)
               '* Existe Scrool Horizontal
               If bVisible And ExisteScrool_V Then
                  bVisible = (.RowPos(.Row + 1) < (.RowHeight(.Row) * mvarMaxLin) - 60)
               End If
               If bVisible Then
                  bVisible = ((n.Left - .Left) >= 0) And ((n.Left - .Left) < .Width)
               End If
            Else
               bVisible = False
               If Not .RowIsVisible(.Row) And Not isScroll Then
                  RunTopRow = True
                  If .RowPos(.Row) > .Height - 90 Then
                     .TopRow = .TopRow + 1
                     If (.Top + .RowPos(.Row) + 40) < .Top Then
                        n.Top = .Top + .CellHeight
                     ElseIf (.Top + .RowPos(.Row) + 40) > .Top + .Height Then
                        n.Top = .Top + .Height - .CellHeight
                     Else
                        n.Top = .Top + .RowPos(.Row) + 40
                     End If
                   Else
                     If .TopRow > .FixedRows Then
                        .TopRow = .TopRow - 1
                     ElseIf .Row = .Rows - 1 Then
                        If .RowPos(.Row) > (.Height - 40) - .RowHeight(.Row) Then
                           '.TopRow = .TopRow - 1
                        End If
                     End If
                     If (.Top + .RowPos(.Row) + 40) < .Top Then
                        n.Top = .Top + .CellHeight
                     ElseIf (.Top + .RowPos(.Row) + 40) > .Top + .Height Then
                        n.Top = .Top + .Height - .CellHeight
                     Else
                        n.Top = .Top + .RowPos(.Row) + 40
                     End If
                  End If
                  bVisible = True
                  RunTopRow = False
                  'DoEvents
               End If
            End If
            
            n.Visible = bVisible
            
            TotWidth = n.Left + n.Width - .Left
            If .Width < TotWidth Then
               n.Width = (.ColWidth(i) - (TotWidth - .Width)) + IIf(ExisteScrool_H, -60, 160)
            End If
         End If
         i = i + 1
      Next
      '* Tratar os Limites do Grid
'      If Not .RowIsVisible(.Row) Then
'        If .Row >= mvarMaxLin Then
'           .TopRow = .Row - (mvarMaxLin - 2)
'        ElseIf .Row < .TopRow Then
'           .TopRow = .Row
'        End If
'      End If
'      .MergeCells = MergeAux
   End With
'DoEvents
'End Sub
'Retorno:
'   Resume Next
Saida:
   Grd.Redraw = RedrawAntes
   Grd.MergeCells = MergeAntes
   mvarModoEdicao = bAuxModoExib
End Sub
Public Sub UpDownGrid(KeyCode As Integer, Shift As Integer)
   Dim MyControl As Object
   Dim i As Integer
   Dim Style_Aux As Integer
   Dim bRedraw As Boolean
      
   On Error Resume Next
   
   If Grd Is Nothing Then Exit Sub
   If RunColorLine Then Exit Sub
   
   bRedraw = Grd.Redraw
'   Grd.Redraw = False
   
   Set MyControl = Grd.Parent.ActiveControl
   
   If TypeName(MyControl) = "ComboBox" Then
      If MyControl.List(MyControl.ListCount - 1) = Chr(255) Then
         MyControl.Remove MyControl.ListCount - 1
      End If
      If Shift <> 2 And Not KeyCode = vbKeyReturn Then
         GoTo Saida
      End If
   End If
   
   With Grd
      Select Case KeyCode
         Case vbKeyUp
            If .Row > 1 Then
               .Row = .Row - 1
               DoEvents
               If MyControl.Visible And MyControl.Enabled Then
                  MyControl.SetFocus
               End If
            End If
            If TypeName(MyControl) = "ComboBox" Then
               If Shift = 2 Then
                  If MyControl.ListIndex <= MyControl.ListCount And KeyCode <> vbKeyControl Then
                     DoEvents
                     If MyControl.ListIndex = MyControl.ListCount - 1 Then
                        MyControl.AddItem Chr(255)
                     End If
                     MyControl.ListIndex = MyControl.ListIndex + 1
                  End If
               End If
            End If
            
         Case vbKeyDown
            If .Row < .Rows - 1 Then
               .Row = .Row + 1
               DoEvents
               If MyControl.Visible And MyControl.Enabled Then
                  MyControl.SetFocus
               End If
            ElseIf .Row = .Rows - 1 Then
               If .RowPos(.Row) > (.Height - 40) - .RowHeight(.Row) Then
                  .TopRow = .TopRow + 1
               End If
            End If
            If TypeName(MyControl) = "ComboBox" Then
               If Shift = 2 Then
                  If MyControl.ListIndex >= 0 And KeyCode <> vbKeyControl Then
                     MyControl.ListIndex = MyControl.ListIndex - 1
                  End If
               End If
            End If
            
         Case vbKeyReturn
            For i = mvarCollLin.Count To 1 Step -1
               If TypeName(mvarCollLin(i)) <> "String" Then
                  'GoTo Saida
                  Exit For
               End If
            Next
            If MyControl Is mvarCollLin(i) Then
               If .Row < .Rows - 1 Then
                  Call ClsCtrl.SetTag(Grd, "LINHA_ANTES", .Row)
                  .Row = .Row + 1
                  DoEvents
                  For i = 1 To mvarCollLin.Count
                     If UCase(TypeName(mvarCollLin(i))) <> "STRING" Then
                        If i > Grd.FixedCols And mvarCollLin(i).Enabled And mvarCollLin(i).Visible Then
                           mvarCollLin(i).SetFocus
                           GoTo Saida
                        End If
                     End If
                  Next
               End If
               
            End If
         Case vbKeyLeft
            If Shift Then
               For i = 1 To mvarCollLin.Count
                  If UCase(TypeName(mvarCollLin(i))) <> "STRING" Then
                     If MyControl Is mvarCollLin(i) Then
                        
                        If i <= Grd.FixedCols + 1 Then
                           If mvarCollLin(mvarCollLin.Count).Enabled And mvarCollLin(mvarCollLin.Count).Visible Then
                              DoEvents
                              mvarCollLin(mvarCollLin.Count).SetFocus
                           End If
                        Else
                           i = IIf(i < 2, mvarCollLin.Count + 1, i)
                           If mvarCollLin(i - 1).Enabled And mvarCollLin(i - 1).Visible Then
                              DoEvents
                              mvarCollLin(i - 1).SetFocus
                           End If
                        End If
                        GoTo Saida
                     End If
                  End If
               Next
            End If
         Case vbKeyRight
            If Shift Then
               For i = 1 To mvarCollLin.Count
                  If UCase(TypeName(mvarCollLin(i))) <> "STRING" Then
                     If MyControl Is mvarCollLin(i) Then
                        If i = mvarCollLin.Count Then
                           If mvarCollLin(Grd.FixedCols + 1).Enabled And mvarCollLin(Grd.FixedCols + 1).Visible Then
                              DoEvents
                              mvarCollLin(Grd.FixedCols + 1).SetFocus
                           End If
                        Else
                           If mvarCollLin(i + 1).Enabled And mvarCollLin(i + 1).Visible Then
                              DoEvents
                              mvarCollLin(i + 1).SetFocus
                           End If
                        End If
                        GoTo Saida
                     End If
                  End If
               Next
            End If
      End Select
   End With
   If TypeName(MyControl) = "ComboBox" Then
      If MyControl.List(MyControl.ListCount - 1) = Chr(255) Then
         MyControl.Remove MyControl.ListCount - 1
      End If
   End If
'   Grd.Redraw = bRedraw
   If ClsDsr.InArray(KeyCode, Array(vbKeyUp, vbKeyDown)) Then ', vbKeyReturn, vbKeyLeft, vbKeyRight)) Then
      Call PintarLinha
   End If
Saida:
'   Grd.Redraw = bRedraw
End Sub
Public Sub VisualizarLinha(Optional bVisible As Boolean = True)
   Dim n As Variant
   For Each n In mvarCollLin
      If UCase(TypeName(n)) <> "STRING" Then
         n.Visible = bVisible
      End If
   Next
End Sub
Public Sub HabilitarLinha(Optional bEnabled As Boolean = True)
   Dim n As Variant, i As Integer
   On Error Resume Next
   i = 0
'   For Each n In mvarCollLin
'      i = i + 1
'      If UCase(TypeName(n)) <> "STRING" Then
'         If i >= Grd.FixedCols Then
'            n.Enabled = (bEnabled And n.Enabled)
'            If ClsCtrl.GetTag(mvarCollLin(i), "BackColor") = "" Then
'               Call ClsCtrl.SetTag(mvarCollLin(i), "BackColor", n.BackColor)
'            End If
'            If bEnabled Then
'               n.BackColor = ClsCtrl.GetTag(mvarCollLin(i), "BackColor")
'            Else
'               n.BackColor = &H80000016
'            End If
'         Else
'            'n.Enabled = False
'            'n.BackColor = &H80000016
'         End If
'      End If
'   Next
   Err = 0
End Sub
Private Sub Class_Initialize()
   mvarModoEdicao = True
   mvarSorted = True
   mvarMinLin = 2
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
End Sub

Private Sub mvarCtrl_GotFocus()
   Call ClsDsr.SelecionarTexto(mvarCtrl)
End Sub

Private Sub mvarCtrl_LostFocus()
   Dim n As Variant
   For Each n In mvarCollLin
      If TypeName(n) <> "String" Then
         If n Is Grd.Parent.ActiveControl Then
            Set mvarCtrl = Nothing
            Set mvarCtrl = Grd.Parent.Controls(24)
            Set mvarCtrl = Grd.Parent.ActiveControl
         End If
      End If
   Next
   
End Sub

Private Sub mvarCtrl_ObjectEvent(Info As EventInfo)
   Dim i As Integer
   Select Case Info.Name
      Case "Change"
      Case "Click"
      Case "DblClick"
      Case "KeyDown"
         Call UpDownGrid(Info.EventParameters(0), Info.EventParameters(1))
      Case Else
  End Select

End Sub

Private Sub Grd_Click()
   Dim n As Variant
   Dim iCol As Integer, iRow As Integer
   Dim DateAux As Date, i As Integer
   Dim Pos As Integer
   
   iCol = Grd.MouseCol
   iRow = Grd.MouseRow
   If iRow >= Grd.FixedRows And iCol + 1 <= mvarCollLin.Count Then
      Select Case UCase(TypeName(mvarCollLin(iCol + 1)))
         Case "CHECKBOX"
            DoEvents
            Set n = mvarCollLin(iCol + 1)
            If ClsCtrl.GetTag(Grd, "CEL(" & CStr(iRow) & "," & CStr(iCol) & ")") = "True" Then
               Call ClsCtrl.SetTag(Grd, "CEL(" & CStr(iRow) & "," & CStr(iCol) & ")", "False")
               n.Value = vbUnchecked
            Else
               Call ClsCtrl.SetTag(Grd, "CEL(" & CStr(iRow) & "," & CStr(iCol) & ")", "True")
               n.Value = vbChecked
            End If
            Grd.CellPictureAlignment = flexAlignCenterCenter ' flexAlignCenterCenter
      End Select
   Else
      If mvarSorted Then
         mvarModoEdicao = False
         RunColorLine = True
      
         Grd.Redraw = False
      
         For i = Grd.FixedCols To Grd.Cols - 1
            If Grd.Row >= Grd.FixedRows Then
               Grd.Col = i
               Grd.CellBackColor = Grd.BackColor
               Grd.CellForeColor = Grd.ForeColor
            End If
         Next
         
         If mvarSortCol = iCol Then
            mvarSortAscending = Not mvarSortAscending
'            mcarSortAscending = Not (ClsCtrl.GetTag(GrdLov, "SortAsc" & CStr(iCol)) = "True")
         Else
            mvarSortAscending = True
         End If
         mvarSortCol = iCol
      
'         Call ClsCtrl.SetTag(GrdLov, "SortAsc" & CStr(mvarSortCol), IIf(mcarSortAscending, "True", "False"))
      
         iRow = Grd.Row
         If Grd.FixedRows <= Grd.Rows - 1 Then
            Grd.Row = Grd.FixedRows
         End If
         Grd.Col = iCol
         If Grd.Rows - Grd.Row <= Grd.Rows - 1 Then
            Grd.RowSel = Grd.Rows - Grd.Row
         End If
      
         If Trim(Grd.TextMatrix(iRow, iCol)) = "" Then
             Grd.Sort = 2 + mvarSortAscending '*1= flexSortGenericAscending
         Else
            n = Asc(Mid(Trim(Grd.TextMatrix(iRow, iCol)), 1, 1))
            On Error Resume Next
            i = 0
            Pos = InStr(Grd.TextMatrix(iRow, iCol), "/")
            If Pos = 2 Or Pos = 3 Then
               i = InStr(Pos + 1, Grd.TextMatrix(iRow, iCol), "/") - Pos
            End If
            DateAux = CDate(Grd.TextMatrix(iRow, iCol))
            If Err = 0 And (i = 3 Or i = 2) Then
               Grd.Sort = 9  '* flexSortCustom
            ElseIf Asc("0") <= n And n <= Asc("9") Then
               Grd.Sort = 4 + mvarSortAscending '* 3=flexSortNumericAscending
            Else
               Grd.Sort = 8 + mvarSortAscending '*7= flexSortStringAscending
            End If
         End If
         Grd.Row = 0
         mvarModoEdicao = True
         RunColorLine = False
         
         Grd.Redraw = True
         Call Me.EditaLinha(1, False)
      End If
   End If

End Sub

Private Sub Grd_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    On Error Resume Next
    Dim dtmRow1 As Date, dtmRow2 As Date
    
'    mcarSortAscending= (ClsCtrl.GetTag(Grd, "SortAsc" & CStr(mvarSortCol)) <> "False")
    
    With Grd
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

Private Sub Grd_EnterCell()
   Dim i As Integer
   If RunColorLine Then Exit Sub
'    DoEvents
   If Grd.Enabled And Grd.Visible Then
      If Grd.Row >= Grd.FixedRows And Grd.SelectionMode = flexSelectionByRow Then
         RunColorLine = True
         If SetColorLine Then
            For i = Grd.FixedCols To Grd.Cols - 1
               If CollLin.Count > i Then
                  If UCase(TypeName(CollLin(i + 1))) = "STRING" Then
                     Grd.Col = i
                     Grd.CellBackColor = Grd.BackColorSel
                     Grd.CellForeColor = Grd.ForeColorSel
                  End If
               End If
            Next
         Else
            i = Grd.FixedCols
            If CollLin.Count >= i + 1 Then
               If UCase(TypeName(CollLin(i + 1))) = "STRING" Or Not mvarModoEdicao Then
                  Grd.Col = i
                  Grd.CellBackColor = Grd.BackColorSel
                  If Grd.CellPicture = 0 Then
                     Grd.CellForeColor = Grd.ForeColorSel
                  End If
               End If
            End If
         End If
         RunColorLine = False
      End If
   End If
   If Not mvarModoEdicao Then Exit Sub
   
   If Grd.MousePointer <> vbDefault Then
      mvarModoEdicao = False
      If Val(ClsCtrl.GetTag(Grd, "LINHA_ANTES")) <> 0 Then
         If Val(ClsCtrl.GetTag(Grd, "LINHA_ANTES")) < Grd.Rows - 1 Then
            Grd.Row = Val(ClsCtrl.GetTag(Grd, "LINHA_ANTES"))
         End If
      End If
      mvarModoEdicao = True
      Exit Sub
   End If
   Call EnterCell

End Sub
Private Sub Grd_GotFocus()
   Dim n As Variant
   If Grd.MousePointer <> vbDefault Then Exit Sub
'   Call Me.EditaLinha(Grd.Row)
'   For Each n In mvarCollLin
'      If UCase(TypeName(n)) <> "STRING" Then
'         If n.Visible And n.Enabled Then
'            n.SetFocus
'            Exit For
'         End If
'      End If
'   Next
End Sub
Private Sub Grd_KeyUp(KeyCode As Integer, Shift As Integer)
   Call UpDownGrid(KeyCode, Shift)
End Sub
Private Sub Grd_LeaveCell()
   Dim i As Integer
   If RunColorLine Then
      'Grd.CellBackColor = Grd.BackColor
      'Grd.CellForeColor = Grd.ForeColor
      Exit Sub
   End If

   If Grd.Enabled And Grd.Visible Then
      If Grd.Row >= Grd.FixedRows And Grd.SelectionMode = flexSelectionByRow Then
         RunColorLine = True
         If SetColorLine Then
            For i = Grd.FixedCols To Grd.Cols - 1
              If CollLin.Count > i Then
                 If UCase(TypeName(CollLin(i + 1))) = "STRING" Then
                    Grd.Col = i
                    Grd.CellBackColor = Grd.BackColor
                    Grd.CellForeColor = Grd.ForeColor
                 End If
              End If
            Next
         Else
            i = Grd.FixedCols
'            If UCase(TypeName(CollLin(i + 1))) = "STRING" Or Not mvarModoEdicao Then
               Grd.Col = i
               Grd.CellBackColor = Grd.BackColor
               If Grd.CellPicture = 0 Then
                  Grd.CellForeColor = Grd.ForeColorSel
               End If
               
'            End If
         End If
         RunColorLine = False
      End If
   End If
   If Not mvarModoEdicao Then Exit Sub
   
   If Grd.MousePointer <> vbDefault Then Exit Sub
   Call LeaveCell
End Sub
Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Grd.MousePointer <> vbDefault Then
       If Val(ClsCtrl.GetTag(Grd, "LINHA_ANTES")) < Grd.Rows Then
          Grd.Row = Val(ClsCtrl.GetTag(Grd, "LINHA_ANTES"))
       End If
    End If
End Sub
Private Sub Grd_RowColChange()
   If RunColorLine Then Exit Sub
   If Not mvarModoEdicao Then Exit Sub
   If Grd.MousePointer <> vbDefault Then Exit Sub
      
   Call MoverLinha
   Call HabilitarLinha(True)
End Sub
Private Sub Grd_Scroll()
   Dim bAux As Boolean
   Dim n As Variant
   Dim i As Integer
   
   If Not mvarModoEdicao Then Exit Sub
''   bAux = Grd.Redraw
''   Grd.Redraw = False
''   RunColorLine = True

   For Each n In mvarCollLin
      If UCase(TypeName(n)) <> "STRING" And UCase(TypeName(n)) <> "CHECKBOX" Then
         n.Visible = False
''         Grd.Col = i
''         Grd.CellBackColor = Grd.BackColorSel
''         Grd.CellForeColor = Grd.ForeColorSel
      End If
      i = i + 1
   Next
'   Call MoverLinha
''   RunColorLine = False
''   Grd.Redraw = bAux
'   Grd.Redraw = False
'   RunTopRow = True
'   DoEvents
'    Call MoverLinha(True)
'   Grd.Redraw = bAux
'   RunTopRow = bAux
End Sub
