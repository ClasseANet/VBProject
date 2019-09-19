VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjReportControl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public Function AddReportRecord(Control As ReportControl, Parent As ReportRecord, Columns As Variant, Optional Icon, Optional HasCheckbox, Optional TreeColumn As Integer = 0, Optional GroupCaption) As ReportRecord
Public Function AddReportRecord(Control As Object, Parent As Object, Columns As Variant, Optional Icon, Optional HasCheckbox, Optional TreeColumn As Integer = 0, Optional GroupCaption) As Object
   Dim xRecord As Object 'ReportRecord
   Dim xItem   As Object 'ReportRecordItem
   Dim i       As Integer
   
   
   If Parent Is Nothing Then
      Set xRecord = Control.Records.Add
   Else
      Set xRecord = Parent.Childs.Add()
      Control.Columns(TreeColumn).TreeColumn = True
   End If
   
   Set xItem = xRecord.AddItem(Columns(0))
   If Not IsMissing(Icon) Then xItem.Icon = Icon
   If Not IsMissing(HasCheckbox) Then xItem.HasCheckbox = HasCheckbox
   If Not IsMissing(GroupCaption) Then xItem.GroupCaption = GroupCaption
   
   
   For i = 1 To Control.Columns.Count - 1 'UBound(Columns)
      If i <= UBound(Columns) Then
         xRecord.AddItem Columns(i)
      Else
         xRecord.AddItem ""
      End If
   Next
   'Set Item = Record.AddItem(Price)
   'Item.Format = "$ %s"
   
   Set AddReportRecord = xRecord
End Function
Private Sub ClipBoardSetData(MyString As String)
   Dim hGlobalMemory    As Long
   Dim lpGlobalMemory   As Long
   Dim hClipMemory      As Long
   Dim x                As Long

   ' Allocate moveable global memory.
   '-------------------------------------------
   hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

   ' Lock the block to get a far pointer
   ' to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)

   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      MsgBox "Could not unlock memory location. Copy aborted."
      GoTo OutOfHere2
   End If

   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      MsgBox "Could not open the Clipboard. Copy aborted."
      Exit Sub
   End If

   ' Clear the Clipboard.
   x = EmptyClipboard()

   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
   If CloseClipboard() = 0 Then
   MsgBox "Could not close Clipboard."
   End If
End Sub
'FillRCFromRS(ByRef pRecordSet As Object, ByRef pReportControl As Object, Optional bDoEvents As Boolean = False)  ' XtremeReportControl)
Private Sub FillRCFromExcel(ByRef pReportControl As Object, sFileXls As String, pSheet As String, Optional bDoEvents As Boolean = False)
   Dim iLin    As Double
   Dim iCol    As Double
   Dim iLinIni As Double
   Dim iLinMax As Double
   Dim iLinMin As Double
   Dim sValue  As String
   
   Dim sColIni   As String
   Dim nColFim   As Double
   Dim sColFim   As String
   
   Dim bLoop  As Boolean
   Dim nFator As Integer
   
   Dim nIndex     As Double
   Dim i          As Double
   Dim xRptCtrl   As Object   ' XtremeReportControl
   Dim xRecord    As Object   ' XtremeReportControl.ReportRecord
   Dim xItem      As Object   ' XtremeReportControl.ReportRecordItem
   
         
   'On Error GoTo ErrHandler
   Screen.MousePointer = vbHourglass
   If sFileXls = "" Then GoTo Saida
   
   '******************
   '* Recuperar nome da Planilha
   Dim oExcelApp             As Object 'Excel.Application
   Dim oFileXls              As Object 'Excel.Workbook
   Dim oSheet                As Object 'Excel.Worksheet
   
   Set oExcelApp = CreateObject("Excel.Application")
   Set oFileXls = oExcelApp.Workbooks.Open(sFileXls)
   Set oSheet = oFileXls.Worksheets(pSheet)
   
   '******************
   '* Popular Grid
   pReportControl.Records.DeleteAll
   pReportControl.Columns.DeleteAll
   pReportControl.Populate
   
   iLinIni = 1 'xVal(Me.TxtLinIni.Text)
   iLinIni = IIf(iLinIni = 0, 1, iLinIni)
   
   iCol = 1
   sValue = Trim(oSheet.Cells(iLinIni, iCol))
   While sValue <> ""
      pReportControl.Columns.Add iCol, sValue, 100, True
      iCol = iCol + 1
      sValue = Trim(oSheet.Cells(iLinIni, iCol))
   Wend
   
   iLinMax = oSheet.UsedRange.Rows.Count
   iLin = iLinMax
   iCol = 1
   iLinMin = 0
   nFator = 1
   bLoop = Trim(oSheet.Cells(iLin, iCol)) = ""
   While bLoop
      If nFator = 1 Then
         iLinMax = iLin
         iLin = iLin - (nFator * CInt((iLinMax - iLinMin) / 2))
      Else
         iLin = iLin - (nFator * CInt((iLinMax - iLin) / 2))
      End If
      
      
      bLoop = Not (Trim(oSheet.Cells(iLin, iCol)) <> "" And Trim(oSheet.Cells(iLin + 1, iCol)) = "")
      
      If Trim(oSheet.Cells(iLin + 1, iCol)) = "" Then
         nFator = 1
      Else
         nFator = -1
         iLinMin = iLin
      End If
      If iLin <= iLinIni Then
         iLin = iLinIni + 1
         bLoop = False
      End If
   Wend
   iLinMax = iLin
      
   nIndex = 1
   nColFim = pReportControl.Columns.Count
   sColFim = ""
   nIndex = Int(nColFim / 26)
   If nIndex > 0 Then
      If nColFim = 26 Then
         sColFim = sColFim & Chr(Asc("@") + nColFim)
      Else
         sColFim = sColFim & Chr(Asc("@") + nIndex)
      End If
   End If
   nIndex = nColFim Mod 26
   If nIndex > 0 Then
      sColFim = sColFim & Chr(Asc("@") + nIndex)
   End If
   If nColFim > 0 Then
      sColFim = sColFim & iLinMax
   End If
   
   sColIni = "A" & iLinIni + 1
   If sColFim = "" Then sColFim = sColIni
   If oSheet.Range(sColIni & ":" & sColFim).Copy Then
      pReportControl.Paste
      Clipboard.Clear
      If pReportControl.Rows.Count > 1 And pReportControl.Columns.Count > 1 Then
         For nIndex = pReportControl.Rows.Count - 1 To 0 Step -1
            If pReportControl.Records(nIndex).Item(1).Value = "" Then
               pReportControl.RemoveRecordEx pReportControl.Records(nIndex)
            Else
               Exit For
            End If
         Next
      End If
   End If
   
   pReportControl.Populate
   'Me.LblStatus.Caption = "Qtd. Linhas : " & pReportControl.Rows.Count
   GoTo Saida
ErrHandler:
  ClsMensagem.ExibirErro
Saida:
   oFileXls.Close
   oExcelApp.Quit
   Set oFileXls = Nothing
   Set oExcelApp = Nothing
   Set oSheet = Nothing
   
   Screen.MousePointer = vbDefault
End Sub
Public Sub FillRCFromRS(ByRef pRecordSet As Object, ByRef pReportControl As Object, Optional bDoEvents As Boolean = False, Optional ByRef pCollColumn As Collection, Optional pCurrency As Boolean = False)    ' XtremeReportControl)
   Dim xField     As Object   ' ADODB.Field
   Dim xRecord    As Object   ' XtremeReportControl.ReportRecord
   Dim xItem      As Object   ' XtremeReportControl.ReportRecordItem
   Dim nIndex     As Integer
   Dim nMouse     As Integer
   Dim nType      As Integer
   
   On Error GoTo TrataErro
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   pReportControl.Records.DeleteAll
   pReportControl.Populate
   
   pReportControl.Columns.DeleteAll
   pReportControl.Redraw
   Set pCollColumn = New Collection
   For Each xField In pRecordSet.Fields
      pReportControl.Columns.Add nIndex, xField.Name, 100, True
      pCollColumn.Add nIndex, xField.Name
      nIndex = nIndex + 1
   Next xField
      
Dim nMetodo As Integer
Dim nTime As Integer
Dim sFileXls As String
nMetodo = 1
   If nMetodo = 1 Then
      Do While Not pRecordSet.EOF
         Set xRecord = pReportControl.Records.Add()
         For nIndex = 0 To pRecordSet.Fields.Count - 1
            Set xItem = xRecord.AddItem("")
            With xItem
               If IsNull(pRecordSet.Collect(nIndex)) Then
                  .Value = Empty
               Else
                  If pCurrency Then
                     nType = pRecordSet.Fields(nIndex).Type
                     If nType = adCurrency Or nType = adDecimal Or nType = adDouble Or nType = adNumeric Then
                        .Value = pRecordSet.Collect(nIndex)
                        If .Value < 0 Then
                           .Caption = "(" & Mid(ClsGeneral.ValBr(pRecordSet.Collect(nIndex)), 2) & ")"
                           .ForeColor = vbRed
                        Else
                           .Caption = ClsGeneral.ValBr(pRecordSet.Collect(nIndex))
                        End If
                     Else
                        .Value = pRecordSet.Collect(nIndex)
                     End If
                  Else
                     .Value = pRecordSet.Collect(nIndex)
                  End If
               End If
            End With
         Next
         If bDoEvents Then DoEvents
         pRecordSet.MoveNext
      Loop
      pReportControl.Populate
   ElseIf nMetodo = 2 Then
      '* METODO APENAS PARA TESTE
      nTime = Second(Now)
      sFileXls = "C:\TMP\TMP.xls"
      Call ClsBanco.RecordSetToExcel(pRecordSet, Nome:="Plan1", NomeArq:=sFileXls, ExibeMsg:=False)
      Call FillRCFromExcel(pReportControl, sFileXls, "Plan1")
      Call ClsAmbiente.ExcluirArquivo(sFileXls)
      MsgBox Second(Now) - nTime
   End If
   Screen.MousePointer = nMouse
   Exit Sub
TrataErro:
   If Err = 457 Then
      Resume Next
   Else
      MsgBox Err & "-" & Error
   End If
End Sub
Public Function GetGrdColumnIndex(pGrd As Object, pCaption As String) As Integer
   Dim i As Integer
   On Error Resume Next
   With pGrd
       For i = 0 To .Columns.Count
           If UCase(Trim(.Columns.Column(i).Caption)) = UCase(Trim(pCaption)) Then
              GetGrdColumnIndex = i
              Exit For
           End If
       Next i
   End With
End Function
Public Sub OrdenarGrd(pReportControl As Object, pColChave As String, pColPai As String, pColTree As String)
   Dim i           As Integer
   Dim j           As Integer
   Dim nColun      As Integer
   Dim sCodPai     As String
   Dim iColunPai   As Long
   
   On Error Resume Next
   nColun = GetGrdColumnIndex(pReportControl, pColTree)
   pReportControl.Columns(nColun).TreeColumn = True
   pReportControl.Populate
   
   For i = pReportControl.Records.Count - 1 To 0 Step -1
       nColun = GetGrdColumnIndex(pReportControl, pColChave)
       sCodPai = pReportControl.Records.Record(i).Item(nColun).Value
       iColunPai = GetGrdColumnIndex(pReportControl, pColPai)
       
       For j = pReportControl.Records.Count - 1 To 0 Step -1
           If Trim(pReportControl.Records.Record(j).Item(iColunPai).Value & "") = Trim(sCodPai & "") Then
              pReportControl.Records(i).Childs.InsertAt pReportControl.Records(i).Childs.Count, pReportControl.Records.Record(j)
              pReportControl.Records.RemoveAt j
              pReportControl.Populate
              'Exit For
           End If
       Next j
       pReportControl.Records.Record(i).Expanded = True
       pReportControl.Populate
   Next i
End Sub
Private Sub RSToClipboard(pRecordSet As Object)
   Dim sLinha  As String
   Dim varRows As Variant
   Dim intCol  As Integer
   Dim lngRow  As Long

   With pRecordSet
      If Not (.EOF And .BOF) Then
         .MoveFirst
         sLinha = ""
         varRows = pRecordSet.GetRows
         For lngRow = 0 To UBound(varRows, 2)
            For intCol = 0 To UBound(varRows, 1)
               sLinha = sLinha & varRows(intCol, lngRow) & vbTab
            Next
            sLinha = sLinha & vbCrLf
         Next
      End If
   End With
   Call ClipBoardSetData(sLinha)
End Sub
