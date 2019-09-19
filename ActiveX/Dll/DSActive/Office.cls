VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Office"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Public ExcelSheet As Excel.Application
'Public FileXls    As Excel.Workbook
Public Function GridToExcel(Grd As Object, Optional Nome, Optional ByVal isVisible As Boolean = False, Optional ByRef pForm, Optional ByVal TopFlood, Optional ByVal ExcluiArq As Boolean = True, Optional ByVal NomeArq, Optional ByVal ExibeMsg As Boolean = True) As Boolean
'   Dim xlApp As Excel.Application
'   Dim xlBook As Excel.Workbook
'   Dim xlSheet As Excel.WorkSheet
   
   Dim xlApp      As Object
   Dim xlBook     As Object
   Dim xlSheet    As Object

   Dim i          As Integer
   Dim j          As Integer
   Dim k          As Integer
   Dim sMsg       As String
   Dim sCaption   As String
   
   Dim iLin       As Integer
   Dim iCol       As Integer
   
   Dim ProgBar    As New CProgBar32
   
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   
   '*
   Grd.Redraw = False
   iLin = Grd.Row
   iCol = Grd.Col
   
   If IsMissing(pForm) Then Set pForm = Grd.Parent
   ExibeFlood = Not IsMissing(pForm)

   If IsMissing(Nome) Then Nome = UCase(Grd.Name)
   If IsMissing(NomeArq) Then NomeArq = UCase(Nome)
   
   If ExibeFlood Then
      Set ProgBar.Parent = pForm
      ProgBar.Create 0, pForm.Height - 600, pForm.Width - 90, 15
      'ProgBar.Create Grd.Left, Grd.Top + Grd.Height, Grd.Width, 15
      ProgBar.MinValue = 1
      ProgBar.MaxValue = Grd.Rows
      sCaption = pForm.Caption
   End If
   
   If LCase(Right(NomeArq, 3)) <> "xls" Then
      If InStr(NomeArq, ".") <> 0 Then
         NomeArq = Mid(NomeArq, 1, InStr(NomeArq, ".") - 1)
      End If
      NomeArq = NomeArq & ".xls"
   End If
      
   If ExcluiArq Then
      Call ClsDos.Del(ClsDos.DrvTmp & NomeArq)
   Else
      ExcluiArq = Not ClsDos.FileExists(ClsDos.DrvTmp & NomeArq)
   End If
   
   Set xlApp = CreateObject("Excel.Application")
   If Not ExcluiArq Then
      On Error Resume Next
      Set xlBook = xlApp.Workbooks.Open(ClsDos.DrvTmp & NomeArq)
      Set xlSheet = xlBook.Worksheets(Nome)
      If Err = 0 Then
         Call ClsMsg.ExibirAviso("Planilha '" & Mid(NomeArq, 1, Len(NomeArq) - 4) & "' Já Existe.", ClsMsg.LoadMsg(1))
         GoTo Saida
      Else
        Set xlSheet = xlBook.Worksheets.Add(, , 1)
      End If
      On Error GoTo TrataErro
      xlSheet.Name = Left(Nome, 31)
   Else
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Worksheets(1)
      On Error Resume Next
      Set xlSheet = xlBook.Worksheets(Nome)
      If Err <> 0 Then
         Set xlSheet = xlBook.Worksheets(1)
         xlSheet.Name = Left(Nome, 31)
      End If
   End If

   For i = 0 To Grd.Rows - 1
      ProgBar.Value = i + 1
      Grd.Parent.Caption = sCaption & " [" & i & "/" & Grd.Rows - 1 & "]"
      Grd.Parent.Refresh
      
      k = 0
      For j = 0 To Grd.Cols - 1
         If Grd.ColWidth(j) > 0 Then
            k = k + 1
            If i = 0 Then
               xlSheet.Columns(k).ColumnWidth = Grd.ColWidth(j) / 120
            End If
            
            If IsNumeric(Grd.TextMatrix(i, j)) Then
               If ClsDsr.InArray(Grd.ColAlignment(j), Array(1, 2, 9)) Then
                  xlSheet.Cells(i + 1, k) = "'" & Grd.TextMatrix(i, j)
               Else
                  xlSheet.Cells(i + 1, k) = Grd.TextMatrix(i, j)
               End If
            Else
               If IsDate(Grd.TextMatrix(i, j)) Then
                  xlSheet.Cells(i + 1, k) = CDate(Grd.TextMatrix(i, j))
                  If xlSheet.Columns(k).ColumnWidth < 9 Then
                     xlSheet.Columns(k).ColumnWidth = 9
                  End If
               Else
                  xlSheet.Cells(i + 1, k) = Grd.TextMatrix(i, j)
               End If
            End If
            xlSheet.Cells(i + 1, k).Font.Name = Grd.Font.Name '* MS Sans Serif
            xlSheet.Cells(i + 1, k).Font.SIZE = Grd.Font.SIZE '* 8,25
            
            '***********
            '* Verificar reconhecimento da cor
            'xlApp.Visible = True
            'Grd.Redraw = True
            Grd.Row = i
            Grd.Col = j
            
            If Grd.CellAlignment = 0 Then
               If ClsDsr.InArray(Grd.ColAlignment(j), Array(1, 2, 9)) Then
                  xlSheet.Cells(i + 1, k).HorizontalAlignment = 2
               ElseIf ClsDsr.InArray(Grd.ColAlignment(j), Array(3, 4, 5)) Then
                  xlSheet.Cells(i + 1, k).HorizontalAlignment = 3
               ElseIf ClsDsr.InArray(Grd.ColAlignment(j), Array(6, 7, 8)) Then
                  xlSheet.Cells(i + 1, k).HorizontalAlignment = 4
               End If
            Else
               xlSheet.Cells(i + 1, k).HorizontalAlignment = 2
            End If
            If ClsDsr.xVal(Grd.CellBackColor) = 0 Then
               xlSheet.Cells(i + 1, k).Interior.Color = Grd.BackColor
            Else
               xlSheet.Cells(i + 1, k).Interior.Color = Grd.CellBackColor
            End If
            If ClsDsr.xVal(Grd.CellForeColor) = 0 Then
               xlSheet.Cells(i + 1, k).Font.Color = Grd.ForeColor
            Else
               xlSheet.Cells(i + 1, k).Font.Color = Grd.CellForeColor
            End If
                        
            If xlSheet.Cells(i + 1, k).Interior.Color = 0 Then xlSheet.Cells(i + 1, k).Interior.Color = vbWhite
            If xlSheet.Cells(i + 1, k).Font.Color = 0 Then xlSheet.Cells(i + 1, k).Font.Color = vbBlack
            
            If xlSheet.Cells(i + 1, k).Interior.Color = vbWhite Or ClsDsr.xVal(Grd.CellBackColor) = 0 Then
               If i < Grd.FixedRows Then
                  xlSheet.Cells(i + 1, k).Interior.ColorIndex = 15
                  xlSheet.Cells(i + 1, k).Font.Color = vbBlack
               End If
               If j < Grd.FixedCols Then
                  xlSheet.Cells(i + 1, k).Interior.ColorIndex = 15
                  xlSheet.Cells(i + 1, k).Font.Color = vbBlack
               End If
            End If
            
            '********
            '* Bordas
            If xlSheet.Cells(i + 1, k).Interior.ColorIndex = 15 Then
               xlSheet.Cells(i + 1, k).Borders(-4107).LineStyle = 1
               xlSheet.Cells(i + 1, k).Borders(-4107).Weight = -4138
               xlSheet.Cells(i + 1, k).Borders(-4107).ColorIndex = -4105
               xlSheet.Cells(i + 1, k).Borders(-4160).LineStyle = 1
               xlSheet.Cells(i + 1, k).Borders(-4160).Weight = -4138
               xlSheet.Cells(i + 1, k).Borders(-4160).ColorIndex = -4105
               xlSheet.Cells(i + 1, k).Borders(-4131).LineStyle = 1
               xlSheet.Cells(i + 1, k).Borders(-4131).Weight = -4138
               xlSheet.Cells(i + 1, k).Borders(-4131).ColorIndex = -4105
               xlSheet.Cells(i + 1, k).Borders(-4152).LineStyle = 1
               xlSheet.Cells(i + 1, k).Borders(-4152).Weight = -4138
               xlSheet.Cells(i + 1, k).Borders(-4152).ColorIndex = -4105
            End If
         End If
         DoEvents
      Next
      DoEvents
      If ClsCtrl.GetTag(Grd, "CANCEL") = "True" Then
         Call ClsCtrl.SetTag(Grd, "CANCEL", False)
         i = Grd.Rows - 1
      End If
   Next
   
   If ExibeFlood Then
      pForm.Caption = sCaption
      Call ProgBar.DestroyProgBar
   End If
   
   If isVisible Then xlApp.Visible = True
   On Error Resume Next
   If ExcluiArq Then
      Call ClsDos.Del(ClsDos.DrvTmp & NomeArq)
   End If
   
   If ClsDos.FileExists(ClsDos.DrvTmp & NomeArq) Then
      Call xlSheet.SaveAs(ClsDos.DrvTmp & NomeArq & "z")
      Call ClsDos.Del(ClsDos.DrvTmp & NomeArq)
      Call xlSheet.SaveAs(ClsDos.DrvTmp & NomeArq)
      Call ClsDos.Del(ClsDos.DrvTmp & NomeArq & "z")
   Else
      Call xlSheet.SaveAs(ClsDos.DrvTmp & NomeArq)
   End If
   
   If Err = 0 Then
      sMsg = "O Arquivo '" & ClsDos.DrvTmp & NomeArq & "'"
      sMsg = sMsg & " foi salvo com sucesso!!!"
      If isVisible Then
         If ExibeMsg Then
            Call ClsMsg.ExibirAviso(sMsg, ClsMsg.LoadMsg(1))
         End If
      Else
         If ExibeMsg Then
            sMsg = sMsg & vbNewLine & vbNewLine
            sMsg = sMsg & "Deseja Visualizá-lo?"
            If vbYes = ClsMsg.ExibirPergunta(sMsg, ClsMsg.LoadMsg(1)) Then
               isVisible = True
               xlApp.Visible = True
            End If
         End If
      End If
   End If

Saida:
   On Error Resume Next
   If Not isVisible Then
      xlBook.Close
      xlApp.Quit
   End If
   Set xlApp = Nothing
   Set xlBook = Nothing
   Set xlSheet = Nothing
   Set ProgBar = Nothing
TrataErro:
   If Err <> 1004 Then
      ClsMsg.ShowError
   End If
   
   Grd.Row = iLin
   Grd.Col = iCol
   Grd.Redraw = True
   Screen.MousePointer = vbDefault
End Function
Public Function RecordSetToExcel(ByRef Rs As Object, Optional ByVal Nome, Optional ByVal isVisible As Boolean = False, Optional ByRef pForm, Optional ByVal TopFlood, Optional ByVal ExcluiArq As Boolean = True, Optional ByVal NomeArq, Optional ByVal ExibeMsg As Boolean = True) As Boolean
   Dim xlApp         As Object
   Dim xlBook        As Object
   Dim xlSheet       As Object

   Dim i             As Integer
   Dim k             As Integer
   Dim sMsg          As String
      
   Dim ProgBar       As New CProgBar32
   Dim ExibeFlood    As Boolean
   Dim sCaption      As String
   Dim lTmpArq       As Boolean 'Indica se o arquivo será gravado no diretório tmp
   
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
      
   ExibeFlood = Not IsMissing(pForm)
      
   If IsMissing(Nome) Then Nome = "TAB"
   If IsMissing(NomeArq) Then NomeArq = Nome
   
   If InStr(1, NomeArq, "\") > 0 Then
      lTmpArq = False
   Else
      lTmpArq = True
   End If
   
   If ExibeFlood Then
      Set ProgBar.Parent = pForm
      ProgBar.Create 0, pForm.Height - 600, pForm.Width - 90, 15
      ProgBar.MinValue = 1
      ProgBar.MaxValue = Rs.RecordCount
      sCaption = pForm.Caption
   End If

   If LCase(Right(NomeArq, 3)) <> "xls" Then
      If InStr(NomeArq, ".") <> 0 Then
         NomeArq = Mid(NomeArq, 1, InStr(NomeArq, ".") - 1)
      End If
      NomeArq = NomeArq & ".xls"
   End If
   

   If ExcluiArq Then
      If lTmpArq Then
         Call ClsDos.Del(ClsDos.DrvTmp & NomeArq)
      Else
         Call ClsDos.Del((NomeArq))
      End If
   Else
      If lTmpArq Then
         ExcluiArq = Not ClsDos.FileExists(ClsDos.DrvTmp & NomeArq)
      Else
         ExcluiArq = Not ClsDos.FileExists(NomeArq)
      End If
   End If
   
   Set xlApp = CreateObject("Excel.Application")
   If Not ExcluiArq Then
      On Error Resume Next
      If lTmpArq Then
         Set xlBook = xlApp.Workbooks.Open(ClsDos.DrvTmp & NomeArq)
      Else
         Set xlBook = xlApp.Workbooks.Open(NomeArq)
      End If
      Set xlSheet = xlBook.Worksheets(Nome)
      If Err = 0 Then
         Call ClsMsg.ExibirAviso("Planilha '" & Mid(NomeArq, 1, Len(NomeArq) - 4) & "' Já Existe.", ClsMsg.LoadMsg(1))
         GoTo Saida
      Else
        Set xlSheet = xlBook.Worksheets.Add(, , 1)
      End If
      On Error GoTo TrataErro
      xlSheet.Name = Left(Nome, 31)
   Else
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Worksheets(1)
      On Error Resume Next
      Set xlSheet = xlBook.Worksheets(Nome)
      If Err <> 0 Then
         Set xlSheet = xlBook.Worksheets(1)
         xlSheet.Name = Left(Nome, 31)
      End If
      
   End If
   
   For i = 0 To Rs.Fields.Count - 1
      k = i + 1
      xlSheet.Columns(k).ColumnWidth = Len(Rs.Fields(i).Name) * 1.1
      xlSheet.Cells(1, k) = UCase(Rs.Fields(i).Name)
      xlSheet.Cells(1, k).Font.Name = "MS Sans Serif"
      xlSheet.Cells(1, k).Font.SIZE = 8.25
      xlSheet.Cells(1, k).Interior.ColorIndex = 15
      If xlSheet.Cells(1, k).Interior.ColorIndex = 15 Then
         xlSheet.Cells(1, k).Borders(-4107).LineStyle = 1
         xlSheet.Cells(1, k).Borders(-4107).Weight = -4138
         xlSheet.Cells(1, k).Borders(-4107).ColorIndex = -4105
         xlSheet.Cells(1, k).Borders(-4160).LineStyle = 1
         xlSheet.Cells(1, k).Borders(-4160).Weight = -4138
         xlSheet.Cells(1, k).Borders(-4160).ColorIndex = -4105
         xlSheet.Cells(1, k).Borders(-4131).LineStyle = 1
         xlSheet.Cells(1, k).Borders(-4131).Weight = -4138
         xlSheet.Cells(1, k).Borders(-4131).ColorIndex = -4105
         xlSheet.Cells(1, k).Borders(-4152).LineStyle = 1
         xlSheet.Cells(1, k).Borders(-4152).Weight = -4138
         xlSheet.Cells(1, k).Borders(-4152).ColorIndex = -4105
      End If
   Next
   
'   Rs.MoveFirst
'   While Not Rs.EOF
'      j = Rs.AbsolutePosition + 1
'      k = 0
'      For i = 0 To Rs.Fields.Count - 1
'         k = i + 1
'         xlSheet.Cells(j, k).Font.Name = "MS Sans Serif"
'         xlSheet.Cells(j, k).Font.SIZE = 8.25
'
'         nMax = Len(Rs.Fields(i).Value) * 1.1
'         nMax = IIf(nMax < xlSheet.Columns(k).ColumnWidth, xlSheet.Columns(k).ColumnWidth, nMax)
'         nMax = IIf(nMax > 50, 50, nMax)
'         xlSheet.Columns(k).ColumnWidth = nMax
'         If IsNumeric(Rs.Fields(i).Value) Then
'            If ClsDsr.InArray(Rs.Fields(i).Type, Array(adVarChar)) Then
'               xlSheet.Cells(j, k) = "'" & Rs.Fields(i).Value
'            Else
'               xlSheet.Cells(j, k) = Rs.Fields(i).Value
'            End If
'         Else
'            If IsDate(Rs.Fields(i).Value) Then
'               If ClsDsr.InArray(Rs.Fields(i).Value, Array(adDate, adDBDate, adDBTime, adDBTimeStamp)) Then
'                  xlSheet.Cells(j, k) = CDate(Rs.Fields(i).Value)
'               Else
'                  xlSheet.Cells(j, k) = Rs.Fields(i).Value
'               End If
'            Else
'               xlSheet.Cells(j, k) = Rs.Fields(i).Value
'            End If
'         End If
'         DoEvents
'      Next
'      DoEvents
'      If ExibeFlood Then
'         ProgBar.Value = Rs.AbsolutePosition
'         pForm.Caption = sCaption & " [" & Rs.AbsolutePosition & "/" & Rs.RecordCount & "]"
'         pForm.Refresh
'         If ClsCtrl.GetTag(pForm, "CANCEL") = "True" Then
'            Call ClsCtrl.SetTag(pForm, "CANCEL", False)
'            Rs.MoveLast
'         End If
'      End If
'      Rs.MoveNext
'   Wend
   
   Rs.MoveFirst
   xlSheet.range("A2").CopyFromRecordset Rs
   
   If ExibeFlood Then
      pForm.Caption = sCaption
      Call ProgBar.DestroyProgBar
   End If
   
   If isVisible Then xlApp.Visible = True
   On Error Resume Next
   If ExcluiArq Then
      If lTmpArq Then
         Call ClsDos.Del(ClsDos.DrvTmp & NomeArq)
      Else
         Call ClsDos.Del((NomeArq))
      End If
   End If
   If lTmpArq Then
      If ClsDos.FileExists(ClsDos.DrvTmp & NomeArq) Then
         Call xlSheet.SaveAs(ClsDos.DrvTmp & NomeArq & "z")
         Call ClsDos.Del(ClsDos.DrvTmp & NomeArq)
         Call xlSheet.SaveAs(ClsDos.DrvTmp & NomeArq)
         Call ClsDos.Del(ClsDos.DrvTmp & NomeArq & "z")
      Else
         Call xlSheet.SaveAs(ClsDos.DrvTmp & NomeArq)
      End If
   Else
      If ClsDos.FileExists(NomeArq) Then
         Call xlSheet.SaveAs(NomeArq & "z")
         Call ClsDos.Del((NomeArq))
         Call xlSheet.SaveAs(NomeArq)
         Call ClsDos.Del(NomeArq & "z")
      Else
         Call xlSheet.SaveAs(NomeArq)
      End If
   End If
   If Err = 0 Then
      If lTmpArq Then
         sMsg = "O Arquivo '" & ClsDos.DrvTmp & NomeArq & "'"
      Else
         sMsg = "O Arquivo '" & NomeArq & "'"
      End If
      sMsg = sMsg & " foi salvo com sucesso!!!"
      If isVisible Then
         If ExibeMsg Then
            Call ClsMsg.ExibirAviso(sMsg, ClsMsg.LoadMsg(1))
         End If
      Else
         If ExibeMsg Then
            sMsg = sMsg & vbNewLine & vbNewLine
            sMsg = sMsg & "Deseja Visualizá-lo?"
            If vbYes = ClsMsg.ExibirPergunta(sMsg, ClsMsg.LoadMsg(1)) Then
               isVisible = True
               xlApp.Visible = True
            End If
         End If
      End If
   End If
   
Saida:
   On Error Resume Next
   If Not isVisible Then
      xlBook.Close
      xlApp.Quit
   End If
   Set xlApp = Nothing
   Set xlBook = Nothing
   Set xlSheet = Nothing
   Set ProgBar = Nothing
   Screen.MousePointer = vbDefault
   Exit Function
TrataErro:
   If Err = 1004 Or Err = 9 Then
      Resume Next
   Else
      ClsMsg.ShowError
   End If
   GoTo Saida
End Function
'Public Sub SetContained(frm As Object)
'   Dim n
'   Dim m
'   On Error Resume Next
'   For Each n In frm.Controls
'      Call ClsCtrl.SetTag(n, "CONTAINED", n.Container.Name)
'   Next
'End Sub
