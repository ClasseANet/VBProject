Attribute VB_Name = "DSR"
Option Explicit
Private Const GWL_STYLE As Long = (-16&)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public Sub SaveParam(pCODPARAM As String, pVLPARAM As String, Optional pDSCPARAM, Optional pCODSIS)

Public Sub ExecuteScript(xConn As Object, pPathFile As String, Optional pTerminator As String = ";")
   Dim Sql As String
   Dim SqlAux As String
   Dim sStatus As String
   
   'Dim x As DS_BANCO
   
   If ExisteArquivo(pPathFile) Then
      Sql = ReadTextFile(pPathFile)
      Sql = Replace(Sql, Chr(239), "")
      Sql = Replace(Sql, Chr(187), "")
      Sql = Replace(Sql, Chr(191), "")
      
      While InStr(Sql, "/*") <> 0
          Sql = Mid(Sql, 1, InStr(Sql, "/*") - 1) & Mid(Sql, InStr(InStr(Sql, "/*"), Sql, "*/") + 2)
      Wend
      While InStr(Sql, "--") <> 0 And InStr(InStr(Sql, "--"), Sql, Chr(13)) <> 0
          Sql = Mid(Sql, 1, InStr(Sql, "--") - 1) & Mid(Sql, InStr(InStr(Sql, "--"), Sql, Chr(13)) + 2)
      Wend
      If InStr(Sql, "--") <> 0 Then
         Sql = Mid(Sql, 1, InStr(Sql, "--") - 1)
      End If
           
      While InStr(Sql, ";")
         SqlAux = Mid(Sql, 1, InStr(Sql, pTerminator))
         If TypeName(xConn) = "DS_BANCO" Then
            If Not xConn.Executa(SqlAux) Then
               sStatus = sStatus & "Erro : " & SqlAux & vbNewLine
            End If
         Else
            xConn.Execute SqlAux
         End If
         Sql = Mid(Sql, InStr(Sql, pTerminator) + 1)
      Wend
   End If
   If Trim(sStatus) <> "" Then
      sStatus = Now() & vbNewLine & sStatus
      Call WriteIniFile(App.Path & "\" & "ExeScr.log", Right(pPathFile, InStr(StrReverse(pPathFile), "\") - 1), "STATUS", sStatus)
'      MsgBox sStatus
   End If
End Sub
'================================================
'================================================
Public Sub SetRunTimeFormProperty(pForm As Form)
   Dim CurStyle As Long
   Dim NewStyle As Long


   CurStyle = GetWindowLong(pForm.hwnd, GWL_STYLE)
   NewStyle = SetWindowLong(pForm.hwnd, GWL_STYLE, CurStyle) ' Xor (WS_BORDER)) ' Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
   'Call SetWindowLong(pForm.hwnd, GWL_STYLE, GetWindowLong(pForm.hwnd, GWL_STYLE) Xor (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
   'Call SetWindowPos(pForm.hwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub
Public Function CriarToolbar(pSys As Object, pNmToolBar As String) As Object
   Dim oToolBar As Object 'CommandBars
   Dim oBar     As Object 'CommandBar
   Dim n        As Object 'CommandBar
   
   Set oToolBar = pSys.MDI.CommandBars
   With pSys
      '* Verificar se Toolbar Existe
      For Each n In oToolBar
         If n.Title = pNmToolBar Then
            Set oBar = n
            Exit For
         End If
      Next
      
      '* Se Toolbar não Existe então cria
      If oBar Is Nothing Then
         Set oBar = oToolBar.Add(pNmToolBar, 4)  ' 0=xtpBarTop, 4=xtpBarFloating
         oBar.Visible = False
      End If
   End With
   Set CriarToolbar = oBar
End Function
'Public Function CriarButtonToolbar(pToolbar As Object, pType As XTPControlType, pId As Long, _
         Optional pCaption As String, Optional pCategory As String, Optional pStyle As Integer = 2, _
         Optional pBeginGroup As Boolean, Optional pIconId As Long, Optional pChecked As Boolean, _
         Optional pParameter) As Object
Public Function CriarButtonToolbar(pToolbar As Object, pType As Integer, pId As Long, _
         Optional pCaption As String, Optional pCategory As String, Optional pStyle As Integer = 2, _
         Optional pBeginGroup As Boolean, Optional pIconId As Long, Optional pChecked As Boolean, _
         Optional pParameter) As Object
   
   Dim oControl As Object 'CommandBarControl
      
   With pToolbar
      Set oControl = .Controls.Find(pType, pId)
      If oControl Is Nothing Then
         Set oControl = .Controls.Add(pType, pId, pCaption)
         With oControl
            .Category = pCategory
            .IconId = pIconId
            .Checked = pChecked
            .BeginGroup = pBeginGroup
            .Style = pStyle
            If Not IsMissing(pParameter) Then
               .Parameter = pParameter
            End If
         End With
      End If
   End With
   
   Set CriarButtonToolbar = oControl
End Function

Public Sub MontarToolbarDinamico(ByRef pMDI As Object)
'   Dim Control As CommandBarControl
'   Dim ToolBar As CommandBar
'   Dim TBCmdBar As Object
'
'   Dim Sql        As String
'   Dim MyRs       As Object
'   Dim nOrdAntes  As Integer
   
'   Set ToolBar = pMDI.CommandBars.Add("Standard", xtpBarTop)
'
'   Sql = "Select * "
'   Sql = Sql & " From GBARCMD"
'   Sql = Sql & " Where CODSIS = " & SqlStr(Sys.CODSIS)
'   Sql = Sql & " Order By GRUPO, ORDEM, ID"
'
'   nOrdAntes = 0
'   If Sys.xDb.AbreTabela(Sql, MyRs) Then
'      While Not MyRs.EOF
'         Set Control = ToolBar.Controls.Add(XTPControlType.xtpControlButton, MyRs("ID"), MyRs("DSCMODU"))
'         If MyRs.AbsolutePosition > 1 Then
'            Control.BeginGroup = (nOrdAntes = xVal(MyRs("ORDEM")))
'         End If
'         Control.Style = IIf(IsNull(MyRs("IMAGEM")), xtpButtonCaption, xtpButtonIcon)
'         Control.Style = xtpButtonIcon
'
'
'         nOrdAntes = xVal(MyRs("ORDEM"))
'         MyRs.MoveNext
'      Wend
'   End If

End Sub
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
Public Function ValorReal(pValor As String) As Currency
   Dim sValor As String
   Dim nFator As Integer
      
   sValor = Trim(pValor)
   nFator = 1
   If InStr(sValor, "(") <> 0 Or InStr(sValor, "-") Then
      nFator = -1
   End If
   
   
   sValor = Replace(sValor, "(", "")
   sValor = Replace(sValor, ")", "")
   sValor = Replace(sValor, "-", "")
   ValorReal = Val(sValor) * nFator
End Function
