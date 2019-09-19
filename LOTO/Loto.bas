Attribute VB_Name = "Loto"
'Private xDb As adodb.Connection
Global WS As Workspace
Global xDb   As Database
Global xDbMestre As Database
Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer
    Dim f As String
    Dim g As String
    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err, False, True)
    FileExists = IIf(Err = 70, True, FileExists)

    Close intFileNum

    Err = 0
End Function
Public Function CriarBD(pdBase As String, Optional TIPO = "ACCESS") As Boolean
   '================================================================
   '= Última Alteração : 15/03/99                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '********************************************************************
   '**                                                                **
   '** OBJETIVO : Criar um Banco de Dados ACCESS                      **
   '**                                                                **
   '** Recebe: dBase$   - Caminho e Banco de Dados (Ex.:"C:\DSR\X.MDB)**
   '**         Tipo$    - Tipo de Banco de Dados                      **
   '**                    chave                                       **
   '**                                                                **
   '** Retorna: Banco de Dados Criado no caminho especificado.        **
   '**          e True/False indicando se o Banco foi criado ou não.  **
   '**                                                                **
   '********************************************************************
   'dbVersion30
   Dim NewBd As Database
   Dim Pos%, Resto$, Path$
   Dim mPointer%
   On Error GoTo fim
   mPointer% = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   
   CriarBD = False
   '* Acaba com BD problema
   '    Kill dBase$
   '* Cria novo BD
   '* Cria caminho
   Pos = InStr(pdBase, "\")
   Resto = pdBase
   While Pos <> 0
      Resto = Mid(Resto, Pos + 1)
      Pos = InStr(Resto, "\")
   Wend
   Path$ = Mid(pdBase, 1, Len(pdBase) - Len(Resto))
   Call MakePath(Path)
   'TODO: Obsolete DAO method used. Switch to newer method
   
   Set NewBd = CreateDatabase(pdBase, dbLangGeneral, 0) '";LANGID=0x0809;CP=1252;COUNTRY=0", 0)
   NewBd.Close
   
   'TODO: Obsolete DAO method used. Switch to newer method
   Set NewBd = OpenDatabase(pdBase, True, False)
   NewBd.Close
   
   CriarBD = True
   Screen.MousePointer = mPointer%
   
   Exit Function
fim:
   If Errors(0).Number = 53 Then Resume Next
   If Errors(0).Number <> 3204 Then '* Banco já Existe
      'ShowError
   End If
   Screen.MousePointer = mPointer%
End Function
Public Function CriarTabela(pdBase As Variant, Tabela$, Vet_Atrib(), Vet_Indices(), Optional pAPAGAR = False) As Integer
   '------------------------------------------------------------------------
   ' Funcao     : Cria_Tab
   ' Autor      : Diogenes
   ' Data       :
   ' Parametro  : cTabela - Nome da tabela a ser criada
   ' Retorno    : true/false - se criada com sucesso
   ' Obj.       : cria a tabela para impressao de relatorios do
   '              sistema.
   '------------------------------------------------------------------------
   Dim i%, mPointer%
   Dim cTab As New TableDef
   Dim Fld() As New DAO.Field
   Dim Ind() As New DAO.Index
   
   
   CriarTabela = False
   mPointer% = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   On Error GoTo fim
   
   '* Chama Rotina de deleção da tabela
   If pAPAGAR Then
      'Call ExcluirTabela(Tabela)
   End If
   
   cTab.Name = Tabela
   ReDim Fld(0 To (UBound(Vet_Atrib, 1) - 1))
   ReDim Ind(0 To IIf(UBound(Vet_Indices, 1) = 0, 0, (UBound(Vet_Indices, 1) - 1)))
   '*************
   '* Cria Campos
   '*************
   On Error Resume Next
   Fld(0).Name = Vet_Atrib(0, 1)
   If Err = 9 Then '* Subscript out of range
      '* Vetor com Uma Dimensão
      On Error GoTo fim
      Err = 0
      For i = 0 To UBound(Vet_Atrib) - 1
         Fld(i).Name = Vet_Atrib(i)(0)
         Fld(i).Type = Vet_Atrib(i)(1)
         If Fld(i).Type = dbText Then 'Text
            Fld(i).AllowZeroLength = True
         End If
         Fld(i).Size = Vet_Atrib(i)(2)
         cTab.Fields.Append Fld(i)  ' Add field to collection.
      Next
   Else
      '* Vetor com Duas Dimensões
      On Error GoTo fim
      For i = 0 To UBound(Vet_Atrib, 1) - 1
         Fld(i).Name = Vet_Atrib(i, 1)
         Fld(i).Type = Vet_Atrib(i, 2)
         If Fld(i).Type = dbText Then 'Text
            Fld(i).AllowZeroLength = True
         End If
         Fld(i).Size = Vet_Atrib(i, 3)
         cTab.Fields.Append Fld(i)  ' Add field to collection.
      Next
   End If
   '**************
   '* Cria Indices
   '**************
   If UBound(Vet_Indices, 1) > 0 Then
      On Error Resume Next
      Ind(0).Name = Vet_Indices(0, 1)
      If Err = 9 Then '* Subscript out of range
         '* Vetor com Uma Dimensão
         Err = 0
         On Error GoTo fim
         For i = 0 To UBound(Vet_Indices, 1) - 1
            Ind(i).Name = Vet_Indices(i)(0)
            Ind(i).Fields = Vet_Indices(i)(1)
            If UCase(Ind(i).Name) = "PRIMARY KEY" Or UCase(Ind(i).Name) = "PRIMARYKEY" Or UCase(Ind(i).Name) = "PRIMARY" Then
               Ind(i).Primary = True
               Ind(i).Unique = True
            End If
            cTab.Indexes.Append Ind(i)  ' Add index to collection.
         Next
      Else
         On Error GoTo fim
         For i = 0 To UBound(Vet_Indices, 1) - 1
            Ind(i).Name = Vet_Indices(i, 1)
            Ind(i).Fields = Vet_Indices(i, 2)
            If UCase(Ind(i).Name) = "PRIMARY KEY" Or UCase(Ind(i).Name) = "PRIMARYKEY" Or UCase(Ind(i).Name) = "PRIMARY" Then
               Ind(i).Primary = True
               Ind(i).Unique = True
            End If
            cTab.Indexes.Append Ind(i)  ' Add index to collection.
         Next
      End If
   End If
   
   '* Cria Tabela
   pdBase.TableDefs.Append cTab
   CriarTabela = True
   Screen.MousePointer = mPointer%
   Exit Function
fim:
   Screen.MousePointer = mPointer%
'   ShowError
End Function
Public Function MakePath(ByVal strDir As String, Optional ByVal fAllowIgnore) As Boolean
    If IsMissing(fAllowIgnore) Then
        fAllowIgnore = True
    End If
    
    Do
        If MakePathAux(strDir) Then
            MakePath = True
            Exit Function
        Else
            Dim strMsg As String
            Dim iRet As Integer
            
'            strMsg = ResolveResString(resMAKEDIR) & LF$ & strDir
            iRet = MsgBox(strMsg, IIf(fAllowIgnore, vbAbortRetryIgnore, vbRetryCancel) Or vbExclamation Or vbDefaultButton2, "")
            Select Case iRet
            Case vbAbort, vbCancel
'                ExitSetup frmCopy, gintRET_ABORT
            Case vbIgnore
                MakePath = False
                Exit Function
            Case vbRetry
            End Select
        End If
    Loop
End Function
Public Function MakePathAux(ByVal strDirName As String) As Boolean
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Add trailing backslash
    '
   If Right$(strDirName, 1) <> "\" Then
        strDirName = strDirName & "\"
    End If

    strOldPath = CurDir$
    MakePathAux = False
    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.
    '
    '
    intOffset = InStr(intAnchor + 1, strDirName, "\")
    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, "\")
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir strPath
            If Err Then
                ' We must create this directory
                Err = 0
                #If Win32 And LOGGING Then
                    NewAction gstrKEY_CREATEDIR, """" & strPath & """"
                #End If
                MkDir strPath
                #If Win32 And LOGGING Then
                    If Err Then
                        LogError ResolveResString(resMAKEDIR) & " " & strPath
                        AbortAction
                        GoTo Done
                    Else
                        CommitAction
                    End If
                #End If
            End If
        End If
    Loop Until intAnchor = 0

    MakePathAux = True
Done:
    ChDir strOldPath

    Err = 0
End Function
Public Sub CriarBancoDeDados(ByRef pDb As Variant, sArquivo As String)
   Dim StrConect As String
   Dim VetAtrib(), VetInd(0)
   Call CriarBD(sArquivo)

   StrConect = "Provider=Microsoft.Jet.OLEDB.4.0;"
   StrConect = StrConect & "Data Source=" & sArquivo
    
   If WS Is Nothing Then
      Set DBEngine = Nothing
      Set WS = DBEngine.CreateWorkspace("WsEngine", "admin", "")
   End If
   
   StrConect = ";"
   Set pDb = WS.OpenDatabase(sArquivo, False, False, mvarStrConect$)
 
'   Set pDb = New adodb.Connection
'   pDb.CommandTimeout = 15
'   pDb.CursorLocation = adUseClient
'   pDb.ConnectionString = StrConect
'   pDb.Open
   
   'ReDim VetAtrib(1)
   'VetAtrib(0) = Array("IDJOGO", "10", "6")
   'Call CriarTabela(pDb, "JOGOS", VetAtrib(), VetInd())
   
   ReDim VetAtrib(4)
   VetAtrib(0) = Array("IDJOGO", dbText, "6")
   VetAtrib(1) = Array("IDCARTAO", dbDouble, "10")
   VetAtrib(2) = Array("VALIDO", dbText, "1")
   VetAtrib(3) = Array("VERIFICADO", dbText, "1")
   Call CriarTabela(pDb, "CARTOES", VetAtrib(), VetInd())
   
   ReDim VetAtrib(3)
   VetAtrib(0) = Array("IDJOGO", dbText, "6")
   VetAtrib(1) = Array("IDCARTAO", dbDouble, "10")
   VetAtrib(2) = Array("NUMERO", dbDouble, "2")
   Call CriarTabela(pDb, "NUMEROS", VetAtrib(), VetInd())
   
   ReDim VetAtrib(3)
   VetAtrib(0) = Array("IDJOGO", dbText, "6")
   VetAtrib(2) = Array("IDCARTAO", dbDouble, "10")
   VetAtrib(1) = Array("SOMA", dbDouble, "2")
   Call CriarTabela(pDb, "SOMAS", VetAtrib(), VetInd())
   pDb.Close
End Sub
Public Function Copy(Orig$, dest$)
   Dim nMsg$, nTipo&, NL
   Dim Resp%
   NL = vbLf
   On Error Resume Next
   If FileExists(Orig$) Then
      Kill dest$
      'Call Del(dest$)
      FileCopy Orig$, dest$
   Else
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(11) + UCase(Orig$), ClsMsg.LoadMsg(12))
      Resp = vbCancel
      Exit Function
   End If
   Resp = vbYes
   Select Case Err
      Case 71
         While Resp = vbYes
            nMsg = ClsMsg.LoadMsg(13) + NL + NL
            nMsg = nMsg & ClsMsg.LoadMsg(14) + NL
            nMsg = nMsg & ClsMsg.LoadMsg(15)
            nTipo = vbYesNo + vbCritical + vbDefaultButton1
            Resp = MsgBox(nMsg, nTipo, ClsMsg.LoadMsg(16))
            If Resp = vbYes Then
               Err = 0
               FileCopy Orig$, dest$
            End If
         Wend
      Case 70
         While Resp = vbOK
            nMsg = ClsMsg.LoadMsg(7) + NL + NL
            nMsg = nMsg & ClsMsg.LoadMsg(56) + NL
            nTipo = vbYesNo + vbCritical + vbDefaultButton1
            Resp = MsgBox(nMsg, nTipo, ClsMsg.LoadMsg(16))
            If Resp = vbYes Then
               Err = 0
               FileCopy Orig$, dest$
            End If
         Wend
   End Select
   Copy = Resp
End Function
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
      sCaption = pForm.Caption
   End If
   
   If LCase(Right(NomeArq, 3)) <> "xls" Then
      If InStr(NomeArq, ".") <> 0 Then
         NomeArq = Mid(NomeArq, 1, InStr(NomeArq, ".") - 1)
      End If
      NomeArq = NomeArq & ".xls"
   End If
      
   If ExcluiArq Then
      If FileExists(App.Path & "\" & NomeArq) Then
         Kill App.Path & "\" & NomeArq
      End If
   Else
      ExcluiArq = Not FileExists(App.Path & "\" & NomeArq)
   End If
   
   Set xlApp = CreateObject("Excel.Application")
   If Not ExcluiArq Then
      On Error Resume Next
      Set xlBook = xlApp.Workbooks.Open(App.Path & "\" & NomeArq)
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
               If Grd.ColAlignment(j) = 1 Or Grd.ColAlignment(j) = 2 Or Grd.ColAlignment(j) = 9 Then
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
            xlSheet.Cells(i + 1, k).Font.Size = Grd.Font.Size '* 8,25
            
            '***********
            '* Verificar reconhecimento da cor
            'xlApp.Visible = True
            'Grd.Redraw = True
            Grd.Row = i
            Grd.Col = j
            
            If Grd.CellAlignment = 0 Then
               If Grd.ColAlignment(j) = 1 Or Grd.ColAlignment(j) = 2 Or Grd.ColAlignment(j) = 9 Then
                  xlSheet.Cells(i + 1, k).HorizontalAlignment = 2
               ElseIf ClsDsr.InArray(Grd.ColAlignment(j), Array(3, 4, 5)) Then
                  xlSheet.Cells(i + 1, k).HorizontalAlignment = 3
               ElseIf ClsDsr.InArray(Grd.ColAlignment(j), Array(6, 7, 8)) Then
                  xlSheet.Cells(i + 1, k).HorizontalAlignment = 4
               End If
            Else
               xlSheet.Cells(i + 1, k).HorizontalAlignment = 2
            End If
            If Val(Grd.CellBackColor) = 0 Then
               xlSheet.Cells(i + 1, k).Interior.Color = Grd.BackColor
            Else
               xlSheet.Cells(i + 1, k).Interior.Color = Grd.CellBackColor
            End If
            If Val(Grd.CellForeColor) = 0 Then
               xlSheet.Cells(i + 1, k).Font.Color = Grd.ForeColor
            Else
               xlSheet.Cells(i + 1, k).Font.Color = Grd.CellForeColor
            End If
                        
            If xlSheet.Cells(i + 1, k).Interior.Color = 0 Then xlSheet.Cells(i + 1, k).Interior.Color = vbWhite
            If xlSheet.Cells(i + 1, k).Font.Color = 0 Then xlSheet.Cells(i + 1, k).Font.Color = vbBlack
            
            If xlSheet.Cells(i + 1, k).Interior.Color = vbWhite Or Val(Grd.CellBackColor) = 0 Then
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
      'If ClsCtrl.GetTag(Grd, "CANCEL") = "True" Then
      '   Call ClsCtrl.SetTag(Grd, "CANCEL", False)
      '   i = Grd.Rows - 1
      'End If
   Next
   
   If ExibeFlood Then
      pForm.Caption = sCaption
   End If
   
   If isVisible Then xlApp.Visible = True
   On Error Resume Next
   If ExcluiArq Then
      If FileExists(App.Path & "\" & NomeArq) Then
         Kill App.Path & "\" & NomeArq
      End If
   End If
   
   If FileExists(App.Path & "\" & NomeArq) Then
      Call xlSheet.SaveAs(App.Path & "\" & NomeArq & "z")
      Call ClsDos.Del(App.Path & "\" & NomeArq)
      Call xlSheet.SaveAs(App.Path & "\" & NomeArq)
      Call ClsDos.Del(App.Path & "\" & NomeArq & "z")
   Else
      Call xlSheet.SaveAs(App.Path & "\" & NomeArq)
   End If
   
   If Err = 0 Then
      sMsg = "O Arquivo '" & App.Path & "\" & NomeArq & "'"
      sMsg = sMsg & " foi salvo com sucesso!!!"
      If isVisible Then
         If ExibeMsg Then
            MsgBox sMsg
         End If
      Else
         If ExibeMsg Then
            sMsg = sMsg & vbNewLine & vbNewLine
            sMsg = sMsg & "Deseja Visualizá-lo?"
            If vbYes = MsgBox(sMsg, vbYesNo) Then
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
   If Err <> 1004 And Err <> 0 Then
      MsgBox Err & " - " & Error
   End If
   
   Grd.Row = iLin
   Grd.Col = iCol
   Grd.Redraw = True
   Screen.MousePointer = vbDefault
End Function

