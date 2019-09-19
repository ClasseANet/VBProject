VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Banco"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function ClonarRS(ByVal pRecordSet As Object, Optional pFiltro As String) As Object
   Dim NewRs   As ADODB.Recordset
   Dim i       As Long
   
   On Error GoTo TratarErro
   Screen.MousePointer = vbHourglass
   DoEvents
   
   pRecordSet.MoveFirst
   '**************************************
   '* Inicia a montagem do novo RecordSet
   '**************************************
   Set NewRs = New ADODB.Recordset

   NewRs.ActiveConnection = Nothing
   NewRs.CursorLocation = adUseClient
   NewRs.LockType = adLockBatchOptimistic
   
   For i = 0 To pRecordSet.Fields.Count - 1
      NewRs.Fields.Append pRecordSet.Fields(i).Name, pRecordSet.Fields(i).Type, pRecordSet.Fields(i).DefinedSize, adFldUpdatable + adFldIsNullable + adFldMayBeNull
      NewRs.Fields(i).NumericScale = pRecordSet.Fields(i).NumericScale
      NewRs.Fields(i).Precision = pRecordSet.Fields(i).Precision
   Next
   
   NewRs.Open
   If pFiltro <> "" Then
      pRecordSet.Filter = ""
      pRecordSet.Filter = pFiltro
   End If
   Do While Not pRecordSet.EOF
      DoEvents
      NewRs.AddNew
      For i = 0 To pRecordSet.Fields.Count - 1
         If Not IsNull(pRecordSet.Fields(i).Value) Then
            Select Case GetTypeField(pRecordSet.Fields(i).Name, pRecordSet)
               Case vbDecimal
                  NewRs.Fields(i).Value = CDbl(pRecordSet.Fields(i).Value & "")
               Case vbString
                  NewRs.Fields(i).Value = pRecordSet.Fields(i).Value & ""
               Case vbDate
                  If pRecordSet.Fields(i).Value & "" <> "" Then
                     NewRs.Fields(i).Value = CDate(pRecordSet.Fields(i).Value & "")
                  End If
            End Select
         End If
      Next
      NewRs.Update
      
      pRecordSet.MoveNext
   Loop
   
   Set ClonarRS = NewRs
   GoTo Fim
   
TratarErro:
   MsgBox "Erro ao montar novo RS. Erro :" & Err.Description, vbExclamation + vbOKOnly, "ClonarRS"
Fim:
   Screen.MousePointer = vbDefault
End Function
Public Function CriarRS(pColl As Collection) As Object
   Dim Rs As ADODB.Recordset
   Dim n
   
   Set Rs = New ADODB.Recordset
   
   With Rs
       .ActiveConnection = Nothing
       .CursorLocation = adUseClient
       .LockType = adLockBatchOptimistic
       With .Fields
           For Each n In pColl
               If n.Tipo = adChar Then
                   .Append n.Nome, n.Tipo, n.Tamanho
               Else
                   .Append n.Nome, n.Tipo
               End If
           Next
       End With
       .Open
   End With
   Set CriarRS = Rs
   Set Rs = Nothing
End Function
'---------------------------------------------------------------------------------------
' Método/Função       : CruzRef
' Data e hora         :
' Criado/Alterado por : Marcio Luis Severo
' Parametro           :  pRecordset - Recordset com os dados
'                        pClsDetahles - objeto do tipo Detalhes com a descrição do campo padrão
'                       pCampo - String GetTag com os campos para formar a referencia nomeados CAMPO1,
'                       CAMPO2, ... CAMPOn e mais o campo que será usado para buscar o valor que deve
'                       ser nomeado como VALOR
'                       pQtdCampo - Inteiro com a quantidade de campos do tag pCampo exceto o campo de valor
'                       pTotal - Booleano que indica se os valores serão totalizados
' Retorna             :  Recordset criado
' Funcionalidade      : Cria um Recordset de referencia cruzada à partir do recordset passado
'---------------------------------------------------------------------------------------
Public Function CruzRef(pRecordSet As ADODB.Recordset, pClsDetahles As Object, pCampo As String _
         , pQtdCampo As Integer, Optional pTotalLinha As Boolean = False _
         , Optional pTotalColuna As Boolean = False, Optional pQtdValor As Integer = 1 _
         ) As Object
   
   Dim clDet      As Detalhes
   Dim cll        As New Collection
   Dim Rs         As ADODB.Recordset
   Dim cTitCampo  As String
   Dim cNomeCampo As String
   Dim n          As Integer
   Dim m          As Integer
   Dim cCampoRef  As String
   
   pRecordSet.MoveFirst
   If pRecordSet.EOF = True Then
       Exit Function
   End If
   
   On Error GoTo CzFIM
   
   'carrega os dados para definir as colunas do recordset
   
   'Cria a coluna de Referencia
   'Set clDet = New Detalhes
   Set clDet = New Detalhes
   clDet.Nome = pClsDetahles.Nome
   clDet.Tipo = pClsDetahles.Tipo
   clDet.Tamanho = pClsDetahles.Tamanho
   cll.Add clDet
   
   Set clDet = New Detalhes
   clDet.Nome = "Tipo"
   clDet.Tipo = adChar
   clDet.Tamanho = 20
   cll.Add clDet
   
   'Cria as colunas de dados
   Do Until pRecordSet.EOF = True
      cTitCampo = ""
      For n = 1 To pQtdCampo
         cNomeCampo = "CAMPO" & n
         cTitCampo = cTitCampo & pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cNomeCampo)).Value & "" & "-"
      Next
      cTitCampo = VBA.Left(cTitCampo, Len(cTitCampo) - 1)
      If cTitCampo = "" Then cTitCampo = "Sem Id"
      If Not ClsGeneral.ExisteItem(cll, cTitCampo) Then
          Set clDet = New Detalhes
          clDet.Nome = cTitCampo
          clDet.Tipo = adDouble
          cll.Add clDet, cTitCampo
      End If
      pRecordSet.MoveNext
   Loop
   
   'Cria a coluna de total
   If pTotalLinha Then
      Set clDet = New Detalhes
      clDet.Nome = "TOTAL"
      clDet.Tipo = adDouble
      cll.Add clDet
   End If
   
   'Cria o Recordset
   Set Rs = CriarRS(cll)
   
   'Se tem Total Geral
   If pTotalColuna = True Then
      Rs.AddNew
      'cCampoRef = pRecordset.Fields(pClsDetahles.Campo).Value & ""
      Rs.Fields(pClsDetahles.Nome) = "ÿTOTAL"
      Rs.Update
      Rs.MoveFirst
   End If
   
   Dim cValTit As String
   
   'Popula o recordset
   pRecordSet.MoveFirst
   Do Until pRecordSet.EOF = True
      For m = 1 To pQtdValor
         cValTit = "VALOR" & m
         
         cCampoRef = pRecordSet.Fields(pClsDetahles.Campo).Value & ""
         cTitCampo = ""
         For n = 1 To pQtdCampo
            cNomeCampo = "CAMPO" & n
            cTitCampo = cTitCampo & pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cNomeCampo)).Value & "" & "-"
         Next
         cTitCampo = VBA.Left(cTitCampo, Len(cTitCampo) - 1)
         If cTitCampo = "" Then cTitCampo = "Sem Id"
         Rs.Filter = pClsDetahles.Nome & " = " & xDb.SqlStr(cCampoRef) & " AND Tipo = " & xDb.SqlStr(ClsGeneral.GetTag(pCampo, cValTit))
         
         If Rs.EOF = True Then
            Rs.AddNew
            Rs.Fields(pClsDetahles.Nome) = Trim(cCampoRef)
            Rs.Fields("Tipo").Value = ClsGeneral.GetTag(pCampo, cValTit)
            Rs.Fields(cTitCampo) = ClsGeneral.xVal(pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cValTit)).Value & "", 2)
            If pTotalLinha Then
              Rs.Fields("TOTAL") = ClsGeneral.xVal(Rs.Fields("TOTAL") + pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cValTit)), 2)
            End If
            Rs.Update
            Rs.Filter = ""
            Rs.MoveFirst
            If pTotalColuna Then
               Rs.Find pClsDetahles.Nome & " = " & xDb.SqlStr("ÿTOTAL")
               Rs.Fields(cTitCampo) = ClsGeneral.xVal(Rs.Fields(cTitCampo) + pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cValTit)), 2)
               If pTotalLinha Then
                 Rs.Fields("TOTAL") = ClsGeneral.xVal(Rs.Fields("TOTAL") + pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cValTit)), 2)
               End If
               Rs.Update
               Rs.MoveFirst
            End If
         Else
            Rs.Fields(cTitCampo) = pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cValTit))
            If pTotalLinha Then
               Rs.Fields("TOTAL") = Rs.Fields("TOTAL") + pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cValTit))
            End If
            Rs.Update
            Rs.Filter = ""
            Rs.MoveFirst
            If pTotalColuna Then
               Rs.Find pClsDetahles.Nome & " = " & xDb.SqlStr("ÿTOTAL")
               Rs.Fields(cTitCampo) = Rs.Fields(cTitCampo) + pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cValTit))
               If pTotalLinha Then
                 Rs.Fields("TOTAL") = Rs.Fields("TOTAL") + pRecordSet.Fields(ClsGeneral.GetTag(pCampo, cValTit))
               End If
               Rs.Update
               Rs.MoveFirst
            End If
         End If
         Rs.Filter = ""
      Next
      pRecordSet.MoveNext
   Loop
   
   Set CruzRef = Rs
   Set Rs = Nothing
   Set pRecordSet = Nothing
   Exit Function

CzFIM:
   'ShowError
   MsgBox Err & "-" & Error
   Exit Function
   Resume
End Function
Public Sub ExecuteScript(ByVal xConn As Object, pPathFile As String, Optional pTerminator As String = "", Optional pRoleName As String = "db_ddladmin", Optional pMemberName As String = "USU_VERIF")
   Dim Sql As String
   Dim SqlAux As String
   Dim sStatus As String
   Dim sTerminator As String
   Dim sTerminator1 As String
   Dim sTerminator2 As String
   Dim sGO As String
   
   If InStr(UCase(pPathFile), "REV43") <> 0 Or InStr(UCase(pPathFile), "REV44") <> 0 Then
      Dim MyConn As Object
      Dim sConect As String
      Set MyConn = CreateObject("ADODB.Connection") ' New ADODB.Connection
      With MyConn
         .CommandTimeout = 300
         .CursorLocation = 3
         
         sConect = "Provider=SQLOLEDB;"
         sConect = sConect & "Initial Catalog=" & xConn.dbName & ";"
         sConect = sConect & "Data Source=" & xConn.Server & ";"
         sConect = sConect & "Integrated Security=SSPI;"
         
         .ConnectionString = sConect
         .Open
         
         .Execute "[sp_addrolemember] @rolename = '" & pRoleName & "', @membername = '" & pMemberName & "'"
         .Close
      End With
       Set MyConn = Nothing
   End If
   
   
   'Dim x As DS_BANCO
   sTerminator = IIf(pTerminator = "", ";", pTerminator)
   If ClsAmbiente.ExisteArquivo(pPathFile) Then
      
      Sql = ClsAmbiente.ReadTextFile(pPathFile)
      Sql = Replace(Sql, Chr(239), "")
      Sql = Replace(Sql, Chr(187), "")
      Sql = Replace(Sql, Chr(191), "")
      
      While InStr(Sql, "/*") <> 0
          Sql = Mid(Sql, 1, InStr(Sql, "/*") - 1) & Mid(Sql, InStr(InStr(Sql, "/*"), Sql, "*/") + 2)
      Wend
      While InStr(Sql, "--") <> 0
         If InStr(InStr(Sql, "--"), Sql, Chr(13)) <> 0 Then
            Sql = Mid(Sql, 1, InStr(Sql, "--") - 1) & Mid(Sql, InStr(InStr(Sql, "--"), Sql, Chr(13)) + 2)
         Else
            Sql = Mid(Sql, 1, InStr(Sql, "--") - 1)
         End If
      Wend
      If InStr(Sql, "--") <> 0 Then
         Sql = Mid(Sql, 1, InStr(Sql, "--") - 1)
      End If
                 
      'sGO = "GO" & Chr(13) & Chr(10)
      sGO = Chr(10) & "GO"
      sTerminator1 = UCase(sTerminator)
      sTerminator2 = UCase(IIf(sTerminator1 = sGO, ";", sGO))
      While InStr(UCase(Sql), sTerminator1) <> 0 Or InStr(UCase(Sql), sTerminator2) <> 0
         If InStr(UCase(Sql), sTerminator1) = 0 And InStr(UCase(Sql), sTerminator2) = 0 Then
            sTerminator = ""
         Else
            If InStr(UCase(Sql), sTerminator1) = 0 Or InStr(UCase(Sql), sTerminator2) = 0 Then
               sTerminator = IIf(InStr(UCase(Sql), sTerminator1) = 0, sTerminator2, sTerminator1)
            Else
               sTerminator = IIf(InStr(UCase(Sql), sTerminator1) < InStr(UCase(Sql), sTerminator2), sTerminator1, sTerminator2)
            End If
         End If
         
         SqlAux = Mid(Sql, 1, InStr(UCase(Sql), sTerminator) - 1)
         While ClsGeneral.InArray(Asc(Mid(SqlAux, 1, 1)), Array(10, 13)): SqlAux = Mid(SqlAux, 2): Wend
         While ClsGeneral.InArray(Asc(Right(SqlAux, 1)), Array(10, 13)): SqlAux = Mid(SqlAux, 1, Len(SqlAux) - 1): Wend
         If Mid(UCase(SqlAux), 1, Len("EXECUTE ")) = "EXECUTE " Then SqlAux = Mid(SqlAux, Len("EXECUTE ") + 1)
         If Mid(UCase(SqlAux), 1, Len("EXEC ")) = "EXEC " Then SqlAux = Mid(SqlAux, Len("EXEC ") + 1)
         
         'Debug.Print SqlAux
         If Trim(SqlAux) <> "" Then
            If TypeName(xConn) = "DS_BANCO" Then
               If Not xConn.Executa(SqlAux, pExibeMsg:=False) Then
                  SqlAux = Replace(SqlAux, "EXEC ", " ")
                  SqlAux = Replace(SqlAux, "EXECUTE ", " ")
                  If Not xConn.Executa(SqlAux, pExibeMsg:=(Trim(sStatus) = "")) Then
                     sStatus = sStatus & "Erro : " & SqlAux & vbNewLine
                  End If
               End If
            Else
               xConn.Execute SqlAux
            End If
         End If
         Sql = Mid(Sql, InStr(UCase(Sql), sTerminator) + Len(sTerminator))
      Wend
   End If
   If Trim(sStatus) <> "" Then
      sStatus = Now() & vbNewLine & sStatus
      Call ClsAmbiente.WriteIniFile(App.Path & "\" & "ExeScr.log", Right(pPathFile, InStr(StrReverse(pPathFile), "\") - 1), "STATUS", sStatus)
'      MsgBox sStatus
   End If
End Sub
'*********************************************************************
'* o Funcionalidade:
'*    "  Retorna um valor constante do tipo de campo de um recordset
'* o Parâmetros:
'*    "  pFieldName :Nome do Campo
'*    "  pRecordSet : ADODB.RecordSet
'* o Retorno: Constante membro de VBA.VbVarType
'*    "  Const vbString  = 8
'*    "  Const vbDecimal = 14
'*    "  Const vbDate    = 7
'*********************************************************************
Public Function GetTypeField(pFieldName As String, pRecordSet As Object) As VbVarType
   Dim nType   As Long
   
   If Not pRecordSet Is Nothing Then
      Select Case pRecordSet.Fields(pFieldName).Type
         Case DataTypeEnum.adChar, DataTypeEnum.adVarChar, DataTypeEnum.adChapter, _
              DataTypeEnum.adLongVarChar
'DAO      dbChar, dbMemo, dbText, adChar, adVarChar, adWChar, adVarWChar
            
            nType = vbString
         
         Case DataTypeEnum.adBigInt, DataTypeEnum.adInteger, DataTypeEnum.adBinary, _
              DataTypeEnum.adCurrency, DataTypeEnum.adDecimal, DataTypeEnum.adDouble, _
              DataTypeEnum.adNumeric
'         dbBigInt, dbBinary, dbBoolean, dbByte, dbCurrency, dbDecimal, _
          dbDouble, dbInteger, dbLong, dbNumeric, dbSingle, dbLongBinary, _
          dbGUID, dbVarBinary, adNumeric, adSmallInt, adSingle, adInteger, _
          adDouble, adDecimal
            
            nType = vbDecimal
         
         Case DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTimeStamp
'         dbDate, dbTime, dbTimeStamp, adDBTimeStamp
            nType = vbDate
                  
      End Select
   End If
   GetTypeField = nType
End Function
Public Function RecordSetToExcel(ByRef Rs As Object, Optional ByVal Nome, Optional ByVal isVisible As Boolean = False, Optional ByRef pForm, Optional ByVal TopFlood, Optional ByVal ExcluiArq As Boolean = True, Optional ByVal NomeArq, Optional ByVal ExibeMsg As Boolean = True) As Boolean
   Dim xlApp         As Object
   Dim xlBook        As Object
   Dim xlSheet       As Object

   Dim i             As Integer
   Dim k             As Integer
   Dim sMsg          As String
      
   'Dim ProgBar       As New CProgBar32
   Dim ExibeFlood    As Boolean
   Dim sCaption      As String
   Dim lTmpArq       As Boolean 'Indica se o arquivo será gravado no diretório tmp
   Dim sPathTemp     As String
   
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
      
   ExibeFlood = Not IsMissing(pForm)
      
   If IsMissing(Nome) Then Nome = "TAB"
   If IsMissing(NomeArq) Then NomeArq = Nome
   
   sPathTemp = ClsAmbiente.GetTempFolder
   
   If InStr(1, NomeArq, "\") > 0 Then
      lTmpArq = False
   Else
      lTmpArq = True
   End If
   
   If ExibeFlood Then
'      Set ProgBar.Parent = pForm
'      ProgBar.Create 0, pForm.Height - 600, pForm.Width - 90, 15
'      ProgBar.MinValue = 1
'      ProgBar.MaxValue = Rs.RecordCount
'      sCaption = pForm.Caption
   End If

   If LCase(Right(NomeArq, 3)) <> "xls" Then
      If InStr(NomeArq, ".") <> 0 Then
         NomeArq = Mid(NomeArq, 1, InStr(NomeArq, ".") - 1)
      End If
      NomeArq = NomeArq & ".xls"
   End If
   

   If ExcluiArq Then
      If lTmpArq Then
         Call ClsAmbiente.ExcluirArquivo(sPathTemp & NomeArq, ExibeMsg)
      Else
         Call ClsAmbiente.ExcluirArquivo((NomeArq), ExibeMsg)
      End If
   Else
      If lTmpArq Then
         ExcluiArq = Not ClsAmbiente.ExisteArquivo(sPathTemp & NomeArq)
      Else
         ExcluiArq = Not ClsAmbiente.ExisteArquivo(NomeArq)
      End If
   End If
   
   Set xlApp = CreateObject("Excel.Application")
   If Not ExcluiArq Then
      On Error Resume Next
      If lTmpArq Then
         Set xlBook = xlApp.Workbooks.Open(sPathTemp & NomeArq)
      Else
         Set xlBook = xlApp.Workbooks.Open(NomeArq)
      End If
      Set xlSheet = xlBook.Worksheets(Nome)
      If Err = 0 Then
         Call ClsMensagem.ExibirAviso("Planilha '" & Mid(NomeArq, 1, Len(NomeArq) - 4) & "' Já Existe.", "Atenção")
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
      xlSheet.Cells(1, k).Font.Size = 8.25
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
   
   '***************
   '* Copiar Para o Excel
   Rs.MoveFirst
   'Dim MApp As Excel.Application
   xlSheet.Range("A2").CopyFromRecordset Rs
   With xlSheet
      xlApp.Intersect(.Range(.Rows(2), .Rows(Rs.RecordCount + 1)), .Range(.Columns(1), .Columns(Rs.Fields.Count))).Value = xlApp.Transpose(Rs.GetRows(Rs.RecordCount))
      Rs.MoveFirst
      .Application.Intersect(.Range(.Rows(2), .Rows(Rs.RecordCount + 1)), .Range(.Columns(1), .Columns(Rs.Fields.Count))).Value = .Application.Transpose(Rs.GetRows(Rs.RecordCount))
   End With
   
   If ExibeFlood Then
'      pForm.Caption = sCaption
'      Call ProgBar.DestroyProgBar
   End If
   
   If isVisible Then xlApp.Visible = True
   On Error Resume Next
   If ExcluiArq Then
      If lTmpArq Then
         Call ClsAmbiente.ExcluirArquivo(sPathTemp & NomeArq, ExibeMsg)
      Else
         Call ClsAmbiente.ExcluirArquivo((NomeArq), ExibeMsg)
      End If
   End If
   If lTmpArq Then
      If ClsAmbiente.ExisteArquivo(sPathTemp & NomeArq) Then
         Call xlSheet.SaveAs(sPathTemp & NomeArq & "z")
         Call ClsAmbiente.ExcluirArquivo(sPathTemp & NomeArq, ExibeMsg)
         Call xlSheet.SaveAs(sPathTemp & NomeArq)
         Call ClsAmbiente.ExcluirArquivo(sPathTemp & NomeArq & "z", ExibeMsg)
      Else
         Call xlSheet.SaveAs(sPathTemp & NomeArq)
      End If
   Else
      If ClsAmbiente.ExisteArquivo(NomeArq) Then
         Call xlSheet.SaveAs(NomeArq & "z")
         Call ClsAmbiente.ExcluirArquivo((NomeArq), ExibeMsg)
         Call xlSheet.SaveAs(NomeArq)
         Call ClsAmbiente.ExcluirArquivo(NomeArq & "z", ExibeMsg)
      Else
         Call xlSheet.SaveAs(NomeArq)
      End If
   End If
   If Err = 0 Then
      If lTmpArq Then
         sMsg = "O Arquivo '" & sPathTemp & NomeArq & "'"
      Else
         sMsg = "O Arquivo '" & NomeArq & "'"
      End If
      sMsg = sMsg & " foi salvo com sucesso!!!"
      If isVisible Then
         If ExibeMsg Then
            Call ClsMensagem.ExibirAviso(sMsg, "Atenção")
         End If
      Else
         If ExibeMsg Then
            sMsg = sMsg & vbNewLine & vbNewLine
            sMsg = sMsg & "Deseja Visualizá-lo?"
            If vbYes = ClsMensagem.ExibirPergunta(sMsg, "Atenção") Then
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
'   Set ProgBar = Nothing
   
   Screen.MousePointer = vbDefault
   Exit Function

TrataErro:
   If Err = 1004 Or Err = 9 Then
      Resume Next
   Else
      ClsMensagem.ExibirErro
   End If
   GoTo Saida
End Function
Public Function SqlDate(ByVal DT As String, Optional Format_Date As Integer = 3, Optional InsereNull As Boolean = True, Optional pDbTipo As Integer = 1) As String
   Dim MyFormat As String
   Dim Pos As Integer
   Dim sData As String
   Dim dData As Date
   Dim sAux As String
   
   On Error GoTo TrataErro
'* em Sql 103 - dd/mm/yy e 101 - mm/dd/yy
   If Val(DT) = 0 Then
      SqlDate = IIf(InsereNull, "Null", "")
   Else
      Select Case pDbTipo
         Case 0:     SqlDate = "CDATE('" & Format$(DT, "mm/dd/yyyy") & "')" 'eDbTipo.Access
         Case 1 'eDbTipo.SQL_SERVER
            '"mm/dd/yyyy" = "mm/dd/yyyy"
            Select Case Format_Date
               Case 1: MyFormat = "mm/dd/yyyy"                 'eSysDate.Data
               Case 2: MyFormat = "mm/dd/yyyy" & " hh:mm:ss"   'eSysDate.Data_Hora
               Case 3: MyFormat = "hh:mm:ss"                   'eSysDate.HORA
               Case Else: MyFormat = "mm/dd/yyyy"
            End Select
            DT = Trim(DT)
            sData = ClsGeneral.StrZero(Mid(DT, 1, InStr(DT, "/") - 1), 2)
            Pos = InStr(DT, "/") + 1
            If InStr(Pos, DT, "/") - Pos >= 0 Then
               sData = sData & "/" & ClsGeneral.StrZero(Mid(DT, Pos, InStr(Pos, DT, "/") - Pos), 2)
            End If
            Pos = InStr(Pos, DT, "/") + 1
            
            sAux = DT
            If InStr(DT, " ") <> 0 Then
               sAux = Mid(DT, 1, InStr(DT, " ") - 1)
            End If
            If Len(Mid(sAux, Pos)) = 2 Then
               sData = sData & "/" & VBA.Right$(Mid(sAux, Pos), 2)
            ElseIf Len(Mid(sAux, Pos)) = 4 Then
               If Mid(Mid(sAux, Pos), 1, 2) > 18 Then
                  sData = sData & "/" & Mid(sAux, Pos)
               Else
                  sData = sData & "/" & VBA.Right$(Mid(sAux, Pos), 2)
               End If
            Else
               sData = sData & "/" & VBA.Right$(Mid(sAux, Pos), 2)
            End If
            dData = VBA.CDate(sData)
            sData = ""
            If Format_Date = 1 Then 'eSysDate.Data
               sData = ClsGeneral.StrZero(Day(dData), 2) & "/" & ClsGeneral.StrZero(Month(dData), 2) & "/" & Year(dData)
            ElseIf Format_Date = 2 Then 'eSysDate.HORA
               sData = ClsGeneral.StrZero(Hour(DT), 2) & ":" & ClsGeneral.StrZero(Minute(DT), 2) & ":" & ClsGeneral.StrZero(Second(DT), 2)
            ElseIf Format_Date = 3 Then 'eSysDate.Data_Hora
               sData = ClsGeneral.StrZero(Day(dData), 2) & "/" & ClsGeneral.StrZero(Month(dData), 2) & "/" & Year(dData)
               sData = sData & " " & ClsGeneral.StrZero(Hour(DT), 2) & ":" & ClsGeneral.StrZero(Minute(DT), 2) & ":" & ClsGeneral.StrZero(Second(DT), 2)
            Else
               sData = ClsGeneral.StrZero(Day(dData), 2) & "/" & ClsGeneral.StrZero(Month(dData), 2) & "/" & Year(dData)
            End If
            
            
            'SqlDate = " CONVERT(DATETIME, '" & Format$(DT, MyFormat) & "', 103) "
            SqlDate = "CAST('" & sData & "' As DATETIME) "

         Case 2:     SqlDate = "TO_DATE('" & Format$(DT, "mm/dd/yyyy") & "')" 'eDbTipo.ORACLE
         'Case Else:               SqlDate = "CDATE('" & Format$(DT, "mm/dd/yyyy") & "')"
      End Select
   End If
Exit Function
TrataErro:
   Call ClsMensagem.ExibirStop("Formato de Data Inválido", "Atenção")
End Function
Public Function SqlNum(ByVal Num As String, Optional InsereNull As Boolean = False) As String
   If InsereNull And ClsGeneral.xVal(Num) = 0 Then
      SqlNum = "Null"
   Else
      If Trim(Num$) = "" Then
         SqlNum = "0"
      Else
         Num = ClsGeneral.xVal(Num)
         SqlNum = Replace(Replace(Num, ".", ""), ",", ".")
      End If
   End If
End Function
Public Function SqlStr(ByVal Txt As String, Optional InsereNull As Boolean = False, Optional pDbTipo As Integer = 1) As String
   Dim sAux_Antes  As String
   Dim sAux_Depois As String
   
   If InsereNull And Trim(Txt) = "" Then
      SqlStr = "Null"
   Else
      If InStr(Txt, "'") <> 0 Then
         sAux_Antes = Mid(Txt, 1, InStr(Txt, "'") - 1)
         sAux_Depois = Mid(Txt, InStr(Txt, "'") + 1)
         If pDbTipo = 1 Then
            Txt = SqlStr(sAux_Antes) & " + CHAR(39)  + " & SqlStr(sAux_Depois)
         ElseIf pDbTipo = 2 Then
            Txt = SqlStr(sAux_Antes) & " + ""'"" + " & SqlStr(sAux_Depois)
         ElseIf pDbTipo = 0 Then
            Txt = SqlStr(sAux_Antes) & " + ""'"" + " & SqlStr(sAux_Depois)
         Else
            Txt = SqlStr(sAux_Antes) & " + ""'"" + " & SqlStr(sAux_Depois)
         End If
         SqlStr = Txt
      Else
         SqlStr = "'" & Txt & "'"
      End If
   End If
End Function
