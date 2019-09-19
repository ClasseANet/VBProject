VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_Money"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Public Property Let IDLOJA(ByVal vData As Integer)
    mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   If mvarIDLOJA = 0 Then mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Function GerarOFCFromRs(pRs As Object, Optional pPath As String, Optional pFileName As String, Optional bNDoc As Boolean = True, Optional bMsg As Boolean = False) As Boolean
   Dim ODlg  As Object
   Dim sArq As String
   'Dim sArqName As String
   Dim Sql As String
   
   Dim pDSCTransf As String
   Dim pDSCConta As String
   Dim pNumBanco As String
   Dim pNumConta As String
   Dim pTpConta As String
   Dim pDTINI As String
   Dim pDTFIM As String
   Dim pSaldo As String
   
   Dim pTpLan As String
   Dim pTpTransa As String
   Dim pData As String
   Dim pValor As String
   Dim pId As String
   Dim pCheque As String
   Dim pFavorecido As String
   Dim pDesp   As String
   Dim pSubDesp As String
   Dim pMemo As String
   
   On Error GoTo TrataErro
   
   pRs.MoveFirst
   pDTINI = "01/01/9999"
   pDTFIM = "01/01/2000"
   While Not pRs.EOF
      pDTINI = IIf(CDate(pRs("DTVENCIMENTO")) < CDate(pDTINI), pRs("DTVENCIMENTO"), pDTINI)
      pDTFIM = IIf(CDate(pRs("DTVENCIMENTO")) > CDate(pDTFIM), pRs("DTVENCIMENTO"), pDTFIM)
      
      Select Case pRs("TPLAN") & ""
         Case "C": pSaldo = xVal(pSaldo) + xVal(pRs("VALOR") & "")
         Case "D": pSaldo = xVal(pSaldo) - xVal(pRs("VALOR") & "")
         Case Else: pSaldo = xVal(pSaldo) + xVal(pRs("VALOR") & "")
      End Select
      pRs.MoveNext
   Wend
   
   sArq = pPath & pFileName
   If pPath = "" Or pFileName = "" Then
      pPath = ResolvePathName(pPath)
      pFileName = "P3R_" & Format(pDTFIM, "ddmmyy") & ".qif"
      sArq = pPath & pFileName
      If pPath = "" Then
         pPath = mvarSys.PathTmp
         Set ODlg = CriarObjeto("MSComDlg.CommonDialog") 'CreateObject("MSComDlg.CommonDialog")
         sArq = SelecionarArquivo(ODlg, "Arquivos Money/Quicken", sArq, "*.ofc|*.ofc")
      End If
   End If
   pFileName = GetNameFromPath(sArq, pPath)
   
   If Trim(pFileName) = "" Then Exit Function
         
   If ExisteArquivo(sArq) Then
      Call ExcluirArquivo(sArq, False)
   End If
   Open sArq For Output As #1
   pRs.MoveFirst
   '***********
   '* Cabeçalho
   Sql = "Select NUMBANCO, NUMAGENCIA, NUMCONTA, DVCONTA, TPCONTA"
   Sql = Sql & " From FCCORRENTE"
   Sql = Sql & " Where IDLOJA=" & pRs("IDLOJA")
   Sql = Sql & "  And IDCONTA=" & pRs("IDCONTA")
   With mvarSys.xDb
      If .AbreTabela(Sql) Then
         pNumBanco = Trim(.RsAux("NUMBANCO") & "")
         pNumConta = Trim(.RsAux("NUMCONTA") & "")
         If Not IsNull(.RsAux("DVCONTA")) Then
            If Trim(.RsAux("DVCONTA")) <> "" Then
               pNumConta = pNumConta & "-" & .RsAux("DVCONTA")
            End If
         End If
         pTpConta = IIf(Trim(.RsAux("TPCONTA") & "") = "P", "1", "0")
      End If
   End With
      
   Call MontarCabecalhoOFC(pNumBanco, pNumConta, pTpConta, pDTINI, pDTFIM, pSaldo)

   '***********
   '* Detalhe
   pRs.MoveFirst
   While Not pRs.EOF
      pTpLan = pRs("TPLAN") & ""
      pDSCTransf = pRs("DSCTRANSF") & ""
      pTpTransa = pRs("TPTRANSA") & ""
      pData = pRs("DTVENCIMENTO") & ""
      pValor = pRs("VALOR") & ""
      pId = IIf(bNDoc, pRs("NDOC") & "", "")
      pCheque = IIf(bNDoc, pRs("NDOC") & "", "")
      pFavorecido = pRs("FAVORECIDO") & ""
      pDesp = pRs("DSCDESP") & ""
      pSubDesp = pRs("DSCSUBDESP") & ""
      pMemo = pRs("HISTORICO") & IIf(pRs("NDOC") & "" <> "", " (Doc.: " & pRs("NDOC") & ")", "")
           
      Call MontarDetalheOFC(pTpLan, pTpTransa, pData, pValor, pId, pCheque, pFavorecido, pDesp, pSubDesp, pMemo, pDSCTransf)
      pRs.MoveNext
   Wend
   '***********
   '* Rodapé
   Print #1, "  </STMTRS>"
   Print #1, " </ACCTSTMT>"
   Print #1, "</OFC>"
   
   Close #1
   GerarOFCFromRs = True
   If bMsg Then Call ExibirInformacao("Arquivo gerado com sucesso.", "Exportar Movimento")
   Exit Function
TrataErro:
   GerarOFCFromRs = False
   If bMsg Then Call ExibirStop("Erro ao gerar arquivo." & vbNewLine & Err.Number & " - " & Err.Description, "Exportar Movimento")
   On Error Resume Next
   Close #1
End Function
Private Function MontarDetalheOFC(pTipo As String, pTransa As String, pData As String, pValor As String, _
      pId As String, pCheque As String, pFavorecido As String, pDesp As String, _
      pSubDesp As String, pMemo As String, pDSCTransf As String) As Integer
   
   pData = Format(pData, "yyyymmdd")
   If xVal(pValor) < 0 Then pValor = -1 * xVal(pValor)
   pValor = Replace(Replace(ValBr(pValor), ".", ""), ",", ".")
   
   
   Print #1, "  <STMTTRN>"
   Print #1, "   <TRNTYPE>0" '& IIf(Trim(pTipo) = "", "0", pTipo)
   Print #1, "   <DTPOSTED>" & IIf(Trim(pData) = "", "0", pData)
   If pTipo = "0" Or pTipo = "C" Then
      Print #1, "   <TRNAMT>" & pValor
   ElseIf pTipo = "1" Or pTipo = "D" Then
      Print #1, "   <TRNAMT>-" & pValor
   Else
      Print #1, "   <TRNAMT>" & pValor
   End If
   Print #1, "   <FITID>" & IIf(Trim(pId) = "", "0", pId)
   Print #1, "   <CHKNUM>" & IIf(Trim(pCheque) = "", "0", pCheque)
   Print #1, "   <NAME>" & IIf(Trim(pFavorecido) = "", "0", Mid(pFavorecido, 1, 32))
   If Trim(pMemo) <> "" Then
       Print #1, "   <MEMO>" & pMemo
   End If
   Print #1, "  </STMTTRN>"
End Function
Private Function MontarCabecalhoOFC(Optional pNumBanco As String = "001", Optional pNumConta As String = "123", _
        Optional pTipo As String = "2", Optional pDTINI As String, Optional pDTFIM As String, Optional pSaldo As String = "0") As Integer

   pDTINI = Format(pDTINI, "yyyymmdd")
   pDTFIM = Format(pDTFIM, "yyyymmdd")
   pSaldo = Replace(Replace(ValBr(pSaldo), ".", ""), ",", ".")


   Print #1, "<OFC>"
   Print #1, " <DTD>2"
   Print #1, " <CPAGE>1252"
   Print #1, " <ACCTSTMT>"
   Print #1, "  <ACCTFROM>"
   Print #1, "   <BANKID>" & IIf(Trim(pNumBanco) = "", "0", pNumBanco)
   Print #1, "   <ACCTID>" & IIf(Trim(pNumConta) = "", "0", pNumConta)
   Print #1, "   <ACCTTYPE>" & IIf(Trim(pTipo) = "", "0", pTipo)
   Print #1, "  </ACCTFROM>"
   
   Print #1, "  <STMTRS>"
   Print #1, "   <DTSTART>" & IIf(Trim(pDTINI) = "", "0", pDTINI)
   Print #1, "   <DTEND>" & IIf(Trim(pDTFIM) = "", "0", pDTFIM)
   Print #1, "   <LEDGER>" & IIf(Trim(pSaldo) = "", "0", pSaldo)
End Function
Public Function GerarQIFFromRs(pRs As Object, Optional pPath As String, Optional pFileName As String, Optional bNDoc As Boolean = True, Optional bMsg As Boolean = False) As Boolean
   Dim ODlg  As Object
   Dim sArq As String
   'Dim sArqName As String
   Dim Sql As String
   
   Dim pDSCTransf As String
   Dim pDSCConta As String
   Dim pNumBanco As String
   Dim pNumConta As String
   Dim pTpConta As String
   Dim pDTINI As String
   Dim pDTFIM As String
   Dim pSaldo As String
   
   Dim pTpLan As String
   Dim pTpTransa As String
   Dim pData As String
   Dim pValor As String
   Dim pId As String
   Dim pCheque As String
   Dim pFavorecido As String
   Dim pDesp   As String
   Dim pSubDesp As String
   Dim pMemo As String
   
   On Error GoTo TrataErro
   
   pRs.MoveFirst
   pDTINI = "01/01/9999"
   pDTFIM = "01/01/2000"
   While Not pRs.EOF
      pDTINI = IIf(CDate(pRs("DTVENCIMENTO")) < CDate(pDTINI), pRs("DTVENCIMENTO"), pDTINI)
      pDTFIM = IIf(CDate(pRs("DTVENCIMENTO")) > CDate(pDTFIM), pRs("DTVENCIMENTO"), pDTFIM)

      Select Case pRs("TPLAN") & ""
         Case "C": pSaldo = xVal(pSaldo) + xVal(pRs("VALOR") & "")
         Case "D": pSaldo = xVal(pSaldo) - xVal(pRs("VALOR") & "")
         Case Else: pSaldo = xVal(pSaldo) + xVal(pRs("VALOR") & "")
      End Select
      pRs.MoveNext
   Wend
   
   sArq = pPath & pFileName
   If pPath = "" Or pFileName = "" Then
      pPath = ResolvePathName(pPath)
      pFileName = "P3R_" & Format(pDTFIM, "ddmmyy") & ".qif"
      sArq = pPath & pFileName
      If pPath = "" Then
         pPath = mvarSys.PathTmp
         Set ODlg = CriarObjeto("MSComDlg.CommonDialog") 'CreateObject("MSComDlg.CommonDialog")
         sArq = SelecionarArquivo(ODlg, "Arquivos Money/Quicken", sArq, "*.ofc|*.ofc")
      End If
   End If
   pFileName = GetNameFromPath(sArq, pPath)
   
   If Trim(pFileName) = "" Then Exit Function
         
   If ExisteArquivo(sArq) Then
      Call ExcluirArquivo(sArq, False)
   End If
   Open sArq For Output As #1
   pRs.MoveFirst
   '***********
   '* Cabeçalho
   Sql = "Select DSCCONTA, NUMBANCO, NUMAGENCIA, NUMCONTA, DVCONTA"
   Sql = Sql & ", TPCONTA"
   Sql = Sql & " From FCCORRENTE"
   Sql = Sql & " Where IDLOJA=" & pRs("IDLOJA")
   Sql = Sql & "  And IDCONTA=" & pRs("IDCONTA")
   With mvarSys.xDb
      If .AbreTabela(Sql) Then
         pDSCConta = Trim(.RsAux("DSCCONTA") & "")
         pNumBanco = Trim(.RsAux("NUMBANCO") & "")
         pNumConta = Trim(.RsAux("NUMCONTA") & "")
         If Not IsNull(.RsAux("DVCONTA")) Then
            If Trim(.RsAux("DVCONTA")) <> "" Then
               pNumConta = pNumConta & "-" & .RsAux("DVCONTA")
            End If
         End If
         pTpConta = IIf(Trim(.RsAux("TPCONTA") & "") = "P", "1", "0")
      End If
   End With
      
   Call MontarCabecalhoQIF(pDSCConta, pNumBanco, pNumConta, pTpConta, pDTINI, pDTFIM, pSaldo)

   '***********
   '* Detalhe
   pRs.MoveFirst
   While Not pRs.EOF
      pTpLan = pRs("TPLAN") & ""
      pDSCTransf = pRs("DSCTRANSF") & ""
      pTpTransa = pRs("TPTRANSA") & ""
      pData = pRs("DTVENCIMENTO") & ""
      pValor = pRs("VALOR") & ""
      pId = IIf(bNDoc, pRs("NDOC") & "", "")
      pCheque = IIf(bNDoc, pRs("NDOC") & "", "")
      pFavorecido = pRs("FAVORECIDO") & ""
      pDesp = pRs("DSCDESP") & ""
      pSubDesp = pRs("DSCSUBDESP") & ""
      pMemo = pRs("HISTORICO") & IIf(pRs("NDOC") & "" <> "", " (Doc.: " & pRs("NDOC") & ")", "")
   
      Call MontarDetalheQIF(pTpLan, pTpTransa, pData, pValor, pId, pCheque, pFavorecido, pDesp, pSubDesp, pMemo, pDSCTransf)
      pRs.MoveNext
   Wend
   '***********
   '* Rodapé
   Close #1
   GerarQIFFromRs = True
   If bMsg Then Call ExibirInformacao("Arquivo gerado com sucesso.", "Exportar Movimento")
   Exit Function
TrataErro:
   GerarQIFFromRs = False
   If bMsg Then Call ExibirStop("Erro ao gerar arquivo." & vbNewLine & Err.Number & " - " & Err.Description, "Exportar Movimento")
   On Error Resume Next
   Close #1
End Function
Private Function MontarCabecalhoQIF(Optional pDSCConta As String = "", Optional pNumBanco As String = "001", Optional pNumConta As String = "123", Optional pTipo As String = "2", Optional pDTINI As String, Optional pDTFIM As String, Optional pSaldo As String = "0") As Integer

'   pDTINI = Format(pDTINI, "dd/mm/yy")
'   pDTFIM = Format(pDTFIM, "dd/mm/yy")
'   pSaldo = Replace(Replace(ValBr(pSaldo), ".", ""), ",", ".")

   Print #1, "!Type:Bank"
'   Print #1, "D" & pDTINI
'   Print #1, "T0.00"
'   Print #1, "CX"
'   Print #1, "POpening Balance"
'   If pDSCConta <> "" Then
'      Print #1, "L[" & pDSCConta & "]"
'   End If
'   Print #1, "^"
End Function

Private Function MontarDetalheQIF(pTipo As String, pTransa As String, pData As String, pValor As String, _
      pId As String, pCheque As String, pFavorecido As String, pDesp As String, _
      pSubDesp As String, pMemo As String, pDSCTransf As String) As Integer
   
   pData = Format(pData, "dd/mm/yy")
   If xVal(pValor) < 0 Then pValor = -1 * xVal(pValor)
   pValor = Replace(Replace(ValBr(pValor), ".", ""), ",", ".")
   
   Print #1, "D" & pData
   Print #1, "T" & IIf(pTipo = "D", "-", "") & pValor
   Print #1, "N" & pId
   Print #1, "P" & pFavorecido
   If pTransa = "T" Then
      Print #1, "L[" & pDSCTransf & "]"
   Else
      If Trim(pDesp) <> "" Then
         Print #1, "L" & pDesp & IIf(Trim(pSubDesp) <> "", ":" & Trim(pSubDesp), "")
      End If
   End If
   If Trim(pMemo) <> "" Then
      Print #1, "M" & pMemo
   End If
   
   Print #1, "^"
End Function


