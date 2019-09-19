VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "General"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function Between(Vl, Min, Max) As Boolean
   Between = False
   If Vl >= Min And Vl <= Max Then
      Between = True
   End If
End Function
Public Function BinToBoolean(nVal As Long) As Boolean
    If nVal = 0 Then
        BinToBoolean = False
    Else
        BinToBoolean = True
    End If
End Function
Public Function BooleanToBin(bVal As Boolean) As Long
    If bVal Then
        BooleanToBin = 1
    Else
        BooleanToBin = 0
    End If
End Function
Public Function BuscaPeriodo(ByVal pSemana As String, ByRef pDataIni As Date, ByRef pDataFim As Date, Optional pExibeMensagem As Boolean = True) As Boolean
   Dim sANO          As String
   Dim nSEMANA       As Integer
   Dim dPriDiaSemana As Date
   Dim dData         As Date
   Dim nFator        As Integer
   Dim nUltSem       As Integer
   Dim nAux          As Integer
   Dim sAux          As Integer

   On Error GoTo Fim:
   
   pSemana = StrZero(pSemana, 4)
   nSEMANA = Mid(pSemana, 3, 2)
   sANO = "2" & StrZero(Mid(pSemana, 1, 2), 3)
      
   '********
   '* Determina Nº da última semana do ano
   dData = CDate("31/12/" & sANO)
   nAux = "2" & StrZero(Mid(BuscaSemana(dData), 1, 2), 3)
   nUltSem = Mid(BuscaSemana(dData), 3, 2)
   While nAux <> Year(dData)
      dData = DateAdd("d", -1, dData)
      nAux = "2" & StrZero(Mid(BuscaSemana(dData), 1, 2), 3)
      nUltSem = Mid(BuscaSemana(dData), 3, 2)
   Wend
   
   If nSEMANA <= 0 Or nSEMANA > nUltSem Then
      If pExibeMensagem = True Then MsgBox "Semana inválida.", vbInformation
      BuscaPeriodo = False
      Exit Function
   End If
   
   '********
   '* Determina Nº da última semana do Ano Anterior
   dData = CDate("31/12/" & CInt(sANO) - 1)
   nAux = "2" & StrZero(Mid(BuscaSemana(dData), 1, 2), 3)
   nUltSem = Mid(BuscaSemana(dData), 3, 2)
   While nAux <> Year(dData)
      dData = DateAdd("d", -1, dData)
      nAux = "2" & StrZero(Mid(BuscaSemana(dData), 1, 2), 3)
      nUltSem = Mid(BuscaSemana(dData), 3, 2)
   Wend
   '****
   '* Percorre a última semana da ano anterior até o 1º dia da 1º semana do ano desejado
   nAux = CInt(Mid(BuscaSemana(dData), 3, 2))
   While nAux <> 1
      dData = DateAdd("d", 1, dData)
      nAux = CInt(Mid(BuscaSemana(dData), 3, 2))
   Wend
   dPriDiaSemana = dData
   
   
   If DateAdd("ww", nSEMANA - 1, dPriDiaSemana) > CDate("31/12/" & sANO) Then
      If pExibeMensagem = True Then MsgBox "Semana inválida.", vbInformation
      BuscaPeriodo = False
      Exit Function
   End If
   
   pDataIni = DateAdd("ww", nSEMANA - 1, dPriDiaSemana)
   pDataFim = pDataIni + 6
      
   BuscaPeriodo = True
   Exit Function

Fim:
End Function
'************************************************************
'* FUNÇÃO QUE DETERMINA A SEMANA DE TRABALHO                *
'* Baseado nas seguintes regras da ABNT NBR9577:            *
'*    1 - A semana possui 7 dias                            *
'*    2 - A semana vai de Segunda a Domingo                 *
'*    3 - A semana 01 é a 1º semana que possui pelo menos 4 *
'*        dias do Ano novo                                  *
'*    4 - Sendo Segunda-Feira o 1º dia da semana, a semana  *
'*        01 é a semana que possui a 1º Quinta-Feira do Ano *
'*                                                          *
'************************************************************
Public Function BuscaSemana(pData As Date) As String
   Dim nFator        As Integer
   Dim dDataAux      As Date
   Dim nAux          As Integer
   Dim Dia1Semana    As Integer
   Dim QtdMin        As Integer
    
   Dia1Semana = vbMonday
   QtdMin = 4
   
   If IsDate(pData) = False Then
      MsgBox "Data inválida!", vbInformation
      Exit Function
   End If
         
   '* Determina se Dia 1º esta na 1ºSemana.
   '* Se não estiver, deveremos subttrair 1 da função padrão do VB
   nAux = 0
   dDataAux = CDate("01/01/" & Year(pData))
   If 7 - Weekday(dDataAux, Dia1Semana) + 1 < QtdMin Then nAux = 1
   
   '* Se estamos na 1º semana verificamos o último dia do ano anterior
   '* e determinamos se esta data pertence ao ano anterior
   If Format(pData, "WW", Dia1Semana) - nAux <= 0 Then
      dDataAux = CDate("31/12/" & Year(pData) - 1)
      BuscaSemana = BuscaSemana(dDataAux)

   '* Se estamos na última semana, definimos o dia 1º do ano seguinte
   '* e determinamos se a data pertence ao ano seguinte.
   ElseIf Format(pData, "WW", Dia1Semana) > 52 And Day(pData) > 31 - (7 - QtdMin) Then
      dDataAux = CDate("01/01/" & Year(pData) + 1)
      If 7 - Weekday(dDataAux, Dia1Semana) + 1 < QtdMin Or Weekday(dDataAux, Dia1Semana) = 1 Then
         BuscaSemana = Format(pData, "YY") & StrZero(Format(pData, "WW", Dia1Semana) - 1, 2)
      Else
         BuscaSemana = Format(dDataAux, "YY") & StrZero(Format(dDataAux, "WW", Dia1Semana), 2)
      End If
   Else
      BuscaSemana = Format(pData, "YY") & StrZero(Format(pData, "WW", Dia1Semana) - nAux, 2)
   End If
End Function
Public Function Calcular_Digito_CPF(wrk_fatr, wrk_limt, DigitoCPF)
   Dim Soma, digito, Ind As Integer
   
   Soma = 0
   Ind = 1
   While Ind <= wrk_limt
       Soma = Soma + DigitoCPF(Ind) * wrk_fatr
       wrk_fatr = wrk_fatr - 1
       Ind = Ind + 1
   Wend
   digito = Soma - (Int(Soma / 11) * 11)
   If digito = 0 Or digito = 1 Then
       digito = 0
   Else
       digito = 11 - digito
   End If
   
   Calcular_Digito_CPF = digito
End Function
Public Function DiaDaSemana(pDate As Date, Optional bPortugues = True) As String
   If bPortugues Then
      Select Case Weekday(pDate)
         Case 1: DiaDaSemana = "Domingo"
         Case 2: DiaDaSemana = "Segunda-Feira"
         Case 3: DiaDaSemana = "Terça-Feira"
         Case 4: DiaDaSemana = "Quarta-Feira"
         Case 5: DiaDaSemana = "Quinta-Feira"
         Case 6: DiaDaSemana = "Sexta-Feira"
         Case 7: DiaDaSemana = "Sábado"
      End Select
   Else
      DiaDaSemana = WeekdayName(Weekday(pDate))
   End If
End Function
Public Function ExisteIndex(pObj As Object, pIndex As Integer) As Boolean
   Dim Value
   On Error Resume Next
   
   Value = pObj(pIndex)
   If Err = 449 Then
      Err = 0
      Set Value = pObj(pIndex)
   End If
   
   ExisteIndex = (Err = 0)
End Function
Public Function ExisteItem(pColl As Collection, pItem As String) As Boolean
   Dim x  As Variant
   
   Err = 0
   On Error Resume Next
   x = pColl(pItem)
   If Err <> 0 Then
      Err = 0
      Set x = pColl(pItem)
   End If
   If (Err = 0) Then
      ExisteItem = (pColl.Count > 0)
   Else
      ExisteItem = (Err = 0)
   End If
End Function
'Public Function eFeriado(xConn As Object, ByVal pData As String, Optional pBanco As String = "BANCO_3R", Optional pTabela As String = "TB_GFERIADO") As Boolean
Public Function eFeriado(xConn As Object, ByVal pData As String, Optional pTabela As String = "GFERIADO", Optional pCampoDT As String = "DATA", Optional pCampoEscopo As String = "ESCOPO") As Boolean
   Dim Tb As Object
   Dim Sql As String
   
   If Not IsDate(pData) Then
      Exit Function
   End If
   
   pData = Format(pData, "dd/mm/yyyy")
   Sql = "Select * "
   Sql = Sql & " From " & pTabela
   Sql = Sql & " Where Day(" + pCampoDT + ")=" & Day(CDate(pData))
   Sql = Sql & " And Month(" + pCampoDT + ")=" & Month(CDate(pData))
   Sql = Sql & " And Year(" + pCampoDT + ")=" & Year(CDate(pData))
   If xConn.AbreTabela(Sql) Then
      eFeriado = True
   Else
      Sql = "Select * "
      Sql = Sql & " From " & pTabela
      Sql = Sql & " Where Day(" + pCampoDT + ")=" & Day(CDate(pData))
      Sql = Sql & " And Month(" + pCampoDT + ")=" & Month(CDate(pData))
      Sql = Sql & " And " + pCampoEscopo + "=1"
      If xConn.AbreTabela(Sql) Then
         eFeriado = True
      End If
   End If
   
'   Set Tb = CriarObjeto(pBanco & "." & pTabela)
'   With Tb
'      Set .xdb = xConn
'      If .Pesquisar(Ch_DATA:=pData) Then
'         eFeriado = True
'      Else
'         Sql = " ESCOPO=1"
'         Sql = Sql & " And Day(DATA)=" & Day(CDate(pData))
'         Sql = Sql & " And Month(DATA)=" & Month(CDate(pData))
'         If .Pesquisar(Ch_Where:=Sql) Then
'            eFeriado = True
'         End If
'      End If
'   End With
End Function
Public Sub ExecuteLink(ByVal sLinkTo As String)

'* Execute o link to http://www.XXX.com.br/
'* (if possible) - or the new 'mailto:XXX@yyyy.com.br'

    On Error Resume Next
    
    Dim lRet As Long
    Dim lOldCursor As Long
    
    lOldCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    lRet = ShellExecute(0, "open", sLinkTo, "", vbNull, 1) 'SW_SHOWNORMAL)
    
    If lRet >= 0 And lRet <= 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error Opening Link to " & sLinkTo & vbCrLf & vbCrLf & Err.LastDllError, , "frmAbout::ExecuteLink"
    End If
    Screen.MousePointer = vbDefault
    
End Sub
Public Function ExtractResData(Id, Tipo, Arquivo As String, Optional pFileBuf) As Boolean
   Dim nInt As Integer
   Dim byteFileBuf() As Byte 'This must be byte rather than String, so no Unicode conversion takes place
   Dim nVez As Integer
   Dim sPath   As String
   
   On Error GoTo Fim
   
   Call ClsAmbiente.GetNameFromPath(Arquivo, sPath)
   If sPath <> "" Then
      Call ClsAmbiente.CriarDiretorio(sPath)
   End If
   Call ClsAmbiente.ExcluirArquivo(Arquivo)
   
   nInt = FreeFile
   Open Arquivo$ For Binary Access Write As nInt
      If IsMissing(pFileBuf) Then
         byteFileBuf = LoadResData(Id, Tipo)
      End If
      Put nInt, , byteFileBuf
   GoTo Saida
Fim:
   nVez = nVez + 1
   If nVez < 5 Then
      Resume
   Else
      Resume Next
   End If
   
Saida:
    Close nInt
    Err = 0
    ExtractResData = ClsAmbiente.ExisteArquivo(Arquivo$)
    Exit Function
End Function
Public Function FormatarData(pStrDate As String, Optional pNull As Boolean = False) As String
   Dim sText As String
   Dim nPos  As Integer
   Dim nPos2  As Integer

   sText = pStrDate
   
   If Not pNull Then
      If Trim(sText) = "" Then sText = Format(Now(), "dd/mm/yyyy")
   End If
   
   nPos = InStr(sText, "/")
   nPos2 = InStr(nPos + 1, sText, "/")
   If nPos > 0 Then
      If nPos <= 2 Then
         sText = StrZero(Mid(sText, 1, nPos - 1), 2) & Mid(sText, nPos)
         nPos = InStr(sText, "/")
         nPos2 = InStr(nPos + 1, sText, "/")
      End If
      If nPos2 = 0 Then
         sText = sText & Format(Now(), "/yyyy")
         nPos = InStr(sText, "/")
         nPos2 = InStr(nPos + 1, sText, "/")
      End If
      If nPos2 - nPos <= 2 And nPos2 - nPos > 0 Then
         sText = Mid(sText, 1, nPos) & StrZero(Mid(sText, nPos + 1, nPos2 - nPos - 1), 2) & Mid(sText, nPos2)
      End If
   End If
   
   sText = Replace(sText, "/", "")
   If Len(sText) <= 2 Then
      sText = StrZero(sText, 2) + Format(Now(), "/mm/yyyy")
   ElseIf Len(sText) = 3 Then
      If Val(Mid(sText, 2, 2)) > 12 Then
         sText = Mid(sText, 1, 2) + "/" + StrZero(Mid(sText, 3, 2), 2) + Format(Now(), "/yyyy")
      Else
         sText = StrZero(Mid(sText, 1, 1), 2) + "/" + StrZero(Mid(sText, 2, 2), 2) + Format(Now(), "/yyyy")
      End If
   ElseIf Len(sText) = 4 Then
      sText = Mid(sText, 1, 2) + "/" + StrZero(Mid(sText, 3, 2), 2) + Format(Now(), "/yyyy")
   ElseIf Len(sText) = 5 Then
      sText = Mid(sText, 1, 2) + "/" + StrZero(Mid(sText, 3, 2), 2) + Mid(Format(Now(), "/yyyy"), 1, 3) & StrZero(Mid(sText, 5), 2)
   ElseIf Len(sText) > 5 Then
      If Mid(sText, 5, 4) >= Left(Year(Now), Len(Mid(sText, 5, 4))) Then
         sText = Mid(sText, 1, 2) + "/" + Mid(sText, 3, 2) + "/" + Left(Left(Year(Now) - 100, 4 - Len(Mid(sText, 5, 4))) + Mid(sText, 5, 4), 4)
      Else
         sText = Mid(sText, 1, 2) + "/" + Mid(sText, 3, 2) + "/" + Left(Left(Year(Now), 4 - Len(Mid(sText, 5, 4))) + Mid(sText, 5, 4), 4)
      End If
   End If
   If IsDate(sText) Then
      FormatarData = sText
   End If
End Function
Public Function FormatarHora(pStrHour As String, Optional pSegundo As Boolean = False) As String
   Dim sHora As String
   
   sHora = Replace(pStrHour, ":", "")
   If Len(sHora) <= 2 Then
      If Val(sHora) < 23 Then
         sHora = StrZero(sHora, 2) & ":00"
      ElseIf Val(sHora) >= 24 And Val(sHora) < 60 Then
         sHora = "00:" & StrZero(sHora, 2)
      ElseIf Val(sHora) >= 60 Then
         sHora = StrZero(Mid(sHora, 1, 1), 2) & ":" & StrZero(Mid(sHora, 1, 1), 2)
      End If
   ElseIf Len(sHora) = 3 Then
      sHora = StrZero(Mid(sHora, 1, 1), 2) & ":" & Mid(sHora, 2, 2)
   ElseIf Len(sHora) >= 4 Then
      sHora = Mid(sHora, 1, 2) & ":" & Mid(sHora, 3, 2)
   End If
   If pSegundo Then
      If Len(Replace(pStrHour, ":", "")) = 5 Then
         sHora = sHora & ":" & StrZero(Mid(Replace(pStrHour, ":", ""), 5, 1), 2)
      ElseIf Len(Replace(pStrHour, ":", "")) = 6 Then
         sHora = sHora & ":" & Mid(Replace(pStrHour, ":", ""), 5, 2)
      Else
         sHora = sHora & ":00"
      End If
   End If
   
   If IsDate(sHora) Then
      FormatarHora = sHora
   End If
End Function
Public Function FormatarNome(pNome As String, Optional Somente1Maiuscula As Boolean) As String
   Dim Palavras As Collection
   Dim sNome As String
   Dim sAux
   
   Set Palavras = GetWords(pNome)
   
   For Each sAux In Palavras
      sAux = Trim(sAux)
      If Mid(sAux, 1, 1) <> "" And Len(sAux) > 3 Then
         If Asc(Mid(sAux, 1, 1)) < 123 Then
            sAux = UCase(Mid(sAux, 1, 1)) & IIf(Somente1Maiuscula, LCase(Mid(sAux, 2)), Mid(sAux, 2))
         End If
      End If
      sNome = sNome & " " & sAux
   Next
   FormatarNome = Trim(sNome)
End Function
Public Function GetTag(ByRef pControle As Variant, ByVal pNome As String, Optional pPadrao As String) As String
   Dim PosIni  As Long
   Dim PosFim  As Long
   Dim StrTAG  As String
   Dim i       As Integer
   
   On Error GoTo Saida
   
   pNome = "|" & Trim(pNome) & "="
   
   If UCase(TypeName(pControle)) = "STRING" Then
      StrTAG = pControle
   Else
      StrTAG = pControle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(pNome))
   If PosIni > 0 Then
      PosIni = PosIni + Len(Trim(pNome))
      PosFim = InStr(PosIni, StrTAG, "|")
      i = 0
      While Mid(StrTAG, PosIni + i, 1) = "|"
         i = i + 1
      Wend
      If i > 0 Then
         PosFim = InStr(PosIni + (i - 1), StrTAG, "|")
      End If
      PosFim = IIf(PosFim = 0, Len(StrTAG), PosFim - 1)
      StrTAG = Mid(StrTAG, PosIni, PosFim - PosIni + 1)
   Else
      StrTAG = ""
   End If
   GetTag = StrTAG
Saida:
   If StrTAG = "" Then
      GetTag = pPadrao
   End If
End Function
Public Function GetWords(ByVal StrLinha As String) As Collection
   Dim MyColl As Collection, Palavra As String

   
   Set MyColl = New Collection
   
   StrLinha$ = Trim(StrLinha$)
   While Trim(StrLinha$) <> ""
      Palavra = RichWordOver(StrLinha$, 0, 0, 1)
      If Trim(Palavra) = "" Then
         Palavra = ","
      Else
         MyColl.Add Palavra
      End If
      StrLinha$ = Trim(Mid(StrLinha$, Len(Palavra$) + 1))
   Wend
   Set GetWords = New Collection
   Set GetWords = MyColl
End Function
Public Function GetWords_AndOR(pTexto As String, Optional ByRef Palavras_And As Collection, Optional ByRef Palavras_Or As Collection, Optional pCampo) As String
   Dim PosAnd  As Integer
   Dim PosE    As Integer
   Dim PosOR   As Integer
   Dim PosOU   As Integer
   Dim Pos     As Integer
   Dim sAux    As String
   Dim n       As Variant
   Dim Sql     As String

   sAux = UCase(pTexto)
   PosOR = InStr(sAux, " OU ")
   If InStr(sAux, " OR ") <> 0 Then
      PosOR = IIf(InStr(sAux, " OR ") >= PosAnd, PosAnd, InStr(sAux, " OR "))
   End If
   PosAnd = InStr(sAux, " E ")
   If InStr(sAux, " AND ") <> 0 Then
      PosAnd = IIf(InStr(sAux, " AND ") >= PosAnd, PosAnd, InStr(sAux, " AND "))
   End If
   
   If PosOR <> 0 And PosAnd <> 0 Then
      MsgBox "Consulta Inválida.", vbExclamation, App.ProductName
      Exit Function
   End If
   
   sAux = UCase(Replace(pTexto, vbNewLine, " "))
   PosE = InStr(sAux, " E ")
   PosAnd = InStr(sAux, " AND ")
   If PosAnd <> 0 And PosE <> 0 Then
      Pos = IIf(PosAnd > PosE, PosE, PosAnd)
   Else
      Pos = PosAnd + PosE
   End If
   Set Palavras_And = New Collection
   If PosAnd = 0 And PosE = 0 Then
      Palavras_And.Add Trim(sAux)
   Else
      While Pos > 0
         Palavras_And.Add Trim(Mid(sAux, 1, Pos))
         If Pos = PosAnd Then
            sAux = Trim(Mid(sAux, Pos + 4)) + " "
         ElseIf Pos = PosE Then
            sAux = Trim(Mid(sAux, Pos + 2)) + " "
         End If
         PosE = InStr(sAux, " E ")
         PosAnd = InStr(sAux, " AND ")
         If PosAnd <> 0 And PosE <> 0 Then
            Pos = IIf(PosAnd > PosE, PosE, PosAnd)
         Else
            Pos = PosAnd + PosE
         End If
      Wend
      If Trim(sAux) <> "" Then
         Palavras_And.Add Trim(sAux)
      End If
   End If
   sAux = UCase(Replace(pTexto, vbNewLine, " "))
   PosOU = InStr(sAux, " OU ")
   PosOR = InStr(sAux, " OR ")
   If PosOR <> 0 And PosOU <> 0 Then
      Pos = IIf(PosOR > PosOU, PosOU, PosOR)
   Else
      Pos = PosOR + PosOU
   End If
   Set Palavras_Or = New Collection
   If PosOR <> 0 Or PosOU <> 0 Then
      While Pos > 0
         Palavras_And.Add Trim(Mid(sAux, 1, Pos))
         If Pos = PosOR Then
            sAux = Trim(Mid(sAux, Pos + 3)) + " "
         ElseIf Pos = PosOU Then
            sAux = Trim(Mid(sAux, Pos + 3)) + " "
         End If
         PosOU = InStr(sAux, " OU ")
         PosOR = InStr(sAux, " OR ")
         If PosOR <> 0 And PosOU <> 0 Then
            Pos = IIf(PosOR > PosOU, PosOU, PosOR)
         Else
            Pos = PosOR + PosOU
         End If
      Wend
      If Trim(sAux) <> "" Then
         Palavras_Or.Add Trim(sAux)
      End If
   End If
   
   Sql = ""
   If Not IsMissing(pCampo) Then
      If Palavras_And.Count > 0 Then
         Sql = Sql & " And ( " & pCampo
         For Each n In Palavras_And
            Sql = Sql & " like '%" & n & "%'"
            Sql = Sql & " AND " & pCampo
         Next
         Sql = Mid(Sql, 1, Len(Sql) - Len(" AND " & pCampo)) & ")"
      ElseIf Palavras_Or.Count > 0 Then
         Sql = Sql & " And ( " & pCampo
         For Each n In Palavras_And
            Sql = Sql & " like '%" & n & "%'"
            Sql = Sql & " OR " & pCampo
         Next
         Sql = Mid(Sql, 1, Len(Sql) - Len(" OR " & pCampo)) & ")"
      End If
   End If
   GetWords_AndOR = Sql
End Function
Public Function InputBoxPassword(prompt, Optional Title, Optional Default) As String
   Const NV_INPUTBOX = &H5000&
   Dim Ret As String
   
   SetTimer 0, NV_INPUTBOX, 10, AddressOf TimerProc
   Ret = InputBox(prompt) ', Title, Default)
   InputBoxPassword = Ret
End Function
Public Function InArray(Valor As Variant, VETOR As Variant) As Boolean
   Dim j As Variant
   InArray = False
   For Each j In VETOR
       If Valor = j Then
         InArray = True
         Exit For
      End If
   Next
End Function
Public Function isAlfaNum(ch As String) As Boolean
   isAlfaNum = (ch >= "0" And ch <= "9") Or (UCase(ch) >= "A" And UCase(ch) <= "Z")
   isAlfaNum = isAlfaNum Or (UCase(ch) >= "À" And UCase(ch) <= "Ý")    '* Or (Ch = "_")
End Function
Public Sub LimparTela(frm As Object)
   Dim i As Integer
   On Error Resume Next
   
   For i% = 0 To frm.Controls.Count - 1
      
      Select Case UCase(TypeName(frm.Controls(i)))
         
         Case "TEXTBOX", "FLATEDIT"
            frm.Controls(i) = ""
         
         Case "MASKEDBOX"
            Dim sMask As String
            sMask = frm.Controls(i).Mask
            frm.Controls(i).Mask = ""
            frm.Controls(i).Text = ""
            frm.Controls(i).Mask = sMask
         
         Case "LABEL"
            If frm.Controls(i).BorderStyle = 1 Then
               frm.Controls(i) = ""
            End If
         
         Case "OPTIONBUTTON":
            If frm.Controls(i).Index = 0 Then
               frm.Controls(i).Value = True
            End If
         
         Case "COMBOBOX"
           If frm.Controls(i).ListCount > 0 Then
              frm.Controls(i).ListIndex = 0
           Else
              frm.Controls(i).ListIndex = -1
           End If
           
         Case "CHECKBOX"
            frm.Controls(i).Value = 0
      End Select
   Next
End Sub
Public Function RichWordOver(ByVal RchTxt As Variant, x As Single, y As Single, Optional Posicao = 1, Optional VerifDclImplicta = True) As String
   Dim pt As PointAPI
   Dim Pos As Integer
   Dim ch As String
   Dim StartPos As Integer, EndPos As Integer
   Dim Txt As String, TxtLen As Integer
   
   Dim LineCount As Single, CurrLinePos As Single, OverAllCursorPos As Single
   Dim ChrsBeforeCurrLine As Single, CurrLineLen As Single, CurrLineCursorPos As Single
   

   ' Convert the position to pixels.
   
   ' Get the character number
   If x = 0 And y = 0 Then
      If IsMissing(Posicao) Then
         LineCount = SendMessageLong(RchTxt.hwnd, EM_GETLINECOUNT, 0, 0&)
         OverAllCursorPos = SendMessageLong(RchTxt.hwnd, EM_GETSEL, 0, 0&) \ &H10000
         CurrLinePos = SendMessageLong(RchTxt.hwnd, EM_LINEFROMCHAR, OverAllCursorPos, 0&)
         ChrsBeforeCurrLine = SendMessageLong(RchTxt.hwnd, EM_LINEINDEX, CurrLinePos, 0&)
         CurrLineLen = SendMessageLong(RchTxt.hwnd, EM_LINELENGTH, -1, 0&)
         CurrLineCursorPos = OverAllCursorPos + 1 - ChrsBeforeCurrLine
         Pos = OverAllCursorPos - 2 * (CurrLineCursorPos)
         If Pos < 0 Then Exit Function
      Else
         Pos = Posicao
      End If
   
   Else
      pt.x = x \ Screen.TwipsPerPixelX
      pt.y = y \ Screen.TwipsPerPixelY
      Pos = SendMessage(RchTxt.hwnd, EM_CHARFROMPOS, 0&, pt)
   End If
      
   If Pos <= 0 Then Exit Function

   ' Find the start of the word.
   If TypeName(RchTxt) = "String" Then
      Txt = RchTxt
   Else
      Txt = RchTxt.Text
   End If
   
   Posicao = Pos
   Do While Trim(Mid$(Txt, Pos, 1)) = ""
      Pos = Pos + 1
      If Pos > Len(Txt) Then
         Pos = Posicao
         Exit Do
      End If
   Loop
   For StartPos = Pos To 1 Step -1
      ch = Mid$(Txt, StartPos, 1)
      ' Allow digits, letters, and underscores.
      If Not isAlfaNum(ch) Then
         Exit For
      End If
   Next StartPos
   StartPos = StartPos + 1

   ' Find the end of the word.
   Dim Dcl_Implicita
   TxtLen = Len(Txt)
   For EndPos = Pos To TxtLen
      ch = Mid$(Txt, EndPos, 1)
      ' Allow digits, letters, and underscores.
      If Not (isAlfaNum(ch) Or ch = "_") Then
         If VerifDclImplicta Then
           'If Not VarWithDclImplicit(Ch) Then
           If InStr(Right(Trim(ch), 1), "%!&$#") = 0 Then
               Exit For
            End If
         Else
            Exit For
         End If
      End If
   Next EndPos
   EndPos = EndPos - 1

   If StartPos <= EndPos Then
      RichWordOver = Mid$(Txt, StartPos, EndPos - StartPos + 1)
   End If
End Function
Public Sub RetiraPreposicao(ByRef pString As String, Optional ByRef pClString As Collection)
   Dim n
   Dim sNvString        As String
   Dim sPalavras        As String
   Dim clPalavras       As Collection
   Dim clPreposicoes    As Collection
   Dim clNvString       As Collection
   
   On Error Resume Next
   Set clPreposicoes = New Collection
   
   clPreposicoes.Add "A", "A"
   clPreposicoes.Add "AS", "AS"
   clPreposicoes.Add "E", "E"
   clPreposicoes.Add "O", "O"
   clPreposicoes.Add "OS", "OS"
   clPreposicoes.Add "DA", "DA"
   clPreposicoes.Add "DAS", "DAS"
   clPreposicoes.Add "DE", "DE"
   clPreposicoes.Add "DO", "DO"
   clPreposicoes.Add "DOS", "DOS"
   
   clPreposicoes.Add "DUM", "DUM"
   clPreposicoes.Add "DUNS", "DUNS"
   clPreposicoes.Add "DUMA", "DUMA"
   clPreposicoes.Add "DUMAS", "DUMAS"
   clPreposicoes.Add "DELE", "DELE"
   clPreposicoes.Add "DELES", "DELES"
   clPreposicoes.Add "DELA", "DELA"
   clPreposicoes.Add "DELAS", "DELAS"
   clPreposicoes.Add "DESTE", "DESTE"
   clPreposicoes.Add "DESTES", "DESTES"
   clPreposicoes.Add "DESTA", "DESTA"
   clPreposicoes.Add "DESTAS", "DESTAS"
   clPreposicoes.Add "DESSE", "DESSE"
   clPreposicoes.Add "DESSES", "DESSES"
   clPreposicoes.Add "DESSA", "DESSA"
   clPreposicoes.Add "DESSAS", "DESSAS"
   clPreposicoes.Add "AQUELE", "AQUELE"
   clPreposicoes.Add "AQUELES", "AQUELES"
   clPreposicoes.Add "AQUELA", "AQUELA"
   clPreposicoes.Add "AQUELAS", "AQUELAS"
   clPreposicoes.Add "DAQUELE", "DAQUELE"
   clPreposicoes.Add "DAQUELES", "DAQUELES"
   clPreposicoes.Add "DAQUELA", "DAQUELA"
   clPreposicoes.Add "DAQUELAS", "DAQUELAS"
   clPreposicoes.Add "DISSO", "DISSO"
   clPreposicoes.Add "DISTO", "DISTO"
   clPreposicoes.Add "DAQUILO", "DAQUILO"
   clPreposicoes.Add "DAI", "DAI"
   clPreposicoes.Add "DALI", "DALI"
   
   clPreposicoes.Add "NESSE", "NESSE"
   clPreposicoes.Add "NESSES", "NESSES"
   clPreposicoes.Add "NESSA", "NESSA"
   clPreposicoes.Add "NESSAS", "NESSAS"
   clPreposicoes.Add "NESTE", "NESTE"
   clPreposicoes.Add "NESTES", "NESTES"
   clPreposicoes.Add "NESTA", "NESTA"
   clPreposicoes.Add "NESTAS", "NESTAS"
   clPreposicoes.Add "NAQUELE", "NAQUELE"
   clPreposicoes.Add "NAQUELES", "NAQUELES"
   clPreposicoes.Add "NAQUELA", "NAQUELA"
   clPreposicoes.Add "NAQUELAS", "NAQUELAS"
   clPreposicoes.Add "NISSO", "NISSO"
   clPreposicoes.Add "NISTO", "NISTO"
   clPreposicoes.Add "NAQUILO", "NAQUILO"
   
   clPreposicoes.Add "EM", "EM"
   clPreposicoes.Add "NO", "NO"
   clPreposicoes.Add "NOS", "NOS"
   clPreposicoes.Add "NA", "NA"
   clPreposicoes.Add "NAS", "NAS"
   clPreposicoes.Add "UM", "UM"
   clPreposicoes.Add "UNS", "UNS"
   clPreposicoes.Add "NUM", "NUN"
   clPreposicoes.Add "NUNS", "NUNS"
   clPreposicoes.Add "UMA", "UMA"
   clPreposicoes.Add "UMAS", "UMAS"
   clPreposicoes.Add "NUMA", "NUNA"
   clPreposicoes.Add "NUMAS", "NUMAS"
   
   clPreposicoes.Add "DOUTRO", "DOUTRO"
   clPreposicoes.Add "DOUTROS", "DOUTROS"
   clPreposicoes.Add "DOUTRA", "DOUTRA"
   clPreposicoes.Add "DOUTRAS", "DOUTRAS"
   
   
   Set clNvString = Nothing
   Set clNvString = New Collection
   Set clPalavras = New Collection
   
   Dim ClsGeneral As Object
   Set ClsGeneral = CreateObject("xLIB.General")
   Set clPalavras = ClsGeneral.GetWords(pString)
   
   For Each n In clPalavras
       sPalavras = UCase(Trim(n))
       If Not ClsGeneral.ExisteItem(clPreposicoes, Trim(sPalavras)) Then
          If Trim(sNvString) <> "" Then sNvString = sNvString & " "
          sNvString = sNvString & sPalavras
          If Not ClsGeneral.ExisteItem(clNvString, sPalavras) Then
             clNvString.Add sPalavras, CStr(sPalavras)
          End If
       End If
   Next
   
   pString = Trim(sNvString)
   Set pClString = clNvString
End Sub
Public Sub ScrollText(ByVal hwnd As Long, Optional ByVal ScrollTo As Long = 7)
   SendMessage hwnd, WMVSCROLL, ScrollTo, 0
   'SendMessage hwnd, &H115, 7, 0
End Sub
Public Sub SelecionarTexto(ByRef Obj As Object)
   Dim Tam As Long
   
   On Error Resume Next
   If (UCase(TypeName(Obj)) = "TEXTBOX") Or (UCase(TypeName(Obj)) = "MASKEDBOX") Or (UCase(TypeName(Obj)) = "COMBOBOX") Or (UCase(TypeName(Obj)) = "FLATEDIT") Then
      Tam = Len(Obj)
      If UCase(TypeName(Obj)) = "MASKEDBOX" Then
         Tam = IIf(Len(Obj.Mask) > Tam, Len(Obj.Mask), Tam)
      End If
      Obj.SelStart = 0
      Obj.SelLength = Tam + 1
   End If
End Sub
Public Function SendTab(frm As Object, ByVal Key As Integer, Optional Tipo As Variant, _
                        Optional Obj As Variant, Optional Maiuscula = True, _
                        Optional Tamanho As Integer = 13, _
                        Optional Qtd_Dec As Integer = 2) As Integer
   
   On Error GoTo Fim
   Dim bNum As Boolean
   
   Dim MyxGeneral  As Object 'XLib.General

   If MyxGeneral Is Nothing Then Set MyxGeneral = CreateObject("xLIB.General")
   
   
   If IsMissing(Tipo) Then Tipo = -1
   If Key% = vbKeyCancel Or Key = 22 Then
      SendTab = Key%
      Exit Function
   End If
   If IsMissing(Obj) Then Set Obj = frm.ActiveControl
   'Key% = IIf(Key% = 39, 34, Key%) 'Trocar apóstrofo por aspas
   If Key% = 13 Then
      Select Case UCase(TypeName(frm.ActiveControl))
         Case "TEXTBOX", "FLATEDIT"
            If Not frm.ActiveControl.MultiLine Then
               SendKeys "{TAB}"
            End If
         Case "OPTIONBUTTON", "MASKEDBOX", "COMBOBOX", "CHECKBOX", "SSOPTION", "SSCHECK", "RADIOBUTTON"

            SendKeys "{TAB}"
      End Select
      SendTab = (Key%)
   Else
      '* Verifica se o caracter digitado é um número
      bNum = (Key >= vbKey0 And Key <= vbKey9)
      If Tipo = vbSingle Or Tipo = vbDouble Or Tipo = vbInteger Or Tipo = vbCurrency Or Tipo = vbDate Then
         If Not bNum And Key <> 8 Then
            bNum = False
            '* não é número
            If Tipo = vbSingle Or Tipo = vbDouble Or Tipo = vbInteger Or Tipo = vbCurrency Then
               If Key = 46 Then Key = 44 '* Ponto
               If Key = 44 Then  '* Vírgula
                  'If InStr(Obj.Text, ",") = 0 Then Obj.Text = "0,"
                  If IsMissing(Obj) Then Set Obj = frm.ActiveControl
                  If MyxGeneral.xVal(Obj.Text) = 0 Then
                     Obj.Text = "0,"
                     Key = 0
                     SendTab = 0
                     SendKeys "{END}"
                  End If
                  
               Else
                  Key = Asc(" ")
                  Beep
                  Exit Function
               End If
            End If
            If Tipo = vbDate Then
               If Key <> Asc("/") Then
                  
                  Key = Asc(" ")
                  Beep
                  Exit Function
               End If
            End If
         End If
      End If
      If Maiuscula Then
         SendTab = Asc(UCase(Chr$(Key%)))
      Else
         SendTab = Key%
      End If
      If Tipo = vbSingle Or Tipo = vbDouble Then
         If Key% = 46 Then '* Ponto
            Key% = 44      '* Virgula
            SendTab = 0
         End If
         If Not IsMissing(Obj) Then
            Tipo = vbCurrency
         End If
      End If
      If Tipo <> -1 Then
         If Tipo = vbCurrency Then
            Dim Ctrl As Object
            Set Ctrl = Obj
            SendTab = MyxGeneral.TratarMoeda(Key%, Ctrl, Tamanho, Qtd_Dec)
         End If
      Else
         If Maiuscula Then
            SendTab = Asc(UCase(Chr$(Key%)))
         Else
            SendTab = Key%
         End If
      End If
      If Tipo = vbDate Then
         If Not IsMissing(Obj) Then
            If bNum Then
            
               If Len(Obj.Text) = 1 Or Len(Obj.Text) = 4 Then
                  Obj.Text = Obj.Text & Chr(Key%) & "/"
                  SendTab = 0
                  SendKeys "{END}"
               End If
               If Len(Obj.Text) = 2 Or Len(Obj.Text) = 5 Then
                  Obj.Text = Obj.Text & "/" & Chr(Key%)
                  SendTab = 0
                  SendKeys "{END}"
               End If
               If Len(Obj.Text) = 10 And Obj.SelLength <> 10 Then
                    SendTab = 0
               End If
            ElseIf Key% = Asc("/") Then
                If Len(Obj.Text) <> 2 Or Len(Obj.Text) = 5 Then
                   If Len(Obj.Text) = 1 Then
                      Obj.Text = Format$(Obj.Text, "00") & "/"
                      SendTab = 0
                      SendKeys "{END}"

                   ElseIf Len(Obj.Text) = 4 Then
                      Obj.Text = Mid(Obj.Text, 1, 3) & Format$(Mid(Obj.Text, 4, 1), "00") & "/"
                      SendTab = 0
                      SendKeys "{END}"
                   Else
                      SendTab = 0
                   End If
                End If
            End If
         End If
      End If
   End If
   
   Exit Function
Fim:
   MsgBox "Dsr Error", vbCritical + vbOKOnly, "SendTab"
End Function
Public Function SendSMS(ByVal pUrl As String) As String
   Dim Status As String
   Dim oWHttp As Object
        
   'Set OWHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
   If oWHttp Is Nothing Then Set oWHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
   If oWHttp Is Nothing Then Set oWHttp = CreateObject("WinHttp.WinHttpRequest")
   If oWHttp Is Nothing Then Set oWHttp = CreateObject("MSXML2.ServerXMLHTTP")
   If oWHttp Is Nothing Then Set oWHttp = CreateObject("Microsoft.XMLHTTP")
       
   With oWHttp
      .Open "GET", pUrl
      .Send
    
      Status = .ResponseText
   End With
   SendSMS = Status
End Function
Public Sub SetRunTimeFormProperty(pForm As Object)
   Dim CurStyle As Long
   Dim NewStyle As Long


   CurStyle = GetWindowLong(pForm.hwnd, GWL_STYLE)
   NewStyle = SetWindowLong(pForm.hwnd, GWL_STYLE, CurStyle) ' Xor (WS_BORDER)) ' Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
   'Call SetWindowLong(pForm.hwnd, GWL_STYLE, GetWindowLong(pForm.hwnd, GWL_STYLE) Xor (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
   'Call SetWindowPos(pForm.hwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub
' Função auxiliar para codificar a mensagem para formato URL
Public Function SetTag(ByRef pControle As Variant, ByVal pNome As String, ByVal pValor As String) As String
   Dim StrTAG As String, StrAux As String
   Dim PosIni As Long, PosFim As Long
   
   pNome = "|" & Trim(pNome) & "="
   
   If UCase(TypeName(pControle)) = "STRING" Then
      StrTAG = pControle
   Else
      StrTAG = pControle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(pNome))
   If PosIni > 0 Then
      PosFim = InStr(PosIni + 1, StrTAG$, "|")
      PosFim = IIf(PosFim = 0, Len(StrTAG) + 1, PosFim)
      StrAux = Mid(StrTAG, 1, PosIni - 1) & Mid(StrTAG, PosIni, Len(pNome)) & Trim(pValor)
      StrAux = StrAux & Mid(StrTAG, PosFim, (Len(StrTAG) - PosFim) + 1)
      StrTAG = StrAux
   Else
      If Trim(StrTAG) = "" Then
         StrTAG = pNome & pValor
      Else
         If UCase(TypeName(pControle)) = "STRING" Then
            StrTAG = pControle & pNome & pValor
         Else
            StrTAG = pControle.Tag & pNome & pValor
         End If
      End If
   End If
   If UCase(TypeName(pControle)) = "STRING" Then
      pControle = StrTAG
   Else
      pControle.Tag = StrTAG
   End If
   SetTag = StrTAG
End Function
Public Sub SetVisualTheme(pSys As Object, Optional pForm As Object)
    On Error Resume Next
    Dim Form As Form
    Dim Ctrl As Object
                        
   If pForm Is Nothing Then
      For Each Form In Forms
         For Each Ctrl In Form.Controls
            Ctrl.VisualTheme = pSys.MDI.CommandBars.VisualTheme
         Next
      Next
   Else
      For Each Ctrl In pForm.Controls
         Ctrl.VisualTheme = pSys.MDI.CommandBars.VisualTheme
      Next
   End If
End Sub
Public Sub ShowError(Optional TxtAux = "", Optional pExibeMsg As Boolean = True)
   Dim Txt As String
   Dim Num As Long, Dsc As String
   Dim TxtTela As String
   Dim NumErrors As Long, DscErrors As String
   
   
   Dim lTitle$, lHelpFile$, lHelpContext$

   
   If Err = 0 Then Exit Sub
   
   Num = Err
   Dsc = Error
   Screen.MousePointer = vbDefault
   On Error Resume Next
'   Errors.Refresh
'   NumErrors = Errors(0).Number
'   DscErrors = Errors(0).Description
   
   'Anote o conteúdo da mensagem abaixo e avise ao analista responsável.
   'Txt$ = Txt$ & LoadMsg(9) & vbLf & vbLf
   
   If Num = NumErrors Then
      '"O Seguinte erro ocorreu : "
      '"Número : "
'      Txt$ = Txt$ + ClsMsg.LoadMsg(21) & vbLf & vbLf & DscErrors & vbLf
'      Txt$ = Txt$ & ClsMsg.LoadMsg(22) & NumErrors
      Txt$ = Txt$ + "O Seguinte erro ocorreu : " + vbLf + vbLf + DscErrors + vbLf
      Txt$ = Txt$ + "Número : " & NumErrors
'xxx      lTitle = Errors(0).Source
'xxx      lHelpFile = Errors(0).HelpFile
'xxx      lHelpContext = Errors(0).HelpContext
   Else
      '"O Seguinte erro ocorreu : "
      '"Número : "
'      Txt$ = Txt$ + ClsMsg.LoadMsg(21) & vbLf & vbLf & Dsc & vbLf
'      Txt$ = Txt$ & ClsMsg.LoadMsg(22) & Num
      Txt$ = Txt$ + "O Seguinte erro ocorreu : " + vbLf + vbLf + Dsc + vbLf
      Txt$ = Txt$ + "Número : " + Num
      
      lTitle = "ERRO"
      lHelpFile = ""
      lHelpContext = ""

   End If
'   Beep
   TxtTela = ""
   If Not Screen.ActiveForm Is Nothing Then
      TxtTela = "Tela : " & Screen.ActiveForm.Name
      If Not Screen.ActiveForm.ActiveControl Is Nothing Then
         TxtTela = TxtTela & "." & Screen.ActiveForm.ActiveControl.Name
         If Screen.ActiveForm.ActiveControl.Index <> "" Then
            TxtTela = TxtTela & "(" & CStr(Screen.ActiveForm.ActiveControl.Index) & ")"
         End If
      End If
   End If
   If TxtAux <> "" Then
      Txt = Txt & vbLf & "Auxiliar : " & vbLf & TxtAux
   End If
   On Error GoTo 0
   If pExibeMsg Then
      If Num = NumErrors Then
         MsgBox Txt, vbMsgBoxHelpButton + vbExclamation, lTitle, lHelpFile, lHelpContext
      Else
         MsgBox Txt, vbMsgBoxHelpButton + vbExclamation, lTitle
      End If
'      Call MsgBox(Txt$, 48, Errors(0).Source)
   End If
   
   '***************
   '* Gravar Erro no Arquivo de Log
   '***************
   If Num <> NumErrors Then
      DscErrors = Dsc
      NumErrors = Num
   End If

   Call ClsMensagem.GravaError(CStr(TxtAux), TxtTela$, NumErrors, DscErrors)
   On Error GoTo 0
   'Errors.Refresh
End Sub
Public Function StrZero(pValor As Variant, pQtd As Integer, Optional pCaracter = "0") As String
   If pQtd <= 0 Then
      StrZero = pValor
   Else
      pValor = Trim(Str(Val(pValor)))
      pQtd = IIf(pQtd >= Len(pValor), pQtd, Len(pValor))
      StrZero = Right(String(pQtd, pCaracter) & pValor, pQtd)
   End If
End Function
Public Function TratarMoeda(Key%, ByRef Obj As Object, Optional Tamanho As Integer, Optional Qtd_Dec As Integer = 2) As String
   Dim qtCasaDecimal As String
   Dim Max           As String
   Dim Numero        As String
   Dim i             As Integer
   Dim NumMax#, TamMaxNum%, TamMaxCarac%, Qtd_Ponto%
   
   Dim MySepDec As String
   MySepDec = ","

   Qtd_Ponto = Int((Obj.MaxLength - Qtd_Dec - 1) / 3) - 1
   Qtd_Ponto = IIf(Qtd_Ponto < 0, 0, Qtd_Ponto)
   Max = ""
   If Not IsMissing(Tamanho) Then
      If (Tamanho - Qtd_Dec) / 3 = CInt((Tamanho - Qtd_Dec) / 3) Then
         Obj.MaxLength = Tamanho + (CInt((Tamanho - Qtd_Dec) / 3))
      Else
         Obj.MaxLength = Tamanho + (CInt((Tamanho - Qtd_Dec) / 3)) + 1
      End If
   End If
   TamMaxNum% = Obj.MaxLength - (Qtd_Dec + Qtd_Ponto + 2)
   
   TamMaxCarac% = TamMaxNum% + Qtd_Dec + 1
   For i = 1 To TamMaxNum%
      Max = Max + "9"
   Next
  
   If Trim(Max) = "" Then
      NumMax# = 100
   Else
      NumMax# = CDbl(Max)
   End If
   Numero = Obj.Text
   Select Case Key
      Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57:
         If Len(Trim$(Numero)) = 0 Then
            Numero = 0
         End If
         If (Trim$(Numero)) = MySepDec Then
            Numero = 0
         End If
         If CDbl(Numero) >= NumMax# Then
            For i = 1 To Len(Trim$(Numero))
               If Mid$(Trim$(Numero), i, 1) = MySepDec Then
                  Beep
                  TratarMoeda = 0
                  Exit For
               End If
            Next i
            qtCasaDecimal = Len(Trim$(Numero)) - i
            If Not (qtCasaDecimal > -1 And _
                   (Len(Trim$(Numero)) - 1) < TamMaxCarac%) Then
               Beep
               TratarMoeda = 0
               Exit Function
            End If
         End If
         TratarMoeda = Key
      Case 8  'TAb
         TratarMoeda = Key
      Case Asc(MySepDec)
         TratarMoeda = Key
         For i = 1 To Len(Trim$(Numero))
             If Mid$(Trim$(Numero), i, 1) = MySepDec Then
                Beep
                TratarMoeda = 0
                Exit For
             End If
         Next i
      
      Case 46: TratarMoeda = 44 '*Substitur o ponto pela virgula
      Case Else
         Beep
         TratarMoeda = 0
   End Select
End Function
Public Function UnFormat(ByVal Codigo As String) As String
   Dim sAux    As String
   Dim i       As Integer
   Dim sChar   As String
   Dim nAsc    As Integer
   
   sAux = ""
   For i = 1 To Len(Codigo)
      sChar = UCase(Mid(Codigo, i, 1))
      nAsc = Asc(sChar)
      If (nAsc >= 65 And nAsc <= 90) Or (nAsc >= 48 And nAsc <= 57) Then
         sAux = sAux & Mid(Codigo, i, 1)
      End If
   Next
   
   UnFormat = Trim(sAux)
End Function
Public Function UrlEncode(ByVal urlText As String) As String
    Dim i As Long
    Dim ansi() As Byte
    Dim ascii As Integer
    Dim encText As String

    ansi = StrConv(urlText, vbFromUnicode)

    encText = ""
    For i = 0 To UBound(ansi)
        ascii = ansi(i)

        Select Case ascii
        Case 48 To 57, 65 To 90, 97 To 122
            encText = encText & Chr(ascii)
        Case 32
            encText = encText & "+"
        Case Else
            If ascii < 16 Then
                encText = encText & "%0" & Hex(ascii)
            Else
                encText = encText & "%" & Hex(ascii)
            End If
        End Select
    Next i
    
    UrlEncode = encText
End Function
Public Function ValBr(ByVal pNum As String, Optional pCasaDec As Integer = 2, Optional pArredTrunc As Integer = 1) As String
   Dim sAux    As String
   Dim nValor  As Currency
   Dim sSep    As String
   
   nValor = xVal(pNum)
   sSep = Mid(FormatNumber(1.1, 1), 2, 1)
   
   Select Case pArredTrunc
      Case 0 '* Formatar para [pCasaDec] ou deixar qtd atual se este for maior que [pCasaDec]
         sAux = FormatNumber(nValor - Fix(nValor), 20, vbUseDefault, vbUseDefault, vbTrue)
         sAux = StrReverse(xVal(StrReverse(Mid(sAux, 3))))
         If Len(sAux) < pCasaDec Then
            sAux = Left(sAux & String(pCasaDec, "0"), pCasaDec)
         End If
         ValBr = CStr(FormatNumber(Fix(nValor), 0)) & sSep & sAux
      Case 1 '* Arredondar decimal
         ValBr = FormatNumber(nValor, pCasaDec, vbUseDefault, vbUseDefault, vbTrue)
      Case 2 '* Truncar decimal
         sAux = FormatNumber(nValor - Fix(nValor), 20, vbUseDefault, vbUseDefault, vbTrue)
         ValBr = CStr(Fix(nValor)) & Mid(sAux, 2, pCasaDec + 1)
   End Select
End Function
Public Function ValidarCPF(ByVal NumCPF As String) As Boolean
   Dim digito, i As Integer
   Dim DigitoCPF(11)
   
   'Atribui o cpf digitado a um vetor
   NumCPF = UnFormat(NumCPF)
   For i = 1 To 11
       DigitoCPF(i) = Val(Mid$(NumCPF, i, 1))
   Next i
   
   'Calcula primeiro digito verificador
   digito = Calcular_Digito_CPF(10, 9, DigitoCPF)
   If digito <> DigitoCPF(10) Then
       ValidarCPF = False
       Exit Function
   End If
   
   'Calcula segundo digito verificador
   digito = Calcular_Digito_CPF(11, 10, DigitoCPF)
   If digito <> DigitoCPF(11) Then
       ValidarCPF = False
       Exit Function
   End If
   
   ValidarCPF = True
End Function
Public Function ValidarCNPJ(ByVal NumCNPJ) As Boolean
   
   Dim Resto%, Produto%, Tamanho%, i%, j%
   Dim Matriz As String
   Dim CNPJ_Resultante As String
   
   NumCNPJ = UnFormat(NumCNPJ)
   If Len(NumCNPJ) <> 14 Then
       ValidarCNPJ = False
       Exit Function
   End If
   
   Matriz = "2345678923456"
   Tamanho% = 12
   CNPJ_Resultante = Left$(NumCNPJ, Tamanho%)
   
   For j% = 1 To 2
       Produto% = 0
       For i% = Tamanho% To 1 Step -1
           Produto% = Produto% + Val(Mid$(CNPJ_Resultante, i%, 1)) * Val(Mid$(Matriz, Tamanho% - i% + 1, 1))
       Next i%
       Resto% = Produto% Mod 11
       CNPJ_Resultante = CNPJ_Resultante & IIf(Resto% <= 1, "0", Format$(11 - Resto%, "0"))
       Tamanho% = Tamanho% + 1
   Next j%
   
   ValidarCNPJ = (NumCNPJ = CNPJ_Resultante)

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
Public Sub Wait(pSecond As Integer)
   Dim Time01 As Long
   Dim Time02 As Long

   Time01 = (Hour(Time) * 60 * 60) + (Minute(Time) * 60) + Second(Time)
   While Time02 - Time01 < pSecond
      Time02 = (Hour(Time) * 60 * 60) + (Minute(Time) * 60) + Second(Time)
      DoEvents
   Wend
End Sub
Public Function xDate(pDate As String, Optional ComHora = True, Optional pFormatoData As String = "DMA") As String
   Dim dDate As Date
   Dim sDate As String
   Dim sHora As String
   Dim Pos As Integer
   Dim sAux As String
   Dim sMes As String
   
   On Error GoTo TrataErro
   
   sDate = Trim(pDate)
   
   If sDate = "" Then Exit Function
   
   sDate = Format(sDate, "dd/mm/yyyy hh:mm:ss")
   
   
   Pos = InStr(sDate, " ")
   If Pos <> 0 Then
      sHora = Mid(sDate, Pos + 1)
      sDate = Mid(sDate, 1, Pos - 1)
   End If
   sAux = Trim(sDate)
   
   sDate = StrZero(Mid(sDate, 1, InStr(sDate, "/") - 1), InStr(InStr(sDate, "/"), sDate, "/") - 1)
   Pos = InStr(sAux, "/") + 1
   
   sMes = Month("28/" & Mid(sAux, Pos, InStr(Pos, sAux, "/") - Pos) & "/1900")
   
   sDate = sDate & "/" & StrZero(sMes, 2)
   
   Pos = InStr(Pos, sAux, "/") + 1
   If Len(Mid(sAux, Pos)) = 2 Then
      sDate = sDate & "/" & VBA.Right$(Mid(sAux, Pos), 2)
   ElseIf Len(Mid(sAux, Pos)) = 4 Then
      If Mid(Mid(sAux, Pos), 1, 2) > 18 Then
         sDate = sDate & "/" & Mid(sAux, Pos)
      Else
         sDate = sDate & "/" & VBA.Right$(Mid(sAux, Pos), 2)
      End If
   Else
      sDate = sDate & "/" & VBA.Right$(Mid(sDate, Pos), 2)
   End If
         
   dDate = VBA.CDate(sDate)
   sDate = ""
   
   If pFormatoData = "DMA" Then
      sDate = StrZero(Day(dDate), 2) & "/" & StrZero(Month(dDate), 2) & "/" & Year(dDate)
   ElseIf pFormatoData = "MDA" Then
      sDate = StrZero(Month(dDate), 2) & "/" & StrZero(Day(dDate), 2) & "/" & Year(dDate)
   ElseIf pFormatoData = "AMD" Then
      sDate = Year(dDate) & "/" & StrZero(Month(dDate), 2) & "/" & StrZero(Day(dDate), 2)
   Else
      sDate = StrZero(Day(dDate), 2) & "/" & StrZero(Month(dDate), 2) & "/" & Year(dDate)
   End If
   
   If xVal(Hour(pDate)) <> 0 And ComHora Then
      sDate = sDate & " " & StrZero(Hour(pDate), 2) & ":" & StrZero(Minute(pDate), 2) & ":" & StrZero(Second(pDate), 2)
   End If
   
   xDate = sDate
Exit Function
TrataErro:
   xDate = sDate
   Call ClsMensagem.ExibirStop("Formato de Data Inválido", "Atenção")
End Function
Public Function xVal(ByVal pNum As String, Optional pQtdCasaDec = 5) As Double
   Dim PosV As Integer
   Dim PosP As Integer
   
   '* Verificar Formatação da 'String'
   pNum = Trim(Replace(pNum, "R$", ""))
   
   PosV = InStr(pNum, ",")
   PosP = InStr(pNum, ".")
   
   If PosV <> 0 And PosP <> 0 Then
      '* Entrar apenas se número estiver Formatado
      '* e desformatá-lo para recuperar valor numérico.
      
      If PosV < PosP Then
         '* 'String' com formatação Americana
         '* trocar para formato brasileiro.
         '* Ex:. 23,455,654.98 ==> 23.455.654,98
         
         pNum = Replace(pNum, ",", "#")
         pNum = Replace(pNum, ".", ",")
         pNum = Replace(pNum, "#", ".")
      End If
   End If
   '* desformatar 'String'
   pNum = Format(pNum, "##." & String(pQtdCasaDec, "0"))

   
   '* Recuperar valor numérico da 'string'.
   xVal = Val(Replace(Replace(pNum, ".", ""), ",", "."))
End Function
Public Function isTransparent(ByVal hwnd As Long) As Boolean
   On Error Resume Next
   Dim Msg As Long
   Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
     isTransparent = True
   Else
     isTransparent = False
   End If
   If Err Then
     isTransparent = False
   End If
End Function
Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
   Dim Msg As Long
   On Error Resume Next
   If Perc < 0 Or Perc > 255 Then
     MakeTransparent = 1
   Else
     Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
     Msg = Msg Or WS_EX_LAYERED
     SetWindowLong hwnd, GWL_EXSTYLE, Msg
     SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
     MakeTransparent = 0
   End If
   If Err Then
     MakeTransparent = 2
   End If
End Function
Public Function MakeOpaque(ByVal hwnd As Long) As Long
   Dim Msg As Long
   On Error Resume Next
   Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   Msg = Msg And Not WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, Msg
   SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
   MakeOpaque = 0
   If Err Then
     MakeOpaque = 2
   End If
End Function
Private Function Transparency(ByVal hwnd As Long, Optional ByVal Col As Long = vbBlack, Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
   ' Return : True if there is no error.
   ' hWnd   : hWnd of the window to make transparent
   ' Col : Color to make transparent if TrMode=False
   ' PcTransp  : 0 à 255 >> 0 = transparent  -:- 255 = Opaque
   Dim DisplayStyle As Long
   
   On Error GoTo Saida
   
   DisplayStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
   If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
      DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
      Call SetWindowLong(hwnd, GWL_EXSTYLE, DisplayStyle)
   End If
   Transparency = (SetLayeredWindowAttributes(hwnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_ALPHA, LWA_COLORKEY)) <> 0)
Saida:
    If Not Err.Number = 0 Then Err.Clear
End Function
Public Sub ActiveTransparency(pForm As Object, d As Boolean, F As Boolean, Perc As Integer, Optional Color As Long)
   Dim bResult As Boolean
   
   If d And F Then
   'Makes color (here the background color of the shape) transparent
   'upon value of T_Transparency
       bResult = Transparency(pForm.hwnd, Color, Perc, False)
   ElseIf d Then
       'Makes form, including all components, transparent
       'upon value of T_Transparency
       bResult = Transparency(pForm.hwnd, 0, Perc, True)
   Else
       'Restores the form opaque.
       bResult = Transparency(pForm.hwnd, , 255, True)
   End If
End Sub
