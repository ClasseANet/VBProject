Attribute VB_Name = "XDSR"
Global Const BLOCK_SIZE = 16384
'************************************************************************************
'* 18 FUNÇÕES DE DSR100.DLL                                                         *
'************************************************************************************
'* Public Function Copy(Orig$, dest$)                                               *
'* Public Sub Del(Arq$)                                                             *
'* Public Function EliminarString(ByVal Palavra$, ByVal Caracter$, _                *
'*                               Optional CaseSensitive = True) As String           *
'* Public Sub ExibirAviso(Txt$, Optional Tit$)                                      *
'* Public Function ExibirPergunta%(Txt$, Optional Tit$)                             *
'* Public Sub ExibirStop(Txt$, Optional Tit$)                                       *
'* Public Function FileExists(ByVal strPathName As String) As Boolean               *
'* Public Sub GravaError(TxtAux$, TxtTela$)                                         *
'* Public Function InArray(Valor As Variant, VETOR As Variant)                      *
'* Public Function LoadMsg$(Num%)                                                   *
'* Public Function MakePath(ByVal strDir$, Optional ByVal fAllowIgnore) As Boolean  *
'* Public Function MakePathAux(ByVal strDirName As String) As Boolean               *
'* Public Sub ShowError(Optional TxtAux = "")                                       *
'* Public Function SqlDate$(ByVal DT$)                                              *
'* Public Function SqlNum(ByVal Num$) As String                                     *
'* Public Function SqlStr(Txt$) As String                                           *
'* Public Function StrReplace(ByVal txtin$, ByVal txtfrom$, ByVal txtto$)           *
'* Public Function StrZero(Valor As Variant, Num As Integer, _                      *
'*                         Optional Caracter = "0") As String                       *
'* Public Function Tiraponto$(ByVal com_ponto As String)                            *
'* public Function xVal(ByVal Num$, Optional NumCasaDec = 5)                        *
'************************************************************************************
'****************************************
'* Variáveis globais e funções          *
'* para suprir a ausência de DSR100.DLL *
'****************************************

Global gIdioma As Long
Global gDrvErro As String
Global gDrvTmp As String
Global gSepDec As String
Global gSepMil As String
Global gSepDt As String
Global gDtMask As String
Global gComputerName As String
Public Function SqlStr(Txt As String) As String
   SqlStr = "'" + CStr(Txt) + "'"
End Function

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
   Errors.Refresh
   NumErrors = Errors(0).Number
   DscErrors = Errors(0).Description
   
   'Anote o conteúdo da mensagem abaixo e avise ao analista responsável.
   'Txt$ = Txt$ & LoadMsg(9) & vbLf & vbLf
   
   If Num = NumErrors Then
      'O Seguinte erro ocorreu : "
      'Número : "
      Txt$ = Txt$ + LoadMsg(21) & vbLf & vbLf & DscErrors & vbLf
      Txt$ = Txt$ & LoadMsg(22) & NumErrors
      lTitle = Errors(0).Source
      lHelpFile = Errors(0).HelpFile
      lHelpContext = Errors(0).HelpContext
   Else
      'O Seguinte erro ocorreu : "
      'Número : "
      Txt$ = Txt$ & LoadMsg(21) & vbLf & vbLf
      Txt$ = Txt$ & Dsc & vbLf
      Txt$ = Txt$ & LoadMsg(22) & Num
      
      lTitle = "ERRO"
      lHelpFile = ""
      lHelpContext = ""

   End If
'   Beep
   TxtTela = ""
   If Not Screen.ActiveForm Is Nothing Then
      TxtTela = "Tela : " & Screen.ActiveForm.NAME
      If Not Screen.ActiveForm.ActiveControl Is Nothing Then
         TxtTela = TxtTela & "." & Screen.ActiveForm.ActiveControl.NAME
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
   On Error Resume Next
   Call GravaError(CStr(TxtAux), TxtTela$)
   On Error GoTo 0
   Errors.Refresh
End Sub
Public Function SqlDate$(ByVal DT$)
   If Val(DT$) = 0 Then
      SqlDate$ = "Null"
   Else
      SqlDate$ = "CDATE('" + Format$(DT$, gDtMask) + "')"
   End If
End Function
Public Function SqlNum(ByVal Num$) As String
    If Trim(Num$) = "" Then
        SqlNum = "0"
    Else
        SqlNum = StrReplace(Tiraponto(Num$), ",", ".")
    End If
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
'-----------------------------------------------------------
' FUNCTION: MakePathAux
'
' Creates the specified directory path.
'
' No user interaction occurs if an error is encountered.
' If user interaction is desired, use the related
'   MakePathAux() function.
'
' IN: [strDirName] - name of the dir path to make
'
' Returns: True if successful, False if error.
'-----------------------------------------------------------
'
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
Public Function InArray(Valor As Variant, VETOR As Variant)
   Dim j As Variant
   InArray = False
   For Each j In VETOR
       If Valor = j Then
         InArray = True
         Exit For
      End If
   Next
End Function
Public Sub ExibirStop(Txt$, Optional Tit$)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Exibir um aviso na tela.                        **
'**                                                            **
'** Recebe: Mensagem$ - Aviso a ser exibido                    **
'**         Tit$   - Título da Mensagem (Opcional)          **
'**                                                            **
'** Retorna : Aviso centralizado na tela com um botão de OK e  **
'**           um ícone de Stop.                                **
'**                                                            **
'****************************************************************
    Dim Mouse%
    Mouse = Screen.MousePointer
    If TypeName(Tit$) = "Nothing" Then
       Tit$ = LoadMsg(1)
    End If
    Screen.MousePointer = vbDefault
    MsgBox Txt$, vbCritical, Tit$
    Screen.MousePointer = Mouse
End Sub
Public Sub Del(Arq$)
   If FileExists(Arq$) Then
      On Error GoTo fim
      Kill Arq$
   End If
   Exit Sub
fim:
   ShowError
End Sub
Public Function EliminarString(ByVal Palavra$, ByVal Caracter$, Optional CaseSensitive = True) As String
    Dim pos%, Com_Carac$
    
    If CaseSensitive Then
       pos% = InStr(Palavra, Caracter)
    Else
      pos% = InStr(UCase(Palavra), UCase(Caracter))
    End If
    
    Com_Carac = Palavra$
    While pos% <> 0
        Com_Carac = Left$(Com_Carac, pos% - 1) + Mid$(Com_Carac, pos% + Len(Caracter$))
        If CaseSensitive Then
           pos% = InStr(Com_Carac$, Caracter)
        Else
           pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
        End If
    Wend

    If CaseSensitive Then
       pos% = InStr(Com_Carac$, Caracter)
    Else
       pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
    End If
    
    If pos% > 0 Then
       If CaseSensitive Then
          pos% = InStr(pos% + Len(Caracter$), Com_Carac, Caracter)
       Else
          pos% = InStr(pos% + Len(Caracter$), UCase(Com_Carac), UCase(Caracter))
       End If
       If pos% > 0 Then Com_Carac = Left$(Com_Carac, pos% - 1)
    End If

    EliminarString = Com_Carac
End Function
Public Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

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
Public Function LoadMsg$(Num%)
   Dim Idioma As eIdioma
   Dim msg As String
   On Error GoTo fim
   '* Idioma usado
   Idioma = mvarIdioma 'gIdioma 'Português = 5000, Inglês = 6000, ...
   'Num% = Idioma% + Num%
   'Select Case Idioma
   '   Case 5000: Msg$ = LoadResString(Num) 'LoadPortugues$(Num%)
   '   Case 6000: Msg$ = LoadResString(Num) 'LoadIngles$(Num%)
   'End Select
   'LoadMsg$ = Msg$
   If Num > 9999 Then
      Num = 5000 + (Num - (Val(Mid(CStr(Num), 1, Len(CStr(Num)) - 3)) * 1000))
   End If
   LoadMsg$ = LoadResString(gIdioma + Num)
   Exit Function
fim:
   ShowError
End Function
Public Sub ExibirAviso(Txt$, Optional Tit$)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Exibir um aviso na tela.                        **
'**                                                            **
'** Recebe: Mensagem$ - Aviso a ser exibido                    **
'**         Tit$   - Título da Mensagem (Opcional)          **
'**                                                            **
'** Retorna : Aviso centralizado na tela com um botão de OK e  **
'**           um ícone de Exclamação.                          **
'**                                                            **
'****************************************************************
   Dim Mouse%
   Mouse = Screen.MousePointer
   If TypeName(Tit$) = "Nothing" Then
      Tit$ = LoadMsg(1)
   End If
   Screen.MousePointer = vbDefault
   MsgBox Txt$, vbExclamation, Tit$
   DoEvents
   Screen.MousePointer = vbDefault
End Sub
Public Function ExibirPergunta%(Txt$, Optional Tit$)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Exibir uma pergunta na tela.                    **
'**                                                            **
'** Recebe: Mensagem$ - Pergunta a ser exibida                 **
'**         Tit$   - Título da Mensagem (Opcional)          **
'**                                                            **
'** Retorna : Resposta da pergunta centralizada com botões Yes **
'**           e No e com um ícone de Interrogação.             **
'**                                                            **
'****************************************************************
    Dim Mouse%
    Mouse = Screen.MousePointer
    If TypeName(Tit$) = "Nothing" Then
       Tit$ = LoadMsg(1)
    End If
    Screen.MousePointer = vbDefault
    ExibirPergunta% = MsgBox(Trim(Txt$), vbQuestion + vbYesNo, Tit$)
    DoEvents
    Screen.MousePointer = Mouse
End Function
Public Sub GravaError(TxtAux$, TxtTela$)
   Dim ArqLog$, ArqTmp$
   Dim pos%, Limpar As Boolean
   Dim Txt$, i%
   Dim nFile As Long
   Dim nFile1 As Long
   Dim nFile2 As Long
   
   
   Dim gDrvErro As String
   
   On Error Resume Next
   
   ArqLog$ = "Error.log"
   ArqTmp$ = "Error.tmp"
   gDrvErro = gDrvTmp
   nFile = FreeFile
   If Not FileExists(gDrvErro & ArqLog$) Then
      Open gDrvErro & ArqLog$ For Output As #nFile
      Print #nFile, "0 Erros"
      Close #nFile
   End If
   Limpar = (FileLen(gDrvErro & ArqLog$) > 1000000) ' > 1 Mb
   Call Copy(gDrvErro & ArqLog$, gDrvErro & ArqTmp$)
   
   nFile1 = FreeFile
   Open gDrvErro & ArqLog$ For Output As #nFile1
   nFile2 = FreeFile
   Open gDrvErro & ArqTmp$ For Input As #nFile2
   Line Input #nFile2, Txt
   If Limpar Then
      Print #nFile1, "4 Erros"
   Else
      pos = 0
      While pos = 0
         pos = InStr(Txt, "Erros")
         If pos > 2 Then
            If pos > 5 Then
               Print #nFile1, CStr(CLng(Mid(Txt, pos - 5, 5)) + 1) & " Erros"
            Else
               Print #nFile1, CStr(CLng(Mid(Txt, 1, pos - 2)) + 1) & " Erros"
            End If
         End If
         If Not EOF(nFile2) And pos = 0 Then Line Input #nFile2, Txt
         pos = IIf(EOF(nFile2), 1, pos)
      Wend
   End If
   Print #nFile1, "====================================================================="
   Print #nFile1, "Date     : " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
   Print #nFile1, "Computer : " & gComputerName
   Print #nFile1, "Source   : " & Errors(0).Source
   Print #nFile1, "Erro     : " & Errors(0).Number & " - " & Mid(Errors(0).Description, 1, 50)
   Txt = Errors(0).Description
   While Len(Txt) >= 50
      Txt = Mid(Txt, 51)
      Print #nFile1, Space(18) & Mid(Txt, 1, 50)
   Wend
   Print #nFile1, "Help     : " & Errors(0).HelpContext & " - " & Errors(0).HelpFile
   If TxtTela <> "" Then
      Print #nFile, TxtTela
   End If
   If TxtAux <> "" Then
      Print #nFile1, "Auxiliar : " & TxtAux
   End If
   Print #nFile1, "====================================================================="
   Do While Not EOF(nFile2)
      Line Input #nFile2, Txt
      Print #nFile1, Txt
      If Limpar And Mid(Txt, 1, 5) = "=====" Then '* Se Arquivo > 1Mb
         i = i + 1
         If i = 6 Then Exit Do '* Gravar apenas os 3 últimos Erros
      End If
   Loop
   Close #nFile2
   '   Print #1, "====================================================================="
   Close #nFile
   Kill gDrvErro & ArqTmp$
End Sub
Public Function Tiraponto$(ByVal com_ponto As String)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Tirar pontos separador de centenas em uma string**
'**                                                            **
'** Recebe: com_ponto - String do número                       **
'**                                                            **
'** Retorna : string do número sem os pontos                   **
'**                                                            **
'****************************************************************
    Dim a%
    
    a% = InStr(com_ponto$, gSepMil$)
    While a% <> 0
        com_ponto = Left$(com_ponto$, a% - 1) + Mid$(com_ponto$, a% + 1)
        a% = InStr(com_ponto$, gSepMil$)
    Wend

    a% = InStr(com_ponto$, gSepDec$)
    
    If a% > 0 Then
       a = InStr(a% + 1, com_ponto$, gSepDec$)
       If a% > 0 Then com_ponto$ = Left$(com_ponto$, a% - 1)
    End If

    Tiraponto$ = com_ponto$

End Function

Public Function StrReplace(ByVal txtin$, ByVal txtfrom$, ByVal txtto$)
'================================================================
'= Última Alteração : 20/01/99                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Substitui o texto de Txtfrom$ para o texto de   **
'**            TxtOut$ na string TxtIn$.                       **
'**                                                            **
'** Recebe: TxtIn$   - string a ser alterada                   **
'**         TxtFrom$ - texto a ser substituido                 **
'**         TxtOut$   - novo texto                             **
'**                                                            **
'** Retorna: string alterada                                   **
'**                                                            **
'****************************************************************
   Dim TxtOut$, LenIn%, LenFrom%, pos%

   LenIn = Len(txtin)
   LenFrom = Len(txtfrom)
   If LenFrom < 1 Or LenIn < 1 Then
      StrReplace = txtin
      Exit Function
   End If
   TxtOut = ""
   pos = InStr(txtin, txtfrom)
   While pos > 0
      TxtOut = TxtOut + Left(txtin, pos - 1) + txtto
      txtin = Right(txtin, Len(txtin) - pos - LenFrom + 1)
      pos = InStr(txtin, txtfrom)
   Wend
   TxtOut = TxtOut + txtin
   StrReplace = TxtOut
End Function
Public Function StrZero(Valor As Variant, Num As Integer, Optional Caracter = "0") As String
   Dim i%, Zeros$
   Zeros = String(Num%, Caracter)
   StrZero = Right(Zeros + Trim(Str(Val(Valor))), Num%)
End Function
Public Function Copy(Orig$, dest$)
   Dim nMsg$, nTipo&, NL
   Dim Resp%
   NL = vbLf
   On Error Resume Next
   If FileExists(Orig$) Then
      Kill Arq$
      FileCopy Orig$, dest$
   Else
      Call ExibirAviso(LoadMsg(11) + UCase(Orig$), LoadMsg(12))
      Resp = vbCancel
      Exit Function
   End If
   Resp = vbYes
   Select Case Err
      Case 71
         While Resp = vbYes
            nMsg = LoadMsg(13) + NL + NL
            nMsg = nMsg & LoadMsg(14) + NL
            nMsg = nMsg & LoadMsg(15)
            nTipo = vbYesNo + vbCritical + vbDefaultButton1
            Resp = MsgBox(nMsg, nTipo, LoadMsg(16))
            If Resp = vbYes Then
               Err = 0
               FileCopy Orig$, dest$
            End If
         Wend
      Case 70
         While Resp = vbOK
            nMsg = LoadMsg(7) + NL + NL
            nMsg = nMsg & LoadMsg(56) + NL
            nTipo = vbYesNo + vbCritical + vbDefaultButton1
            Resp = MsgBox(nMsg, nTipo, LoadMsg(16))
            If Resp = vbYes Then
               Err = 0
               FileCopy Orig$, dest$
            End If
         Wend
   End Select
   Copy = Resp
End Function
Public Function xVal(ByVal Num$, Optional NumCasaDec = 5)
   Dim PosV As Integer
   Dim PosP As Integer
   
   '* Verificar Formatação da 'String'
   Num$ = Trim(EliminarString(Num$, "R$"))
   
   PosV = InStr(Num, ",")
   PosP = InStr(Num, ".")
   
   If PosV <> 0 And PosP <> 0 Then
      '* Entrar apenas se número estiver Formatado
      '* e desformatá-lo para recuperar valor numérico.
      
      If PosV < PosP Then
         '* 'String' com formatação Americana
         '* trocar para formato brasileiro.
         '* Ex:. 23,455,654.98 ==> 23.455.654,98
         
         Num = StrReplace(Num, ",", "#")
         Num = StrReplace(Num, ".", ",")
         Num = StrReplace(Num, "#", ".")
      End If
      
      '* desformatar 'String'
      Num = Format(Num, "##." & String(NumCasaDec, "0"))
      
   End If
   
   '* Recuperar valor numérico da 'string'.
   xVal = Val(StrReplace(Num, ",", "."))
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
Public Function isAlfaNum(ch As String) As Boolean
   isAlfaNum = (ch >= "0" And ch <= "9") Or (UCase(ch) >= "A" And UCase(ch) <= "Z")
   isAlfaNum = isAlfaNum Or (UCase(ch) >= "À" And UCase(ch) <= "Ý")    '* Or (Ch = "_")
End Function
Public Function RichWordOver(ByVal RchTxt As Variant, x As Single, y As Single, Optional Posicao = 1, Optional VerifDclImplicta = True) As String
   Dim pt As PointAPI
   Dim pos As Integer
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
         pos = OverAllCursorPos - 2 * (CurrLineCursorPos)
         If pos < 0 Then Exit Function
      Else
         pos = Posicao
      End If
   
   Else
      pt.x = x \ Screen.TwipsPerPixelX
      pt.y = y \ Screen.TwipsPerPixelY
      pos = SendMessage(RchTxt.hwnd, EM_CHARFROMPOS, 0&, pt)
   End If
      
   If pos <= 0 Then Exit Function

   ' Find the start of the word.
   If TypeName(RchTxt) = "String" Then
      Txt = RchTxt
   Else
      Txt = RchTxt.Text
   End If
   
   Posicao = pos
   Do While Trim(Mid$(Txt, pos, 1)) = ""
      pos = pos + 1
      If pos > Len(Txt) Then
         pos = Posicao
         Exit Do
      End If
   Loop
   For StartPos = pos To 1 Step -1
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
   For EndPos = pos To TxtLen
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


