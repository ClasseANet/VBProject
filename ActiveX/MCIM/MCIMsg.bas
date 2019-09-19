Attribute VB_Name = "MCIMsg"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal strUser As String, lngBuffer As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Sub main()
   FrmInicio.Show
End Sub
'Public Const MAX_PATH = 260
'Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'Public Function GetComputerName() As String
'    '-------------------------------------------------------------
'    ' Returns the name of the computer.
'    '-------------------------------------------------------------
'    '   API declarations:
'    '-------------------------------------------------------------
'    '   Private Const MAX_PATH = 260
'    '   Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'    '-------------------------------------------------------------
'   Dim sBuffer As String
'   Dim lRet    As Long
'   Dim nSize   As Long
'
'   nSize = MAX_PATH
'   sBuffer = Space$(MAX_PATH)
'
'
'
'   lRet = GetComputerNameA(sBuffer, nSize)
'   If lRet <> 0 Then
'      GetComputerName = UCase$(Left$(Trim$(sBuffer), Len(Trim$(sBuffer)) - 1))
'      GetComputerName = EliminarString(GetComputerName, Chr(0))
'   Else
'      GetComputerName = ""
'   End If
'End Function
'Public Function EliminarString(ByVal Palavra$, ByVal Caracter$, Optional CaseSensitive = True) As String
'    Dim Pos%, Com_Carac$
'
'    If CaseSensitive Then
'       Pos% = InStr(Palavra, Caracter)
'    Else
'      Pos% = InStr(UCase(Palavra), UCase(Caracter))
'    End If
'
'    Com_Carac = Palavra$
'    While Pos% <> 0
'        Com_Carac = Left$(Com_Carac, Pos% - 1) + Mid$(Com_Carac, Pos% + Len(Caracter$))
'        If CaseSensitive Then
'           Pos% = InStr(Com_Carac$, Caracter)
'        Else
'           Pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
'        End If
'    Wend
'
'    If CaseSensitive Then
'       Pos% = InStr(Com_Carac$, Caracter)
'    Else
'       Pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
'    End If
'
'    If Pos% > 0 Then
'       If CaseSensitive Then
'          Pos% = InStr(Pos% + Len(Caracter$), Com_Carac, Caracter)
'       Else
'          Pos% = InStr(Pos% + Len(Caracter$), UCase(Com_Carac), UCase(Caracter))
'       End If
'       If Pos% > 0 Then Com_Carac = Left$(Com_Carac, Pos% - 1)
'    End If
'
'    EliminarString = Com_Carac
'End Function
'
Public Function GetTag(ByRef Controle As Variant, ByVal VarName As String, Optional VarDefault As String) As String
   Dim PosIni As Long, posfim As Long
   Dim StrTAG As String
   Dim i%
   
   On Error GoTo Saida
   
   VarName = "|" & Trim(VarName) & "="
   
   If UCase(TypeName(Controle)) = "STRING" Then
      StrTAG = Controle
   Else
      StrTAG = Controle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(VarName))
   If PosIni > 0 Then
      PosIni = PosIni + Len(Trim(VarName))
      posfim = InStr(PosIni, StrTAG$, "|")
      i = 0
      While Mid(StrTAG$, PosIni + i, 1) = "|"
         i = i + 1
      Wend
      If i > 0 Then
         posfim = InStr(PosIni + (i - 1), StrTAG$, "|")
      End If
      posfim = IIf(posfim = 0, Len(StrTAG$), posfim - 1)
      StrTAG$ = Mid(StrTAG$, PosIni, posfim - PosIni + 1)
   Else
      StrTAG$ = ""
   End If
   GetTag = StrTAG$
Saida:
   If StrTAG$ = "" Then
      GetTag = VarDefault
   End If
End Function
Public Function SetTag(ByRef Controle As Variant, ByVal VarName As String, ByVal VarValor As String) As String
   Dim StrTAG As String, StrAux As String
   Dim PosIni As Long, posfim As Long
   
   VarName = "|" & Trim(VarName) & "="
   
   If UCase(TypeName(Controle)) = "STRING" Then
      StrTAG = Controle
   Else
      StrTAG = Controle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(VarName))
   If PosIni > 0 Then
      posfim = InStr(PosIni + 1, StrTAG$, "|")
      posfim = IIf(posfim = 0, Len(StrTAG) + 1, posfim)
      StrAux = Mid(StrTAG, 1, PosIni - 1) & Mid(StrTAG, PosIni, Len(VarName)) & Trim(VarValor)
      StrAux = StrAux & Mid(StrTAG, posfim, (Len(StrTAG) - posfim) + 1)
      StrTAG = StrAux
   Else
      If Trim(StrTAG) = "" Then
         StrTAG = VarName & VarValor
      Else
         If UCase(TypeName(Controle)) = "STRING" Then
            StrTAG = Controle & VarName & VarValor
         Else
            StrTAG = Controle.Tag & VarName & VarValor
         End If
      End If
   End If
   If UCase(TypeName(Controle)) = "STRING" Then
      Controle = StrTAG
   Else
      Controle.Tag = StrTAG
   End If
   SetTag = StrTAG
End Function
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = VBA.Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Function NetworkUserID() As String
  Dim lngBufferSize As Long
  Dim strUser As String
  
  Dim glngReturnStatus As Long
    
  On Error GoTo NetworkUserID_EH

  NetworkUserID = "Usuário desconhecido."
    
  lngBufferSize = 255
  strUser = Space$(lngBufferSize)

  glngReturnStatus = GetUserName(strUser, lngBufferSize)
  If glngReturnStatus = SUCCESS Then
    strUser = Left$(strUser, lngBufferSize - 1)
  Else
    Err = glngReturnStatus
  End If
  NetworkUserID = strUser
  Exit Function

NetworkUserID_EH:
  NetworkUserID = "ErrorInCall"
  Exit Function
  
End Function
Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
      SetTopMostWindow = False
   End If
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
      Txt$ = Txt$ + ClsMsg.LoadMsg(21) & vbLf & vbLf & DscErrors & vbLf
      Txt$ = Txt$ & ClsMsg.LoadMsg(22) & NumErrors
      lTitle = Errors(0).Source
      lHelpFile = Errors(0).HelpFile
      lHelpContext = Errors(0).HelpContext
   Else
      'O Seguinte erro ocorreu : "
      'Número : "
      Txt$ = Txt$ + ClsMsg.LoadMsg(21) & vbLf & vbLf & Dsc & vbLf
      Txt$ = Txt$ & ClsMsg.LoadMsg(22) & Num
      
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

   'Call GravaError(CStr(TxtAux), TxtTela$, NumErrors, DscErrors)
   On Error GoTo 0
'   Errors.Refresh
End Sub
Public Function SendTab(frm As Object, ByVal Key As Integer, Optional Tipo As Variant, _
                        Optional Obj As Variant, Optional Maiuscula = True, _
                        Optional Tamanho As Integer = 13, _
                        Optional Qtd_Dec As Integer = 2) As Integer
   '================================================================
   '= Última Alteração : 08/12/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Trata os caracteres de digitação e permiti a    **
   '**            validação do ActiveControl com [Enter]          **
   '**                                                            **
   '** Recebe: Key  - Código ASCII do caracter digitado.          **
   '**         Tipo - Tipo de Dados                               **
   '**                será apagado.                               **
   '**                                                            **
   '** Retorna:Key  - Caracter convertido em maiúsculo.           **
   '**                                                            **
   '****************************************************************
   On Error GoTo Fim
   Dim bNum As Boolean
   
   If IsMissing(Tipo) Then Tipo = -1
   If Key% = vbKeyCancel Or Key = 22 Then
      SendTab = Key%
      Exit Function
   End If
   If IsMissing(Obj) Then Set Obj = frm.ActiveControl
   'Key% = IIf(Key% = 39, 34, Key%) 'Trocar apóstrofo por aspas
   If Key% = 13 Then
      Select Case UCase(TypeName(frm.ActiveControl))
         Case "TEXTBOX"
            If Not frm.ActiveControl.MultiLine Then
               SendKeys "{TAB}"
            End If
         Case "OPTIONBUTTON", "MASKEDBOX", "COMBOBOX", "CHECKBOX", "SSOPTION", "SSCHECK"
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
                  If xVal(Obj.Text) = 0 Then
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
            SendTab = TratarMoeda(Key%, Ctrl, Tamanho, Qtd_Dec)
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

   If SendTab <> 13 Then
      Call SetTag(frm, "SUJA", "True")
   End If
   
   Exit Function
Fim:
   Call ShowError("SendTab")
End Function
Function xVal(ByVal Num$, Optional NumCasaDec = 5)
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
Public Function StrReplace(ByVal TxtIn As String, ByVal TxtFrom As String, ByVal TxtTo As String) As String
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
   Dim TxtOut$, LenIn%, LenFrom%, Pos%

   LenIn = Len(TxtIn)
   LenFrom = Len(TxtFrom)
   If LenFrom < 1 Or LenIn < 1 Then
      StrReplace = TxtIn
      Exit Function
   End If
   TxtOut = ""
   Pos = InStr(TxtIn, TxtFrom)
   While Pos > 0
      TxtOut = TxtOut + Left(TxtIn, Pos - 1) + TxtTo
      TxtIn = Right(TxtIn, Len(TxtIn) - Pos - LenFrom + 1)
      Pos = InStr(TxtIn, TxtFrom)
   Wend
   TxtOut = TxtOut + TxtIn
   StrReplace = TxtOut
End Function
Public Function TratarMoeda$(Key%, ByRef Obj As Object, Optional Tamanho As Integer, Optional Qtd_Dec As Integer = 2)
   Dim qtCasaDecimal$, Max$, Numero$
   Dim i%, NumMax#, TamMaxNum%, TamMaxCarac%, Qtd_Ponto%

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
         If (Trim$(Numero)) = mvarSepDec$ Then
            Numero = 0
         End If
         If CDbl(Numero) >= NumMax# Then
            For i = 1 To Len(Trim$(Numero))
               If Mid$(Trim$(Numero), i, 1) = mvarSepDec$ Then
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
      Case Asc(mvarSepDec$)
         TratarMoeda = Key
         For i = 1 To Len(Trim$(Numero))
             If Mid$(Trim$(Numero), i, 1) = mvarSepDec$ Then
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

Public Function EliminarString(ByVal Palavra$, ByVal Caracter$, Optional CaseSensitive = True) As String
    Dim Pos%, Com_Carac$
    
    If CaseSensitive Then
       Pos% = InStr(Palavra, Caracter)
    Else
      Pos% = InStr(UCase(Palavra), UCase(Caracter))
    End If
    
    Com_Carac = Palavra$
    While Pos% <> 0
        Com_Carac = Left$(Com_Carac, Pos% - 1) + Mid$(Com_Carac, Pos% + Len(Caracter$))
        If CaseSensitive Then
           Pos% = InStr(Com_Carac$, Caracter)
        Else
           Pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
        End If
    Wend

    If CaseSensitive Then
       Pos% = InStr(Com_Carac$, Caracter)
    Else
       Pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
    End If
    
    If Pos% > 0 Then
       If CaseSensitive Then
          Pos% = InStr(Pos% + Len(Caracter$), Com_Carac, Caracter)
       Else
          Pos% = InStr(Pos% + Len(Caracter$), UCase(Com_Carac), UCase(Caracter))
       End If
       If Pos% > 0 Then Com_Carac = Left$(Com_Carac, Pos% - 1)
    End If

    EliminarString = Com_Carac
End Function
Public Sub SelecionarTexto(ByRef Obj As Object)
'================================================================
'= Última Alteração : 02/01/98                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Selecionar todo texto do objeto que receber o   **
'**            foco, tal função geralmente é usado no evento   **
'**            GotFocus                                        **
'**                                                            **
'** Recebe:    Obj - Objeto a ser selecionado                  **
'**                                                            **
'** Retorna:   objeto selecionado.                            **
'**                                                            **
'****************************************************************
   Dim Tam As Long
   
   On Error Resume Next
   If (UCase(TypeName(Obj)) = "TEXTBOX") Or (UCase(TypeName(Obj)) = "MASKEDBOX") Or (UCase(TypeName(Obj)) = "COMBOBOX") Then
      Tam = Len(Obj)
      If UCase(TypeName(Obj)) = "MASKEDBOX" Then
         Tam = IIf(Len(Obj.Mask) > Tam, Len(Obj.Mask), Tam)
      End If
      Obj.SelStart = 0
      Obj.SelLength = Tam + 1
   End If
End Sub
Public Function ExisteItem(MyColl As Object, Item As String) As Boolean
   Dim x  As Variant
   
   Err = 0
   On Error Resume Next
   x = MyColl(Item)
   If Err = 438 Then
      Err = 0
      Set x = MyColl(Item)
   End If
   If (Err = 0) Then
      ExisteItem = (MyColl.Count > 0)
   Else
      ExisteItem = (Err = 0)
   End If
End Function

