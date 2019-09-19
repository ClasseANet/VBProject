Attribute VB_Name = "XDSR"
''************************************************************************************
''* 17 FUNÇÕES DE DSR100.DLL                                                         *
''************************************************************************************
''* Public Function SendTab(Frm As Object, ByVal Key As Integer, Optional Tipo As Variant, Optional Obj As Variant) As Integer
''* Public Sub ConfigForm(Frm As Object, Optional FrmIcone = "", Optional FundoTela = "FUNDO", Optional Centrar = True, Optional Pintar = True, Optional BtnIcone = True)
''* Public Sub BotaoIcon(Frm As Object, Optional Cursor = "PRESS", Optional CursorID = "LUPA")
''* Public Sub CentrarForm(MDI_name As Object, form_name As Object)
''* Public Function LoadMsg(Num As Integer) As String
''* Public Sub ShowError(Optional TxtAux = "")
''* Public Sub GravaError(TxtAux$, TxtTela$)
''* Public Sub Del(Arq$)
''* Public Function Copy(Orig$, dest$)
''* Public Function TratarMoeda$(Key%, ByRef Obj As Object)
''* Public Sub PintarFundo(Img As Object, Optional FundoTela = "FUNDO", Optional Frm = Nothing)
''* Function FileExists(ByVal strPathName As String) As Boolean
''* Public Sub LoadPctMouse(Key, Tipo, ByRef Ctrl As Object)
''* Private Function MakePathAux(ByVal strDirName As String) As Boolean
''* Public Function MakePath(ByVal strDir As String, Optional ByVal fAllowIgnore) As Boolean
''* Public Sub ClsCtrl.Set_Focus(ByVal objeto As Object)
''* Public Sub HabilitarBotao(ByRef objeto As Object, ByVal Bool%, Optional Pct As Variant)
''* Public Sub HabilitarObj(ByRef objeto As Object, ByVal Bool%, Optional ByRef Mnu)
''* Public Sub SelecionarTexto(ByRef Obj As Object)
''* Public Sub ExibirAviso(Txt As String, Optional Tit)
''* Public Function Data_Padrao_Windows$(ByVal DIA%, ByVal Mes%, ByVal Ano%)
''* Public Function EliminarString(ByVal Palavra$, ByVal Caracter$, Optional CaseSensitive = True) As String
''* Public Function InArray(Valor As Variant, VETOR As Variant)
''* Public Sub MontarTreeView(ByRef Dbase As Object, ByRef Tree As Object, Sql$, Root%, Optional IdRoot = "R", Optional DscRoot = "Sistema", Optional LenKEy = 3, Optional Imagem = "", Optional cTAG, Optional Flood = "", Optional LblFlood = "")
''* Public Sub MontarDbCombo(ByRef Dbase As Object, ByRef Combo As Object, Sql As String, Dsc As String, Optional Id = "", Optional IdView = False, Optional ComboAux, Optional Limpa = True)
''* Public Function clsdsr.StrZero(Valor As Variant, Num As Integer, Optional Caracter = "0") As String
''* Public Function LocalizarCombo(Cmb As Object, Chave As String, Optional SetCombo = True) As Integer
''* Public Sub clsdsr.LimparTela(Frm As Object)
''* Public Sub ZerarMask(ByRef Msk As Object, Optional ByRef MAscara)
''* Public Sub ExibirStop(Txt$, Optional Tit$)
''* Public Function SqlStr(Txt As String) As String
''* Public Sub ExibirInformacao(Txt$, Optional Tit$)
''* Public Function ExibirPergunta(Txt As String, Optional Tit As Variant) As Integer
''* Public Sub CentrarObj(ObjMain As Object, ObjChild As Object, Optional Tip)
''* Private Sub SetPctOrder(Grd As Object, Sql$)
''* Public Function MontarMSGrid(ByRef DataControl As Object, ByRef Grd As Object, Arr As Variant, Sql As String, Optional ByVal GrdTam = 0) As Boolean
''* Public Function PesquisarMSGrid%(Grd As Object, Chave$, Coluna%, Optional MatchCase = False)
''* Public Function MSGrdProcurar(ByRef DTC As Object, ByRef Grd As Object, ByRef Pnl As Object, Key) As String
''* Public Sub SelRowMSGrid(Grd As Object, Lin%)
''* Public Sub RefreshMSGrid(ByRef DataControl As Object, ByRef Grd As Object)
''* Public Sub OrdenarMSGrid(DTC As Object, Grd As Object, Ind%)
''* Public Sub SetHourglass(hWnd As Long)
''* Public Sub SetDefault(hWnd As Long)
''* Public Function LoadOriMsg$(Num%)
''* Public Sub MontarCombo(ByRef Combo As Object, Coll As Collection, Optional Propiedade = "") ' As Property)
''* Public Function StrReplace(ByVal TxtIn As String, ByVal TxtFrom As String, ByVal TxtTo As String) As String
''* Public Function String_Sem_Acento(str) As String
''* Public Function Between(Vl, Min, Max)
''* Public Sub SetPicture(Controle As Object, Key$, Optional Tipo = vbResBitmap)
''* Public Function SetFormatDT_Number() As Boolean
'
'Public Const CB_FINDSTRING = &H14C                  ' Used to search a Combo
'Public Const LB_FINDSTRING = &H18F                  ' Used to search a List box
'Public Const CB_SHOWDROPDOWN = &H14F
'Public Const CB_GETITEMHEIGHT = &H154
'
''* Cursor
'Public Const IDC_WAIT = 32514&   ' Hourglass
'Public Const IDC_ARROW = 32512&
'Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long
'Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function ReleaseCapture Lib "user32" () As Long
'
'Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
''************************************************************************************
''****************************************
''* Variáveis globais e funções          *
''* para suprir a ausência de DSR100.DLL *
''****************************************
'Global gDrvTmp As String
'Global gDrvErro As String
'Global gSepDec As String
'Global gSepMil As String
'Global gSepDt As String
'Global gDtMask As String
'Global gComputerName As String
'Global gIdioma As Integer
'Public Function SendTab(Frm As Object, ByVal Key As Integer, Optional Tipo As Variant, Optional Obj As Variant) As Integer
''================================================================
''= Última Alteração : 08/12/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Trata os caracteres de digitação e permiti a    **
''**            validação do ActiveControl com [Enter]          **
''**                                                            **
''** Recebe: Key  - Código ASCII do caracter digitado.          **
''**         Tipo - Tipo de Dados                               **
''**                será apagado.                               **
''**                                                            **
''** Retorna:Key  - Caracter convertido em maiúsculo.           **
''**                                                            **
''****************************************************************
'    On Error GoTo fim
'    If IsMissing(Tipo) Then Tipo = -1
'    Key% = IIf(Key% = 39, 34, Key%) 'Trocar apóstrofo por SqlStr
'    If Key% = 13 Then
'       Select Case TypeName(Frm.ActiveControl)
'          Case "TextBox"
'             If Not Frm.ActiveControl.MultiLine Then
'                SendKeys "{TAB}"
'             End If
'           Case "OptionButton", "MaskEdBox", "ComboBox", "CheckBox", "SSOption", "SSCheck"
'             SendKeys "{TAB}"
'       End Select
'       SendTab = (Key%)
'    Else
'        If (Key < 48 Or Key > 57) And Key <> 8 And Tipo = vbInteger Then
'           Key = Asc(" ")
'           Beep
'           Exit Function
'        End If
'        SendTab = Asc(UCase(Chr$(Key%)))
'        If Tipo <> -1 Then
'           If Tipo = vbCurrency Then
'              Dim Ctrl As Object
'              Set Ctrl = Obj
'              SendTab = TratarMoeda(Key%, Ctrl)
'           End If
'        Else
'           SendTab = Asc(UCase(Chr$(Key%)))
'        End If
'    End If
'    On Error Resume Next
'    Frm.Suja = Frm.Suja Or (SendTab <> 13)
'
'    If Err = 438 Or Err = 0 Then
'       On Error GoTo 0
'    Else
'       ShowError
'    End If
'Exit Function
'fim:
'   Call ShowError("DSR100.DSR.SendTab")
'End Function
'Public Sub ConfigForm(Frm As Object, Optional FrmIcone = "", Optional FundoTela = "FUNDO", Optional Centrar = True, Optional Pintar = True, Optional BtnIcone = True)
'   Call ResizeForm(Frm)
'   If Centrar Then Call CentrarForm(Screen, Frm)
'   If Pintar And FundoTela <> "" Then CallClsCtrl.PintarFundo(Frm.ImgFundo, FundoTela)
'   If BtnIcone Then Call BotaoIcon(Frm)
'   If FrmIcone <> "" Then Frm.Icon = FrmIcone
'End Sub
'Public Sub CentrarForm(MDI_name As Object, form_name As Object)
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Centralizar um Form em outro form base.         **
''**                                                            **
''** Recebe:  MDI_name  - MDIForm base                          **
''**          form_name - Form a ser centralizado               **
''**                                                            **
''** Retorna: Form centralizado                                 **
''**                                                            **
''****************************************************************
'   Dim HH!, LL!, tt!, WW!
'   Dim Dif%, i%
'   On Error Resume Next
'   If form_name.WindowState = vbMaximized Then Exit Sub
'   WW! = form_name.Width
'   HH! = form_name.Height
'
'   If form_name.MDIChild And TypeName(MDI_name) <> "Screen" Then
'      tt! = (MDI_name.ScaleHeight - HH!) / 2
'      LL! = (MDI_name.ScaleWidth - WW!) / 2
'   Else
'     tt! = (MDI_name.Height - HH!) / 2
'     LL! = (MDI_name.Width - WW!) / 2
'     If form_name.MDIChild Then
'        tt! = tt! - 1000
'        LL! = LL! + 60
'     End If
'   End If
'   form_name.Move LL!, tt!
'End Sub
'Public Sub BotaoIcon(Frm As Object, Optional Cursor = "PRESS", Optional CursorID = "LUPA")
'   Dim i%
''   On Error Resume Next
'   For i% = 0 To Frm.Controls.Count - 1
'      Select Case UCase(TypeName(Frm.Controls(i)))
'         Case "SSCOMMAND", "COMMANDBUTTON", "TOOLBAR"
'            Frm.Controls(i).MousePointer = vbCustom
'            Frm.Controls(i).MouseIcon = LoadResPicture(Cursor, vbResCursor)
'         Case "LABEL"
'            If UCase(Frm.Controls(i).Name) = "LBLID" And Frm.Controls(i).MousePointer <> vbCustom Then
'               Frm.Controls(i).MousePointer = vbCustom
'               Frm.Controls(i).MouseIcon = LoadResPicture(CursorID, vbResCursor)
'               Frm.Controls(i).ForeColor = 16711680
'            End If
'      End Select
'   Next
'End Sub
'Public Function LoadMsg(Num As Integer) As String
'   Dim Idioma%, msg$
'   On Error GoTo fim
'   '* Idioma usado
'   mvarIdioma% = IIf(mvarIdioma% = 0, 5000, mvarIdioma)
'   Idioma% = mvarIdioma% 'Português = 5000, Inglês = 6000, ...
'
'   'Num% = Idioma% + Num%
'   'Select Case Idioma%
'   '   Case 5000: Msg$ = LoadResString(Num) 'LoadPortugues$(Num%)
'   '   Case 6000: Msg$ = LoadResString(Num) 'LoadIngles$(Num%)
'   'End Select
'   'LoadMsg$ = Msg$
'   If Num > 9999 Then
'      Num = 5000 + (Num - (Val(Mid(CStr(Num), 1, Len(CStr(Num)) - 3)) * 1000))
'   End If
'   LoadMsg$ = LoadResString(mvarIdioma + Num)
'   Exit Function
'fim:
'   ShowError
'End Function
'Public Sub ShowError(Optional TxtAux = "")
'   Dim Txt$, Num&, Dsc$
'   Dim TxtTela$
'   If Err = 0 Then Exit Sub
'   Errors.Refresh
'   Num = Err
'   Dsc = Error
'   Screen.MousePointer = vbDefault
'   On Error Resume Next
'   'Anote o conteúdo da mensagem abaixo e avise ao analista responsável.
'   Txt$ = Txt$ & ClsMSG.LoadMsg(9) & vbLf & vbLf
'   'O Seguinte erro ocorreu : "
'   If Num = Errors(0).Number Then
'      Txt$ = Txt$ + clsmsg.LoadMsg(21) & vbLf & vbLf & Errors(0).Description & vbLf
'      'Número : "
'      Txt$ = Txt$ & ClsMSG.LoadMsg(22) & Errors(0).Number
'   Else
'      Txt$ = Txt$ + clsmsg.LoadMsg(21) & vbLf & vbLf & Dsc & vbLf
'      If Errors(0).Number <> 0 Then
''         Txt$ = Txt$ & Errors(0).Description & vbLf
'      End If
'      'Número : "
'      Txt$ = Txt$ & ClsMSG.LoadMsg(22) & Num
'   End If
''   Beep
'   TxtTela = ""
'   If Not Screen.ActiveForm Is Nothing Then
'      TxtTela = "Tela : " & Screen.ActiveForm.Name
'      If Not Screen.ActiveForm.ActiveControl Is Nothing Then
'         TxtTela = TxtTela & "." & Screen.ActiveForm.ActiveControl.Name
'         If Screen.ActiveForm.ActiveControl.index <> "" Then
'            TxtTela = TxtTela & "(" & CStr(Screen.ActiveForm.ActiveControl.index) & ")"
'         End If
'      End If
'   End If
'   If TxtAux <> "" Then
'      Txt = Txt & vbLf & "Auxiliar : " & vbLf & TxtAux
'   End If
'   MsgBox Txt, vbMsgBoxHelpButton + vbExclamation, Errors(0).Source, Errors(0).HelpFile, Errors(0).HelpContext
'
''   Call MsgBox(Txt$, 48, Errors(0).Source)
'
'   '***************
'   '* Gravar Erro no Arquivo de Log
'   '***************
'  Call GravaError(CStr(TxtAux), TxtTela$)
'  On Error GoTo 0
'  Errors.Refresh
'End Sub
'Public Sub GravaError(TxtAux$, TxtTela$)
'   Dim ArqLog$, ArqTmp$
'   Dim pos%, Limpar As Boolean
'   Dim Txt$, i%
'   ArqLog$ = "Error.log"
'   ArqTmp$ = "Error.tmp"
'   mvarDrvErro = DrvTmp
'   If Not FileExists(mvarDrvErro & ArqLog$) Then
'      Open mvarDrvErro & ArqLog$ For Output As #1
'         Print #1, "0 Erros"
'      Close #1
'   End If
'   Limpar = (FileLen(mvarDrvErro & ArqLog$) > 1000000) ' > 1 Mb
'   Call Copy(mvarDrvErro & ArqLog$, mvarDrvErro & ArqTmp$)
'   Open mvarDrvErro & ArqLog$ For Output As #1
'   Open mvarDrvErro & ArqTmp$ For Input As #2
'      Line Input #2, Txt
'      If Limpar Then
'         Print #1, "4 Erros"
'      Else
'         pos = 0
'         While pos = 0
'            pos = InStr(Txt, "Erros")
'            If pos > 2 Then
'               If pos > 5 Then
'                  Print #1, CStr(CLng(Mid(Txt, pos - 5, 5)) + 1) & " Erros"
'               Else
'                  Print #1, CStr(CLng(Mid(Txt, 1, pos - 2)) + 1) & " Erros"
'               End If
'            End If
'            If Not EOF(2) And pos = 0 Then Line Input #2, Txt
'            pos = IIf(EOF(2), 1, pos)
'         Wend
'      End If
'      Print #1, "====================================================================="
'      Print #1, "Date     : " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
'      Print #1, "Computer : " & ComputerName
'      Print #1, "Source   : " & Errors(0).Source
'      Print #1, "Erro     : " & Errors(0).Number & " - " & Mid(Errors(0).Description, 1, 50)
'      Txt = Errors(0).Description
'      While Len(Txt) >= 50
'         Txt = Mid(Txt, 51)
'         Print #1, Space(18) & Mid(Txt, 1, 50)
'      Wend
'      Print #1, "Help     : " & Errors(0).HelpContext & " - " & Errors(0).HelpFile
'      If TxtTela <> "" Then
'         Print #1, TxtTela
'      End If
'      If TxtAux <> "" Then
'         Print #1, "Auxiliar : " & TxtAux
'      End If
'      Print #1, "====================================================================="
'      Do While Not EOF(2)
'         Line Input #2, Txt
'         Print #1, Txt
'         If Limpar And Mid(Txt, 1, 5) = "=====" Then '* Se Arquivo > 1Mb
'            i = i + 1
'            If i = 6 Then Exit Do '* Gravar apenas os 3 últimos Erros
'         End If
'      Loop
'   Close #2
''   Print #1, "====================================================================="
'   Close #1
'   Call Del(mvarDrvErro & ArqTmp$)
'End Sub
'Public Sub Del(Arq$)
'   If FileExists(Arq$) Then
'      On Error GoTo fim
'      Kill Arq$
'   End If
'   Exit Sub
'fim:
'   ShowError
'End Sub
'Public Function Copy(Orig$, dest$)
'   Dim nMsg$, nTipo&, NL
'   Dim Resp%
'   NL = vbLf
'   On Error Resume Next
'   If FileExists(Orig$) Then
'      Call Del(dest$)
'      FileCopy Orig$, dest$
'   Else
'      Call ExibirAviso(LoadMsg(11) + UCase(Orig$), ClsMSG.LoadMsg(12))
'      Resp = vbCancel
'      Exit Function
'   End If
'   Resp = vbYes
'   Select Case Err
'      Case 71
'         While Resp = vbYes
'            nMsg = clsmsg.LoadMsg(13) + NL + NL
'            nMsg = nMsg & ClsMSG.LoadMsg(14) + NL
'            nMsg = nMsg & ClsMSG.LoadMsg(15)
'            nTipo = vbYesNo + vbCritical + vbDefaultButton1
'            Resp = MsgBox(nMsg, nTipo, ClsMSG.LoadMsg(16))
'            If Resp = vbYes Then
'               Err = 0
'               FileCopy Orig$, dest$
'            End If
'         Wend
'      Case 70
'         While Resp = vbOK
'            nMsg = clsmsg.LoadMsg(7) + NL + NL
'            nMsg = nMsg & ClsMSG.LoadMsg(56) + NL
'            nTipo = vbYesNo + vbCritical + vbDefaultButton1
'            Resp = MsgBox(nMsg, nTipo, ClsMSG.LoadMsg(16))
'            If Resp = vbYes Then
'               Err = 0
'               FileCopy Orig$, dest$
'            End If
'         Wend
'   End Select
'   Copy = Resp
'End Function
'Public Function TratarMoeda$(Key%, ByRef Obj As Object)
'   Dim qtCasaDecimal$, Max$, Numero$
'   Dim i%, NumMax#, TamMaxNum%, TamMaxCarac%, Qtd_dec%, Qtd_Ponto%
'
'   Qtd_dec% = 2
'   Qtd_Ponto = Int((Obj.MaxLength - Qtd_dec - 1) / 3) - 1
'   Qtd_Ponto = IIf(Qtd_Ponto < 0, 0, Qtd_Ponto)
'   Max = ""
'   TamMaxNum% = Obj.MaxLength - (Qtd_dec + Qtd_Ponto + 2)
'   TamMaxCarac% = TamMaxNum% + Qtd_dec + 1
'   For i = 1 To TamMaxNum%
'      Max = Max + "9"
'   Next
'
'   NumMax# = CDbl(Max)
'   Numero = Obj.Text
'   Select Case Key
'      Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57:
'         If Len(Trim$(Numero)) = 0 Then
'            Numero = 0
'         End If
'         If (Trim$(Numero)) = mvarSepDec$ Then
'            Numero = 0
'         End If
'         If CDbl(Numero) >= NumMax# Then
'            For i = 1 To Len(Trim$(Numero))
'               If Mid$(Trim$(Numero), i, 1) = mvarSepDec$ Then
'                  Beep
'                  TratarMoeda = 0
'                  Exit For
'               End If
'            Next i
'            qtCasaDecimal = Len(Trim$(Numero)) - i
'            If Not (qtCasaDecimal > -1 And _
'                   (Len(Trim$(Numero)) - 1) < TamMaxCarac%) Then
'               Beep
'               TratarMoeda = 0
'               Exit Function
'            End If
'         End If
'         TratarMoeda = Key
'      Case 8  'TAb
'         TratarMoeda = Key
'      Case Asc(mvarSepDec$)
'         TratarMoeda = Key
'         For i = 1 To Len(Trim$(Numero))
'             If Mid$(Trim$(Numero), i, 1) = mvarSepDec$ Then
'                Beep
'                TratarMoeda = 0
'                Exit For
'             End If
'         Next i
'      Case Else
'         Beep
'         TratarMoeda = 0
'   End Select
'End Function
'Public Sub ResizeForm(Frm As Object)
'   Dim wStart As Integer
'   Dim hStart As Integer
''   Dim wFactor As Single
''   Dim hFactor As Single
'   Dim LoopIndex As Integer
'   Dim Ctrl As Object
'   Dim Ind As Byte
'   Const DesignX = 800 '640
'   Const designY = 600 '480
'   wStart = Screen.Width / Screen.TwipsPerPixelX
'   hStart = Screen.Height / Screen.TwipsPerPixelY
'   mvarwFactor = wStart / DesignX
'   mvarhFactor = hStart / designY
'   If wStart = DesignX Then Exit Sub
'   On Error Resume Next
'   For LoopIndex = 0 To Frm.Controls.Count - 1
'      Set Ctrl = Frm.Controls(LoopIndex)
'      Ctrl.Left = Ctrl.Left * wFactor
'      Ctrl.Width = Ctrl.Width * wFactor
'      Ctrl.FontSize = Ctrl.FontSize * wFactor
'      Ctrl.Height = Ctrl.Height * hFactor
'      If TypeOf Ctrl Is MSFlexGrid Then 'SpreadSheet Then
'          Ctrl.Row = 1
'         For Ind = 1 To Ctrl.Cols
'            Ctrl.Col = Ind
'            Ctrl.ColWidth(Ind) = Ctrl.ColWidth(Ind) _
'            * (wFactor + 0.02)
'         Next
'         Ctrl.Col = 1
'         For Ind = 1 To Ctrl.Rows
'            Ctrl.Row = Ind
'            Ctrl.RowHeight(Ind) = Ctrl.RowHeight(Ind) _
'            * (hFactor - 0.1)
'         Next
'      End If
'      Ctrl.Top = Ctrl.Top * hFactor
'      Set Ctrl = Nothing
'   Next
'   Frm.Width = Frm.Width * wFactor
'   Frm.Height = Frm.Height * hFactor
'End Sub
'Public Sub PintarFundo(Img As Object, Optional FundoTela = "FUNDO", Optional Frm = Nothing)
'   Dim i%, J%, Tam!, Larg!
'   Dim WW&, HH&
'   On Error GoTo fim
'   If Frm Is Nothing Then Set Frm = Img.Parent
'
'   Img.BorderStyle = vbBSNone
'   If Img.Picture = 0 Then
'      Img.Picture = LoadResPicture(FundoTela, vbResBitmap)
'   End If
'   '*****
'   '* Definir Dimensão da Imagem
'   '*****
'   Tam = Img.Width
'   Larg = Img.Height
'   Larg = IIf(Larg < Tam, Larg, Tam)
'   Tam = IIf(Larg < Tam, Larg, Tam)
'   '*****
'   '* Definir Dimensão do Objeto
'   '*****
'   On Error Resume Next
'   WW = Frm.ScaleWidth
'   HH = Frm.ScaleHeight
'   If Err <> 0 Then WW = Frm.Width
'   If Err <> 0 Then HH = Frm.Height
'   Err = 0
'   '*****
'   '* Pintar Objeto
'   '*****
'   For i = 0 To Int(WW / Tam)
'      For J = 0 To Int(HH / Tam)
'         Frm.PaintPicture Img.Picture, i * Tam, J * Larg, Tam, Larg, 0, 0
'      Next
'   Next
'fim:
'   If Err = 3265 Then 'Resource with identifier 'qq' not found
'     If Idioma = 5000 Then
'         Call ExibirStop("Fundo de Tela """ & FundoTela & """ não existe.", ClsMSG.LoadMsg(1))
'     Else
'         Call ExibirStop("Resource with identifier " & FundoTela & " not found.", ClsMSG.LoadMsg(1))
'     End If
'     Err = 0
'   Else
'       ShowError
'   End If
'End Sub
'Public Function FileExists(ByVal strPathName As String) As Boolean
'    Dim intFileNum As Integer
'
'    On Error Resume Next
'
'    '
'    'Remove any trailing directory separator character
'    '
'    If Right$(strPathName, 1) = "\" Then
'        strPathName = Left$(strPathName, Len(strPathName) - 1)
'    End If
'
'    '
'    'Attempt to open the file, return value of this function is False
'    'if an error occurs on open, True otherwise
'    '
'    intFileNum = FreeFile
'    Open strPathName For Input As intFileNum
'
'    FileExists = IIf(Err, False, True)
'    FileExists = IIf(Err = 70, True, FileExists)
'
'    Close intFileNum
'
'    Err = 0
'End Function
'Public Sub LoadPctMouse(Key, Tipo, ByRef Ctrl As Object)
'   Ctrl.MousePointer = vbCustom
'   Ctrl.MouseIcon = LoadResPicture(Key, Tipo)
'End Sub
'Public Sub ExecuteLink(ByVal sLinkTo As String)
'
''* Execute o link to http://www.DSR.com.br/
''* (if possible) - or the new 'mailto:dramos@mandic.com.br'
'
'    On Error Resume Next
'
'    Dim lRet As Long
'    Dim lOldCursor As Long
'
'    lOldCursor = Screen.MousePointer
'
'    Screen.MousePointer = vbHourglass
'    lRet = ShellExecute(0, "open", sLinkTo, "", vbNull, SW_SHOWNORMAL)
'
'    If lRet >= 0 And lRet <= 0 Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Error Opening Link to " & sLinkTo & vbCrLf & vbCrLf & Err.LastDllError, , "frmAbout::ExecuteLink"
'    End If
'    Screen.MousePointer = vbDefault
'
'End Sub
'Public Function AppAtiva(Aplic As Object)
'   Dim ERRO%, Estilo%
'   Dim Txt$
'    If Aplic.PrevInstance Then
'        On Error Resume Next
'        'tenta ativar aplicação que já estava rodando (caption de MDIForm)
'        AppActivate Aplic.ExeName
'        ERRO% = Err
'        On Error GoTo 0
'        'testa erro (aplicação pode estar com o foco em um form não MDIChild)
'        If ERRO% <> 0 Then
'           Txt$ = "Já existe uma cópia da aplicação rodando."
'           Txt$ = Txt$ + Chr$(10) + "Pressione Ok para fechar esta mensagem"
'           Txt$ = Txt$ + Chr$(10) + "e Alt+Tab para localizar a aplicação."
'            Estilo% = vbSystemModal + vbExclamation
'            Screen.MousePointer = vbDefault
'            MsgBox Txt$, Estilo%, Aplic.Description
'            DoEvents
'        Else
'            'maximiza aplicação que já estava rodando
'            SendKeys "% X"
'        End If
'        'encerra aplicação cópia
'        AppAtiva = True
'    End If
'End Function
'Public Function MakePath(ByVal strDir As String, Optional ByVal fAllowIgnore) As Boolean
'    If IsMissing(fAllowIgnore) Then
'        fAllowIgnore = True
'    End If
'
'    Do
'        If MakePathAux(strDir) Then
'            MakePath = True
'            Exit Function
'        Else
'            Dim strMsg As String
'            Dim iRet As Integer
'
''            strMsg = ResolveResString(resMAKEDIR) & LF$ & strDir
'            iRet = MsgBox(strMsg, IIf(fAllowIgnore, vbAbortRetryIgnore, vbRetryCancel) Or vbExclamation Or vbDefaultButton2, "")
'            Select Case iRet
'            Case vbAbort, vbCancel
''                ExitSetup frmCopy, gintRET_ABORT
'            Case vbIgnore
'                MakePath = False
'                Exit Function
'            Case vbRetry
'            End Select
'        End If
'    Loop
'End Function
''-----------------------------------------------------------
'' FUNCTION: MakePathAux
''
'' Creates the specified directory path.
''
'' No user interaction occurs if an error is encountered.
'' If user interaction is desired, use the related
''   MakePathAux() function.
''
'' IN: [strDirName] - name of the dir path to make
''
'' Returns: True if successful, False if error.
''-----------------------------------------------------------
''
'Private Function MakePathAux(ByVal strDirName As String) As Boolean
'    Dim strPath As String
'    Dim intOffset As Integer
'    Dim intAnchor As Integer
'    Dim strOldPath As String
'
'    On Error Resume Next
'
'    '
'    'Add trailing backslash
'    '
'   If Right$(strDirName, 1) <> "\" Then
'        strDirName = strDirName & "\"
'    End If
'
'    strOldPath = CurDir$
'    MakePathAux = False
'    intAnchor = 0
'
'    '
'    'Loop and make each subdir of the path separately.
'    '
'    '
'    intOffset = InStr(intAnchor + 1, strDirName, "\")
'    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
'    Do
'        intOffset = InStr(intAnchor + 1, strDirName, "\")
'        intAnchor = intOffset
'
'        If intAnchor > 0 Then
'            strPath = Left$(strDirName, intOffset - 1)
'            ' Determine if this directory already exists
'            Err = 0
'            ChDir strPath
'            If Err Then
'                ' We must create this directory
'                Err = 0
'                #If Win32 And LOGGING Then
'                    NewAction gstrKEY_CREATEDIR, """" & strPath & """"
'                #End If
'                MkDir strPath
'                #If Win32 And LOGGING Then
'                    If Err Then
'                        LogError ResolveResString(resMAKEDIR) & " " & strPath
'                        AbortAction
'                        GoTo Done
'                    Else
'                        CommitAction
'                    End If
'                #End If
'            End If
'        End If
'    Loop Until intAnchor = 0
'
'    MakePathAux = True
'Done:
'    ChDir strOldPath
'
'    Err = 0
'End Function
'Public Sub ClsCtrl.Set_Focus(ByVal objeto As Object)
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Obter Foco para um determinado objeto, se ele   **
''**            estiver inativo se tornará ativo e visível.     **
''**                                                            **
''** Recebe: objeto - Objeto a receber o foco                   **
''**                                                            **
''** Retorna : Objeto focado                                    **
''**                                                            **
''****************************************************************
'   On Error GoTo fim
'   DoEvents
'   If objeto.Enabled = True And objeto.Visible = True Then
'      objeto.SetFocus
'   Else
'      Call HabilitarObj(objeto.Parent, True)
'      Call HabilitarObj(objeto, True)
'      objeto.SetFocus
'   End If
'Exit Sub
'fim:
'   ShowError
'End Sub
'Public Sub HabilitarBotao(ByRef objeto As Object, ByVal Bool%, Optional Pct As Variant)
'   On Error GoTo fim
'   If Not IsMissing(Pct) Then
'      objeto.Caption = ""
'      Select Case UCase(TypeName(Pct))
'         Case "STRING": objeto.Picture = LoadPicture(Pct)
'         Case "INTEGER": objeto.Picture = LoadResPicture(Pct, vbResBitmap)
'         Case "PICTURE", "LONG": objeto.Picture = Pct
'      End Select
'   End If
'   objeto.MousePointer = IIf(Bool, vbDefault, vbNoDrop)
'Exit Sub
'fim:
'   ShowError
'End Sub
'Public Sub HabilitarObj(ByRef objeto As Object, ByVal Bool%, Optional ByRef Mnu)
'   objeto.Enabled = Bool%
'   objeto.Visible = Bool%
'   If Not IsMissing(Mnu) Then Mnu.Enabled = Bool
'End Sub
'Public Sub SelecionarTexto(ByRef Obj As Object)
''================================================================
''= Última Alteração : 02/01/98                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Selecionar todo texto do objeto que receber o   **
''**            foco, tal função geralmente é usado no evento   **
''**            GotFocus                                        **
''**                                                            **
''** Recebe:    Obj - Objeto a ser selecionado                  **
''**                                                            **
''** Retorna:   objeto selecionado.                            **
''**                                                            **
''****************************************************************
'   If (TypeName(Obj) = "TextBox") Or (TypeName(Obj) = "MaskEdBox") Then
'      Obj.SelStart = 0
'      Obj.SelLength = Len(Obj) + 1
'   End If
'End Sub
'Public Sub ExibirAviso(Txt As String, Optional Tit)
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Exibir um aviso na tela.                        **
''**                                                            **
''** Recebe: Mensagem$ - Aviso a ser exibido                    **
''**         Tit$   - Título da Mensagem (Opcional)          **
''**                                                            **
''** Retorna : Aviso centralizado na tela com um botão de OK e  **
''**           um ícone de Exclamação.                          **
''**                                                            **
''****************************************************************
'   Dim Mouse%
'   Mouse = Screen.MousePointer
'   If TypeName(Tit) = "Nothing" Then
'      Tit = clsmsg.LoadMsg(1)
'   End If
'   Screen.MousePointer = vbDefault
'   MsgBox Txt$, vbExclamation, Tit
'   DoEvents
'   Screen.MousePointer = vbDefault
'End Sub
'Public Function Data_Padrao_Windows$(ByVal DIA%, ByVal Mes%, ByVal Ano%)
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Recuperar o formato de data padrão que o windows**
''**            está utilizando.                                **
''** Recebe:  dia% - Dia atual.                                 **
''**          Mes% - Mês atual.                                 **
''**          Ano% - Ano atual.                                 **
''**                                                            **
''** Retorna: Formato de data do Sistema Operacional.           **
''**                                                            **
''****************************************************************
'    mvarSepDt$ = "/"
'    Select Case mvarDtMask$
'        Case "dd/mm/yyyy"
'            Data_Padrao_Windows$ = Format$(DIA%, "00") + mvarSepDt$ + Format$(Mes%, "00") + mvarSepDt$ + Right$(CStr(Ano%), 4)
'        Case "mm/dd/yyyy"
'            Data_Padrao_Windows$ = Format$(Mes%, "00") + mvarSepDt + Format$(DIA%, "00") + mvarSepDt + Right$(CStr(Ano%), 4)
'        Case "yyyy/mm/dd"
'            Data_Padrao_Windows$ = Right$(CStr(Ano%), 4) + mvarSepDt + Format$(Mes%, "00") + mvarSepDt + Format$(DIA%, "00")
'        Case Else
'            Data_Padrao_Windows$ = Format$(DIA%, "00") + mvarSepDt$ + Format$(Mes%, "00") + mvarSepDt$ + Right$(CStr(Ano%), 4)
'    End Select
'End Function
'Public Function EliminarString(ByVal Palavra$, ByVal Caracter$, Optional CaseSensitive = True) As String
'    Dim pos%, Com_Carac$
'
'    If CaseSensitive Then
'       pos% = InStr(Palavra, Caracter)
'    Else
'      pos% = InStr(UCase(Palavra), UCase(Caracter))
'    End If
'
'    Com_Carac = Palavra$
'    While pos% <> 0
'        Com_Carac = Left$(Com_Carac, pos% - 1) + Mid$(Com_Carac, pos% + Len(Caracter$))
'        If CaseSensitive Then
'           pos% = InStr(Com_Carac$, Caracter)
'        Else
'           pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
'        End If
'    Wend
'
'    If CaseSensitive Then
'       pos% = InStr(Com_Carac$, Caracter)
'    Else
'       pos% = InStr(UCase(Com_Carac$), UCase(Caracter))
'    End If
'
'    If pos% > 0 Then
'       If CaseSensitive Then
'          pos% = InStr(pos% + Len(Caracter$), Com_Carac, Caracter)
'       Else
'          pos% = InStr(pos% + Len(Caracter$), UCase(Com_Carac), UCase(Caracter))
'       End If
'       If pos% > 0 Then Com_Carac = Left$(Com_Carac, pos% - 1)
'    End If
'
'    EliminarString = Com_Carac
'End Function
'Public Function InArray(Valor As Variant, VETOR As Variant)
'   Dim J As Variant
'   InArray = False
'   For Each J In VETOR
'       If Valor = J Then
'         InArray = True
'         Exit For
'      End If
'   Next
'End Function
'Public Sub MontarTreeView(ByRef Dbase As Object, ByRef Tree As Object, Sql$, Root%, Optional IdRoot = "R", Optional DscRoot = "Sistema", Optional LenKEy = 3, Optional Imagem = "", Optional cTAG, Optional Flood = "", Optional LblFlood = "")
'   Dim NodX As Object
'   Dim Cod$, CodRoot$, Txt$, msg$
'   Dim dyTre As Recordset
'   Dim k%, Vet() As String
'   Dim NvZero%, Tam%
'   Dim Flood_Ini&, Flood_Fim&
'   On Error GoTo fim
'   ReDim Vet(0)
'   '* Configurações de Flood
'   If Flood <> "" Then
'      Flood_Ini = Flood.Value
'      Flood_Fim = CDbl(Flood.Tag)
'      Flood.Enabled = True
'      Flood.Visible = True
'      If TypeName(LblFlood) = "Label" Then
'         LblFlood.Enabled = True
'         LblFlood.Caption = Trim(CStr(Flood_Ini)) & " %"
'         LblFlood.Visible = True
'         LblFlood.Refresh
'      End If
'   End If
'   If Trim(Sql$) = "" Then
'      Sql$ = Tree.Tag
'   Else
'      Tree.Tag = Sql$
'   End If
''   If Root Then If DscRoot = "Sistema" Then DscRoot = App.Title
'
'   Set NodX = Tree.Nodes.Add(, , IdRoot, DscRoot)     ' Root
'   Tree.Nodes(Tree.Nodes.Count).Sorted = Tree.Sorted
'
'   If Imagem <> "" Then NodX.Image = Imagem
'
'   Call Dbase.AbreTabela(Sql$, dyTre)
'   If Not Dbase.CodeSql Then Exit Sub
'   If dyTre.EOF Then Exit Sub
'   dyTre.MoveFirst
'   NvZero = IIf(Root, dyTre(2), 1)
'   While Not dyTre.EOF
'      Cod$ = "'" + dyTre(0) + "'"
'      If dyTre(2) = NvZero Then    'Se foi definido Raíz e Nível
'         CodRoot = IdRoot
'      Else
'         On Error Resume Next
'         Tam% = dyTre(LenKEy)
'         If Err = 3265 Then Tam% = 2 'Item não encontrado na coleção
'         On Error GoTo 0
'         On Error GoTo fim
'         CodRoot = "'" & Mid(dyTre(0), 1, IIf(Len(dyTre(0)) <= Tam%, Len(dyTre(0)), Len(dyTre(0)) - Tam%)) & "'"
'      End If
'      Txt$ = dyTre(1)
'      On Error Resume Next
'                                         'tvwChild =4
'      Set NodX = Tree.Nodes.Add(CodRoot, 4, Cod$, Txt$)
'      Tree.Nodes(Tree.Nodes.Count).Sorted = Tree.Sorted
'      If Not IsMissing(cTAG) Then
'         Tree.Nodes(Tree.Nodes.Count).Tag = Tree.Nodes(Tree.Nodes.Count).Tag & "|" & UCase(cTAG) & "=" & dyTre(cTAG)
'         If Err = 3265 Then On Error GoTo 0 'Item not found in this collection.
'      End If
'      If Err <> 0 Then 'NÃO EXISTE O NÓ PAI ( CODROOT)
'         k = k + 1
'         ReDim Preserve Vet(3 * k)
'         Vet(3 * k) = Cod$
'         Vet((3 * k) - 2) = Txt$
'         Select Case Err
'            Case 35601: msg$ = "**Erro -> Item sem Nível Superior Correspondente"
'         End Select
'         Vet((3 * k) - 1) = msg$
'          On Error GoTo 0
'      End If
'      On Error GoTo fim
'      '*** Expand tree to see all nodes. ***
'      'nodX.EnsureVisible
'      If Imagem <> "" Then NodX.Image = Imagem
'      dyTre.MoveNext
'      If Not dyTre.EOF Then
'         If Flood <> "" Then
'            Flood.Value = CInt(Flood_Ini + ((dyTre.PercentPosition / 100) * (Flood_Fim - Flood_Ini)))
'            If TypeName(LblFlood) = "Label" Then
'               If Trim(CStr(Flood.Value)) & " %" <> LblFlood.Caption Then
'                  LblFlood.Caption = Trim(CStr(Flood.Value)) & " %"
'                  LblFlood.Refresh
'               End If
'            End If
'         End If
'      End If
'   Wend
'   dyTre.Close
'   If UBound(Vet) <> 0 Then
'      Set NodX = Tree.Nodes.Add(, , "ERRO", "ERROS")
'      For k = (LBound(Vet) + 1) To (UBound(Vet) / 3)
''         On Error Resume Next
'                                           'tvwChild=4
'         Set NodX = Tree.Nodes.Add("ERRO", 4, Vet(3 * k), Vet((3 * k) - 2) & "-" & Vet((3 * k) - 1))
'         Tree.Nodes(Tree.Nodes.Count).Sorted = Tree.Sorted
'      Next
'   End If
'   Tree.HideSelection = False
'   If Root Then Tree.Nodes(2).EnsureVisible
'   Tree.SelectedItem = Tree.Nodes(1)
'   Tree.Refresh
'Exit Sub
'fim:
'   ReDim Vet(0)
'   ShowError
'End Sub
'Public Sub MontarDbCombo(ByRef Dbase As Object, ByRef Combo As Object, Sql As String, Dsc As String, Optional Id = "", Optional IdView = False, Optional ComboAux, Optional Limpa = True)
''===================================================================
''= Última Alteração : 28/11/97                                     =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)              =
''===================================================================
''*******************************************************************
''**                                                               **
''** OBJETIVO : Montar ComboBox                                    **
''**                                                               **
''** Recebe: Combo    - Nome do ComboBox.                          **
''**         SQL$     - Query que extrairá os dados para o Combo   **
''**         DSC$     - Decrição do dado do Combo                  **
''**         ID       - Código do dado do Combo (Opcional)         **
''**         IDVIEW   - Determina se o ID será visível ou não      **
''**         ComboAux - Se IDVIEW=True deve ser passado o Combo    **
''**                    Auxiliar que conterá os valores de ID      **
''**                                                               **
''** Retorna : Combo Montado                                       **
''**                                                               **
''*******************************************************************
'   Dim mPointer%, DyCb As Recordset
'   mPointer% = Screen.MousePointer
'   Screen.MousePointer = vbHourglass
'   '* Limpa Combo
'   If Limpa Then Combo.Clear
'   If Trim(Sql) = "" Then
'      Sql = Combo.Tag
'   Else
'      Combo.Tag = Sql$
'   End If
'   Call Dbase.AbreTabela(Sql$, DyCb)
'   If Dbase.CodeSql = True Then
'      'Carrega o Combo
'      DyCb.MoveFirst
'      Do While Not DyCb.EOF
'         If Id = "" Then
'            Combo.AddItem DyCb(Dsc$)
'         Else
'            If IdView Then
'               Combo.AddItem DyCb(Id) & " - " & DyCb(Dsc$)
'            Else
'               Combo.AddItem DyCb(Dsc$)
'               If IsMissing(ComboAux) Then
'                  Combo.ItemData(Combo.ListCount - 1) = CInt(DyCb(Id))
'               Else
'                  ComboAux.AddItem CStr(DyCb(Id))
'               End If
'            End If
'         End If
'         DyCb.MoveNext
'      Loop
'      Combo.Tag = Combo.Tag & "LOAD"
'      If Combo.ListCount > 0 Then Combo.ListIndex = 0 ' Combo.List(0)
'      Combo.Tag = Mid(Combo.Tag, 1, Len(Combo.Tag) - 4)
'      DyCb.Close
'   End If
'   Combo.Enabled = True
'   Combo.Visible = True
'   Screen.MousePointer = mPointer%
'End Sub
'Public Function clsdsr.StrZero(Valor As Variant, Num As Integer, Optional Caracter = "0") As String
'   Dim i%, Zeros$
'   Zeros = String(Num%, Caracter)
'   clsdsr.StrZero = Right(Zeros + Trim(str(Val(Valor))), Num%)
'End Function
'Public Function LocalizarCombo(Cmb As Object, Chave As String, Optional SetCombo = True) As Integer
'   LocalizarCombo = SendMessageAny(Cmb.hWnd, CB_FINDSTRING, -1, ByVal Chave$)
'   If SetCombo Then
'      If Cmb.ListCount <> 0 Then
'         If Cmb.Style = 2 Then
'            Cmb.ListIndex = LocalizarCombo
'         Else
'            Cmb = Cmb.List(LocalizarCombo)
'         End If
'      End If
'   End If
'End Function
'Public Sub clsdsr.LimparTela(Frm As Object)
'   Dim i%
'   On Error Resume Next
'   For i% = 0 To Frm.Controls.Count - 1
'      Select Case UCase(TypeName(Frm.Controls(i)))
'         Case "TEXTBOX": Frm.Controls(i) = ""
'         Case "MASKEDBOX": Call ZerarMask(Frm.Controls(i))
'         Case "LABEL": If Frm.Controls(i).BorderStyle = 1 Then Frm.Controls(i) = ""
'         Case "OPTIONBUTTON": If Frm.Controls(i).index = 0 Then Frm.Controls(i).Value = True
'         Case "COMBOBOX": Frm.Controls(i) = Frm.Controls(i).List(0)
'         Case "CHECKBOX": Frm.Controls(i).Value = 0
'      End Select
'   Next
'End Sub
'Public Sub ZerarMask(ByRef Msk As Object, Optional ByRef MAscara)
'   If IsMissing(MAscara) Then MAscara = Msk.Mask
'   Msk.Mask = ""
'   Msk.Text = ""
'   Msk.Mask = MAscara
'End Sub
'Public Sub ExibirStop(Txt$, Optional Tit$)
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Exibir um aviso na tela.                        **
''**                                                            **
''** Recebe: Mensagem$ - Aviso a ser exibido                    **
''**         Tit$   - Título da Mensagem (Opcional)          **
''**                                                            **
''** Retorna : Aviso centralizado na tela com um botão de OK e  **
''**           um ícone de Stop.                                **
''**                                                            **
''****************************************************************
'    Dim Mouse%
'    Mouse = Screen.MousePointer
'    If TypeName(Tit$) = "Nothing" Then
'       Tit$ = clsmsg.LoadMsg(1)
'    End If
'    Screen.MousePointer = vbDefault
'    MsgBox Txt$, vbCritical, Tit$
'    DoEvents
'    Screen.MousePointer = Mouse
'End Sub
'Public Function SqlStr(Txt As String) As String
'   SqlStr = "'" & Txt & "'"
'End Function
'Public Sub ExibirInformacao(Txt$, Optional Tit$)
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Exibir uma informação.                          **
''**                                                            **
''** Recebe: Mensagem$ - Informação ser exibida                 **
''**         Tit$   - Título da Mensagem (Opcional)          **
''**                                                            **
''** Retorna : Informacao Centralizada com um botão de OK e     **
''**           um ícone de Informação.                          **
''**                                                            **
''****************************************************************
'    Dim Mouse%
'    Mouse = Screen.MousePointer
'    If TypeName(Tit$) = "Nothing" Then
'       Tit$ = LoadMsg(1)
'    End If
'    Screen.MousePointer = vbDefault
'    MsgBox Txt$, vbInformation, Tit$
'    DoEvents
'    Screen.MousePointer = Mouse
'End Sub
'Public Function ExibirPergunta(Txt As String, Optional Tit As Variant) As Integer
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Exibir uma pergunta na tela.                    **
''**                                                            **
''** Recebe: Mensagem$ - Pergunta a ser exibida                 **
''**         Tit$   - Título da Mensagem (Opcional)          **
''**                                                            **
''** Retorna : Resposta da pergunta centralizada com botões Yes **
''**           e No e com um ícone de Interrogação.             **
''**                                                            **
''****************************************************************
'    Dim Mouse%
'    Mouse = Screen.MousePointer
'    If TypeName(Tit) = "Nothing" Then
'       Tit = LoadMsg(1)
'    End If
'    Screen.MousePointer = vbDefault
'    ExibirPergunta% = MsgBox(Trim(Txt$), vbQuestion + vbYesNo, Tit)
'    DoEvents
'    Screen.MousePointer = Mouse
'End Function
'Public Sub RefreshMSGrid(ByRef DataControl As Object, ByRef Grd As Object)
'   Dim i%, J%, Max%, Fixed%, Lin%, LinTop%
'   Dim PosCampo%, pos&
'   Dim Campo$, Letra$
'   Dim Cab As Variant
'   On Error GoTo fim
''* Salvar número de coluna fixas e a linha corrente
'   Fixed% = Grd.FixedCols
'   Grd.FixedCols = 0
'   Lin% = Grd.Row
'   LinTop% = Grd.TopRow
''* Montar Cabeçalho do Grid Existente
'   ReDim Cab((4 * Grd.Cols) - 1)
'   For i = 0 To (Grd.Cols - 1)
'      J = i + 1
'      Cab((J% * 4) - 4) = Grd.TextMatrix(0, i)
'      Cab((J% * 4) - 2) = Grd.ColWidth(i) / 120
'      Cab((J% * 4) - 1) = Grd.ColAlignment(i)
'   Next
''* Atualizar Dados Exibidos no Grid
'   DataControl.Refresh
'   For i = 0 To (Grd.Cols - 1)
'      J = i + 1
'      Cab((J% * 4) - 3) = Grd.TextMatrix(0, i)
'   Next
'   Max% = (UBound(Cab) - LBound(Cab) + 1) / 4
''* Esconder todas a colunas
'   Grd.Visible = False
'   Grd.Refresh
''* Mostrar e Formatar as colunas selecionadas
''   Grd.Width = 260
'   For J = 1 To Max%
'      For i = 0 To Grd.Cols - 1
'         If Grd.TextMatrix(0, i) = Cab((J% * 4) - 3) Then
'            Grd.TextMatrix(0, i) = Cab((J% * 4) - 4)      'Título
'            Grd.ColWidth(i) = Cab((J% * 4) - 2) * 120   'Tamanho
'            Grd.ColAlignment(i) = Cab((J% * 4) - 1)     'Alinhamento
''            If Grd.Width >= 9440 Then Grd.Width = 9440
''            Grd.Width = Grd.Width + Grd.ColWidth(i)
'            Grd.Parent.Width = IIf(Grd.Parent.Width > Grd.Width, Grd.Parent.Width, Grd.Width + 240)
'         End If
'      Next
'   Next
'   If Grd.Width >= 9295 Then
'      Grd.Left = 60
'      Grd.Width = 9440
'   End If
''* Restaurar Cores do Cabeçalho
'   Grd.Row = 0
'   Grd.Col = Val(Grd.Tag)
'   Grd.CellBackColor = vbBlue
'   Grd.CellForeColor = vbWhite
''* Salvar número de coluna fixas e a linha corrente
'   Grd.FixedCols = Fixed%
'   If Grd.Rows > LinTop% Then
'      Lin% = IIf(Lin% = 0, 1, Lin%)
'      Grd.Row = IIf(Lin% >= Grd.Rows, Grd.Rows - 1, Lin)
'      Grd.TopRow = LinTop%
'   End If
'   Call SelRowMSGrid(Grd, Lin%)
'   Grd.Visible = True
'   Grd.Refresh
'   Grd.SetFocus
'   Exit Sub
'fim:
'   If Err <> 5 And Err <> 438 Then ShowError
'End Sub
'Public Sub SelRowMSGrid(Grd As Object, Lin%)
'
''Área de manipulação de dados
'    If Grd.Rows <= Lin Then Exit Sub
'    Grd.Row = Lin
'    Grd.Col = IIf(Grd.FixedCols = 0, 0, Grd.FixedCols - 1)
'    Grd.RowSel = Lin
'    Grd.ColSel = Grd.Cols - 1
'End Sub
'Public Function MSGrdProcurar(ByRef DTC As Object, ByRef Grd As Object, ByRef Pnl As Object, Key) As String
''================================================================
''= Última Alteração : 10/07/98                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Criticar Texto de um "Panel" para executar um   **
''**            PesquisarDbGrid%                                **
''**                                                            **
''** Recebe:  DTC     - Data Object referenciado pelo DBGrid   **
''**          Grd     - Grd de Pesquisa                        **
''**          Pnl     - Panel que contem a texto de pesquisa    **
''**          Chave   - Chave de pesquisa                       **
''**                                                            **
''** Retorna: Linha do Grid selecionada.                        **
''**                                                            **
''**                                                            **
''****************************************************************
'   Dim Chave$, RowTop%, Coluna%
'   On Error GoTo fim
'   If Grd.Rows <= 1 Then Exit Function
'   MSGrdProcurar = LTrim(Pnl + IIf(Key > 0, UCase$(Chr$(Key)), ""))
'   If Len(MSGrdProcurar) > 0 Then
'      Chave$ = MSGrdProcurar
'      Coluna% = Val(Grd.Tag)
'      RowTop% = PesquisarMSGrid(Grd, Chave$, Coluna%)
'      Grd.TopRow = RowTop%
'      Call SelRowMSGrid(Grd, RowTop)
'   End If
'   Exit Function
'fim:
'    ShowError
'End Function
'Public Function PesquisarMSGrid%(Grd As Object, Chave$, Coluna%, Optional MatchCase = False)
''================================================================
''= Última Alteração : 13/01/98                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Pesquisar item no Grd                          **
''**                                                            **
''** Recebe:  Grd    - Grd de Pesquisa                        **
''**          chave$  - Chave de Pesquisa                       **
''**          coluna% - Coluna de Pesquisa                      **
''**                                                            **
''** Retorna: linha do Grd                                     **
''**                                                            **
''****************************************************************
'   Dim Min%             '* Primeira linha do Grd com dados
'   Dim Max%             '* Última linha do Grd com dados
'   Dim rowi%, Lin%      '* Linha inicial e atual do Grd
'   Dim Tam%             '* Tamanho da chave
'   Dim Vl$              '* Valor da celula a ser comparada
'   Dim isNum As Boolean '* Comparação numérica
'   Dim Ordem As Boolean '* Ordem da Coluna True = Crescente
'   Dim Ch_Vl As Boolean, Ch_Vl2 As Boolean
'
'   If Grd.Rows <= 1 Then Exit Function
'   rowi% = Grd.Row
'
'   Tam% = Len(Chave$)
'   Min% = 1
'   Max% = IIf(Grd.Rows < 1, 0, Grd.Rows - 1)
'   If isNum Then
'      Vl$ = CDbl(Grd.TextMatrix(Max, Coluna))
'   Else
'      Vl$ = String_Sem_Acento(Left$(Grd.TextMatrix(Max, Coluna), Tam%))
'   End If
'
'   If Not MatchCase Then
'      Vl$ = UCase(Vl$)
'      Chave$ = UCase(Chave)
'      Chave$ = String_Sem_Acento(Chave)
'   End If
'   isNum = (Between(Asc(Mid(Grd.TextMatrix(Max, Coluna), 1, 1)), vbKey0, vbKey9) And _
'             Between(Asc(Mid(Grd.TextMatrix(1, Coluna), 1, 1)), vbKey0, vbKey9))
'   If isNum Then
'      Ordem = (CDbl(Grd.TextMatrix(Max, Coluna)) > CDbl(Grd.TextMatrix(1, Coluna)))
'   Else
'      Ordem = (Grd.TextMatrix(Max, Coluna) > Grd.TextMatrix(1, Coluna))
'   End If
'   If isNum Then
'      Ch_Vl = IIf(Ordem, CDbl(Chave$) > CDbl(Vl$), CDbl(Chave$) < CDbl(Vl$))
'   Else
'      Ch_Vl = IIf(Ordem, Chave$ > Vl$, Chave$ < Vl$)
'   End If
'
'
'   If Ch_Vl Then
'      Lin% = Max%
'   Else
'      If isNum Then
'         Ch_Vl = IIf(Ordem, CDbl(Chave$) <= CDbl(Grd.TextMatrix(Min, Coluna)), CDbl(Chave$) >= CDbl(Grd.TextMatrix(Min, Coluna)))
'      Else
'         Ch_Vl = IIf(Ordem, Chave$ <= Left$(Grd.TextMatrix(Min, Coluna), Tam%), Chave$ >= Left$(Grd.TextMatrix(Min, Coluna), Tam%))
'      End If
'
'      If Ch_Vl Then
'         Lin% = Min%
'      Else
'         Do While Min% <= Max%
'            Lin% = (Max% + Min%) / 2
'            If isNum Then
'               Vl$ = Grd.TextMatrix(Lin, Coluna)
'               Ch_Vl = IIf(Ordem, CDbl(Chave$) > CDbl(Vl$), CDbl(Chave$) < CDbl(Vl$))
'               Ch_Vl2 = IIf(Ordem, CDbl(Chave$) < CDbl(Vl$), CDbl(Chave$) > CDbl(Vl$))
'            Else
'               Vl$ = String_Sem_Acento(Left$(Grd.TextMatrix(Lin, Coluna), Tam%))
'               If Not MatchCase Then Vl$ = UCase(Vl$)
'               Ch_Vl = IIf(Ordem, Chave$ > Vl$, Chave$ < Vl$)
'               Ch_Vl2 = IIf(Ordem, Chave$ < Vl$, Chave$ > Vl$)
'            End If
'            If Ch_Vl Then
'               Min% = Lin% + 1
'            ElseIf Ch_Vl2 Then
'               Max% = Lin% - 1
'            Else
'               Do While Chave$ = Vl$
'                  Lin% = Lin% - 1
'                  Vl$ = String_Sem_Acento(Left$(Grd.TextMatrix(Lin, Coluna), Tam%))
'                  If Not MatchCase Then Vl$ = UCase(Vl$)
'               Loop
'               Lin% = Lin% + 1
'               Exit Do
'            End If
'         Loop
'      End If
'   End If
'   If Min% > Max% Then
'      Lin% = Max% 'rowi%
'   End If
'
'   Dim MAX_LINHAS_GRID%
'   MAX_LINHAS_GRID% = 0
'
'   If Lin% >= Grd.TopRow + MAX_LINHAS_GRID% - 2 Then
'      If Lin% >= Grd.Rows - MAX_LINHAS_GRID% + 2 Then
'         rowi% = Grd.Rows - MAX_LINHAS_GRID% - 1
'         rowi% = IIf(rowi% < 0, 1, rowi%)
'         Grd.TopRow = rowi%
'      Else
'         Grd.TopRow = Lin%
'      End If
'   ElseIf Lin% < Grd.TopRow Then
'      Grd.TopRow = IIf(Lin% < 1, 1, Lin%)
'   End If
'
'   PesquisarMSGrid% = IIf(Lin% < 1, 0, Lin%)
'End Function
'Public Function MontarMSGrid(ByRef DataControl As Object, ByRef Grd As Object, Arr As Variant, Sql As String, Optional ByVal GrdTam = 0) As Boolean
'   Dim i%, J%, Max%, Cols%, Tam%, ColVal%, Fixed%
'   Dim PosCampo%, pos&
'   Dim Sql_Ori$, Sql_Rest$
'   Dim Campo$, Letra$, Aux$
'   Dim wFactor As Single, hFactor As Single
'   Dim lEOF As Boolean
'   wFactor = (Screen.Width / Screen.TwipsPerPixelX) / 800
'   hFactor = (Screen.Height / Screen.TwipsPerPixelY) / 600
'   On Error GoTo fim
'   Sql = UCase(Sql)
'   If Trim(Sql) = "" Then Exit Function
''* Salvar número de coluna fixas
'   Fixed% = Grd.FixedCols
'   If Grd.Cols = 0 Then Grd.Cols = 1
'   Grd.FixedCols = 0
''* Tratar Query para manter o índice
''* se o Grid já estiver ordenado.
'   If Grd.Tag <> "" And Grd.Rows > 1 Then
'      Sql_Ori$ = Sql
'      pos = InStr(Sql, "ORDER BY")
''      Grd.Visible = True
'      If pos <> 0 Then
'         pos = pos + 8
'         Campo$ = ""
'         Letra = Mid(Sql, pos, 1)
'         While Letra = " "
'            pos = pos + 1
'            Letra = Mid(Sql, pos, 1)
'         Wend
'         Do While (Letra <> " " And Letra <> ";")
'            Campo$ = Campo + Letra
'            pos = pos + 1
'            Letra = Mid(Sql, pos, 1)
'            If pos > Len(Sql) Then Exit Do
'         Loop
'      '* Define Coluna do Índice
'         If Trim(str(Val(Campo))) <> Campo And Campo <> "0" Then
'            i = 0
'            Sql = Mid(Sql, 1, InStr(Sql, Campo))
'            While InStr(Sql, ",") <> 0
'               i = i + 1
'               Sql = Mid(Sql, InStr(Sql, ",") + 1)
'            Wend
'            Campo = str(i + 1)
'         End If
'         If Grd.Tag <> 0 Then
'            Sql = Mid(Sql_Ori, 1, InStr(Sql_Ori, " ORDER BY")) + " ORDER BY " + CStr(IIf(Grd.Tag < 1, 1, Grd.Tag + 1))
'         Else
'            Sql = Sql_Ori
'         End If
'      End If
'   End If
''* Gerar Dados Exibidos no Grid
'   If TypeName(DataControl) = "MSRDC" Then
'      DataControl.Sql = Sql
'   Else
'      DataControl.RecordSource = Sql
'   End If
'   DataControl.Refresh
'
''* Esconder todas a colunas
'   Grd.Visible = False
'
'   If IsEmpty(Arr) Then
'      Max = 0
'   Else
'      Max% = (UBound(Arr) - LBound(Arr) + 1) / 4
'   '* Mostrar e Formatar as colunas selecionadas
'      If GrdTam = 0 Then
'         Grd.Width = (260 + 80) * wFactor
'      Else
'         Grd.Width = (GrdTam + 80) * wFactor
'      End If
'   End If
'   If Grd.Rows = 0 Then Grd.Rows = 1
'   Grd.Row = 0
'   If Max > 0 Then '* Número de Colunas
'      For i = 0 To Grd.Cols - 1
'         Grd.ColWidth(i) = 0
'      Next
'      For J = 1 To Max%
'          For i = 0 To Grd.Cols - 1
'             Grd.Col = i
'             If Grd.Text = Arr((J% * 4) - 3) Then
'                Grd.Text = Arr((J% * 4) - 4)       'Título
'                Grd.ColWidth(i) = Arr((J% * 4) - 2) * 120 * wFactor 'Tamanho
'                Grd.ColAlignment(i) = Arr((J% * 4) - 1)     'Alinhamento
'                If GrdTam = 0 Then Grd.Width = Grd.Width + Grd.ColWidth(i)
'                Grd.Parent.Width = IIf(Grd.Parent.Width > Grd.Width, Grd.Parent.Width, Grd.Width + 240)
'                Exit For
'             End If
'             If i > Max% And J > Max% Then
'                Grd.ColWidth(i) = 0
'             End If
'          Next
'      Next
'   Else
'      If TypeName(DataControl) = "MSRDC" Then
'         Cols = DataControl.Resultset.rdoColumns.Count - 1
'         For i = 0 To Cols
'            Tam = DataControl.Resultset.rdoColumns(i).size
'            Tam = IIf(DataControl.Resultset.rdoColumns(i).size < Len(DataControl.Resultset.rdoColumns(i).Name), Len(DataControl.Resultset.rdoColumns(i).Name), Tam)
'            Tam = IIf(Tam > 30, 30, Tam)
'            Grd.ColWidth(i) = Tam * 120 * wFactor
'         Next
'      Else
'         'DataControl.RecordSource = Sql
'      End If
'   End If
'   If Grd.Width >= Grd.Parent.Width * wFactor Then
'      Grd.Left = 60 * wFactor
'      Grd.Width = Grd.Parent.Width * wFactor
'   End If
'   If TypeName(DataControl) = "MSRDC" Then
'      lEOF = DataControl.Resultset.EOF
'   Else
'      lEOF = DataControl.Recordset.EOF
'   End If
'   If lEOF Then
'      If Grd.Cols > Fixed Then Grd.FixedCols = Fixed%
'      Grd.Col = 0
'      Grd.Visible = True
'      Grd.Refresh
'      Exit Function
'   End If
'   pos = InStr(Sql, "ORDER BY")
'   Grd.Visible = True
'   If pos <> 0 Then
'      Aux = Mid(Sql, pos)
'      If InStr(Aux, ",") = 0 Then '* Só Existe um campo no ORDER BY
'         pos = pos + 8
'         Campo$ = ""
'         Letra = Mid(Sql, pos, 1)
'         While Letra = " "
'            pos = pos + 1
'            Letra = Mid(Sql, pos, 1)
'         Wend
'         Do While (Letra <> " " And Letra <> ";")
'            Campo$ = Campo + Letra
'            pos = pos + 1
'            Letra = Mid(Sql, pos, 1)
'            If pos > Len(Sql) Then Exit Do
'         Loop
'      '* Define Coluna do Índice
'         If Trim(str(Val(Campo))) <> Campo And Campo <> "0" Then
'            i = 0
'            Sql = Mid(Sql, 1, InStr(Sql, Campo) - 1)
'            While InStr(Sql, ",") <> 0
'               i = i + 1
'               Sql = Mid(Sql, InStr(Sql, ",") + 1)
'            Wend
'            Campo = str(i + 1)
'         End If
'         Grd.Tag = Val(Campo) - 1
'         If Grd.Tag = CStr(-1) Then
'            Sql = Mid(Sql, 1, InStr(Sql, Campo))
'            Grd.Tag = 0
'            While InStr(Sql, ",") <> 0
'               Grd.Tag = Grd.Tag + 1
'               Sql = Mid(Sql, 1, InStr(Sql, ",") - 1)
'            Wend
'         End If
'         Grd.Col = IIf(Val(Grd.Tag) >= Grd.Cols, Grd.Cols - 1, Val(Grd.Tag))
''         Grd.CellBackColor = vbBlue
''         Grd.CellForeColor = vbWhite
''         Grd.Col = 0
'      End If
'   End If
'   If Grd.Rows > 1 Then Grd.Row = 1
''   For i = 0 To Fixed% - 1
''      If Grd.ColWidth(i) = 0 Then Fixed% = Fixed% + 1
''   Next
'   While Fixed >= Grd.Cols
'      Fixed = Fixed% - 1
'   Wend
'   '***************
'   '* Configurar Picture de Ordenação no CabeçaLho do Grid
'   '***************
'   Call SetPctOrder(Grd, Sql)
'   Grd.Col = 0
'   On Error GoTo fim
'   Grd.FixedCols = Fixed%
'   Grd.Visible = True
'   Grd.Refresh
'   MontarMSGrid = True
'   Exit Function
'fim:
'   MontarMSGrid = False
'    ShowError
'End Function
'Private Sub SetPctOrder(Grd As Object, Sql$)
'   Dim n, Desc As Boolean
'   Dim pos%, i%
'   On Error Resume Next
'   Sql = UCase(Sql)
'   pos = InStr(Sql, "ORDER BY")
'   Sql = Mid(Sql, pos + 9)
'   Desc = (InStr(UCase(Sql), " DESC ") <> 0)
'   With Grd
'      i = 0
'      For Each n In .Parent.Controls
'         If UCase(n.Name) = "PCTORDER" Then
'            n.Visible = True
'            n.Move .Left + .CellLeft + .CellWidth - 240, .Top + 60
'            Call SetPicture(.Parent.Controls(i), IIf(Desc, "ORDER_UP", "ORDER_DOWN"))
'            Exit For
'         End If
'         i = i + 1
'      Next
'   End With
'   Err = 0
'End Sub
'Public Sub CentrarObj(ObjMain As Object, ObjChild As Object, Optional Tip)
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Centralizar um Form em outro form base.         **
''**                                                            **
''** Recebe:  MDI_name  - MDIForm base                          **
''**          form_name - Form a ser centralizado               **
''**                                                            **
''** Retorna: Form centralizado                                 **
''**                                                            **
''****************************************************************
'   Dim HH!, LL!, tt!, WW!
'   On Error Resume Next
'   If IsMissing(Tip) Then Tip = ""
'
'   WW! = ObjChild.ScaleWidth
'   If Err = 438 Then WW! = ObjChild.Width: Err = 0
'   HH! = ObjChild.ScaleHeight
'   If Err = 438 Then HH! = ObjChild.Height: Err = 0
'   LL! = ObjChild.ScaleLeft
'   If Err = 438 Then LL! = ObjChild.Left: Err = 0
'   tt! = ObjChild.ScaleTop
'   If Err = 438 Then tt! = ObjChild.Top: Err = 0
'   If Tip = "V" Or Tip = "" Then
'      tt! = (ObjMain.ScaleHeight - HH!) / 2
'      If Err = 438 Then
'         HH! = ObjChild.Height
'         tt! = (ObjMain.Height - HH!) / 2
'         Err = 0
'      End If
'   End If
'   If Tip = "H" Or Tip = "" Then
'      LL! = (ObjMain.ScaleWidth - WW!) / 2
'      If Err = 438 Then
'         WW! = ObjChild.Width
'         LL! = (ObjMain.Width - WW!) / 2
'         Err = 0
'      End If
'   End If
'
'   ObjChild.Move LL!, tt!
'End Sub
'Public Sub OrdenarMSGrid(DTC As Object, Grd As Object, Ind%)
'   Dim pos%, Sql$, i%, Fixed%, Ord As Boolean
'   Dim MyIndex As index, MyField As Field
'   Dim Cab()
'   Dim Rc1 As Recordset, Rc2 As Recordset
'
'   On Error GoTo fim
'   Screen.MousePointer = vbHourglass
'   If Ind < 0 Then Exit Sub
'   If Grd.Cols <= Ind Then Exit Sub
'   Grd.Tag = Ind%
''* Salvar Estrutura do Grid
'   Fixed% = Grd.FixedCols
'   Grd.FixedCols = 0
'   ReDim Cab(Grd.Cols - 1, 2)
'   Grd.Row = 0
'   For i = 0 To Grd.Cols - 1
'      Grd.Col = i
'      Cab(i, 0) = Grd                 'Título
'      Cab(i, 1) = Grd.ColWidth(i)     'Tamanho
'      Cab(i, 2) = Grd.ColAlignment(i) 'Alinhamento
'   Next
'   If Ind < 0 Then Exit Sub
'   If True Then
'      If TypeName(DTC) = "MSRDC" Then
'         Sql = DTC.Sql
'      Else
'         Sql = DTC.RecordSource
'      End If
'      pos = InStr(UCase(Sql), "ORDER BY")
'      If pos <> 0 Then
'         Ord = (InStr(Mid(Sql, pos), str(Ind + 1)) <> 0)
'         Ord = (Ord And (InStr(Mid(Sql, pos), " DESC ") = 0))
'      Else
'         Ord = False
'         pos = Len(Sql) + 1
'      End If
'      If pos <> 0 Then
'         Sql$ = Trim(Mid(Sql, 1, pos - 1))
'         Sql$ = Sql$ & " order by " & Trim(str(Ind + 1))
'         Sql = Sql & IIf(Ord, " DESC ", "")
'         If TypeName(DTC) = "MSRDC" Then
'            DTC.Sql = Sql
'         Else
'            DTC.RecordSource = Sql$
'         End If
'         DTC.Refresh
'      End If
'   Else
'      Set Rc1 = DTC.Recordset
'      Rc1.Sort = DTC.Recordset.Fields(Ind).Name
'      Set Rc2 = Rc1.OpenRecordset(Rc1.Type)
'      Set DTC.Recordset = Rc2
'   End If
'   Set Rc1 = Nothing
'   Set Rc2 = Nothing
'
'   'Recuperar Estrutura do Grid
'   Grd.Row = 0
'   For i = 0 To Grd.Cols - 1
'      Grd.Col = i%
'      Grd = Cab(i, 0)                 'Título
'      Grd.ColWidth(i) = Cab(i, 1)     'Tamanho
'      Grd.ColAlignment(i) = Cab(i, 2) 'Alinhamento
'   Next
'   Grd.FixedCols = Fixed%
'   Grd.Col = Ind%
'   '***************
'   '* Configurar Picture de Ordenação no CabeçaLho do Grid
'   '***************
'   Call SetPctOrder(Grd, Sql)
''   Grd.CellBackColor = vbBlue
''   Grd.CellForeColor = vbWhite
'   On Error GoTo fim
'   Call SelRowMSGrid(Grd, 1)
'   Screen.MousePointer = vbDefault
'   Exit Sub
'fim:
'    ShowError
'   Screen.MousePointer = vbDefault
'End Sub
'Public Function LoadOriMsg$(Num%)
'  On Error GoTo fim
'  LoadOriMsg$ = LoadResString(Num)
'  Exit Function
'fim:
'   ShowError
'End Function
'Public Sub SetDefault(hWnd As Long)
'  DoEvents
'  Call ReleaseCapture
'  Call SetCursor(LoadCursor(0, IDC_ARROW))
'  Screen.MousePointer = vbDefault
'End Sub
'Public Sub SetHourglass(hWnd As Long)
'  DoEvents
'  Call SetCapture(hWnd)
'  Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
'  Screen.MousePointer = vbHourglass
'End Sub
'Public Sub MontarCombo(ByRef Combo As Object, Coll As Collection, Optional Propiedade = "") ' As Property)
'   Dim n As Variant, Dsc$
'   Dim mPointer%, i%
'   mPointer% = Screen.MousePointer
'   Screen.MousePointer = vbHourglass
'
'   Combo.Clear
'   i = 0
'   For Each n In Coll
'      Select Case UCase(Propiedade)
'         Case "IDNF": Dsc = n.IDNF
'         Case "DSCPIECE": Dsc = n.DSCPIECE
'         Case "IDPIECE": Dsc = n.IDPIECE
'         Case Else: Dsc = n
'      End Select
'      Combo.AddItem Dsc
''      If Combo.Sorted Then
''         Combo.ItemData(LocalizarCombo(Combo, Dsc)) = i
''      End If
'      i = i + 1
'   Next
'   On Error Resume Next
'   Combo.Enabled = True
'   Combo.Visible = True
'   Combo.ListIndex = 0
'   Screen.MousePointer = vbDefault
'   On Error GoTo 0
'End Sub
'Public Function StrReplace(ByVal TxtIn As String, ByVal TxtFrom As String, ByVal TxtTo As String) As String
''================================================================
''= Última Alteração : 20/01/99                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Substitui o texto de Txtfrom$ para o texto de   **
''**            TxtOut$ na string TxtIn$.                       **
''**                                                            **
''** Recebe: TxtIn$   - string a ser alterada                   **
''**         TxtFrom$ - texto a ser substituido                 **
''**         TxtOut$   - novo texto                             **
''**                                                            **
''** Retorna: string alterada                                   **
''**                                                            **
''****************************************************************
'   Dim TxtOut$, LenIn%, LenFrom%, pos%
'
'   LenIn = Len(TxtIn)
'   LenFrom = Len(TxtFrom)
'   If LenFrom < 1 Or LenIn < 1 Then
'      StrReplace = TxtIn
'      Exit Function
'   End If
'   TxtOut = ""
'   pos = InStr(TxtIn, TxtFrom)
'   While pos > 0
'      TxtOut = TxtOut + Left(TxtIn, pos - 1) + TxtTo
'      TxtIn = Right(TxtIn, Len(TxtIn) - pos - LenFrom + 1)
'      pos = InStr(TxtIn, TxtFrom)
'   Wend
'   TxtOut = TxtOut + TxtIn
'   StrReplace = TxtOut
'End Function
'Public Function String_Sem_Acento(str) As String
''ALTERAR FUNÇÃO PARA SUPORTAR CARACTERES MINUSCULOS
''só para maiúsculas
'
'Dim strNova, Car$
'Dim Tam%, i%
'
'Tam = Len(str)
'strNova = ""
'If Tam <> 0 Then
'    For i% = 1 To Tam
'        Car = Mid$(str, i, 1)
'        Select Case Asc(Car)
'            Case 192, 193, 194, 195, 196, 197: strNova = strNova + "A"
'            Case 200, 201, 202, 203: strNova = strNova + "E"
'            Case 204, 205, 206, 207: strNova = strNova + "I"
'            Case 210, 211, 212, 213, 214: strNova = strNova + "O"
'            Case 217, 218, 219, 229: strNova = strNova + "U"
'            Case 199: strNova = strNova + "C"
'            Case 209: strNova = strNova + "N"
'            Case 221: strNova = strNova + "Y"
'            Case Else: strNova = strNova + Car
'        End Select
'    Next i
'Else
'    strNova = str
'End If
'String_Sem_Acento = strNova
'
'End Function
'Public Function Between(Vl, Min, Max)
'   Between = False
'   If Vl >= Min And Vl <= Max Then
'      Between = True
'   End If
'End Function
'Public Sub SetPicture(Controle As Object, Key$, Optional Tipo = vbResBitmap)
'   Controle.Picture = LoadResPicture(Key, Tipo)
'End Sub
'Public Function SetFormatDT_Number() As Boolean
''================================================================
''= Última Alteração : 28/11/97                                  =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Definir formato de data e número                **
''**                                                            **
''** Recebe:                                                    **
''**                                                            **
''** Retorna: Mensagem com os formatos definidos se o sistema   **
''**          operacional estiver utilizando formatos diferntes **
''**                                                            **
''****************************************************************
'
''* Função deve ser revisada para uma utilização generalizada
'
'   Dim i%, Txt$, Aux$, aux1$, DT$, tmp_data$, tmp_ano%
'   Dim DIA$, Mes$, Ano$
'
'   Aux$ = Format$(1000, "#,##0.00")
'   mvarSepDec$ = Mid$(Aux$, 6, 1)
'   mvarSepMil$ = Mid$(Aux$, 2, 1)
'   mvarSepDt$ = "/"
'   Aux$ = CStr(Date)
'   For i% = 2 To 5
'      aux1$ = Mid$(Aux$, i%, 1)
'      'procura o primeiro caracter que não seja um dígito
'      If aux1 < "0" Or aux1 > "9" Then
'         mvarSepDt$ = aux1$
'         Exit For
'      End If
'   Next
'   DIA = Format$(Day(Aux$), "00")
'   Mes = Format$(Month(Aux$), "00")
'   Select Case Len(Aux$)
'      Case 8: Ano = Right$(Year(Now), 2)
'      Case 10: Ano = Year(Aux$)
'   End Select
'   Select Case Aux$
'      Case DIA + mvarSepDt$ + Mes + mvarSepDt$ + Ano
'         mvarFormatoData$ = "DMA"
'         mvarDtMask$ = "dd" + mvarSepDt$ + "mm" + mvarSepDt$ + "yyyy"
'         mvarDtMaskAux$ = "dd" + mvarSepDt$ + "mm" + mvarSepDt$ + "yy"
'      Case Mes + mvarSepDt$ + DIA + mvarSepDt$ + Ano
'         mvarFormatoData$ = "MDA"
'         mvarDtMask$ = "mm" + mvarSepDt$ + "dd" + mvarSepDt$ + "yyyy"
'         mvarDtMaskAux$ = "mm" + mvarSepDt$ + "dd" + mvarSepDt$ + "yy"
'      Case Ano + mvarSepDt$ + Mes + mvarSepDt$ + DIA
'         mvarFormatoData$ = "AMD"
'         mvarDtMask$ = "yyyy" + mvarSepDt$ + "mm" + mvarSepDt$ + "dd"
'         mvarDtMaskAux$ = "yy" + mvarSepDt$ + "mm" + mvarSepDt$ + "dd"
'   End Select
'   DT$ = Format$(Now, mvarDtMask$)
'
'   'testa formato data/número
'   If mvarDtMask$ <> "dd" + mvarSepDt$ + "mm" + mvarSepDt$ + "yyyy" Then
'      Txt$ = LoadMsg(17) + Chr$(10) + Chr$(10)
'      Txt$ = Txt$ + LoadMsg(19)
'      GoTo fim
'   End If
'   If Not (mvarSepDt$ = "/" And mvarSepMil$ = "." And mvarSepDec$ = ",") Then
'   ' And GetStringFromIni("Intl", "sShortDate", "win.ini") = UCase(mvarDtMask$)) Then
'   'If Not (mvarSepDt$ = "/" And mvarSepMil$ = "," And mvarSepDec$ = ".") Then
'      If Not (mvarSepDt$ = "." And mvarSepMil$ = "," And mvarSepDec$ = ".") Then
'      Txt$ = ClsMsg.LoadMsg(18) + Chr$(10) + Chr$(10)
'      Txt$ = Txt$ + ClsMsg.LoadMsg(19) + Chr$(10)
'      Txt$ = Txt$ + ClsMsg.LoadMsg(20) + Chr$(10)
''      Txt$ = Txt$ + "Número, utilize 9,999.99" + Chr$(10)
'       GoTo fim
'       End If
'   End If
'   Call WritePrivateProfileString("INTL", "SSHORTDATE", mvarDtMask$, "WIN.INI")
'
'
'    'testa formato ano da data
'    tmp_ano% = 0
'    tmp_data$ = Trim$(CStr(CVDate(DT$)))
'    For i% = Len(tmp_data$) To 1 Step -1
'       If Mid$(tmp_data$, i%, 1) = mvarSepDt$ Then
'          tmp_ano% = Val(Trim$(Mid$(tmp_data, i% + 1)))
'          Exit For
'       End If
'    Next i%
'    SetFormatDT_Number = True
'Exit Function
'
'fim:
'   Screen.MousePointer = vbDefault
'   MsgBox Txt$, vbCritical
'   DoEvents
'   SetFormatDT_Number = False
'End Function
'
'
