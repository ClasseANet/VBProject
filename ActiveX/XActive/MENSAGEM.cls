VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MENSAGEM"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Enum eIdioma
   PORTUGUES = 5000
   INGLES = 6000
   FRANCES = 7000
   ESPANHOL = 8000
End Enum
Private mvarIdioma As eIdioma
Private mvarDrvErro As String
Public Property Get DrvErro() As String
    DrvErro = mvarDrvErro
End Property
Public Property Let DrvErro(vData As String)
    mvarDrvErro = vData
End Property
Public Property Get Idioma() As eIdioma
    Idioma = mvarIdioma
End Property
Public Property Let Idioma(vData As eIdioma)
    mvarIdioma = vData
   If ClsMsg.Idioma <> vData Then
    ClsMsg.Idioma = vData
   End If
End Property
Public Function NL(Optional Num = 0)
   Dim i%
   For i = i To Num
      NL = NL + vbLf  '=vblf
   Next
End Function

Private Sub Class_Terminate()
   On Error Resume Next
   mvarIdioma = 0
End Sub
Public Function LoadMsg(Num As Integer) As String
   Dim Idioma As eIdioma
   Dim msg As String
   On Error GoTo Fim
   '* Idioma usado
   Idioma = mvarIdioma 'Português = 5000, Inglês = 6000, ...
   'Num% = Idioma + Num%
   'Select Case Idioma
   '   Case 5000: Msg = LoadResString(Num) 'LoadPortugues$(Num%)
   '   Case 6000: Msg = LoadResString(Num) 'LoadIngles$(Num%)
   'End Select
   'LoadMsg = Msg
   If Num > 9999 Then
      Num = 5000 + (Num - (Val(Mid(CStr(Num), 1, Len(CStr(Num)) - 3)) * 1000))
   End If
   LoadMsg = LoadResString(mvarIdioma + Num)
   Exit Function
Fim:
   ShowError
End Function
Public Function LoadPict(Num, Tipo As Integer)
   On Error GoTo Fim
   LoadPict = LoadResPicture(Num, Tipo)
   Exit Function
Fim:
   ShowError
End Function
Public Function LoadOriMsg$(Num%)
  On Error GoTo Fim
  LoadOriMsg$ = LoadResString(Num)
  Exit Function
Fim:
   ShowError
End Function

'------------------------------------------------------------
'this sub displays the error message with it's Err code
'and prompts to show the Errors collection if it
'is a data access type error
'------------------------------------------------------------
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

   Call GravaError(CStr(TxtAux), TxtTela$, NumErrors, DscErrors)
   On Error GoTo 0
   Errors.Refresh
End Sub
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
       Tit$ = ClsMsg.LoadMsg(1)
    End If
    Screen.MousePointer = vbDefault
    MsgBox Txt$, vbCritical, Tit$
    DoEvents
    Screen.MousePointer = Mouse
End Sub
Public Function ExibirPergunta(Txt As String, Optional Tit As Variant, Optional DefaultYes = True) As Integer
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
    If TypeName(Tit) = "Nothing" Then
       Tit = ClsMsg.LoadMsg(1)
    End If
    Screen.MousePointer = vbDefault
    If DefaultYes Then
       ExibirPergunta% = MsgBox(Trim(Txt$), vbQuestion + vbYesNo, Tit)
    Else
       ExibirPergunta% = MsgBox(Trim(Txt$), vbQuestion + vbYesNo + vbDefaultButton2, Tit)
    End If
'    DoEvents
    Screen.MousePointer = Mouse
End Function
Public Sub ExibirInformacao(Txt$, Optional Tit$)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Exibir uma informação.                          **
'**                                                            **
'** Recebe: Mensagem$ - Informação ser exibida                 **
'**         Tit$   - Título da Mensagem (Opcional)          **
'**                                                            **
'** Retorna : Informacao Centralizada com um botão de OK e     **
'**           um ícone de Informação.                          **
'**                                                            **
'****************************************************************
    Dim Mouse%
    Mouse = Screen.MousePointer
    If TypeName(Tit$) = "Nothing" Then
       Tit$ = ClsMsg.LoadMsg(1)
    End If
    Screen.MousePointer = vbDefault
    MsgBox Txt$, vbInformation, Tit$
    DoEvents
    Screen.MousePointer = Mouse
End Sub
Public Sub ExibirAviso(Txt As String, Optional Tit)
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
   If IsMissing(Tit) Then
      Tit = ClsMsg.LoadMsg(1)
   End If
   Screen.MousePointer = vbDefault
   MsgBox Txt$, vbExclamation, Tit
   DoEvents
   Screen.MousePointer = vbDefault
End Sub
Public Sub GravaError(TxtAux$, TxtTela$, Optional NumErro = "", Optional DscErro = "")
   Dim ArqLog$, ArqTmp$
   Dim Pos%, Limpar As Boolean
   Dim Txt$, i%
   
   On Error Resume Next
   
   ArqLog$ = "Error.log"
   ArqTmp$ = "Error.tmp"
   mvarDrvErro = ClsDos.DrvTmp
   If Dir(mvarDrvErro, vbDirectory) = "" Then
      Call MkDir(mvarDrvErro)
   End If
   If Not ClsDos.FileExists(mvarDrvErro & ArqLog$) Then
      Open mvarDrvErro & ArqLog$ For Output As #1
         Print #1, "0 Erros"
      Close #1
   End If
   Limpar = (FileLen(mvarDrvErro & ArqLog$) > 1000000) ' > 1 Mb
   Call ClsDos.Copy(mvarDrvErro & ArqLog$, mvarDrvErro & ArqTmp$)
   Open mvarDrvErro & ArqLog$ For Output As #1
   Open mvarDrvErro & ArqTmp$ For Input As #2
      Line Input #2, Txt
      If Limpar Then
         Print #1, "4 Erros"
      Else
         Pos = 0
         While Pos = 0
            Pos = InStr(Txt, "Erros")
            If Pos > 2 Then
               If Pos > 5 Then
                  Print #1, CStr(CLng(Mid(Txt, Pos - 5, 5)) + 1) & " Erros"
               Else
                  Print #1, CStr(CLng(Mid(Txt, 1, Pos - 2)) + 1) & " Erros"
               End If
            End If
            If Not EOF(2) And Pos = 0 Then Line Input #2, Txt
            Pos = IIf(EOF(2), 1, Pos)
         Wend
      End If
      Print #1, "====================================================================="
      Print #1, "Date     : " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
      Print #1, "Computer : " & ClsDsr.ComputerName
      Print #1, "Source   : " & Errors(0).Source
      If NumErro = Errors(0).Number Then
         Print #1, "Erro     : " & Errors(0).Number & " - " & Mid(Errors(0).Description, 1, 50)
      Else
         Print #1, "Erro     : " & NumErro & " - " & Mid(DscErro, 1, 50)
      End If
      Txt = Errors(0).Description
      While Len(Txt) >= 50
         Txt = Mid(Txt, 51)
         Print #1, Space(18) & Mid(Txt, 1, 50)
      Wend
      Print #1, "Help     : " & Errors(0).HelpContext & " - " & Errors(0).HelpFile
      If TxtTela <> "" Then
         Print #1, TxtTela
      End If
      If TxtAux <> "" Then
         Print #1, "Auxiliar : " & TxtAux
      End If
      Print #1, "====================================================================="
      Do While Not EOF(2)
         Line Input #2, Txt
         Print #1, Txt
         If Limpar And Mid(Txt, 1, 5) = "=====" Then '* Se Arquivo > 1Mb
            i = i + 1
            If i = 6 Then Exit Do '* Gravar apenas os 3 últimos Erros
         End If
      Loop
   Close #2
'   Print #1, "====================================================================="
   Close #1
   Call ClsDos.Del(mvarDrvErro & ArqTmp$)
End Sub
Private Sub Class_Initialize()
   mvarIdioma = GetSetting("DSR", "General Format", "Idioma", PORTUGUES) '5000-Portugues
End Sub
