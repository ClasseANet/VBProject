VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mensagem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub ExibirStop(pTexto As String, Optional pTITULO As String)
    Dim Mouse As Integer
    
    Mouse = Screen.MousePointer
    If Trim(pTITULO) = "" Then
       pTITULO = Traduzir("Atenção!")
    End If
    Screen.MousePointer = vbDefault
    MsgBox pTexto, vbCritical, pTITULO
    DoEvents
    Screen.MousePointer = Mouse
End Sub
Public Function ExibirPergunta(pTexto As String, Optional pTITULO As String, Optional pDefaultYes = True) As Integer
    Dim Mouse As Integer
    
    Mouse = Screen.MousePointer
    If Trim(pTITULO) = "" Then
       pTITULO = Traduzir("Atenção!")
    End If
    Screen.MousePointer = vbDefault
    If pDefaultYes Then
       ExibirPergunta% = MsgBox(Trim(pTexto), vbQuestion + vbYesNo, pTITULO)
    Else
       ExibirPergunta% = MsgBox(Trim(pTexto), vbQuestion + vbYesNo + vbDefaultButton2, pTITULO)
    End If
'    DoEvents
    Screen.MousePointer = Mouse
End Function
Public Sub ExibirInformacao(pTexto As String, Optional pTITULO As String)
    Dim Mouse As Integer
    
    Mouse = Screen.MousePointer
    If Trim(pTITULO) = "" Then
       pTITULO = Traduzir("Atenção!")
    End If
    Screen.MousePointer = vbDefault
    MsgBox pTexto, vbInformation, pTITULO
    DoEvents
    Screen.MousePointer = Mouse
End Sub
Public Sub ExibirAviso(pTexto As String, Optional pTITULO As String)
   Dim Mouse As Integer
   
   Mouse = Screen.MousePointer
    If Trim(pTITULO) = "" Then
       pTITULO = Traduzir("Atenção!")
    End If
   Screen.MousePointer = vbDefault
   MsgBox pTexto, vbExclamation, pTITULO
   DoEvents
   Screen.MousePointer = vbDefault
End Sub
Public Sub ExibirErro()
   Dim Mouse   As Integer
   Dim sTexto  As String
   Dim pTITULO As String
   
   Mouse = Screen.MousePointer
   
   sTexto = "O seguinte erro ocorreu : " & vbNewLine & vbNewLine
   sTexto = sTexto & "Number : " & Err.Number & vbNewLine
   sTexto = sTexto & "Description : " & Err.Description & vbNewLine
   sTexto = sTexto & "Help File : " & Err.HelpFile
   
   If Trim(pTITULO) = "" Then
      pTITULO = Traduzir("Atenção!")
   End If
   
   Screen.MousePointer = vbDefault
   MsgBox sTexto, vbError, pTITULO
   DoEvents
   Screen.MousePointer = vbDefault
End Sub
Public Sub ExibirResultado(pSys As Object, Optional pResultado As Boolean = True, Optional pNumPisca As Integer, Optional pMsg As String = "")
   Dim sMsg As String
    
   'pNumPisca = 1
   'pResultado = False
    
   If pMsg = "" Then
      If pResultado Then
         pMsg = "Operação realizada com Sucesso!!"
         If pNumPisca = 0 Then pNumPisca = 1
      Else
         pMsg = "Falha na operação!!"
         If pNumPisca = 0 Then pNumPisca = 2
      End If
   End If

   With pSys.MDI
      .Timer.Interval = 500
      .CommandBars.Tag = ClsGeneral.SetTag(.CommandBars.Tag, "EXIBIRRESULTADO", 1)
      .CommandBars.Tag = ClsGeneral.SetTag(.CommandBars.Tag, "NUMPISCA", pNumPisca)
      .CommandBars.Tag = ClsGeneral.SetTag(.CommandBars.Tag, "MSGPOSITIVA", IIf(pResultado, 1, 0))
      .CommandBars.Tag = ClsGeneral.SetTag(.CommandBars.Tag, "MSGRESULTADO", pMsg)
   End With
End Sub
Public Sub GravaError(TxtAux$, TxtTela$, Optional NumErro = "", Optional DscErro = "")
   Dim ArqLog$, ArqTmp$
   Dim Pos%, Limpar As Boolean
   Dim Txt$, i%
   Dim mvarDrvErro  As String
   
   On Error Resume Next
   
   ArqLog$ = "Error.log"
   ArqTmp$ = "Error.tmp"
   mvarDrvErro = "C:\TMP\"  'ClsDos.DrvTmp
   If Dir(mvarDrvErro, vbDirectory) = "" Then
      Call MkDir(mvarDrvErro)
   End If
   If Not ClsAmbiente.ExisteArquivo(mvarDrvErro & ArqLog$) Then
      Open mvarDrvErro & ArqLog$ For Output As #1
         Print #1, "0 Erros"
      Close #1
   End If
   Limpar = (FileLen(mvarDrvErro & ArqLog$) > 1000000) ' > 1 Mb
   Call ClsAmbiente.CopiarArquivo(mvarDrvErro & ArqLog$, mvarDrvErro & ArqTmp$)
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
      Print #1, "Computer : " & ClsAmbiente.ComputerName
      Print #1, "Source   : " & Error 'Errors(0).Source
'      If NumErro = Errors(0).Number Then
'         Print #1, "Erro     : " & Errors(0).Number & " - " & Mid(Errors(0).Description, 1, 50)
'      Else
         Print #1, "Erro     : " & NumErro & " - " & Mid(DscErro, 1, 50)
'      End If
'      Txt = Errors(0).Description
      While Len(Txt) >= 50
         Txt = Mid(Txt, 51)
         Print #1, Space(18) & Mid(Txt, 1, 50)
      Wend
'      Print #1, "Help     : " & Errors(0).HelpContext & " - " & Errors(0).HelpFile
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
   Call ClsAmbiente.ExcluirArquivo(mvarDrvErro & ArqTmp$)
End Sub
Public Function Traduzir(pString As String, Optional pIdioma As Double) As String
   Traduzir = pString
End Function

