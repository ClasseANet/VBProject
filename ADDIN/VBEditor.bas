Attribute VB_Name = "VBEditor"
Option Explicit
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Public Type PointAPI
    x As Long
    Y As Long
End Type

Public Const CFM_BACKCOLOR = &H4000000
'Public Const WM_USER = &H400
Public Const LF_FACESIZE = 32

'Public Const EM_GETSEL = &HB0
'Public Const EM_SETSEL = &HB1
'Public Const EM_GETLINECOUNT = &HBA
'Public Const EM_LINEINDEX = &HBB
'Public Const EM_LINELENGTH = &HC1
'Public Const EM_LINEFROMCHAR = &HC9
'Public Const EM_CHARFROMPOS& = &HD7
'Public Const EM_SETCHARFORMAT = (WM_USER + 68)

Public Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60
    
    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer            ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer           ' /* Amount to space between letters  */
    crBackColor As Long        ' /* Background color                 */
    lLCID As Long               ' /* Locale ID                        */
    dwReserved As Long         ' /* Reserved. Must be 0              */
    sStyle As Integer            ' /* Style handle                     */
    wKerning As Integer            ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte     ' /* Underline type                   */
    bAnimation As Byte         ' /* Animated text like maRchTxting ants */
    bRevAuthor As Byte         ' /* Revision author index            */
    bReserved1 As Byte
End Type

Global Const gwDesenvolvedor = 1
Global Const gwWebSite = 2
Global Const gweMail = 3
Global Const gwTelefone = 4
Global Const gwData_Hora = 5
Global Const gwNome_do_Projeto = 6
Global Const gwNome_do_Modulo = 7
Global Const gwNome_do_Arquivo = 8
Global Const gwNome_da_Funcao = 9
Global Const gwParametros = 10
Global Const gwComentario = 11
Global Const gwComentario_do_Usuario = 12
Global Const gwConfiguracao = 13
Global Const gwConfigurar = 14
Global Const gwGeral = 15
Global Const gwIndentacao = 16
Global Const gwEspacos = 17
Global Const gwNome_do_Modelo = 18
Global Const gwSalvar = 19
Global Const gwExcluir = 20
Global Const gwCaracter = 21
Global Const gwModelo = 22
Global Const gwExibir_Somente_Itens_Selecionados = 23
Global Const gwCodigo_do_Usuario = 24
Global Const gwErro = 25
Global Const gwTratamento_de_Erro = 26
Global Const gwIndentar_Com_Função = 27
Global Const gwIndentar_Comentário = 28
Global Const gwIndentar_Select_Case = 29
Global Const gwIdioma = 30
Global Const gwBanco_de_Dados = 31
Global Const gwPortugues = 32
Global Const gwIngles = 33
Global Const gwFrances = 34
Global Const gwEspanhol = 35
Global Const gwInserir_Linha_em_Branco_Antes_da_Funcao = 36
Global Const gwProg_Auxiliar = 37
Global Const gwAbrir = 38
Global Const gwArquivo = 39
Public Sub GetConfig()
   Dim MySet As Variant, i%
   On Error Resume Next
   With Sys
      With .Edit
         .ExeAuxiliar = GetSetting(.AppName, "Outros", "ExeAuxiliar", "")
      
         .Desenvolvedor = GetSetting(.AppName, "Desenvolvedor", "Nome", "Diogenes S. Ramos")
         .WebSite = GetSetting(.AppName, "Desenvolvedor", "WebSite", "")
         .eMail = GetSetting(.AppName, "Desenvolvedor", "eMail", "disantos@ig.com.br")
         .Telefone = GetSetting(.AppName, "Desenvolvedor", "Telefone", "")
         .Idioma = GetSetting(.AppName, "General", "Idioma", 5000)   '* Português
         
         .ErrorLabel = GetSetting(.AppName, "Error", "ErrorLabel", "Trata_Erro")
         .ErrorFunction = GetSetting(.AppName, "Error", "ErrorFunction", "MsgBox Cstr(Err) & " & """" & " - " & """" & " & Error")
         .SaidaLabel = GetSetting(.AppName, "Error", "SaidaLabel", "Saida")
         .SaidaFunction = GetSetting(.AppName, "Error", "SaidaFunction", "Exit Sub")
         .SpcIndent = GetSetting(.AppName, "Indentar", "SpcIndent", "3")
                  
         .UserComment = GetSetting(.AppName, "Template", "UserComment", "")
         .CharComment = GetSetting(.AppName, "Template", "CharComment", "'* ")
         .Template = GetSetting(.AppName, "Template", "Nome", "Padrão")
         MySet = GetAllSettings(.AppName, "Templates")
         If Not IsEmpty(MySet) Then
            For i = LBound(MySet, 1) To UBound(MySet, 1)
               .Templates.Add MySet(i, 0) & " " & MySet(i, 1), MySet(i, 0)
            Next
         End If
         If .Templates.Count = 0 Then
            .Templates.Add "Padrão 1|5|6|7|8|9|10|11|", "Padrão"
         End If
         .IndentFunction = GetSetting(.AppName, "Indent", "IndentFunction", True)
         .IndentComment = GetSetting(.AppName, "Indent", "IndentComment", True)
         .IndentSelect = GetSetting(.AppName, "Indent", "IndentSelect", True)
         .LineBlankBefore = GetSetting(.AppName, "Indent", "LineBlankBefore", False)
      End With
      
      With .Constru
         .LoadIni = GetSetting(.AppName, "Setup", "LoadIni", False)
         .ExibeSubPasta = GetSetting(.AppName, "Setup", "ExibeSubPasta", False)
         .SalvarOnLine = GetSetting(.AppName, "Setup", "SalvarOnLine", False)
      End With
      
      With .Proj
         '*** [ General ] ***
         .MICRO = GetSetting(.AppName, "General", "MICRO", "RUN TIME")
         .Idioma = GetSetting(.AppName, "General", "IDIOMA", 5000)   '* Português
'''''''         'DSACTIVE.Idioma = .Idioma
         '   BANCO.Sys_Idioma = Sys.Idioma
         .FundoTela = GetSetting(.AppName, .AppName, "FUNDOTELA", "FUNDO")
         .DrvErro = GetSetting(.AppName, .AppName, "DRVERRO", App.PATH & "\Erro\")
         If Trim(.DrvErro) = "" Then
            .DrvErro = App.PATH & "\Erro\"
         End If
      '      sys.DrvDrive = SysMdi.Drv1.List(0) + "\"
'         .dbDrive_Orig = .dbDrive
      End With
   End With
End Sub
Public Function LinhaCab(Index As Integer, Optional NmFuncao = "") As String
   Dim NmProjeto$, NmArq$, NmModulo$
   Dim StartLine&, StarColumn&, EndLine&, EndColumn&
   Dim MyCodePane As CodePane
   Dim Str$

   Set MyCodePane = VBInstance.ActiveCodePane
   With Sys.Edit
      LinhaCab = .CharComment
      Select Case Index
          Case 1:
             LinhaCab = LinhaCab & LoadRes(gwDesenvolvedor)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwDesenvolvedor)))
             LinhaCab = LinhaCab & " : " & .Desenvolvedor
          Case 2:
             LinhaCab = LinhaCab & LoadRes(gwWebSite)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwWebSite)))
             LinhaCab = LinhaCab & " : " & .WebSite
          Case 3:
             LinhaCab = LinhaCab & LoadRes(gweMail)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gweMail)))
             LinhaCab = LinhaCab & " : " & .eMail
          Case 4:
             LinhaCab = LinhaCab & LoadRes(gwTelefone)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwTelefone)))
             LinhaCab = LinhaCab & " : " & .Telefone
          Case 5:
             LinhaCab = LinhaCab & LoadRes(gwData_Hora)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwData_Hora)))
             LinhaCab = LinhaCab & " : " & Format(Now, "dd/mm/yyyy hh:mm:ss")
          Case 6:
             NmProjeto$ = VBInstance.ActiveVBProject.Name & " (" & VBInstance.ActiveVBProject.FileName & ")"
             LinhaCab = LinhaCab & LoadRes(gwNome_do_Projeto)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwNome_do_Projeto)))
             LinhaCab = LinhaCab & " : " & NmProjeto$
          Case 7:
             NmModulo$ = MyCodePane.CodeModule.Parent.Name
             If Trim(NmFuncao) <> "" Then
                NmModulo$ = MyCodePane.CodeModule.Parent.Name
                NmModulo$ = NmModulo$ & " (" & GetNameFromPath(MyCodePane.CodeModule.Parent.FileNames(1)) & ")"
             End If
             LinhaCab = LinhaCab & LoadRes(gwNome_do_Modulo)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwNome_do_Modulo)))
             LinhaCab = LinhaCab & " : " & NmModulo$
          Case 8:
             NmArq$ = GetNameFromPath(MyCodePane.CodeModule.Parent.FileNames(1))
             LinhaCab = LinhaCab & LoadRes(gwNome_do_Arquivo)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwNome_do_Arquivo)))
             LinhaCab = LinhaCab & " : " & NmArq$
          Case 9:
             If Trim(NmFuncao) = "" Then
                MyCodePane.GetSelection StartLine&, StarColumn&, EndLine&, EndColumn&
                NmFuncao = MyCodePane.CodeModule.ProcOfLine(StartLine&, vbext_ProcKind.vbext_pk_Proc)
             End If
             LinhaCab = LinhaCab & LoadRes(gwNome_da_Funcao)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwNome_da_Funcao)))
             LinhaCab = LinhaCab & " : " & NmFuncao
          Case 10:
             LinhaCab = LinhaCab & LoadRes(gwParametros)
             LinhaCab = LinhaCab & Space(16 - Len(LoadRes(gwParametros)))
             LinhaCab = LinhaCab & " : "
          Case 11:
             Str$ = ""
             Str$ = Str$ & .CharComment
             Str$ = Str$ & LoadRes(gwComentario)
             Str$ = Str$ & Space(16 - Len(LoadRes(gwComentario)))
             Str$ = Str$ & " : " & vbCrLf
             Str$ = Str$ & .CharComment & vbCrLf
             Str$ = Str$ & .CharComment
             LinhaCab = Str$
          Case 12:
             LinhaCab = LinhaCab & .UserComment
          Case Else:
             LinhaCab = ""
       End Select
   End With
End Function

Public Sub SaveConfig()
   Dim lODBC$, lVersao$, lName$
   Dim Pos%, n As Variant
   With Sys
      With .Edit
         Call SaveSetting(.AppName, "Outros", "ExeAuxiliar", .ExeAuxiliar)
      
         Call SaveSetting(.AppName, "Desenvolvedor", "Nome", .Desenvolvedor)
         Call SaveSetting(.AppName, "Desenvolvedor", "WebSite", .WebSite)
         Call SaveSetting(.AppName, "Desenvolvedor", "eMail", .eMail)
         Call SaveSetting(.AppName, "Desenvolvedor", "Telefone", .Telefone)
         Call SaveSetting(.AppName, "General", "Idioma", .Idioma)
         Call SaveSetting(.AppName, "Error", "ErrorLabel", .ErrorLabel)
         Call SaveSetting(.AppName, "Error", "ErrorFunction", .ErrorFunction)
         Call SaveSetting(.AppName, "Error", "SaidaLabel", .SaidaLabel)
         Call SaveSetting(.AppName, "Error", "SaidaFunction", .SaidaFunction)
         Call SaveSetting(.AppName, "Indentar", "SpcIndent", .SpcIndent)
         Call SaveSetting(.AppName, "Template", "Nome", .Template)
         Call SaveSetting(.AppName, "Template", "UserComment", .UserComment)
         Call SaveSetting(.AppName, "Template", "CharComment", .CharComment)
         For Each n In .Templates
            Pos = InStr(n, " ")
            If Pos <> 0 Then
               Call SaveSetting(.AppName, "Templates", Mid(n, 1, Pos - 1), Mid(n, Pos + 1))
            End If
         Next
         Call SaveSetting(.AppName, "Indent", "IndentFunction", .IndentFunction)
         Call SaveSetting(.AppName, "Indent", "IndentComment", .IndentComment)
         Call SaveSetting(.AppName, "Indent", "IndentSelect", .IndentSelect)
         Call SaveSetting(.AppName, "Indent", "LineBlankBefore", .LineBlankBefore)
      End With
      With .Constru
         Call SaveSetting(.AppName, "Setup", "LoadIni", .LoadIni)
         Call SaveSetting(.AppName, "Setup", "ExibeSubPasta", .ExibeSubPasta)
         Call SaveSetting(.AppName, "Setup", "SalvarOnLine", .SalvarOnLine)
      End With
   End With

'   lODBC = Me.OptODBC(0)
'   lVersao = UCase(Me.CmbVersao.List(Me.CmbVersao.ListIndex))
'   lName = UCase(Me.TxtDbName)
   
   '*** [ Database Format ] ***
'   Call SaveSetting(Sys.AppName, "Database Format", "DBODBC", lODBC$)
'   Call SaveSetting(Sys.AppName, "Database Format", "DBVERSAO", lVersao$)
'   Call SaveSetting(Sys.AppName, "Database Format", "DBNAME", lName$)
   
   '*** [ Database Drive ] ***
'   Call SaveSetting(Sys.AppName, "Database Drive", "DBDRIVE", Me.TxtDbDrive)
'   Call SaveSetting(Sys.AppName, "Database Drive", "DRVRPT", Me.TxtDrvRpt)
   
   '*** [ Setup ] ***
'   Call SaveSetting(Sys.AppName, "Setup", "MICRO", Me.TxtMicro)
'   Call SaveSetting(Sys.AppName, "Setup", "FUNDOTELA", Me.ShpFundo.Tag)
'   Call SaveSetting(Sys.AppName, "Setup", "IDIOMA", Me.CmbIdioma)
'   Call SaveSetting(Sys.AppName, "Setup", "DrvErro", Me.TxtDrvErro)
'
'   Call SaveSetting(Sys.AppName, "Setup", "LOADINI", Me.LstSetup.Selected(0))
'   Call SaveSetting(Sys.AppName, "Setup", "EXIBESUBPASTA", Me.LstSetup.Selected(1))
End Sub
Public Sub SomarLinha(ByVal StrLinha As String, ByRef NumLinha As Integer)
   Dim Palavras As New Collection
   If Trim(StrLinha) = "" Then
      NumLinha = NumLinha + 1
      Exit Sub
   End If
   Set Palavras = GetPalavras(StrLinha)
   If Not InArray(UCase(Palavras(1)), Array("ATTRIBUTE", "VERSION")) Then
      NumLinha = NumLinha + 1
   End If
   Set Palavras = Nothing
End Sub
Public Function LoadRes(Index As Integer)
   LoadRes = LoadResString(Sys.Proj.Idioma + Index)
End Function
Public Function IndentarFuncao(Optional pCodePane, Optional pMember, Optional TextFunc = "", Optional IndComm, Optional IndFunc, Optional IndSele, Optional SpcInd, Optional LineBlank, Optional ExibeFlood = True) As String
   Dim StrLinha, StrLinhaAdd$, StrLinhaSub$, StrAux$
   Dim LinhaStart, QtdLin%, Pos%, Linha%, NumAux%, PosTxt%
   Dim StartLine&, StarColumn&, EndLine&, EndColumn&
   Dim MyCodePane As New CodePane
   Dim Indent As Integer, Somou As Boolean, LinDcl As Boolean
   Dim DeclarationLines As Boolean
   Dim NmFunc As String, IndentarProjeto As Boolean
   Dim IsMissingTextFunc As Boolean
   Dim TextoFinal As String
   Dim ProcKind As vbext_ProcKind
   
   
   StrLinhaAdd = ""
   StrLinhaAdd = StrLinhaAdd & "|If |With |Do |While |For |Select Case "
   StrLinhaAdd = StrLinhaAdd & "|Case |Else |ElseIf |Enum |Type |"
   StrLinhaSub = ""
   StrLinhaSub = StrLinhaSub & "|End If |End With |Loop |Wend |Next "
   StrLinhaSub = StrLinhaSub & "|End Select |End Sub |End Function "
   StrLinhaSub = StrLinhaSub & "|End Enum |End Type |"
   
   
   On Error GoTo Fim
   
   If IsMissing(IndComm) Then IndComm = Sys.Edit.IndentComment
   If IsMissing(IndFunc) Then IndFunc = Sys.Edit.IndentFunction
   If IsMissing(IndSele) Then IndSele = Sys.Edit.IndentSelect
   If IsMissing(LineBlank) Then LineBlank = Sys.Edit.LineBlankBefore
   If IsMissing(SpcInd) Then SpcInd = Sys.Edit.SpcIndent
   IsMissingTextFunc = (TextFunc = "")
   If IsMissing(pCodePane) Then
      Set MyCodePane = VBInstance.ActiveCodePane
   Else
      IndentarProjeto = True
      Select Case TypeName(pCodePane)
         Case "CodePane":    Set MyCodePane = pCodePane
         Case "CodeModule":  Set MyCodePane = pCodePane.CodePane
         Case "VBComponent": Set MyCodePane = pCodePane.CodeModule.CodePane
      End Select
   End If
   If Not IsMissing(pMember) Then
      NmFunc = pMember.Name
   End If
   
   '* Recuperar Nome da Função
   ProcKind = vbext_ProcKind.vbext_pk_Proc
   If Trim(NmFunc) = "" And IsMissingTextFunc Then
      ExibeFlood = True And ExibeFlood
      Call MyCodePane.GetSelection(StartLine&, StarColumn&, EndLine&, EndColumn&)
      NmFunc = MyCodePane.CodeModule.ProcOfLine(StartLine&, ProcKind)
   Else
      ExibeFlood = False
   End If
   DoEvents
  '* Recuperar Linhas da Função
Inicio:
   StrLinha = ""
   StrAux = ""
   TextoFinal = ""
   If IsMissingTextFunc Then
      If StartLine& <= MyCodePane.CodeModule.CountOfDeclarationLines And NmFunc = "" Then
         LinhaStart = 1
         QtdLin = MyCodePane.CodeModule.CountOfDeclarationLines
      Else
         LinhaStart = MyCodePane.CodeModule.ProcStartLine(NmFunc, ProcKind)
         QtdLin = MyCodePane.CodeModule.ProcCountLines(NmFunc, ProcKind)
      End If
      TextFunc = Trim(MyCodePane.CodeModule.Lines(LinhaStart, QtdLin))
   End If
   DeclarationLines = IsMissingTextFunc And (LinhaStart <= MyCodePane.CodeModule.CountOfDeclarationLines)
   
   '* Eliminar Linhas em Branco antes da Função
   StrLinha = ""
   StrAux = ""
   TextoFinal = ""
   If Trim(TextFunc) = "" Then Exit Function
   Pos = InStr(TextFunc, vbCrLf)
   If Pos > 0 Then
      StrLinha = Trim(Mid(TextFunc, 1, Pos - 1))
      If Mid(StrLinha, 1, 1) = "'" Then StrAux = StrLinha & vbCrLf
      Do While Trim(StrLinha) = "" Or Mid(StrLinha, 1, 1) = "'"
         TextFunc = Mid(TextFunc, Pos + 2)
         Pos = InStr(TextFunc, vbCrLf)
         StrLinha = Trim(Mid(TextFunc, 1, Pos - 1))
         If Mid(StrLinha, 1, 1) = "'" Then
            StrAux = StrLinha & vbCrLf
         Else
            If InStr(StrLinha, NmFunc) <> 0 Then Exit Do
         End If
      Loop
   End If
   TextFunc = StrAux & TextFunc
   
   If LineBlank And Not DeclarationLines Then
      TextFunc = " " & vbCrLf & TextFunc
   End If
   
   '*******************
   '* Indentar Linhas *
   '*******************
   LinDcl = False

   If IndentarProjeto Then
      If Not AtuFlood(-1) Then
         IndentarFuncao = "End"
         GoTo Fim
      End If
   End If
   
   PosTxt = InStr(TextFunc, vbCrLf)
   StrLinha = Trim(Mid(TextFunc, 1, PosTxt - 1))
   TextFunc = Mid(TextFunc, PosTxt + 2)
   Do While PosTxt <> 0
      Linha = Linha + 1
      If ExibeFlood Then
         If Not AtuFlood(Linha, QtdLin, "Pocessando... Linha " & Trim(CStr(Linha)) & " / " & Trim(CStr(QtdLin))) Then
            IndentarFuncao = "End"
            GoTo Fim
         End If
      End If
        
      If InStr(StrLinha, "Sub ") <> 0 Or InStr(StrLinha, "Function ") <> 0 Or InStr(StrLinha, "Property ") <> 0 Then
         LinDcl = True
         Indent = 0
      End If
      
      '*********
      '* Linha de Comentário
      If Mid(StrLinha, 1, 1) = "'" Then
         If Not IndComm Then
            Indent% = Indent% - SpcInd
         End If
      End If
      
      '*********
      '* Linha de Else
      If Mid(Trim(StrLinha), 1, 4) = "Else" Or Mid(Trim(StrLinha), 1, 6) = "ElseIf" Then
         Indent% = Indent% - SpcInd
      End If
      
      '*****************
      '* Identar Linha *
      '*****************
      'Indent% = IIf(Indent% < 0, 0, Indent%)
      If Indent% < 0 Then
         StrLinha = Space(0) & StrLinha
      Else
        StrLinha = Space(Indent%) & StrLinha
      End If
      
      '*********
      '* Linha de Comentário
      If Mid(Trim(StrLinha), 1, 1) = "'" Then
         If Not IndComm Then
            Indent% = Indent% + SpcInd
         End If
      End If
      
      '*********
      '* Linha de Select Case
      If Mid(Trim(StrLinha), 1, Pos + 4) = "Select Case" Then
         Indent% = Indent% + SpcInd
         If IndSele Then
            Indent% = Indent% + SpcInd
         End If
      End If
      
      '*****************************
      '** Recuperar Próxima Linha **
      '*****************************
      If TextFunc = "" Or (Linha = 1 And Trim(StrLinha) = "") Then
         TextoFinal = TextoFinal & StrLinha & IIf(LineBlank, vbCrLf, "")
      Else
        TextoFinal = TextoFinal & StrLinha & vbCrLf
      End If
      PosTxt = InStr(TextFunc, vbCrLf)
      If PosTxt = 0 Then
         StrLinha = Trim(TextFunc)
      Else
         StrLinha = Trim(Mid(TextFunc, 1, PosTxt - 1))
         TextFunc = Mid(TextFunc, PosTxt + 2)
      End If
      
      
      If LinDcl Then
         LinDcl = False
         If IndFunc And Not DeclarationLines Then
            Indent = SpcInd
         Else
            Indent = 0
         End If
      End If
         
      If Somou Then
         Indent = Indent + SpcInd
         Somou = False
      End If
    
      Pos = InStr(StrLinha & " ", " ")
      Select Case Mid(StrLinha, 1, Pos - 1)
         Case "End"
            Pos = InStr(Pos + 1, StrLinha & " ", " ")
            If Pos = 0 Then
               StrAux = StrLinha
            Else
               StrAux = Mid(StrLinha, 1, Pos - 1)
            End If
            If IndSele And StrAux = "End Select" Then
               Indent = Indent - SpcInd
            End If
         Case "If"
            StrAux = Mid(StrLinha, 1, Pos - 1)
            Pos = InStr(StrLinha, " Then")
            If Len(Trim(Mid(StrLinha, Pos + 5))) > 0 And Mid(Trim(Mid(StrLinha, Pos + 5)), 1, 1) <> "'" Then
               StrAux = ""
            End If
         Case "Select"
            If Mid(StrLinha, 1, Pos + 4) = "Select Case" Then
               StrAux = Mid(StrLinha, 1, Pos - 1)
            End If
         Case "Case"
            StrAux = Mid(StrLinha, 1, Pos - 1)
            Indent = Indent - SpcInd
         Case "Public", "Private"
            If Mid(StrLinha, Pos + 1, 4) = "Enum" Or Mid(StrLinha, Pos + 1, 4) = "Type" Then
               StrAux = Mid(StrLinha, Pos + 1, 4)
            End If
         Case Else
            StrAux = Mid(StrLinha, 1, Pos - 1)
      End Select
      
      Somou = False
      Select Case True
         Case (InStr(StrLinhaAdd, "|" & StrAux & " |") <> 0)
            Somou = True
         Case (InStr(StrLinhaSub, "|" & StrAux & " |") <> 0)
            Indent = Indent - SpcInd
         Case Else
            Indent = IIf(Indent = 0, Indent = SpcInd, Indent)
      End Select
   Loop
   If Trim(TextFunc) <> "" Then
      TextoFinal = TextoFinal & Trim(TextFunc)
   End If
   IndentarFuncao = ""
   IndentarFuncao = TextoFinal
   If IsMissingTextFunc Then
      Call MyCodePane.CodeModule.DeleteLines(LinhaStart, QtdLin)
      Call MyCodePane.CodeModule.InsertLines(LinhaStart, IndentarFuncao)
      If ExibeFlood Then Call FimFlood
   End If
   If ProcKind >= 1 And ProcKind <= 2 Then
      ProcKind = ProcKind + 1
      GoTo Inicio
   End If
   Set MyCodePane = Nothing
   Exit Function
Fim:
   If Err = 35 Then '* Sub or Function not defined
      ProcKind = ProcKind + 1
      Resume
   ElseIf Err = 429 Then '* ActiveX component can't create object
      Exit Function
   ElseIf Err <> 0 Then
      MsgBox CStr(Err) & "-" & Err.Description & " ( " & NmFunc & " )"
      Resume Next
   End If
End Function
Public Function FindDeadConstants(Rtf As Control, VbTexto As String, Optional EscopoLocal = True) As String
   Dim FileHandle As Integer
   Dim FileContents As String
   Dim PosOfDcl As Long
   Dim StartOfWord As Long
   
   Dim EndOfWord As Long
   Dim ConstantName As String
   Dim OriginalConstantName As String
   Dim MyMenber As MEMBER
   
   Dim TxtDcl As String
   Dim RtfTexto As String
   Dim CollDcl As New Collection, n As Variant
   
   Dim Pos%, Linha$, Aux As Long
   Dim DeadConst As Boolean
   Dim IniComp As Integer
   Dim LinhaDcl As String
   Dim lSelStart As Long
   Dim i%
   
   '* Abrir o arquivo ou simplesmente recuperar
   '* o texto a ser avaliado
   If FileExists(Trim(VbTexto)) Then
      FileHandle = FreeFile
      Open VbTexto For Binary Access Read As FileHandle
         VbTexto = Input$(LOF(FileHandle), FileHandle)
      Close FileHandle
   Else
      If Rtf.Text = "" Then
         Rtf.Text = VbTexto
      End If
   End If
   RtfTexto = Rtf.Text
'   lSelStart = Rtf.SelStart
   
   'loop through all the module-level
   'constants:
   CollDcl.Add " Dim "
   CollDcl.Add " Const "
   CollDcl.Add " Static "
   If Not EscopoLocal Then
      CollDcl.Add " Public "
      CollDcl.Add " Global "
   End If
   
   For Each n In CollDcl
      TxtDcl = n
      Do
         If InStr(PosOfDcl + 1, " " & RtfTexto, TxtDcl) = 0 Then
            PosOfDcl = InStr(PosOfDcl + 1, RtfTexto, vbCrLf & Trim(TxtDcl) & " ")
            PosOfDcl = IIf(PosOfDcl = 0, 0, PosOfDcl + 2)
         Else
            PosOfDcl = InStr(PosOfDcl + 1, " " & RtfTexto, TxtDcl)
         End If
         
         If PosOfDcl > 0 Then
            'we've found a constant:
            StartOfWord = PosOfDcl + Len(Trim(TxtDcl)) + 1
            Pos = InStr(PosOfDcl + 1, RtfTexto, vbCrLf)
            If Pos <> 0 Then
               LinhaDcl = Mid(RtfTexto, StartOfWord, Pos - StartOfWord)
            End If
            Pos = 0
            Aux = Rtf.GetLineFromChar(StartOfWord)
            For i = 1 To Aux
               Pos = InStr(Pos + 2, RtfTexto, vbCrLf)
            Next
            Pos = Pos + 2
            Aux = InStr(Pos + 2, RtfTexto, vbCrLf)
            Aux = IIf(Aux = 0, Len(RtfTexto), Aux)
            If Mid(Trim(Mid(RtfTexto, Pos, Aux - Pos)), 1, 1) <> "'" Then
               Do
                  OriginalConstantName = Trim(RichWordOver(Rtf, 0, 0, StartOfWord))
                  ConstantName = OriginalConstantName
                  EndOfWord = StartOfWord + Len(OriginalConstantName)
         
                  'if the constant is not
                  'referenced beyond its
         
                  'declaration, then it's dead:
                  IniComp = IIf(RtfTexto = VbTexto, EndOfWord, 1)
                  If (0 = InVbWord(IniComp, LCase(VbTexto), LCase(ConstantName))) Then
                     DeadConst = True
                     If VarWithDclImplicit(LCase(ConstantName)) Then
                           ConstantName = Mid(ConstantName, 1, Len(ConstantName) - 1)
                           If (0 <> InVbWord(IniComp, LCase(VbTexto), LCase(ConstantName))) Then
                              DeadConst = False
                           End If
                     End If
                  Else
                     DeadConst = False
                  End If
                  
                  If Not DeadConst Then
                     Aux = 0
                     IniComp = InVbWord(IniComp + 1, LCase(VbTexto), LCase(ConstantName))
                     Do
                        If IniComp <> 0 Then
                           Pos = InStr(IniComp, VbTexto, vbCrLf)
'                           While InStr(Pos - Aux, VbTexto, vbCrLf) = Pos And InStr(Pos - Aux, VbTexto, vbCrLf) <> 0 And (Pos - Aux > 0)
'                               Aux = Aux + 2
'                           Wend
                           i = 0
                           While (Asc(Mid(VbTexto, IniComp - i, 1)) <> 13 And Asc(Mid(VbTexto, IniComp - i, 1)) <> 10) Or IniComp = i
                              i = i + 1
                           Wend
                           Aux = IniComp - i + 1
                           Linha = Mid(VbTexto, Aux, Pos - Aux)
'                           Linha = Mid(VbTexto, InStr(Pos - Aux, VbTexto, vbCrLf) + 2, Pos - InStr(Pos - Aux, VbTexto, vbCrLf) - 2)
                           '* se Linha Comentário então a Variável é morta
                           If Mid(Trim(Linha), 1, 1) = "'" Then
                              DeadConst = True
                           Else
                              If InArray(RichWordOver(Trim(Linha), 0, 0, 1), Array("Dim", "Const", "Static", "Public", "Global")) Then
                                 Aux = InStr(IniComp, VbTexto, "End Sub")
                                 Pos = InStr(IniComp, VbTexto, "End Function")
                                 If Pos = 0 Or Aux = 0 Then
                                    Aux = IIf(Aux < Pos, Pos, Aux)
                                 Else
                                   Aux = IIf(Aux < Pos, Aux, Pos)
                                 End If
                                 Pos = InStr(IniComp, VbTexto, "End Property")
                                 If Pos <> 0 Then
                                    Aux = IIf(Aux < Pos, Aux, Pos)
                                 End If
                                 If InArray(Trim(TxtDcl), Array("Public", "Global", "Const")) Then
                                    IniComp = Aux
                                 Else
                                    IniComp = InVbWord(IniComp + 1, LCase(VbTexto), LCase(ConstantName))
                                 End If
                                 
                                 If IniComp = 0 Then
                                    DeadConst = True
                                    Exit Do
                                 ElseIf IniComp < Aux Then
                                    DeadConst = False
                                    Exit Do
                                 Else
                                    DeadConst = True
                                 End If
                              Else
                                 DeadConst = False
                                 Exit Do
                              End If
                           End If
                           IniComp = InVbWord(IniComp + 1, LCase(VbTexto), LCase(ConstantName))
                        End If
                     Loop Until IniComp = 0
                  End If
                  '* Colorir todas as referências.
         '               If Not DeadConst Then
         '                  Pos = InStr(LCase(VbTexto), LCase(ConstantName))
         '                  While Pos <> 0
         '                     Rtf.SelStart = Pos
         '                     Rtf.SelLength = Len(OriginalConstantName)
         '                     Rtf.SelColor = vbCyan
         '                     Rtf.SelText = OriginalConstantName
         '                     Pos = InStr(Pos + Len(ConstantName), LCase(VbTexto), LCase(ConstantName))
         '                  Wend
         '               End If
                  If DeadConst Then
                     With Rtf
                        .SelStart = StartOfWord - 1
                        .SelLength = EndOfWord - StartOfWord
                        'Call SetBackColorSel(Rtf, iif(DeadConst,vbRed, vbWhite))
                        .SelColor = IIf(DeadConst, vbRed, vbBlack)
                        .SelStrikeThru = DeadConst
                        .SelBold = DeadConst
                        'FrmAddIn.TxtCode.SelText
                     End With
                  End If
         
                  If InStr(LinhaDcl, ",") <> 0 Then
                     StartOfWord = StartOfWord + InStr(LinhaDcl, ",") + 1
                     LinhaDcl = Trim(Mid(LinhaDcl, InStr(LinhaDcl, ",") + 1))
                  Else
                     StartOfWord = -1
                  End If
               Loop Until (StartOfWord = -1)
            End If
         End If
      Loop Until PosOfDcl = 0
   Next
   DoEvents
'   Rtf.SelStart = lSelStart
   Set CollDcl = Nothing
End Function
Public Function InVbWord(Optional Start, Optional VbTexto, Optional VbWord)
   Start = IIf(IsMissing(Start), 1, Start)
   InVbWord = 0
   If IsMissing(VbTexto) Or IsMissing(VbTexto) Then
      Exit Function
   End If
   Select Case True
      Case InStr(Start, LCase(VbTexto), LCase(" " & VbWord & vbCrLf)) <> 0
         InVbWord = InStr(Start, LCase(VbTexto), LCase(" " & VbWord & vbCrLf)) + 1
      Case InStr(Start, LCase(VbTexto), LCase(vbCrLf & VbWord & " ")) <> 0
         InVbWord = InStr(Start, LCase(VbTexto), LCase(vbCrLf & VbWord & " ")) + 2
      Case InStr(Start, LCase(VbTexto), LCase(" " & VbWord & ",")) <> 0
         InVbWord = InStr(Start, LCase(VbTexto), LCase(" " & VbWord & ",")) + 1
      Case InStr(Start, LCase(VbTexto), LCase(" " & VbWord & ")")) <> 0
         InVbWord = InStr(Start, LCase(VbTexto), LCase(" " & VbWord & ")")) + 1
      Case InStr(Start, LCase(VbTexto), LCase(" " & VbWord & "(")) <> 0
         InVbWord = InStr(Start, LCase(VbTexto), LCase(" " & VbWord & "(")) + 1
      Case InStr(Start, LCase(VbTexto), LCase("." & VbWord & " ")) <> 0
         InVbWord = InStr(Start, LCase(VbTexto), LCase("." & VbWord & " ")) + 1
      Case InStr(Start, LCase(VbTexto), LCase(" " & VbWord & ".")) <> 0
         InVbWord = InStr(Start, LCase(VbTexto), LCase(" " & VbWord & ".")) + 1
      Case InStr(Start, LCase(VbTexto), LCase(" " & VbWord & " ")) <> 0
         InVbWord = InStr(Start, LCase(VbTexto), LCase(" " & VbWord & " ")) + 1
   End Select
   If InVbWord = 0 Then
      If VarWithDclImplicit(CStr(VbWord)) Then
         InVbWord = InVbWord(Start, VbTexto, Mid(VbWord, 1, Len(VbWord) - 1))
      End If
   End If
End Function
                       
Public Sub SetBackColorSel(RtfHWnd As Long, vbRed)
    Dim tCF2 As CHARFORMAT2
    tCF2.dwMask = CFM_BACKCOLOR
    tCF2.crBackColor = TranslateColor(vbRed)
    tCF2.cbSize = Len(tCF2)
    'Call SendMessage(RtfHWnd, EM_SETCHARFORMAT, &H1&, tCF2)
    'Call SendMessage(RtfHWnd, (WM_USER + 68), &H1&, tCF2)
    Call SendMessage(RtfHWnd, &H400 + 68, &H1&, tCF2)
End Sub
Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function
Public Function DoAutocomplete(ObjX As Object) As Long
    Dim hWndEdit As Long
    
    If TypeOf ObjX Is TextBox Then
        ' Just set the edit field hWnd to the
        ' textbox hWnd as a textbox is an
        ' edit field
        hWndEdit = ObjX.hwnd
    ElseIf TypeOf ObjX Is ComboBox Then
        ' Get edit field of the combobox
'        hWndEdit = FindWindowEx(ObjX.hwnd, 0, "EDIT", vbNullString)
    Else
        ' No edit field
        DoAutocomplete = 0
        Exit Function
    End If
    
    ' Apply the autocomplete functionality
'    DoAutocomplete = SHAutoComplete(hWndEdit, SHACF_DEFAULT)
    
End Function



Sub CreateMenu(VBInst As VBIDE.Application)
   Dim FrmCurr As Object
   Dim MenuFile As Object
   Dim Mnu As Object
   
   On Error Resume Next
   
   Set FrmCurr = VBInst.ActiveProject.ActiveForm
   Set MenuFile = FrmCurr.AddMenuTemplate("MenuFile", Nothing)
   MenuFile.Properties("Caption").Value = "&File"
   
   Set Mnu = FrmCurr.AddMenuTemplate("File", MenuFile)
   Mnu.Properties("Caption").Value = "&New"
   Mnu.Properties("Index").Value = 0
   
End Sub
