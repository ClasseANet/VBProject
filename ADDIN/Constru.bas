Attribute VB_Name = "Constru"
Option Explicit
Global DB         As New DS_BANCO
Global gConstru   As New CONSTRUTOR
Global glbProj    As VBProject
Global SysMdi     As New FrmAddIn

Global Sys           As New Set_AddIn

Global VBInstance    As VBIDE.VBE       'instance of VB IDE
Global VbApplication As VBIDE.Application
Global gWinConstru   As Form            'used to make sure we only run one instance
Global gWinTabOrder  As VBIDE.Window    'used to make sure we only run one instance
Global gDocTabOrder  As Object          'user doc object

Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetParent0 Lib "user32" (ByVal hwnd&) As Long
Public Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal msg&, ByVal wp&, ByVal lp&)
Public Declare Sub SetFocus Lib "user32" (ByVal hwnd&)

Global Const WM_SYSKEYDOWN = &H104
Global Const WM_SYSKEYUP = &H105
Global Const WM_SYSCHAR = &H106
Global Const VK_F = 70  ' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"
Dim hWndMenu       As Long           'needed to pass the menu keystrokes to VB



'********************************************************************
'* This sub should be executed from the Immediate window in         *
'* order to get this app added to the VBADDIN.INI  file you         *
'* you must change the name in the 2nd argument to reflecty         *
'* the correct name of your project                                 *
'********************************************************************
Public Sub AddToINI()
   Dim ErrCode As Long
   ErrCode = WritePrivateProfileString("Add-Ins32", "VbEditorUtil.Connect", "0", "vbaddin.ini")
   Debug.Print IIf(ErrCode = 1, "Conectado", "Desconectado")
End Sub
Function InRunMode(VBInst As VBIDE.VBE) As Boolean
   InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
End Function
Sub HandleKeyDown(ud As Object, KeyCode As Integer, Shift As Integer)
   If Shift <> 4 Then Exit Sub
   If KeyCode < 65 Or KeyCode > 90 Then Exit Sub
   If VBInstance.DisplayModel = vbext_dm_SDI Then Exit Sub

   If hWndMenu = 0 Then hWndMenu = FindHwndMenu(ud.hwnd)
   PostMessage hWndMenu, WM_SYSKEYDOWN, KeyCode, &H20000000
   KeyCode = 0
   SetFocus hWndMenu
End Sub

Function FindHwndMenu&(ByVal hwnd&)
   Dim h As Long

Loop2:
   h = GetParent0(hwnd)
   If h = 0 Then FindHwndMenu = hwnd: Exit Function
   hwnd = h
   GoTo Loop2
End Function
Public Function PalavraReservada(pTxt As String) As Boolean
   pTxt = Trim(UCase(pTxt))
   PalavraReservada = True
   Select Case True
      Case InArray(pTxt, Array("IF", "THEN", "ELSE", "END", "SELECT", "CASE", "DO", "LOOP", "UNTIL"))
      Case InArray(pTxt, Array("FOR", "EACH", "IN", "TO", "NEXT", "WHILE", "WEND", "WITH", "EXIT"))
      Case InArray(pTxt, Array("GLOBAL", "PUBLIC", "PRIVATE", "LOCAL", "DIM", "SUB", "FUNCTION", "PROPERTY"))
      Case InArray(pTxt, Array("STATIC", "DIM", "REDIM", "CONST", "ENUM", "PRESERV"))
      Case InArray(pTxt, Array("BOOLEAN", "STRING", "LONG", "INTEGER", "VARIANT", "DATE"))
      Case InArray(pTxt, Array("SINGLE", "BYTE", "CURRENCY", "DOUBLE", "OBJECT"))
      Case InArray(pTxt, Array("TRUE", "FALSE", "AND", "OR", "NOT"))
      Case InArray(pTxt, Array("CALL", "GET", "SET", "LET", "NEW", "NOTHING", "EMPTY"))
      Case InArray(pTxt, Array("AS", "ON", "STEP", "GOTO", "ERROR", "NOTHING", "EMPTY"))
      Case InArray(pTxt, Array("DEBUG", "PRINT", "INPUT", "OUTPUT", "OPEN", "CLOSE", "LINE"))
      Case InArray(pTxt, Array("BYVAL", "BYREF", "OPTIONAL"))
      Case InArray(pTxt, Array("OPTION", "EXPLICIT", "WITHEVENTS", "RAISEEVENT", "EVENT"))
      Case Else: PalavraReservada = False
   End Select

End Function
Public Sub IniFlood()
   With FrmFlood
      .Show
      .ZOrder 0
      .LblPercent = "0%"
      .LblPercent.Refresh
      .Left = Screen.Width - .Frme.Width - 240
      .Top = 240

      DoEvents
      .Frme.Enabled = True
      .Visible = True
      .Frme.Visible = True
   End With
End Sub
Public Function AtuFlood(ByVal Value As Integer, Optional Total, Optional Str) As Boolean
   Dim TimeDiff&, TmpEstimado&
   Dim StrMin$, StrSeg$

   AtuFlood = True
   DoEvents
   
   If Not IsMissing(Total) Then
      Total = IIf(Total <= 0, 1, Total)
      Value = (Value / Total) * 100
      Value = IIf(Value >= 100, 100, Value)
   End If
   With FrmFlood
      If .Cancel Then
         .Cancel = False
         AtuFlood = False
         Call FimFlood
         Unload FrmFlood
         Exit Function
      End If
   
      If Not .Visible Then Call IniFlood

      If Not IsMissing(Str) Then
         .Frme.Caption = Trim(CStr(Str))
         .Frme.Refresh
      End If

      If Value <> .ProgBar.Value Then
         .ZOrder 0
         .ProgBar.Value = IIf(Value < 0, .ProgBar.Value, Value)
         .LblPercent = Trim(CStr(.ProgBar.Value)) & "%"
         .LblPercent.Refresh
      End If
   End With
End Function
Public Sub FimFlood()
   Dim n
   With FrmFlood
      If Not .Cancel Then
         .ZOrder 0
         .ProgBar.Value = 100
         .LblPercent = "100%"
         .LblPercent.Refresh
      End If
   End With
   Unload FrmFlood
End Sub

Public Sub GetPropPage(pPage As PROPPAGE, pVbComp As VBComponent, Optional Obj As Variant)
   Dim MyProp As New PROPPAGE
   Dim Img$

   Dim LinhaInicial%, PosRetorno%, PosIniArg%, PosAux%
   Dim StrFunc$, StrAux$, Palavras As Collection
   Dim MyArg As ADDARG, n
   Dim Parenteses%, i%, Pos%, isFunc As Boolean
   Dim StrFrase As String, StrWord As String
   
   Set MyProp = pPage
   With MyProp

      'Set .PROJETO = mvarPROJETO

      'Select Case mvarMe.ActiveControl.Name
      '   Case "LstItens", "TabComp"
      '      Img$ = mvarMe.LstItens.SelectedItem.Icon
      '   Case "TreProj"
      '      Img$ = mvarMe.TreProj.SelectedItem.Image
      'End Select
      Img = "CLASSE"
      If False Then
         Select Case Img$
            Case "CLASSE": .TipoPagina = tpClasse
            Case "COLECAO": .TipoPagina = tpColecao
            Case "EVENTO": .TipoPagina = tpEvento
            Case "FORMULARIO": .TipoPagina = tpForm
            Case "METODO": .TipoPagina = tpMetodo
            Case "MODULO": .TipoPagina = tpModulo
            Case "PROPRIEDADE": .TipoPagina = tpPropriedade
            Case "VARIAVEL", "CONSTANTE": .TipoPagina = tpPropriedade
            Case Else: .TipoPagina = tpNull
         End Select
      End If
      On Error Resume Next
      'On Error GoTo 0
      '* Definir Retorno e Argumentos

      'LinhaInicial = mvarVbComp.CodeModule.ProcStartLine(mvarVbMember.Name, vbext_ProcKind.vbext_pk_Proc)
      LinhaInicial = pVbComp.CodeModule.ProcBodyLine(Obj, vbext_ProcKind.vbext_pk_Proc)
      If Err = 35 Or LinhaInicial = 0 Then
         LinhaInicial = 1
      End If
      
      StrFunc$ = pVbComp.CodeModule.Lines(LinhaInicial, 1)
      StrFrase = StrFunc$
      isFunc = (InStr(UCase(StrFunc), "FUNCTION") <> 0)
      If isFunc Then .Retorno = "Variant"
      PosRetorno = InStr(UCase(StrFrase), ")")
      While PosAux <> 0
         StrFrase = Trim(Mid(StrFrase, PosRetorno + 1))
         PosRetorno = InStr(UCase(StrFrase), ")")
      Wend
      PosIniArg% = InStr(UCase(StrFunc$), "(") + 1
      PosRetorno = IIf(PosRetorno = PosIniArg%, 0, PosRetorno)
      If PosRetorno > 1 Then
         StrFrase = Mid(StrFunc$, PosIniArg%, PosRetorno - PosIniArg%)
      Else
        StrFrase = ""
      End If
      
      Parenteses = 1
      StrAux$ = Mid(StrFunc$, PosRetorno + 1)
      If StrAux$ <> "" Then
         PosRetorno = InStr(UCase(StrAux$), "AS ")
         If PosRetorno <> 0 Then
            .Retorno = Trim(Mid(StrAux$, PosRetorno + 3))
         End If
      End If
      Call .SetArgumentos(StrFrase)
'     .NOME = mvarVbMember.Name
'     .Descricao = mvarVbMember.Description
'     .HelpID = mvarVbMember.HelpContextID
   End With
End Sub
'****************************
'****************************
'****************************
' *** Add VB code in a RTF control
'
'  Call InitColorize
'  Call ColorizeWords(rtfVBCode)
'
' *** Now your VB code in your RTF control is colorized
'Source Code:
'
' #VBIDEUtils#************************************************************
' * Programmer Name  : Waty Thierry
' * Web Site    : www.geocities.com/ResearchTriangle/6311/
' * E-Mail      : waty.thierry@usa.net
' * Date     : 30/10/98
' * Time     : 14:47
' * Module Name    : Colorize_Module
' * Module Filename  : Colorize.bas
' **********************************************************************
' * Comments    : Colorize in black, blue, green the VB keywords
' *
' *
' **********************************************************************
'Option Explicit
'Private gsBlackKeywords    As String
'Private gsBlueKeyWords     As String

'Public Sub InitColorize()
'   ' #VBIDEUtils#************************************************************
'   ' * Programmer Name  : Waty Thierry
'   ' * Web Site         : www.geocities.com/ResearchTriangle/6311/
'   ' * E-Mail           : waty.thierry@usa.net
'   ' * Date             : 30/10/98
'   ' * Time             : 14:47
'   ' * Module Name      : Colorize_Module
'   ' * Module Filename  : Colorize.bas
'   ' * Procedure Name   : InitColorize
'   ' * Parameters       :
'   ' **********************************************************************
'   ' * Comments         : Initialize the VB keywords
'   ' *
'   ' *
'   ' **********************************************************************
'
'   gsBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
'   gsBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*"
'
'End Sub
Public Sub SetAttribute(ByVal StrLinha As String, ByRef vObj As Object)
   Dim Palavras As New Collection
   Dim i%
   Dim StrAux$
   Dim ValueOk As Boolean
   If Trim(StrLinha) = "" Then Exit Sub
   Set Palavras = GetPalavras(StrLinha)
   If UCase(Palavras(1)) = "ATTRIBUTE" Then
      For i = 1 To Palavras.Count
         If ValueOk And Palavras(i) <> """" Then
            StrAux$ = StrAux$ & " " & Palavras(i)
         End If
         If Palavras(i) = "=" Then ValueOk = True
      Next
      StrAux$ = Trim(StrAux$)
      Select Case Palavras(4)
         Case "VB_Description": vObj.PagProp.Descricao = StrAux$
         Case "VB_VarHelpID":   vObj.PagProp.HelpID = StrAux$
      End Select
   End If
   Set Palavras = Nothing
End Sub
Public Function GetStartLine(strFileName As String, StrObj As String) As Long
   Dim nArq As Integer
   Dim Textline As String
   Dim LinhaCorrente As Long
   Dim CodeLine As Boolean
   Dim Achou As Boolean
   Dim Arr As Variant

   If Trim(strFileName) = "" Or Trim(StrObj) = "" Then
      Exit Function
   End If
   StrObj = UCase(StrObj)
   nArq = FreeFile
   On Error Resume Next
   
   Open strFileName For Input Shared As #nArq
   LinhaCorrente = 1
   CodeLine = False
   Do While Not EOF(nArq)
      Line Input #nArq, Textline
      Textline = Trim(Textline)
      If Not CodeLine Then
         Arr = Array("Option", "Public", "Global", "Private", "Dim", "Static", "Function", "Sub", "Property")
         CodeLine = InArray(RichWordOver(Textline, 0, 0, 1), Arr)
         CodeLine = CodeLine Or (Mid(Textline, 1, 1) = "'")
      End If
      If CodeLine Then
         Achou = (InStr(UCase(Textline), StrObj) <> 0 And Mid(Textline, 1, 1) <> "'")
         If Achou Then Exit Do
         LinhaCorrente = LinhaCorrente + 1
      End If
   Loop
   Close #nArq
   GetStartLine = IIf(Achou, LinhaCorrente, 0)
End Function

