VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSR"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Enum eMoeda
   Numerico
   Real
   Dolar
   Outros
End Enum

Enum eHabilitarTela
   EnabledTrue
   EnabledFalse
   EnabledDefault
End Enum

Enum Resp
   SIM = 0
   NAO = 1
End Enum

Private mvarSepDec$
Private mvarSepMil$
Private mvarSepDt$

Private mvarDtMaskAux$
Private mvarDtMask$
Private mvarFormatoData$
Private TaskBarhWnd As Long
Public Property Get SepDec() As String
    SepDec = mvarSepDec
End Property
Public Property Let SepDec(Txt$)
    mvarSepDec = Txt$
End Property
Public Property Get FormatoData() As String
    FormatoData = mvarFormatoData
End Property
Public Property Let FormatoData(Txt$)
    mvarFormatoData = Txt$
End Property

Public Property Get SepDt() As String
    SepDt = mvarSepDt
End Property
Public Property Let SepDt(Txt$)
    mvarSepDt = Txt$
End Property
Public Property Get SepMil() As String
    SepMil = mvarSepMil
End Property

Public Property Let SepMil(Txt$)
    mvarSepMil = Txt$
End Property
Private Sub AbrirDialUp(Conexão As String)
   Dim Exec
   Exec = Shell("C:\Windows\rundll32.exe rnaui.dll,RnaDial " & Conexão, 1)
   AppActivate Exec
   SendKeys "{enter}"
   DoEvents
End Sub
Public Function AlterItem(ByRef cCol As Collection, pKey As String, pItem As Variant) As Boolean
' **********************************************************************
' * Date            : 27/03/03
' * Time            : 14:47
' * Module Name     : AlterItem
' * Module Filename : DSR.cls
' * Procedure Name  : DSR
' * Parameters      : ByRef cCol As Collection, pKey As String, pItem As Variant
' **********************************************************************
' * Comments        : Altera o pItem com a chave pKey numa coleção cCol.
' *                   Caso pKey não exista cadastra o pItem
' *
' **********************************************************************
   On Error GoTo AlterErr
   If ExisteItem(cCol, pKey) Then
      cCol.Remove pKey
   End If
   cCol.Add Item:=pItem, Key:=pKey
   AlterItem = True
   
   GoTo Fim
   
AlterErr:
   AlterItem = False
   Resume Fim
Fim:
End Function
Public Function Aspas(Txt As String) As String
Attribute Aspas.VB_Description = "Inclui o Texto informado entre Aspas Simples ( "" ' "" ) para implemntar uma ""Query"""
   Aspas = "'" & Txt & "'"
End Function
Public Sub ColorizeVbWords(Rtf As Object, Optional PosIni, Optional PosFim, Optional ViewFlood = True, Optional LeftFlood, Optional TopFlood, Optional WidthFlood)
' **********************************************************************
' * Date            : 30/10/98
' * Time            : 14:47
' * Module Name     : Colorize_Module
' * Module Filename : Colorize.bas
' * Procedure Name  : ColorizeWords
' * Parameters      :  rtf As RichTextBox
' **********************************************************************
' * Comments        : Colorize in black, blue, green the VB keywords
' *
' *
' **********************************************************************
   
   Dim sBuffer    As String
   Dim nI         As Long
   Dim nJ         As Long
   Dim sTmpWord   As String
   Dim nStartPos  As Long
   Dim nSelLen    As Long
   Dim nWordPos   As Long
   Dim nWordPosBlack As Long
   Dim nWordPosBlue As Long
   Dim nLinBefore As Long
   Dim LenWord    As Integer
   Dim gsBlackKeyWords As String
   Dim gsBlueKeyWords  As String
   Dim gsErrKeyWords As String
   Dim lAux As Long
   
   Dim ProgBar As New CProgBar32, Value As Integer

   'Dim cHourglass  As class_Hourglass
   'Set cHourglass = New class_Hourglass
   gsBlackKeyWords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin"
   gsBlackKeyWords = gsBlackKeyWords & "*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear"
   gsBlackKeyWords = gsBlackKeyWords & "*Collection*Command*Cos*CreateObject*CurDir"
   gsBlackKeyWords = gsBlackKeyWords & "*DateAdd*DateDiff*DatePart*DateSerial*DateValue"
   gsBlackKeyWords = gsBlackKeyWords & "*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty"
   gsBlackKeyWords = gsBlackKeyWords & "*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime"
   gsBlackKeyWords = gsBlackKeyWords & "*FileLen*Fix*Format*FV*GetAllSettings*GetAttr"
   gsBlackKeyWords = gsBlackKeyWords & "*GetObject*GetSetting*Hex*Hide*Hour*InputBox"
   gsBlackKeyWords = gsBlackKeyWords & "*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty"
   gsBlackKeyWords = gsBlackKeyWords & "*IsError*IsMissing*IsNull*IsNumeric*IsObject"
   gsBlackKeyWords = gsBlackKeyWords & "*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log"
   gsBlackKeyWords = gsBlackKeyWords & "*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer"
   gsBlackKeyWords = gsBlackKeyWords & "*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate"
   gsBlackKeyWords = gsBlackKeyWords & "*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd"
   gsBlackKeyWords = gsBlackKeyWords & "*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn"
   gsBlackKeyWords = gsBlackKeyWords & "*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv"
   gsBlackKeyWords = gsBlackKeyWords & "*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial"
   gsBlackKeyWords = gsBlackKeyWords & "*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType"
   gsBlackKeyWords = gsBlackKeyWords & "*WeekDay*Width*Year*"
   
   
   gsBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary"
   gsBlueKeyWords = gsBlueKeyWords & "*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur"
   gsBlueKeyWords = gsBlueKeyWords & "*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const"
   gsBlueKeyWords = gsBlueKeyWords & "*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare"
   gsBlueKeyWords = gsBlueKeyWords & "*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec"
   gsBlueKeyWords = gsBlueKeyWords & "*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim"
   gsBlueKeyWords = gsBlueKeyWords & "*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase"
   gsBlueKeyWords = gsBlueKeyWords & "*Error*Event*Exit*Explicit*False*For*Function*Get"
   gsBlueKeyWords = gsBlueKeyWords & "*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer"
   gsBlueKeyWords = gsBlueKeyWords & "*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet"
   gsBlueKeyWords = gsBlueKeyWords & "*Module*Name*New*Next*Not*Object*On*Open*Option*Or"
   gsBlueKeyWords = gsBlueKeyWords & "*Output*Print*Private*Property*Public*Put*Random"
   gsBlueKeyWords = gsBlueKeyWords & "*Read*ReDim*Resume*Return*RSet*Seek*Select*Set"
   gsBlueKeyWords = gsBlueKeyWords & "*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then"
   gsBlueKeyWords = gsBlueKeyWords & "*True*Type*UBound*Unlock*Variant*Wend*While*With"
   gsBlueKeyWords = gsBlueKeyWords & "*Xor*Nothing*To*"
   
   gsBlueKeyWords = gsBlueKeyWords & "*End If*End Function*End Sub"
   gsBlueKeyWords = gsBlueKeyWords & "*Select Case*Case Else*"
   
   gsErrKeyWords = "*ENDIF*ENDFUNCTION*ENDSUB*ENDWITH"
   gsErrKeyWords = gsErrKeyWords & "*SELECTCASE*CASEELSE*"
   
   sBuffer = Rtf.Text
   
   sTmpWord = ""
   With Rtf
      .Visible = False
      nLinBefore = .SelStart
      If IsMissing(PosIni) Then PosIni = 1
      If IsMissing(PosFim) Then PosFim = Len(sBuffer) + 1
      
      If ViewFlood Then
         If IsMissing(LeftFlood) Then LeftFlood = .Left
         If IsMissing(TopFlood) Then TopFlood = .Width + (.Top / Screen.TwipsPerPixelY)
         If IsMissing(WidthFlood) Then WidthFlood = .Width / Screen.TwipsPerPixelX
         Set ProgBar.Parent = .Parent
         ProgBar.Create LeftFlood, TopFlood, WidthFlood, 15
         ProgBar.MinValue = PosIni
         ProgBar.MaxValue = PosFim
         'Call ProgBar.SetBarColor(QBColor(3))
      End If
      For nI = PosIni To PosFim
         If ViewFlood Then ProgBar.Value = nI
   
         Select Case Mid(sBuffer, nI, 1)
            Case "A" To "Z", "a" To "z", "_"
               If sTmpWord = "" Then nStartPos = nI
               sTmpWord = sTmpWord & Mid(sBuffer, nI, 1)
          
            Case Chr(34) '* Caracter " (Aspas)
               nSelLen = 1
               For nJ = 1 To 9999999
                  If Mid(sBuffer, nI + 1, 1) = Chr(34) Then
                     nI = nI + 2
                     Exit For
                  Else
                     nSelLen = nSelLen + 1
                     nI = nI + 1
                  End If
               Next
         
            Case Chr(39) '* Caracter ' (Plic)
               .SelStart = nI - 1
               nSelLen = 0
               lAux = Len(sBuffer) + 2 - nI
               For nJ = 1 To lAux
                  If Mid(sBuffer, nI, 2) = vbCrLf Then
                     Exit For
                  Else
                     nSelLen = nSelLen + 1
                     nI = nI + 1
                  End If
               Next
               .SelLength = nSelLen
               .SelColor = RGB(0, 127, 0)
         
            Case Else
               If Not (Len(sTmpWord) = 0) Then
                  '* verificar e corrigir pequenos erros
                  nWordPos = InStr(1, gsErrKeyWords, "*" & UCase(sTmpWord) & "*", 1)
                  If nWordPos <> 0 Then
                     LenWord = Len(sTmpWord)
                     Select Case UCase(sTmpWord)
                        Case "ENDIF":       sTmpWord = "End If"
                        Case "ENDFUNCTION": sTmpWord = "End Function"
                        Case "ENDSUB":      sTmpWord = "End Sub"
                        Case "SELECTCASE":  sTmpWord = "Select Else"
                        Case "CASEELSE":    sTmpWord = "Case Else"
                        'Case "ENDWITH":     sTmpWord = "End With"
                     End Select
                     sBuffer = Mid(sBuffer, 1, nStartPos - 1) & sTmpWord & Mid(sBuffer, nStartPos + LenWord)
                     .Text = sBuffer
                  End If
                  
                  .SelStart = nStartPos - 1
                  .SelLength = Len(sTmpWord)
                  
                  nWordPosBlack = InStr(1, gsBlackKeyWords, "*" & sTmpWord & "*", 1)
                  nWordPosBlue = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)
                  
                  If nWordPosBlack <> 0 Then
                     .SelColor = RGB(0, 0, 0)
                     .SelText = Mid(gsBlackKeyWords, nWordPosBlack + 1, Len(sTmpWord))
                  End If
                  If nWordPosBlue <> 0 Then
                     .SelColor = RGB(0, 0, 127)
                     .SelText = Mid(gsBlueKeyWords, nWordPosBlue + 1, Len(sTmpWord))
                  Else
                     If nWordPosBlack = 0 Then
                        .SelColor = RGB(0, 0, 0)
                        .SelText = sTmpWord
                     End If
                  End If
                  If UCase(sTmpWord) = "REM" Then
                     .SelStart = nI - 4
                     .SelLength = 3
                     For nJ = 1 To 9999999
                        If Mid(sBuffer, nI, 2) = vbCrLf Then
                           Exit For
                        Else
                           .SelLength = .SelLength + 1
                           nI = nI + 1
                        End If
                     Next
                     .SelColor = RGB(0, 127, 0)
                     .SelText = LCase(.SelText)
                  End If
               End If
               sTmpWord = ""
         End Select
      Next
      .SelStart = 0
      .SelStart = nLinBefore
      .Visible = True
      .Enabled = True
      .SetFocus
   End With
   If ViewFlood Then ProgBar.DestroyProgBar
   Set ProgBar = Nothing
End Sub
Public Sub ColorizeVbWords2(Rtf As Object, Optional PosIni, Optional PosFim, Optional ViewFlood = True, Optional LeftFlood, Optional TopFlood, Optional WidthFlood)
' **********************************************************************
' * Date            : 30/10/98
' * Time            : 14:47
' * Module Name     : Colorize_Module
' * Module Filename : Colorize.bas
' * Procedure Name  : ColorizeWords
' * Parameters      :  rtf As RichTextBox
' **********************************************************************
' * Comments        : Colorize in black, blue, green the VB keywords
' *
' *
' **********************************************************************
   
   Dim sBuffer    As String
   Dim nI         As Long
   Dim nJ         As Long
   Dim sTmpWord   As String
   Dim nStartPos  As Long
   Dim nSelLen    As Long
   Dim nWordPos   As Long
   Dim nWordPosBlack As Long
   Dim nWordPosBlue As Long
   Dim nLinBefore As Long
   Dim LenWord    As Integer
   Dim gsBlackKeyWords As String
   Dim gsBlueKeyWords  As String
   Dim gsErrKeyWords As String
   Dim lAux As Long
   
   Dim ProgBar As New CProgBar32, Value As Integer

   'Dim cHourglass  As class_Hourglass
   'Set cHourglass = New class_Hourglass
   gsBlackKeyWords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin"
   gsBlackKeyWords = gsBlackKeyWords & "*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear"
   gsBlackKeyWords = gsBlackKeyWords & "*Collection*Command*Cos*CreateObject*CurDir"
   gsBlackKeyWords = gsBlackKeyWords & "*DateAdd*DateDiff*DatePart*DateSerial*DateValue"
   gsBlackKeyWords = gsBlackKeyWords & "*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty"
   gsBlackKeyWords = gsBlackKeyWords & "*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime"
   gsBlackKeyWords = gsBlackKeyWords & "*FileLen*Fix*Format*FV*GetAllSettings*GetAttr"
   gsBlackKeyWords = gsBlackKeyWords & "*GetObject*GetSetting*Hex*Hide*Hour*InputBox"
   gsBlackKeyWords = gsBlackKeyWords & "*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty"
   gsBlackKeyWords = gsBlackKeyWords & "*IsError*IsMissing*IsNull*IsNumeric*IsObject"
   gsBlackKeyWords = gsBlackKeyWords & "*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log"
   gsBlackKeyWords = gsBlackKeyWords & "*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer"
   gsBlackKeyWords = gsBlackKeyWords & "*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate"
   gsBlackKeyWords = gsBlackKeyWords & "*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd"
   gsBlackKeyWords = gsBlackKeyWords & "*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn"
   gsBlackKeyWords = gsBlackKeyWords & "*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv"
   gsBlackKeyWords = gsBlackKeyWords & "*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial"
   gsBlackKeyWords = gsBlackKeyWords & "*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType"
   gsBlackKeyWords = gsBlackKeyWords & "*WeekDay*Width*Year*"
   
   
   gsBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary"
   gsBlueKeyWords = gsBlueKeyWords & "*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur"
   gsBlueKeyWords = gsBlueKeyWords & "*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const"
   gsBlueKeyWords = gsBlueKeyWords & "*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare"
   gsBlueKeyWords = gsBlueKeyWords & "*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec"
   gsBlueKeyWords = gsBlueKeyWords & "*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim"
   gsBlueKeyWords = gsBlueKeyWords & "*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase"
   gsBlueKeyWords = gsBlueKeyWords & "*Error*Event*Exit*Explicit*False*For*Function*Get"
   gsBlueKeyWords = gsBlueKeyWords & "*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer"
   gsBlueKeyWords = gsBlueKeyWords & "*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet"
   gsBlueKeyWords = gsBlueKeyWords & "*Name*New*Next*Not*Object*On*Open*Option*Or"
   gsBlueKeyWords = gsBlueKeyWords & "*Output*Print*Private*Property*Public*Put*Random"
   gsBlueKeyWords = gsBlueKeyWords & "*Read*ReDim*Resume*Return*RSet*Seek*Select*Set"
   gsBlueKeyWords = gsBlueKeyWords & "*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then"
   gsBlueKeyWords = gsBlueKeyWords & "*True*Type*UBound*Unlock*Variant*Wend*While*With"
   gsBlueKeyWords = gsBlueKeyWords & "*Xor*Nothing*To*"
   
   gsBlueKeyWords = gsBlueKeyWords & "*End If*End Function*End Sub"
   gsBlueKeyWords = gsBlueKeyWords & "*Select Case*Case Else*"
   
   gsErrKeyWords = "*ENDIF*ENDFUNCTION*ENDSUB*ENDWITH"
   gsErrKeyWords = gsErrKeyWords & "*SELECTCASE*CASEELSE*"
   
   sBuffer = Rtf.Text

   Dim StrLin0$, StrLin1$, strIniLine$, StrChar$, StrBlack$, StrBlue$, StrGreen$, StrRed$, StrBold$, StrStrike$, StrEnd$
   Dim StrTexto$
   Dim IniTexto As Boolean
   Dim IniLinha As Boolean
   
   StrLin0 = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fmodern Courier New;}{\f3\fmodern Courier New;}}" & vbNewLine
   StrLin0 = StrLin0 & "{\colortbl\red0\green0\blue0;\red0\green0\blue127;\red0\green127\blue0;\red255\green0\blue0;}"
   StrLin1 = "\deflang1046\pard"
   strIniLine = "\par "
   StrChar = "\plain\f3\fs20"
   StrBlack = "\cf0"
   StrBlue = "\cf1"
   StrGreen = "\cf2"
   StrRed = "\cf3"
   StrBold = "\b"
   StrStrike = "\strike"
   StrEnd = "\par }"

   
Rtf.Text = ""
Rtf.TextRTF = ""
StrTexto = StrLin0 & vbNewLine
StrTexto = StrTexto & StrLin1

'If False Then
'   StrTexto = StrTexto & StrChar & StrBlack & " Option" & StrChar & vbNewLine
'   StrTexto = StrTexto & strIniLine & StrChar & StrBlue & " Azul" & StrChar & vbNewLine
'   StrTexto = StrTexto & strIniLine & StrChar & StrGreen & " Verde" & StrChar & vbNewLine
'   StrTexto = StrTexto & strIniLine & StrChar & StrRed & StrBold & StrStrike & " Vermelho" & StrChar & vbNewLine
'   StrTexto = StrTexto & StrEnd
'   Rtf.TextRtf = StrTexto
'End If
'Exit Sub


   sTmpWord = ""
   With Rtf
      .Visible = False
      nLinBefore = .SelStart
      If IsMissing(PosIni) Then PosIni = 1
      If IsMissing(PosFim) Then PosFim = Len(sBuffer) + 1
      
      If ViewFlood Then
         If IsMissing(LeftFlood) Then LeftFlood = .Left
         If IsMissing(TopFlood) Then TopFlood = .Width + (.Top / Screen.TwipsPerPixelY)
         If IsMissing(WidthFlood) Then WidthFlood = .Width / Screen.TwipsPerPixelX
         Set ProgBar.Parent = .Parent
         ProgBar.Create LeftFlood, TopFlood, WidthFlood, 15
         ProgBar.MinValue = PosIni
         ProgBar.MaxValue = PosFim
         'Call ProgBar.SetBarColor(QBColor(3))
      End If
      IniTexto = True
      For nI = PosIni To PosFim
         If ViewFlood Then
            ProgBar.Value = nI
         End If
   
         Select Case Mid(sBuffer, nI, 1)
            Case "A" To "Z", "a" To "z", "_"
               If sTmpWord = "" Then nStartPos = nI
               sTmpWord = sTmpWord & Mid(sBuffer, nI, 1)
          
            Case Chr(34) '* Caracter " (Aspas)
               nSelLen = 1
               For nJ = 1 To 9999999
                  If Mid(sBuffer, nI + 1, 1) = Chr(34) Then
                     nI = nI + 2
                     Exit For
                  Else
                     nSelLen = nSelLen + 1
                     nI = nI + 1
                  End If
               Next
         
            Case Chr(39) '* Caracter ' (Plic)
               '.SelStart = nI - 1
               'nSelLen = 0
               lAux = Len(sBuffer) + 2 - nI
               For nJ = 1 To lAux
                  If Mid(sBuffer, nI, 2) = vbCrLf Then
                     Exit For
                  Else
                     sTmpWord = sTmpWord & Mid(sBuffer, nI, 1)
                     nSelLen = nSelLen + 1
                     nI = nI + 1
                  End If
               Next
               
               If IniTexto Then
                 StrTexto = StrTexto & StrChar
                 IniTexto = False
               Else
                 StrTexto = StrTexto & strIniLine & StrChar
               End If
               StrTexto = StrTexto & StrGreen & " " & sTmpWord & StrChar & " "
            Case Else
                If Mid(sBuffer, nI, 1) = Chr(13) Then
                    StrTexto = StrTexto & vbNewLine
                    IniLinha = True
                End If
               If Not (Len(sTmpWord) = 0) Then
                  '* verificar e corrigir pequenos erros
                  nWordPos = InStr(1, gsErrKeyWords, "*" & UCase(sTmpWord) & "*", 1)
                  If nWordPos <> 0 Then
                     LenWord = Len(sTmpWord)
                     Select Case UCase(sTmpWord)
                        Case "ENDIF":       sTmpWord = "End If"
                        Case "ENDFUNCTION": sTmpWord = "End Function"
                        Case "ENDSUB":      sTmpWord = "End Sub"
                        Case "SELECTCASE":  sTmpWord = "Select Else"
                        Case "CASEELSE":    sTmpWord = "Case Else"
                        'Case "ENDWITH":     sTmpWord = "End With"
                     End Select
                     sBuffer = Mid(sBuffer, 1, nStartPos - 1) & sTmpWord & Mid(sBuffer, nStartPos + LenWord)
                     '.Text = sBuffer
                  End If
                  
                  '.SelStart = nStartPos - 1
                  '.SelLength = Len(sTmpWord)
                  
                  nWordPosBlack = InStr(1, gsBlackKeyWords, "*" & sTmpWord & "*", 1)
                  nWordPosBlue = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)
                  
                  If IniTexto Then
                    StrTexto = StrTexto & StrChar
                    IniTexto = False
                  Else
                    If IniLinha Then
                       StrTexto = StrTexto & strIniLine
                       IniLinha = False
                    End If
                    StrTexto = StrTexto & StrChar
                  End If
                  
                  If nWordPosBlack <> 0 Then
                     StrTexto = StrTexto & StrBlack & " " & sTmpWord & StrChar & " "
                  ElseIf nWordPosBlue <> 0 Then
                     StrTexto = StrTexto & StrBlue & " " & sTmpWord & StrChar & " "
                  Else
                     If nWordPosBlack = 0 Then
                        StrTexto = StrTexto & StrBlack & " " & sTmpWord & StrChar & " "
                     End If
                  End If
                  If UCase(sTmpWord) = "REM" Then
                     .SelStart = nI - 4
                     .SelLength = 3
                     For nJ = 1 To 9999999
                        If Mid(sBuffer, nI, 2) = vbCrLf Then
                           Exit For
                        Else
                           .SelLength = .SelLength + 1
                           nI = nI + 1
                        End If
                     Next
                     .SelColor = RGB(0, 127, 0)
                     .SelText = LCase(.SelText)
                  End If
               End If
               sTmpWord = ""
         End Select
         'If Mid(sBuffer, nI, 1) = Chr(13) Then
         '   StrTexto = StrTexto & vbNewLine
         'End If
      Next
      .SelStart = 0
      .SelStart = nLinBefore
      .TextRTF = StrTexto & StrEnd
      .Visible = True
      .Enabled = True
      .SetFocus
   End With
   If ViewFlood Then ProgBar.DestroyProgBar
   Set ProgBar = Nothing
End Sub

Public Function Between(Vl, Min, Max)
Attribute Between.VB_Description = "Retorna ""True"" se o Valor informado estiver entr entre os valores Mínimo e Máximo ( Min e Max ) Inclusive."
   Between = False
   If Vl >= Min And Vl <= Max Then
      Between = True
   End If
End Function
Public Function DataUtil(Data As Date) As Boolean
   Dim bUtil   As Boolean
   
   bUtil = WeekDayUtil(Weekday(Data))
   If bUtil Then
      '* Pesquisar em Feridados Anuais (FERIADOS)
      '* Pesquisar em Recesso          (RECESSOEXP)
   End If
   
   DataUtil = bUtil
End Function
Private Function WeekDayUtil(Dia As Integer) As Boolean
   '* Pesquisar em Feridados Anuais
   Select Case Dia
      Case 1: WeekDayUtil = False   '* Domingo
      Case 2: WeekDayUtil = True    '* Segunda
      Case 3: WeekDayUtil = True    '* Terça
      Case 4: WeekDayUtil = True    '* Quarta
      Case 5: WeekDayUtil = True    '* Quinta
      Case 6: WeekDayUtil = True    '* Sexta
      Case 7: WeekDayUtil = False   '* Sábado
   End Select
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
Public Function Data_Padrao_Windows$(ByVal Dia%, ByVal Mes%, ByVal Ano%)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Recuperar o formato de data padrão que o windows**
'**            está utilizando.                                **
'** Recebe:  dia% - Dia atual.                                 **
'**          Mes% - Mês atual.                                 **
'**          Ano% - Ano atual.                                 **
'**                                                            **
'** Retorna: Formato de data do Sistema Operacional.           **
'**                                                            **
'****************************************************************
    Select Case mvarDtMask$
        Case "dd" + mvarSepDt + "MM" + mvarSepDt$ + "yyyy"
            Data_Padrao_Windows$ = Format$(Dia%, "00") + mvarSepDt$ + Format$(Mes%, "00") + mvarSepDt$ + Right$(CStr(Ano%), 4)
        Case "MM" + mvarSepDt + "dd" + mvarSepDt + "yyyy"
            Data_Padrao_Windows$ = Format$(Mes%, "00") + mvarSepDt + Format$(Dia%, "00") + mvarSepDt + Right$(CStr(Ano%), 4)
        Case "yyyy" + mvarSepDt + "MM" + mvarSepDt + "dd"
            Data_Padrao_Windows$ = Right$(CStr(Ano%), 4) + mvarSepDt + Format$(Mes%, "00") + mvarSepDt + Format$(Dia%, "00")
    End Select
End Function
Public Function GetDec(Num, Optional ByVal DC = 2)
'================================================================
'= Última Alteração : 05/07/99                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Retornar a parte decimal do número como Inteiro **
'**                                                            **
'** Recebe: Num     - Número com suas decimais                 **
'**         DC      - Número de Casas a serem retornadas       **
'**                                                            **
'** Retorna : Valor da casa decimal como Inteiro               **
'**                                                            **
'****************************************************************

   Dim Mult&, i%
   Mult = 1
   For i = 1 To DC
      Mult = 10 * Mult
   Next
   GetDec = (Num - Int(Num)) * Mult
End Function
'---------------------------------------------------------------
' DifTime()
' Retorna o tempo decorrido entre duas horas.
'---------------------------------------------------------------
Public Function DifTime(dInicio, dFim, Optional cFormato As Variant) As String
  Dim dIntervalo As Double
  Dim sD As Variant
  
  dIntervalo = TimeValue(dFim) - TimeValue(dInicio)
  sD = Int(CSng(dIntervalo))
'  DifTime = IIf(sD = 0, "", _
'    IIf(sD = 1, sD & " dia, ", sD & " dia, ")) & _
'    Format$(dIntervalo, "hh:mm:ss")
  DifTime = IIf(sD = 0, "", sD & " dia, ") & _
    Format$(dIntervalo, IIf(Vazio(cFormato), "hh:mm:ss", cFormato))
  
End Function
Public Sub DimensionarTela(MDI As Object, frm As Object)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Dimensionar um Form, com MDI_child = True,      **
'**            dentro do espaço livre do SysMdi          **
'** Recebe:  form_name - Form a ser dimensionado               **
'**                                                            **
'** Retorna: Form dimensionado                                 **
'**                                                            **
'****************************************************************
    If VarType(frm) = 9 Then
       frm.Height = MDI.ScaleHeight
       frm.Width = MDI.ScaleWidth
       frm.Top = 0 '30
       frm.Left = 0 '30
    Else
       frm.Height = MDI.ScaleHeight - 50
       frm.Width = MDI.ScaleWidth - 50
       frm.Top = 30
       frm.Left = 30
    End If
   ' Frm.Height = 4900
   ' Frm.Width = 9550
End Sub
Public Function SetFormatDT_Number() As Boolean
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Definir formato de data e número                **
'**                                                            **
'** Recebe:                                                    **
'**                                                            **
'** Retorna: Mensagem com os formatos definidos se o sistema   **
'**          operacional estiver utilizando formatos diferntes **
'**                                                            **
'****************************************************************

'* Função deve ser revisada para uma utilização generalizada

   Dim i%, Txt$, Aux$, aux1$, DT$, tmp_data$, tmp_ano%
   Dim Dia$, Mes$, Ano$
   
   On Error Resume Next
   
   Aux$ = Format$(1000, "#,##0.00")
   mvarSepDec$ = Mid$(Aux$, 6, 1)
   mvarSepMil$ = Mid$(Aux$, 2, 1)
   mvarSepDt$ = "/"
   Aux$ = CStr(Date)
   For i% = 2 To 5
      aux1$ = Mid$(Aux$, i%, 1)
      'procura o primeiro caracter que não seja um dígito
      If aux1 < "0" Or aux1 > "9" Then
         mvarSepDt$ = aux1$
         Exit For
      End If
   Next
   
   Dia = Format$(Day(Aux$), "00")
   Mes = Format$(Month(Aux$), "00")
   Select Case Len(Aux$)
      Case 8: Ano = Right$(Year(Now), 2)
      Case 10: Ano = Year(Aux$)
   End Select
   Select Case Aux$
      Case Dia + mvarSepDt$ + Mes + mvarSepDt$ + Ano
         mvarFormatoData$ = "DMA"
         mvarDtMask$ = "dd" + mvarSepDt$ + "MM" + mvarSepDt$ + "yyyy"
         mvarDtMaskAux$ = "dd" + mvarSepDt$ + "MM" + mvarSepDt$ + "yy"
      Case Mes + mvarSepDt$ + Dia + mvarSepDt$ + Ano
         mvarFormatoData$ = "MDA"
         mvarDtMask$ = "MM" + mvarSepDt$ + "dd" + mvarSepDt$ + "yyyy"
         mvarDtMaskAux$ = "MM" + mvarSepDt$ + "dd" + mvarSepDt$ + "yy"
      Case Ano + mvarSepDt$ + Mes + mvarSepDt$ + Dia
         mvarFormatoData$ = "AMD"
         mvarDtMask$ = "yyyy" + mvarSepDt$ + "MM" + mvarSepDt$ + "dd"
         mvarDtMaskAux$ = "yy" + mvarSepDt$ + "MM" + mvarSepDt$ + "dd"
   End Select
   DT$ = Format$(Now, mvarDtMask$)
   
   'testa formato data/número
   If mvarDtMask$ <> "dd" + mvarSepDt$ + "MM" + mvarSepDt$ + "yyyy" Then
      Txt$ = ClsMsg.LoadMsg(17) + Chr$(10) + Chr$(10)
      Txt$ = Txt$ + ClsMsg.LoadMsg(19)
      GoTo Fim
   End If
   If Not (mvarSepDt$ = "/" And mvarSepMil$ = "." And mvarSepDec$ = ",") Then
   ' And GetStringFromIni("Intl", "sShortDate", "win.ini") = UCase(mvarDtMask$)) Then
   'If Not (mvarSepDt$ = "/" And mvarSepMil$ = "," And mvarSepDec$ = ".") Then
      If Not (mvarSepDt$ = "." And mvarSepMil$ = "," And mvarSepDec$ = ".") Then
      Txt$ = ClsMsg.LoadMsg(18) + Chr$(10) + Chr$(10)
      Txt$ = Txt$ + ClsMsg.LoadMsg(19) + Chr$(10)
      Txt$ = Txt$ + ClsMsg.LoadMsg(20) + Chr$(10)
      Txt$ = Txt$ + vbNewLine
      Txt$ = Txt & "Sep. Dt : " & mvarSepDt$ + vbNewLine
      Txt$ = Txt & "Sep. Grp: " & mvarSepMil$ + vbNewLine
      Txt$ = Txt & "Sep. Dec: " & mvarSepDec$ + vbNewLine
      Txt$ = Txt & "Aux : " & Format$(1000, "#,##0.00") + vbNewLine
      
      
'      Txt$ = Txt$ + "Número, utilize 9,999.99" + Chr$(10)
       GoTo Fim
       End If
   End If
   Call WritePrivateProfileString("INTL", "SSHORTDATE", mvarDtMask, "WIN.INI")
   Call WritePrivateProfileString("INTL", "sDate", mvarSepDt, "WIN.INI")
   Call WritePrivateProfileString("INTL", "sThousand", mvarSepMil, "WIN.INI")
   Call WritePrivateProfileString("INTL", "sDecimal", mvarSepDec, "WIN.INI")

    'testa formato ano da data
    tmp_ano% = 0
    tmp_data$ = Trim$(CStr(CVDate(DT$)))
    For i% = Len(tmp_data$) To 1 Step -1
       If Mid$(tmp_data$, i%, 1) = mvarSepDt$ Then
          tmp_ano% = Val(Trim$(Mid$(tmp_data, i% + 1)))
          Exit For
       End If
    Next i%
    SetFormatDT_Number = True
Exit Function

Fim:
   Screen.MousePointer = vbDefault
   MsgBox Txt$, vbCritical
   DoEvents
   SetFormatDT_Number = False
End Function
Public Function DToMask$(Data As Variant, ByRef Msk As Object)
   Dim Ano As Variant
   Msk.PromptInclude = False
   If IsNull(Data) Or Vazio(Data) Then Data = Format$("31/12/1899", mvarDtMask$)
   On Error Resume Next
   If Year(Data) = 1899 Then
      Call ZerarMask(Msk, Msk.Mask)
   Else
      Msk = Format$(Data, mvarDtMask$)
      If Len(Msk.Mask) < 9 Then
         Msk = Format$(Data, mvarDtMaskAux)
      End If
   End If
   Msk.PromptInclude = True
   DToMask$ = Msk.Text
End Function
Public Function Change_Data(ByRef Campo As Variant)
   Dim strIniData, Ano$, SECULO$, ValorAno%
   Dim TamData%, Bool%
   Dim strData$, dd$, mm$, yy$, Pos%
   Dim mvarDtMaskAux$, mvarDtMask$
   TamData = Len(Trim$(Campo))
   strData = Campo
   If TamData > 0 Then
      strIniData = Left$(Trim$(Campo), 6)
      ValorAno% = Val(Right$(Trim$(Campo), TamData - 6))
      Ano = Right$(Format$(ValorAno%, "0000"), 2)
      SECULO$ = IIf(ValorAno > 1000, Left$(ValorAno, 2), Left$(Year(Now), 2))
      Change_Data = strIniData + SECULO$ + Ano$
   Else
      Change_Data = Campo
   End If
   Select Case UCase(TypeName(Campo))
      Case "MASKEDBOX"
         Select Case Campo.Mask
            Case "##/##/##": Change_Data = Format$(Change_Data, mvarDtMaskAux$)
            Case "##/##/####": Change_Data = Format$(Change_Data, mvarDtMask$)
            Case Else: Change_Data = Format$(Change_Data, Campo.Format)
         End Select
   End Select
End Function
Public Function ComputerName() As String
   Dim lpBuffer As String * 255
   Dim lRet As Long
   Dim i As Integer
   
   lRet = GetComputerName(lpBuffer, 255)
   If lRet <> 0 Then
     ComputerName = UCase$(Left$(Trim$(lpBuffer), Len(Trim$(lpBuffer)) - 1))
     ComputerName = EliminarString(ComputerName, Chr(0))
   Else
     ComputerName = "Não Fornecido."
   End If
End Function
Public Function CTOD$(ByVal DT$)
   If Val(DT$) = 0 Then
      CTOD$ = "#" + str(CDate("00:00:00")) + "#"
   Else
      CTOD$ = "'" + Format$(DT$, mvarDtMask$) + "'"
   End If
End Function
Public Function CTON(ByVal Objeto As String) As Currency
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Converter um texto em número do tipo Currency   **
'**            trocando o separador de decimais de "," p/ "."  **
'**                                                            **
'** Recebe: objeto - O objetom que contém o texto              **
'**                                                            **
'** Retorna : Número de Tipo Currency                          **
'**                                                            **
'****************************************************************
Dim StN_aux$, StN_x%
    
    StN_aux$ = Objeto
    StN_aux$ = Tiraponto(StN_aux)
    
    '*** troca "," por "."
    StN_x% = InStr(StN_aux, ",")
    If StN_x% > 0 Then
       Mid(StN_aux, StN_x%, 1) = "."
    End If

    CTON@ = Val(StN_aux)

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
Public Sub ExecuteLink(ByVal sLinkTo As String)

'* Execute o link to http://www.DSR.com.br/
'* (if possible) - or the new 'mailto:dramos@mandic.com.br'

    On Error Resume Next
    
    Dim lRet As Long
    Dim lOldCursor As Long
    
    lOldCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    lRet = ShellExecute(0, "open", sLinkTo, "", vbNull, SW_SHOWNORMAL)
    
    If lRet >= 0 And lRet <= 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error Opening Link to " & sLinkTo & vbCrLf & vbCrLf & Err.LastDllError, , "frmAbout::ExecuteLink"
    End If
    Screen.MousePointer = vbDefault
    
End Sub
Public Function ExisteItem(MyColl As Collection, Item As String) As Boolean
   Dim x  As Variant
   Err = 0
   On Error Resume Next
   x = MyColl(Item)
   If Err <> 0 Then
      Err = 0
      Set x = MyColl(Item)
   End If
   If (Err = 0) Then
      ExisteItem = (MyColl.Count > 0)
   Else
      ExisteItem = (Err = 0)
   End If
End Function
Public Sub FloodAtu(ByRef Progress As Object, Vl%, Optional ByRef Lbl = "", Optional IniFim = True)
   Dim i%
   Const Mult% = 10
   If Lbl <> "" Then Lbl.Visible = True
   With Progress
      If IniFim And Not .Visible Then Call FloodIni(Progress)
      For i = (.Value * Mult) To (Vl * Mult)
         .Value = IIf(CInt(i / Mult) > .Max, .Max, CInt(i / Mult))
         If Lbl <> "" Then
            Lbl.Visible = True
            Lbl.Caption = CStr(.Value) & "%"
            Lbl.Refresh
         End If
      Next
      If Vl >= .Max Then
         If IniFim Then
            Call FloodFim(Progress)
         Else
            Call FloodAtu(Progress, 0, , False)
         End If
      End If
   End With
End Sub
Public Sub FloodFim(ByRef Progress As Object)
   With Progress
      .Visible = False
   End With
End Sub
Public Sub FloodIni(ByRef Progress As Object)
   With Progress
      .Min = 0
      .Max = 100
      .Value = 0
      .Appearance = 1
      .Visible = True
   End With
End Sub
Public Function FormatDate(pDate As String, Optional ComHora = True) As String
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
   
   If mvarFormatoData = "DMA" Then
      sDate = StrZero(Day(dDate), 2) & "/" & StrZero(Month(dDate), 2) & "/" & Year(dDate)
   ElseIf mvarFormatoData = "MDA" Then
      sDate = StrZero(Month(dDate), 2) & "/" & StrZero(Day(dDate), 2) & "/" & Year(dDate)
   ElseIf mvarFormatoData = "AMD" Then
      sDate = Year(dDate) & "/" & StrZero(Month(dDate), 2) & "/" & StrZero(Day(dDate), 2)
   Else
      sDate = StrZero(Day(dDate), 2) & "/" & StrZero(Month(dDate), 2) & "/" & Year(dDate)
   End If
   
   If xVal(Hour(pDate)) <> 0 And ComHora Then
      sDate = sDate & " " & StrZero(Hour(pDate), 2) & ":" & StrZero(Minute(pDate), 2) & ":" & StrZero(Second(pDate), 2)
   End If
   
   FormatDate = sDate
Exit Function
TrataErro:
   FormatDate = sDate
   Call ClsMsg.ExibirStop("Formato de Data Inválido", "Atenção")
End Function
Public Function FormatarCNPJ(pNum As String, Optional isPJ As Boolean = True)
   Dim sTexto As String
   If isPJ Then
      sTexto = Mid(pNum, 1, 2)
      sTexto = sTexto & "." & Mid(pNum, 3, 3)
      sTexto = sTexto & "." & Mid(pNum, 6, 3)
      sTexto = sTexto & "/" & Mid(pNum, 9, 4)
      sTexto = sTexto & "-" & Mid(pNum, 13, 2)
   Else
      sTexto = Mid(pNum, 1, 3)
      sTexto = sTexto & "." & Mid(pNum, 4, 3)
      sTexto = sTexto & "." & Mid(pNum, 7, 3)
      sTexto = sTexto & "-" & Mid(pNum, 10, 2)
   End If
   FormatarCNPJ = sTexto
End Function
Public Function FormatarMoeda(Campo$, Qtd_Int%, Qtd_Dec%) As String
   Dim Num$, Formato$, Achou$, qtCasaDecimal$, Max$, Numero$, DecZero$
   Dim i%, NumMax#, TamMaxNum%, TamMaxCarac%, Qtd_Ponto%
   
   Qtd_Int = IIf(Qtd_Int = 0, 1, Qtd_Int)
   Qtd_Ponto = Int((Qtd_Int - 1) / 3)
   Qtd_Ponto = IIf(Qtd_Ponto < 0, 0, Qtd_Ponto)
   TamMaxNum% = Qtd_Int - 1
   TamMaxCarac% = TamMaxNum% + Qtd_Dec + 1
   DecZero$ = ""
   For i = 1 To Qtd_Dec
      DecZero$ = DecZero$ + "0"
   Next
   Max = "9"
   For i = 1 To TamMaxNum%
      Max = Max + "9"
   Next
     
   NumMax# = CDbl(Max)
   Formato = "#,###,##0." + DecZero$
   Numero = Format$(Campo, Formato)
   Achou = False
   i = InStr(Numero, mvarSepDec$)
   If InStr(Numero, ".") > 0 Then
      Achou = True
   End If
   qtCasaDecimal = Len(Trim$(Numero)) - i
   
   If Achou Then
      Select Case i:
         Case 1: Numero = "0" + mvarSepDec$
      End Select
   Else
      If Len(Trim$(Numero)) = 0 Then
         Num = "0" + mvarSepDec$ + DecZero$
      Else
         Num = Numero + mvarSepDec$ + DecZero$
     End If
   End If
   Num = Numero + Mid(DecZero$, 1, Qtd_Dec - qtCasaDecimal)

   If CDbl(Num) >= NumMax Then
      Beep
      FormatarMoeda = Numero
      Exit Function
   End If

   FormatarMoeda$ = Num
End Function
Public Function GetDecimal(Num, Optional ByVal QtdCasas = 2)
   Dim Mult&, i%
   Mult = 1
   For i = 1 To QtdCasas
      Mult = 10 * Mult
   Next
   GetDecimal = (Num - Int(Num)) * Mult
End Function
Public Function GetNameFromPath(PathFile As String, Optional PathNome = 2) As String
'================================================================
'= Última Alteração : 03/02/2000                                =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Recupera o nome do arquivo da "String de seu    **
'**            "Path".                                         **
'**                                                            **
'** Recebe: PathFile - Caminho Completo com nome do Arquivo.   **
'**                                                            **
'** Retorna : Nome do Arquivo.                                 **
'**                                                            **
'****************************************************************
   Dim i%
   While InStr(i + 1, PathFile, "\")
      i = InStr(i + 1, PathFile, "\")
   Wend
   If PathNome = 1 Then
      GetNameFromPath = VBA.Left$(PathFile, i)
   ElseIf PathNome = 2 Then
      GetNameFromPath = VBA.Mid$(PathFile, Len(VBA.Left$(PathFile, i)) + 1)
   End If
End Function
Public Function ResolvePathName(ByVal sPath As String) As String
   Dim PosIni As Integer
   Dim PosFim As Integer

   If Right(sPath, 1) <> "\" And Trim(sPath) <> "" Then
      sPath = sPath & "\"
   End If
   If InStr(sPath, "%") <> 0 Then
      PosIni = InStr(sPath, "%")
      PosFim = InStr(PosIni + 1, sPath, "%")
      sPath = Mid(sPath, 1, PosIni - 1) & Environ(Mid(sPath, PosIni + 1, PosFim - PosIni - 1)) & Mid(sPath, PosFim + 1)
   End If

   ResolvePathName = sPath
End Function
Public Function GetTokens(ByVal pString As String, Optional pDelim As String = " ") As Collection
   Dim MyColl  As Collection
   Dim sWord    As String

   
   Set MyColl = New Collection
   
   pString = Trim(pString)
   While Trim(pString) <> ""
      sWord = Mid(pString, 1, IIf(InStr(pString, pDelim) > 0, InStr(pString, pDelim) - 1, Len(pString)))
      If Trim(sWord) <> "" Then
         MyColl.Add sWord
      End If
      pString = Trim(Mid(pString, Len(sWord) + 1))
   Wend
   Set GetTokens = New Collection
   Set GetTokens = MyColl
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
      Call ClsMsg.ExibirAviso("Consulta Inválida.")
      Exit Function
   End If
   
   sAux = UCase(StrReplace(pTexto, vbNewLine, " "))
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
   sAux = UCase(StrReplace(pTexto, vbNewLine, " "))
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
Public Function GetPalavras(ByVal StrLinha As String) As Collection
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
   Set GetPalavras = New Collection
   Set GetPalavras = MyColl
End Function
Public Function GetPalavras_AndOR(pTexto As String, Optional ByRef Palavras_And As Collection, Optional ByRef Palavras_Or As Collection, Optional pCampo) As String
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
      Call ClsMsg.ExibirAviso("Consulta Inválida.")
      Exit Function
   End If
   
   sAux = UCase(StrReplace(pTexto, vbNewLine, " "))
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
   sAux = UCase(StrReplace(pTexto, vbNewLine, " "))
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
   GetPalavras_AndOR = Sql
End Function
Public Function RichWordOver(ByVal RchTxt As Variant, x As Single, y As Single, Optional Posicao = 1, Optional VerifDclImplicta = True) As String
   Dim pt As PointAPI
   Dim Pos As Integer
   Dim Ch As String
   Dim StartPos As Integer, EndPos As Integer
   Dim Txt As String, TxtLen As Integer
   
   Dim LineCount As Single, CurrLinePos As Single, OverAllCursorPos As Single
   Dim ChrsBeforeCurrLine As Single, CurrLineLen As Single, CurrLineCursorPos As Single
   

   ' Convert the position to pixels.
   
   ' Get the character number
   If x = 0 And y = 0 Then
      If IsMissing(Posicao) Then
         LineCount = SendMessageLong(RchTxt.hWnd, EM_GETLINECOUNT, 0, 0&)
         OverAllCursorPos = SendMessageLong(RchTxt.hWnd, EM_GETSEL, 0, 0&) \ &H10000
         CurrLinePos = SendMessageLong(RchTxt.hWnd, EM_LINEFROMCHAR, OverAllCursorPos, 0&)
         ChrsBeforeCurrLine = SendMessageLong(RchTxt.hWnd, EM_LINEINDEX, CurrLinePos, 0&)
         CurrLineLen = SendMessageLong(RchTxt.hWnd, EM_LINELENGTH, -1, 0&)
         CurrLineCursorPos = OverAllCursorPos + 1 - ChrsBeforeCurrLine
         Pos = OverAllCursorPos - 2 * (CurrLineCursorPos)
         If Pos < 0 Then Exit Function
      Else
         Pos = Posicao
      End If
   
   Else
      pt.x = x \ Screen.TwipsPerPixelX
      pt.y = y \ Screen.TwipsPerPixelY
      Pos = SendMessage(RchTxt.hWnd, EM_CHARFROMPOS, 0&, pt)
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
       Ch = Mid$(Txt, StartPos, 1)
       ' Allow digits, letters, and underscores.
       If Not isAlfaNum(Ch) Then
           Exit For
       End If
   Next StartPos
   StartPos = StartPos + 1

   ' Find the end of the word.
   Dim Dcl_Implicita
   TxtLen = Len(Txt)
   For EndPos = Pos To TxtLen
       Ch = Mid$(Txt, EndPos, 1)
       ' Allow digits, letters, and underscores.
       If Not (isAlfaNum(Ch) Or Ch = "_") Then
          If VerifDclImplicta Then
             If Not VarWithDclImplicit(Ch) Then
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
Public Function isAlfaNum(Ch As String) As Boolean
   isAlfaNum = (Ch >= "0" And Ch <= "9") Or (UCase(Ch) >= "A" And UCase(Ch) <= "Z")
   isAlfaNum = isAlfaNum Or (UCase(Ch) >= "À" And UCase(Ch) <= "Ý")    '* Or (Ch = "_")
End Function
Public Function VarWithDclImplicit(Variable As String) As Boolean
  Dim Ch As String
  Ch = Right(Trim(Variable), 1)
  VarWithDclImplicit = (Ch = "%" Or Ch = "!" Or Ch = "&" Or Ch = "$" Or Ch = "#")
End Function
Public Function NameDclImplicit(VarName As String) As String
   Dim NmDcl As String
   Select Case Right(VarName, 1)
      Case "!":      NmDcl = "Single"
      Case "#":      NmDcl = "Double"
      Case "$":      NmDcl = "String"
      Case "%":      NmDcl = "Integer"
      Case "&":      NmDcl = "Long"
      Case "(", ")": NmDcl = "Array"
      Case Else:     NmDcl = "Variant"
   End Select
   NameDclImplicit = NmDcl
End Function
Public Function RichMousePos(RtfHWnd As Long, x As Single, y As Single) As Long
   Dim pt As PointAPI
   pt.x = x \ Screen.TwipsPerPixelX
   pt.y = y \ Screen.TwipsPerPixelY
   RichMousePos = SendMessage(RtfHWnd, EM_CHARFROMPOS, 0&, pt)
End Function
Public Function GetStringFromIni$(lpAppName$, lpKeyName$, Optional lpFilename)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Ler uma string do arquivo .INI, o nome  de  tal **
'**            arquivo deve estar na variável SYS_INI definida **
'**            na declaração de variáveies do Sistema.         **
'**                                                            **
'** Recebe: LpaPPnAME - Nome da classe da string.              **
'**         lpKeyName - Nome da string a ser lida              **
'**                                                            **
'** Retorna : Valor da string                                  **
'**                                                            **
'****************************************************************
    Dim valid As Integer
    Dim lpDefault As String
    Dim lpreturnstring As String
    Dim SIZE As Long
    
    lpDefault = ""
    lpreturnstring = Space$(255)
    SIZE = Len(lpreturnstring)
    
    valid = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpreturnstring$, SIZE, lpFilename)
    GetStringFromIni = UCase$(Left$(lpreturnstring, valid))
End Function
Function GetPathName(ByVal strFilename As String) As String
    Dim intPos As Integer
    Dim strPathOnly As String
    Dim dirTmp As DirListBox
    Dim i As Integer

    On Error Resume Next
    Err = 0

    intPos = Len(strFilename)

    '
    'Change all '/' chars to '\'
    '

    For i = 1 To Len(strFilename)
        If Mid$(strFilename, i, 1) = gstrSEP_DIRALT Then
            Mid$(strFilename, i, 1) = gstrSEP_DIR
        End If
    Next i

    If InStr(strFilename, gstrSEP_DIR) = intPos Then
        If intPos > 1 Then
            intPos = intPos - 1
        End If
    Else
        Do While intPos > 0
            If Mid$(strFilename, intPos, 1) <> gstrSEP_DIR Then
                intPos = intPos - 1
            Else
                Exit Do
            End If
        Loop
    End If

    If intPos > 0 Then
        strPathOnly = VBA.Left$(strFilename, intPos)
        If Right$(strPathOnly, 1) = gstrCOLON Then
            strPathOnly = strPathOnly & gstrSEP_DIR
        End If
    Else
        strPathOnly = CurDir$
    End If

    If Right$(strPathOnly, 1) = gstrSEP_DIR Then
        strPathOnly = VBA.Left$(strPathOnly, Len(strPathOnly) - 1)
    End If

    GetPathName = UCase16(strPathOnly)

    Err = 0
End Function
Function UCase16(ByVal str As String)
    UCase16 = str
End Function
Public Sub HideTaskbar()
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd <> 0 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    End If
End Sub
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
Public Function LeVetBarra(VetorBarra$, Optional Var = "") As Variant
   Dim Pos%, Fim%
   If Trim(VetorBarra$) = "" Then
      LeVetBarra = ""
      Exit Function
   End If
   If Var = "" Then
   Else
      Pos = InStr(VetorBarra$, "|" & UCase(Var) & "=")
      If Pos > 0 Then
         Fim = InStr(Pos + Len(Var) - 1, VetorBarra$, "|")
         Fim = IIf(Fim = 0, Len(VetorBarra$), Fim)
         LeVetBarra = Mid(VetorBarra$, Pos + 11, Fim)
      End If
   End If
End Function
Public Sub HabilitarTela(frm As Object, Optional TipoHabilita As eHabilitarTela = EnabledDefault)
   Dim i%, GravaDefault As Boolean
   
   On Error Resume Next
   If TipoHabilita = EnabledDefault Then
      Select Case UCase(ClsCtrl.GetTag(frm, "HABILITARTELA"))
         Case "S", ""
            Call ClsCtrl.SetTag(frm, "HABILITARTELA", "N")
            GravaDefault = True
         Case Else
            GravaDefault = False
      End Select
      For i% = 0 To frm.Controls.Count - 1
         If TypeName(frm.Controls(i)) <> "Line" And TypeName(frm.Controls(i)) <> "Shape" Then
            If GravaDefault Then
               Call ClsCtrl.SetTag(frm.Controls(i), "ENABLEDDEFAULT", IIf(frm.Controls(i).Enabled, "S", "N"))
            Else
               frm.Controls(i).Enabled = (ClsCtrl.GetTag(frm.Controls(i), "ENABLEDDEFAULT") = "S")
            End If
         End If
      Next
   Else
      For i% = 0 To frm.Controls.Count - 1
         If TypeName(frm.Controls(i)) <> "Line" And TypeName(frm.Controls(i)) <> "Shape" Then
            frm.Controls(i).Enabled = (TipoHabilita = EnabledTrue)
         End If
      Next
   End If
End Sub
Public Sub LimparTela(frm As Object)
   Dim i%
   On Error Resume Next
   
   For i% = 0 To frm.Controls.Count - 1
      
      Select Case UCase(TypeName(frm.Controls(i)))
         
         Case "TEXTBOX"
            frm.Controls(i) = ""
         
         Case "MASKEDBOX"
            Call ZerarMask(frm.Controls(i))
         
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
Public Function OptValue(OptControl As Object) As Integer
   Dim i As Integer
   OptValue = -1
   For i = OptControl.LBound To OptControl.UBound
      If OptControl(i).Value Then
         OptValue = i
         Exit For
      End If
   Next
End Function
Public Sub Pause(ByVal Seconds As Single)
   Call Sleep(Int(Seconds * 1000#))
End Sub
Public Function PrnTxt(ByVal x%, ByVal y%, ByVal Txt$, ByVal FonteTit$, ByVal TamFontTit%, ByVal NegritoTit%, ParamArray Param())
'****************************************************************************
'* Função  : Rotina auxiliar para impressão de texto.                       *
'*                                                                          *
'* Entrada : x% = POSIÇÃO HORIZONTAL (PROP. LEFT) DO TEXTO DE IDENTIFICAÇÃO *
'*           y% = POSIÇÃO VERTICAL (PROP. TOP) DO TEXTO DE IDENTIFICAÇÃO    *
'*           txt$ = TEXTO DE IDENTIFICAÇÃO                                  *
'*           fontetit$                                                      *
'*           tamfonttit%                                                    *
'*           negritotit%                                                    *
'****************************************************************************
   Dim TamanhoTit%, TotX%
   Dim Line%, Alignment%
   
   On Error Resume Next

   Line% = IIf(UBound(Param) >= 0, Param(0), True)
   Alignment% = IIf(UBound(Param) >= 1, Param(1), vbLeftJustify)
   Printer = IIf(UBound(Param) >= 2, Param(2), Printer)

'* Fonte do Texto
   Printer.FontName = FonteTit$
   Printer.FontBold = NegritoTit%
   Printer.FontSize = TamFontTit%
'* Posição do Texto
   TotX% = Printer.Width - 400
   TamanhoTit% = Printer.TextWidth(Txt$)
   If Line Then x% = 100
   Select Case Alignment
      Case vbLeftJustify: Printer.CurrentX = x%
      Case vbRightJustify: Printer.CurrentX = (TotX% - TamanhoTit%)
      Case vbCenter: Printer.CurrentX = (((TotX% - x%) - TamanhoTit%) / 2) + 100
   End Select
   Printer.CurrentY = y%
'* Imprimir Texto
   If Txt$ <> "" Then
      Printer.Print Txt$
   End If
'* Posição Final ddo Texto
   Select Case Alignment
      Case vbLeftJustify: Printer.CurrentX = TamanhoTit% + x%
      Case vbRightJustify: Printer.CurrentX = (TotX% - x%)
      Case vbCenter: Printer.CurrentX = (((TotX% - x%) - TamanhoTit%) / 2) + TamanhoTit%
   End Select
   If Not Line% Then
      Printer.CurrentY = y%
   Else
      Printer.CurrentX = 100
   End If
   PrnTxt = Array(Printer.CurrentX, Printer.CurrentY)
End Function
Public Function RegServer(sServerPath As String, Optional fRegister = True, Optional fMsg As Boolean = True, Optional isActivexExe As Boolean = False) As Boolean
  
   Dim hMod As Long               ' module handle
   Dim lpfn As Long                  ' reg/unreg function address
   Dim sCmd As String             ' msgbox string
   Dim lpThreadID As Long        ' unused, receives the thread ID
   Dim hThread As Long            ' thread handle
   Dim fSuccess As Boolean     ' if things worked
   Dim dwExitCode As Long      ' thread's exit code if it doesn't finish
   
   ' Load the server into memory
   hMod = LoadLibrary(sServerPath)
   
   ' Get the specified function's address and our msgbox string.
   If fRegister Then
      If isActivexExe Then
         lpfn = GetProcAddress(hMod, "ExeRegisterServer")
      Else
         lpfn = GetProcAddress(hMod, "DllRegisterServer")
      End If
     sCmd = "register"
   Else
      If isActivexExe Then
         lpfn = GetProcAddress(hMod, "ExeUnregisterServer")
      Else
         lpfn = GetProcAddress(hMod, "DllUnregisterServer")
      End If
      sCmd = "unregister"
   End If
   
   ' If we got a function address...
   If lpfn Then
     
     ' Create an alive thread and execute the function.
     hThread = CreateThread(ByVal 0, 0, ByVal lpfn, ByVal 0, 0, lpThreadID)
     
     ' If we got the thread handle...
     If hThread Then
       
       ' Wait 10 secs for the thread to finish (the function may take a while...)
       fSuccess = (WaitForSingleObject(hThread, 10000) = WAIT_OBJECT_0)
       
       ' If it didn't finish in 10 seconds...
       If Not fSuccess Then
         ' Something unlikely happened, lose the thread.
         Call GetExitCodeThread(hThread, dwExitCode)
         Call ExitThread(dwExitCode)
       End If
       
       ' Lose the thread handle
       Call CloseHandle(hThread)
     
     End If   ' hThread
   End If   ' lpfn
   
   ' Free the server if we loaded it.
   If hMod Then Call FreeLibrary(hMod)
   
   If fMsg Then
      If fSuccess Then
        MsgBox "Successfully " & sCmd & "ed " & sServerPath   ' past tense
        RegServer = True
      Else
        MsgBox "Failed To " & sCmd & " " & sServerPath, vbExclamation
      End If
   End If
End Function
Public Sub WriteLabel(ByVal Obj As Object, ByVal Txt As String, ByVal x As Single, ByVal y As Single, Optional FontePadrão = True)
   Dim One_Line As String
   Dim Pos As Integer

   Obj.CurrentY = y
   If FontePadrão Then
      Obj.Font.Name = "Times New Roman"
      Obj.Font.SIZE = 10
      Obj.Font.Bold = True
   End If
   Do While Len(Txt) > 0
      '* Achar Próxima Linha.
      Pos = InStr(Txt, vbCrLf)
      If Pos = 0 Then
          '* Última Linha.
          One_Line = Txt
          Txt = ""
      Else
          One_Line = Left$(Txt, Pos - 1)
          Txt = Right$(Txt, _
              Len(Txt) - Pos + 1 - Len(vbCrLf))
      End If
      '* Escrever Label e ulr Linha
      Obj.CurrentX = x
      Obj.Print One_Line
   Loop
End Sub
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
         'Case "OPTIONBUTTON", "MASKEDBOX", "COMBOBOX", "CHECKBOX", "SSOPTION", "SSCHECK"
         Case Else
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
      Call ClsCtrl.SetTag(frm, "SUJA", "True")
   End If
   
   Exit Function
Fim:
   Call ClsMsg.ShowError("DSR100.DSR.SendTab")
End Function
Public Sub SetDefault(hWnd As Long)
  DoEvents
  Call ReleaseCapture
  Call SetCursor(LoadCursor(0, IDC_ARROW))
  Screen.MousePointer = vbDefault
End Sub
Public Sub SetHourglass(hWnd As Long)
  DoEvents
  Call SetCapture(hWnd)
  Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
  Screen.MousePointer = vbHourglass
End Sub
Public Sub ShowTaskbar()
    If TaskBarhWnd <> 0 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End If
End Sub
Public Function String_Sem_Acento(str) As String
'ALTERAR FUNÇÃO PARA SUPORTAR CARACTERES MINUSCULOS
'só para maiúsculas

Dim strNova, Car$
Dim Tam%, i%

Tam = Len(str)
strNova = ""
If Tam <> 0 Then
    For i% = 1 To Tam
        Car = Mid$(str, i, 1)
        Select Case Asc(Car)
            Case 192, 193, 194, 195, 196, 197: strNova = strNova + "A"
            Case 200, 201, 202, 203: strNova = strNova + "E"
            Case 204, 205, 206, 207: strNova = strNova + "I"
            Case 210, 211, 212, 213, 214: strNova = strNova + "O"
            Case 217, 218, 219, 229: strNova = strNova + "U"
            Case 199: strNova = strNova + "C"
            Case 209: strNova = strNova + "N"
            Case 221: strNova = strNova + "Y"
            Case Else:
               'If Asc(Car) <> 8 Then
                  strNova = strNova + Car
               'End If
        End Select
    Next i
Else
    strNova = str
End If
String_Sem_Acento = strNova

End Function

'************************************************************************************
'   Função  : Substitui o texto de txtfrom$ para o texto de txtto$
'           na string txtin$.
'
'   Entrada : txtin$   - string a ser alterada
'             txtfrom$  - texto a ser substituido
'             txtto$  - novo texto
'
'   Saída   : string alterada
'
'   Dt. Mod.:
'   Modif.  :
'************************************************************************************
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
Public Function StrZero(Valor As Variant, Num As Integer, Optional Caracter = "0") As String
   Dim i%, Zeros$
   If Num% <= 0 Then
      StrZero = Valor
   Else
      Zeros = String(Num%, Caracter)
      StrZero = Right(Zeros + Trim(str(Val(Valor))), Num%)
   End If
End Function
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
    
    a% = InStr(com_ponto$, mvarSepMil$)
    While a% <> 0
        com_ponto = Left$(com_ponto$, a% - 1) + Mid$(com_ponto$, a% + 1)
        a% = InStr(com_ponto$, mvarSepMil$)
    Wend

    a% = InStr(com_ponto$, mvarSepDec$)
    
    If a% > 0 Then
       a = InStr(a% + 1, com_ponto$, mvarSepDec$)
       If a% > 0 Then com_ponto$ = Left$(com_ponto$, a% - 1)
    End If

    Tiraponto$ = com_ponto$

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
Public Function UnFormat(ByRef Codigo) As String
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

Function ValidarCGC(ByVal numCGC) As Integer
   
   Dim Resto%, Produto%, Tamanho%, i%, j%
   Dim Matriz$, CGC_Resultante$
   
   If Len(numCGC) <> 14 Then
       ValidarCGC = False
       Exit Function
   End If
   
   Matriz$ = "2345678923456"
   Tamanho% = 12
   CGC_Resultante$ = Left$(numCGC, Tamanho%)
   
   For j% = 1 To 2
       Produto% = 0
       For i% = Tamanho% To 1 Step -1
           Produto% = Produto% + Val(Mid$(CGC_Resultante$, i%, 1)) * Val(Mid$(Matriz$, Tamanho% - i% + 1, 1))
       Next i%
       Resto% = Produto% Mod 11
       CGC_Resultante$ = CGC_Resultante$ & IIf(Resto% <= 1, "0", Format$(11 - Resto%, "0"))
       Tamanho% = Tamanho% + 1
   Next j%
   
   ValidarCGC = (numCGC = CGC_Resultante$)

End Function

Function ValidarCPF(ByVal numCpf As String)
Dim digito, i As Integer
Dim DigitoCPF(11)

'Atribui o cpf digitado a um vetor
For i = 1 To 11
    DigitoCPF(i) = Val(Mid$(numCpf, i, 1))
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
   

Public Function ValidarCPFCGC(CpfCgc As Object, Bool%)
   ValidarCPFCGC = True
   If Len(Trim(CpfCgc.Text)) < 12 Then
      If Not ValidarCPF(Trim(CpfCgc.Text)) Then
         Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(23), ClsMsg.LoadMsg(1))
         ValidarCPFCGC = False
      End If
   Else
      If Not ValidarCGC(Trim(CpfCgc.Text)) Then
         Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(24), ClsMsg.LoadMsg(1))
         ValidarCPFCGC = False
      End If
   End If
   If Not Bool And Not ValidarCPFCGC Then CpfCgc.SetFocus
End Function

Public Function ValidarData(ByRef Campo As Object, Optional ByVal SetFocus_Object = False) As Boolean

Dim ValorAno%, TamData%
Dim Dia$, Mes$, Ano$, SECULO$, DtAux$
Dim Data$, Prp%

If UCase(TypeName(Campo)) = "MASKEDBOX" Then
   Prp% = Campo.PromptInclude
   Campo.PromptInclude = False
   Data = Mid(Campo, 1, 2) & "/" & Mid(Campo, 3, 2) & "/" & Mid(Campo, 5)
   Campo.PromptInclude = True
Else
   Data = Campo
End If

TamData = Len(Trim$(Campo))

If TamData > 0 Then
   If TamData = 5 Then
      Data = Data + "/" + Trim(str(Year(Now())))
      TamData = 10
   End If
   Select Case mvarFormatoData
      Case "DMA"
         Dia = Val(Mid(Data, 1, 2))
         Mes = Val(Mid(Data, 4, 2))
         ValorAno% = Val(Mid(Trim$(Data), 7))
         If ValorAno% = 0 Then ValorAno% = 2000
         Ano = Right$(Format$(ValorAno%, "0000"), 2)
         SECULO$ = IIf(ValorAno > 1000, Left$(ValorAno, 2), Left$(Year(Now), 2))
         Ano = Val(SECULO + Ano)
      Case "MDA"
         Mes = Val(Mid(Data, 1, 2))
         Dia = Val(Mid(Data, 4, 2))
         ValorAno% = Val(Mid(Trim$(Data), 7))
         If ValorAno% = 0 Then ValorAno% = 2000
         Ano = Right$(Format$(ValorAno%, "0000"), 2)
         SECULO$ = IIf(ValorAno > 1000, Left$(ValorAno, 2), Left$(Year(Now), 2))
         Ano = Val(SECULO + Ano)
      Case "AMD"
         ValorAno% = Val(Left$(Trim$(Data), TamData - 6))
         If ValorAno% = 0 Then ValorAno% = 2000
         Ano = Right$(Format$(ValorAno%, "0000"), 2)
         SECULO$ = IIf(ValorAno > 1000, Left$(ValorAno, 2), Left$(Year(Now), 2))
         Ano = Val(SECULO + Ano)
         Mes = Val(Mid(Data, TamData - 6 + 2, 2))
         Dia = Val(Right$(Data, 2))
   End Select
   If (Dia = 0 Or Mes = 0 Or Ano = 0) And UCase(TypeName(Campo)) = "MASKEDBOX" Then
      Campo.PromptInclude = Prp%
      Call ZerarMask(Campo, Campo.Mask)
      Exit Function
   End If
   DtAux = DateSerial(Ano, Mes, Dia)
   ValidarData = True
   If Not Format$(DtAux, mvarDtMask) = Format$(Data, mvarDtMask) Then
      Beep
      Select Case UCase(TypeName(Campo))
         Case "MASKEDBOX"
            Campo.PromptInclude = Prp%
            Call ZerarMask(Campo, Campo.Mask)
         Case Else
            Campo = ""
      End Select
      If SetFocus_Object Then
         Call ClsMsg.ExibirStop(ClsMsg.LoadMsg(25), ClsMsg.LoadMsg(1))
         Campo.SetFocus
      End If
      ValidarData = False
   Else
      Campo = Format$(Data, mvarDtMask)
   End If
End If
End Function

Public Sub ValidarHora(ByRef Campo)
   Dim Hora$
   Hora = Campo
   Hora = EliminarString(Hora, ":")
   Hora = StrZero(Hora, 4)
   If Val(Left(Hora, 2)) > 24 Or Val(Right(Hora, 2)) > 60 Then
      Beep
      Call ClsMsg.ExibirStop(ClsMsg.LoadMsg(26), ClsMsg.LoadMsg(1))
      Campo.SetFocus
   End If
   Campo.Text = Format$(Hora, "00:00")
End Sub
Public Function Vazio(Field As Variant) As Boolean
' Objetivo:  Simulação da função em Clipper empty()
'            No VB, um campo em branco (= "") não é IsEmpty
   
   Vazio = False
   
   Select Case VarType(Field)
     Case vbEmpty
        Vazio = True
     Case vbNull
        Vazio = True
     Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
        If Field = 0 Then Vazio = True
     Case vbDate
        If Field <= CDate("01/01/1901") Then Vazio = True
     Case vbString
        If Trim$(Field) = "" Or _
           Trim$(Field) = "/  /" Or _
           Trim$(Field) = "  /  /  " Or _
           Trim$(Field) = "  /  /    " Or _
           Trim$(Field) = "__/__/__" Or _
           Trim$(Field) = "__/__/____" Or _
           Trim$(Field) = "00:00:00" Then Vazio = True
     '9  vbObject
     '10 vbError
     Case vbBoolean
        If Field Then Vazio = True
   End Select
   Select Case UCase(TypeName(Field))
     Case "FIELD"
        If IsNull(Field) Then Vazio = True
     Case "TEXTBOX", "MASKEDBOX", "COMBOBOX"
        Vazio = Vazio(Val(Field.Text))
   End Select
End Function
Public Function VerificarData(ByVal Data$)
  VerificarData = True
'  If Trim(data$) Then
     If Not IsDate(Data$) Then
        Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(25), ClsMsg.LoadMsg(1))
        VerificarData = False
     End If
 ' End If
End Function
Public Sub Wait(Second As Integer)
   Dim TimeIni&
   TimeIni& = Timer
   While Timer - TimeIni& < Second
      DoEvents
   Wend
End Sub
Public Sub ZerarMask(ByRef Msk As Object, Optional ByRef MAscara)
   If IsMissing(MAscara) Then MAscara = Msk.Mask
   Msk.Mask = ""
   Msk.Text = ""
   Msk.Mask = MAscara
End Sub
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
Function ValBr(ByVal Num As String, Optional Casa_Dec)
   Dim sFormat As String
   Dim sNumber As String
   
   If IsMissing(Casa_Dec) Then
      Casa_Dec = QtdCasaDecimal(xVal(Num), 2)
   End If
   
   If Casa_Dec = 0 Then
      sFormat = "###,##0"
      sNumber = "1.000"
   Else
      sFormat = "###,##0." & String(Casa_Dec, "0")
      sNumber = "1.000," & String(Casa_Dec, "0")
   End If
   ValBr = Format(Num, sFormat)
   
   If sNumber <> Format(1000, sFormat) Then
      ValBr = StrReplace(Num, ".", ",")
   End If
   
   If Trim(ValBr) = "" Then
      If Casa_Dec = 0 Then
         ValBr = "0"
      Else
         ValBr = "0," & String(Casa_Dec, "0")
      End If
   End If
End Function
Function QtdCasaDecimal(pNum As Double, Optional Minimo As Integer = 2) As Integer
   Dim sNum As String
   Dim PosV As Integer
   
   sNum = CStr(pNum)
   PosV = InStr(sNum, ",")
   If PosV > 0 Then
      QtdCasaDecimal = Len(sNum) - PosV
   End If
   QtdCasaDecimal = IIf(QtdCasaDecimal < Minimo, Minimo, QtdCasaDecimal)
End Function
'Public Sub AutoInstalacao(Aplic As Object)
'   Call AutoInstala(Aplic)
'End Sub
Private Sub Class_Initialize()
   On Error Resume Next
'   If CDate(Format(Now, "dd/mm/yyyy")) >= CDate("01/01/2002") Then
'      MsgBox "Without License."
'      VBA.Err.Raise 1
'   End If
   mvarSepDec$ = ","
   mvarSepMil = "."
   mvarSepDt$ = "/"
   mvarDtMask$ = "dd" & mvarSepDt & "MM" & mvarSepDt & "yyyy"
   mvarDtMaskAux$ = "dd" & mvarSepDt & "MM" & mvarSepDt & "yy"
   mvarFormatoData = "DMA"

End Sub
Public Function Extenso(pNumber As String, Optional bMoeda As eMoeda = eMoeda.Real, Optional CasaDec = 2, Optional strMoeda, Optional strMoedas, Optional strCentavo = "Centavo", Optional strCentavos = "Centavos") As String
   Dim Tam        As Integer
   Dim StrNumber  As String
   Dim StrNum     As String
   Dim StrDec     As String
   Dim StrExtenso As String
   Dim sAux       As String
   Dim i          As Integer
   Dim CZero      As Boolean
   
   If Trim(pNumber) = "" Then Exit Function
   
   If IsMissing(strMoeda) Then strMoeda = ""
   If IsMissing(strMoedas) Then strMoedas = ""
   If IsMissing(strCentavo) Then strCentavo = ""
   If IsMissing(strCentavos) Then strCentavos = ""
   
   strMoeda = Trim(strMoeda)
   strMoedas = Trim(strMoedas)
   strCentavo = Trim(strCentavo)
   strCentavos = Trim(strCentavos)
   
   StrNumber = Trim(CStr(xVal(pNumber)))
   
   If InStr(StrNumber, ",") <> 0 Then
      StrNumber = Trim(CStr(xVal(pNumber)))
      
      Tam = InStr(StrNumber, ",") - 1
      
      StrNum = Mid(StrNumber, 1, Tam)
      StrDec = Mid(StrNumber, Tam + 2)
   Else
      StrNumber = Trim(CStr(xVal(pNumber & ",00")))
      Tam = Len(StrNumber)
      StrNum = StrNumber
      StrDec = "0"
   End If
   i = 0
   sAux = StrNum
   While Tam > 0
      i = i + 1
      
      If StrExtenso = "" Then
         StrExtenso = DefExtenso(Right$(sAux, 3), i)
      Else
         If i < 2 Then
            StrExtenso = DefExtenso(Right$(sAux, 3), i) & " e " & StrExtenso
         Else
            If CZero Then
               StrExtenso = DefExtenso(Right$(sAux, 3), i) & " e " & StrExtenso
            Else
               StrExtenso = DefExtenso(Right$(sAux, 3), i) & " " & StrExtenso
            End If
         End If
      End If
      
      CZero = (Mid(Right$("000" & sAux, 3), 1, 1) = "0")
      
      Tam = Tam - 3
      If Tam > 0 Then
         sAux = Mid(sAux, 1, Tam)
      End If
   Wend
   If bMoeda <> eMoeda.Numerico And Trim(StrExtenso) <> "" Then
      Select Case bMoeda
         Case eMoeda.Real:    StrExtenso = StrExtenso & IIf(pNumber = "1", " Real", " Reais")
         Case eMoeda.Dolar:   StrExtenso = StrExtenso & IIf(pNumber = "1", " Dolar", " Dolares")
         Case eMoeda.Outros:  StrExtenso = StrExtenso & IIf(pNumber = "1", " " & strMoeda, " " & strMoedas)
      End Select
   End If
   If xVal(StrDec) > 0 Then
      If Len(Trim(StrDec)) < CasaDec Then
         StrDec = Mid(StrDec & String(CasaDec, "0"), 1, CasaDec)
      End If
      StrExtenso = StrExtenso & IIf(bMoeda, " e ", " Vírgula ")
      StrExtenso = StrExtenso & Extenso(StrDec, False)
      Select Case bMoeda
         Case eMoeda.Numerico: StrExtenso = StrExtenso
         Case eMoeda.Real:     StrExtenso = StrExtenso & IIf(StrDec = "1", " Centavo", " Centavos")
         Case eMoeda.Dolar:    StrExtenso = StrExtenso & IIf(StrDec = "1", " Centavo", " Centavos")
         Case eMoeda.Outros:   StrExtenso = StrExtenso & IIf(StrDec = "1", " " & strCentavo, " " & strCentavos)
      End Select
      
      StrExtenso = StrExtenso ''& IIf(bMoeda, IIf(StrDec = "1", " Centavo", " Centavos"), "")
   End If
   
   Extenso = StrExtenso
End Function
Private Function DefExtenso(sNumber, i) As String
   Dim StrExt     As String
   Dim StrExtU    As String
   Dim StrExtD    As String
   Dim StrExtC    As String
   Dim sNum       As String
   Dim sUnidade   As String
   Dim sDezena    As String
   Dim sCentena   As String
   
   sNumber = xVal(sNumber)
   sNumber = VBA.Right$("000" & sNumber, 3)
   sUnidade = Mid(sNumber, 3, 1)
   sDezena = Mid(sNumber, 2, 1)
   sCentena = Mid(sNumber, 1, 1)
         
   If sDezena = "1" Then
      sNum = sDezena & sUnidade
      Select Case Val(sNum)
         Case 10: StrExtD = "Dez"
         Case 11: StrExtD = "Onze"
         Case 12: StrExtD = "Doze"
         Case 13: StrExtD = "Treze"
         Case 14: StrExtD = "Quatorze"
         Case 15: StrExtD = "Quinze"
         Case 16: StrExtD = "Dezesseis"
         Case 17: StrExtD = "Dezessete"
         Case 18: StrExtD = "Dezoito"
         Case 19: StrExtD = "Dezenove"
      End Select
   Else
      StrExtU = DefUnidadeExtenso(Val(sUnidade))
      StrExtD = DefDezenaExtenso(Val(sDezena))
   End If
   If sCentena = "1" And sDezena = "0" And sUnidade = "0" Then
      StrExtC = "Cem"
   Else
      StrExtC = DefCentenaExtenso(Val(sCentena))
   End If
   
   StrExt = StrExtC
   If StrExt <> "" And StrExtD <> "" Then
      StrExt = StrExt & " e " & StrExtD
   ElseIf StrExt = "" And StrExtD <> "" Then
      StrExt = StrExtD
   End If
   If StrExt <> "" And StrExtU <> "" Then
      StrExt = StrExt & " e " & StrExtU
   ElseIf StrExt = "" And StrExtU <> "" Then
      StrExt = StrExtU
   End If

   Select Case i
      Case 1: StrExt = StrExt & ""
      Case 2: StrExt = StrExt & " Mil"
      Case 3: StrExt = StrExt & IIf(StrExt = "Um", " Milhão", " Milhões")
      Case 4: StrExt = StrExt & IIf(StrExt = "Um", " Bilhão", " Bilhões")
      Case 5: StrExt = StrExt & IIf(StrExt = "Um", " Trilhão", " Trilhões")
      Case 6: StrExt = StrExt & IIf(StrExt = "Um", " Quatrihão", " Quatrihões")
   End Select
   
   DefExtenso = StrExt
End Function
Public Function DefUnidadeExtenso(pNum As Integer) As String
   Select Case pNum
      Case 1: DefUnidadeExtenso = "Um"
      Case 2: DefUnidadeExtenso = "Dois"
      Case 3: DefUnidadeExtenso = "Três"
      Case 4: DefUnidadeExtenso = "Quatro"
      Case 5: DefUnidadeExtenso = "Cinco"
      Case 6: DefUnidadeExtenso = "Seis"
      Case 7: DefUnidadeExtenso = "Sete"
      Case 8: DefUnidadeExtenso = "Oito"
      Case 9: DefUnidadeExtenso = "Nove"
   End Select
End Function
Public Function DefDezenaExtenso(pNum As Integer) As String
   Select Case pNum
      Case 2: DefDezenaExtenso = "Vinte"
      Case 3: DefDezenaExtenso = "Trinta"
      Case 4: DefDezenaExtenso = "Quarenta"
      Case 5: DefDezenaExtenso = "Cinquenta"
      Case 6: DefDezenaExtenso = "Sessenta"
      Case 7: DefDezenaExtenso = "Setenta"
      Case 8: DefDezenaExtenso = "Oitenta"
      Case 9: DefDezenaExtenso = "Noventa"
   End Select
End Function
Public Function DefCentenaExtenso(pNum As Integer) As String
   Select Case pNum
      Case 1: DefCentenaExtenso = "Cento"
      Case 2: DefCentenaExtenso = "Duzentos"
      Case 3: DefCentenaExtenso = "Trezentos"
      Case 4: DefCentenaExtenso = "Quatrocentos"
      Case 5: DefCentenaExtenso = "Quinhentos"
      Case 6: DefCentenaExtenso = "Seiscentos"
      Case 7: DefCentenaExtenso = "Setecentos"
      Case 8: DefCentenaExtenso = "Oitocentos"
      Case 9: DefCentenaExtenso = "Novecentos"
   End Select
End Function
Public Function EncryptMatrix(Text As String, Matrix As String)
   
   If Len(Text) = 2 Then
       EncryptMatrix = Text
       Exit Function
   End If
   
   Dim Mat() As Integer
   Dim Code() As Integer
   Dim Txt() As Integer
   Dim AmountOfRows As Integer
   Dim x As Integer
   Dim y As Integer
   Dim x2 As Integer
   Dim y2 As Integer

   'Get the numbers from Matrix and put them in Mat
   ReDim Preserve Mat(0 To 0)
   Do Until Matrix = ""
       ReDim Preserve Mat(0 To UBound(Mat) + 1)

       If InStr(Matrix, ",") <> 0 Then
           Mat(UBound(Mat)) = Mid(Matrix, 1, InStr(Matrix, ",") - 1)
           Matrix = Right(Matrix, Len(Matrix) - InStr(Matrix, ","))
       Else
           Mat(UBound(Mat)) = Matrix
           Matrix = ""
       End If
   Loop
   Do Until Int(Sqr(UBound(Mat))) = Sqr(UBound(Mat))
       ReDim Preserve Mat(0 To UBound(Mat) + 1)
       Mat(UBound(Mat)) = 3
   Loop
   AmountOfRows = Int(Sqr(UBound(Mat)))
   
   
   ReDim Preserve Txt(1 To AmountOfRows)
   ReDim Preserve Code(1 To AmountOfRows)
   Do Until Len(Text) < AmountOfRows
       'Get values from text
       For x = 1 To UBound(Code)
           Txt(x) = Asc(Mid(Text, 1, 1))
           Text = Right(Text, Len(Text) - 1)
       Next
       
       'Multiply the matrixes
       For x = 1 To UBound(Txt)
           Code(x) = 0
           For y = 1 To AmountOfRows
               Code(x) = Code(x) + (Mat(((UBound(Mat) / AmountOfRows) * (y - 1)) + x) * Txt(y))
           Next
       Next
       
       'Put the new values in the text
       For x = 1 To UBound(Txt)
           'More often then not, the value after being run
           'through the matrix is over 255. Because we can't
           'use the Chr function with numbers over 255 we need
           'a more complex way to put the value in the text.
           'Hear we will do it by finding a number to divide
           'it by then insert that number and the result. This
           'doubles the length, but what can you do?
           y = 1
           Do Until Code(x) / y <= 256 And Round(Code(x) / y, 0) * y = Code(x)
               y = y + 1
               If Code(x) = y Then 'The number is prime. Now we have to
                                   'assign a number to every prime number
                                   'above 255(1-257,2-263,3-269,ect). Then
                                   'insert an identifier character and the
                                   'prime number id.
                   y = 0
                   x2 = 255
                   Do
                       x2 = x2 + 1
                       For y2 = 2 To Code(x)
                           If Int(x2 / y2) = x2 / y2 Then GoTo NextNum
                           If y2 = x2 - 1 Then
                               y = y + 1
                               If x2 = Code(x) Then
                                   x2 = 0
                               End If
                               GoTo NextNum
                           End If
                       Next
NextNum:
                       
                       If x2 = 0 Then Exit Do
                   Loop
                   EncryptMatrix = EncryptMatrix & Chr$(255) & Chr$(y) 'Chr$(255) will be the identifier
                   y = 0
                   Exit Do
               End If
           Loop
           If y <> 0 Then EncryptMatrix = EncryptMatrix & Chr$(y) & Chr$(Code(x) / y)
       Next
   Loop
   EncryptMatrix = EncryptMatrix & Text
End Function
Public Function DecryptMatrix(Text As String, Matrix As String)
   If Len(Text) = 2 Then
       DecryptMatrix = Text
       Exit Function
   End If
   
   Dim Mat() As Double
   Dim Mat2() As Double
   Dim Determinant As Integer
   Dim Code() As Integer
   Dim Txt() As Integer
   Dim AmountOfRows As Integer
   Dim x As Integer
   Dim y As Integer
   Dim x2 As Integer
   Dim y2 As Integer


   'Get the numbers from Matrix and put them in Mat
   ReDim Preserve Mat(0 To 0)
   Do Until Matrix = ""
       ReDim Preserve Mat(0 To UBound(Mat) + 1)

       If InStr(Matrix, ",") <> 0 Then
           Mat(UBound(Mat)) = Mid(Matrix, 1, InStr(Matrix, ",") - 1)
           Matrix = Right(Matrix, Len(Matrix) - InStr(Matrix, ","))
       Else
           Mat(UBound(Mat)) = Matrix
           Matrix = ""
       End If
   Loop
   Do Until Int(Sqr(UBound(Mat))) = Sqr(UBound(Mat))
       ReDim Preserve Mat(0 To UBound(Mat) + 1)
       Mat(UBound(Mat)) = 3
   Loop
   AmountOfRows = Int(Sqr(UBound(Mat)))
   
   'Find the inverse of mat
       'The inverse of the matrix [a b]  is [1 / (a * d - c * b) * d          1 / (a * d - c * b) * (b * -1)]
       '                          [c d]     [1 / (a * d - c * b) * (c * -1)   1 / (a * d - c * b) * a       ]
       'a * d - c * b is called the determinant
   
   'Get the determinant
   Determinant = (Mat(1) * Mat(4)) - (Mat(2) * Mat(3))
   
   'Put the inverse of Mat into Mat2
   ReDim Mat2(1 To UBound(Mat))
   For x = 1 To UBound(Mat)
       If x = 1 Then
           Mat2(x) = (1 / Determinant) * Mat(4)
       ElseIf x = 4 Then
           Mat2(x) = (1 / Determinant) * Mat(1)
       Else
           Mat2(x) = (1 / Determinant) * (Mat(x) * -1)
       End If
   Next
   
   'Put valuse of Mat2 in Mat
   For x = 1 To UBound(Mat)
       Mat(x) = Mat2(x)
   Next
   
   ReDim Preserve Txt(1 To AmountOfRows)
   ReDim Preserve Code(1 To AmountOfRows)
   Do Until Len(Text) < AmountOfRows
       'Get values from code
       For x = 1 To UBound(Txt)
           If Asc(Mid(Text, 1, 1)) = 255 Then
               y = 0
               Code(x) = Asc(Right(Mid(Text, 1, 2), 1))
               x2 = 255
               Do
                   x2 = x2 + 1
                   y2 = 1
                   Do
                       y2 = y2 + 1
                       If Int(x2 / y2) = x2 / y2 Then Exit Do
                       If y2 * 2 > x2 - 1 Then
                           y = y + 1
                           If y = Code(x) Then
                               Code(x) = x2
                               x2 = 0
                           End If
                           Exit Do
                       End If
                   Loop
                   If x2 = 0 Then Exit Do
               Loop
           Else
               Code(x) = Asc(Mid(Text, 1, 1)) * Asc(Right(Mid(Text, 1, 2), 1))
           End If
           Text = Right(Text, LenB(Text) - 2)
       Next
       
       'Multiply the matrixes
       For x = 1 To UBound(Txt)
           Txt(x) = 0
           For y = 1 To AmountOfRows
               Txt(x) = Txt(x) + (Mat(((UBound(Mat) / AmountOfRows) * (y - 1)) + x) * Code(y))
           Next
       Next
       
       'Take value from the Txt and make text
       For x = 1 To UBound(Txt)
           DecryptMatrix = DecryptMatrix & Chr$(Txt(x))
       Next
   Loop
   DecryptMatrix = DecryptMatrix & Text
End Function
Public Function BDateDiff(StartDate As Date, EndDate As Date, Optional aHolidays As Variant) As Integer
   ' Copyright (c) 2002, TAG Consulting
   ' All rights reserved
   ' The following routine may be used in any application so long as
   ' the copyright notice remains intact
   
   Dim AdjEndDate As Date
   Dim AdjStartDate As Date
   Dim bEndLater As Boolean
   Dim bHolidays As Boolean
   Dim EndDays As Integer
   Dim i As Integer
   Dim MondayStart As Date
   Dim MondayEnd As Date
   Dim StartDays As Integer
   
   ' Determine if an array of holiday dates were included in the parameter list
   If IsMissing(aHolidays) Then
      bHolidays = False
   ElseIf Not IsArray(aHolidays) Then
      bHolidays = False
   Else
      bHolidays = True
   End If
   
   ' Determine the direction of the calculation
   Select Case DateDiff("d", StartDate, EndDate)
      Case Is > 0   ' Start date earlier than end date
         ' For weekends, move back to the previous Friday.
         Select Case Weekday(StartDate)
            Case 7:     AdjStartDate = DateAdd("d", -1, StartDate)
            Case 1:     AdjStartDate = DateAdd("d", -2, StartDate)
            Case Else:  AdjStartDate = StartDate
         End Select
         AdjEndDate = EndDate
         bEndLater = True
      Case Is < 0    ' End date earlier than start date
         ' For weekends, move up to the next Monday.
         Select Case Weekday(EndDate)
            Case 7:     AdjStartDate = DateAdd("d", 2, EndDate)
            Case 1:     AdjStartDate = DateAdd("d", 1, EndDate)
            Case Else:  AdjStartDate = EndDate
         End Select
         AdjEndDate = StartDate
         bEndLater = False
      Case 0   ' Trivial calculation averted
         BDateDiff = 0
         Exit Function
   End Select
   
   StartDays = Weekday(AdjStartDate) - 2
   MondayStart = DateAdd("d", StartDays * -1, AdjStartDate)
   EndDays = Weekday(AdjEndDate) - 2
   MondayEnd = DateAdd("d", -1 * EndDays, AdjEndDate)
   
   BDateDiff = (5 / 7) * DateDiff("d", MondayStart, MondayEnd) - StartDays + EndDays

   If bHolidays Then
      For i = LBound(aHolidays, 1) To UBound(aHolidays, 1)
         If IsDate(aHolidays(i)) Then
            If DateDiff("d", AdjStartDate, aHolidays(i)) >= 0 And DateDiff("d", AdjEndDate, aHolidays(i)) <= 0 Then
               BDateDiff = BDateDiff - 1
            End If
         End If
      Next i
   End If
   If Not bEndLater Then
      BDateDiff = BDateDiff * -1
   End If
End Function
Public Function BDateAdd(StartDate As Date, Interval As Integer, Optional aHolidays As Variant) As Date
   ' Copyright (c) 2002, TAG Consulting
   ' All rights reserved
   ' The following routine may be used in any application so long as the copyright notice remains intact
   
   Dim AdjustedDate As Date
   Dim bHolidays As Boolean
   Dim i As Integer
   Dim TempDate As Date
   Dim TestInterval As Integer
   
   ' Determine if an array of holiday dates were included in the parameter list
   If IsMissing(aHolidays) Then
      bHolidays = False
   ElseIf Not IsArray(aHolidays) Then
      bHolidays = False
   Else
      bHolidays = True
   End If
   
   Select Case Interval
      Case Is > 0
         ' For weekends, move back to the previous Friday.
         Select Case Weekday(StartDate)
            Case 7:     AdjustedDate = DateAdd("d", -1, StartDate)
            Case 1:     AdjustedDate = DateAdd("d", -2, StartDate)
            Case Else:  AdjustedDate = StartDate
         End Select
      Case Is < 0
         ' For weekends, move up to the next Monday.
         Select Case Weekday(StartDate)
            Case 7:     AdjustedDate = DateAdd("d", 2, StartDate)
            Case 1:     AdjustedDate = DateAdd("d", 1, StartDate)
            Case Else:  AdjustedDate = StartDate
         End Select
      Case 0
         BDateAdd = StartDate
         Exit Function
   End Select
   
   TempDate = DateAdd("d", (Interval \ 5) * 7 + (Interval Mod 5), AdjustedDate)
   
   ' If the approximation ends on a weekend, adjust to a Monday
   Select Case Weekday(TempDate)
      Case 7: TempDate = DateAdd("d", 2, TempDate)
      Case 1: TempDate = DateAdd("d", 1, TempDate)
   End Select
   
   ' Perform a validation by using the DateDiff and adjusting if necessary
   TestInterval = BDateDiff(StartDate, TempDate)
   If TestInterval <> Interval Then
      TempDate = DateAdd("d", Interval - TestInterval, TempDate)
   End If

   BDateAdd = TempDate

   If bHolidays Then
      For i = LBound(aHolidays, 1) To UBound(aHolidays, 1)
         If IsDate(aHolidays(i)) Then
            If DateDiff("d", AdjustedDate, aHolidays(i)) > 0 And DateDiff("d", BDateAdd, aHolidays(i)) <= 0 Then
               BDateAdd = DateAdd("d", 1, BDateAdd)
               ' Adjust for weekends
               Select Case Weekday(BDateAdd)
                  Case 7: BDateAdd = DateAdd("d", 2, BDateAdd)
                  Case 1: BDateAdd = DateAdd("d", 1, BDateAdd)
               End Select
            End If
         End If
      Next i
   End If
End Function
Public Function ParseLine(pTexto As String, pTamn As Integer) As String
    Dim stTmp$, stFinal$
    Dim cTmpPesq$
    Dim stNew$
    Dim i
    
    If pTexto <> "" Then
       
       For i = 1 To Len(pTexto)
         If Mid(pTexto, i, 1) = Chr$(13) Then
           stTmp$ = stTmp$ + " "
         ElseIf Mid(pTexto, i, 1) = Chr$(10) Then
           
         Else
           stTmp$ = stTmp$ + Mid(pTexto, i, 1)
         End If
       Next
       
       pTexto = RTrim(stTmp$)
       cTmpPesq$ = pTexto
       Do
         stTmp$ = stGetToken(cTmpPesq$, " ")
         If Len(stNew$ + stTmp$) < pTamn Then
            stNew$ = stNew$ + " " + stTmp$
         Else
            stFinal$ = stFinal$ + " " + stNew$ + vbCrLf + stTmp$
            stNew$ = ""
         End If

       Loop Until cTmpPesq$ = ""
       If stFinal$ = "" Then
         stFinal$ = stNew$
       End If
     End If
     ParseLine = stFinal$
End Function
Public Function stGetToken(stLn$, stDelim$) As String
    Dim iOpenQuote%, iCloseQuote%
    Dim iDelim%, stToken$
    On Error GoTo GetTokenError

    iOpenQuote% = InStr(1, stLn$, """")
    iDelim% = InStr(1, stLn$, stDelim$)

    If (iOpenQuote% > 0) And (iOpenQuote% < iDelim%) Then
         iCloseQuote% = InStr(iOpenQuote% + 1, stLn$, """")
         iDelim% = InStr(iCloseQuote% + 1, stLn$, stDelim$)
    End If

    If (iDelim% <> 0) Then
         stToken$ = LTrim$(RTrim$(Mid$(stLn$, 1, iDelim% - 1)))
         stLn$ = Mid$(stLn$, iDelim% + 1)
    Else
         stToken$ = LTrim$(RTrim$(Mid$(stLn$, 1)))
         stLn$ = ""
    End If

    If (Len(stToken$) > 0) Then
         If (Mid$(stToken$, 1, 1) = """") Then
              stToken$ = Mid$(stToken$, 2)
         End If
         If (Mid$(stToken$, Len(stToken$), 1) = """") Then
              stToken$ = Mid$(stToken$, 1, Len(stToken$) - 1)
         End If
    End If
    stGetToken = stToken$

GetTokenExit:
    Exit Function

GetTokenError:
    Resume GetTokenExit
End Function

Private Sub Class_Terminate()
   On Error Resume Next
End Sub
