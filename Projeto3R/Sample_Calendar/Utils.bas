Attribute VB_Name = "Utils"
Option Explicit

Private Const API_NULL As Long = 0

Private Const WM_GETFONT = &H31

Private Const CB_ERR = -1
Private Const CB_GETLBTEXT = &H148
Private Const CB_GETLBTEXTLEN = &H149
'Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETHORIZONTALEXTENT = &H15E
'Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160

Private Const LB_ERR = -1
Private Const LB_GETTEXT = &H189
Private Const LB_GETTEXTLEN = &H18A
Private Const LB_ITEMFROMPOINT = &H1A9
'Private Const LB_GETHORIZONTALEXTENT = &H193
Private Const LB_SETHORIZONTALEXTENT = &H194

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const SM_CXFULLSCREEN = 16
'Private Const SM_CXBORDER = 5
'Private Const SM_CXHSCROLL = 21
'Private Const SM_CXHTHUMB = 10
Private Const SM_CXVSCROLL = 2

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_HSCROLL = &H100000
'Private Const WS_VSCROLL = &H200000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Declare Function GetTextExtentPoint32 Lib "gdi32" _
    Alias "GetTextExtentPoint32A" _
   (ByVal hDC As Long, _
    ByVal lpsz As String, _
    ByVal cbString As Long, _
    lpSize As SIZE) As Long

Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEFAULT_GUI_FONT = 17

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Const GDI_ERROR = &HFFFF

Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

'Required to call CreateFont()
Private Const LOGPIXELSY = 90

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long _
    , ByVal nIndex As Long) As Long

Private Const FW_BOLD = 700
Private Const FW_NORMAL = 400
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" ( _
    ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long _
    , ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long _
    , ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long


Public Function ParseTimeDuration(ByVal strTime As String, ByRef pnMinutes As Long) As Boolean
    pnMinutes = 0
    ParseTimeDuration = False
        
    Dim nI As Long, nLen As Long
    Dim nMeasureStart As Long, nFIdx As Long
    Dim strChI As String
        
    strTime = Trim(strTime)
    nLen = Len(strTime)
    
    If nLen = 0 Then
        Exit Function
    End If
    
    '------------------------------------------
    nMeasureStart = -1
    For nI = 1 To nLen
        strChI = Mid(strTime, nI, 1)
        nFIdx = InStr(1, "-+.,0123456789", strChI)
        If nFIdx <= 0 Then
            nMeasureStart = nI
            Exit For
        End If
    Next
    
    Dim strNumber As String, strMeasure As String
    Dim nMultiplier As Long
            
    If nMeasureStart > 0 Then
        strNumber = Left(strTime, nMeasureStart - 1)
        strMeasure = Mid(strTime, nMeasureStart)
        strMeasure = Trim(strMeasure)
    Else
        strNumber = strTime
    End If
    
    If Len(strNumber) = 0 Then
        Exit Function
    End If

    Dim strM0 As String
    strM0 = Left(strMeasure, 1)
    
    nMultiplier = 1
    If strM0 = "m" Or strM0 = "M" Then
        nMultiplier = 1
    ElseIf strM0 = "h" Or strM0 = "H" Then
        nMultiplier = 60
    ElseIf strM0 = "d" Or strM0 = "D" Then
        nMultiplier = 60 * 24
    ElseIf strM0 = "w" Or strM0 = "W" Then
        nMultiplier = 60 * 24 * 7
    End If

    Dim dblTime As Double
    dblTime = Val(strNumber)
    
    pnMinutes = dblTime * nMultiplier
    ParseTimeDuration = True
End Function

Public Function FormatTimeDuration(ByVal nMinutes As Long, ByVal bAprox As Boolean) As String
    Dim nWeeks As Long, nDays As Long, nHours As Long
    
    nWeeks = nMinutes / (7 * 24 * 60)
    nDays = nMinutes / (24 * 60)
    nHours = nMinutes / 60

    Dim strDuration As String
    
    If (bAprox Or (nMinutes Mod (7 * 24 * 60)) = 0) And nWeeks > 0 Then
        strDuration = nWeeks & " week" & IIf(nWeeks > 1, "s", "")
    
    ElseIf (bAprox Or (nMinutes Mod (24 * 60)) = 0) And nDays > 0 Then
        strDuration = nDays & " day" & IIf(nDays > 1, "s", "")
        
    ElseIf (bAprox Or (nMinutes Mod 60) = 0) And nHours > 0 Then
        strDuration = nHours & " hour" & IIf(nHours > 1, "s", "")
        
    Else
        strDuration = nMinutes & " minute" & IIf(nMinutes > 1, "s", "")
    End If

    FormatTimeDuration = strDuration
End Function

Public Sub FillStandardDurations_0m_2w(cmbDuration As ComboBox, bSnoozeBox As Boolean)
    
    If Not bSnoozeBox Then
        cmbDuration.AddItem "0 minutes"
        cmbDuration.AddItem "1 minute"
    End If
    
    cmbDuration.AddItem "5 minutes"
    cmbDuration.AddItem "10 minutes"
    cmbDuration.AddItem "15 minutes"
    cmbDuration.AddItem "30 minutes"
    
    cmbDuration.AddItem "1 hour"
    cmbDuration.AddItem "2 hours"
    cmbDuration.AddItem "4 hours"
    cmbDuration.AddItem "8 hours"
    
    cmbDuration.AddItem "0.5 day"
    cmbDuration.AddItem "1 day"
    cmbDuration.AddItem "2 days"
    cmbDuration.AddItem "3 days"
    cmbDuration.AddItem "4 days"
    
    cmbDuration.AddItem "1 week"
    cmbDuration.AddItem "2 weeks"
End Sub

Public Function CalcStandardDurations_0m_2wLong(sDuration As String) As Long
    Select Case sDuration
        Case "0 minutes":
            CalcStandardDurations_0m_2wLong = 0
        Case "1 minute":
            CalcStandardDurations_0m_2wLong = 1
        Case "5 minutes":
            CalcStandardDurations_0m_2wLong = 5
        Case "10 minutes":
            CalcStandardDurations_0m_2wLong = 10
        Case "15 minutes":
            CalcStandardDurations_0m_2wLong = 15
        Case "30 minutes":
            CalcStandardDurations_0m_2wLong = 30
        
        Case "1 hour":
            CalcStandardDurations_0m_2wLong = 60
        Case "2 hours":
            CalcStandardDurations_0m_2wLong = 60 * 2
        Case "4 hours":
            CalcStandardDurations_0m_2wLong = 60 * 4
        Case "8 hours":
            CalcStandardDurations_0m_2wLong = 60 * 8
        
        Case "0.5 day":
            CalcStandardDurations_0m_2wLong = 60 * 12
        Case "1 day":
            CalcStandardDurations_0m_2wLong = 60 * 24
        Case "2 days":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 2
        Case "3 days":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 3
        Case "4 days":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 4
        
        Case "1 week":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 7
        Case "2 weeks":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 7 * 2
    End Select
End Function

Public Function CalcStandardDurations_0m_2wString(lDuration As Long) As String
    Select Case lDuration
        Case 0:
            CalcStandardDurations_0m_2wString = "0 minutes"
        Case 1:
            CalcStandardDurations_0m_2wString = "1 minute"
        Case 5:
            CalcStandardDurations_0m_2wString = "5 minutes"
        Case 10:
            CalcStandardDurations_0m_2wString = "10 minutes"
        Case 15:
            CalcStandardDurations_0m_2wString = "15 minutes"
        Case 30:
            CalcStandardDurations_0m_2wString = "30 minutes"
        
        Case 60:
            CalcStandardDurations_0m_2wString = "1 hour"
        Case (60 * 2):
            CalcStandardDurations_0m_2wString = "2 hours"
        Case (60 * 4):
            CalcStandardDurations_0m_2wString = "4 hours"
        Case (60 * 8):
            CalcStandardDurations_0m_2wString = "8 hours"
        
        Case (60 * 12):
            CalcStandardDurations_0m_2wString = "0.5 day"
        Case (60 * 24):
            CalcStandardDurations_0m_2wString = "1 day"
        Case (60 * 24 * 2):
            CalcStandardDurations_0m_2wString = "2 days"
        Case (60 * 24 * 3):
            CalcStandardDurations_0m_2wString = "3 days"
        Case (60 * 24 * 4):
            CalcStandardDurations_0m_2wString = "4 days"
        
        Case (60 * 24 * 7):
            CalcStandardDurations_0m_2wString = "1 week"
        Case (60 * 24 * 7 * 2):
            CalcStandardDurations_0m_2wString = "2 weeks"
    End Select
End Function

Public Function BooleanToBin(bVal As Boolean) As Long
    If bVal Then
        BooleanToBin = 1
    Else
        BooleanToBin = 0
    End If
End Function

Public Function BinToBoolean(nVal As Long) As Boolean
    If nVal = 0 Then
        BinToBoolean = False
    Else
        BinToBoolean = True
    End If
End Function

Public Function ColorToStr(ByVal clr As OLE_COLOR) As String
    Dim strColor As String
  
    strColor = clr Mod 256
    strColor = strColor & ", " & (clr \ 256 Mod 256)
    strColor = strColor & ", " & (clr \ 256 \ 256 Mod 256)
    
    ColorToStr = strColor
End Function

Public Sub CopyFont(fntDest As StdFont, fntSrc As StdFont)
    fntDest.Bold = fntSrc.Bold
    fntDest.Italic = fntSrc.Italic
    fntDest.Name = fntSrc.Name
    fntDest.SIZE = fntSrc.SIZE
    fntDest.Strikethrough = fntSrc.Strikethrough
    fntDest.Underline = fntSrc.Underline
End Sub

Public Function AreFontsDifferent(fnt1 As StdFont, fnt2 As StdFont) As Boolean
    AreFontsDifferent = _
            fnt1.Bold <> fnt2.Bold Or _
            fnt1.Italic <> fnt2.Italic Or _
            StrComp(fnt1.Name, fnt2.Name) <> 0 Or _
            fnt1.SIZE <> fnt2.SIZE Or _
            fnt1.Strikethrough <> fnt2.Strikethrough Or _
            fnt1.Underline <> fnt2.Underline
End Function

    

'=================================================================================
' All the following code is taken from MSDN Magazine article:
' ActiveX and Visual Basic: Enhance the Display of Long Text Strings in a Combobox or Listbox
'
Private Sub mComboAdjustWidth(cboCombo As ComboBox)
'Purpose: Adjust the width of a Combo dropdown to fit the largest item in the list

    Dim lItemLen As Long
    Dim lItemMaxLen As Long
    Dim sItemText As String
    Dim sItemTemp As String
    
    Dim lResult As Long
    
    Dim lLength As Long
    Dim lVerticalScrollbarWidth As Long
    Dim lScreenWidth As Long
    
    Dim i As Integer
    
    'Find the longest item in the list (by number of characters)
    'Note: To be 100% accurate, we should compare the GetTextExtentPoint32 values of each
    'string rather than the Length. Would this be much slower? For now, take the simple
    'way out.
    lItemMaxLen = 0
    sItemText = ""
    For i = 0 To cboCombo.ListCount - 1
        'Note: Informal timings show that CB_GETLBTEXTLEN is about twice as fast as Len()
        lResult = SendMessage(cboCombo.hwnd, CB_GETLBTEXTLEN, i, ByVal 0)
        
        If (lResult = CB_ERR) Then
            gErrHandlerAPI "mComboAdjustWidth"
            GoTo NextItem
        End If
        
        'If (chkSendMessage.Value = vbChecked) Then
        If True Then
            sItemTemp = Space(lResult + 1)
            
            lResult = SendMessage(cboCombo.hwnd, CB_GETLBTEXT, i, ByVal sItemTemp)
            
            If (lResult = CB_ERR) Then
                gErrHandlerAPI "mComboAdjustWidth"
                GoTo NextItem
            End If
            
            sItemTemp = Left(sItemTemp, lResult)
            
            lResult = mlStringLenInControl(cboCombo, sItemTemp)
            'Debug.Print lResult, sItemTemp
        
            'If the current item is longer than the longest found so far...
            If (lResult > lItemMaxLen) Then
                'remember the size and string value of the current item
                lItemMaxLen = lResult
                sItemText = sItemTemp
            End If
        Else
            'If the current item is longer than the longest found so far...
            If (lResult > lItemMaxLen) Then
                'remember the size and string value of the current item
                lItemMaxLen = lResult
                sItemText = Space(lResult + 1)
                
                lResult = SendMessage(cboCombo.hwnd, CB_GETLBTEXT, i, ByVal sItemText)
            
                If (lResult = CB_ERR) Then
                    gErrHandlerAPI "mComboAdjustWidth"
                    GoTo NextItem
                End If
            
                sItemText = Left(sItemText, lItemMaxLen)
                
                'Debug.Print sItemText
            End If
        End If

NextItem:
    Next
   
    'Determine the width of the longest string found, in the context of the Combo font
    lLength = mlStringLenInControl(cboCombo, sItemText)
   
    If (lLength = 0) Then GoTo Exit_
    
    'Account for the window border, 1 pixel either side
    lLength = lLength + 2
    
    'Account for the scrollbar width plus a fudge factor
    lVerticalScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL)
    If (lVerticalScrollbarWidth <> 0) Then
        lLength = lLength + lVerticalScrollbarWidth + 4
    End If
    
    'Account for the maximum screen width
    lScreenWidth = GetSystemMetrics(SM_CXFULLSCREEN)
    If (lLength > lScreenWidth) And (lScreenWidth <> 0) Then
        lLength = lScreenWidth
    End If
    
    'Set the width of the Combo dropdown.
    lResult = SendMessage(cboCombo.hwnd, CB_SETDROPPEDWIDTH, lLength, ByVal 0)
     
    If (lResult = CB_ERR) Then GoTo Err
     
Exit_:
    Exit Sub
    
Err:
    gErrHandlerAPI "mComboAdjustWidth"

End Sub

Public Sub mListBoxAdjustHScroll(ctlControl As Control)
'Purpose: Adjust the horizontal scroll extent of a ListBox to fit the largest item in the
'list

    Dim lItemLen As Integer
    Dim lItemMaxLen As Integer
    Dim sItemText As String
    
    Dim lResult As Long
    
    Dim lLength As Long
    
    Dim i As Integer
    
    Select Case TypeName(ctlControl)
    Case "ListBox", "ComboBox"
    Case Else
        MsgBox "mListBoxAdjustHScroll: Invalid argument value for ctlControl."
        GoTo Exit_
    End Select
    
    'Find the longest item in the ListBox (by number of characters)
    'Note: To be 100% accurate, we should compare the GetTextExtentPoint32 values of each
    'string rather than the Length. Would this be much slower? For now, take the simple
    'way out.
    lItemMaxLen = 0
    sItemText = ""
    For i = 0 To ctlControl.ListCount - 1
        Select Case TypeName(ctlControl)
        Case "ComboBox"
            lResult = SendMessage(ctlControl.hwnd, CB_GETLBTEXTLEN, i, ByVal 0)
        Case "ListBox"
            lResult = SendMessage(ctlControl.hwnd, LB_GETTEXTLEN, i, ByVal 0)
        End Select
        
        If (lResult = CB_ERR) Then
            gErrHandlerAPI "mListBoxAdjustHScroll"
            GoTo NextItem
        End If
        
        'If the current item is longer than the longest found so far...
        If (lResult > lItemMaxLen) Then
            'remember the size and string value of the current item
            lItemMaxLen = lResult
            sItemText = ctlControl.List(i)
        End If
NextItem:
    Next
   
    'Determine the width of the longest string found, in the context of the Combo font
    lLength = mlStringLenInControl(ctlControl, sItemText)
   
    If (lLength = 0) Then GoTo Exit_
    
    'Fudge factor
    lLength = lLength + 4
    
    'Set the horizontal scrollbar extent
    Select Case TypeName(ctlControl)
    Case "ComboBox"
        lResult = SendMessage(ctlControl.hwnd, CB_SETHORIZONTALEXTENT, lLength, ByVal 0)
    Case "ListBox"
        lResult = SendMessage(ctlControl.hwnd, LB_SETHORIZONTALEXTENT, lLength, ByVal 0)
    End Select
    
Exit_:
    Exit Sub
    
Err:
    gErrHandlerAPI "mListBoxAdjustHScroll"

End Sub

Private Function mlStringLenInControl(ctlControl As Control, ByVal vsString As String) As Long
'Purpose: Determine the length in pixels of a string in the device context of ctlControl

    Dim hDC As Long
    Dim lFont As Long
    Dim lFontOld As Long
    Dim uSize As SIZE
   
    Dim lHeight As Long
    Dim lResult As Long
    
    mlStringLenInControl = 0

    With ctlControl
        'Get a handle to the device context for the control
        hDC = GetDC(.hwnd)
        If (hDC = API_NULL) Then GoTo ErrAPI
        
        'lFont = GetStockObject(ANSI_FIXED_FONT)
        lFont = GetStockObject(ANSI_VAR_FONT)
        'lFont = GetStockObject(SYSTEM_FONT)
        'lFont = GetStockObject(DEFAULT_GUI_FONT)
            
        If (lFont = API_NULL) Then GoTo ErrAPI
    
    End With
    
    'Select the font in to the device context, and retain prior font
    lFontOld = SelectObject(hDC, lFont)
    If (lFontOld = 0) Or (lFontOld = GDI_ERROR) Then GoTo ErrAPI
   
    'Determine the width of the string
    lResult = GetTextExtentPoint32(hDC, vsString, Len(vsString), uSize)
    If (lResult = 0) Then GoTo ErrAPI
   
    'Return the string length
    mlStringLenInControl = uSize.cx

Exit_:
    'Reset the device context font and delete the temporary font. Ignore any errors.
    SelectObject hDC, lFontOld
    DeleteObject lFont
     
    'Release the device context handle. Ignore any errors.
    ReleaseDC ctlControl.hwnd, hDC

    Exit Function
    
Err:
    MsgBox "mlStringLenInControl: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    Resume Exit_
    
ErrAPI:
    gErrHandlerAPI "mlStringLenInControl"
    GoTo Exit_

'Resume
End Function

Public Sub gErrHandlerAPI(ByVal vsRoutine As String, Optional ByVal vlMessageId As Variant)

    Dim sBuffer As String
    Dim lReturn As Long
    
    If IsMissing(vlMessageId) Then
        vlMessageId = GetLastError()
    End If
    
    sBuffer = Space(255)
    lReturn = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, vlMessageId, 0, sBuffer, 255, 0)

    If (lReturn > 0) Then
        sBuffer = Left(sBuffer, lReturn - 1)
        MsgBox "WinAPI (" & vsRoutine & "): " & sBuffer
    Else
        MsgBox "WinAPI (" & vsRoutine & "): " & vlMessageId & " - No description exists for this error number."
    End If
    
End Sub

Private Sub mAddScrollBar(oControl As Control)

    Dim lWindowStyle As Long
    
    lWindowStyle = GetWindowLong(oControl.hwnd, GWL_STYLE)
    
    If (lWindowStyle = 0) Then
        gErrHandlerAPI "mAddScrollBar"
        Exit Sub
    End If

    lWindowStyle = lWindowStyle Or WS_HSCROLL
    
    SetLastError 0
    
    lWindowStyle = SetWindowLong(oControl.hwnd, GWL_STYLE, lWindowStyle)

End Sub



