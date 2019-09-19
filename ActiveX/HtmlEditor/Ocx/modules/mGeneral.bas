Attribute VB_Name = "mGeneral"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long   'Frequency is in hertz, Duration is in milliseconds

' Sub to sleep x seconds
Public Sub Sleep(lngSleep As Long)
   Dim lngSleepEnd As Long
   lngSleepEnd = GetTickCount + lngSleep * 1000
   While GetTickCount <= lngSleepEnd
      DoEvents
   Wend
End Sub

' Sub to freeze x seconds
Public Sub Freeze(lngFreeze As Long)
   Dim lngFreezeEnd As Long
   lngFreezeEnd = GetTickCount + lngFreeze * 1000
   While GetTickCount <= lngFreezeEnd
   Wend
End Sub
Public Function TrimNull(sString As String) As String
    TrimNull = Left(sString, InStr(1, sString, vbNullChar) - 1)
End Function

'Sort string arrays
Sub SortArray(inpArray())
    Dim intRet
    Dim intCompare
    Dim intLoopTimes
    Dim strTemp
    
    For intLoopTimes = 1 To UBound(inpArray)
        For intCompare = LBound(inpArray) To UBound(inpArray) - 1
            intRet = StrComp(inpArray(intCompare), _
                     inpArray(intCompare + 1), vbTextCompare)
    
            If intRet = 1 Then 'String1 is greater than String2
                strTemp = inpArray(intCompare)
                inpArray(intCompare) = inpArray(intCompare + 1)
                inpArray(intCompare + 1) = strTemp
            End If
        Next
    Next

End Sub

' For put a windows in the middle of the screen
' FrmChild  = Windows to center
' FrmParent = MDI Windows (Optional)
Public Sub CenterForm(FrmChild As Form, Optional FrmParent As Variant)
    Dim iTop As Integer, iLeft As Integer
    
    If Not IsMissing(FrmParent) Then
        iTop = FrmParent.Top + (FrmParent.ScaleHeight - FrmChild.Height) \ 2
        iLeft = FrmParent.Left + (FrmParent.ScaleWidth - FrmChild.Width) \ 2
    Else
        iTop = (Screen.Height - FrmChild.Height) \ 2
        iLeft = (Screen.Width - FrmChild.Width) \ 2
    End If
    If iTop And iLeft Then
        FrmChild.Move iLeft, iTop
    End If
End Sub

Public Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
   If Right$(sPath, 1) <> "\" Then
        QualifyPath = sPath & "\"
   Else
        QualifyPath = sPath
   End If
      
End Function

Public Function QualifyURL(ByVal strURL As String) As String
    
    Dim URL As String
    
    URL = Trim(LCase(strURL))
    
    If URL = "" Then
        QualifyURL = ""
        Exit Function
    End If
    
    If InStr(1, URL, "http://", vbTextCompare) = 1 Or _
        InStr(1, URL, "https://", vbTextCompare) = 1 Or _
        InStr(1, URL, "ftp://", vbTextCompare) = 1 Or _
        InStr(1, URL, "file://", vbTextCompare) = 1 Or _
        InStr(1, URL, "gopher://", vbTextCompare) = 1 Or _
        InStr(1, URL, "wais:", vbTextCompare) = 1 Or _
        InStr(1, URL, "telnet:", vbTextCompare) = 1 Or _
        InStr(1, URL, "mailto:", vbTextCompare) = 1 Or _
        InStr(1, URL, "news:", vbTextCompare) = 1 _
    Then
        QualifyURL = Trim(strURL)
    Else
        QualifyURL = "http://" & Trim(strURL)
    End If
    
End Function



