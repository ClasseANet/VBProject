Attribute VB_Name = "Tmp"
'Dim m_ProjectInfo As ProjectInfo
'// constants
Const Project_Ver = "1.0"
Private gPaths() As String
Private gFolders() As String
Public blnCancel As Boolean
Private blnBusy As Boolean
'// Event Declarations:
'Event OnOpen(strFile As String, strtitle As String)
'Event OnSaveAs(strFile As String, strtitle As String)
'Event OnNew()
'Event OnClose()
Private blnDragging As Boolean
Private mbSplitting As Boolean    '// Are we splitting?
Private m_blnShowCPL As Boolean

Private Const lVSplitLimit As Long = 500   '// Splitter side limits
Public Function GetCaption(ByVal strPath As String) As String
   If IsDrive(strPath) Then
      GetCaption = strPath
   Else
      While InStr(strPath, "\") <> 0
         strPath = Right$(strPath, Len(strPath) - InStr(strPath, "\"))
      Wend
      GetCaption = strPath
   End If
End Function
Public Function IsDrive(ByVal strString As String) As Boolean
   If strString Like "*:" Or strString Like "*:\" Then
      IsDrive = True
   End If
End Function
Public Function ParseFileName(ByVal strFile As String, ByVal strLine As String, ByRef strFullpath As String, ByRef strtitle As String)
   If InStr(1, strLine, ";") <> 0 Then
      strtitle = Right$(strLine, Len(strLine) - InStr(1, strLine, ";") - 1)
   Else
      strtitle = Right$(strLine, Len(strLine) - InStr(1, strLine, "="))
   End If
   strtitle = EliminarString(strtitle, """")
   strFullpath = GetFolder(strFile) '// folder
   Do While strtitle Like "..\*"
      '// up another folder
      If InStr(strFullpath, "\") = 0 Then strFullpath = strFullpath & "\"
      Pos = 1
      While InStr(Pos, strFullpath, "\")
         Pos = InStr(Pos, strFullpath, "\") + 1
      Wend
      strFullpath = Mid(strFullpath, 1, Pos - 2)

      '        strFullpath = Mid(strFullpath, 1, InStr(strFullpath, "\") - 1)
      strtitle = Right$(strtitle, Len(strtitle) - 3)
   Loop
   strFullpath = strFullpath & "\" & strtitle

   '   strtitle = GetCaption(strtitle)
   Do While InStr(strtitle, "\") <> 0
      strtitle = Mid(strtitle, InStr(strtitle, "\") + 1)
   Loop

   If InStr(1, strLine, ";") <> 0 Then
      strtitle = Mid(strLine, InStr(strLine, "=") + 1, InStr(strLine, ";") - InStr(strLine, "=") - 1) & " (" & strtitle & ")"
      '   Else
      '      strtitle = Mid(strLine, InStr(strLine, "=") + 1)
   End If
   If LCase(Mid(Trim(strLine), 1, 4)) = "form" Then
      strtitle = gConstru.LerPropriedade(strFullpath, "Attribute VB_Name", False) & " (" & strtitle & ")"
   End If
   strtitle = EliminarString(strtitle, """")
End Function
Public Function GetFolder(strPath As String) As String
   Dim Pos%
   On Error Resume Next
   If strPath Like "*:" Or strPath Like "*:\" Then
      GetFolder = strPath
   Else
      'GetFolder = Mid(strPath, 1, InStr(strPath, "\") - 1)
      Pos = 1
      While InStr(Pos, strPath, "\")
         Pos = InStr(Pos, strPath, "\") + 1
      Wend
      GetFolder = Mid(strPath, 1, Pos - 2)
   End If
End Function
Public Function GetExtension(strFileName As String) As String
   GetExtension = LCase$(Mid(strFileName, InStr(strFileName, ".") + 1, 3))
   'GetExtension = LCase$(Right$(strFileName, Len(strFileName) - InStr(strFileName, ".")))
End Function
