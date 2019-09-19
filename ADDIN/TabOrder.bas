Attribute VB_Name = "modMain"
'Option Explicit
'
'
'Private Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal msg&, ByVal wp&, ByVal lp&)
'Private Declare Sub SetFocus Lib "user32" (ByVal hwnd&)
'Private Declare Function GetParent Lib "user32" (ByVal hwnd&) As Long
'Const WM_SYSKEYDOWN = &H104
'Const WM_SYSKEYUP = &H105
'Const WM_SYSCHAR = &H106
'Const VK_F = 70  ' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
'Dim hwndMenu       As Long           'needed to pass the menu keystrokes to VB
'
'Global gVBInstance  As VBIDE.VBE       'instance of VB IDE
'Global gWinWindow   As VBIDE.Window    'used to make sure we only run one instance
'Global gdocTabOrder As Object          'user doc object
'
'
'Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FILENAME$)
'Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"
'Sub AddToINI()
'  'this code adds an entry to VB5.INI
'  'this should be executed from the immediate window
'  Debug.Print WritePrivateProfileString("Add-Ins32", "TabOrder.Connect", "0", "vbaddin.ini")
'End Sub
'Function InRunMode(VBInst As VBIDE.VBE) As Boolean
'  InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
'End Function
'Sub HandleKeyDown(ud As Object, KeyCode As Integer, Shift As Integer)
'  If Shift <> 4 Then Exit Sub
'  If KeyCode < 65 Or KeyCode > 90 Then Exit Sub
'  If gVBInstance.DisplayModel = vbext_dm_SDI Then Exit Sub
'
'  If hwndMenu = 0 Then hwndMenu = FindHwndMenu(ud.hwnd)
'  PostMessage hwndMenu, WM_SYSKEYDOWN, KeyCode, &H20000000
'  KeyCode = 0
'  SetFocus hwndMenu
'End Sub
'Function FindHwndMenu&(ByVal hwnd&)
'  Dim h As Long
'
'Loop2:
'  h = GetParent(hwnd)
'  If h = 0 Then FindHwndMenu = hwnd: Exit Function
'  hwnd = h
'  GoTo Loop2
'End Function
'Public Function Copy(Orig$, dest$)
'   Dim nMsg$, nTipo&, NL
'   Dim Resp%
'   NL = vbLf
'   On Error Resume Next
'   If FileExists(Orig$) Then
'      Call Del(dest$)
'      FileCopy Orig$, dest$
'   Else
''      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(11) + UCase(Orig$), ClsMsg.LoadMsg(12))
'      Resp = vbCancel
'      Exit Function
'   End If
'   Resp = vbYes
'   Select Case Err
'      Case 71
'         While Resp = vbYes
''            nMsg = ClsMsg.LoadMsg(13) + NL + NL
''            nMsg = nMsg & ClsMsg.LoadMsg(14) + NL
''            nMsg = nMsg & ClsMsg.LoadMsg(15)
''            nTipo = vbYesNo + vbCritical + vbDefaultButton1
''            Resp = MsgBox(nMsg, nTipo, ClsMsg.LoadMsg(16))
'            If Resp = vbYes Then
'               Err = 0
'               FileCopy Orig$, dest$
'            End If
'         Wend
'      Case 70
'         While Resp = vbOK
''            nMsg = ClsMsg.LoadMsg(7) + NL + NL
''            nMsg = nMsg & ClsMsg.LoadMsg(56) + NL
'            nTipo = vbYesNo + vbCritical + vbDefaultButton1
''            Resp = MsgBox(nMsg, nTipo, ClsMsg.LoadMsg(16))
'            If Resp = vbYes Then
'               Err = 0
'               FileCopy Orig$, dest$
'            End If
'         Wend
'   End Select
'   Copy = Resp
'End Function
'
'Function FileExists(ByVal strPathName As String) As Boolean
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
'Public Sub Del(Arq$)
'   If FileExists(Arq$) Then
'      On Error GoTo Fim
'      Kill Arq$
'   End If
'   Exit Sub
'Fim:
''   ClsMsg.ShowError
'End Sub
'Public Function GetNameFromPath(PathFile As String, Optional PathNome = 2) As String
''================================================================
''= Última Alteração : 03/02/2000                                =
''= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
''================================================================
''****************************************************************
''**                                                            **
''** OBJETIVO : Recupera o nome do arquivo da "String de seu    **
''**            "Path".                                         **
''**                                                            **
''** Recebe: PathFile - Caminho Completo com nome do Arquivo.   **
''**                                                            **
''** Retorna : Nome do Arquivo.                                 **
''**                                                            **
''****************************************************************
'   Dim i%
'   While InStr(i + 1, PathFile, "\")
'      i = InStr(i + 1, PathFile, "\")
'   Wend
'   If PathNome = 1 Then
'      GetNameFromPath = VBA.Left$(PathFile, i)
'   ElseIf PathNome = 2 Then
'      GetNameFromPath = VBA.Mid$(PathFile, Len(VBA.Left$(PathFile, i)) + 1)
'   End If
'End Function
'
