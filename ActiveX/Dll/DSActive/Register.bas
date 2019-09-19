Attribute VB_Name = "Register"
Option Explicit

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long '(lpThreadAttributes As SECURITY_ATTRIBUTES,
Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'* dwCreationFlags param, call ResumeThread to
'* wake the thread up, specify 0 for an alive thread
Public Const CREATE_SUSPENDED = &H4
'*  dwMilliseconds param, specify 0 for immediate return.
Public Const INFINITE = &HFFFFFFFF   ' Infinite timeout
'*  WaitForSingleObject rtn vals
Public Const STATUS_WAIT_0 = &H0
Public Const STATUS_ABANDONED_WAIT_0 = &H80
Public Const STATUS_TIMEOUT = &H102
Public Const WAIT_FAILED = &HFFFFFFFF '* The state of the specified object is signaled (success)
'Public Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Public Const WAIT_ABANDONED = ((STATUS_ABANDONED_WAIT_0) + 0) '* Thread went away before the mutex got signaled
Public Const WAIT_TIMEOUT = STATUS_TIMEOUT '* dwMilliseconds timed out
Private Const STATUS_PENDING = &H103
Private Const STILL_ACTIVE = STATUS_PENDING
Public Function ProcurarArquivo(CmD As Object, Optional cDialogTitle = "Find File", Optional cfilename = "", Optional cFilter = "*.*", Optional cFilterIndex = 1)
   Dim LenP%, LenF%
       
   On Error GoTo OpenError
   
   With CmD
      .DialogTitle = cDialogTitle
      .FileName = cfilename
      .Filter = cFilter ' "Access Files (*.mdb)|*.mdb"
      .FilterIndex = cFilterIndex
      .Tag = ""
      .CancelError = True
      .Flags = FileOpenConstants.cdlOFNFileMustExist
      .ShowOpen
      LenP% = Len(.FileName)
      LenF% = Len(.FileTitle)
      ProcurarArquivo = UCase(.FileTitle)
      .Tag = UCase(Mid(.FileName, 1, LenP% - LenF%))
   End With
   Exit Function
OpenError:
Screen.MousePointer = vbDefault
CmD.FileName = ""
If Err = 3049 Then
  If MsgBox(Error & vbLf & vbLf & "Attempt to Repair it?", 4 + 48) = vbYes Then
'      Resume AttemptRepair
  End If
End If

If Err <> 32755 And Err <> 3049 Then   'check for common dialog cancelled
'    ShowError
End If

End Function
Function RegCloseKey(ByVal hKey As Long) As Boolean
    Dim lResult As Long

    On Error GoTo 0
    lResult = OSRegCloseKey(hKey)
    RegCloseKey = (lResult = ERROR_SUCCESS)
End Function


