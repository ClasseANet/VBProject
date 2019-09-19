Attribute VB_Name = "Logon"
Option Explicit
Private Const LOGON32_LOGON_INTERACTIVE = 2
Private Const LOGON32_LOGON_NETWORK = 3
Private Const LOGON32_LOGON_BATCH = 4
Private Const LOGON32_LOGON_SERVICE = 5

Private Const LOGON32_PROVIDER_DEFAULT = 0
Private Const LOGON32_PROVIDER_WINNT35 = 1
Private Const LOGON32_PROVIDER_WINNT40 = 2
Private Const LOGON32_PROVIDER_WINNT50 = 3

Private Const CREATE_DEFAULT_ERROR_MODE = &H4000000
Private Const LOGON_WITH_PROFILE = &H1
Private Const LOGON_NETCREDENTIALS_ONLY = &H2

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long ' !!! must be Long for Unicode string
    lpDesktop As Long  ' !!! must be Long for Unicode string
    lpTitle As Long    ' !!! must be Long for Unicode string
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

'  LogonUser() requires that the caller has the following permission
'  Permission                        Display Name
'  --------------------------------------------------------------------
'  SE_TCB_NAME                      Act as part of the operating system

'  CreateProcessAsUser() requires that the caller has the following permissions
'  Permission                        Display Name
'  ---------------------------------------------------------------
'  SE_ASSIGNPRIMARYTOKEN_NAME       Replace a process level token
'  SE_INCREASE_QUOTA_NAME           Increase quotas
  
Private Declare Function LogonUser Lib "ADVAPI32.DLL" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function CreateProcessAsUser Lib "ADVAPI32.DLL" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
' CreateProcessWithLogonW API is available only on Windows 2000 and later.
Private Declare Function CreateProcessWithLogonW Lib "ADVAPI32.DLL" (ByVal lpUsername As String, ByVal lpDomain As String, ByVal lpPassword As String, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32.dll" (ByVal uMode As Long) As Long
        
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
                             
' Version Checking APIs
Private Declare Function GetVersionExA Lib "kernel32.dll" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Const VER_PLATFORM_WIN32_NT = &H2

'********************************************************************

'                   RunAsUser for Windows 2000 and Later
'********************************************************************
Public Function W2KRunAsUser(ByVal UserName As String, ByVal Password As String, ByVal DomainName As String, ByVal CommandLine As String, ByVal CurrentDirectory As String) As Long
   Dim si As STARTUPINFO
   Dim pi As PROCESS_INFORMATION
   
   Dim wUser As String
   Dim wDomain As String
   Dim wPassword As String
   Dim wCommandLine As String
   Dim wCurrentDir As String
   
   Dim Result As Long
   
   si.cb = Len(si)
       
   wUser = StrConv(UserName + Chr$(0), vbUnicode)
   wDomain = StrConv(DomainName + Chr$(0), vbUnicode)
   wPassword = StrConv(Password + Chr$(0), vbUnicode)
   wCommandLine = StrConv(CommandLine + Chr$(0), vbUnicode)
   wCurrentDir = StrConv(CurrentDirectory + Chr$(0), vbUnicode)
   
   Result = CreateProcessWithLogonW(wUser, wDomain, wPassword, LOGON_WITH_PROFILE, 0&, wCommandLine, CREATE_DEFAULT_ERROR_MODE, 0&, wCurrentDir, si, pi)
   ' CreateProcessWithLogonW() does not
   If Result <> 0 Then
      CloseHandle pi.hThread
      CloseHandle pi.hProcess
      W2KRunAsUser = 0
   Else
      W2KRunAsUser = Err.LastDllError
      'MsgBox "CreateProcessWithLogonW() failed with error " & Err.LastDllError, vbExclamation
   End If
End Function
'********************************************************************
'                   RunAsUser for Windows NT 4.0
'********************************************************************
Public Function NT4RunAsUser(ByVal UserName As String, ByVal Password As String, ByVal DomainName As String, ByVal CommandLine As String, ByVal CurrentDirectory As String) As Long
   Dim Result As Long
   Dim hToken As Long
   Dim si As STARTUPINFO
   Dim pi As PROCESS_INFORMATION
   
   Result = LogonUser(UserName, DomainName, Password, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, hToken)
   If Result = 0 Then
      NT4RunAsUser = Err.LastDllError
      ' LogonUser will fail with 1314 error code, if the user account associated
      ' with the calling security context does not have
      ' "Act as part of the operating system" permission
      MsgBox " () failed with error " & Err.LastDllError, vbExclamation
      Exit Function
   End If
   
   si.cb = Len(si)
   Result = CreateProcessAsUser(hToken, 0&, CommandLine, 0&, 0&, False, CREATE_DEFAULT_ERROR_MODE, 0&, CurrentDirectory, si, pi)
   If Result = 0 Then
      NT4RunAsUser = Err.LastDllError
      ' CreateProcessAsUser will fail with 1314 error code, if the user
      ' account associated with the calling security context does not have
      ' the following two permissions
      ' "Replace a process level token"
      ' "Increase Quotoas"
      MsgBox "CreateProcessAsUser() failed with error " & Err.LastDllError, vbExclamation
      CloseHandle hToken
      Exit Function
   End If
   
   CloseHandle hToken
   CloseHandle pi.hThread
   CloseHandle pi.hProcess
   NT4RunAsUser = 0
End Function
'***********************
'* http://www.andreavb.com/forum/viewtopic_4543.html
'* RunAsUser UserNamet, Password, DomainName, CommandLine, CurrentDirectory
Public Function RunAsUser(ByVal UserName As String, ByVal Password As String, ByVal DomainName As String, ByVal CommandLine As String, ByVal CurrentDirectory As String) As Long

   Dim w2kOrAbove As Boolean
   Dim osinfo As OSVERSIONINFO
   Dim Result As Long
   Dim uErrorMode As Long
    
   ' Determine if system is Windows 2000 or later
   osinfo.dwOSVersionInfoSize = Len(osinfo)
   osinfo.szCSDVersion = Space$(128)
   GetVersionExA osinfo
   w2kOrAbove = (osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT And osinfo.dwMajorVersion >= 5)
   If (w2kOrAbove) Then
       Result = W2KRunAsUser(UserName, Password, DomainName, CommandLine, CurrentDirectory)
   Else
       Result = NT4RunAsUser(UserName, Password, DomainName, CommandLine, CurrentDirectory)
   End If
   RunAsUser = Result
End Function

