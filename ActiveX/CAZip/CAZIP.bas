Attribute VB_Name = "BLCAZIP"
Option Explicit
Sub Main()
   Dim sCommand As String
   Dim sSub       As String
   Dim pFiles     As String
   Dim pZipFile   As String
   Dim pExibeMsg  As Boolean
   Dim pPath      As String
   Dim pPathDest  As String
   Dim pHonorDir  As Boolean
   
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   
   sCommand = Trim(UCase(Command$))
   'sCommand = "|SUB=UNZIP|PATH=C:\Arquivos de programas\ClasseA\Admin\Dll\|ZIPFILE=Unzip32.dll.zip|PATHDEST=C:\Arquivos de programas\ClasseA\Admin\Dll\|HONORDIR=False"
   
   sSub = GetTag(sCommand, "SUB")
   
   
   If sSub = "ZIP" Then
      pFiles = GetTag(sCommand, "FILES")
      pZipFile = GetTag(sCommand, "ZIPFILE")
      pExibeMsg = GetTag(sCommand, "EXIBEMSG")
      
      Call Zip(pFiles, pZipFile, pExibeMsg)
   ElseIf sSub = "UNZIP" Then
      pPath = GetTag(sCommand, "PATH")
      pZipFile = GetTag(sCommand, "ZIPFILE")
      pPathDest = GetTag(sCommand, "PATHDEST")
      pHonorDir = GetTag(sCommand, "HONORDIR", True)
      
      If Unzip(pPath, pZipFile, pPathDest, pHonorDir) Then
         'Command$ = SetTag(Command$, "RESULT", True)
      Else
         'Command$ = SetTag(Command$, "RESULT", False)
      End If
   End If
   Screen.MousePointer = vbDefault
TrataErro:
End Sub
Public Sub Zip(pFiles As String, pZipFile As String, Optional bExibeMsg As Boolean = True)
   Dim oZip As CGZipFiles
   Dim sFile

   On Error GoTo TrataErro

   Set oZip = New CGZipFiles
   With oZip
      .ZipFileName = pZipFile                   '"\ZIPTEST.ZIP"
      .UpdatingZip = False                      ' ensures a new zip is created. Are we updating a Zip File ? - This doesn't seem to work - check InfoZip homepage for more info.

      If TypeName(pFiles) = "Collection" Then   'App.Path & "\*.*" Add in the files to the zip - in this case, we want all the ones in the current directory
         'For Each sFile In pFiles
         '   .AddFile sFile
         'Next
      ElseIf TypeName(pFiles) = "String" Then
         .AddFile pFiles
      End If
      If .MakeZipFile <> 0 Then
         If bExibeMsg Then MsgBox .GetLastMessage ' any errors
      End If
    End With
    Set oZip = Nothing

    Exit Sub

TrataErro:
    MsgBox Err.Number & " " & "Form1::cmdZip_Click" & " " & Err.Description
End Sub
Public Function Unzip(pPath As String, pFile As String, Optional pPathDest As String, Optional pHonorDir As Boolean = True) As Boolean
   Dim oUnZip As CGUnzipFiles
   Dim pMessage As String
   Dim sError  As String

   On Error GoTo TrataErro
   'Call RegServer(App.Path & "\Unzip.dll")

   If Trim(pPathDest) = "" Then pPathDest = pPath
   Call CriarDiretorio(pPathDest)

    Set oUnZip = New CGUnzipFiles
    With oUnZip
      .ZipFileName = ResolvePathName(pPath) & pFile
      .ExtractDir = ResolvePathName(pPathDest)
      .HonorDirectories = pHonorDir
      If .Unzip <> 0 Then
         pMessage = .GetLastMessage
         Unzip = False
      Else
         Unzip = True
      End If
    End With
    Set oUnZip = Nothing

    Exit Function

TrataErro:
   sError = "CAZIPEXE.CLCAZIP.UnZip" & vbNewLine
   sError = sError & "Error: " & Err.Number & " - " & Err.Description & vbNewLine
   sError = sError & "Last Message: " & pMessage & vbNewLine
   sError = sError & "File Unzip: " & ResolvePathName(pPath) & pFile & vbNewLine
   MsgBox sError
End Function

