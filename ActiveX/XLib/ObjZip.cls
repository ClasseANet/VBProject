VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public Function Unzip(pPath As String, pFile As String, Optional pPathDest As String, Optional pHonorDir As Boolean = True) As Boolean
Public Sub Unzip(pPath As String, pFile As String, Optional pPathDest As String, Optional pHonorDir As Boolean = True)
   Dim sTag As String
   Dim sExe As String
   If Not L_Unzip(pPath, pFile, pPathDest, pHonorDir) Then
      If ClsAmbiente.ExisteArquivo(ClsAmbiente.ResolvePathName(pPath) & pFile) Then
         sTag = ClsGeneral.SetTag(sTag, "SUB", "UNZIP")
         sTag = ClsGeneral.SetTag(sTag, "PATH", pPath)
         sTag = ClsGeneral.SetTag(sTag, "ZIPFILE", pFile)
         sTag = ClsGeneral.SetTag(sTag, "PATHDEST", pPathDest)
         sTag = ClsGeneral.SetTag(sTag, "HONORDIR", pHonorDir)
         
         sExe = ClsAmbiente.ResolvePathName(ClsAmbiente.GetSpecialFolder(CSIDL_SYSTEM64))
         sExe = sExe & "ClasseA\"
         sExe = sExe & "CAZIP.exe "
         Call ClsAmbiente.SincShell(sExe & sTag, vbHide, True)
      End If
   End If
End Sub
Public Sub Zip(pFiles As String, pZipFile As String, Optional bExibeMsg As Boolean = True, Optional bSubescreveZipFile As Boolean = True)
   Dim sTag As String
   Dim sExe As String
   
   If bSubescreveZipFile Then
      If ClsAmbiente.ExisteArquivo(pZipFile) Then
         Call ClsAmbiente.ExcluirArquivo(pZipFile)
      End If
   End If
   If ClsAmbiente.ExisteArquivo(pFiles) Then
      sTag = ClsGeneral.SetTag(sTag, "SUB", "ZIP")
      sTag = ClsGeneral.SetTag(sTag, "FILES", pFiles)
      sTag = ClsGeneral.SetTag(sTag, "ZIPFILE", pZipFile)
      sTag = ClsGeneral.SetTag(sTag, "EXIBEMSG", bExibeMsg)
      
      sExe = ClsAmbiente.ResolvePathName(ClsAmbiente.GetSpecialFolder(CSIDL_SYSTEM64))
      sExe = sExe & "ClasseA\"
      sExe = sExe & "CAZIP.exe "
      
      Call ClsAmbiente.SincShell(sExe & sTag, vbHide, True)
   End If
End Sub
Private Function L_Unzip(pPath As String, pFile As String, Optional pPathDest As String, Optional pHonorDir As Boolean = True) As Boolean
   Dim oUnZip As CGUnzipFiles
   Dim pMessage As String
   Dim sError  As String

   On Error GoTo TrataErro
   'Call RegServer(App.Path & "\Unzip.dll")

   If Trim(pPathDest) = "" Then pPathDest = pPath
   Call ClsAmbiente.CriarDiretorio(pPathDest)

    Set oUnZip = New CGUnzipFiles
    With oUnZip
      .ZipFileName = ClsAmbiente.ResolvePathName(pPath) & pFile
      .ExtractDir = ClsAmbiente.ResolvePathName(pPathDest)
      .HonorDirectories = pHonorDir
      If .Unzip <> 0 Then
         pMessage = .GetLastMessage
         L_Unzip = False
      Else
         L_Unzip = True
      End If
    End With
    Set oUnZip = Nothing

    Exit Function

TrataErro:
   L_Unzip = False
   'sError = "CAZIPEXE.CLCAZIP.UnZip" & vbNewLine
   'sError = sError & "Error: " & Err.Number & " - " & Err.Description & vbNewLine
   'sError = sError & "Last Message: " & pMessage & vbNewLine
   'sError = sError & "File Unzip: " & ClsAmbiente.ResolvePathName(pPath) & pFile & vbNewLine
   'MsgBox sError
End Function

Private Sub Class_Initialize()
   Dim sZip As String
   Dim sUnZip As String
'Dim i As Long
'For i = 0 To 100
'   Debug.Print i & ": " & ClsAmbiente.GetSpecialFolder(i)
'Next
  Call SetDllSystem(ClsAmbiente.GetSpecialFolder(CSIDL_SYSTEM64) & "ClasseA\Unzip32.dll")
  Call SetDllSystem(ClsAmbiente.GetSpecialFolder(CSIDL_SYSTEM64) & "ClasseA\Zip32.dll")
  Call SetDllSystem(ClsAmbiente.GetSpecialFolder(CSIDL_SYSTEM32) & "ClasseA\Unzip32.dll")
  Call SetDllSystem(ClsAmbiente.GetSpecialFolder(CSIDL_SYSTEM32) & "ClasseA\Zip32.dll")
  Call SetDllSystem(App.Path & "\Unzip32.dll")
  Call SetDllSystem(App.Path & "\Zip32.dll")
End Sub
Private Sub SetDllSystem(pFile As String)
   If ClsAmbiente.ExisteArquivo(pFile) Then
      If UCase(Right(pFile, 12)) = "\UNZIP32.DLL" Then
         If FileLen(pFile) < 140000 Then
            Call ClsAmbiente.ExcluirArquivo(pFile, False)
         End If
      ElseIf UCase(Right(pFile, 10)) = "\ZIP32.DLL" Then
         If FileLen(pFile) > 134000 Then
            Call ClsAmbiente.ExcluirArquivo(pFile, False)
         End If
      End If
   End If
   If ClsAmbiente.ExisteArquivo(pFile) Then
      Call ClsAmbiente.RegServer(pFile, True, False)
   Else
      If UCase(Right(pFile, 12)) = "\UNZIP32.DLL" Then
         Call ClsGeneral.ExtractResData("UNZIP32", "DLL", pFile)
      ElseIf UCase(Right(pFile, 10)) = "\ZIP32.DLL" Then
         Call ClsGeneral.ExtractResData("ZIP32", "DLL", pFile)
      End If
      Call ClsAmbiente.RegServer(pFile, True, False)
   End If
End Sub
