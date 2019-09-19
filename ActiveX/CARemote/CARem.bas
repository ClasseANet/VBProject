Attribute VB_Name = "CARem"
Global xAmbiente  As Object 'XLib.Ambiente
Public Sub Main()
   Dim sPath As String
   Dim sFile As String
   Dim sParam As String
   Dim sCommand   As String
   
   sPath = ResolvePathName(App.Path)
   sFile = "TV.exe"
   sParam = " /run"
   sCommand = sPath & sFile & sParam
   Call ExtrairDependencia
   If Dir(App.Path & "/" & "TV.exe") <> "" Then
      Call SincShell(sCommand, vbNormalFocus, True)
      'Call ExcluirArquivo(sPath & sFile)
      'End
   End If
End Sub
Public Sub ExtrairDependencia()
   If Not ExisteArquivo(App.Path & "\TV.exe") Then
      Call ExtractResData("TV", "CUSTOM", App.Path & "\TV.exe")
   End If
End Sub
Public Function ExtractResData(Id, Tipo, Arquivo As String, Optional pFileBuf) As Boolean
   Dim nInt As Integer
   Dim byteFileBuf() As Byte 'This must be byte rather than String, so no Unicode conversion takes place
   Dim nVez As Integer
   Dim sPath   As String

   On Error GoTo Fim

   Call GetNameFromPath(Arquivo, sPath)
   If sPath <> "" Then
      Call CriarDiretorio(sPath)
   End If
   Call ExcluirArquivo(Arquivo, False)

   nInt = FreeFile
   Open Arquivo$ For Binary Access Write As nInt
      If IsMissing(pFileBuf) Then
         byteFileBuf = LoadResData(Id, Tipo)
      End If
      Put nInt, , byteFileBuf
   GoTo Saida
Fim:
   nVez = nVez + 1
   If nVez < 5 Then
      Resume
   Else
      Resume Next
   End If

Saida:
    Close nInt
    Err = 0
    ExtractResData = ExisteArquivo(Arquivo$)
    Exit Function
End Function
Public Sub SincShell(Comando As String, Optional Modo As VbAppWinStyle = vbMaximizedFocus, Optional EsperaProcesso = True)
   If xAmbiente Is Nothing Then Set xAmbiente = CreateObject("xLIB.Ambiente")
   Call xAmbiente.SincShell(Comando, Modo, EsperaProcesso)
End Sub
Public Function GetNameFromPath(PathFile As String, Optional ByRef PathReturn As String) As String
   Dim i As Integer

   i = InStrRev(PathFile, "\")
   'i = IIf(i = 0, 1, i)
   If PathReturn = "1" Then
      GetNameFromPath = VBA.Left$(PathFile, i)
   Else
      GetNameFromPath = VBA.Mid$(PathFile, Len(VBA.Left$(PathFile, i)) + 1)
   End If
   PathReturn = ResolvePathName(VBA.Left$(PathFile, i))
End Function
Public Function CriarDiretorio(pPath As String, Optional bViewMsg As Boolean = False) As Boolean
   Dim sAux     As String
   Dim sPath    As String
   Dim sOldPath As String
   Dim sIniPath As String
   Dim PosIni   As Integer
   Dim PosAux   As Integer
   Dim nResp    As Integer
   Dim bResult  As Boolean

   On Error Resume Next

   sOldPath = CurDir$

   pPath = ResolvePathName(pPath)
   bResult = (Trim(Dir(pPath, vbDirectory)) <> "")

   If Not bResult Then
      PosIni = InStr(1, pPath, "\")
      While PosIni <> 0
         PosIni = InStr(PosIni + 1, pPath, "\")
         If PosIni > 0 Then
            sPath = Left$(pPath, PosIni - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir sPath
            If Err Then
               ' We must create this directory
               Err = 0
               #If Win32 And LOGGING Then
                  NewAction gstrKEY_CREATEDIR, """" & sPath & """"
               #End If
               MkDir sPath
               #If Win32 And LOGGING Then
                  If Err Then
                     LogError ResolveResString(resMAKEDIR) & " " & sPath
                     AbortAction
                     GoTo Done
                  Else
                     CommitAction
                  End If
               #End If
            End If
         End If
      Wend
      bResult = (Trim(Dir(pPath, vbDirectory)) <> "")
   End If

   ChDir sOldPath

   If Trim(Dir(pPath, vbDirectory)) = "" Then
      nResp = MsgBox("Erro ao criar pasta:" & pPath, vbRetryCancel Or vbExclamation Or vbDefaultButton2, "Ambiente")
      Select Case nResp
         Case vbIgnore, vbAbort, vbCancel
            bResult = False
         Case vbRetry
            bResult = CriarDiretorio(pPath, bViewMsg)
      End Select
   End If
   CriarDiretorio = bResult
   Err = 0
End Function
Public Function ExcluirArquivo(File As String, Optional ViewError As Boolean = True) As Boolean
   If ExisteArquivo(File) Then
      On Error GoTo Fim
      Call Kill(File)
   End If
   ExcluirArquivo = Not ExisteArquivo(File)
   Exit Function
Fim:
   If ViewError Then
      ClsMensagem.ExibirErro
   End If
End Function
Public Function ExisteArquivo(ByVal strPathName As String) As Boolean
   Dim intFileNum   As Integer
   Dim sArq         As String
   Dim sPath        As String

   On Error Resume Next

   strPathName = Trim(strPathName)
   strPathName = Replace(strPathName, """", "")

   Call GetNameFromPath(strPathName, sPath)
   sArq = Mid(strPathName, Len(sPath) + 1)
   If Len(Dir(strPathName, vbArchive)) > 4 And sPath <> "" And sArq <> "" Then
      ExisteArquivo = IIf(Err = 0, True, False)
   Else
      If Right$(strPathName, 1) = "\" Then
          strPathName = VBA.Left$(strPathName, Len(strPathName) - 1)
      End If
      '
      'Attempt to open the file, return value of this function is False
      'if an error occurs on open, True otherwise
      '
      intFileNum = FreeFile
      Open strPathName For Input As intFileNum
      ExisteArquivo = IIf(Err = 0, True, False)
      Close intFileNum
   End If

   Err = 0
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
