VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOS"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private pConcatFile
Private pDrvTmp
Public Property Get ConcatFile() As String
    ConcatFile = pConcatFile
End Property
Public Property Let ConcatFile(Txt$)
    pConcatFile = Txt$
End Property
Public Property Get DrvTmp() As String
    DrvTmp = pDrvTmp
End Property
Public Property Let DrvTmp(Txt$)
    pDrvTmp = Txt$
End Property
Public Function AbrirTxt%(Arq$)
     Dim Hnd%
  On Error GoTo CopyErr
   Call Del(Arq$)
   AbrirTxt% = FreeFile()
   Open Arq$ For Output As #AbrirTxt%
Exit Function
CopyErr:
  Select Case Err
     Case 55: Err = 0
     Case Else: ClsMsg.ShowError
  End Select
End Function
Public Function Copy(Orig As String, Dest As String) As Boolean
   Dim nMsg       As String
   Dim nTipoBox   As Long
   Dim Resp       As Integer

   On Error Resume Next
   
   If FileExists(Orig) Then
      If FileExists(Dest) Then
         Call Del(Dest)
      Else
         Call MakePath(ClsDsr.GetNameFromPath(Dest, 1))
      End If
      FileCopy Orig, Dest

   Else
      Call ClsMsg.ExibirAviso("Arquivo não encontrado: " + UCase(Orig), "Importação")
      Resp = vbCancel
      Exit Function
   End If
   
   
   Resp = vbYes
   Select Case Err
      Case 71
         While Resp = vbYes
            nTipoBox = vbYesNo + vbCritical + vbDefaultButton1
            nMsg = "Drive ou arquivo inválido" + vbNewLine + vbNewLine
            nMsg = nMsg & "Insira um disco no drive ou verifieu o arquivo." + vbNewLine
            nMsg = nMsg & "Deseja continuar?"
            Resp = MsgBox(nMsg, nTipoBox, "Erro!")
            If Resp = vbYes Then
               Err = 0
               FileCopy Orig, Dest
            End If
         Wend
      Case 70
         While Resp = vbOK
            nTipoBox = vbOK + vbCritical + vbDefaultButton1
            nMsg = "Usuário não tem permissão a esta operação." + vbNewLine + vbNewLine
            nMsg = nMsg & "Algum recurso está compartilhando esta informação." + vbNewLine
            Resp = MsgBox(nTipoBox, nTipoBox, "Erro!")
            If Resp = vbYes Then
               Err = 0
               FileCopy Orig, Dest
            End If
         Wend
   End Select
   Copy = (Resp = vbYes)
End Function
Public Function CopyByte%(ByVal SrcName$, Optional ByVal Tam)
    Const lMAXCOPYBUF& = 64512
    Const lMINCOPYBUFSIZE& = 4096
    Const intOPEN% = 1
    Const intGET% = 2
    Const intPUT% = 3
    Const intMEMFAIL% = 4

    Dim intSrcFile As Integer
    Dim intStatus As Integer
    Dim lBytesLeftToWrite As Long
    Dim lBytesThisTime As Long
    Dim byteFileBuf() As Byte 'This must be byte rather than String, so no Unicode conversion takes place
    Dim strMsg As String
    
    Dim QtdRead&, Vol%, VolFile$, mIntConcatFile%

    On Error GoTo CSFError
    
    If IsMissing(Tam) Then Tam = (22 * 64512) '=1.419.264
    '
    'Ensure that the specified source file is available
    '
    If DetectFile(SrcName$) = vbIgnore Then
        CopyByte = vbIgnore
        Exit Function
    End If

    lBytesLeftToWrite = FileLen(SrcName$)
    
    '
    'For error reporting, flag that we're attempting to open the file now
    '
    intStatus = intOPEN

    '
    'Open the source file for reading now
    '
    intSrcFile = FreeFile
    Open SrcName$ For Binary Access Read As intSrcFile
    
    Vol% = 1
    VolFile = Mid(SrcName$, 1, Len(SrcName$) - 2) & Trim(ClsDsr.StrZero(Vol%, 2))
    Call Del(VolFile$)
    mIntConcatFile = FreeFile
    Open VolFile$ For Binary Access Write As mIntConcatFile

    '
    'Initially, we'll try to copy lMAXCOPYBUF bytes at a time.  If our attempt
    'to allocate a copy buffer (Space$(...)) fails, the error handling logic
    'will cause the buffer size to be halved and another allocation attempt to
    'be made.
    '
    lBytesThisTime = lMAXCOPYBUF
    ReDim byteFileBuf(1 To lBytesThisTime) As Byte

    While (lBytesLeftToWrite <> 0)
        '
        'while source file hasn't been read, if the number of bytes left is bigger than
        'the buffer size, reduce the buffer size
        '
        If lBytesThisTime > lBytesLeftToWrite Then
            lBytesThisTime = lBytesLeftToWrite
            ReDim byteFileBuf(1 To lBytesThisTime) As Byte
        End If
        '
        'Set operation status and Get from the source file and Put to the dest file
        '
        intStatus = intGET
        Get intSrcFile, , byteFileBuf
                
        QtdRead = QtdRead + lMAXCOPYBUF
        If QtdRead > Tam Then
           Close mIntConcatFile
           Vol = Vol + 1
           VolFile$ = Mid(SrcName$, 1, Len(SrcName$) - 2) & Trim(ClsDsr.StrZero(Vol%, 2))
           Call Del(VolFile$)
           mIntConcatFile = FreeFile
           Open VolFile$ For Binary Access Write As mIntConcatFile
           QtdRead = lMAXCOPYBUF
        End If
        intStatus = intPUT
        Put mIntConcatFile, , byteFileBuf

        lBytesLeftToWrite = lBytesLeftToWrite - lBytesThisTime
    Wend

    CopyByte = True
    GoTo CSFCleanup

CSFError:
    If Err = 14 Then    'Out of String Space
        lBytesThisTime = lBytesThisTime \ 2
        If lBytesThisTime >= lMINCOPYBUFSIZE Then
            Resume
        Else
            intStatus = intMEMFAIL
        End If
    End If

    strMsg = vbLf & SrcName$

    Select Case intStatus
    Case intOPEN
'        strMsg = ResolveResString(resCANTOPEN) & strMsg
    Case intGET
'        strMsg = ResolveResString(resCANTREAD) & strMsg
    Case intPUT
'        strMsg = ResolveResString(resCANTWRITE) & strMsg & LS$ & ResolveResString(resCHKSPACE)
    Case intMEMFAIL
'        strMsg = ResolveResString(resOUTOFMEMORY) & strMsg
    End Select

'    Select Case MsgError(Error$ & NL & strMsg, vbAbortRetryIgnore Or vbExclamation Or vbDefaultButton2, "Setup Message")
'    Case vbAbort
'        ExitSetup frmCopy, gintRET_ABORT
'    Case vbIGNORE
'        ConcatSplitFile = vbIGNORE
'    Case vbRetry
'        Resume
'    End Select

CSFCleanup:
    Close mIntConcatFile
    Close intSrcFile
    Err = 0
    Exit Function
End Function
'-----------------------------------------------------------
' FUNCTION: ConcatSplitFile
'
' Reads and appends the source file passed in onto the
' previously opened destination file specified by
' mintConcatFile.  mintConcatFile should be opened
' by calling OpenConcatFile() before calling this function.
'
' IN: [strSrcName] - Source file to append to destination
'
' Returns: True if copy was successful, vbIGNORE if user
'          elects to ignore a reported copy error
'-----------------------------------------------------------
'
Function ConcatSplitFile(ByVal strSrcName As String, mIntConcatFile%) As Integer
    Const lMAXCOPYBUF& = 64512
    Const lMINCOPYBUFSIZE& = 4096
    Const intOPEN% = 1
    Const intGET% = 2
    Const intPUT% = 3
    Const intMEMFAIL% = 4

    Dim intSrcFile As Integer
    Dim intStatus As Integer
    Dim lBytesLeftToWrite As Long
    Dim lBytesThisTime As Long
    Dim byteFileBuf() As Byte 'This must be byte rather than String, so no Unicode conversion takes place
    Dim strMsg As String

    On Error GoTo CSFError
    
    '
    'Ensure that the specified source file is available
    '
    If DetectFile(strSrcName) = vbIgnore Then
        ConcatSplitFile = vbIgnore
        Exit Function
    End If

    lBytesLeftToWrite = FileLen(strSrcName)

    '
    'For error reporting, flag that we're attempting to open the file now
    '
    intStatus = intOPEN

    '
    'Open the source file for reading now
    '
    intSrcFile = FreeFile
    Open strSrcName For Binary Access Read As intSrcFile

    '
    'Initially, we'll try to copy lMAXCOPYBUF bytes at a time.  If our attempt
    'to allocate a copy buffer (Space$(...)) fails, the error handling logic
    'will cause the buffer size to be halved and another allocation attempt to
    'be made.
    '
    lBytesThisTime = lMAXCOPYBUF
    ReDim byteFileBuf(1 To lBytesThisTime) As Byte

    While lBytesLeftToWrite <> 0
        '
        'while source file hasn't been read, if the number of bytes left is bigger than
        'the buffer size, reduce the buffer size
        '
        If lBytesThisTime > lBytesLeftToWrite Then
            lBytesThisTime = lBytesLeftToWrite
            ReDim byteFileBuf(1 To lBytesThisTime) As Byte
        End If
        '
        'Set operation status and Get from the source file and Put to the dest file
        '
        intStatus = intGET
        Get intSrcFile, , byteFileBuf

        intStatus = intPUT
        Put mIntConcatFile, , byteFileBuf

        lBytesLeftToWrite = lBytesLeftToWrite - lBytesThisTime
    Wend

    ConcatSplitFile = True
    GoTo CSFCleanup

CSFError:
    If Err = 14 Then    'Out of String Space
        lBytesThisTime = lBytesThisTime \ 2
        If lBytesThisTime >= lMINCOPYBUFSIZE Then
            Resume
        Else
            intStatus = intMEMFAIL
        End If
    End If

    strMsg = vbLf & strSrcName

    Select Case intStatus
    Case intOPEN
'        strMsg = ResolveResString(resCANTOPEN) & strMsg
    Case intGET
'        strMsg = ResolveResString(resCANTREAD) & strMsg
    Case intPUT
'        strMsg = ResolveResString(resCANTWRITE) & strMsg & LS$ & ResolveResString(resCHKSPACE)
    Case intMEMFAIL
'        strMsg = ResolveResString(resOUTOFMEMORY) & strMsg
    End Select

'    Select Case MsgError(Error$ & NL & strMsg, vbAbortRetryIgnore Or vbExclamation Or vbDefaultButton2, "Setup Message")
'    Case vbAbort
'        ExitSetup frmCopy, gintRET_ABORT
'    Case vbIGNORE
'        ConcatSplitFile = vbIGNORE
'    Case vbRetry
'        Resume
'    End Select

CSFCleanup:
    Close intSrcFile
    Err = 0
    Exit Function
End Function
Public Sub Del(File As String, Optional ViewError As Boolean = True)
   If FileExists(File) Then
      On Error GoTo Fim
      Call Kill(File)
   End If
   Exit Sub
Fim:
   If ViewError Then
      ClsMsg.ShowError
   End If
End Sub
Public Sub DelDir(pPath As String)
   On Error GoTo TrataErro
   
   If Right(pPath, 1) <> "\" Then
      pPath = pPath & "\"
   End If
   While Dir(pPath) <> ""
      Call Del(pPath & Dir(pPath))
   Wend
   RmDir pPath
   Exit Sub
TrataErro:
   If Err <> 76 Then '* Path not found
      Call ClsMsg.ShowError
   End If
End Sub
Function DetectFile(ByVal strFilename As String) As Integer
'-----------------------------------------------------------
' FUNCTION: DetectFile
'
' Detects whether the specified file exists.  If it can't
' be found, the user is given the opportunity to abort,
' retry, or ignore finding the file.  This call is used,
' for example, to ensure that a floppy with the specified
' file name is in the drive before continuing.
'
' IN: [strFileName] - name of file to detect, usually
'                     should include full path, Example:
'                     'A:\MYAPP.DAT'
'
' Returns: TRUE if the file was detected, vbIGNORE if
'          the user chose ignore when the file couldn't
'          be found, or calls ExitSetup upon 'Abort'
'-----------------------------------------------------------
'
    Dim strMsg As String

    DetectFile = True
    Do While FileExists(strFilename) = False
        strMsg = ClsMsg.LoadMsg(11) & strFilename
        Select Case MsgBox(strMsg, vbAbortRetryIgnore Or vbExclamation Or vbDefaultButton2, "System Message")
        Case vbAbort
            DetectFile = vbAbort
            'ExitSetup frmCopy, gintRET_ABORT
        Case vbIgnore
            DetectFile = vbIgnore
            Exit Do
        End Select
    Loop
End Function
'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err, False, True)
    FileExists = IIf(Err = 70, True, FileExists)

    Close intFileNum

    Err = 0
End Function
Public Sub FecharTxt(Arq%)
      Close #Arq%
End Sub
'-----------------------------------------------------------
' FUNCTION: MakePath
'
' Creates the specified directory path
'
' IN: [strDirName] - name of the dir path to make
'     [fAllowIgnore] - whether or not to allow the user to
'                      ignore any encountered errors.  If
'                      false, the function only returns
'                      if successful.  If missing, this
'                      defaults to True.
'
' Returns: True if successful, False if error and the user
'          chose to ignore.  (The function does not return
'          if the user selects ABORT/CANCEL on an error.)
'-----------------------------------------------------------
'
Public Function MakePath(ByVal strDir As String, Optional ByVal fAllowIgnore) As Boolean
    If IsMissing(fAllowIgnore) Then
        fAllowIgnore = True
    End If
    
    Do
        If MakePathAux(strDir) Then
            MakePath = True
            Exit Function
        Else
            Dim strMsg As String
            Dim iRet As Integer
            
'            strMsg = ResolveResString(resMAKEDIR) & LF$ & strDir
            iRet = MsgBox(strMsg, IIf(fAllowIgnore, vbAbortRetryIgnore, vbRetryCancel) Or vbExclamation Or vbDefaultButton2, "")
            Select Case iRet
            Case vbAbort, vbCancel
'                ExitSetup frmCopy, gintRET_ABORT
            Case vbIgnore
                MakePath = False
                Exit Function
            Case vbRetry
            End Select
        End If
    Loop
End Function
'-----------------------------------------------------------
' FUNCTION: MakePathAux
'
' Creates the specified directory path.
'
' No user interaction occurs if an error is encountered.
' If user interaction is desired, use the related
'   MakePathAux() function.
'
' IN: [strDirName] - name of the dir path to make
'
' Returns: True if successful, False if error.
'-----------------------------------------------------------
'
Private Function MakePathAux(ByVal strDirName As String) As Boolean
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Add trailing backslash
    '
   If Right$(strDirName, 1) <> "\" Then
        strDirName = strDirName & "\"
    End If

    strOldPath = CurDir$
    MakePathAux = False
    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.
    '
    '
    intOffset = InStr(intAnchor + 1, strDirName, "\")
    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, "\")
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir strPath
            If Err Then
                ' We must create this directory
                Err = 0
                #If Win32 And LOGGING Then
                    NewAction gstrKEY_CREATEDIR, """" & strPath & """"
                #End If
                MkDir strPath
                #If Win32 And LOGGING Then
                    If Err Then
                        LogError ResolveResString(resMAKEDIR) & " " & strPath
                        AbortAction
                        GoTo Done
                    Else
                        CommitAction
                    End If
                #End If
            End If
        End If
    Loop Until intAnchor = 0

    MakePathAux = True
Done:
    ChDir strOldPath

    Err = 0
End Function
Public Function MultVolume%(ByVal strSrcName As String, Optional ByVal Tam)
    Const lMAXCOPYBUF& = 64512
    Const lMINCOPYBUFSIZE& = 4096
    Const intOPEN% = 1
    Const intGET% = 2
    Const intPUT% = 3
    Const intMEMFAIL% = 4

    Dim intSrcFile As Integer
    Dim intStatus As Integer
    Dim lBytesLeftToWrite As Long
    Dim lBytesThisTime As Long
    Dim byteFileBuf() As Byte 'This must be byte rather than String, so no Unicode conversion takes place
    Dim strMsg As String
    
    Dim QtdRead&, Vol%, VolFile$, mIntConcatFile%

    On Error GoTo CSFError
    
    If IsMissing(Tam) Then Tam = (22 * 64512) '=1.419.264
    '
    'Ensure that the specified source file is available
    '
    If DetectFile(strSrcName) = vbIgnore Then
        MultVolume = vbIgnore
        Exit Function
    End If

    lBytesLeftToWrite = FileLen(strSrcName)

    '
    'For error reporting, flag that we're attempting to open the file now
    '
    intStatus = intOPEN

    '
    'Open the source file for reading now
    '
    intSrcFile = FreeFile
    Open strSrcName For Binary Access Read As intSrcFile
    
    Vol% = 1
    VolFile = Mid(strSrcName, 1, Len(strSrcName) - 2) & Trim(ClsDsr.StrZero(Vol%, 2))
    Call Del(VolFile$)
    mIntConcatFile = FreeFile
    Open VolFile$ For Binary Access Write As mIntConcatFile

    '
    'Initially, we'll try to copy lMAXCOPYBUF bytes at a time.  If our attempt
    'to allocate a copy buffer (Space$(...)) fails, the error handling logic
    'will cause the buffer size to be halved and another allocation attempt to
    'be made.
    '
    lBytesThisTime = lMAXCOPYBUF
    ReDim byteFileBuf(1 To lBytesThisTime) As Byte

    While (lBytesLeftToWrite <> 0)
        '
        'while source file hasn't been read, if the number of bytes left is bigger than
        'the buffer size, reduce the buffer size
        '
        If lBytesThisTime > lBytesLeftToWrite Then
            lBytesThisTime = lBytesLeftToWrite
            ReDim byteFileBuf(1 To lBytesThisTime) As Byte
        End If
        '
        'Set operation status and Get from the source file and Put to the dest file
        '
        intStatus = intGET
        Get intSrcFile, , byteFileBuf
        
        QtdRead = QtdRead + lMAXCOPYBUF
        If QtdRead > Tam Then
           Close mIntConcatFile
           Vol = Vol + 1
           VolFile$ = Mid(strSrcName, 1, Len(strSrcName) - 2) & Trim(ClsDsr.StrZero(Vol%, 2))
           Call Del(VolFile$)
           mIntConcatFile = FreeFile
           Open VolFile$ For Binary Access Write As mIntConcatFile
           QtdRead = lMAXCOPYBUF
        End If
        intStatus = intPUT
        Put mIntConcatFile, , byteFileBuf

        lBytesLeftToWrite = lBytesLeftToWrite - lBytesThisTime
    Wend

    MultVolume = True
    GoTo CSFCleanup

CSFError:
    If Err = 14 Then    'Out of String Space
        lBytesThisTime = lBytesThisTime \ 2
        If lBytesThisTime >= lMINCOPYBUFSIZE Then
            Resume
        Else
            intStatus = intMEMFAIL
        End If
    End If

    strMsg = vbLf & strSrcName

    Select Case intStatus
    Case intOPEN
'        strMsg = ResolveResString(resCANTOPEN) & strMsg
    Case intGET
'        strMsg = ResolveResString(resCANTREAD) & strMsg
    Case intPUT
'        strMsg = ResolveResString(resCANTWRITE) & strMsg & LS$ & ResolveResString(resCHKSPACE)
    Case intMEMFAIL
'        strMsg = ResolveResString(resOUTOFMEMORY) & strMsg
    End Select

'    Select Case MsgError(Error$ & NL & strMsg, vbAbortRetryIgnore Or vbExclamation Or vbDefaultButton2, "Setup Message")
'    Case vbAbort
'        ExitSetup frmCopy, gintRET_ABORT
'    Case vbIGNORE
'        ConcatSplitFile = vbIGNORE
'    Case vbRetry
'        Resume
'    End Select

CSFCleanup:
    Close mIntConcatFile
    Close intSrcFile
    Err = 0
    Exit Function
End Function
'-----------------------------------------------------------
' FUNCTION: OpenConcatFile
'
' Opens a file to be the destination for concatenation of
' two or more source files that (typically) have been
' split across disks.
'
' Returns: The handle of the file to use for concatentation
'          if the open was successful, or -1 if the open
'          failed and the user chose to ignore the error.
'-----------------------------------------------------------
'
Function OpenConcatFile() As Integer
    Dim intFileNum As Integer
    Dim strMsg As String

    On Error Resume Next

    Do
        Del pDrvTmp & pConcatFile
'        Kill mstrConcatDrive & mstrCONCATFILE
        Err = 0

        intFileNum = FreeFile
        Open pDrvTmp & pConcatFile For Binary Access Write As intFileNum

        If Err > 0 Then
'            strMsg = ResolveResString(421) & NL$ & DRVTMP & CONCATFILE
'            strMsg = strMsg & NL$ & ResolveResString(422)
'            Select Case MsgError(strMsg, vbAbortRetryIgnore Or vbExclamation Or vbDefaultButton2, "Setup Message")
'            Case vbAbort
'                ExitSetup frmCopy, gintRET_ABORT
'            Case vbIGNORE
'                OpenConcatFile = -1
'                Exit Function
'            End Select
        End If
    Loop While Err > 0

    OpenConcatFile = intFileNum
End Function
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
Public Function RunFile(sCommand As String, Optional ExecuteWith As String) As Long
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   
   If Trim(ExecuteWith) = "" Then
      RunFile = ShellExecute(0&, vbNullString, sCommand, vbNullString, vbNullString, vbNormalFocus)
   Else
      RunFile = ShellExecute(0&, vbNullString, ExecuteWith, sCommand, vbNullString, vbNormalFocus)
   End If
   
   Screen.MousePointer = vbDefault
   Exit Function
TrataErro:
   Screen.MousePointer = vbDefault
   Call ClsMsg.ShowError
End Function
Public Sub OpenFile(sCommand As String)
   On Error GoTo TrataErro
   ShellExecute 0&, "open", sCommand, vbNullString, vbNullString, 5 'SW_SHOW=5
   Exit Sub
TrataErro:
   Call ClsMsg.ShowError
End Sub
'---------------------------------------------------------------
' TruncatePath()
' Trunca o caminho de um arquivo para "C:\...\PATH\ARQUIVO.EXT"
'---------------------------------------------------------------
Public Function TruncatePath(ByVal sFileName _
  As String, iMaxLen As Integer) As String
  
  If Len(sFileName) Then
    Dim iPos As Integer, iPos0 As Integer, _
    iPos1 As Integer, iPos2 As Integer, _
    iPos3 As Integer, iPos4 As Integer
    'iPos = SGBkwdInstrS(0, _
    Left$(sFileName, Len(sFileName) - 1), "\")
    iPos = InStr(0, Left$(sFileName, Len(sFileName) - 1), "\")
    iPos0 = InStr(sFileName, ":")
    iPos1 = InStr(sFileName, "\")
    iPos2 = InStr(iPos1, _
    sFileName, "\"): iPos2 = iPos1 + iPos2
    iPos3 = InStr(iPos2, sFileName, "\"): iPos3 = iPos2 + iPos3
    iPos4 = InStr(iPos3, sFileName, "\"): iPos4 = iPos3 + iPos4
    If Len(sFileName) > iMaxLen Then
      If (iPos4 <> 0) And iPos4 + Len(Right$(sFileName, iPos)) _
        <= iMaxLen - 2 Then
        sFileName = Left$(sFileName, iPos4) & "..." & Right$(sFileName, _
          Len(sFileName) - iPos)
      ElseIf (iPos3 > 0) And iPos3 + Len(Mid$(sFileName, _
        iPos)) <= iMaxLen - 2 Then
        sFileName = Left$(sFileName, iPos3) & "..." & _
          Right$(sFileName, Len(sFileName) - iPos)
      ElseIf (iPos3 > 0) And iPos3 + Len(Mid$(sFileName, iPos)) _
        <= iMaxLen - 2 Then
        sFileName = Left$(sFileName, iPos2) & "..." & Right$(sFileName, _
          Len(sFileName) - iPos)
      Else
        sFileName = Left$(sFileName, iPos0 + 1) & "..." & Right$(sFileName, _
          Len(sFileName) - iPos)
      End If
    End If
  End If
  
  TruncatePath = Left$(sFileName, Len(sFileName) - 1)
  
End Function

Private Sub Class_Initialize()
   pDrvTmp = "C:\TMP\"
   pConcatFile = "CONCAT.DIO"
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
End Sub


