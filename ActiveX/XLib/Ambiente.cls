VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ambiente"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public Const CSIDL_DESKTOP = &H0       '// The Desktop - virtual folder
'Public Const CSIDL_TEMPORARY = 1
'Public Const CSIDL_PROGRAMS = 2        '// Program Files
'Public Const CSIDL_CONTROLS = 3        '// Control Panel - virtual folder
'Public Const CSIDL_PRINTERS = 4        '// Printers - virtual folder
'Public Const CSIDL_DOCUMENTS = 5       '// C:\Documents and Settings\dsr\My Documents
'Public Const CSIDL_FAVORITES = 6       '// C:\Documents and Settings\dsr\Favorites
'Public Const CSIDL_STARTUP = 7         '// C:\Documents and Settings\dsr\Start Menu\Programs\Startup
'Public Const CSIDL_RECENT = 8          '// C:\Documents and Settings\dsr\Recent
'Public Const CSIDL_SENDTO = 9          '// C:\Documents and Settings\dsr\SendTo
'Public Const CSIDL_BITBUCKET = 10      '// Recycle Bin - virtual folder
'Public Const CSIDL_STARTMENU = 11      '// C:\Documents and Settings\dsr\Start Menu
'Public Const CSIDL_DESKTOPFOLDER = 16  '// C:\Documents and Settings\dsr\Desktop
'Public Const CSIDL_DRIVES = 17         '// My Computer - virtual folder
'Public Const CSIDL_NETWORK = 18        '// Network Neighbourhood - virtual folder
'Public Const CSIDL_NETHOOD = 19        '// C:\Documents and Settings\dsr\NetHood
'Public Const CSIDL_FONTS = 20          '// C:\WINNT\Fonts
'Public Const CSIDL_SHELLNEW = 21       '// C:\Documents and Settings\dsr\Templates
''22 - C:\Documents and Settings\All Users\Start Menu
''23 - C:\Documents and Settings\All Users\Start Menu\Programs
''24 - C:\Documents and Settings\All Users\Start Menu\Programs\Startup
''25 - C:\Documents and Settings\All Users\Desktop
''26 - C:\Documents and Settings\dsr\Application Data
''27 - C:\Documents and Settings\dsr\PrintHood
''28 - C:\Documents and Settings\dsr\Local Settings\Application Data
''31 - C:\Documents and Settings\All Users\Favorites
''32 - C:\Documents and Settings\dsr\Local Settings\Temporary Internet Files
''33 - C:\Documents and Settings\dsr\Cookies
''34 - C:\Documents and Settings\dsr\Local Settings\History
''35 - C:\Documents and Settings\All Users\Application Data
'Public Const CSIDL_WINDOWS = 36        '// C:\WINNT
'Public Const CSIDL_SYSTEM32 = 37       '// C:\WINNT\system32
'Public Const CSIDL_PROGRAM_FILES = 38  '// C:\Program Files
''39 - C:\Documents and Settings\dsr\My Documents\My Pictures
''40 - C:\Documents and Settings\dsr
''Public Const CSIDL_SYSTEM64 = 41       '// C:\WINNT\SysWOW64
'Public Const CSIDL_COMMON = 43         '// C:\Program Files\Common Files

Public Function AbrirTxt(Arq As String) As Integer
   Dim Hnd As Integer
  
   On Error GoTo CopyErr
   Call ExcluirArquivo(Arq)
   AbrirTxt = FreeFile()
   Open Arq For Output As #AbrirTxt
Exit Function
CopyErr:
  Select Case Err
     Case 55: Err = 0
     Case Else: MsgBox Err.Number & " - " & Err.Description
  End Select
End Function
Public Function AppAtiva(Aplic As Object) As Boolean
   Dim nERRO As Integer, cTxt As String
   
   If Aplic.PrevInstance Then
      On Error Resume Next
      '* Tenta ativar aplicação que já estava rodando (caption de MDIForm)
      AppActivate Aplic.EXEName
      nERRO = Err
      On Error GoTo 0
      '* Testa erro
      '* Aplicação pode estar com o foco em um form não MDIChild
      If nERRO <> 0 Then
         cTxt = "Já existe uma cópia da aplicação rodando."
         cTxt = cTxt & Chr$(10) & "Pressione Ok para fechar esta mensagem"
         cTxt = cTxt & Chr$(10) & "e Alt+Tab para localizar a aplicação."
         Screen.MousePointer = vbDefault
         
         MsgBox cTxt, vbSystemModal + vbExclamation, Aplic.Title
         DoEvents
      Else
         '* Maximiza aplicação que já estava rodando
         SendKeys "% X"
      End If
      '* Encerra aplicação cópia
      AppAtiva = True
   End If
End Function
Public Sub CapturarTela(ByRef PicBox As Object)
   PicBox.Cls
   BitBlt PicBox.hDC, 0, 0, (Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY), GetDC(GetDesktopWindow), 0, 0, &HCC0020
End Sub
Public Function CapturarTelaSis(ByVal pDestinoJPG As String, Optional bMsg As Boolean) As Boolean
   Dim sFileBmp As String
   On Error Resume Next
   
   sFileBmp = Mid(pDestinoJPG, 1, Len(pDestinoJPG) - 3) & "bmp"
   
   If ExisteArquivo(pDestinoJPG) Then Call ExcluirArquivo(pDestinoJPG)
   If ExisteArquivo(sFileBmp) Then Call ExcluirArquivo(sFileBmp)
   
   
   DoEvents
   Call keybd_event(vbKeySnapshot, 1, 0, 0)
   DoEvents
   SavePicture Clipboard.GetData(vbCFBitmap), sFileBmp
   CapturarTelaSis = True
   
   Call ConvertBMPtoJPG(sFileBmp, pDestinoJPG, bMsg)
   If ExisteArquivo(pDestinoJPG) Then
      ExcluirArquivo (sFileBmp)
   End If
End Function
Public Function ComputerName() As String
   Dim lpBuffer As String * 255
   Dim lRet As Long
   Dim i As Integer
   
   lRet = GetComputerName(lpBuffer, 255)
   If lRet <> 0 Then
     ComputerName = UCase$(Left$(Trim$(lpBuffer), Len(Trim$(lpBuffer)) - 1))
     ComputerName = Replace(ComputerName, Chr(0), "")
   Else
     ComputerName = "Não Fornecido."
   End If
End Function
Public Sub ConvertBMPtoJPG(pOriBMP As String, pDestJPG As String, Optional bMsg As Boolean)
  Dim TmpImage    As imgdes ' Image descriptors
  Dim Tmp2Image   As imgdes
  Dim rCode       As Long
  Dim quality     As Long
  Dim VbItCount   As Long
  Dim bdat        As BITMAPINFOHEADER ' Reserve space for BMP struct
  
  Const NO_ERROR = 0
  
  quality = 75
  ' Get info on the file we're to load
  rCode = bmpinfo(pOriBMP, bdat)
  If (rCode <> NO_ERROR) And bMsg Then
      MsgBox "Cannot find file", 0, "Error encountered!"
      Exit Sub
  End If

  VbItCount = bdat.biBitCount
  If (VbItCount >= 16) Then ' 16-, 24-, or 32-bit image is loaded into 24-bit buffer
      VbItCount = 24
  End If

  ' Allocate space for an image
  rCode = allocimage(TmpImage, bdat.biWidth, bdat.biHeight, VbItCount)
  If (rCode <> NO_ERROR) And bMsg Then
    MsgBox "Not enough memory", 0, "Error encountered!"
    Exit Sub
  End If

  ' Load image
  rCode = loadbmp(pOriBMP, TmpImage)
  If (rCode <> NO_ERROR) Then
      freeimage TmpImage ' Free image on error
      MsgBox "Cannot load file", 0, "Error encountered!"
      Exit Sub
  End If

  If (VbItCount = 1) Then ' If we loaded a 1-bit image, convert to 8-bit grayscale
      ' because jpeg only supports 8-bit grayscale or 24-bit color images
    rCode = allocimage(Tmp2Image, bdat.biWidth, bdat.biHeight, 8)
    If (rCode = NO_ERROR) Then
        rCode = convert1bitto8bit(TmpImage, Tmp2Image)
        freeimage TmpImage ' Replace 1-bit image with grayscale image copyimgdes tmp2image, tmpimage
    End If
  End If

  ' Save image
  rCode = savejpg(pDestJPG, TmpImage, quality)
  freeimage TmpImage

End Sub
Public Function CopiarArquivo(Orig As String, Dest As String) As Boolean
   Dim nMsg       As String
   Dim nTipoBox   As Long
   Dim Resp       As Integer

   On Error Resume Next
   
   If ExisteArquivo(Orig) Then
      If ExisteArquivo(Dest) Then
         Call Kill(Dest)
      Else
         Call CriarDiretorio(GetNameFromPath(Dest, 1))
      End If
      FileCopy Orig, Dest

   Else
      Call MsgBox("Arquivo não encontrado: " + UCase(Orig), "Importação")
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
   CopiarArquivo = (Resp = vbYes)
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
Public Function CriarObjeto(pObjeto As String, Optional bMsg As Boolean = True) As Object
   Const SubFolder = "ClasseA"
   
   Dim DsAuto     As Object
   
   Dim MyObject   As Object
   Dim sProjeto   As String
   Dim sArq       As String
   Dim sReg       As String
   Dim sPath      As String
   Dim sPathSis   As String
   Dim sPathNew   As String
   Dim bAchou     As Boolean
         
   On Error GoTo TrataErro
   
   sProjeto = Split(pObjeto, ".")(0)
   If sProjeto = "" Then
      sArq = ""
   ElseIf sProjeto = "aCadastro" Then
      sArq = "CADASTRO.DLL"
   ElseIf sProjeto = "pAutoInstall" Then
      sArq = "AutoInstall.DLL"
   Else
      sArq = sProjeto & ".dll"
   End If
   
   If sArq = "AutoInstall.DLL" Then
      Set MyObject = CreateObject(pObjeto)
      
   ElseIf Len(sArq) > 4 Then
      On Error Resume Next
      Set MyObject = CreateObject(pObjeto)
      If Err <> 0 Then
         sPathSis = GetRegisterDir(sArq)
         If sPathSis <> "" Then
            sPath = ResolvePathName(GetSpecialFolder(38) & SubFolder) '* 38=CSIDL_PROGRAM_FILES
            sReg = "*.reg"
            
            bAchou = (Len(Dir(sPath & sReg)) > 4)
            If bAchou Then
               sPath = ProcuraArquivo(sPath, sReg)
               sReg = Dir(sPath & sReg)
            Else
               sPath = ProcuraArquivo(sPath, sReg)
               sReg = Dir(sPath & sReg)
               bAchou = ExisteArquivo(sPath & sReg)
            End If
                        
            sPathNew = ""
            If bAchou Then
               sPathNew = ResolvePathName(ReadIniFile(sPath & sReg, "Setup", "PATHSETUP"))
            End If
                        
            If sPathNew = "" Then
               On Error Resume Next
               Set MyObject = CreateObject(pObjeto)
               If Err <> 0 Then
                  On Error GoTo TrataErro
                  Call RegServer(sPathSis & sArq, False, False)
                  Call RegServer(sPathSis & sArq, True, False)
               End If
               Err = 0
               On Error GoTo TrataErro
            Else
               'If DsAuto Is Nothing Then
               '   Set DsAuto = CreateObject("pAutoInstall.AutoInstall")
               'End If
               'If Not DsAuto.AtualizaDLL(sPathSis & sArq, sPathNew & sArq, False) Then
               '   Call DsAuto.AtualizaDLL(sPathSis & sArq, sPathNew & sArq, False)
               'End If
               If Not AtualizaDLL(sPathSis & sArq, sPathNew & sArq, False) Then
                  Call AtualizaDLL(sPathSis & sArq, sPathNew & sArq, False)
               End If
               
            End If
         End If
               
         If MyObject Is Nothing Then
            Set MyObject = CreateObject(pObjeto)
         End If
      End If
   End If
   
   Set CriarObjeto = MyObject
GoTo Saida
TrataErro:
   If bMsg Then
      MsgBox "Erro em 'CriarObjeto( " & pObjeto & " )' : " & Err & " - " & Error
   End If
Saida:
   Set DsAuto = Nothing
End Function
Private Function AtualizaDLL(pDLLMaq As String, pDLLNova As String, Optional pObrigatorio As Boolean = False, Optional ArquivoLog As String) As Boolean
   Dim nVersaoMaq    As Double
   Dim nVersaoNova   As Double
   Dim sVersaoMaq    As String
   Dim sVersaoNova   As String
   Dim sPathMaq      As String
   
   On Error Resume Next
   
   nVersaoNova = 0
   nVersaoMaq = 0
   sVersaoNova = "0"
   sVersaoMaq = "0"
   
   sPathMaq = StrReverse(Mid(StrReverse(pDLLMaq), InStr(StrReverse(pDLLMaq), "\") + 1))
   If ExisteArquivo(pDLLNova) Then
      nVersaoNova = GetFileVersionNumber(pDLLNova)
      sVersaoNova = GetFileVersion(pDLLNova)
      
      If ExisteArquivo(pDLLMaq) Then
         nVersaoMaq = GetFileVersionNumber(pDLLMaq)
         sVersaoMaq = GetFileVersion(pDLLMaq)
      Else
         nVersaoMaq = -1
         sVersaoMaq = "-1"
      End If
   End If
   
   If nVersaoNova > nVersaoMaq Or pObrigatorio Then
      If ExisteArquivo(pDLLMaq) Then
         'On Error Resume Next
         'Call RegServer(pDLLMaq, False, False)
         Call ExcluirArquivo(pDLLMaq)
      Else
         Call CriarDiretorio(sPathMaq)
      End If
      
      Call FileCopy(pDLLNova, pDLLMaq)
      AtualizaDLL = RegServer(pDLLMaq, True, False)
      'Call GravaAtualizaDll_Log(ArquivoLog)
      
   ElseIf nVersaoNova <= nVersaoMaq Then
      AtualizaDLL = True
   End If
End Function
Public Function Decrypt2(ByVal Password As String, Optional Key As String) As String
   Dim p    As String
   Dim b    As String
   Dim s    As String
   Dim i    As Integer
   Dim j    As Integer
   Dim A1   As Integer
   Dim A2   As Integer
   Dim A3   As Integer
   
   If Trim(Key) = "" Then Key = Encrypt2("231072150500", "DIO")
   j = 1
   For i = 1 To Len(Key)
     p = p & Asc(Mid$(Key, i, 1))
   Next
   For i = 1 To Len(Password) Step 2
     A1 = Asc(Mid$(p, j, 1))
     j = j + 1
     If j > Len(p) Then j = 1
     b = Mid$(Password, i, 2)
     A3 = Val("&H" + b)
     A2 = A1 Xor A3
     s = s + Chr$(A2)
   Next
   Decrypt2 = Mid(s, 3)
End Function
Public Function Encrypt2(ByVal Password As String, Optional Key As String) As String
   Dim p    As String
   Dim b    As String
   Dim s    As String
   Dim i    As Integer
   Dim j    As Integer
   Dim A1   As Integer
   Dim A2   As Integer
   Dim A3   As Integer
   
   Dim sAux As String
   
   If Len(Trim(Password)) < 2 Then
      sAux = "PI"
   Else
      sAux = Right$(Password, 1) & Left$(Password, 1)
   End If
   Password = sAux & Password
   If Trim(Key) = "" Then Key = Encrypt2("231072150500", "DIO")
   j = 1
   For i = 1 To Len(Key)
     p = p & Asc(Mid$(Key, i, 1))
   Next
    
   For i = 1 To Len(Password)
     A1 = Asc(Mid$(p, j, 1))
     j = j + 1
     If j > Len(p) Then j = 1
     A2 = Asc(Mid$(Password, i, 1))
     A3 = A1 Xor A2
     b = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b
     s = s + b
   Next
   Encrypt2 = s
End Function
Public Function ExcluirDiretorio(Diretorio As String) As Boolean
   Dim fso As Object
   Set fso = CriarObjeto("scripting.FileSystemObject")
   
   On Error GoTo Trata_Erro
   
   ExcluirDiretorio = True
   fso.DeleteFolder Diretorio, True
   Exit Function
   
Trata_Erro:
   ExcluirDiretorio = False
   
End Function
Public Sub EscreverTxt(nArq As Integer, sTexto As String)
   Print #nArq, sTexto
End Sub
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
Public Function ExisteDiretorio(ByVal strPath As String) As Boolean
   strPath = ResolvePathName(strPath)
   ExisteDiretorio = (Trim(Dir(strPath, vbDirectory)) <> "")
End Function
Public Function FecharPrograma(pPrograma As String) As Boolean
'Closes Windows Calculator
   Dim hWindow As Long
   Dim hThread As Long
   Dim hProcess As Long
   Dim lProcessId As Long
   Dim lngResult As Long
   Dim lngReturnValue As Long

   hWindow = FindWindow(vbNullString, pPrograma)
   hThread = GetWindowThreadProcessId(hWindow, lProcessId)
   hProcess = OpenProcess(SYNCHRONIZE, 0&, lProcessId)
   lngReturnValue = PostMessage(hWindow, WM_CLOSE, 0&, 0&)
   lngResult = WaitForSingleObject(hProcess, INFINITE)

   'Does the handle still exist?
   DoEvents
   hWindow = FindWindow(vbNullString, pPrograma)
   If IsWindow(hWindow) = 1 Then
      'The handle still exists. Use the TerminateProcess function
      'to close all related processes to this handle. See the
      'article for more information.
      'MsgBox "Handle still exists."
      FecharPrograma = False
   Else
      'Handle does not exist.
      'MsgBox "All Program Instances Closed."
      FecharPrograma = True
   End If
End Function
Public Sub FecharTxt(Arq As Integer)
      Close #Arq
End Sub
Public Function GetFileExtension(ByVal pFilename As String) As String
    Dim nPos As Integer

    nPos = InStrRev(pFilename, ".")
    If nPos > 0 Then
      GetFileExtension = Mid(pFilename, nPos + 1)
    End If
End Function
Public Function GetFileVersion(ByVal pFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
   Dim sVerInfo As VERINFO
   Dim strVer As String
   
   On Error GoTo GFVError
   
   If IsMissing(fIsRemoteServerSupportFile) Then
      fIsRemoteServerSupportFile = False
   End If
   
   '
   'Get the file version into a VERINFO struct, and then assemble a version string
   'from the appropriate elemen
   '
   If GetFileVerStruct(pFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
      strVer = ""
      strVer = strVer & Format$(sVerInfo.nMSHi, "000") & "."
      strVer = strVer & Format$(sVerInfo.nMSLo, "000") & "."
      strVer = strVer & Format$(sVerInfo.nLSHi, "000") & "."
      strVer = strVer & Format$(sVerInfo.nLSLo, "000")
      GetFileVersion = strVer
   Else
      GetFileVersion = ""
   End If
   
   Exit Function
    
GFVError:
   GetFileVersion = ""
   If Err = 48 Then
      MsgBox "ERRO : " & Err & " - " & Error
   End If
   Err = 0
End Function
Public Function GetFileVersionNumber(pFilename As String) As Double
   Dim Pos  As Integer
   Dim nVer As Double
   Dim sVer As String
   Dim sAux As String
   Dim PosA As Integer
   
   On Error Resume Next
   
   sAux = ""
   PosA = 0
   
   sVer = GetFileVersion(pFilename)
   Pos = InStr(sVer, ".")
   If Pos <> 0 Then
      While Pos <> 0
         Pos = InStr(PosA + 1, sVer, ".")
         sAux = sAux & Right("000" + Mid(sVer, PosA + 1, IIf(Pos = 0, Len(sVer) + 1, Pos) - PosA - 1), 3)
         PosA = Pos
      Wend
   End If
   sAux = IIf(Trim(sAux) = "", "0", Trim(sAux))
   GetFileVersionNumber = Val(sAux)
   
'   Pos = InStr(sVer, ".")
'   While Pos <> 0
'      sVer = Mid(sVer, 1, Pos - 1) & Mid(sVer, Pos + 1)
'      Pos = InStr(sVer, ".")
'   Wend
'   sVer = IIf(sVer = "", sVer = "0", sVer)
'   GetFileVersionNumber = Val(sVer)
End Function
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
Public Function GetRegisterDir(ByVal sArq As String, Optional AppPath As String) As String
   Dim sRoot      As String
   Dim sPath      As String
   Dim bAchou     As Boolean
   Dim sAux       As String
   
   
   sPath = ResolvePathName(AppPath)
   bAchou = ExisteArquivo(AppPath & sArq)
         
   '*****************
   '* \Windows\SysWOW64
   If Not bAchou Then
      sRoot = ResolvePathName(GetSpecialFolder(CSIDL_SYSTEM64))
      
      sPath = ResolvePathName(sRoot & gSubFolder)
      bAchou = ExisteArquivo(sPath & sArq)
      If Not bAchou Then
         sPath = ProcuraArquivo(sPath, sArq, False)
         If sPath <> "" Then bAchou = ExisteArquivo(sPath & sArq)
         If Not bAchou Then
            sPath = sRoot
            bAchou = ExisteArquivo(sPath & sArq)
         End If
      End If
      '*****************
      '* \Windows\System32
      If Not bAchou Then
         sRoot = ResolvePathName(GetSpecialFolder(CSIDL_SYSTEM32))
         
         sPath = ResolvePathName(sRoot & gSubFolder)
         bAchou = ExisteArquivo(sPath & sArq)
         If Not bAchou Then
            sPath = ProcuraArquivo(sPath, sArq, False)
            If sPath <> "" Then bAchou = ExisteArquivo(sPath & sArq)
            If Not bAchou Then
               sPath = sRoot
               bAchou = ExisteArquivo(sPath & sArq)
            End If
         End If
      
         '*****************
         '* \Program Files
         If Not bAchou Then
            sRoot = ResolvePathName(GetSpecialFolder(CSIDL_PROGRAM_FILES))
            
            sPath = ResolvePathName(sRoot & SubFolder)
            bAchou = ExisteArquivo(sPath & sArq)
            If Not bAchou Then
               sPath = ProcuraArquivo(sPath, sArq, False)
               If sPath <> "" Then bAchou = ExisteArquivo(sPath & sArq)
               If Not bAchou Then
                  sPath = sRoot
                  bAchou = ExisteArquivo(sPath & sArq)
               End If
            End If
   
      
            '*****************
            '* \Windows\Common
            If Not bAchou Then
               sRoot = ResolvePathName(GetSpecialFolder(CSIDL_COMMON))
               
               sPath = ResolvePathName(sRoot & SubFolder)
               bAchou = ExisteArquivo(sPath & sArq)
               If Not bAchou Then
                  sPath = ProcuraArquivo(sPath, sArq, False)
                  If sPath <> "" Then bAchou = ExisteArquivo(sPath & sArq)
                  If Not bAchou Then
                     sPath = sRoot
                     bAchou = ExisteArquivo(sPath & sArq)
                  End If
               End If
      
               '*****************
               '* \Windows
               If Not bAchou Then
                  sRoot = GetSpecialFolder(CSIDL_WINDOWS)
                  
                  sPath = ResolvePathName(sRoot & SubFolder)
                  bAchou = ExisteArquivo(sPath & sArq)
                  If Not bAchou Then
                     sPath = ProcuraArquivo(sPath, sArq, False)
                     If sPath <> "" Then bAchou = ExisteArquivo(sPath & sArq)
                     If Not bAchou Then
                        sPath = sRoot
                        bAchou = ExisteArquivo(sPath & sArq)
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   sPath = IIf(bAchou, sPath, "")
   
   GetRegisterDir = ResolvePathName(sPath)
End Function
Public Function GetSerialNumber(Optional sDrive As String = "C:\") As Long
   GetVolumeInformation sDrive, vbNullString, 0, GetSerialNumber, 0, 0, vbNullString, 0
End Function

Public Function GetShortName(sFile As String) As String
   Dim sShortFile As String * 67
   Dim lResult As Long

   'Make a call to the GetShortPathName API
   lResult = GetShortPathName(sFile, sShortFile, _
   Len(sShortFile))

   'Trim out unused characters from the string.
   GetShortName = Left$(sShortFile, lResult)

End Function

Public Function GetSpecialFolder(CSIDL As Long) As String
    Dim sPath  As String
    Dim IDL    As ITEMIDLIST
    Dim nhWnd  As Long
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
   If CSIDL = CSIDL_TEMPORARY Then
      GetSpecialFolder = ResolvePathName(GetTempFolder)
   Else
      GetSpecialFolder = ""
      If SHGetSpecialFolderLocation(nhWnd, CSIDL, IDL) = 0 Then
         '
         ' Get the path from the ID list, and return the folder.
         '
         sPath = Space$(260)
         If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            GetSpecialFolder = ResolvePathName(Left$(sPath, InStr(sPath, vbNullChar) - 1) & "")
         End If
     End If
   End If
End Function

Public Function GetTempFolder() As String
   ' Returns the path to the user's Temp folder. To boot, Windows
   ' requires that a temporary folder exist, so this should always
   ' safely return a path to one. Just in case, though, check the
   ' return value of GetTempPath.
   
   Dim strTempPath As String
   Dim lngTempPath As Long
   
   ' Fill string with null characters.
   strTempPath = String(144, vbNullChar)
   ' Get length of string.
   lngTempPath = Len(strTempPath)
   ' Call GetTempPath, passing in string length and string.
   If (GetTempPath(lngTempPath, strTempPath) > 0) Then
      ' GetTempPath returns path into string.
      ' Truncate string at first null character.
      GetTempFolder = ResolvePathName(VBA.Left(strTempPath, InStr(1, strTempPath, vbNullChar) - 1))
   Else
      GetTempFolder = ""
   End If
End Function
Public Function GetUserName() As String
  Dim nSize As Long
  Dim sUser As String
  
  Dim nStatus As Long
    
  On Error GoTo ErroID

  GetUserName = "Unknow"
  nSize = 255
  
  sUser = Space$(nSize)

  nStatus = GetUserNameA(sUser, nSize)
  If nStatus = 1& Then
    sUser = Left$(sUser, nSize - 1&)
  Else
    Err = nStatus
  End If
  
  GetUserName = sUser
  
  Exit Function
ErroID:
  GetUserName = "Error"
End Function
Public Sub GravarArquivoLog(pPath As String, pNomeArq As String, pTITULO As String, pConteudo As String, Optional bHora As Boolean = True)

   Dim cArq As String
   Dim cPath As String
   Dim cNmArquivo As String

   cPath = pPath & IIf(VBA.Right$(pPath, 1) <> "\", "\", "")
   cNmArquivo = pNomeArq
   
   If bHora = True Then
      cNmArquivo = cNmArquivo & Format(Now, "yyyymmddhhMM")
   Else
      cNmArquivo = cNmArquivo & Format(Now, "yyyymmdd")
   End If
   
   cNmArquivo = cNmArquivo & ".txt"
   
   Open cPath & cNmArquivo For Append As #1
         
   Print #1, pTITULO & " - " & Format(Now, "dd/mm/yyyy hh:MM:ss") & vbNewLine
   Print #1, "_____________________________________________________________" & vbNewLine
   Print #1, pConteudo
   
   Close #1
End Sub
Public Function RegServer(sServerPath As String, Optional fRegister = True, Optional fMsg As Boolean = True, Optional isActivexExe As Boolean = False) As Boolean
   Dim hMod       As Long    ' module handle
   Dim lpfn       As Long    ' reg/unreg function address
   Dim sCmd       As String  ' msgbox string
   Dim lpThreadID As Long    ' unused, receives the thread ID
   Dim hThread    As Long    ' thread handle
   Dim fSuccess   As Boolean ' if things worked
   Dim dwExitCode As Long    ' thread's exit code if it doesn't finish
   
   ' Load the server into memory
   hMod = LoadLibrary(sServerPath)
   
   ' Get the specified function's address and our msgbox string.
   If fRegister Then
      If isActivexExe Then
         lpfn = GetProcAddress(hMod, "ExeRegisterServer")
      Else
         lpfn = GetProcAddress(hMod, "DllRegisterServer")
      End If
     sCmd = "register"
   Else
      If isActivexExe Then
         lpfn = GetProcAddress(hMod, "ExeUnregisterServer")
      Else
         lpfn = GetProcAddress(hMod, "DllUnregisterServer")
      End If
      sCmd = "unregister"
   End If
   
   ' If we got a function address...
   If lpfn Then
     
     ' Create an alive thread and execute the function.
     hThread = CreateThread(ByVal 0, 0, ByVal lpfn, ByVal 0, 0, lpThreadID)
     
     ' If we got the thread handle...
     If hThread Then
       
       ' L_Wait 10 secs for the thread to finish (the function may take a while...)
       fSuccess = (WaitForSingleObject(hThread, 10000) = WAIT_OBJECT_0)
       
       ' If it didn't finish in 10 seconds...
       If Not fSuccess Then
         ' Something unlikely happened, lose the thread.
         Call GetExitCodeThread(hThread, dwExitCode)
         Call ExitThread(dwExitCode)
       End If
       
       ' Lose the thread handle
       Call CloseHandle(hThread)
     
     End If   ' hThread
   End If   ' lpfn
   
   ' Free the server if we loaded it.
   If hMod Then Call FreeLibrary(hMod)
   
   RegServer = fSuccess
   
   If fMsg Then
      If fSuccess Then
         MsgBox "Successfully " & sCmd & "ed " & sServerPath   ' past tense
      Else
        MsgBox "Failed To " & sCmd & " " & sServerPath, vbExclamation
      End If
   End If
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
Public Sub IniCommonControls()
   InitCommonControls
End Sub
Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
    Dim dwFlags As Long
    Dim WebTest As Boolean
    
    On Error Resume Next
    
    ConnType = ""
    WebTest = InternetGetConnectedState(dwFlags, 0&)
    Select Case WebTest
        Case dwFlags And CONNECT_LAN: ConnType = "LAN"
        Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
        Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
        Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
        Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configurada"
        'Case dwflags And CONNECT_RAS: ConnType = "Remota"
    End Select
   IsWebConnected = WebTest
End Function
Private Function FindFile(RootPath As String, FileName As String) As String
   Dim lNullPos As Long
   Dim lResult As Long
   Dim sBuffer As String

   On Error GoTo FileFind_Error 'Allocate buffer

   ' Change mouse cursor
   Screen.MousePointer = vbHourglass
   sBuffer = Space(260 * 2) 'Find the file
   lResult = SearchTreeForFile(RootPath, FileName, sBuffer)
   'Trim null, if exists
   If lResult Then
      lNullPos = InStr(sBuffer, vbNullChar)
      If Not lNullPos Then
         sBuffer = Left(sBuffer, lNullPos - 1)
      End If
      'Return filename
      FindFile = GetNameFromPath(sBuffer, 1)
   Else 'Nothing found
      FindFile = vbNullString
   End If
   ' Change mouse cursor
   Screen.MousePointer = vbDefault

   Exit Function
FileFind_Error:
   FindFile = vbNullString
End Function
Public Function ProcuraArquivo(ByVal pPath As String, ByVal pArq As String, Optional pAdmin As Boolean = True) As String
   Dim sAux    As String
   Dim sPath   As String
   Dim bAchou  As Boolean
   Dim sPath0  As String
   Dim i       As Integer
   Dim nVezes  As Integer
   Dim bAdmin  As Boolean
   
   On Error GoTo TrataErro
   nVezes = 10000
   
   
   pPath = ResolvePathName(pPath)
   sAux = pPath
   ChDir sAux
   sPath0 = sAux
   
   bAchou = ExisteArquivo(sAux & pArq)
   If bAchou Then
      sPath = pPath
   Else
      sAux = Dir(sAux, vbDirectory)
      While sAux <> ""
         sAux = Dir(Attributes:=VbFileAttribute.vbDirectory)
         bAdmin = IIf(UCase(sAux) = "ADMIN", pAdmin, True)
         
         If Mid(sAux, 1, 1) <> "." And sAux <> "" And bAdmin Then
            If (GetAttr(pPath & sAux) And vbDirectory) = vbDirectory Then
               sAux = ResolvePathName(pPath & sAux)
               bAchou = ExisteArquivo(sAux & pArq)
               
               If bAchou Then
                  sPath = sAux
                  sAux = ""
               Else
                  sPath = ProcuraArquivo(sAux, pArq)
                  If sPath = "" Then
                     '*********************
                     '* Retorna ao diretório anterior
                     ChDir sPath0
                     If sAux <> Dir(pPath, vbDirectory) Then
                        i = 0
                        While i < nVezes
                           i = i + 1
                           If sAux = ResolvePathName(pPath & Dir(Attributes:=vbDirectory)) Then
                              i = nVezes + 1
                           End If
                        Wend
                     End If
                  Else
                    sAux = ""
                  End If
               End If
            End If
         End If
      Wend
   End If
   ProcuraArquivo = sPath
Exit Function
TrataErro:
   If Err = 16 Then
      Resume Next
   ElseIf Err = 76 Then   '*Path not Fouund
      Resume Next
   Else
      MsgBox Err.Number & " - " & Err.Description, vbCritical, "TrataErro AutoInstall"
      Resume Next
   End If
End Function
Public Function ProgramaAtivo(pPrograma As String) As Boolean
   Dim hWindow As Long
   hWindow = FindWindow(vbNullString, pPrograma)
   If IsWindow(hWindow) = 1 Then
      ProgramaAtivo = True
   Else
      ProgramaAtivo = False
   End If
End Function

Public Function GetRegKey(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
   Dim hCurKey As Long
   Dim lValueType As Long
   Dim strBuffer As String
   Dim lDataBufferSize As Long
   Dim intZeroPos As Integer
   Dim lRegResult As Long

   ' Set up default value
   If Not IsEmpty(Default) Then
      GetRegKey = Default
   Else
      GetRegKey = ""
   End If

   ' Open the key and get length of string
   lRegResult = RegOpenKey(hKey, strPath, hCurKey)
   lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

   If lValueType = 1 Then
      'initialise string buffer and retrieve string
      strBuffer = String(lDataBufferSize, " ")
      lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)

      'String de formato ' format string
      intZeroPos = InStr(strBuffer, Chr$(0))
      If intZeroPos > 0 Then
         GetRegKey = Left$(strBuffer, intZeroPos - 1)
      Else
         GetRegKey = strBuffer
      End If
   End If
lRegResult = RegCloseKey(hCurKey)
End Function
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = 1 Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = 3 Then
        

            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function
Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    'Open the key
    RegOpenKey hKey, strPath, Ret
    'Get the key's content
    GetString = RegQueryStringValue(Ret, strValue)
    'Close the key
    RegCloseKey Ret
End Function



Public Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, Optional DefaultValue As String) As String
   Dim strBuffer As String
   
   strBuffer = Space$(255)
   If GetPrivateProfileString(strSection, strKey, "", strBuffer, 255, strIniFile) > 0 Then
      If InStr(strBuffer, Chr$(0)) > 0 Then
        strBuffer = VBA.Left$(strBuffer, InStr(strBuffer, Chr$(0)) - 1)
      End If
      ReadIniFile = RTrim$(strBuffer)
   Else
      ReadIniFile = DefaultValue
   End If
End Function
Public Function ReadTextFile(strPath As String) As String
    On Error GoTo ErrTrap
    Dim intFileNumber As Integer
    
    If Dir(strPath) = "" Then Exit Function
    intFileNumber = FreeFile
    Open strPath For Input As #intFileNumber
    
    ReadTextFile = Input(LOF(intFileNumber), #intFileNumber)
ErrTrap:
    Close #intFileNumber
End Function
Public Function SetMDI(ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
   SetMDI = SetParent(hWndChild, hWndNewParent)
End Function
Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
      SetTopMostWindow = False
   End If
End Function

Public Sub SincShell(Comando As String, Optional Modo As VbAppWinStyle = vbMaximizedFocus, Optional EsperaProcesso = True)
   Dim IDProcess  As Long
   Dim hProcess   As Long
   Dim ExitCode   As Long
   Dim Ret        As Long
   
   On Error GoTo TrataErro

   IDProcess = Shell(Comando, Modo)
   If EsperaProcesso Then
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 1, IDProcess)
      Do
         Ret = GetExitCodeProcess(hProcess, ExitCode)
         DoEvents
      Loop While (ExitCode = STILL_ACTIVE)
      Ret = CloseHandle(hProcess)
   End If
Exit Sub
TrataErro:
   MsgBox CStr(Err) & " - " & CStr(Error)
End Sub
Public Function WriteIniFile(ByVal strIniFile As String, strSection As String, strKey As String, strValue As String) As Boolean
   Dim intLen As Integer
   
   If Not ExisteArquivo(strIniFile) Then
      intLen = AbrirTxt(strIniFile)
      Call FecharTxt(intLen)
   End If
   intLen = 0
   intLen = WritePrivateProfileString(strSection, strKey, strValue, strIniFile)
   WriteIniFile = (intLen > 0)
End Function
Public Function WriteTextFile(ByVal pFile As String, pText As String) As Boolean
   Dim intLen As Integer
   
   On Error GoTo Saida
'   If ExisteArquivo(pFile) Then
      intLen = FreeFile
      Open pFile For Output As #intLen
      Print #intLen, pText
      Close #intLen
'   End If
   WriteTextFile = (intLen > 0)
Saida:
End Function
