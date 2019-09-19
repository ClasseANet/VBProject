Attribute VB_Name = "DSLOAD"
Option Explicit
'============================================================
'============================================================
Type VERINFO                                            'Version FIXEDFILEINFO
    strPad1 As Long                                     'Pad out struct version
    strPad2 As Long                                     'Pad out struct signature
    nMSLo As Integer                                    'Low word of ver # MS DWord
    nMSHi As Integer                                    'High word of ver # MS DWord
    nLSLo As Integer                                    'Low word of ver # LS DWord
    nLSHi As Integer                                    'High word of ver # LS DWord
    strPad3(1 To 16) As Byte                            'Skip some of VERINFO struct (16 bytes)
    FileOS As Long                                      'Information about the OS this file is targeted for.
    strPad4(1 To 16) As Byte                            'Pad out the resto of VERINFO struct (16 bytes)
End Type
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public Type POINTAPI   ' pt
   x As Long
   y As Long
End Type
' WaitForSingleObject rtn vals
Public Const STATUS_WAIT_0 = &H0
Public Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0) ' The State of the specified object is signaled (success)

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long '(lpThreadAttributes As SECURITY_ATTRIBUTES,
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFilename As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
Public Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFilename As String, lVerHandle As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Sub lmemcpy Lib "VB5STKIT.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public m_hMod As Long
'============================================================
'============================================================
Global Const gSubFolder = "ClasseA"

Global gLocalReg     As String
Global gSetupFile    As String
Global gMenuDinamico As Boolean
Global gDebug        As Boolean
Global gDebugGotoShow As Boolean
Global gIDUSU        As String
Global gReload       As Boolean
'============================================================
'============================================================
Public gCaption1  As String
Public gCaption2  As String
Public gCaption3  As String
'============================================================
'============================================================
'#Const ComRef = False
Global MySplash As Object
'#If ComRef Then
'   Global Sys     As SysA.SetA
'   Global Splash  As Conexao.Splash
'#Else
   Global Sys     As Object
   Global Splash  As Object
'#End If
Global ClTelas As Collection
'============================================================
Dim cProgs   As Collection

Public Sub Main()
   'Dim i As Integer: For i = 1 To 100: Debug.Print Environ(i): Next
   'Call ValidarCNPJ("12185606000134")
   'sCommand = "/RUNAS /CODSIS:"
   Dim sCommand As String
   
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   
   sCommand = Trim(UCase(Command$))
   gLocalReg = App.Path & "\" & App.EXEName & ".reg"
   gSetupFile = "SETUP.INI"
   gDebug = (InStr(sCommand, "DEBUG") <> 0)
            
   If gDebug Then
      If vbYes = MsgBox("Goto Show?", vbYesNo, "Debug") Then
         gDebugGotoShow = True
         gDebug = False
      End If
   End If
   If Not gDebug And InStr(UCase(App.Path), "\SISTEMAS\") = 0 Then
      Call MyLoadgCODSIS(True)
      ExibeSplah
      If Dir("C:\Tmp\" & gCODSIS & "\", vbDirectory) = "" Then Call CriarDiretorio("C:\Tmp\" & gCODSIS & "\")
      Call DesativarProgramas
      Call DesativarProgramas
      Set xAmbiente = Nothing
      Call AutoInstall
      Call AtivarProgramas
   End If
   Call MyLoadgCODSIS
   
   Call Proprietario
   Call MyRunAs
   
   '***************
   '* verificando se existe o código do sistema no parâmentro
   '* do executavel para atribuir ao codsis
   If sCommand <> "" Then
      If InStr(1, sCommand, "/CODSIS:") > 0 Then
         '* Existe o  codigo do sistema no parâmetro do executavel
         sCommand = Mid(sCommand, InStr(1, sCommand, "/CODSIS:", vbTextCompare))
         If InStr(1, sCommand, " ", vbTextCompare) > 0 Then
            sCommand = Trim(Mid(sCommand, 1, InStr(1, sCommand, " ", vbTextCompare)))
         End If
         If Mid$(sCommand, 9) <> "" Then
            gCODSIS = Mid$(sCommand, 9)
         End If
      End If
   End If
      
   If gCODSIS = "" Then MyLoadgCODSIS
   Call LimpaInstaciaObj
   Call InstaciaObj
   
   '*********
   '* Testa se já existe uma cópia da aplicação rodando e define formato Data e número.
   If gIDUSU = "" Then
      If AppAtiva(App) Then
         If gDebug Then MsgBox "End"
         End
      End If
   End If
   
   If Not Sys Is Nothing Then
      With Sys
         .ExePath = App.Path
      End With
   End If
   If gDebug Then MsgBox "Entrar ExibeSenha"
   Call ExibeSenha
   Exit Sub
TrataErro:
   Call UnloadIni
   Set Splash = Nothing
   MsgBox Err.Number & " - " & Err.Description, Title:="Projeto3R.DSLOAD.Main"
   Resume Next
End Sub
Private Sub ExibeSplah()
   Dim sArq       As String
   Dim sCODSIS    As String
   Dim sAppPath   As String
   Dim sExeName   As String
   Dim sCommand   As String
   Dim sPathExe   As String
   Dim sPathSetup As String
   Dim SetupIni   As String
   Dim sStatus    As String
   Dim LerIni     As Boolean
   Dim sPathSis   As String
   Dim sPathNewFile As String
   Dim sAux       As String
   
   On Error GoTo TrataErro
   
   '*************
   '* SetupFile = ReadIniFile(gLocalReg, "Setup", "Setup.Ini", App.Path)
   sCommand = Trim(UCase(Command$))
   sCODSIS = IIf(Trim(gCODSIS) = "", App.EXEName, gCODSIS)
   sAppPath = L_ResolvePathName(App.Path)
   sExeName = App.EXEName
   sPathExe = L_ResolvePathName(L_GetTag(sCommand, "SourcePath", ""))
   
   
   '*************
   '* PathSetup = "\\guarani\sistemas\admin\"
   SetupIni = sAppPath & gSetupFile
   If gDebug Then MsgBox "FileExists(" & SetupIni & ")"
   If Not L_ExisteArquivo(SetupIni) Then
      sPathSetup = L_ReadIniFile(gLocalReg, "Setup", "PATHSETUP")
   Else
      sPathSetup = sAppPath
   End If
   sPathSetup = L_ResolvePathName(sPathSetup)
   SetupIni = sPathSetup & gSetupFile
   
   
   If gDebug Then MsgBox "FileExists(" & SetupIni & ")"
   If L_ExisteArquivo(SetupIni) Then
      sStatus = L_ReadIniFile(SetupIni, "AutoInstall", "Status", "0")
      LerIni = (sStatus = "1" Or sStatus = "2")
      If gDebug Then MsgBox "LerIni = " & LerIni
      If LerIni Then
         sPathNewFile = L_ResolvePathName(L_ReadIniFile(SetupIni, "AutoInstall Files", "Path", ""))
         
         If Trim(sPathNewFile) <> "" Then
            
            sArq = "SPL" & sCODSIS & ".dll"
            sPathSis = L_GetRegisterDir(sArq, App.Path)
            If Trim(sPathSis) <> "" Then
               If gDebug Then
                  sAux = "File = " & sArq & vbNewLine
                  sAux = sAux & "sPathSis = " & sPathSis & vbNewLine
                  sAux = sAux & "sPathNewFile = " & sPathNewFile & vbNewLine
                  sAux = sAux & " ResolvePathName(" & sPathNewFile & ") = " & L_ResolvePathName(sPathNewFile) & vbNewLine
                  sAux = sAux & L_GetFileVersionNumber(sPathSis & sArq) & " < " & L_GetFileVersionNumber(sPathNewFile & sArq) & " = " & (L_GetFileVersionNumber(sPathSis & sArq) < L_GetFileVersionNumber(sPathNewFile & sArq)) & vbNewLine
                  MsgBox sAux
               End If
               If L_GetFileVersionNumber(sPathSis & sArq) < L_GetFileVersionNumber(sPathNewFile & sArq) Then
                     
                  Call L_RegServer(sPathSis & sArq, False, False)
                  Call Kill(sPathSis & sArq)
                  Call FileCopy(sPathNewFile & sArq, sPathSis & sArq)
                  Call L_RegServer(sPathSis & sArq, True, False)
               End If
            End If
         End If
      End If
   End If
   
   GoTo Saida
TrataErro:
   If Err <> 0 Then
MsgBox Err.Number & " - " & Err.Description     'ShowError("Sub Main()")
   Resume 0
End If

Saida:
   On Error Resume Next
   If InStr(UCase(App.Path), "\SISTEMAS\") = 0 Then
      Set MySplash = CreateObject("SPL" & gCODSIS & ".Tela")
      If Not MySplash Is Nothing Then
         MySplash.Show
      End If
   End If
End Sub
Private Sub Proprietario()
   Dim MyFunc As Object
   
   Set MyFunc = CreateObject("CLA.PROPRIETARIO")
   With MyFunc
      .AppPath = App.Path
      .AppExe = App.EXEName
      .SetupFile = gSetupFile
      .Executa
   End With
   Set MyFunc = Nothing
End Sub
Private Sub AutoInstall()
   If gDebug Then MsgBox "AutoInstall"
   Dim AutoIn     As Object
   Dim sCODSIS    As String
   Dim sAppPath   As String
   Dim sExeName   As String
   Dim sCommand   As String
   
   On Error GoTo TrataErro
   
   '*************
   '* gSetupFile = ReadIniFile(gLocalReg, "Setup", "SetupFile", App.Path)
   Call MyLoadgCODSIS(True)
   sCODSIS = IIf(Trim(gCODSIS) = "", App.EXEName, gCODSIS)
   sAppPath = App.Path
   sExeName = App.EXEName
   sCommand = Trim(Command$) & IIf(gDebug, " /Debug", "")
   If gDebug Then
      If InStr(sCommand, "DEBUG") = 0 Then
         sCommand = sCommand & " /Debug"
      End If
   End If
   
   If Trim(sCommand) <> "" Then Screen.MousePointer = vbHourglass
   
   If AutoIn Is Nothing Then
      If gDebug Then MsgBox "Antes do VerificaAutoInstall"
      '* Não verifica Autonstall.dll se o executável for temporário, ou seja
      '* se o executável estiver sendo atualizado.
      If gDebug Then MsgBox "InStr(sCommand, |SOURCEPATH) R:" & InStr(sCommand, "|SOURCEPATH")
      
      If InStr(sCommand, "|SOURCEPATH") = 0 Then
         Call VerificaAutoInstall
      End If
      
      If gDebug Then MsgBox "Depois do VerificaAutoInstall"
      If gDebug Then MsgBox "Set AutoIn = CreateObject(pAutoInstall.AutoInstall)"
      
      Set AutoIn = CreateObject("pAutoInstall.AutoInstall")
      
      If gDebug Then MsgBox "AutoInstall Created"
   End If
   
'* Function VerificaSetup(pCODSIS As String, AppPath As String, pExeName As String, pCommand As String, Optional PathSetup As String, Optional gSetupFile As String = "SETUP.INI") As Boolean
   If gDebug Then MsgBox "AutoIn.VerificaSetup(pCODSIS:=" & gCODSIS & ", AppPath:=" & sAppPath & ", pExeName:=" & sExeName & ", pCommand:=" & sCommand & ", PathSetup:='" & "', gSetupFile:=" & gSetupFile & ")"
   
   If AutoIn Is Nothing Then
      If vbNo = MsgBox("O sistema não verificou atualizações no servidor, você poderá estar executando uma versão desatualizada." & vbNewLine & vbNewLine & "Deseja continuar?", vbQuestion + vbYesNo, "AutoInstall") Then
         End
      End If
   Else
      If Not AutoIn.VerificaSetup(pCODSIS:=sCODSIS, AppPath:=sAppPath, pExeName:=sExeName, pCommand:=sCommand, PathSetup:="", SetupFile:=gSetupFile) Then
         Set AutoIn = Nothing
         End
      End If
   End If
   
   Screen.MousePointer = vbDefault
   If gDebug Then MsgBox "Out AutoIn.VerificaSetup"
   GoTo Saida
TrataErro:
   If Err <> 0 Then
      MsgBox Err.Number & " - " & Err.Description, vbCritical, gCODSIS & " TrataErro AutoInstall"
      Resume Next
   End If
Saida:
   gCODSIS = ""
   Set AutoIn = Nothing
End Sub
Public Sub ExibeSenha(Optional pTrocaConexao As Boolean = False, Optional pxDb As Object, Optional pIDUSU As String)
   Dim bAux As Boolean
   Dim sAux As String
   Dim i As Integer
   Dim nAux As Integer
   
   On Error GoTo TrataErro

   If gDebug Then MsgBox "Em ExibeSenha"
   If Splash Is Nothing Then
      Set Splash = CriarObjeto("CONEXAO.Splash")
   End If
   If Splash Is Nothing Then
      MsgBox "Tela de senha não foi criada.", vbInformation
   Else
      If Not Sys Is Nothing Then
         Sys.ExePath = App.Path
         Sys.LocalReg = gLocalReg
         'Sys.IDLOJA = ReadIniFile(gLocalReg, "Config", "LOJA", "1")
      End If
      
      If gDebug Then MsgBox "Splash Not Nothing"
      With Splash
         Set .Sys = Sys
         .DebugSys = False
         .CODSIS = gCODSIS
         .Alias = gALIAS
         .dbTipo = gDBTipo
         .Server = gSERVER
         .dbName = gDBNAME
         .UID = gDBUSER
         .PWD = gDBPWD
         
         .ExeVersao = LoadVersion
         .ExePath = App.Path & "\"
         .ExeFile = App.EXEName
         If Not MyLoadPicture Is Nothing Then
            Set .ImgTitulo = MyLoadPicture
         End If
         
         If Sys.GetParam("Caption1", gCODSIS) <> "" Then gCaption1 = Sys.GetParam("Caption1", gCODSIS)
         If Sys.GetParam("Caption2", gCODSIS) <> "" Then gCaption2 = Sys.GetParam("Caption2", gCODSIS)
         If Sys.GetParam("Caption3", gCODSIS) <> "" Then gCaption3 = Sys.GetParam("Caption3", gCODSIS)
         .Caption1 = gCaption1
         .Caption2 = gCaption2
         .Caption3 = gCaption3
         
         DoEvents
         Screen.MousePointer = vbDefault
            
         If pTrocaConexao Then
            .IDUSU = MDI.StatusBar.Pane(1).Text
         Else
            If Not .VerificaLicenca1 Then
               If UCase(Mid(App.Path, 1, 12)) <> "C:\SISTEMAS\" Then
                  End
               End If
            End If
         End If
         gIDUSU = .IDUSU
         If gDebug Then MsgBox "gIDUSU=" & gIDUSU
         If gIDUSU = "" Then
            If gDebug Then MsgBox "Show"
            Set MySplash = Nothing
            'If pTrocaConexao Then
               .xdb.SrvDesconecta
            'End If
            .Show
         Else
            If gDebug Then MsgBox "Conectar"
            .IDUSU = gIDUSU
            .Conectar
         End If
         
         Dim dHHIni As Date
         Dim nTimeout As Long
         nTimeout = 5 '*Segundos
         dHHIni = CDate(Format(Now, "hh:mm:ss"))

         If gDebug Then MsgBox "Conectado=" & .xdb.Conectado
         Do While Not .xdb.Conectado
            If .Cancelado Then
               If pTrocaConexao Then
                  .Conectar
                  If .xdb.Conectado Then
                     pIDUSU = .IDUSU
                  End If
                  Exit Do
               Else
                  End
               End If
            End If

            DoEvents
            If gIDUSU <> "" Then '*Se não exibir senha, verificar tempo de conexão
               If DateDiff("s", dHHIni, CDate(Format(Now, "hh:mm:ss"))) > nTimeout Then
                  .Alias = "PRODUCAO"
                  If InStr(UCase(.Server), "\SQLEXPRESS") <> 0 Then
                     .Server = Environ("COMPUTERNAME") & "\SQLEXPRESS"
                  Else
                     .Server = Environ("COMPUTERNAME")
                  End If
                  .dbTipo = 1
                  .dbName = gDBNAME
                  .UID = gDBUSER
                  .PWD = gDBPWD
                  If Not ExisteArquivo(.LocalReg) Then
                     Call .GerarLocalReg
                  End If
                  If Dir(Environ("Programfiles") & "\Microsoft SQL Server\", vbDirectory) <> "" Or Dir(Environ("Programfiles") & "(x86)\Microsoft SQL Server\", vbDirectory) <> "" Then
                     .Conectar
                  End If
                  If .Conectado Then
                     bAux = ExisteArquivo(Sys.PathSetup & gSetupFile)
                     If Not bAux Then
                        If ExisteArquivo(Environ("Programfiles") & "(x86)\ClasseA\Admin\Dll\" & gSetupFile) Then
                           Sys.PathSetup = Environ("Programfiles") & "(x86)\ClasseA\Admin\Dll\"
                        End If
                     End If
                     If bAux Then
                        Call WriteIniFile(Sys.PathSetup & gSetupFile, "Database Format", "SERVER", .Server)
                        If InStr(UCase(Sys.PathSetup), "(X86)") <> 0 Then
                           Call WriteIniFile(Sys.PathSetup & gSetupFile, "AutoInstall Files", "Path", Sys.PathSetup)
                        End If
                        If ExisteArquivo(Sys.PathSetup & "Bak\" & gSetupFile & ".zip") Then
                           Call Unzip(Sys.PathSetup & "Bak\", gSetupFile & ".zip", Sys.PathTmp, False)
                           If ExisteArquivo(Sys.PathTmp & gSetupFile) Then
                              sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, "AutoInstall Files", "Path"))
                              If sAux <> "" Then Call WriteIniFile(Sys.PathSetup & gSetupFile, "AutoInstall Files", "Path", sAux)
                              i = 1
                              sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, "AutoInstall Files", "Path" & i))
                              While sAux <> ""
                                 If sAux <> "" Then Call WriteIniFile(Sys.PathSetup & gSetupFile, "AutoInstall Files", "Path" & i, sAux)
                                 i = i + 1
                                 sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, "AutoInstall Files", "Path" & i))
                              Wend
                              
                              sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path"))
                              If sAux <> "" Then Call WriteIniFile(Sys.PathSetup & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path", sAux)
                              i = 1
                              sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path" & i))
                              While sAux <> ""
                                 If sAux <> "" Then Call WriteIniFile(Sys.PathSetup & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path" & i, sAux)
                                 i = i + 1
                                 sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path" & i))
                              Wend
                           End If
                           
                        End If
                     End If
                  Else
                     If vbYes = ExibirPergunta("Tempo de conexão expirado!" & vbNewLine & "Deseja verificar sua conexão?") Then
                        Set MySplash = Nothing
                        .ShowCon vbModal
                     'End With
                     Else
                        'MsgBox "Tempo de conexão expirado!", vbInformation, "Conexão"
                        End
                     End If
                  End If
               End If
            End If
         Loop
                  
         DoEvents
         Screen.MousePointer = vbHourglass

         If gDebug Then MsgBox "pTrocaConexao=" & pTrocaConexao & " Or Splash.xDb.Conectado=" & .xdb.Conectado
         If .xdb.Conectado Or pTrocaConexao Then
         If gDebug Then MsgBox "pTrocaConexao=" & pTrocaConexao & " Or Splash.xDb.Conectado=" & .xdb.Conectado
            If .Cancelado Then
               Set Sys.xdb = IIf(pTrocaConexao, .xdb, pxDb)
               gIDUSU = pIDUSU
            Else
               Set Sys.xdb = .xdb
               gIDUSU = .IDUSU
            End If
            
            If gDebug Then MsgBox "IDUSU : " & gIDUSU & vbNewLine & "CODSIS : " & gCODSIS
            Call AtualizaBD
            
            With Sys
               'Set .XDb = XDb
               .IDUSU = Trim(gIDUSU)
               .CODSIS = gCODSIS
               .LocalReg = gLocalReg
               '********
               '* Se pasta de Administração alterou re-executa Sistema.
               If gDebug Then MsgBox ".PathSetup : " & .PathSetup & vbNewLine & "ReadIniFile(gLocalReg, 'Setup', 'PATHSETUP', '') : " & ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "") & vbNewLine & " Reexecuta : " & (.PathSetup <> ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "") And .PathSetup <> "")
               If ResolvePathName(.PathSetup) <> ResolvePathName(ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "")) And .PathSetup <> "" Then
                  sAux = ResolvePathName(.GetParam(pCODPARAM:="PATHSETUP", pCODSIS:="GLOBAL", Default:="", pDescricao:="Pasta de Administração"))
                  If Len(Trim(sAux)) > 1 Then
                     Call WriteIniFile(gLocalReg, "Setup", "PATHSETUP", sAux)
                  End If
                  Call .SetRegPathSetup(gLocalReg)
                  Main
                  Exit Sub
               End If
               
               If gIDUSU = "" Then
                  End
               End If
               
               nAux = ReadIniFile(gLocalReg, "Config", "LOJA")
               .GetIniVars pCODSIS:=gCODSIS, pIDLOJA:=nAux, pIniFile:=gSetupFile, pAppPath:=App.Path
               '.SaveIniVars
               
               Screen.MousePointer = vbHourglass
               If gDebug Then MsgBox "Antes MDI.Show"
'xxx               MontaMenu
               MDI.Show
               Set MySplash = Nothing
                  
               On Error Resume Next
               Set .DefaultIcon = MDI.Icon
               '               'Verifica visualizção da versão
               '               Screen.MousePointer = vbHourglass
               '               s_VerificaVisualVersao
               
            End With
         Else
            End
         End If
      End With
   End If
   Exit Sub
TrataErro:
   If gIDUSU <> "" Then
      gIDUSU = ""
      Set MySplash = Nothing
      Splash.Show
      On Error Resume Next
   End If
   Call UnloadIni
   'Set XDb = Nothing
   Set Splash = Nothing
   
'MsgBox Err.Number & " - " & Err.Description   'ShowError("Sub Main()")
   Resume Next
End Sub
Public Sub InstaciaObj()
   If gDebug Then MsgBox "InstaciaObj"
   Dim TpErro  As String
   Dim Erro429 As Boolean

   On Error GoTo TrataErro

   TpErro = "DSActive"
   If gDebug Then MsgBox "Criou DsActive"


   TpErro = "SysA"
'   #If ComRef Then
'      Set Sys = New SysA.SetA
'   #Else
      Set Sys = CriarObjeto("SysA.SetA")
'   #End If
   Sys.LocalReg = gLocalReg
   Sys.CODSIS = gCODSIS
   If gDebug Then MsgBox "Criou SysA"

   TpErro = "CONEXAO"
   Set Splash = CriarObjeto("CONEXAO.Splash")
   If gDebug Then MsgBox "Criou SysA"

   Call MyInstaciaObj

   GoTo Saida
TrataErro:
   If Err = 429 Then
      Erro429 = True
      Resume Next
   Else
      ShowError "Instacia de Objetos [" & TpErro & "]"
   End If
Saida:
   If Erro429 Then
      Err = 429
      ShowError "Instacia de Objetos [" & TpErro & "]"
   End If
End Sub
Private Sub LimpaInstaciaObj()
   If gDebug Then MsgBox "LimpaInstaciaObj"
   On Error GoTo TrataErro
   
   Set Sys = Nothing
   Set Splash = Nothing
   
   Call MyLimpaInstaciaObj
   
   Exit Sub
TrataErro:
   ShowError "Limpa Instacia de Objetos"
End Sub
Private Sub UnloadIni()
   Dim n As Variant
   On Error Resume Next
   Unload Splash
   Set Splash = Nothing
   For Each n In Forms
      If UCase("FrmInicio") = UCase(n.Name) Or UCase("FrmSenha") = UCase(n.Name) Then
         Unload n
         Exit For
      End If
   Next
End Sub
Private Sub VerificaAutoInstall()
   Dim sArq       As String
   Dim sCODSIS    As String
   Dim sAppPath   As String
   Dim sExeName   As String
   Dim sCommand   As String
   Dim sPathExe   As String
   Dim sPathSetup As String
   Dim SetupIni   As String
   Dim sStatus    As String
   Dim LerIni     As Boolean
   Dim sPathSis   As String
   Dim sPathNewFile As String
   Dim sAux       As String
   
   On Error GoTo TrataErro
      
   '*************
   '* SetupFile = ReadIniFile(gLocalReg, "Setup", "Setup.Ini", App.Path)
   sCommand = Trim(UCase(Command$))
   sCODSIS = IIf(Trim(gCODSIS) = "", App.EXEName, gCODSIS)
   sAppPath = L_ResolvePathName(App.Path)
   sExeName = App.EXEName
   sPathExe = L_ResolvePathName(L_GetTag(sCommand, "SourcePath", ""))
   
   
   '*************
   '* PathSetup = "\\guarani\sistemas\admin\"
   SetupIni = sAppPath & gSetupFile
   If gDebug Then MsgBox "FileExists(" & SetupIni & ")"
   If Not L_ExisteArquivo(SetupIni) Then
      sPathSetup = L_ReadIniFile(gLocalReg, "Setup", "PATHSETUP")
   Else
      sPathSetup = sAppPath
   End If
   sPathSetup = L_ResolvePathName(sPathSetup)
   SetupIni = sPathSetup & gSetupFile
   
   
   If gDebug Then MsgBox "FileExists(" & SetupIni & ")"
   If L_ExisteArquivo(SetupIni) Then
      sStatus = L_ReadIniFile(SetupIni, "AutoInstall", "Status", "0")
      LerIni = (sStatus = "1" Or sStatus = "2")
      If gDebug Then MsgBox "LerIni = " & LerIni
      If LerIni Then
         sPathNewFile = L_ResolvePathName(L_ReadIniFile(SetupIni, "AutoInstall Files", "Path", ""))
         
         If Trim(sPathNewFile) <> "" Then
            
            sArq = "xLib.dll"
            sPathSis = L_GetRegisterDir(sArq, App.Path)
            If Trim(sPathSis) <> "" Then
               If gDebug Then
                  sAux = "File = " & sArq & vbNewLine
                  sAux = sAux & "sPathSis = " & sPathSis & vbNewLine
                  sAux = sAux & "sPathNewFile = " & sPathNewFile & vbNewLine
                  sAux = sAux & " ResolvePathName(" & sPathNewFile & ") = " & L_ResolvePathName(sPathNewFile) & vbNewLine
                  sAux = sAux & L_GetFileVersionNumber(sPathSis & sArq) & " < " & L_GetFileVersionNumber(sPathNewFile & sArq) & " = " & (L_GetFileVersionNumber(sPathSis & sArq) < L_GetFileVersionNumber(sPathNewFile & sArq)) & vbNewLine
                  MsgBox sAux
               End If
               If L_GetFileVersionNumber(sPathSis & sArq) < L_GetFileVersionNumber(sPathNewFile & sArq) Then
                  Call L_RegServer(sPathSis & sArq, False, False)
                  Call Kill(sPathSis & sArq)
                  Call FileCopy(sPathNewFile & sArq, sPathSis & sArq)
                  Call L_RegServer(sPathSis & sArq, True, False)
               End If
            End If
            
            sArq = "AutoInstall.dll"
            sPathSis = L_GetRegisterDir(sArq, App.Path)
            If Trim(sPathSis) <> "" Then
               If gDebug Then
                  sAux = "File = " & sArq & vbNewLine
                  sAux = sAux & "sPathSis = " & sPathSis & vbNewLine
                  sAux = sAux & "sPathNewFile = " & sPathNewFile & vbNewLine
                  sAux = sAux & " ResolvePathName(" & sPathNewFile & ") = " & L_ResolvePathName(sPathNewFile) & vbNewLine
                  sAux = sAux & L_GetFileVersionNumber(sPathSis & sArq) & " < " & L_GetFileVersionNumber(sPathNewFile & sArq) & " = " & (L_GetFileVersionNumber(sPathSis & sArq) < L_GetFileVersionNumber(sPathNewFile & sArq)) & vbNewLine
                  MsgBox sAux
               End If
               If L_GetFileVersionNumber(sPathSis & sArq) < L_GetFileVersionNumber(sPathNewFile & sArq) Then
                  Call L_RegServer(sPathSis & sArq, False, False)
                  Call Kill(sPathSis & sArq)
                  Call FileCopy(sPathNewFile & sArq, sPathSis & sArq)
                  Call L_RegServer(sPathSis & sArq, True, False)
               End If
            End If
            sArq = "SysA.dll"
            sPathSis = L_GetRegisterDir(sArq, App.Path)
            If Trim(sPathSis) <> "" Then
               If gDebug Then
                  sAux = "File = " & sArq & vbNewLine
                  sAux = sAux & "sPathSis = " & sPathSis & vbNewLine
                  sAux = sAux & "sPathNewFile = " & sPathNewFile & vbNewLine
                  sAux = sAux & " ResolvePathName(" & sPathNewFile & ") = " & L_ResolvePathName(sPathNewFile) & vbNewLine
                  sAux = sAux & L_GetFileVersionNumber(sPathSis & sArq) & " < " & L_GetFileVersionNumber(sPathNewFile & sArq) & " = " & (L_GetFileVersionNumber(sPathSis & sArq) < L_GetFileVersionNumber(sPathNewFile & sArq)) & vbNewLine
                  MsgBox sAux
               End If
               If L_GetFileVersionNumber(sPathSis & sArq) < L_GetFileVersionNumber(sPathNewFile & sArq) Then
                  Call L_RegServer(sPathSis & sArq, False, False)
                  Call Kill(sPathSis & sArq)
                  Call FileCopy(sPathNewFile & sArq, sPathSis & sArq)
                  Call L_RegServer(sPathSis & sArq, True, False)
               End If
            End If
         End If
      End If
   End If
   
   GoTo Saida
TrataErro:
   If Err <> 0 Then
MsgBox Err.Number & " - " & Err.Description & vbNewLine & vbNewLine & sPathSis & sArq     'ShowError("Sub Main()")
   Resume 0
End If
Saida:
End Sub
Public Sub UnloadSystem()
   On Error Resume Next
   Dim nObject  As Object
   
   '*********************
   '* Registra LogOut e Libera Licença
   '*********************
   If Splash Is Nothing Then
      Set Splash = CriarObjeto("CONEXAO.Splash")
   End If
   Call Splash.RegistraLogOut(Sys.IDUSU, Sys.CODSIS)
   
   
   '*********************
   '* Desconecta do Banco de Dados
   '*********************
   Call Sys.xdb.SrvDesconecta
   
   '*********************
   '* Libera Memória
   '*********************
   FreeLibrary m_hMod
   
   '*********************
   '* Limpar Colecao ClTelas e Forms
   '*********************
   For Each nObject In Forms
      Unload nObject
   Next
   If Not ClTelas Is Nothing Then
      For Each nObject In ClTelas
         ClTelas.Remove nObject.Index
      Next
      Set ClTelas = Nothing
   End If
      
   Set Sys = Nothing
   Set Splash = Nothing
End Sub
Private Sub MyRunAs()
   Dim sUserName     As String
   Dim sPassword     As String
   Dim sDomainName   As String
   Dim sCommandLine  As String
   Dim sCurrentDir   As String
   
   Dim sCommand      As String
   Dim sRegistry     As String
   Dim sArq       As String
   Dim sPathSis   As String
   Dim sResult    As String
   Dim sMsg       As String
   Dim sCODSIS    As String
   Dim sAppPath   As String
   Dim sExeName   As String
   Dim sPathSetup As String
   Dim sStatus    As String
   Dim LerIni     As Boolean
   Dim SetupIni   As String
   Dim SetupReg   As String
   
   Dim bErro      As Boolean
   Dim MyRegistro As Object
   
   
   On Error GoTo TrataErro
         
   sCommand = UCase(Command$)
   If InStr(sCommand, "/RUNAS") = 0 Then
      sCODSIS = IIf(Trim(gCODSIS) = "", App.EXEName, gCODSIS)
      sAppPath = L_ResolvePathName(App.Path)
      sExeName = App.EXEName
   
      If InStr(UCase(sAppPath), "\SISTEMAS\") <> 0 Then
         sAppPath = Environ("programfiles") & "\ClasseA\" & Replace(App.ProductName, " ", "") & "\"
      End If

      'sCommand = ""

      '*************
      '* Verificar Registro Local
      SetupReg = sAppPath & sExeName & ".reg"
      If L_ExisteArquivo(SetupReg) Then
         sRegistry = L_ReadIniFile(SetupReg, "General", "Registry", "")
      End If
         
      '*************
      '* PathSetup = "\\guarani\sistemas\admin\"
      SetupIni = sAppPath & gSetupFile
      If L_ExisteArquivo(SetupIni) Then
         sPathSetup = sAppPath
      Else
         sPathSetup = L_ReadIniFile(gLocalReg, "Setup", "PATHSETUP")
      End If
      sPathSetup = L_ResolvePathName(sPathSetup)
      
      SetupIni = sPathSetup & gSetupFile
      If L_ExisteArquivo(SetupIni) Then
         sStatus = L_ReadIniFile(SetupIni, "AutoInstall", "Status", "0")
         LerIni = (sStatus = "1" Or sStatus = "2")
         If LerIni Then
            If sRegistry <> L_ReadIniFile(SetupIni, "General", "Registry", "") Then
               sRegistry = L_ReadIniFile(SetupIni, "General", "Registry", "")
               Call L_WriteIniFile(SetupReg, "General", "Registry", sRegistry)
            End If
         End If
      End If
      
      'sCommand = ""
      If InStr(sCommand, "/RUNAS") = 0 And sRegistry <> "" Then
         sRegistry = L_Decrypt2(sRegistry)
         sUserName = L_GetTag(sRegistry, "USERNAME", "")
         sPassword = L_GetTag(sRegistry, "PASSWORD", "")
         sDomainName = L_GetTag(sRegistry, "DOMAINNAME", "")
         sCommandLine = sAppPath & sExeName & ".exe /RUNAS " & sCommand
         sCurrentDir = sAppPath
                           
         Set MyRegistro = CreateObject("CARegistro.RunAs")
         If sUserName = "" Or sPassword = "" Or sDomainName = "" Then
            sMsg = "Registro inválido! " & vbNewLine
            sMsg = sMsg & "Favor gerar chave de registro com usuário Administrador." & vbNewLine & vbNewLine
            sMsg = sMsg & "Caso você não tenha este privilégio entre em contato com Administrador da Rede."
            MsgBox sMsg, vbCritical & vbOKOnly, "Registro"
            
            Set MyRegistro = CreateObject("CARegistro.RunAs")
            MyRegistro.Show
            Set MyRegistro = Nothing
            End
         ElseIf sUserName <> "" And sPassword <> "" And sDomainName <> "" Then
            sMsg = MyRegistro.Execute(sUserName, sPassword, sDomainName, sCommandLine, sCurrentDir)
            If sMsg = "0" Then
               End
            End If
         End If
      End If
   End If
   GoTo Saida
TrataErro:
   MsgBox Err & "-" & Error
   bErro = True
   Resume Next
Saida:

End Sub
Function L_GetDepFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    L_GetDepFileVerStruct = False

    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = 32767

    On Error GoTo Failed

    iFile = FreeFile

    Open strFilename For Input Access Read Lock Read Write As #iFile

    ' Loop through each line, looking for the key
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If VBA.Left$(strLine, cchVersionKey) = strVersionKey Then
            ' We've found the version key.  Copy everything after the equals sign
            Dim strVersion As String

            strVersion = Mid$(strLine, cchVersionKey + 1)

            'Parse and store the version information
            L_PackVerInfo strVersion, sVerInfo

            L_GetDepFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend

    Close iFile
    Exit Function

Failed:
    L_GetDepFileVerStruct = False
End Function
Public Function L_GetFileVersion(ByVal strFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
    Dim sVerInfo As VERINFO
    Dim strVer As String

    On Error GoTo GFVError

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If

    '
    'Get the file version into a VERINFO struct, and then assemble a version string
    'from the appropriate elements.
    '
    If L_GetFileVerStruct(strFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
        strVer = Format$(sVerInfo.nMSHi) & "." & Format$(sVerInfo.nMSLo) & "."
        strVer = strVer & Format$(sVerInfo.nLSHi) & "." & Format$(sVerInfo.nLSLo)
        L_GetFileVersion = strVer
    Else
        L_GetFileVersion = ""
    End If

    Exit Function

GFVError:
    L_GetFileVersion = ""
    Err = 0
End Function
Private Function L_GetFileVersionNumber(pFilename As String) As Double
   Dim Pos  As Integer
   Dim nVer As Double
   Dim sVer As String
   Dim sAux As String
   Dim PosA As Integer
   
   On Error Resume Next
   
   sAux = ""
   PosA = 0
   
   sVer = L_GetFileVersion(pFilename)
   Pos = InStr(sVer, ".")
   If Pos <> 0 Then
      While Pos <> 0
         Pos = InStr(PosA + 1, sVer, ".")
         sAux = sAux & Right("000" + Mid(sVer, PosA + 1, IIf(Pos = 0, Len(sVer) + 1, Pos) - PosA - 1), 3)
         PosA = Pos
      Wend
   End If
   sAux = IIf(Trim(sAux) = "", "0", Trim(sAux))
   L_GetFileVersionNumber = Val(sAux)
End Function
Private Function L_GetFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
    Const strFIXEDFILEINFO$ = "\"

    Dim lVerSize As Long
    Dim lVerHandle As Long
    Dim lpBufPtr As Long
    Dim byteVerData() As Byte
    Dim fFoundVer As Boolean

    L_GetFileVerStruct = False
    fFoundVer = False

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If

    If fIsRemoteServerSupportFile Then
        L_GetFileVerStruct = L_GetRemoteSupportFileVerStruct(strFilename, sVerInfo)
        fFoundVer = True
    Else
        '
        'Get the size of the file version info, allocate a buffer for it, and get the
        'version info.  Next, we query the Fixed file info portion, where the internal
        'file version used by the Windows VerInstallFile API is kept.  We then copy
        'the fixed file info into a VERINFO structure.
        '
        lVerSize = GetFileVersionInfoSize(strFilename, lVerHandle)
        If lVerSize > 0 Then
            ReDim byteVerData(lVerSize)
            If GetFileVersionInfo(strFilename, lVerHandle, lVerSize, byteVerData(0)) <> 0 Then ' (Pass byteVerData array via reference to first element)
                If VerQueryValue(byteVerData(0), strFIXEDFILEINFO & "", lpBufPtr, lVerSize) <> 0 Then
                    lmemcpy sVerInfo, lpBufPtr, lVerSize
                    fFoundVer = True
                    L_GetFileVerStruct = True
                End If
            End If
        End If
    End If

    If Not fFoundVer Then
        '
        ' We were unsuccessful in finding the version info from the file.
        ' One possibility is that this is a dependency file.
        '
        If UCase(L_Extension(strFilename)) = "DEP" Then
            L_GetFileVerStruct = L_GetDepFileVerStruct(strFilename, sVerInfo)
        End If
    End If
End Function
Function L_GetRemoteSupportFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = 32767

    On Error GoTo Failed

    iFile = FreeFile

    Open strFilename For Input Access Read Lock Read Write As #iFile

    ' Loop through each line, looking for the key
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If VBA.Left$(strLine, cchVersionKey) = strVersionKey Then
            ' We've found the version key.  Copy everything after the equals sign
            Dim strVersion As String

            strVersion = Mid$(strLine, cchVersionKey + 1)

            'Parse and store the version information
            L_PackVerInfo strVersion, sVerInfo

            'Convert the format 1.2.3 from the .VBR into
            '1.2.0.3, which is really want we want
            sVerInfo.nLSLo = sVerInfo.nLSHi
            sVerInfo.nLSHi = 0

            L_GetRemoteSupportFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend

    Close iFile
    Exit Function

Failed:
    L_GetRemoteSupportFileVerStruct = False
End Function
Private Function L_ExisteArquivo(ByVal strPathName As String) As Boolean
   Dim intFileNum   As Integer
   Dim sArq         As String
   Dim sPath        As String
   
   On Error Resume Next
   
   strPathName = Trim(strPathName)
   strPathName = Replace(strPathName, """", "")
   
   Call L_GetNameFromPath(strPathName, sPath)
   sArq = Mid(strPathName, Len(sPath) + 1)
   If Len(Dir(strPathName, vbArchive)) > 4 And sPath <> "" And sArq <> "" Then
      L_ExisteArquivo = IIf(Err = 0, True, False)
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
      L_ExisteArquivo = IIf(Err = 0, True, False)
      Close intFileNum
   End If
   
   Err = 0
End Function
Public Function L_Extension(ByVal strFilename As String) As String
    Dim intPos As Integer

    L_Extension = ""

    intPos = Len(strFilename)

    Do While intPos > 0
        Select Case Mid$(strFilename, intPos, 1)
            Case "."
                L_Extension = Mid$(strFilename, intPos + 1)
                Exit Do
            Case "\", "/"
                Exit Do
            'End Case
        End Select

        intPos = intPos - 1
    Loop
End Function
Private Function L_GetNameFromPath(PathFile As String, Optional ByRef PathReturn As String) As String
   Dim i As Integer
   
   i = InStrRev(PathFile, "\")
   i = IIf(i = 0, 1, i)
   If PathReturn = "1" Then
      L_GetNameFromPath = VBA.Left$(PathFile, i)
   Else
      L_GetNameFromPath = VBA.Mid$(PathFile, Len(VBA.Left$(PathFile, i)) + 1)
   End If
   PathReturn = L_ResolvePathName(VBA.Left$(PathFile, i))
End Function
Public Function L_RegServer(sServerPath As String, Optional fRegister = True, Optional fMsg As Boolean = True, Optional isActivexExe As Boolean = False) As Boolean
   Dim hMod As Long               ' module handle
   Dim lpfn As Long                  ' reg/unreg function address
   Dim sCmd As String             ' msgbox string
   Dim lpThreadID As Long        ' unused, receives the thread ID
   Dim hThread As Long            ' thread handle
   Dim fSuccess As Boolean     ' if things worked
   Dim dwExitCode As Long      ' thread's exit code if it doesn't finish
   
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
       
       ' Wait 10 secs for the thread to finish (the function may take a while...)
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
   
   If fMsg Then
      If fSuccess Then
        MsgBox "Successfully " & sCmd & "ed " & sServerPath   ' past tense
        L_RegServer = True
      Else
        MsgBox "Failed To " & sCmd & " " & sServerPath, vbExclamation
      End If
   End If
End Function
Public Function L_ResolvePathName(ByVal sPath As String) As String
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

   L_ResolvePathName = sPath
End Function
Sub L_PackVerInfo(ByVal strVersion As String, sVerInfo As VERINFO)
    Dim intOffset As Integer
    Dim intAnchor As Integer

    On Error GoTo PVIError

    intOffset = InStr(strVersion, ".")
    If intOffset = 0 Then
        sVerInfo.nMSHi = Val(strVersion)
        GoTo PVIMSLo
    Else
        sVerInfo.nMSHi = Val(VBA.Left$(strVersion, intOffset - 1))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, ".")
    If intOffset = 0 Then
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSHi
    Else
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, ".")
    If intOffset = 0 Then
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSLo
    Else
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, ".")
    If intOffset = 0 Then
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor))
    Else
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
    End If

    Exit Sub

PVIError:
    sVerInfo.nMSHi = 0
PVIMSLo:
    sVerInfo.nMSLo = 0
PVILSHi:
    sVerInfo.nLSHi = 0
PVILSLo:
    sVerInfo.nLSLo = 0
End Sub
Public Function L_ProcuraArquivo(ByVal pPath As String, ByVal pArq As String, Optional pAdmin As Boolean = True) As String
   Dim sAux    As String
   Dim sPath   As String
   Dim bAchou  As Boolean
   Dim sPath0  As String
   Dim i       As Integer
   Dim nVezes  As Integer
   Dim bAdmin  As Boolean
   
   On Error GoTo TrataErro
   nVezes = 10000
   
   
   sAux = L_ResolvePathName(pPath)
   ChDir sAux
   sPath0 = sAux
   
   bAchou = L_ExisteArquivo(sAux & pArq)
   If bAchou Then
      sPath = pPath
   Else
      sAux = Dir(sAux, vbDirectory)
      While sAux <> ""
         sAux = Dir(Attributes:=VbFileAttribute.vbDirectory)
         bAdmin = IIf(UCase(sAux) = "ADMIN", pAdmin, True)
         
         If InStr(sAux, ".") = 0 And sAux <> "" And bAdmin Then
            If (GetAttr(pPath & sAux) And vbDirectory) = vbDirectory Then
               sAux = L_ResolvePathName(pPath & sAux)
               bAchou = L_ExisteArquivo(sAux & pArq)
               
               If bAchou Then
                  sPath = sAux
                  sAux = ""
               Else
                  sPath = L_ProcuraArquivo(sAux, pArq)
                  If sPath = "" Then
                     '*********************
                     '* Retorna ao diretório anterior
                     ChDir sPath0
                     If sAux <> Dir(pPath, vbDirectory) Then
                        i = 0
                        While i < nVezes
                           i = i + 1
                           If sAux = L_ResolvePathName(pPath & Dir(Attributes:=vbDirectory)) Then
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
   L_ProcuraArquivo = sPath
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
Public Function L_ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, Optional DefaultValue As String) As String
   Dim strBuffer As String
   
   strBuffer = Space$(255)
   If GetPrivateProfileString(strSection, strKey, "", strBuffer, 255, strIniFile) > 0 Then
      If InStr(strBuffer, Chr$(0)) > 0 Then
        strBuffer = VBA.Left$(strBuffer, InStr(strBuffer, Chr$(0)) - 1)
      End If
      L_ReadIniFile = RTrim$(strBuffer)
   Else
      L_ReadIniFile = DefaultValue
   End If
End Function
Private Function L_WriteIniFile(ByVal strIniFile As String, strSection As String, strKey As String, strValue As String) As Boolean
   Dim intLen As Integer
   
   If Not L_ExisteArquivo(strIniFile) Then
      intLen = L_AbrirTxt(strIniFile)
      Close #intLen
   End If
   intLen = 0
   intLen = WritePrivateProfileString(strSection, strKey, strValue, strIniFile)
   L_WriteIniFile = (intLen > 0)
End Function
Private Function L_AbrirTxt(Arq As String) As Integer
   Dim Hnd As Integer
  
   On Error GoTo CopyErr
   Call L_ExcluirArquivo(Arq)
   L_AbrirTxt = FreeFile()
   Open Arq For Output As #L_AbrirTxt
Exit Function
CopyErr:
  Select Case Err
     Case 55: Err = 0
     Case Else: MsgBox Err.Number & " - " & Err.Description
  End Select
End Function
Private Function L_ExcluirArquivo(File As String, Optional ViewError As Boolean = True) As Boolean
   If L_ExisteArquivo(File) Then
      On Error GoTo Fim
      Call Kill(File)
   End If
   L_ExcluirArquivo = Not L_ExisteArquivo(File)
   Exit Function
Fim:
   If ViewError Then
      MsgBox Err & " - " & Error
   End If
End Function
Public Function L_GetRegisterDir(ByVal sArq As String, Optional AppPath As String) As String
   Dim sRoot      As String
   Dim sPath      As String
   Dim bAchou     As Boolean
   Dim sAux       As String
   
   
   sPath = L_ResolvePathName(AppPath)
   bAchou = L_ExisteArquivo(AppPath & sArq)
         
   '*****************
   '* \Windows\SysWOW64
   If Not bAchou Then
      sRoot = L_ResolvePathName(L_GetSpecialFolder(41)) 'CSIDL_SYSTEM64))
      
      sPath = L_ResolvePathName(sRoot & gSubFolder)
      bAchou = L_ExisteArquivo(sPath & sArq)
      If Not bAchou Then
         sPath = L_ProcuraArquivo(sPath, sArq, False)
         If sPath <> "" Then bAchou = L_ExisteArquivo(sPath & sArq)
         If Not bAchou Then
            sPath = sRoot
            bAchou = L_ExisteArquivo(sPath & sArq)
         End If
      End If
      '*****************
      '* \Windows\System32
      If Not bAchou Then
         sRoot = L_ResolvePathName(L_GetSpecialFolder(37)) 'CSIDL_SYSTEM32))
         
         sPath = L_ResolvePathName(sRoot & gSubFolder)
         bAchou = L_ExisteArquivo(sPath & sArq)
         If Not bAchou Then
            sPath = L_ProcuraArquivo(sPath, sArq, False)
            If sPath <> "" Then bAchou = L_ExisteArquivo(sPath & sArq)
            If Not bAchou Then
               sPath = sRoot
               bAchou = L_ExisteArquivo(sPath & sArq)
            End If
         End If
      
         '*****************
         '* \Program Files
         If Not bAchou Then
            sRoot = L_ResolvePathName(Environ("PROGRAMFILES")) 'GetSpecialFolder(CSIDL_PROGRAM_FILES))
            
            sPath = L_ResolvePathName(sRoot & gSubFolder)
            bAchou = L_ExisteArquivo(sPath & sArq)
            If Not bAchou Then
               sPath = L_ProcuraArquivo(sPath, sArq, False)
               If sPath <> "" Then bAchou = L_ExisteArquivo(sPath & sArq)
               If Not bAchou Then
                  sPath = sRoot
                  bAchou = L_ExisteArquivo(sPath & sArq)
               End If
            End If
   
      
            '*****************
            '* \Program Files\Common
            If Not bAchou Then
               sRoot = L_ResolvePathName(Environ("CommonProgramFiles")) '""GetSpecialFolder(CSIDL_COMMON))
               
               sPath = L_ResolvePathName(sRoot & gSubFolder)
               bAchou = L_ExisteArquivo(sPath & sArq)
               If Not bAchou Then
                  sPath = L_ProcuraArquivo(sPath, sArq, False)
                  If sPath <> "" Then bAchou = L_ExisteArquivo(sPath & sArq)
                  If Not bAchou Then
                     sPath = sRoot
                     bAchou = L_ExisteArquivo(sPath & sArq)
                  End If
               End If
      
               '*****************
               '* \Windows
               If Not bAchou Then
                  sRoot = Environ("SystemRoot") 'L_GetSpecialFolder(CSIDL_WINDOWS)
                  
                  sPath = L_ResolvePathName(sRoot & gSubFolder)
                  bAchou = L_ExisteArquivo(sPath & sArq)
                  If Not bAchou Then
                     sPath = L_ProcuraArquivo(sPath, sArq, False)
                     If sPath <> "" Then bAchou = L_ExisteArquivo(sPath & sArq)
                     If Not bAchou Then
                        sPath = sRoot
                        bAchou = L_ExisteArquivo(sPath & sArq)
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   sPath = IIf(bAchou, sPath, "")
   
   L_GetRegisterDir = L_ResolvePathName(sPath)
End Function
Public Function L_GetSpecialFolder(CSIDL As Long) As String
    Dim sPath  As String
    Dim IDL    As ITEMIDLIST
    Dim nhWnd  As Long
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
   If CSIDL = 1 Then 'CSIDL_TEMPORARY
      L_GetSpecialFolder = L_ResolvePathName(GetTempFolder)
   Else
      L_GetSpecialFolder = ""
      If SHGetSpecialFolderLocation(nhWnd, CSIDL, IDL) = 0 Then
         '
         ' Get the path from the ID list, and return the folder.
         '
         sPath = Space$(260)
         If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            L_GetSpecialFolder = L_ResolvePathName(Left$(sPath, InStr(sPath, vbNullChar) - 1) & "")
         End If
     End If
   End If
End Function
Private Function L_GetTag(ByRef pControle As Variant, ByVal pNome As String, Optional pPadrao As String) As String
   Dim PosIni  As Long
   Dim PosFim  As Long
   Dim StrTAG  As String
   Dim i       As Integer
   
   On Error GoTo Saida
   
   pNome = "|" & Trim(pNome) & "="
   
   If UCase(TypeName(pControle)) = "STRING" Then
      StrTAG = pControle
   Else
      StrTAG = pControle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(pNome))
   If PosIni > 0 Then
      PosIni = PosIni + Len(Trim(pNome))
      PosFim = InStr(PosIni, StrTAG, "|")
      i = 0
      While Mid(StrTAG, PosIni + i, 1) = "|"
         i = i + 1
      Wend
      If i > 0 Then
         PosFim = InStr(PosIni + (i - 1), StrTAG, "|")
      End If
      PosFim = IIf(PosFim = 0, Len(StrTAG), PosFim - 1)
      StrTAG = Mid(StrTAG, PosIni, PosFim - PosIni + 1)
   Else
      StrTAG = ""
   End If
   L_GetTag = StrTAG
Saida:
   If StrTAG = "" Then
      L_GetTag = pPadrao
   End If
End Function
Private Function L_Decrypt2(ByVal Password As String, Optional Key As String) As String
   Dim P    As String
   Dim B    As String
   Dim S    As String
   Dim i    As Integer
   Dim j    As Integer
   Dim A1   As Integer
   Dim A2   As Integer
   Dim A3   As Integer
   
   If Trim(Key) = "" Then Key = L_Encrypt2("231072150500", "DIO")
   j = 1
   For i = 1 To Len(Key)
     P = P & Asc(Mid$(Key, i, 1))
   Next
   For i = 1 To Len(Password) Step 2
     A1 = Asc(Mid$(P, j, 1))
     j = j + 1
     If j > Len(P) Then j = 1
     B = Mid$(Password, i, 2)
     A3 = Val("&H" + B)
     A2 = A1 Xor A3
     S = S + Chr$(A2)
   Next
   L_Decrypt2 = Mid(S, 3)
End Function
Private Function L_Encrypt2(ByVal Password As String, Optional Key As String) As String
   Dim P    As String
   Dim B    As String
   Dim S    As String
   Dim i    As Integer
   Dim j    As Integer
   Dim A1   As Integer
   Dim A2   As Integer
   Dim A3   As Integer
   
   Dim sAux As String
   
   If Len(Trim(Password)) < 2 Then
      sAux = "PI"
   Else
      sAux = VBA.Right$(Password, 1) & VBA.Left$(Password, 1)
   End If
   Password = sAux & Password
   If Trim(Key) = "" Then Key = L_Encrypt2("231072150500", "DIO")
   j = 1
   For i = 1 To Len(Key)
     P = P & Asc(Mid$(Key, i, 1))
   Next
    
   For i = 1 To Len(Password)
     A1 = Asc(Mid$(P, j, 1))
     j = j + 1
     If j > Len(P) Then j = 1
     A2 = Asc(Mid$(Password, i, 1))
     A3 = A1 Xor A2
     B = Hex$(A3)
     If Len(B$) < 2 Then B$ = "0" + B
     S = S + B
   Next
   L_Encrypt2 = S
End Function
Private Sub AtualizaBD()
   Dim sPathSetup As String
   Dim sFileZip   As String
   Dim Sql        As String
   Dim MyRs       As Object
   Dim sArqAtu    As String
   Dim nAtu       As Long
   Dim i          As Integer
   Dim bNum       As Boolean
   Dim sFileSql   As String
   Dim sPathSql   As String
   Dim bAux       As Boolean
   Dim bExiste    As Boolean
   Dim bDate      As Boolean
      
   On Error Resume Next
      
   '*****
   '* excluir Arquivo de Log
   Call ExcluirArquivo(App.Path & "\" & "ExeScr.log")
      
   '*****
   '* Localiza Script.zia
   sPathSetup = ResolvePathName(ReadIniFile(gLocalReg, "Setup", "PATHSETUP", ""))
   'sPathSql = Environ("TEMP") & "\Sql"
   sPathSql = Sys.PathTmp & "Sql"
   
   sFileZip = Sys.xdb.dbName
   If InStr(sFileZip, "_") <> 0 Then
      sFileZip = Mid(sFileZip, 1, InStr(sFileZip, "_") - 1)
   End If
   sFileZip = sFileZip & "REV"
   
   If ExisteArquivo(sPathSetup & sFileZip & ".zia") Then
      sFileZip = sFileZip & ".zia"
   ElseIf ExisteArquivo(sPathSetup & sFileZip & ".zip") Then
      sFileZip = sFileZip & ".zip"
   Else
      sFileZip = ""
   End If
   If sFileZip <> "" Then
      '*****
      '* Verifica Versão do Banco
      Sql = "Select IDBD, DSCBD, VSBD, ATUBD, DTATU"
      Sql = Sql & ", ARQATU "
      Sql = Sql & " From VERSAOBD"
      If Sys.xdb.AbreTabela(Sql, MyRs) Then
         sArqAtu = MyRs("ARQATU") & ""
         'nAtu = MyRs("VSBD") & ""
         nAtu = MyRs("ATUBD") & ""
      End If
      'bDate = CDate(FileDateTime("C:\Program Files (x86)\ClasseA\Admin\Dll\G3RREV.zia")) > CDate(MyRs("DTATU"))
      bDate = CDate(FileDateTime(sPathSetup & sFileZip)) > CDate(MyRs("DTATU"))
      If bDate Then
         If ExisteArquivo(sPathSql & "\*.*") Then
            Kill sPathSql & "\*.Sql"
         End If
      
         Call Unzip(sPathSetup, sFileZip, sPathSql & "\", False)
         'Call ExcluirArquivo(sPathSetup & sFileZip)
      End If
      Set MyRs = Nothing
            
      sArqAtu = IIf(UCase(sArqAtu) = "REV" & StrZero(nAtu, 2) & ".SQL", "Rev" & StrZero(nAtu + 1, 2) & ".sql", sArqAtu)
      nAtu = nAtu + 1
      If Not ExisteArquivo(sPathSql & "\" & sArqAtu) Then
         If bDate Then
            Sql = "Update VERSAOBD Set DTATU=" & SqlDate(FileDateTime("C:\Program Files (x86)\ClasseA\Admin\Dll\G3RREV.zia"))
            Call Sys.xdb.Executa(Sql)
         End If
      Else
         While ExisteArquivo(sPathSql & "\" & sArqAtu)
            For i = 1 To Len(sArqAtu)
               bNum = IsNumeric(Mid(sArqAtu, i, 1))
               If bNum Then Exit For
            Next
            If bNum Then
               sArqAtu = Mid(sArqAtu, 1, i - 1) & StrZero(nAtu, 2) & ".sql"
            Else
               Exit Sub
            End If
            
            '*****
            '* Executa Atualização
            If ExisteArquivo(sPathSql & "\" & sArqAtu) Then
               Call ExecuteScript(Sys.xdb, sPathSql & "\" & sArqAtu)
            End If
            
            nAtu = nAtu + 1
            sArqAtu = IIf(sArqAtu = "", "REV00.sql", sArqAtu)
         Wend
      End If
      '******
      '* Verifica se o Banco é Local ou Remoto
      bAux = (Sys.xdb.Server <> Sys.xdb.ServerName("[Remote]"))
      If gIDUSU = "DIO" Then
'         If ExibirPergunta("Atualiza Menu e Pesquisas?", "Acesso Restrito", False) = vbYes Then
'            bAux = True
'         End If
      End If
      If bAux And bDate Then
         '*****
         '* Executa Menu
         sPathSql = Sys.PathTmp & "Sql"
         'If Not Sys.XDB.AbreTabela("Select Distinct ALTERSTAMP From MODULO") Then
         If Sys.xdb.Alias <> "WEB" Then
            DoEvents
            sFileSql = Dir(sPathSql & "\*InsertMenu.Sql")
            If ExisteArquivo(sPathSql & "\" & sFileSql) Then
               If UCase(Mid(Sys.ExePath, 1, 12)) = "C:\SISTEMAS\" Then sPathSql = App.Path & "\Script"
               Call ExecuteScript(Sys.xdb, sPathSql & "\" & sFileSql)
            End If
         End If
         '*****
         '* Executa Pesquisas
         sPathSql = Sys.PathTmp & "Sql"
         'If Not Sys.XDB.AbreTabela("Select Distinct ALTERSTAMP From GPESQUISA") Then
         If Sys.xdb.Alias <> "WEB" Then
            sFileSql = Dir(sPathSql & "\*Pesquisas.Sql")
            If ExisteArquivo(sPathSql & "\" & sFileSql) Then
               If UCase(Mid(Sys.ExePath, 1, 12)) = "C:\SISTEMAS\" Then sPathSql = App.Path & "\Script"
               Call ExecuteScript(Sys.xdb, sPathSql & "\" & sFileSql)
            End If
         End If
      End If
      '*****
      '* Apagar pasta
      Kill sPathSql & "\*.Sql"
      'Call ExcluirArquivo(sPathSql & "\*.*")
      'Call ExcluirDiretorio(sPathSql)
   End If
End Sub
Private Sub DesativarProgramas()
   On Error Resume Next
'   If ProgramaAtivo("CABio") Then Call FecharPrograma("CABio")
'Exit Sub
   Dim sArq As String
   Dim sPath As String
   
   If InStr(App.Path, "\Sistemas\") <> 0 Then
      sPath = Environ("programfiles") & "\ClasseA\Projeto3R\"
   Else
      sPath = App.Path & "\"
   End If
   sArq = Dir$(sPath & "*.exe", vbNormal)
   Set cProgs = New Collection
   Do While sArq <> ""
      If UCase(sArq) <> UCase(App.EXEName & ".exe") And UCase(sArq) <> UCase(App.Title & ".exe") And sArq <> "P3R.exe" Then
         sArq = Mid(sArq, 1, Len(sArq) - 4)
         If ProgramaAtivo(sArq) Then
            cProgs.Add sPath & sArq & ".exe", sArq
            Call FecharPrograma(sArq)
         End If
      End If
      sArq = Dir$
   Loop
End Sub
Private Sub AtivarProgramas()
   Dim n As Variant
   If Not cProgs Is Nothing Then
      For Each n In cProgs
         Call SincShell(CStr(n), vbHide, False)
      Next
   End If
End Sub
