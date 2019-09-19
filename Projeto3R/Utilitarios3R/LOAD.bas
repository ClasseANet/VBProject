Attribute VB_Name = "LOAD"
Option Explicit
'============================================================
'============================================================
Public Type POINTAPI   ' pt
   x As Long
   y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
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
Global gIDUSU        As String
'============================================================
'============================================================
Public gCaption1  As String
Public gCaption2  As String
Public gCaption3  As String
'============================================================
'============================================================
#If ComRef Then
   Global Sys     As SysA.SetA
   Global XDb     As XBANCO01.DS_BANCO
   Global Splash  As Conexao.Splash
   Global DsAuto  As DSACTIVE.AutoInstall
   Global DsDsr   As DSACTIVE.DSR
   Global DsMsg   As DSACTIVE.MENSAGEM
   Global DsLoad  As DSACTIVE.DS_LOAD
#Else
   Global Sys     As Object
   Global XDb     As Object
   Global XDbMaua As Object
   Global Splash  As Object
   Global DsAuto  As Object
   Global DsDsr   As Object
   Global DsMsg   As Object
   Global DsLoad  As Object
#End If
Global ClTelas As Collection
'============================================================

Public Sub Main()
   'Dim i As Integer: For i = 1 To 100: Debug.Print Environ(i): Next
    Dim sCommand As String
    
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   
   sCommand = Trim(UCase(Command$))
   gLocalReg = L_ResolvePathName(App.Path) & App.EXEName & ".reg"
   gSetupFile = "SETUP.INI"
   gDebug = (InStr(sCommand, "DEBUG") <> 0)
   
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
        gCODSIS = Mid$(sCommand, 9)
      End If
   End If
   
   If gCODSIS = "" Then
      MyLoadgCODSIS
   End If
   
   Call LimpaInstaciaObj
   Call AutoInstall
   Call InstaciaObj
   
   '*********
   '* Testa se já existe uma cópia da aplicação rodando e define formato Data e número.
   If Not DsLoad Is Nothing And gIDUSU = "" Then
      DsLoad.Aplic = App
      If gDebug Then MsgBox "DsLoad.SetFormat"
      Call DsLoad.SetFormat
      If gDebug Then MsgBox "DsLoad.Ativa"
      If DsLoad.Ativa Then
         If gDebug Then MsgBox "End"
         End
      End If
   End If
   Set DsLoad = Nothing
   
   If gDebug Then MsgBox "Entrar ExibeSenha"
   Call ExibeSenha
   Exit Sub
TrataErro:
   Call UnloadIni
   Set XDb = Nothing
   Set Splash = Nothing
   
MsgBox Err.Number & " - " & Err.Description   'ShowError("Sub Main()")
   Resume Next
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
   Set AutoIn = Nothing
End Sub
Public Sub ExibeSenha(Optional pTrocaConexao As Boolean = False, Optional pxDb As Object, Optional pIDUSU As String)

   On Error GoTo TrataErro

   If gDebug Then MsgBox "Em ExibeSenha"
   If Splash Is Nothing Then
      MsgBox "Tela de senha não foi criada.", vbInformation
   Else
      If gDebug Then MsgBox "Splash Not Nothing"
      With Splash
         .DebugSys = False
         .CODSIS = gCODSIS
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
            
         If gDebug Then MsgBox "gIDUSU=" & gIDUSU
         If gIDUSU = "" Then
            If gDebug Then MsgBox "Show"
            .Show
         Else
            If gDebug Then MsgBox "Conectar"
            .IDUSU = gIDUSU
            .Conectar
         End If
         
         Dim dHHIni As Date
         Dim nTimeout As Long
         nTimeout = 15 '*Segundos
         dHHIni = CDate(Format(Now, "hh:mm:ss"))
         
         If gDebug Then MsgBox "Conectado=" & .Conectado
         Do While Not .Conectado
            If .Cancelado Then
               If pTrocaConexao Then
                  Exit Do
               Else
                  End
               End If
            End If
            DoEvents
            If gIDUSU <> "" Then '*Se não exibir senha, verificar tempo de conexão
               If DateDiff("s", dHHIni, CDate(Format(Now, "hh:mm:ss"))) > nTimeout Then
                  MsgBox "Tempo de conexão expirado!", vbInformation, "Conexão"
                  End
               End If
            End If
         Loop
         
         
         DoEvents
         Screen.MousePointer = vbHourglass
         If gDebug Then MsgBox "pTrocaConexao Or .Conectado" & (.Conectado Or pTrocaConexao)
         If .Conectado Or pTrocaConexao Then
            If .Cancelado Then
               Set XDb = pxDb
               gIDUSU = pIDUSU
            Else
               Set XDb = .XDb
               gIDUSU = .IDUSU
            End If
            
            If gDebug Then MsgBox "IDUSU : " & gIDUSU & vbNewLine & "CODSIS : " & gCODSIS
            
            With Sys
               Set .XDb = XDb
               .IDUSU = Trim(gIDUSU)
               .CODSIS = gCODSIS
               .LocalReg = gLocalReg
               '********
               '* Se pasta de Administração alterou re-executa Sistema.
               If gDebug Then MsgBox ".PathSetup : " & .PathSetup & vbNewLine & "ReadIniFile(gLocalReg, 'Setup', 'PATHSETUP', '') : " & DsAuto.ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "") & vbNewLine & " Reexecuta : " & (.PathSetup <> ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "") And .PathSetup <> "")
               If ResolvePathName(.PathSetup) <> ResolvePathName(ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "")) And .PathSetup <> "" Then
                  
                  Dim sAux As String
                  sAux = ResolvePathName(.GetParam(pCODPARAM:="PATHSETUP", pCODSIS:="GLOBAL", Default:="", pDescricao:="Pasta de Administração"))
                  If Trim(sAux) <> "" Then
                     Call WriteIniFile(gLocalReg, "Setup", "PATHSETUP", sAux)
                  End If
                  Call .SetRegPathSetup(gLocalReg)
                  Main
                  Exit Sub
               End If
               
               If gIDUSU = "" Then End
               .GetIniVars pCODSIS:=gCODSIS, pIniFile:=gSetupFile, pAppPath:=App.Path
               '.SaveIniVars
               
               Screen.MousePointer = vbHourglass
               If gDebug Then MsgBox "Antes MDI.Show"
'xxx               MontaMenu
               MDI.Show

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
      Splash.Show
      On Error Resume Next
   End If
   Call UnloadIni
   Set XDb = Nothing
   Set Splash = Nothing
   
MsgBox Err.Number & " - " & Err.Description   'ShowError("Sub Main()")
   Resume Next
End Sub
Private Sub InstaciaObj()
   If gDebug Then MsgBox "InstaciaObj"
   Dim TpErro  As String
   Dim Erro429 As Boolean

   On Error GoTo TrataErro

   TpErro = "XBanco"
   Set XDb = CriarObjeto("XBANCO01.DS_BANCO")
   If gDebug Then MsgBox "Criou xBanco"

   TpErro = "DSActive"
   Set DsAuto = CriarObjeto("DSACTIVE.AutoInstall")
   Set DsDsr = CriarObjeto("DSACTIVE.DSR")
   Set DsLoad = CriarObjeto("DSACTIVE.DS_LOAD")
   Set DsMsg = CriarObjeto("DSACTIVE.MENSAGEM")
   If gDebug Then MsgBox "Criou DsActive"

   TpErro = "SysA"
   Set Sys = CriarObjeto("SysA.SetA")
   Sys.LocalReg = gLocalReg
   Set Sys.XDb = XDb
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
      DsMsg.ShowError "Instacia de Objetos [" & TpErro & "]"
   End If
Saida:
   If Erro429 Then
      Err = 429
      DsMsg.ShowError "Instacia de Objetos [" & TpErro & "]"
   End If
End Sub
Private Sub LimpaInstaciaObj()
   If gDebug Then MsgBox "LimpaInstaciaObj"
   On Error GoTo TrataErro
   
   Set XDb = Nothing
   Set Sys = Nothing
   Set Splash = Nothing
   Set DsAuto = Nothing
   Set DsDsr = Nothing
   Set DsLoad = Nothing
   Set DsMsg = Nothing
   
   Call MyLimpaInstaciaObj
   
   Exit Sub
TrataErro:
   DsMsg.ShowError "Limpa Instacia de Objetos"
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
   
   Set DsAuto = CreateObject("DSACTIVE.AutoInstall")
   Set DsDsr = CreateObject("DSACTIVE.DSR")
   
   '*************
   '* SetupFile = ReadIniFile(gLocalReg, "Setup", "Setup.Ini", App.Path)
   sCommand = Trim(UCase(Command$))
   sCODSIS = IIf(Trim(gCODSIS) = "", App.EXEName, gCODSIS)
   sAppPath = L_ResolvePathName(App.Path)
   sExeName = App.EXEName
   sPathExe = L_GetTag(sCommand, "SourcePath", "")
   
   
   '*************
   '* PathSetup = "\\guarani\sistemas\admin\"
   SetupIni = sAppPath & gSetupFile
   If gDebug Then MsgBox "FileExists(" & SetupIni & ")"
   If Not L_ExisteArquivo(SetupIni) Then
      sPathSetup = L_ReadIniFile(gLocalReg, "Setup", "PATHSETUP")
   Else
      sPathSetup = sAppPath
   End If
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
            sPathSis = DsAuto.GetRegisterDir(sArq, App.Path)
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
                  Call DsDsr.RegServer(sPathSis & sArq, False, False)
                  Call Kill(sPathSis & sArq)
                  Call FileCopy(sPathNewFile & sArq, sPathSis & sArq)
                  Call DsDsr.RegServer(sPathSis & sArq, True, False)
               End If
            End If
            
            sArq = "AutoInstall.dll"
            sPathSis = DsAuto.GetRegisterDir(sArq, App.Path)
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
                  Call DsDsr.RegServer(sPathSis & sArq, False, False)
                  Call Kill(sPathSis & sArq)
                  Call FileCopy(sPathNewFile & sArq, sPathSis & sArq)
                  Call DsDsr.RegServer(sPathSis & sArq, True, False)
               End If
            End If
            If Now() <= "21\11\2008" Then
               Call DsDsr.RegServer(sPathSis & "xLib.dll", True, False)
               Call DsDsr.RegServer(sPathSis & "AutoInstall.dll", True, False)
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
   Call Sys.XDb.SrvDesconecta
   Call XDb.SrvDesconecta
   
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
   Set XDb = Nothing
   Set XDbMaua = Nothing
   Set Splash = Nothing
   Set DsAuto = Nothing
   Set DsDsr = Nothing
   Set DsMsg = Nothing
   Set DsLoad = Nothing
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
   
   sCODSIS = IIf(Trim(gCODSIS) = "", App.EXEName, gCODSIS)
   sAppPath = L_ResolvePathName(App.Path)
   sExeName = App.EXEName
   
   If UCase(Mid(sAppPath, 1, Len("C:\SISTEMAS"))) = "C:\SISTEMAS" Then
      sAppPath = Environ("programfiles") & "\ClasseA\Producao\"
   End If

   'sCommand = ""
   If InStr(sCommand, "/RUNAS") = 0 Then
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
Private Function L_GetFileVersionNumber(pFilename As String) As Double
   Dim Pos  As Integer
   Dim nVer As Double
   Dim sVer As String
   Dim sAux As String
   Dim PosA As Integer
   
   On Error Resume Next
   
   sAux = ""
   PosA = 0
   
   sVer = DsAuto.GetFileVersion(pFilename)
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
   
'   Pos = InStr(sVer, ".")
'   While Pos <> 0
'      sVer = Mid(sVer, 1, Pos - 1) & Mid(sVer, Pos + 1)
'      Pos = InStr(sVer, ".")
'   Wend
'   sVer = IIf(sVer = "", sVer = "0", sVer)
'   L_GetFileVersionNumber = Val(sVer)
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
Private Function L_ResolvePathName(ByVal sPath As String) As String
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
Private Function L_ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String, Optional DefaultValue As String) As String
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
   Dim b    As String
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
     b = Mid$(Password, i, 2)
     A3 = Val("&H" + b)
     A2 = A1 Xor A3
     S = S + Chr$(A2)
   Next
   L_Decrypt2 = Mid(S, 3)
End Function
Private Function L_Encrypt2(ByVal Password As String, Optional Key As String) As String
   Dim P    As String
   Dim b    As String
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
     b = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b
     S = S + b
   Next
   L_Encrypt2 = S
End Function
