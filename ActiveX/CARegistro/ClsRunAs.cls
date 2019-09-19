VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RunAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarAut    As FrmAutenticacao
Attribute mvarAut.VB_VarHelpID = -1

Private mvarExePath     As String
Private mvarExeFile     As String
Private mvarLocalReg    As String
Private mvarCODSIS      As String
Private mvarXDb         As Object

Private mvarUserName    As String
Private mvarPassword    As String
Private mvarDomainName  As String
Private mvarCommandLine As String
Private mvarCurrentDir  As String
Private mvarCommand     As String

Public Property Let Command(ByVal vData As String)
   mvarCommand = vData
End Property
Public Property Get Command() As String
   Command = mvarCommand
End Property
Public Property Let CurrentDir(ByVal vData As String)
   mvarCurrentDir = vData
End Property
Public Property Get CurrentDir() As String
   CurrentDir = mvarCurrentDir
End Property
Public Property Let CommandLine(ByVal vData As String)
   mvarCommandLine = vData
End Property
Public Property Get CommandLine() As String
   CommandLine = mvarCommandLine
End Property
Public Property Let DomainName(ByVal vData As String)
   mvarDomainName = vData
End Property
Public Property Get DomainName() As String
   DomainName = mvarDomainName
End Property
Public Property Let Password(ByVal vData As String)
   mvarPassword = vData
End Property
Public Property Get Password() As String
   Password = mvarPassword
End Property
Public Property Let UserName(ByVal vData As String)
   mvarUserName = vData
End Property
Public Property Get UserName() As String
   UserName = mvarUserName
End Property
Public Property Let ExeFile(ByVal vData As String)
   mvarExeFile = vData
End Property
Public Property Get ExeFile() As String
   ExeFile = mvarExeFile
End Property
Public Property Let LocalReg(ByVal vData As String)
   mvarLocalReg = vData
End Property
Public Property Get LocalReg() As String
   mvarLocalReg = mvarExePath & mvarExeFile & ".reg"
   LocalReg = mvarLocalReg
End Property
Public Property Let CODSIS(ByVal vData As String)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let ExePath(ByVal vData As String)
   mvarExePath = vData
End Property
Public Property Get ExePath() As String
   ExePath = mvarExePath
End Property
Public Property Set XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Get XDb() As Object
   If mvarXDb Is Nothing Then
      Set mvarXDb = CreateObject("XBANCO01.DS_BANCO")
   End If
   Set XDb = mvarXDb
End Property
Private Sub Class_Initialize()
   Set mvarAut = New FrmAutenticacao
End Sub
Public Sub Show(Optional pModo As VBRUN.FormShowConstants = vbModal)
   mvarAut.Show pModo
End Sub
Private Sub mvarAut_Activate()
   If GetTag(mvarAut, "1Vez", "S") = "S" Then
      Call SetTag(mvarAut, "1Vez", "S")
      If mvarUserName <> "" Or mvarAut.TxtUserName.Text <> "" Then
         mvarAut_CmdAutenticacaoClick
      End If
      Screen.MousePointer = vbDefault
      mvarAut.MousePointer = vbDefault
   End If
End Sub
Private Sub mvarAut_CmdAutenticacaoClick()
   Dim i          As Integer
   Dim bOk        As Boolean
   Dim CtlObj     As Control
   Dim sCommand   As String
   Dim sTAG       As String
   Dim bUsuValido As Boolean
   
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   mvarAut.MousePointer = vbHourglass
   
   bUsuValido = True
   With mvarAut
      Call SetTag(.FraVerif, "CAPTION", .FraVerif.Caption)
      .FraVerif.Caption = "Verificando..."
      .PrbFlood.Visible = True
      .Refresh
      .FraVerif.Refresh
            
      mvarUserName = Trim(.TxtUserName.Text)
      mvarPassword = Trim(.TxtPassword.Text)
      mvarDomainName = Trim(.CmbDomainName.Text)
      
      If mvarUserName <> "" And mvarPassword <> "" And mvarDomainName <> "" Then
         Dim nCtrl As Control
         For Each nCtrl In .Controls
            If Mid(nCtrl.Name, 1, Len("ImgStatus")) = "ImgStatus" Then
               nCtrl.Visible = False
            End If
         Next
         'Call Wait(1)
         If mvarCommand <> "" Then
            If mvarUserName <> GetTag(mvarCommand, "USERNAME", "") Or _
               mvarPassword <> GetTag(mvarCommand, "PASSWORD", "") Or _
               mvarDomainName <> GetTag(mvarCommand, "DOMAINNAME", "") Then
               
               mvarCommand = ""
            End If
         End If
         
         If InStr(mvarCommand, "RUNAS") = 0 Then
            mvarCurrentDir = ResolvePathName(App.Path)
            If Environ("WINDIR") <> Mid(App.Path, 1, Len(Environ("WINDIR"))) Then
               mvarCurrentDir = ResolvePathName(ResolvePathName(Environ("WINDIR")) & "system32\ClasseA")
            End If
            mvarCommandLine = mvarCurrentDir & "CAReg.exe"
            
            sTAG = SetTag(sTAG, "USERNAME", mvarUserName)
            sTAG = SetTag(sTAG, "PASSWORD", mvarPassword)
            sTAG = SetTag(sTAG, "DOMAINNAME", mvarDomainName)
            sTAG = SetTag(sTAG, "COMMANDLINE", mvarCommandLine)
            sTAG = SetTag(sTAG, "CURRENTDIR", mvarCurrentDir)
         
            mvarCommandLine = mvarCommandLine & " /RUNAS:" & sTAG
            If 0 = RunAsUser(mvarUserName, mvarPassword, mvarDomainName, mvarCommandLine, mvarCurrentDir) Then
               Unload mvarAut
               Exit Sub
            Else
               bUsuValido = False
            End If
         End If
                     
         For i = .LblVerif.LBound To .LblVerif.uBound
            Set CtlObj = .Controls.Add("VB.IMAGE", "ImgStatus" & i)
            Set CtlObj.Picture = .ImgVerif.Picture
            Set CtlObj.Container = .FraVerif
            CtlObj.Stretch = True
            CtlObj.Width = 240
            CtlObj.Height = 240
            CtlObj.Left = .LblVerif(i).Left - 360
            CtlObj.Top = .LblVerif(i).Top
            CtlObj.Visible = True
            CtlObj.Refresh
            
            bOk = False
            
            If bUsuValido Then
               Select Case i
                  Case 0: bOk = Verif00_AcessoFTP
                  Case 1: bOk = Verif01_AcessoPastaArq
                  Case 2: bOk = Verif02_AcessoPastaSis
                  Case 3: bOk = Verif03_AcessoRegistro
                  Case 4: bOk = Verif04_AcessoServBD
                  Case 5: bOk = Verif05_AcessoBanco
                  Case 6: bOk = Verif06_AcessoServDados
                  Case 7: bOk = Verif07_AcessoIniReg
               End Select
            End If
                           
            'Call Wait(1)
            If bOk Then
               Set CtlObj.Picture = .ImgOK.Picture
            Else
               Set CtlObj.Picture = .ImgNOK.Picture
            End If
            CtlObj.Visible = True
            CtlObj.Refresh
            '.FraVerif.Refresh
            '.Refresh
            
         Next
      Else
         MsgBox "Usuário ou Senha inválido.", vbExclamation + vbOKOnly, "Autenticação"
      End If
   End With
   
   GoTo Saida
   Exit Sub
TrataErro:
   If Err = 727 Then 'There is already a control with the name 'IMG0'
      Set CtlObj = mvarAut.Controls("ImgStatus" & i)
      Resume Next
   End If
Saida:
   On Error Resume Next
   mvarAut.PrbFlood.Visible = False
   mvarAut.FraVerif.Caption = GetTag(mvarAut.FraVerif, "CAPTION", "")
   mvarAut.FraVerif.Refresh
   If mvarAut.TxtUserName.Visible Then
      mvarAut.TxtUserName.SetFocus
   End If
   Screen.MousePointer = vbDefault
   mvarAut.MousePointer = vbDefault
End Sub
Private Sub mvarAut_CmdSair()
   Unload mvarAut
End Sub
Private Sub mvarAut_Load()
   Dim ComputerSystem As Object
   
   On Error Resume Next
   With mvarAut
      .FraVerif.Caption = "Verificação"
      .TxtUserName.Text = ""
      .TxtPassword.Text = ""
      .CmbDomainName.Clear
      .CmbDomainName.ListIndex = -1
   
      Set ComputerSystem = CreateObject("WinNTSystemInfo")
      .CmbDomainName.AddItem ComputerSystem.DomainName
      If ComputerSystem.DomainName <> ComputerSystem.ComputerName Then
         .CmbDomainName.AddItem ComputerSystem.ComputerName
      End If
      If .CmbDomainName.ListCount > 0 Then
         .CmbDomainName.ListIndex = 0
      End If
       
      If mvarUserName <> "" Then .TxtUserName.Text = mvarUserName
      If mvarPassword <> "" Then .TxtPassword.Text = mvarPassword
      If mvarDomainName <> "" Then .CmbDomainName.Text = mvarDomainName
   End With
End Sub
Private Function Verif00_AcessoFTP() As Boolean
   Dim sPath      As String
   Dim sFile      As String
   Dim sFilePath  As String
   Dim sLinha     As String
   Dim NumArq     As Integer
   Dim bConect    As Boolean
   
   Dim MyFTP      As Object 'VersaoFTP.TL_VerifVersao
   On Error GoTo TrataErro
      
   sPath = ResolvePathName(Environ("TEMP"))
   sFile = "CARegistro.txt"
   sFilePath = sPath & sFile
   
   '*********
   '* Criar Arquivo
   If ExisteArquivo(sFilePath) Then
      Call ExcluirArquivo(sFilePath, False)
   End If
   sLinha = "Arquivo de teste de acesso gerado por CARegistro"
   NumArq = FreeFile()
   Open sFilePath For Output As #NumArq
   Write #NumArq, sLinha
   Close #NumArq
   
   '*********
   '* Testar conexao FTP
   Set MyFTP = CriarObjeto("VersaoFTP.TL_VerifVersao")
   'Set MyFTP = New VersaoFTP.TL_VerifVersao
   With MyFTP
      bConect = .ConectarFTP(.FtpIP, .FtpUsu, .FTPPwd)
      If bConect Then
         .EnviarArquivo sPath, sFile, .RemotePath, sFile
         If ExisteArquivo(sFilePath) Then Call ExcluirArquivo(sFilePath, False)
         .BaixarArquivo .RemotePath, sFile, sPath, sFile
         .AbrirPasta .RemotePath
         .ExcluirArquivo sFile, False
         .DesconectarFTP
     
         Verif00_AcessoFTP = ExisteArquivo(sFilePath)
     End If
   End With
   
   GoTo Saida
TrataErro:
   Verif00_AcessoFTP = False
   MsgBox Err & "-" & Error, vbCritical + vbOKOnly, "Autenticação"
Saida:
   On Error Resume Next
   If ExisteArquivo(sFilePath) Then
      Call ExcluirArquivo(sFilePath, False)
   End If
   If bConect Then
      Call MyFTP.DesconectarFTP
   End If
   Set MyFTP = Nothing
   DoEvents
End Function
Private Function Verif01_AcessoPastaArq() As Boolean
   Dim sFilePath  As String
   
   On Error GoTo TrataErro

   sFilePath = ResolvePathName(Environ("PROGRAMFILES")) & "CLASSEA\CARegistro.txt"
   
   Verif01_AcessoPastaArq = GerarArquivoTeste(sFilePath)
   
   GoTo Saida
TrataErro:
   MsgBox Err & "-" & Error, vbCritical + vbOKOnly, "Autenticação"
Saida:
   DoEvents
End Function
Private Function Verif02_AcessoPastaSis() As Boolean
   Dim sFilePath  As String
   
   On Error GoTo TrataErro

   sFilePath = ResolvePathName(Environ("WINDIR")) & "System32\CLASSEA\CARegistro.txt"
   
   Verif02_AcessoPastaSis = GerarArquivoTeste(sFilePath)
   GoTo Saida
TrataErro:
   MsgBox Err & "-" & Error, vbCritical + vbOKOnly, "Autenticação"
Saida:
   DoEvents
End Function
Private Function Verif03_AcessoRegistro() As Boolean
   Dim nIdioma As Long
   On Error GoTo TrataErro
   
   nIdioma = GetSetting("DSR", "General Format", "Idioma", 5000)
   Call SaveSetting("DSR", "General Format", "Idioma", nIdioma)
   
   Verif03_AcessoRegistro = True
   
   GoTo Saida
TrataErro:
   MsgBox Err & "-" & Error, vbCritical + vbOKOnly, "Autenticação"
Saida:
   DoEvents
End Function
Private Function Verif04_AcessoServBD() As Boolean
   Dim sFilePath  As String
   Dim sSetupIni  As String
   Dim MyLocalReg As String
   
   On Error GoTo TrataErro

   MyLocalReg = mvarLocalReg
   If MyLocalReg = "" Then
      MyLocalReg = DefineLocalReg("Setup", "PATHSETUP")
   End If
   
   sSetupIni = ReadIniFile(MyLocalReg, "Setup", "PATHSETUP")
   sSetupIni = ResolvePathName(sSetupIni) & "Setup.ini"
   
   If ExisteArquivo(sSetupIni) Then
      sFilePath = ReadIniFile(sSetupIni, "Database Format", "SERVER")
   Else
      sFilePath = ReadIniFile(MyLocalReg, "Conection 0", "SERVER")
   End If
   sFilePath = "\\" & ResolvePathName(sFilePath) & "c$\CARegistro.txt"
   
   Verif04_AcessoServBD = GerarArquivoTeste(sFilePath)
   
   GoTo Saida
TrataErro:
   MsgBox Err & "-" & Error, vbCritical + vbOKOnly, "Autenticação"
Saida:
   DoEvents
End Function
Private Function Verif05_AcessoBanco() As Boolean
   Dim MyLocalReg As String
   Dim sSetupIni  As String
   Dim sSERVER    As String
   Dim sDBNAME    As String
   Dim sUID       As String
   Dim sPwd       As String
   
   Dim bisODBC    As Boolean
   Dim nDBTIPO    As Integer
   Dim sdbVersao  As String
   Dim bisADO     As Boolean
   
   On Error GoTo TrataErro
   
   MyLocalReg = mvarLocalReg
   If MyLocalReg = "" Then
      MyLocalReg = DefineLocalReg("Setup", "PATHSETUP")
   End If

   sSetupIni = ReadIniFile(MyLocalReg, "Setup", "PATHSETUP")
   sSetupIni = ResolvePathName(sSetupIni) & "Setup.ini"
   
   If ExisteArquivo(sSetupIni) Then
      bisODBC = ReadIniFile(sSetupIni, "Database Format", "isODBC", False)
      nDBTIPO = ReadIniFile(sSetupIni, "Database Format", "DBTIPO", 1)
      sdbVersao = ReadIniFile(sSetupIni, "Database Format", "dbVersao", "7")
      bisADO = ReadIniFile(sSetupIni, "Database Format", "isADO", True)
      
      sSERVER = ReadIniFile(sSetupIni, "Database Format", "SERVER")
      sDBNAME = ReadIniFile(sSetupIni, "Database Format", "DBNAME")
      sUID = ReadIniFile(sSetupIni, "Database Format", "UID")
      sPwd = ReadIniFile(sSetupIni, "Database Format", "PWD")

   Else
      bisODBC = ReadIniFile(sSetupIni, "Conection 0", "isODBC", False)
      nDBTIPO = ReadIniFile(sSetupIni, "Conection 0", "DBTIPO", 1)
      sdbVersao = ReadIniFile(sSetupIni, "Conection 0", "dbVersao", "7")
      bisADO = ReadIniFile(sSetupIni, "Conection 0", "isADO", True)
      
      sSERVER = ReadIniFile(MyLocalReg, "Conection 0", "SERVER")
      sDBNAME = ReadIniFile(MyLocalReg, "Conection 0", "DBNAME")
      sUID = ReadIniFile(MyLocalReg, "Conection 0", "UID")
      sPwd = ReadIniFile(MyLocalReg, "Conection 0", "PWD")
   End If
   
   Set mvarXDb = CreateObject("XBANCO01.DS_BANCO")
   If sSERVER <> "" And sDBNAME <> "" And sUID <> "" And sPwd <> "" Then
      With mvarXDb
         .isODBC = bisODBC
         .DBTIPO = nDBTIPO
         .dbVersao = sdbVersao
         .isADO = bisADO
         .SERVER = sSERVER
         .DBNAME = sDBNAME
         .UID = sUID
         .Pwd = Decrypt2(sPwd)
         .SrvConecta
      End With
   End If
   Verif05_AcessoBanco = mvarXDb.Conectado
   
   GoTo Saida
TrataErro:
   MsgBox Err & "-" & Error, vbCritical + vbOKOnly, "Autenticação"
Saida:
   DoEvents
End Function

Private Function Verif06_AcessoServDados() As Boolean
   Dim sFilePath  As String
   Dim sSetupIni  As String
   Dim MyLocalReg As String
   
   On Error GoTo TrataErro

   MyLocalReg = mvarLocalReg
   If MyLocalReg = "" Then
      MyLocalReg = DefineLocalReg("Setup", "PATHSETUP")
   End If

   sSetupIni = ReadIniFile(MyLocalReg, "Setup", "PATHSETUP")
   sSetupIni = ResolvePathName(sSetupIni) & "Setup.ini"
   
   If ExisteArquivo(sSetupIni) Then
      sFilePath = ReadIniFile(sSetupIni, "AutoInstall Files", "Path")
   Else
      sFilePath = ReadIniFile(MyLocalReg, "Setup", "PATHSETUP")
   End If
   sFilePath = sFilePath & "CARegistro.txt"
   
   Verif06_AcessoServDados = GerarArquivoTeste(sFilePath)
   
   GoTo Saida
TrataErro:
   MsgBox Err & "-" & Error, vbCritical + vbOKOnly, "Autenticação"
Saida:
   DoEvents
End Function
Private Function Verif07_AcessoIniReg() As Boolean
   
   Dim sRegistro  As String
   Dim sPath      As String
   Dim CollPath   As Collection
   Dim CollSetup  As Collection
   Dim vAux       As Variant
   Dim sStatus    As String
   Dim LerIni     As Boolean
   
   Dim sAppPath   As String
   Dim sArq       As String
   
   Dim sSetupReg  As String
   Dim sSetupIni  As String
   
   On Error GoTo TrataErro
   
   sRegistro = ""
   sRegistro = SetTag(sRegistro, "USERNAME", mvarUserName)
   sRegistro = SetTag(sRegistro, "PASSWORD", mvarPassword)
   sRegistro = SetTag(sRegistro, "DOMAINNAME", mvarDomainName)
   sRegistro = Encrypt2(sRegistro)
   
   '*****************
   '* Localizar arquivo REG
   If sRegistro <> "" Then
      sAppPath = ResolvePathName(ResolvePathName(Environ("PROGRAMFILES")) & "CLASSEA")
      Set CollPath = New Collection
      sPath = Dir(sAppPath, vbDirectory)
      Do While sPath <> ""    ' Start the loop.
         If sPath <> "." And sPath <> ".." Then
            If (GetAttr(sAppPath & sPath) And vbDirectory) = vbDirectory Then
               CollPath.Add sPath
            End If
         End If
         sPath = Dir
      Loop
      
      Set CollSetup = New Collection
      For Each vAux In CollPath
         sArq = Dir(ResolvePathName(sAppPath & vAux) & "*.reg", vbArchive)
         sSetupReg = ResolvePathName(sAppPath & vAux) & sArq
         If Trim(sSetupReg) <> "" And ExisteArquivo(sSetupReg) Then
            If sSetupIni <> ReadIniFile(sSetupReg, "Setup", "PATHSETUP") Then
               sSetupIni = ReadIniFile(sSetupReg, "Setup", "PATHSETUP")
               CollSetup.Add sSetupIni
            End If
            Call WriteIniFile(sSetupReg, "General", "Registry", sRegistro)
            sArq = Dir
            While sArq <> ""
               sSetupReg = ResolvePathName(sAppPath & vAux) & sArq
               Call WriteIniFile(sSetupReg, "General", "Registry", sRegistro)
               sArq = Dir
            Wend
         End If
      Next
               
   '   '*************
   '   '* PathSetup = "\\guarani\sistemas\admin\"
      For Each vAux In CollSetup
         sSetupIni = vAux & "Setup.ini"
         If ExisteArquivo(sSetupIni) Then
            sStatus = ReadIniFile(sSetupIni, "AutoInstall", "Status", "0")
            LerIni = (sStatus = "1" Or sStatus = "2")
            If LerIni Then
               Call WriteIniFile(sSetupIni, "General", "Registry", sRegistro)
            End If
         End If
      Next
   End If
   If Verif06_AcessoServDados Then
      Verif07_AcessoIniReg = True
   End If
   GoTo Saida
TrataErro:
   MsgBox Err & "-" & Error, vbCritical + vbOKOnly, "Autenticação"
Saida:
   DoEvents
End Function
Public Function Execute(Optional pUserName As String, Optional pPassword As String, Optional pDomainName As String, Optional pCommandLine As String, Optional pCurrentDir As String, Optional pTAG As String) As Long
   
   mvarUserName = IIf(Trim(pUserName) = "", mvarUserName, pUserName)
   mvarPassword = IIf(Trim(pPassword) = "", mvarPassword, pPassword)
   mvarDomainName = IIf(Trim(pDomainName) = "", mvarDomainName, pDomainName)
   mvarCommandLine = IIf(Trim(pCommandLine) = "", mvarCommandLine, pCommandLine)
   mvarCurrentDir = IIf(Trim(pCurrentDir) = "", mvarCurrentDir, pCurrentDir)
   
   If Trim(pTAG) <> "" Then
      pUserName = Trim(GetTag(pTAG, "USERNAME", ""))
      pPassword = Trim(GetTag(pTAG, "PASSWORD", ""))
      pDomainName = Trim(GetTag(pTAG, "DOMAINNAME", ""))
      pCommandLine = Trim(GetTag(pTAG, "COMMANDLINE", ""))
      pCurrentDir = Trim(GetTag(pTAG, "CURRENTDIR", ""))
   
      mvarUserName = IIf(Trim(pUserName) = "", mvarUserName, pUserName)
      mvarPassword = IIf(Trim(pPassword) = "", mvarPassword, pPassword)
      mvarDomainName = IIf(Trim(pDomainName) = "", mvarDomainName, pDomainName)
      mvarCommandLine = IIf(Trim(pCommandLine) = "", mvarCommandLine, pCommandLine)
      mvarCurrentDir = IIf(Trim(pCurrentDir) = "", mvarCurrentDir, pCurrentDir)
   End If
   
   Execute = -1
   If mvarUserName <> "" And mvarPassword <> "" And mvarDomainName <> "" Then
      Execute = RunAsUser(mvarUserName, mvarPassword, mvarDomainName, mvarCommandLine, mvarCurrentDir)
   End If
End Function
Private Function DefineLocalReg(Optional pChave As String, Optional pVar As String) As String
   Dim sPath      As String
   Dim CollPath   As Collection
   Dim CollSetup  As Collection
   Dim vAux       As Variant
   Dim sAppPath   As String
   Dim sArq       As String
   Dim sSetupReg  As String
   Dim sAux       As String

   sAppPath = ResolvePathName(ResolvePathName(Environ("PROGRAMFILES")) & "CLASSEA")
   Set CollPath = New Collection
   sPath = Dir(sAppPath, vbDirectory)
   Do While sPath <> ""    ' Start the loop.
      If sPath <> "." And sPath <> ".." Then
         If (GetAttr(sAppPath & sPath) And vbDirectory) = vbDirectory Then
            CollPath.Add sPath
         End If
      End If
      sPath = Dir
   Loop

   Set CollSetup = New Collection
   For Each vAux In CollPath
      sArq = Dir(ResolvePathName(sAppPath & vAux) & "*.reg", vbArchive)
      sSetupReg = ResolvePathName(sAppPath & vAux) & sArq
      If Trim(sSetupReg) <> "" And ExisteArquivo(sSetupReg) Then
         If pChave = "" Then
            DefineLocalReg = sSetupReg
            Exit For
         Else
            sAux = ReadIniFile(sSetupReg, pChave, pVar)
            If sAux <> "" Then
               DefineLocalReg = sSetupReg
               Exit For
            End If
         End If
         Exit For
      End If
   Next
End Function
Private Function DefineSetupIni() As String
   Dim sPath      As String
   Dim CollPath   As Collection
   Dim CollSetup  As Collection
   Dim vAux       As Variant
   Dim sAppPath   As String
   Dim sArq       As String
   Dim sSetupReg  As String
   Dim sSetupIni  As String
   
   sAppPath = ResolvePathName(ResolvePathName(Environ("PROGRAMFILES")) & "CLASSEA")
   Set CollPath = New Collection
   sPath = Dir(sAppPath, vbDirectory)
   Do While sPath <> ""    ' Start the loop.
      If sPath <> "." And sPath <> ".." Then
         If (GetAttr(sAppPath & sPath) And vbDirectory) = vbDirectory Then
            CollPath.Add sPath
         End If
      End If
      sPath = Dir
   Loop
   
   Set CollSetup = New Collection
   For Each vAux In CollPath
      sArq = Dir(ResolvePathName(sAppPath & vAux) & "*.reg", vbArchive)
      sSetupReg = ResolvePathName(sAppPath & vAux) & sArq
      If Trim(sSetupReg) <> "" And ExisteArquivo(sSetupReg) Then
         sSetupIni = ReadIniFile(sSetupReg, "Setup", "PATHSETUP") & "Setup.ini"
         If ExisteArquivo(sSetupIni) Then
            DefineSetupIni = sSetupIni
            Exit For
         End If
      End If
   Next
End Function
Private Function GerarArquivoTeste(pArquivo As String) As Boolean
   Dim sLinha  As String
   Dim NumArq  As Integer
   
   On Error Resume Next
   
   If ExisteArquivo(pArquivo) Then
      Call ExcluirArquivo(pArquivo, False)
   End If

   sLinha = "Arquivo de teste de acesso gerado por CARegistro"
   NumArq = FreeFile()
   Open pArquivo For Output As #NumArq
    Write #NumArq, sLinha
   Close #NumArq
   
   If ExisteArquivo(pArquivo) Then
      Call ExcluirArquivo(pArquivo, False)
   End If
   
   GerarArquivoTeste = (Err = 0)
End Function
