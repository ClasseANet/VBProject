Attribute VB_Name = "Padrao"
Option Explicit
Global XDb  As DS_BANCO
Global Sys  As New System

Public Splash As Object 'FrmInicio
Public Sub GetConfig()
   Dim IniFile    As String
   Dim LerIni     As Boolean
   Dim sAppName   As String
   
   sAppName = UCase(App.EXEName)

   '***************************
   '*** [ Database Format ] ***
   '***************************
   With XDb
      .isODBC = GetSetting(sAppName, "Database Format", "isODBC", False)
      .dbTipo = GetSetting(sAppName, "Database Format", "DBTIPO", eDbTipo.SQL_SERVER)
      .dbVersao = GetSetting(sAppName, "Database Format", "DBVERSAO", "7.0")
      .isADO = GetSetting(sAppName, "Database Format", "isADO", True)
      Select Case XDb.dbTipo
         Case eDbTipo.Access
            .dbDrive = GetSetting(sAppName, "Database Format", "DBDRIVE", "C:\DSR\" + UCase(sAppName) + "\")
            .dbName = GetSetting(sAppName, "Database Format", "DBNAME", UCase(Sys.AppExeName) & ".mdb")
         Case eDbTipo.SQL_SERVER
            .Server = GetSetting(sAppName, "Database Format", "SERVER", "SERVIDOR_SQL")
            .dbName = GetSetting(sAppName, "Database Format", "DBNAME", UCase(sAppName))
         Case eDbTipo.ORACLE
            .Server = GetSetting(sAppName, "Database Format", "SERVER", "SERVIDOR_SQL")
            .dbName = GetSetting(sAppName, "Database Format", "DBNAME", UCase(sAppName))
      End Select
      .DSN = GetSetting(sAppName, "Database Format", "DSN", IIf(.isODBC, UCase(sAppName), ""))
      .UID = GetSetting(sAppName, "Database Format", "UID", "USU_VERIF")
      .PWD = GetSetting(sAppName, "Database Format", "PWD", "DIPLOMATA")
   End With
   With XDb
      IniFile = App.Path & "\" & App.EXEName & ".ini"
      
      LerIni = False
      If FileExists(IniFile) Then
         LerIni = (ReadIniFile(IniFile, "General", "Status") = "1")
         If LerIni Then
            If ReadIniFile(IniFile, "Database Format", "isODBC") <> "" Then
               .isODBC = ReadIniFile(IniFile, "Database Format", "isODBC")
            End If
            If ReadIniFile(IniFile, "Database Format", "dbTipo") <> "" Then
               .dbTipo = ReadIniFile(IniFile, "Database Format", "dbTipo")
            End If
            If ReadIniFile(IniFile, "Database Format", "dbVersao") <> "" Then
               .dbVersao = ReadIniFile(IniFile, "Database Format", "dbVersao")
            End If
            If ReadIniFile(IniFile, "Database Format", "isADO") <> "" Then
               .isADO = ReadIniFile(IniFile, "Database Format", "isADO")
            End If
            If ReadIniFile(IniFile, "Database Drive", "dbDrive") <> "" Then
               .dbDrive = ReadIniFile(IniFile, "Database Format", "dbDrive")
            End If
            If ReadIniFile(IniFile, "Database Format", "Server") <> "" Then
               .Server = ReadIniFile(IniFile, "Database Format", "Server")
            End If
            If ReadIniFile(IniFile, "Database Format", "dbName") <> "" Then
               .dbName = ReadIniFile(IniFile, "Database Format", "dbName")
            End If
            If ReadIniFile(IniFile, "Database Format", "DSN") <> "" Then
               .DSN = ReadIniFile(IniFile, "Database Format", "DSN")
            End If
            If ReadIniFile(IniFile, "Database Format", "UID") <> "" Then
               .UID = ReadIniFile(IniFile, "Database Format", "UID")
            End If
            If ReadIniFile(IniFile, "Database Format", "Pwd") <> "" Then
               .PWD = ReadIniFile(IniFile, "Database Format", "Pwd")
            End If
         End If
      End If
   End With
   
   '**************************
   '*** [ Database Drive ] ***
   '**************************
   'Sys.DrvRptDB = GetSetting(sAppName, "Report", "DRVRPTDB", UCase(App.Path) + "\RPT\")
   Sys.DrvRpt = GetSetting(sAppName, "Report", "DRVRPT", UCase(App.Path) + "\")
   
   '**************************
   '***     [ Setup ]      ***
   '**************************
   Sys.DSVM = GetSetting(sAppName, "Setup", "DSVM", False)

   If LerIni Then
      If ReadIniFile(IniFile, "Setup", "DSVM") <> "" Then
         Sys.DSVM = ReadIniFile(IniFile, "Setup", "DSVM")
      End If
   End If
   
   Sys.Idioma = GetSetting(sAppName, "Setup", "IDIOMA", eIdioma.PORTUGUES)
      
'   dsr100.Idioma = Sys.Idioma
'   XBANCO01.dbTipo = XDb.dbTipo
'   BANCO.Idioma = Sys.Idioma
   
'   Sys.FundoTela = GetSetting(sAppName, "Setup", "FUNDOTELA", "FUNDO")
'   Sys.DrvTmpErro = GetSetting(sAppName, "Setup", "DRVERRO", "C:\Tmp\" & UCase(sAppName) & "\Erro\")
   
'   If Trim(Sys.DrvTmpErro) = "" Then Sys.DrvTmpErro = "C:\Tmp\" & UCase(sAppName) & "\Erro\"
'   If Dir(Sys.DrvTmpErro, vbDirectory) = "" Then Call MakePath(Sys.DrvTmpErro)

   
   '   sys.DrvDrive = MDI.Drv1.List(0) + "\"
'   Sys.dbDrive_Orig = XDb.dbDrive
'   Sys.ExibeToolBar = GetSetting(sAppName, "Setup", "EXIBETOOLBAR", True)
'   Sys.ExibeListaCad = GetSetting(sAppName, "Setup", "EXIBELSTCADASTRO", False)
'   Sys.SaiComESC = GetSetting(sAppName, "Setup", "SAICOMESC", False)
'   Sys.Resize = GetSetting(sAppName, "Setup", "REDIMENSIONA", False)
End Sub
Public Sub SaveConfig()
   Dim sAppName   As String
   
   sAppName = UCase(Sys.AppExeName)
   
   '************* [ General ] *******************
   '** Configurações necessárias a serem lidas **
   '** pelas bibliotecas DSR100 e xBanco01.    **
   '*********************************************
   Call SaveSetting("DSR", "General Format", "Idioma", Sys.Idioma)
   Call SaveSetting("DSR", "General Format", "dbTipo", XDb.dbTipo)
   
   '*** [ Database Format ] ***
   Call SaveSetting(sAppName, "Database Format", "isODBC", XDb.isODBC)
   Call SaveSetting(sAppName, "Database Format", "DBTIPO", XDb.dbTipo)
   Call SaveSetting(sAppName, "Database Format", "DBVERSAO", XDb.dbVersao)
   Call SaveSetting(sAppName, "Database Format", "isADO", XDb.isADO)
   
   Call SaveSetting(sAppName, "Database Format", "SERVER", XDb.Server)
   Call SaveSetting(sAppName, "Database Format", "DBNAME", XDb.dbName)
   Call SaveSetting(sAppName, "Database Format", "DBDRIVE", XDb.dbDrive)
   Call SaveSetting(sAppName, "Database Format", "DSN", XDb.DSN)
   Call SaveSetting(sAppName, "Database Format", "UID", XDb.UID)
   Call SaveSetting(sAppName, "Database Format", "PWD", XDb.PWD)
   
  
   '*** [ Setup ] ***
   Call SaveSetting(sAppName, "Setup", "DSVM", Sys.DSVM)
   Call SaveSetting(sAppName, "Setup", "IDIOMA", Sys.Idioma)
End Sub
Public Sub SplashFlood(pValue As Integer, Optional pMsg As String = "", Optional pVisible As Boolean = True)
   Dim n As Form
'   If Splash Is Nothing Then
'      For Each n In Forms
'         If UCase("FrmInicio") = UCase(n.Name) Then
'            Set Splash = n
'            Exit For
'         End If
'      Next
'   End If
   If Not Splash Is Nothing Then
      Splash.Flood.Visible = pVisible
      If pVisible Then
         Splash.LblMsg.Caption = pMsg
         Splash.Flood.Value = pValue
      Else
         pMsg = ""
         Splash.Flood.Value = 0
      End If
      Splash.LblMsg.Visible = True
      Splash.LblMsg.Caption = pMsg
      Splash.LblMsg.Refresh
   End If
End Sub
Public Sub UnloadIni()
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

