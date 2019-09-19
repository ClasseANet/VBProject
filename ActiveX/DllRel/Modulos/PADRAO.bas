Attribute VB_Name = "PADRAO"
Option Explicit
Global Const gCODSIS = "SGO"
Global Const gCaption1 = "Sistema"
Global Const gCaption2 = "de Gestão"
Global Const gCaption3 = "Operacional"

Global MdiMenu As Object   '* Menu.ControlMenu

#If ComRef Then
   Global MyPM    As PMClasse
#Else
   Global MyPM    As Object
#End If
Global ClTelas As Collection
'============================
'= Classes                  =
'============================
Global SysGO    As New SETTING
'Global SysGO    As New SysA.SetA
Global BANCO    As New BANCO_SGO

Public Sub GetConfig()
   Dim IniFile As String
   Dim LerIni As Boolean


   '***************************
   '*** [ Database Format ] ***
   '***************************

'   With XDb
'      .isODBC = GetSetting(SysA.CODSIS, "Database Format", "isODBC", False)
'      .dbTipo = GetSetting(SysA.CODSIS, "Database Format", "DBTIPO", eDbTipo.SQL_SERVER)
'      .dbVersao = GetSetting(SysA.CODSIS, "Database Format", "DBVERSAO", "7.0")
'      .isADO = GetSetting(SysA.CODSIS, "Database Format", "isADO", True)
'      Select Case XDb.dbTipo
'         Case eDbTipo.Access
'            .dbDrive = GetSetting(SysA.CODSIS, "Database Drive", "DBDRIVE", "C:\DSR\" + UCase(SysA.CODSIS) + "\")
'            .dbName = GetSetting(SysA.CODSIS, "Database Format", "DBNAME", UCase(SysGO.AppExeName) & ".mdb")
'         Case eDbTipo.SQL_SERVER
'            .Server = GetSetting(SysA.CODSIS, "Database Format", "SERVER", "SERVIDOR_SQL")
'            .dbName = GetSetting(SysA.CODSIS, "Database Format", "DBNAME", UCase(SysA.CODSIS))
'         Case eDbTipo.ORACLE
'            .Server = GetSetting(SysA.CODSIS, "Database Format", "SERVER", "SERVIDOR_SQL")
'            .dbName = GetSetting(SysA.CODSIS, "Database Format", "DBNAME", UCase(SysA.CODSIS))
'      End Select
'      .DSN = GetSetting(SysA.CODSIS, "Database Format", "DSN", IIf(.isODBC, UCase(SysA.CODSIS), ""))
'      .UID = GetSetting(SysA.CODSIS, "Database Format", "UID", "USU_VERIF")
'      .PWD = GetSetting(SysA.CODSIS, "Database Format", "PWD", "DIPLOMATA")
'   End With
'   With XDb
'      IniFile = SysGO.AppExeName & ".ini"
'      IniFile = App.Path & "\" & IniFile
'      If FileExists(IniFile) Then
'         LerIni = (ReadIniFile(IniFile, "General", "Status") = "1")
'         If LerIni Then
'            If ReadIniFile(IniFile, "Database Format", "isODBC") <> "" Then
'               .isODBC = ReadIniFile(IniFile, "Database Format", "isODBC")
'            End If
'            If ReadIniFile(IniFile, "Database Format", "dbTipo") <> "" Then
'               .dbTipo = ReadIniFile(IniFile, "Database Format", "dbTipo")
'            End If
'            If ReadIniFile(IniFile, "Database Format", "dbVersao") <> "" Then
'               .dbVersao = ReadIniFile(IniFile, "Database Format", "dbVersao")
'            End If
'            If ReadIniFile(IniFile, "Database Format", "isADO") <> "" Then
'               .isADO = ReadIniFile(IniFile, "Database Format", "isADO")
'            End If
'            If ReadIniFile(IniFile, "Database Drive", "dbDrive") <> "" Then
'               .dbDrive = ReadIniFile(IniFile, "Database Drive", "dbDrive")
'            End If
'            If ReadIniFile(IniFile, "Database Format", "Server") <> "" Then
'               .Server = ReadIniFile(IniFile, "Database Format", "Server")
'            End If
'            If ReadIniFile(IniFile, "Database Format", "dbName") <> "" Then
'               .dbName = ReadIniFile(IniFile, "Database Format", "dbName")
'            End If
'            If ReadIniFile(IniFile, "Database Format", "DSN") <> "" Then
'               .DSN = ReadIniFile(IniFile, "Database Format", "DSN")
'            End If
'            If ReadIniFile(IniFile, "Database Format", "UID") <> "" Then
'               .UID = ReadIniFile(IniFile, "Database Format", "UID")
'            End If
'            If ReadIniFile(IniFile, "Database Format", "Pwd") <> "" Then
'               .PWD = ReadIniFile(IniFile, "Database Format", "Pwd")
'            End If
'         End If
'      End If
'   End With
'
'   '**************************
'   '***     [ Setup ]      ***
'   '**************************
'   SysGO.DSVM = GetSetting(SysA.CODSIS, "Setup", "DSVM", False)
'   SysGO.Idioma = GetSetting(SysA.CODSIS, "Setup", "IDIOMA", eIdioma.PORTUGUES)
'
'   SysGO.FundoTela = GetSetting(SysA.CODSIS, "Setup", "FUNDOTELA", "FUNDO")
'   SysGO.DrvTmpErro = GetSetting(SysA.CODSIS, "Setup", "DRVERRO", "C:\Tmp\" & UCase(SysA.CODSIS) & "\Erro\")
'
'   If Trim(SysGO.DrvTmpErro) = "" Then SysGO.DrvTmpErro = "C:\Tmp\" & UCase(SysA.CODSIS) & "\Erro\"
'   If Dir(SysGO.DrvTmpErro, vbDirectory) = "" Then Call MakePath(SysGO.DrvTmpErro)
'
'   SysGO.dbDrive_Orig = XDb.dbDrive
'   SysGO.ExibeToolBar = GetSetting(SysA.CODSIS, "Setup", "EXIBETOOLBAR", True)
'   SysGO.ExibeLstCadastro = GetSetting(SysA.CODSIS, "Setup", "EXIBELSTCADASTRO", False)
'   SysGO.SaiComESC = GetSetting(SysA.CODSIS, "Setup", "SAICOMESC", False)
'   SysGO.Resize = GetSetting(SysA.CODSIS, "Setup", "REDIMENSIONA", False)

End Sub
'Public Sub SaveConfig()
'
'   '************* [ General ] *******************
'   '** Configurações necessárias a serem lidas **
'   '** pelas bibliotecas DSR100 e xBanco01.    **
'   '*********************************************
'   Call SaveSetting("DSR", "General Format", "Idioma", SysGO.Idioma)
'   Call SaveSetting("DSR", "General Format", "dbTipo", XDb.dbTipo)
'
'   '*** [ Database Format ] ***
'   Call SaveSetting(SysA.CODSIS, "Database Format", "isODBC", XDb.isODBC)
'   Call SaveSetting(SysA.CODSIS, "Database Format", "DBTIPO", XDb.dbTipo)
'   Call SaveSetting(SysA.CODSIS, "Database Format", "DBVERSAO", XDb.dbVersao)
'   Call SaveSetting(SysA.CODSIS, "Database Format", "isADO", XDb.isADO)
'
'   Call SaveSetting(SysA.CODSIS, "Database Format", "SERVER", XDb.Server)
'   Call SaveSetting(SysA.CODSIS, "Database Format", "DBNAME", XDb.dbName)
'   Call SaveSetting(SysA.CODSIS, "Database Format", "DSN", XDb.DSN)
'   Call SaveSetting(SysA.CODSIS, "Database Format", "UID", XDb.UID)
'   Call SaveSetting(SysA.CODSIS, "Database Format", "PWD", XDb.PWD)
'
'   '*** [ Database Drive ] ***
'   Call SaveSetting(SysA.CODSIS, "Database Drive", "DBDRIVE", XDb.dbDrive)
'
'   '*** [ Setup ] ***
'   Call SaveSetting(SysA.CODSIS, "Setup", "DSVM", SysGO.DSVM)
'   Call SaveSetting(SysA.CODSIS, "Setup", "FUNDOTELA", SysGO.FundoTela)
'   Call SaveSetting(SysA.CODSIS, "Setup", "IDIOMA", SysGO.Idioma)
'   Call SaveSetting(SysA.CODSIS, "Setup", "DRVERRO", SysGO.DrvTmpErro)
'   Call SaveSetting(SysA.CODSIS, "Setup", "EXIBETOOLBAR", SysGO.ExibeToolBar)
'   Call SaveSetting(SysA.CODSIS, "Setup", "EXIBELSTCADASTRO", SysGO.ExibeLstCadastro)
'   Call SaveSetting(SysA.CODSIS, "Setup", "SAICOMESC", SysGO.SaiComESC)
'   Call SaveSetting(SysA.CODSIS, "Setup", "REDIMENSIONA", SysGO.Resize)
'End Sub
Public Sub GetConfigSis()
   Dim IniFile     As String
   Dim sRptEstacao As String
   Dim sRptBanco   As String
   Dim sRptINI     As String
      
   IniFile = App.EXEName & ".ini"
   IniFile = App.Path & "\" & IniFile
   If FileExists(IniFile) Then
      sRptINI = ReadIniFile(IniFile, "General", "RptDrive")
   End If
   
   '**************************
   '*** [ Database Drive ] ***
   '**************************
   If sRptINI = "" Then
      sRptBanco = GetParam("DRVRPT", Default:=UCase(App.Path) + "\RPT\")
      sRptEstacao = GetSetting(SysA.CODSIS, "Report Drive", "DRVRPT", "")
      If sRptEstacao <> SysGO.DrvRpt And Trim(sRptEstacao) <> "" Then
         SysGO.DrvRpt = sRptEstacao
      Else
         SysGO.DrvRpt = sRptBanco
      End If
   Else
      SysGO.DrvRpt = sRptINI
   End If
   
   If Dir(SysGO.DrvRpt, vbDirectory) = "" Then
      Call MakePath(SysGO.DrvRpt)
   End If
   SysGO.DrvRpt = SysGO.DrvRpt & IIf(Right(SysGO.DrvRpt, 1) = "\", "", "\")
   
   SysGO.DSVM = GetSetting(SysA.CODSIS, "Setup", "DSVM", False)
      
   SysGO.NomeEmpresa = GetParam("EMPRESA", SysA.CODSIS)
End Sub
Public Sub SaveConfigSis()
   Call SaveParam("DRVRPT", SysGO.DrvRpt)
   Call SaveSetting(SysA.CODSIS, "Setup", "DSVM", SysGO.DSVM)
End Sub
Public Sub SaveParam(pCODPARAM As String, pVLPARAM As String, Optional pDSCPARAM, Optional pCODSIS)
  Dim bExiste As Boolean
  
  If IsMissing(pCODSIS) Then pCODSIS = SysA.CODSIS
    
  With BANCO.TB_PARAM
     bExiste = .Pesquisar(pCODSIS, pCODPARAM)
     .CODSIS = pCODSIS
     .CODPARAM = pCODPARAM
     .VLPARAM = pVLPARAM
     If Not IsMissing(pDSCPARAM) Then .DSCPARAM = pDSCPARAM

     If bExiste Then
        .Alterar True
     Else
        .Incluir True
     End If
  End With
End Sub
Public Function GetParam(CODPARAM As String, Optional CODSIS, Optional Default)
   Dim Sql As String
   Dim MyRs As Object
   Dim MyPARAM As TB_PARAM
   
   If IsMissing(CODSIS) Then CODSIS = SysA.CODSIS
   If IsMissing(Default) Then Default = ""
   
   CODPARAM = UCase(Trim(Mid(CODPARAM, 1, 10)))
   CODSIS = UCase(CStr(CODSIS))
      
   Set MyPARAM = New TB_PARAM
   With MyPARAM
      'Set .XDb = XDb
      If .Pesquisar(CODSIS, CODPARAM) Then
         GetParam = .VLPARAM
      Else
         GetParam = Default
         .CODPARAM = UCase(Trim(Mid(CODPARAM, 1, 10)))
         .CODSIS = Mid(InputBox("Inclua o Parâmetro : 'Código do Sistema'", SysA.CODSIS, SysA.CODSIS), 1, 30)  'Mid(UCase(CStr(CODSIS)), 1, 10)
         If Trim(.CODSIS) = "" Then
            .CODSIS = SysA.CODSIS
         End If
         .DSCPARAM = Mid(InputBox("Inclua a Descrição do Parâmetro " & UCase(Trim(Mid(CODPARAM, 1, 10))), SysA.CODSIS), 1, 30)
         .VLPARAM = Default
         .Incluir True
      End If
   End With
   Set MyPARAM = Nothing
End Function
'******************************************************
'******************************************************
'***** Pertence ao novo Padrão.bas ********************
'******************************************************
'******************************************************
Public Function LoadVersion() As String
   Dim sVer    As String
   
   LoadVersion = App.Major & "." & App.Minor & "." & App.Revision
   
End Function
