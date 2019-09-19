VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PROPRIETARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim gLocalReg As String

Private mvarAppPAth As String
Private mvarAppExe  As String
Private mvarSetupFile As String
Public Property Let AppPath(ByVal vData As String)
   mvarAppPAth = vData
End Property
Public Property Get AppPath() As String
   AppPath = mvarAppPAth
End Property
Public Property Let AppExe(ByVal vData As String)
   mvarAppExe = vData
End Property
Public Property Get AppExe() As String
   AppExe = mvarAppExe
End Property
Public Property Let SetupFile(ByVal vData As String)
   mvarSetupFile = vData
End Property
Public Property Get SetupFile() As String
   SetupFile = mvarSetupFile
End Property
Public Sub Executa()
   gLocalReg = mvarAppPAth & "\" & mvarAppExe & ".reg"

   'Call CriarUsuarios
   Call Limpa_Setup_G3RREV
End Sub
Private Sub Limpa_Setup_G3RREV()
   Dim sFile As String
   sFile = Environ("PROGRAMFILES") & "\ClasseA\Admin\Dll\Setup.ini"
   If ExisteArquivo(sFile) Then
      If CDate(FileDateTime(sFile)) <= CDate("10/02/2011") Then
         Call ExcluirArquivo(sFile)
      End If
   End If
   sFile = Environ("PROGRAMFILES") & "\ClasseA\Admin\Dll\G3RREV.zia"
   If ExisteArquivo(sFile) Then
      If CDate(FileDateTime(sFile)) <= CDate("10/02/2011") Then
         Call ExcluirArquivo(sFile)
      End If
   End If
End Sub
Private Sub CriarBancoDeDados()
   Dim xConn   As Object
   Dim sConect As String
   Dim Sql     As String
    
   Dim sDbName As String
   Dim sServer As String
    
   sDbName = "G3R"
   sServer = Environ("computername") & "\SQLEXPRESS"
   
   sConect = "Provider=SQLOLEDB;"
   sConect = sConect & "Initial Catalog=" & sDbName & ";"
   sConect = sConect & "Data Source=" & sServer & ";"
   sConect = sConect & "Integrated Security=SSPI;"
    
   Set xConn = CreateObject("ADODB.Connection") ' New ADODB.Connection
   With xConn
      .CommandTimeout = 300
      .CursorLocation = adUseClient
      .ConnectionString = sConect
      .Open
      If .State = adStateOpen Then
         Call R02Users(xConn)
      End If
   End With
End Sub
Private Sub R02Users(xConn As ADODB.Connection)
   Dim Sql As String
   
   Sql = "USE [G3R];" & vbNewLine
   Sql = Sql & "-----------------------------------------------------------------------------------" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.server_principals WHERE name = N'USU_VERIF') DROP LOGIN [USU_VERIF];" & vbNewLine
   Sql = Sql & "CREATE LOGIN [USU_VERIF] WITH PASSWORD=N'MINOTAURO', DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[us_english], CHECK_EXPIRATION=OFF, CHECK_POLICY=ON;" & vbNewLine
   Sql = Sql & "--ALTER LOGIN [USU_VERIF] DISABLE;" & vbNewLine
   Sql = Sql & "-----------------------------------------------------------------------------------" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'USU_VERIF') DROP USER [USU_VERIF];" & vbNewLine
   Sql = Sql & "CREATE USER [USU_VERIF] FOR LOGIN [USU_VERIF] WITH DEFAULT_SCHEMA=[dbo];" & vbNewLine
   Sql = Sql & "GRANT CONNECT TO [USU_VERIF] AS [dbo];" & vbNewLine
   Sql = Sql & "-------------------------------------------------------------------------------------" & vbNewLine
   Sql = Sql & "" & vbNewLine
   Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_datareader', @membername = 'USU_VERIF';" & vbNewLine
   Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_datawriter', @membername = 'USU_VERIF';" & vbNewLine
   Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_backupoperator', @membername = 'USU_VERIF';" & vbNewLine
   Sql = Sql & "EXEC sys.sp_addsrvrolemember @loginame = N'USU_VERIF', @rolename = N'sysadmin';" & vbNewLine
   Sql = Sql & "-----------------------------------------------------------------------------------" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.server_principals WHERE name = N'DBA') DROP LOGIN [DBA];" & vbNewLine
   Sql = Sql & "CREATE LOGIN [DBA] WITH PASSWORD=N'BRASIL', DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[us_english], CHECK_EXPIRATION=OFF, CHECK_POLICY=ON;" & vbNewLine
   Sql = Sql & "--ALTER LOGIN [DBA] DISABLE;" & vbNewLine
   Sql = Sql & "-------------------------------------------------------------------------------------" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'DBA') DROP USER [DBA];" & vbNewLine
   Sql = Sql & "CREATE USER [DBA] WITHOUT LOGIN WITH DEFAULT_SCHEMA=[dbo];" & vbNewLine
   Sql = Sql & "GRANT CONNECT TO [DBA] AS [dbo];" & vbNewLine
   Sql = Sql & "-------------------------------------------------------------------------------------" & vbNewLine
   Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_owner', @membername = 'DBA';" & vbNewLine
   Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_accessadmin', @membername = 'DBA';" & vbNewLine
   
   xConn.Execute Sql
   
End Sub
