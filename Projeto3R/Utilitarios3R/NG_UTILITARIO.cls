VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_UTILITARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eTipoSMS
    PosicaoDia = 1
    LembrarSessao
    LembrarCreme
End Enum
Private mvarPane        As Object

Private mvarSys         As Object
Private mvarIDCOLIGADA  As Integer
Private mvarIDLOJA      As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   Set gSys = vData
   
   mvarIDCOLIGADA = xVal(mvarSys.Propriedades("IDCOLIGADA"))
   mvarIDLOJA = xVal(mvarSys.Propriedades("IDLOJA"))
'   mvarIDAGENDA = mvarSys.Propriedades("IDAGENDA")
End Property
'Public Sub F_FecharDia()
   'Call F_Backup
   'Call EnviarSMS(eTipoSMS.PosicaoDia)
'    Screen.MousePointer = vbDefault
'End Sub
Public Sub F_SINCDB(Optional pExeExterno As Boolean)
   Dim oFtp       As Object
   Dim MyRs       As Object
   Dim RemPath    As String
   Dim LocPath    As String
   Dim LocAux     As String
   Dim sFile      As String
   Dim sFileBak   As String
   Dim sUNIDADE   As String
   Dim bOk        As Boolean
   Dim bNOk       As Boolean
   Dim bAux       As Boolean
   Dim Sql        As String
   Dim bIsWeb     As Boolean
   'Dim sUID       As String
   Dim sConn      As String
   Dim bMultiDb   As Boolean
   Dim bBaixou    As Boolean
   Dim bConectado As Boolean
   Dim iConection As Integer
   Dim sMsg       As String
   Dim nIDCOLIGADA As Integer
   Dim nIDLOJA    As Integer
   'dim sNMLOJA as string
   Dim bMesmoArq  As Boolean
   Dim sTag As String
   Dim TlCal      As Object
   
   Screen.MousePointer = vbHourglass
   If pExeExterno Then
      nIDCOLIGADA = mvarSys.IDCOLIGADA
      nIDLOJA = mvarSys.IDLOJA
   Else
      nIDCOLIGADA = mvarSys.IDCOLIGADA
      nIDLOJA = mvarIDLOJA
      iConection = ReadIniFile(mvarSys.LocalReg, "conections", "Last", 0)
      If iConection <> 0 Then
         Set mvarSys = mvarSys.MDI.ReloadConnection(0, False)
      End If
   End If
   Sql = "Select IDLOJA, IDCOLIGADA, NOME"
   Sql = Sql & " From OLOJA"
   Sql = Sql & " Where IDCOLIGADA=" & nIDCOLIGADA
   Sql = Sql & " And IDLOJA>1"
   Sql = Sql & " And ATIVO=1"
   bMultiDb = mvarSys.xDb.Abretabela(Sql, MyRs)
   
   On Error Resume Next
   If Not pExeExterno Then
      With mvarSys.MDI
         .ProgressBar2.Visible = True
         .CommandBars.StatusBar.Pane(3).Handle = .ProgressBar2.hwnd
      End With
   End If
   '* Conectar FTP
   If bMultiDb Then
      LocPath = mvarSys.PathTmp
      While Not MyRs.EOF
         mvarSys.IDLOJA = xVal(MyRs("IDLOJA") & "")
         sUNIDADE = Replace(MyRs("NOME") & "", " ", "")
         sFile = MyRs("IDCOLIGADA") & sUNIDADE & ".zip"
         sFileBak = MyRs("IDCOLIGADA") & sUNIDADE & ".bak"
         
         '* Limpar Pasta
         'Call ExcluirArquivo(LocPath & sFile)
         Call ExcluirArquivo(LocPath & sFileBak)
         
         RemPath = mvarSys.GetParam("FtpBakPath", Default:="", pDescricao:="Caminho do Ftp de Backup")
         RemPath = "/" & sUNIDADE & IIf(Mid(RemPath, 1, 1) = "/", "", "/") & RemPath
      
         '* Baixar Back-up
         LocAux = ResolvePathName(mvarSys.Propriedades("PATHBACKUP", True))
         If LocAux = "" Then
            LocAux = LocPath & "Backup\"
            LocAux = ProcuraArquivo(LocPath & "Backup\", sFile)
            LocAux = IIf(LocAux = "", LocPath & "Backup\", LocAux)
         End If
         
         bBaixou = False
         bMesmoArq = False
         If ExisteArquivo(LocAux & sFile) Then
            If DateDiff("h", FileDateTime(LocAux & sFile), mvarSys.xDb.SysDate(3)) <= 6 Then
               bBaixou = True
            Else
               If Weekday(mvarSys.xDb.SysDate) = 2 Then
                  bBaixou = DateDiff("d", FileDateTime(LocAux & sFile), mvarSys.xDb.SysDate) <= 2
               End If
            End If
            If bBaixou Then
               If ExisteArquivo(LocPath & sFile) Then
                  bMesmoArq = (FileDateTime(LocAux & sFile) = FileDateTime(LocPath & sFile))
                  If bMesmoArq Then
                     sTag = SetTag(sTag, "MESMOARQ" & mvarSys.IDLOJA, "1")
                  End If
               End If
            Else
               If pExeExterno Then
                  bBaixou = ExisteArquivo(LocPath & sFile)
                  If bBaixou Then
                     bMesmoArq = (FileDateTime(LocAux & sFile) = FileDateTime(LocPath & sFile))
                     If bMesmoArq Then
                        sTag = SetTag(sTag, "MESMOARQ" & mvarSys.IDLOJA, "1")
                     End If
                  End If
               End If
            End If
         Else
           '* Limpar Pasta
           Call ExcluirArquivo(LocPath & sFile)
         End If
         If bBaixou Then
            If Not bMesmoArq Then Call CopiarArquivo(LocAux & sFile, LocPath & sFile)
         Else
            If Not bConectado Then
               bIsWeb = IsWebConnected
               If Err <> 0 Then bIsWeb = True
               If bIsWeb Then
                  Set oFtp = ConectarFTP()
                  bConectado = False
                  If Not oFtp Is Nothing Then
                     bConectado = oFtp.Conectado
                  End If
                  If Not bConectado Then
                     sMsg = sMsg & vbNewLine & "O Sistema não conectou o servidor!"
                  End If
               Else
                  sMsg = sMsg & vbNewLine & "O Sistema não identificou a conexão de Internet!"
               End If
            End If
            If bConectado Then
               bBaixou = oFtp.BaixarArquivo(RemPath, sFile, LocPath, sFile)
               If bBaixou And Not ExisteArquivo(LocAux & sFile) Then
                  Call CopiarArquivo(LocPath & sFile, LocAux & sFile)
               End If
            End If
         End If
         If Not bBaixou And Not pExeExterno Then
            sMsg = sMsg & vbNewLine & "[" & sUNIDADE & "] O sistema não conseguiu conectar Servidor!"
         End If
         MyRs.MoveNext
      Wend
      
      MyRs.MoveFirst
      While Not MyRs.EOF
         mvarSys.IDLOJA = xVal(MyRs("IDLOJA") & "")
         sUNIDADE = Replace(MyRs("NOME") & "", " ", "")
         sFile = MyRs("IDCOLIGADA") & sUNIDADE & ".zip"
         sFileBak = MyRs("IDCOLIGADA") & sUNIDADE & ".bak"
         
         LocAux = LocPath
         bBaixou = ExisteArquivo(LocAux & sFile)
         If Not bBaixou Then
            LocAux = ResolvePathName(mvarSys.Propriedades("PATHBACKUP", True))
            If LocAux = "" Then
               LocAux = LocPath & "Backup\"
               LocAux = ProcuraArquivo(LocPath & "Backup\", sFile)
               LocAux = IIf(LocAux = "", LocPath & "Backup\", LocAux)
            End If
            bBaixou = ExisteArquivo(LocAux & sFile)
            If bBaixou Then
               Call CopiarArquivo(LocAux & sFile, LocPath & sFile)
            End If
         End If
            
         bNOk = False
         If pExeExterno Then
            If "1" = GetTag(sTag, "MESMOARQ" & mvarSys.IDLOJA, "0") Then
               bBaixou = False
            End If
         End If
         If bBaixou Then
            '* Descompactar e Restaurar Back-ups
            sMsg = sMsg & vbNewLine & "[" & sUNIDADE & "] " & Format(FileDateTime(LocAux & sFile), "dd/mm/yy hh:mm") & " Operação realizada com sucesso!"
            Call Unzip(LocPath, sFile, LocPath, False)
            If Not pExeExterno Then Call ExcluirArquivo(LocPath & sFile, False)
         ElseIf Not pExeExterno Then
            bNOk = True
            sMsg = sMsg & vbNewLine & "[" & sUNIDADE & "] O Sistema não conseguiu baixar o back-up."
            If ExisteArquivo(LocPath & sFile) Then
               If vbYes = ExibirPergunta("O Sistema não baixou o back-up. Deseja continuar com arquivo em disco?") Then
                  bBaixou = True
                  bNOk = False
               End If
            End If
         End If
         If bBaixou Then
            bAux = F_RESTOREDATABASE(mvarSys.xDb, LocPath & sFileBak, "G3R_" & sUNIDADE, sMsg, sUNIDADE)
            bOk = bOk Or bAux
            bNOk = bNOk Or Not bAux
            Call ExcluirArquivo(Replace(LocPath & sFile, ".zip", ".bak"), False)
         End If
                        
         '* Limpar Pasta
         'Call ExcluirArquivo(LocPath & sFile)
         'Call ExcluirArquivo(LocPath & sFileBak)
         If bBaixou Then
            sConn = "Conection " & MyRs.AbsolutePosition
            If Trim(mvarSys.LocalReg) <> "" Then
               Call WriteIniFile(mvarSys.LocalReg, sConn, "ALIAS", UCase(sUNIDADE))
               Call WriteIniFile(mvarSys.LocalReg, sConn, "DBNAME", "G3R_" & UCase(sUNIDADE))
               Call WriteIniFile(mvarSys.LocalReg, sConn, "isODBC", mvarSys.xDb.isODBC)
               Call WriteIniFile(mvarSys.LocalReg, sConn, "DBTIPO", mvarSys.xDb.DBTIPO)
               Call WriteIniFile(mvarSys.LocalReg, sConn, "isADO", mvarSys.xDb.isADO)
               If InStr(mvarSys.xDb.Server, ComputerName & "\") <> 0 Then
                  Call WriteIniFile(mvarSys.LocalReg, sConn, "SERVER", Replace(mvarSys.xDb.Server, ComputerName, "[Local]"))
               Else
                  Call WriteIniFile(mvarSys.LocalReg, sConn, "SERVER", mvarSys.xDb.Server)
               End If
               Call WriteIniFile(mvarSys.LocalReg, sConn, "UID", mvarSys.xDb.UID)
               Call WriteIniFile(mvarSys.LocalReg, sConn, "PWD", Encrypt2(mvarSys.xDb.PWD))
            End If
         End If
         MyRs.MoveNext
      Wend
   Else
      ExibirInformacao "Não há base a ser sincronizada!"
   End If
      
   mvarSys.IDLOJA = nIDLOJA
   If Not pExeExterno Then
      With mvarSys.MDI
         .ProgressBar2.Visible = False
         .CommandBars.StatusBar.Pane(3).Handle = .ProgressBar.hwnd
      End With
   
      If bMultiDb Then
         If iConection <> 0 Then Call mvarSys.MDI.ReloadConnection(iConection)
      End If
   End If
      
   If bMultiDb Then
      If Not pExeExterno Then
         If bOk Then
            DoEvents
            SendKeys "{F5}"
            'Set TlCal = CriarObjeto("Calendario3R.TL_Calendario")
            'Set TlCal.Sys = mvarSys
            'Call TlCal.RefreshCalendar
            ExibirInformacao sMsg & IIf(bNOk, vbNewLine & vbNewLine & "Tente mais tarde.", "")
         Else
            ExibirInformacao "Houve uma falha na operação!" & vbNewLine & sMsg & IIf(oFtp.Conectado, "", vbNewLine & "Conexão ao servidor não estabelecida!") & vbNewLine & vbNewLine & "Tente mais tarde."
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Function F_RESTOREDATABASE(ByRef pxDb As Object, pFileBak As String, pDataBaseName As String, ByRef sMsg As String, Optional pAlias As String = "Produção") As Boolean
   Dim Sql     As String
   Dim bOk     As Boolean
   Dim sUID    As String
   Dim DataPath   As String
   Dim MyRs       As Object
      
   With pxDb
      Set MyRs = Nothing
      Sql = "Select * From SysFiles"
      If .Abretabela(Sql, MyRs) Then
         DataPath = MyRs("FileName")
         DataPath = Mid(DataPath, 1, Len(DataPath) - InStr(StrReverse(DataPath), "\"))
      Else
         DataPath = GetSpecialFolder(38) & "Microsoft SQL Server\MSSQL.1\MSSQL\DATA\"
      End If
      
      DataPath = ResolvePathName(DataPath)
      sUID = .UID
   End With
   
   Sql = ""
   Sql = Sql & "USE [master];" & vbNewLine
   Sql = Sql & " Alter Database [" & pDataBaseName & "]"
   Sql = Sql & " SET SINGLE_USER With ROLLBACK IMMEDIATE;" & vbNewLine
   
   Sql = Sql & "RESTORE DATABASE [" & pDataBaseName & "]"
   Sql = Sql & " FROM DISK = '" & pFileBak & "'"
   Sql = Sql & " WITH FILE=1,  NOUNLOAD, REPLACE, STATS = 10"
   Sql = Sql & ", MOVE 'G3R_Data' To '" & DataPath & pDataBaseName & ".mdf'"
   Sql = Sql & ", MOVE 'G3R_Log'  To '" & DataPath & pDataBaseName & ".ldf';" & vbNewLine
   
   Sql = Sql & " Alter Database [" & pDataBaseName & "]"
   Sql = Sql & " SET MULTI_USER With ROLLBACK IMMEDIATE;"
                                       
   Sql = Sql & "USE [" & pxDb.dbName & "];" & vbNewLine
   bOk = pxDb.executa(Sql)
   If bOk Then
      Sql = "USE [" & pDataBaseName & "];" & vbNewLine
      Sql = Sql & "IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'DBA') DROP USER [DBA];" & vbNewLine
      Sql = Sql & "IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'" & sUID & "') DROP USER [" & sUID & "];" & vbNewLine
      Sql = Sql & "CREATE USER [" & sUID & "] FOR LOGIN [" & sUID & "] WITH DEFAULT_SCHEMA=[dbo];" & vbNewLine
      Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_datareader', @membername = '" & sUID & "';" & vbNewLine
      Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_datawriter', @membername = '" & sUID & "';" & vbNewLine
      Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_backupoperator', @membername = '" & sUID & "';" & vbNewLine
      Sql = Sql & "EXEC sys.sp_addsrvrolemember @loginame = N'" & sUID & "', @rolename = N'sysadmin';"
      
      Sql = Sql & "USE [" & pxDb.dbName & "];" & vbNewLine
      bOk = pxDb.executa(Sql)
      If Not bOk Then
        sMsg = sMsg & vbNewLine & "[" & pAlias & "] Erro ao criar usuário de Banco."
      End If
   Else
     sMsg = sMsg & vbNewLine & "[" & pAlias & "] Erro na restauração do Banco."
   End If
   
   F_RESTOREDATABASE = bOk
End Function
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Function F_Backup(Optional pExporta As Boolean = True, Optional pExibeResult As Boolean = True) As Boolean
   Dim Sql As String
   Dim sDBName As String
   Dim sBakName As String
   Dim sPathBk As String
   Dim bResult As Boolean
   Dim sText   As String
   Dim nMouse As Integer
      
   'nMouse = Screen.MousePointer
   'Screen.MousePointer = vbHourglass
      
   sDBName = mvarSys.xDb.dbName
   sBakName = Replace(mvarSys.Propriedades("NMLOJA"), " ", "")  '* NMCOLIGADA
   
   If InStr(sBakName, " ") <> 0 Then
      sBakName = Trim(Mid(sBakName, 1, InStr(sBakName, " ")))
   End If
   sBakName = mvarSys.Propriedades("IDCOLIGADA") & sBakName
   sPathBk = ""
   'sPathBk = GetSpecialFolder(38) & "\ClasseA\Admin\Backups\"
   'sPathBk = mvarSys.PathTmp
   'Call CriarDiretorio(sPathBk)
   
   Sql = "DBCC SHRINKDATABASE(N'" & sDBName & "' )"
   mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Otimizando banco de dados..."
   bResult = mvarSys.xDb.executa(Sql)
   
   Sql = "BACKUP DATABASE [" & sDBName & "]"
   Sql = Sql & " TO  DISK = N'" & sPathBk & sBakName & ".bak'"
   Sql = Sql & " With NOFORMAT"
   Sql = Sql & " , INIT"
   Sql = Sql & " , NAME = N'" & sBakName & "-Full Database Backup'"
   Sql = Sql & " , SKIP"
   Sql = Sql & " , NOREWIND"
   Sql = Sql & " , NOUNLOAD"
   Sql = Sql & " , STATS = 10;"

   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL\", sBakName & ".bak")
   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.1\", sBakName & ".bak")
   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.2\", sBakName & ".bak")
   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.3\", sBakName & ".bak")
   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\", sBakName & ".bak")

   If ExisteArquivo(ResolvePathName(sPathBk) & sBakName & ".bak") Then
      mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Excluindo arquivo " & ResolvePathName(sPathBk) & sBakName & ".bak"
      Call ExcluirArquivo(ResolvePathName(sPathBk) & sBakName & ".bak", False)
   End If
   
   DoEvents
   mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Backup Iniciado..."
   bResult = mvarSys.xDb.executa(Sql)
   mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Backup Finalizado!"
   DoEvents
   
   'nIndex = mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Index
   'If nIndex >= 0 Then
   '   mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Backup Ok"
   'End If
   sText = ""
   If bResult Then
      mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Copiando Backup..."
      sText = sText & "Back-up realizado com sucesso." & vbNewLine
      Call F_CopiarBackup(sPathBk)
   Else
      sText = sText & "Back-up não realizado." & vbNewLine
   End If
   
   If pExporta And bResult Then
      If IsWebConnected Then
         'If ExibirPergunta("Deseja armazenar o Back-up para Internet?") = vbYes Then
            mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Exportando Backup..."
            bResult = F_ExportarBackup(sPathBk)
         'End If
         sText = sText & vbNewLine
         If bResult Then
            sText = sText & "Arquivo salvo no servidor com sucesso!"
         Else
            sText = sText & "Falha na conexão, arquivo NÃO foi salvo no servidor."
         End If
      
      ElseIf bResult Then
         sText = sText & vbNewLine
         sText = "O Sistema não identificou sua conexão de Internet e não pode armazenar seu Backup."
      End If
   End If
   If pExibeResult Then
      If sText <> "" Then ExibirInformacao sText
   End If
   
   F_Backup = bResult
'   Screen.MousePointer = nMouse
End Function
Public Function F_CopiarBackup(Optional sPathBk As String = "") As Boolean
   Dim sBakName   As String
   Dim sFile      As String
   Dim sText      As String
   Dim sZipFile   As String
   Dim nResult    As Double
   
   DoEvents
   sPathBk = ResolvePathName(sPathBk)
   sBakName = mvarSys.Propriedades("IDCOLIGADA") & UnFormat(Replace(mvarSys.Propriedades("NMLOJA"), " ", ""))
   sFile = sBakName & ".bak"
   Call SetPathBak(sPathBk, sFile)
      
   If ExisteArquivo(sPathBk & sFile) Then
      mvarSys.MDI.CommandBars.StatusBar(3).Text = "Compactando " & sPathBk & sFile
      sZipFile = sBakName & ".zip"
      Call Zip(sPathBk & sFile, sPathBk & sZipFile)
      nResult = ExisteArquivo(sPathBk & sZipFile)
      If nResult <> 0 Then
         If mvarSys.Propriedades("PATHBACKUP") = "" Then
            mvarSys.Propriedades("PATHBACKUP") = mvarSys.PathTmp + "Backup\" + CStr(mvarSys.IDCOLIGADA) + mvarSys.Propriedades("NMLOJA") + "\"
         End If
         
         If InStr(mvarSys.ExePath, ":\Sistemas\Dsr\") = 0 Then
            Call CopiarArquivo(sPathBk & sZipFile, mvarSys.ExePath & "Backup\" & sZipFile)
         End If
         'Call CopiarArquivo(sPathBk & sZipFile, mvarsys.PathTmp & sZipFile)
         Call CopiarArquivo(sPathBk & sZipFile, ResolvePathName(mvarSys.Propriedades("PATHBACKUP")) & sZipFile)
      End If
   Else
      sText = sText & "Arquivo copiado com sucesso." & vbNewLine & vbNewLine
      sText = sText & "'" & sPathBk & sFile & "'"
   End If
   F_CopiarBackup = (nResult = 0)
End Function
Public Function F_ExportarBackup(Optional sPathBk As String = "") As Boolean
   Dim sBakName   As String
   Dim sFile      As String
   Dim sText      As String
   Dim sNMLoja    As String
   Dim sZipFile   As String
   Dim sZipDest   As String
   Dim nResult    As Double
   Dim FtpBakPath As String
   Dim bExiste    As Boolean
   
   DoEvents

   sNMLoja = UnFormat(Replace(mvarSys.Propriedades("NMLOJA"), " ", ""))
   sPathBk = ResolvePathName(sPathBk)
   sBakName = mvarSys.Propriedades("IDCOLIGADA") & sNMLoja
   sFile = sBakName & ".bak"
   Call SetPathBak(sPathBk, sFile)

   FtpBakPath = mvarSys.GetParam("FtpBakPath", Default:="", pDescricao:="Caminho do Ftp de Backup")
   FtpBakPath = "/" & sNMLoja & IIf(Mid(FtpBakPath, 1, 1) = "/", "", "/") & FtpBakPath
   
   If ExisteArquivo(sPathBk & sFile) Then
      mvarSys.MDI.CommandBars.StatusBar(3).Text = "Compactando " & sFile
      mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Compactando " & sFile
      sZipFile = sBakName & ".zip"
      If ExisteArquivo(sPathBk & sZipFile) And ExisteArquivo(sPathBk & sFile) Then
         bExiste = (DateDiff("s", FileDateTime(sPathBk & sFile), FileDateTime(sPathBk & sZipFile)) < (3 * 60))
      End If
      If Not bExiste Then
         Call Zip(sPathBk & sFile, sPathBk & sZipFile)
      End If
      nResult = ExisteArquivo(sPathBk & sZipFile)
      If nResult <> 0 Then
         mvarSys.MDI.CommandBars.StatusBar.FindPane(104).Text = "Exportando Backup... "
         
         Dim oFtp        As Object
         Dim bConectado  As Boolean
         
         Set oFtp = ConectarFTP()
         If Not oFtp Is Nothing Then bConectado = oFtp.Conectado
         
         If bConectado Then
            
            With oFtp
               Set .Sys = mvarSys
               
               Call .CriarPasta(FtpBakPath)
               Call .ExcluirArquivoRemoto(FtpBakPath & sZipFile, False)
               DoEvents
               nResult = .EnviarArquivo(sPathBk, sZipFile, FtpBakPath, sZipFile, False)
               DoEvents
               nResult = IIf(nResult = 0, 0, .Conectado)
               
               Call CriarArqNuvem(oFtp)
               .DesconectarFTP
            End With
         End If
         Set oFtp = Nothing
         If nResult = 0 Then
            sText = sText & "Arquivo exportado com sucesso." & vbNewLine & vbNewLine
         Else
            sText = sText & "Arquivo não foi exportado." & vbNewLine & vbNewLine
         End If
      End If
      Call ExcluirArquivo(sPathBk & sZipFile)
   Else
      sText = sText & "Arquivo exportado com sucesso." & vbNewLine & vbNewLine
      sText = sText & "'" & sPathBk & sFile & "'"
   End If
   F_ExportarBackup = (nResult = 0)
   
End Function
Private Sub CriarArqNuvem(oFtp As Object)
   Dim sNMLoja As String
   Dim sBakName As String
   Dim sBakFile As String
   Dim sBakPath As String
   Dim sNFile As String
   Dim sNPath As String
   Dim sNPathFile As String
   Dim QtdFile As Integer
   Dim FtpBakPath As String
   
   sNMLoja = UnFormat(Replace(mvarSys.Propriedades("NMLOJA"), " ", ""))
   sBakName = mvarSys.Propriedades("IDCOLIGADA") + sNMLoja
   sBakPath = ResolvePathName(mvarSys.Propriedades("PATHBACKUP"))
   sBakFile = sBakPath + sBakName + ".zip"
   
   sNFile = "N" + sBakName + ".txt"
   sNPath = mvarSys.Propriedades("PATHBACKUP")
   sNPathFile = sNPath + sNFile
   
   Call AbrirTxt(mvarSys.Propriedades("PATHBACKUP") + sNFile)
   QtdFile = xVal(ReadIniFile(sNPathFile, "sNFile", "QTD", 0))
   
   Call WriteIniFile(sNPathFile, sNFile, "NAME", sBakName + ".zip")
   Call WriteIniFile(sNPathFile, sNFile, "DATE", FileDateTime(sBakFile))
   Call WriteIniFile(sNPathFile, sNFile, "VERSION", GetFileVersionNumber(sBakFile))
   Call WriteIniFile(sNPathFile, sNFile, "LENGHT", FileLen(sBakFile))
   Call WriteIniFile(sNPathFile, sNFile, "PATH", sBakPath)
   
   With oFtp
      FtpBakPath = mvarSys.GetParam("FtpBakPath")
      FtpBakPath = "/" & sNMLoja & IIf(Mid(FtpBakPath, 1, 1) = "/", "", "/") & FtpBakPath
      
      Call .ExcluirArquivoRemoto(FtpBakPath & sNFile, False)
      Call .EnviarArquivo(sNPath, sNFile, FtpBakPath, sNFile, False)
   End With
End Sub
Private Sub SetPathBak(ByRef sPathBk, sFile As String)
   If sPathBk = "" Then
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles%") & "Microsoft SQL Server\MSSQL.1\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles%") & "Microsoft SQL Server\MSSQL.2\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles%") & "Microsoft SQL Server\MSSQL.3\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles%") & "Microsoft SQL Server\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles%(x86)") & "Microsoft SQL Server\MSSQL.1\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles%(x86)") & "Microsoft SQL Server\MSSQL.2\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles%(x86)") & "Microsoft SQL Server\MSSQL.3\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles%(x86)") & "Microsoft SQL Server\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles% (x86)") & "Microsoft SQL Server\MSSQL.1\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles% (x86)") & "Microsoft SQL Server\MSSQL.2\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles% (x86)") & "Microsoft SQL Server\MSSQL.3\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(ResolvePathName("%programfiles% (x86)") & "Microsoft SQL Server\", sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(mvarSys.GetParam("PATHBACKUP", Default:=mvarSys.PathTmp + "Backup\" + CStr(mvarSys.IDCOLIGADA) + mvarSys.Propriedades("NMLOJA") + "\"), sFile)
      If sPathBk = "" Then sPathBk = ProcuraArquivo(mvarSys.PathTmp, sFile)
   End If
End Sub

Public Function F_Enviar_eMail(pFile As String, pTit As String, pHtml As String, pAttach As String, pNome As String, pDestino As String)
   Dim nMouse As Integer
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
                        
   If InStr(UCase(mvarSys.ExePath), "C:\SISTEMAS") <> 0 Then pDestino = "diogenes72@bol.com.br"
   
   With FrmCadMail
      Set .mvarSys = mvarSys
      
      .sAssunto = pTit
      .sNomeDest = pNome
      .sDestino = pDestino
      .sMensagem = pHtml
      .sFileHtm = pFile
      .sAnexo = pAttach
      .Show vbModal
      
      F_Enviar_eMail = .bResult
   End With
   Screen.MousePointer = nMouse
End Function
Public Function GetHtml_eMail(pIDEVENTO As Long, Optional ByRef pFile As String, Optional ByRef pAttach As String, Optional ByRef pNome As String, Optional ByRef pData As String, Optional ByRef pDestino As String) As String
   Dim sPath As String
   Dim sNMTarefa As String
   Dim sPathTmp As String
   Dim sAtta As String
   Dim sFile As String
   Dim sHtml As String
   Dim Sql As String
   Dim Rs  As Object
   Dim i As Integer
   Dim cAux As Collection
   
   sNMTarefa = "Bemvindo"
   
   '********
   '* Definir Pastas de Trabalho
   sPath = mvarSys.ExePath & "HTM\"
   If InStr(UCase(sPath), UCase("C:\Sistemas")) <> 0 Then
      sPath = Environ("programfiles") & "\ClasseA\Projeto3R\Htm\"
   Else
      sPath = mvarSys.ExePath & "HTM\"
   End If
   sPathTmp = Environ("TEMP") & "\" & mvarSys.CODSIS & "\"
   
   '********
   '* Definir Arquivo de Trabalho
   sAtta = sNMTarefa & ".gif"
   sFile = sNMTarefa & ".htm"
   
   '********
   '* Limpar Pasta de Trabalho
   pAttach = sPathTmp & sAtta
   pFile = sPathTmp & sFile
   If ExisteArquivo(pAttach) Then Call ExcluirArquivo(pAttach, False)
   If ExisteArquivo(pFile) Then Call ExcluirArquivo(pFile, False)
   
   '********
   '* Recuperar HTML
   If ExisteArquivo(sPath & sFile) Then
      Call CopiarArquivo(sPath & sFile, pFile)
   Else
      Call ExtractResData(UCase(sNMTarefa), "HTML", pFile)
   End If
   '********
   '* Recuperar Arquivo Anexo
   If ExisteArquivo(sPath & sAtta) Then
      Call CopiarArquivo(sPath & sAtta, pAttach)
   Else
      Call ExtractResData(UCase(sNMTarefa), "GIF", pAttach)
   End If
   
   If ExisteArquivo(pFile) Then sHtml = ReadTextFile(pFile)
   
   If Trim(sHtml) <> "" Then
      Dim sNome As String, sData As String, sDIASEMANA As String, sHORA As String
      Dim sUNIDADE As String, sENDERECO As String, sTELEFONE As String
           
      With mvarSys.xDb
         Sql = "Select E.STARTDATETIME, C.NOME, C.SEXO, C.EMAIL, L.NOME [NMLOJA]"
         Sql = Sql & ", L.ENDERECO, L.TELEFONE1"
         Sql = Sql & " From OEVENTOAGENDA E"
         Sql = Sql & " JOIN OCLIENTE C On E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE"
         Sql = Sql & " JOIN OLOJA L On E.IDLOJA=L.IDLOJA"
         Sql = Sql & " Where E.IDLOJA=" & mvarSys.IDLOJA
         Sql = Sql & " And E.IDEVENTO=" & pIDEVENTO
         If .Abretabela(Sql, Rs) Then
            sNome = Rs("NOME") & ""
            Set cAux = GetWords(sNome)
            sNome = "Sr(a) "
            For i = 1 To cAux.Count
               sNome = sNome & " " & cAux(i)
               If i <> 1 Then
                 If Len(cAux(i)) > 2 Then
                    Exit For
                 End If
               End If
            Next
            sData = Format(Rs("STARTDATETIME"), "DD/MM/YY")
            sDIASEMANA = StrConv(WeekdayName(Weekday(sData)), vbProperCase)
            sHORA = Format(Rs("STARTDATETIME"), "HH:MM")
            sUNIDADE = mvarSys.Propriedades("FromDisplayName") 'Rs("NMLOJA")
            sENDERECO = Rs("ENDERECO")
            sTELEFONE = Rs("TELEFONE1")
                        
            pNome = sNome
            pData = Rs("STARTDATETIME")
            pDestino = Rs("EMAIL")
         End If
      End With
      
      '**************************
      '***** POPULAR HTML *******
      '**************************
      sHtml = Replace(sHtml, "@@NOME", sNome)
      sHtml = Replace(sHtml, "@@DATA", sData)
      sHtml = Replace(sHtml, "@@DIASEMANA", sDIASEMANA)
      sHtml = Replace(sHtml, "@@HORA", sHORA)

      sHtml = Replace(sHtml, "@@UNIDADE", sUNIDADE)
      sHtml = Replace(sHtml, "@@ENDERECO", sENDERECO)
      sHtml = Replace(sHtml, "@@TELEFONE", sTELEFONE)
   End If
   
   i = AbrirTxt(pFile)
   Call EscreverTxt(i, sHtml)
   Call FecharTxt(i)
   
   GetHtml_eMail = sHtml
End Function

Public Function F_EnviarFechamento(sDIA As String, bExibeResult As Boolean) As Boolean
   Dim bResult As Boolean
   Dim nMouse  As Integer
   Dim xMail   As Object
   Dim sFile As String
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   Set xMail = CriarObjeto("CAMail.SendMail")
   Call MontarMail(mvarSys, xMail)
   With xMail
      .Subject = "Fechamento dia " & sDIA
      .Message = GetHtmleMailDiario(sDIA)
      
      If mvarSys.Propriedades("FECHATELAAG") = 1 Then
         sFile = Environ("TEMP") & "\" & "AG" & Format(mvarSys.xDb.SysDate, "ddmmyy") & ".jpg"
         If ExisteArquivo(sFile) Then .Attachment = sFile
      End If
      
      .Connect
      .Send
      .Disconnect
      bResult = .SendSuccesful
      
      If bExibeResult Then
         ExibirInformacao IIf(bResult, "Mensagem enviada com Sucesso!!", "Envio Falhou!!")
      End If
   End With
   Set xMail = Nothing
   
   F_EnviarFechamento = bResult
   Screen.MousePointer = nMouse
End Function
Public Function F_EnviarRPS(sDIA As String, bExibeResult As Boolean) As Boolean
   Dim bResult As Boolean
   Dim nMouse  As Integer
   Dim xMail   As Object
   Dim sFile As String
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   sFile = GerarArquivoRPS(pDTFechamento:=sDIA)
      
   If Trim(sFile) <> "" Then
      Set xMail = CriarObjeto("CAMail.SendMail")
      Call MontarMail(mvarSys, xMail, Sys.GetParam("LstEMailArqFin"))
      
      With xMail
         
         If ExisteArquivo(sFile) Then
            .Subject = "Arquivo: " & GetNameFromPath(sFile)
            .Message = GetHtmleMailRPS(sFile)
            
            If Trim(sFile) <> "" Then .Attachment = sFile
            
            .Connect
            .Send
            .Disconnect
            bResult = .SendSuccesful
            
            If bExibeResult Then
               ExibirInformacao IIf(bResult, "Mensagem enviada com Sucesso!!", "Envio Falhou!!")
            End If
         End If
      End With
      Set xMail = Nothing
   End If
   F_EnviarRPS = bResult
   Screen.MousePointer = nMouse
End Function
Public Function F_EnviarInventario(sDIA As String, bExibeResult As Boolean) As Boolean
   Dim bResult As Boolean
   Dim nMouse  As Integer
   Dim xMail   As Object
   Dim sFile As String
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   If ExisteArquivo(sFile) Then
      Set xMail = CriarObjeto("CAMail.SendMail")
      Call MontarMail(mvarSys, xMail, Sys.GetParam("LstEMailSupri"))
      With xMail
         'sFile = GerarArquivoInventario
         
            .Subject = "Inventário: " & sDIA
            .Message = GetHtmleMailRPS(sFile)
            
            If Trim(sFile) <> "" Then .Attachment = sFile
            
            .Connect
            .Send
            .Disconnect
            bResult = .SendSuccesful
            
            If bExibeResult Then
               ExibirInformacao IIf(bResult, "Mensagem enviada com Sucesso!!", "Envio Falhou!!")
            End If
      End With
      Set xMail = Nothing
   End If
   F_EnviarInventario = bResult
   Screen.MousePointer = nMouse
End Function

Public Function F_EnviarFaturamento(sDIA As String, bExibeResult As Boolean) As Boolean
   Dim bResult As Boolean
   Dim nMouse  As Integer
   Dim xMail   As Object
   Dim sFile As String
   
   Dim i As Integer
   Dim sData As String
   Dim nULTFAT As Integer
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   If mvarSys.GetParam("NFE", Default:="0") = "0" Then
      Exit Function
   Else
      '* Se o Último mês faturado não é mês passado, ou seja, a diferença do mês atual é maisor que 1
      nULTFAT = xVal(mvarSys.GetParam("ULTFAT", pSECAO:="FIN", Default:="0"))
      If InArray(Month(sDIA) - nULTFAT, Array(1, -11)) Then
         Exit Function
      End If
   End If
   
   sData = DateAdd("d", -1, "01" & Format(sDIA, "/mm/yyyy"))
   
   Set xMail = CriarObjeto("CAMail.SendMail")
   Call MontarMail(mvarSys, xMail, Sys.GetParam("LstEMailArqFin"))
   With xMail
      .Subject = mvarSys.Propriedades("RZABREV")
      .Subject = .Subject & " - Faturamento - "
      .Subject = .Subject & StrConv(MonthName(Month(sData), True), vbProperCase)
      .Subject = .Subject & "/" & Format(sData, "yy")
      .Message = GetHtmleMailFaturamento(sData)
      
      'sFile = GerarArquivoRPS
      'If ExisteArquivo(sFile) Then .Attachment = sFile
      
      .Connect
      .Send
      .Disconnect
      bResult = .SendSuccesful
      
      If bExibeResult Then
         ExibirInformacao IIf(bResult, "Mensagem enviada com Sucesso!!", "Envio Falhou!!")
      End If
   End With
   If bResult Then
      Call mvarSys.SaveParam("ULTFAT", Month(sData), pCATPARAM:="FIN", pDSCPARAM:="Último Mês de Faturamento Enviado")
   End If
   F_EnviarFaturamento = bResult
   Screen.MousePointer = nMouse
End Function
Public Function F_EnviarOcorrencias(sDIA As String, bExibeResult As Boolean) As Boolean
   Dim bResult As Boolean
   Dim nMouse  As Integer
   Dim xMail   As Object
   Dim sFile As String
   
   Dim Sql As String
   Dim sTexto As String
      
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
      
   Sql = "Select * "
   Sql = Sql & " From ODIARIO"
   Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA
   Sql = Sql & " And DTDIARIO=" & SqlDate(sDIA)
      
   If mvarSys.xDb.Abretabela(Sql) Then
      sTexto = mvarSys.xDb.RsAux("DSCDIARIO")
      If sTexto <> "" Then
         
         Set xMail = CriarObjeto("CAMail.SendMail")
         Call MontarMail(mvarSys, xMail, Sys.GetParam("LstEMailOcorrencia"))
         
         With xMail
            .Subject = "Ocorrências: " & UCase(Format(sDIA, "dd/mmm"))
            .Message = Replace(sTexto, Chr(13) + Chr(10), "<br>") 'sTexto 'GetHtmleMailDiario(sDIA)
            
            .Connect
            .Send
            .Disconnect
            bResult = .SendSuccesful
            
            If bExibeResult Then
               ExibirInformacao IIf(bResult, "Mensagem enviada com Sucesso!!", "Envio Falhou!!")
            End If
         End With
         Set xMail = Nothing
      End If
   End If
   
   F_EnviarOcorrencias = bResult
   Screen.MousePointer = nMouse
End Function
Public Function GetHtmleMailRPS(pArquivo As String) As String
   Dim sPath As String
   Dim sFile As String
   Dim sPathFile As String
   Dim sHtml As String
   Dim Sql As String
   Dim nRPS1 As Long
   Dim i As Integer
   
   
   sFile = "RPS3R.htm"
   sPath = Environ("TEMP") & "\" & mvarSys.CODSIS & "\"
   sPathFile = sPath & sFile
   
   Call ExtractResData("RPS", "HTML", sPathFile)
   If ExisteArquivo(sPathFile) Then
      sHtml = ReadTextFile(sPathFile)
      Call ExcluirArquivo(sPathFile, False)
   End If
   If sHtml = "" Then
      Call CopiarArquivo(App.Path & "\" & sFile, sPathFile)
      sHtml = ReadTextFile(sPathFile)
   End If
   
   If Trim(sHtml) <> "" Then
      Dim nArq As Integer, nTPLinha As Integer
      Dim sNMLoja As String, sCODLOTE As String
      Dim sDTINI As String, sDTFIM As String
      Dim sQTDRPS As String, sVLTOTAL As String
      Dim sDETALHE As String, sLinha As String
      Dim sVLVENDA As String
      
      sNMLoja = UCase(mvarSys.Propriedades("NMLOJA"))
      sCODLOTE = Mid(Right(pArquivo, 12), 4, 5)
           
      nArq = 1
      sDETALHE = "<br>"
      sDETALHE = sDETALHE & "RECIBO&nbsp;&nbsp;&nbsp;&nbsp;DATA EMISS &nbsp;&nbsp;&nbsp;&nbsp; VALOR &nbsp;&nbsp;&nbsp;&nbsp; C.P.F. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; NOME" & "<br>"
      sDETALHE = sDETALHE & String(80, "=") & "<br>" & vbNewLine  '103
      'sDETALHE = " 1.000284  20/10/2011  R$     55,00  000.000.000-00  Ana Lúcia Nobre de Miranda Marques"

      Open pArquivo For Input As #nArq
      Do While Not EOF(nArq)
         Line Input #nArq, sLinha
         nTPLinha = xVal(Mid(sLinha, 1, 2))
         Select Case nTPLinha
            Case 10
               sDTINI = Mid(sLinha, 42, 2) & "/" & Mid(sLinha, 40, 2) & "/" & Mid(sLinha, 36, 4)
               sDTFIM = Mid(sLinha, 50, 2) & "/" & Mid(sLinha, 48, 2) & "/" & Mid(sLinha, 44, 4)
            Case 20
                  nRPS1 = xVal(IIf(nRPS1 = 0, xVal(Mid(sLinha, 9, 15)), nRPS1))
                                    
                  sDETALHE = sDETALHE & Space(1) & Trim(Mid(sLinha, 4, 5)) & "." & StrZero(xVal(Mid(sLinha, 9, 15)), 6) & "&nbsp;"
                  sDETALHE = sDETALHE & Space(1) & Mid(sLinha, 30, 2) & "/" & Mid(sLinha, 28, 2) & "/" & Mid(sLinha, 24, 4) & "&nbsp;"
                  sDETALHE = sDETALHE & Space(1) & "R$ "
                  For i = 1 To 8 - Len(ValBr(xVal(Mid(sLinha, 701, 15)) / 100, 2))
                     sDETALHE = sDETALHE & "&nbsp;"
                  Next
                  sDETALHE = sDETALHE & Space(1) & ValBr(xVal(Mid(sLinha, 701, 15)) / 100, 2) & "&nbsp;"
                  sDETALHE = sDETALHE & Space(1) & Mid(sLinha, 37, 3) & "." & Mid(sLinha, 40, 3) & "." & Mid(sLinha, 43, 3) & "-" & Mid(sLinha, 45, 2) & "&nbsp;"
                  sDETALHE = sDETALHE & Space(1) & Trim(Mid(sLinha, 78, 115))
                  sDETALHE = sDETALHE & "<br>" & vbNewLine
            Case 90
               sQTDRPS = xVal(Mid(sLinha, 3, 8))
               sVLTOTAL = "R$ " & ValBr(xVal(Mid(sLinha, 11, 15)) / 100, 2)
               
               sCODLOTE = sCODLOTE & " (" & nRPS1 & "-" & nRPS1 + xVal(sQTDRPS) - 1 & ")"
         End Select
      Loop
      
      Sql = "Select Sum(I.QTDVENDA*VLUNIT)-Sum(VLDESC)+Sum(VLACRESC) [VLVENDA]"
      Sql = Sql & " From  CVENDA V"
      Sql = Sql & " Join CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA"
      Sql = Sql & " Join SPRODUTO P    On I.IDLOJA=P.IDLOJA AND I.IDPROD=P.IDPROD And P.ESERVICO=1"
      Sql = Sql & " Where V.IDLOJA= " & SqlNum(mvarIDLOJA)
      Sql = Sql & " And V.DTVENDA >= " & SqlDate(sDTINI & " 00:00")
      Sql = Sql & " And V.DTVENDA <= " & SqlDate(sDTFIM & " 23:59")
      
      If mvarSys.xDb.Abretabela(Sql) Then
         If xVal(mvarSys.xDb.RsAux("VLVENDA") & "") = xVal(sVLTOTAL) Then
            sVLVENDA = "."
         Else
            sVLVENDA = "Venda: R$" & ValBr(mvarSys.xDb.RsAux("VLVENDA") & "") & " ?"
         End If
      End If
      
      Close #nArq
      
      '**************************
      '***** POPULAR HTML *******
      '**************************
      sHtml = Replace(sHtml, "@@NMLOJA", sNMLoja)
      sHtml = Replace(sHtml, "@@CODLOTE", sCODLOTE)
      sHtml = Replace(sHtml, "@@DTINI", sDTINI)
      sHtml = Replace(sHtml, "@@DTFIM", sDTFIM)
      sHtml = Replace(sHtml, "@@QTDRPS", sQTDRPS)
      sHtml = Replace(sHtml, "@@VLTOTAL", sVLTOTAL)
      sHtml = Replace(sHtml, "@@VLVENDA", sVLVENDA)
      sHtml = Replace(sHtml, "@@DETALHE", sDETALHE)
      
   End If
   
   GetHtmleMailRPS = sHtml
End Function
Public Function GetHtmleMailInventario(pArquivo As String) As String
   Dim sPath As String
   Dim sFile As String
   Dim sPathFile As String
   Dim sHtml As String
   Dim Sql As String
   Dim nRPS1 As Long
   Dim i As Integer
   
   sFile = "RPS3R.Htm"
   sPath = Environ("TEMP") & "\" & mvarSys.CODSIS & "\"
   sPathFile = sPath & sFile
   
   Call ExtractResData("RPS", "HTML", sPathFile)
   If ExisteArquivo(sPathFile) Then
      sHtml = ReadTextFile(sPathFile)
      Call ExcluirArquivo(sPathFile, False)
   End If
   If sHtml = "" Then
      Call CopiarArquivo(App.Path & "\" & sFile, sPathFile)
      sHtml = ReadTextFile(sPathFile)
   End If
   
   If Trim(sHtml) <> "" Then
      Dim nArq As Integer, nTPLinha As Integer
      Dim sNMLoja As String, sCODLOTE As String
      Dim sDTINI As String, sDTFIM As String
      Dim sQTDRPS As String, sVLTOTAL As String
      Dim sDETALHE As String, sLinha As String
      Dim sVLVENDA As String
      
      sNMLoja = UCase(mvarSys.Propriedades("NMLOJA"))
      sCODLOTE = Mid(Right(pArquivo, 12), 4, 5)
           
      nArq = 1
      sDETALHE = "<br>"
      sDETALHE = sDETALHE & "RECIBO&nbsp;&nbsp;&nbsp;&nbsp;DATA EMISS &nbsp;&nbsp;&nbsp;&nbsp; VALOR &nbsp;&nbsp;&nbsp;&nbsp; C.P.F. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; NOME" & "<br>"
      sDETALHE = sDETALHE & String(80, "=") & "<br>" & vbNewLine  '103
      'sDETALHE = " 1.000284  20/10/2011  R$     55,00  000.000.000-00  Ana Lúcia Nobre de Miranda Marques"

      Open pArquivo For Input As #nArq
      Do While Not EOF(nArq)
         Line Input #nArq, sLinha
         nTPLinha = xVal(Mid(sLinha, 1, 2))
         Select Case nTPLinha
            Case 10
               sDTINI = Mid(sLinha, 42, 2) & "/" & Mid(sLinha, 40, 2) & "/" & Mid(sLinha, 36, 4)
               sDTFIM = Mid(sLinha, 50, 2) & "/" & Mid(sLinha, 48, 2) & "/" & Mid(sLinha, 44, 4)
            Case 20
                  nRPS1 = xVal(IIf(nRPS1 = 0, xVal(Mid(sLinha, 9, 15)), nRPS1))
                                    
                  sDETALHE = sDETALHE & Space(1) & Trim(Mid(sLinha, 4, 5)) & "." & StrZero(xVal(Mid(sLinha, 9, 15)), 6) & "&nbsp;"
                  sDETALHE = sDETALHE & Space(1) & Mid(sLinha, 30, 2) & "/" & Mid(sLinha, 28, 2) & "/" & Mid(sLinha, 24, 4) & "&nbsp;"
                  sDETALHE = sDETALHE & Space(1) & "R$ "
                  For i = 1 To 8 - Len(ValBr(xVal(Mid(sLinha, 701, 15)) / 100, 2))
                     sDETALHE = sDETALHE & "&nbsp;"
                  Next
                  sDETALHE = sDETALHE & Space(1) & ValBr(xVal(Mid(sLinha, 701, 15)) / 100, 2) & "&nbsp;"
                  sDETALHE = sDETALHE & Space(1) & Mid(sLinha, 37, 3) & "." & Mid(sLinha, 40, 3) & "." & Mid(sLinha, 43, 3) & "-" & Mid(sLinha, 45, 2) & "&nbsp;"
                  sDETALHE = sDETALHE & Space(1) & Trim(Mid(sLinha, 78, 115))
                  sDETALHE = sDETALHE & "<br>" & vbNewLine
            Case 90
               sQTDRPS = xVal(Mid(sLinha, 3, 8))
               sVLTOTAL = "R$ " & ValBr(xVal(Mid(sLinha, 11, 15)) / 100, 2)
               
               sCODLOTE = sCODLOTE & " (" & nRPS1 & "-" & nRPS1 + xVal(sQTDRPS) - 1 & ")"
         End Select
      Loop
      
      Sql = "Select Sum(I.QTDVENDA*VLUNIT)-Sum(VLDESC) [VLVENDA]"
      Sql = Sql & " From  CVENDA V"
      Sql = Sql & " Join CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA"
      Sql = Sql & " Join SPRODUTO P    On I.IDLOJA=P.IDLOJA AND I.IDPROD=P.IDPROD And P.ESERVICO=1"
      Sql = Sql & " Where V.IDLOJA= " & SqlNum(mvarIDLOJA)
      Sql = Sql & " And V.DTVENDA >= " & SqlDate(sDTINI & " 00:00")
      Sql = Sql & " And V.DTVENDA <= " & SqlDate(sDTFIM & " 23:59")
      
      If mvarSys.xDb.Abretabela(Sql) Then
         If xVal(mvarSys.xDb.RsAux("VLVENDA") & "") = xVal(sVLTOTAL) Then
            sVLVENDA = "."
         Else
            sVLVENDA = "Venda: R$" & ValBr(mvarSys.xDb.RsAux("VLVENDA") & "") & " ?"
         End If
      End If
      
      Close #nArq
      
      '**************************
      '***** POPULAR HTML *******
      '**************************
      sHtml = Replace(sHtml, "@@NMLOJA", sNMLoja)
      sHtml = Replace(sHtml, "@@CODLOTE", sCODLOTE)
      sHtml = Replace(sHtml, "@@DTINI", sDTINI)
      sHtml = Replace(sHtml, "@@DTFIM", sDTFIM)
      sHtml = Replace(sHtml, "@@QTDRPS", sQTDRPS)
      sHtml = Replace(sHtml, "@@VLTOTAL", sVLTOTAL)
      sHtml = Replace(sHtml, "@@VLVENDA", sVLVENDA)
      sHtml = Replace(sHtml, "@@DETALHE", sDETALHE)
      
   End If
   
   GetHtmleMailInventario = sHtml
End Function
Public Function GetHtmleMailDiario(pDia As String) As String
   Dim sFile As String
   Dim sPathFile As String
   Dim sHtml As String
   
   sFile = "Diario3R.htm"
   sPathFile = Environ("TEMP") & "\" & mvarSys.CODSIS & "\" & sFile
   
   Call ExtractResData("DIARIO", "HTML", sPathFile)
   If ExisteArquivo(sPathFile) Then
      sHtml = ReadTextFile(sPathFile)
      Call ExcluirArquivo(sPathFile, False)
   End If
   
   If Trim(sHtml) <> "" Then
      Dim Sql As String
      Dim oMovCC As Object
      
      Dim nAno As Integer
      Dim nMes As Integer
      Dim sAux As String
      
      Dim sNMLoja As String, sANO2 As String, sANO1 As String, sANO0 As String
      Dim sQTDSDIA As String, sQTDPDIA As String, sQTDTDIA As String
      Dim sVLSDIA As String, sVLPDIA As String
      Dim sVLTDIA1 As String, sVLTDIA2 As String
      Dim sVLDescDIA As String, sVLDescMes As String
      Dim sVLFATDIA As String, sVLFATMes As String
      Dim sVLDINDIA As String, sVLDEBDIA As String, sVLCREDIA As String, sVLCHEDIA As String
      Dim sVLCAIXADIA As String, sQTDPROD As String
      
      Dim sQTDSMES As String, sQTDPMES As String, sQTDTMES As String
      Dim sVLSMES As String, sVLPMES As String
      Dim sVLTMES1 As String, sVLTMES2 As String
      Dim sVLDINMES As String, sVLDEBMES As String, sVLCREMES As String, sVLCHEMES As String
      Dim sDIATRABMES As String, sMEDMES As String
      
      Dim sQTDT02 As String, sQTDT01 As String, sQTDT00 As String
      Dim sANO02 As String, sANO01 As String, sANO00 As String
      Dim sMES As String, sDIA As String
      
      Dim sPREVVENDAMES As String, sPREVPRODMES As String
      Dim sLinha As String, sLinhaPDia As String, sLinhaPMes As String
      Dim sVLTPDIA As String, sVLTPMES As String
      
      sNMLoja = mvarSys.Propriedades("NMLOJA")
      sANO0 = Format(mvarSys.xDb.SysDate(), "yyyy")
      sANO1 = xVal(sANO0) - 1
      sANO2 = xVal(sANO0) - 2
      
      sQTDSDIA = "--": sQTDPDIA = "--": sQTDTDIA = "--"
      sVLSDIA = "--": sVLPDIA = "--": sVLTDIA1 = "--": sVLTDIA2 = "--"
      sVLDINDIA = "--": sVLDEBDIA = "--": sVLCREDIA = "--": sVLCHEDIA = "--"
      sVLCAIXADIA = "--": sQTDPROD = "--"
      sVLDescDIA = "--"
      
      sQTDSMES = "--": sQTDPMES = "--": sQTDTMES = "--"
      sVLSMES = "--": sVLPMES = "--": sVLTMES1 = "--": sVLTMES2 = "--"
      sVLDINMES = "--": sVLDEBMES = "--": sVLCREMES = "--": sVLCHEMES = "--"
      sDIATRABMES = "--": sMEDMES = "--"
      sVLDescMes = "--"
      
      sQTDT02 = "--": sQTDT01 = "--":  sQTDT00 = "--"
      
      sPREVVENDAMES = "--": sPREVPRODMES = "--"
      
      sDIA = " " & Format(pDia, "dd/mm/yy") & " - " & StrConv(LCase$(Format(pDia, "DDDD")), vbProperCase)
      sMES = " " & StrConv(LCase$(Format(pDia, "MMMM")), vbProperCase) & "/" & Format(pDia, "YY")
      
      sNMLoja = mvarSys.Propriedades("NMLOJA")
      sANO0 = Year(pDia)
      sANO1 = xVal(sANO0) - 1
      sANO2 = xVal(sANO0) - 2
      
      
      With mvarSys.xDb
         '@@LINHAPDIA     @@LINHAPMES
         sLinha = sLinha & "<tr>"
         sLinha = sLinha & "  <td bgcolor=""" & "#E8FADC""" & "><strong><font size=""" & "2""" & "><font color=""" & "#006633""" & ">@@DSCTRATAMENTO</font></font></strong></td>"
         sLinha = sLinha & "  <td><p align=""" & "center""" & "><font size=""" & "2""" & ">@@QTDTRATAMENTO</font></p></td>"
         sLinha = sLinha & "</tr>"
         
         '************
         '* Define Linhas de Serviços do Dia
         Sql = "Select E.DSCTRATAMENTO, Count(*) [QTDTRATAMENTO]" & vbNewLine
         Sql = Sql & " From OSESSAO S" & vbNewLine
         Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Join OTPTRATAMENTO E On E.IDLOJA=S.IDLOJA And E.IDTPTRATAMENTO=S.IDTPTRATAMENTO" & vbNewLine
         Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Convert(varchar, A.DTATEND,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO<>1" & vbNewLine
         Sql = Sql & " Group By E.DSCTRATAMENTO" & vbNewLine
         Sql = Sql & " Having Count(*)>0"
         
         sLinhaPDia = ""
         sVLTPDIA = "0"
         If .Abretabela(Sql) Then
            While Not .RsAux.EOF
               sAux = sLinha & vbNewLine
               sAux = Replace(sAux, "@@DSCTRATAMENTO", .RsAux("DSCTRATAMENTO"))
               sAux = Replace(sAux, "@@QTDTRATAMENTO", xVal(.RsAux("QTDTRATAMENTO"), 0))
               sLinhaPDia = sLinhaPDia & sAux
               sVLTPDIA = xVal(sVLTPDIA) + xVal(.RsAux("QTDTRATAMENTO"), 0)
               .RsAux.MoveNext
            Wend
         End If
      
         '************
         '* Define Linhas de Serviços do Mês
         Sql = "Select E.DSCTRATAMENTO, Count(*) [QTDTRATAMENTO]" & vbNewLine
         Sql = Sql & " From OSESSAO S" & vbNewLine
         Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Join OTPTRATAMENTO E On E.IDLOJA=S.IDLOJA And E.IDTPTRATAMENTO=S.IDTPTRATAMENTO" & vbNewLine
         Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Year(A.DTATEND)=" & Format(pDia, "yyyy") & vbNewLine
         Sql = Sql & " And Month(A.DTATEND)=" & Format(pDia, "mm") & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO<>1" & vbNewLine
         Sql = Sql & " Group By E.DSCTRATAMENTO" & vbNewLine
         Sql = Sql & " Having Count(*)>0"
         
         sLinhaPMes = ""
         sVLTPMES = "0"
         If .Abretabela(Sql) Then
            While Not .RsAux.EOF
               sAux = sLinha & vbNewLine
               sAux = Replace(sAux, "@@DSCTRATAMENTO", .RsAux("DSCTRATAMENTO"))
               sAux = Replace(sAux, "@@QTDTRATAMENTO", xVal(.RsAux("QTDTRATAMENTO"), 0))
               sLinhaPMes = sLinhaPMes & sAux
               sVLTPMES = xVal(sVLTPMES) + xVal(.RsAux("QTDTRATAMENTO"), 0)
               .RsAux.MoveNext
            Wend
         End If
            
         '************
         '* Serviços e Produtos Dia
         'Sql = "Select P.ESERVICO, SUM(I.QTDVENDA) [QTD], SUM(I.QTDVENDA*I.VLUNIT) [VALOR]"
         Sql = "Select P.ESERVICO, SUM(I.QTDVENDA) [QTD], SUM(I.QTDVENDA*I.VLUNIT) [VALOR]" & vbNewLine
         Sql = Sql & " From CITENSVENDA I" & vbNewLine
         Sql = Sql & " Join CVENDA V On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA" & vbNewLine
         Sql = Sql & " Join SPRODUTO P On I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD AND P.EVENDA=1" & vbNewLine
         Sql = Sql & " Where I.IDLOJA = " & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Convert(VarChar, V.DTVENDA,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
         Sql = Sql & " And isNull(V.IDFATURA,0)=0" & vbNewLine
         Sql = Sql & " And isNull(V.FLGCANCELADA,0)=0" & vbNewLine
         Sql = Sql & " And V.SITVENDA=10" & vbNewLine
         Sql = Sql & " Group By P.ESERVICO"
         If .Abretabela(Sql) Then
            While Not .RsAux.EOF
               If .RsAux("ESERVICO") = 1 Then
                  sQTDSDIA = ValBr(.RsAux("QTD"), 0)
                  sVLSDIA = "R$ " & ValBr(.RsAux("VALOR"))
               Else
                  sQTDPDIA = ValBr(.RsAux("QTD"), 0)
                  sVLPDIA = "R$ " & ValBr(.RsAux("VALOR"))
               End If
               sVLTDIA1 = xVal(sVLTDIA1) + .RsAux("VALOR")
               .RsAux.MoveNext
            Wend
            
            sVLDescDIA = "0"
            Sql = " Select Sum(V.VLDESC) [DESC]" & vbNewLine
            Sql = Sql & " From CVENDA V" & vbNewLine
            Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
            Sql = Sql & " And Convert(VarChar, V.DTVENDA,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
            If .Abretabela(Sql) Then
               sVLDescDIA = xVal(.RsAux("DESC") & "")
            End If
            sVLFATDIA = "0"
            Sql = " Select isNull(Sum(F.VALOR),0) [FAT]" & vbNewLine
            Sql = Sql & " From FFATURA F " & vbNewLine
            Sql = Sql & " Left Join CVENDA V ON F.IDLOJA=V.IDLOJA And F.IDVENDA=V.IDVENDA" & vbNewLine
            Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
            Sql = Sql & " And Convert(VarChar, F.DTEMISSAO,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
            Sql = Sql & " And ("
            Sql = Sql & " V.DTVENDA Is Null Or"
            Sql = Sql & " V.DTVENDA>" & SqlStr(Format(pDia, "dd/mm/yyyy 23:59"))
            Sql = Sql & ")"
            If .Abretabela(Sql) Then
               sVLFATDIA = xVal(.RsAux("FAT") & "")
            End If
            
            sVLTDIA1 = "R$ " & ValBr(xVal(sVLTDIA1) - xVal(sVLDescDIA) - xVal(sVLFATDIA))
            sVLDescDIA = "(R$ " & ValBr(sVLDescDIA) & ")"
            sVLFATDIA = "(R$ " & ValBr(sVLFATDIA) & ")"
         End If
         
         '************
         '* Testes do Dia
         sQTDTDIA = "0"
         Sql = "Select COUNT(*) [QTD]" & vbNewLine
         Sql = Sql & " From OSESSAO S" & vbNewLine
         Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Convert(varchar, A.DTATEND,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO=1" & vbNewLine
         'Sql = Sql & " Having Count(*)>0"
         If .Abretabela(Sql) Then
            sQTDTDIA = ValBr(.RsAux("QTD"), 0)
         End If
         Sql = "Select isNull(COUNT(*),0) [QTD]" & vbNewLine
         Sql = Sql & " From OSESSAO S" & vbNewLine
         Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Join OTPTRATAMENTO T On S.IDLOJA=T.IDLOJA And S.IDTPTRATAMENTO=T.IDTPTRATAMENTO" & vbNewLine
         Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Convert(varchar, A.DTATEND,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO=2" & vbNewLine
         Sql = Sql & " And S.NUMSESSAO = 1" & vbNewLine
         Sql = Sql & " And T.FLGAVALIACAO=1"
         If .Abretabela(Sql) Then
            sQTDTDIA = sQTDTDIA & " / " & ValBr(.RsAux("QTD"), 0)
         Else
            sQTDTDIA = sQTDTDIA & " / 0"
         End If
         
         '************
         '* Forma de Pagamento Dia
         Sql = "Select P.IDFORMAPGTO, SUM(P.VLPGTO)-SUM(V.VLTROCO) [Valor]" & vbNewLine
         Sql = Sql & " From CPGTOSVENDA P "
         Sql = Sql & " Join CVENDA V On P.IDLOJA=V.IDLOJA And P.IDVENDA=V.IDVENDA" & vbNewLine
         Sql = Sql & " Where P.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         'Sql = Sql & " And Convert(varchar, P.DTPGTO,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
         Sql = Sql & " And Convert(varchar, V.DTVENDA,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
         Sql = Sql & " And isNull(V.IDFATURA,0)=0"
         Sql = Sql & " And isNull(V.FLGCANCELADA,0)=0"
         Sql = Sql & " And V.SITVENDA=10"
         Sql = Sql & " Group By P.IDFORMAPGTO"
         If .Abretabela(Sql) Then
            While Not .RsAux.EOF
               If .RsAux("IDFORMAPGTO") = 1 Then '* Dinheiro
                  sVLDINDIA = "R$ " & ValBr(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 2 Then '* Débito
                  sVLDEBDIA = "R$ " & ValBr(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 3 Then '* Crédito
                  sVLCREDIA = "R$ " & ValBr(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 4 Then '* Cheque
                  sVLCHEDIA = "R$ " & ValBr(.RsAux("Valor"))
               End If
               sVLTDIA2 = xVal(sVLTDIA2) + xVal(.RsAux("Valor"))
               .RsAux.MoveNext
            Wend
            sVLTDIA2 = "R$ " & ValBr(xVal(sVLTDIA2))
         End If
         
         '************
         '* Conta Caixa
         Set oMovCC = CriarObjeto("Financ3R.TL_MOVCC")
         Set oMovCC.Sys = mvarSys
         oMovCC.IDLOJA = mvarIDLOJA
         oMovCC.IDCONTA = 1
         Sql = oMovCC.QueryGrid
         Sql = Sql & " Order By NLAN"
         If .Abretabela(Sql) Then
            .RsAux.MoveLast
            sVLCAIXADIA = "R$ " & ValBr(.RsAux("Saldo"))
         End If
         
         '************
         '* Estoque
         Sql = "Select E.IDPROD, E.SLDDISPONIVEL" & vbNewLine
         Sql = Sql & " From SESTPROD E" & vbNewLine
         Sql = Sql & " Join SPRODUTO P On P.IDLOJA=E.IDLOJA And P.IDPROD=E.IDPROD" & vbNewLine
         Sql = Sql & " Where E.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And P.EVENDA=1" & vbNewLine
         Sql = Sql & " And P.ESERVICO=0"
         
         Sql = "Select Sum(QTDITEM) [SLDDISPONIVEL]" & vbNewLine
         Sql = Sql & " From SMOVEST M Join SPRODUTO P On P.IDLOJA=M.IDLOJA AND P.IDPROD=M.IDPROD" & vbNewLine
         Sql = Sql & " Where M.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And P.EVENDA = 1" & vbNewLine
         Sql = Sql & " And P.ESERVICO=0"
         If .Abretabela(Sql) Then
            sQTDPROD = ValBr(.RsAux("SLDDISPONIVEL") & "", 0)
         End If
         
         '************
         '* Serviços e Produtos Mês
         Sql = "Select P.ESERVICO, SUM(I.QTDVENDA) [QTD], SUM(I.QTDVENDA*I.VLUNIT) [VALOR]" & vbNewLine
         Sql = Sql & " From CITENSVENDA I" & vbNewLine
         Sql = Sql & " Join CVENDA V On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA" & vbNewLine
         Sql = Sql & " Join SPRODUTO P On I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD AND P.EVENDA=1" & vbNewLine
         Sql = Sql & " Where I.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Month(V.DTVENDA)=" & SqlNum(Month(pDia)) & vbNewLine
         Sql = Sql & " And Year(V.DTVENDA)=" & SqlNum(Year(pDia)) & vbNewLine
         Sql = Sql & " Group By P.ESERVICO"
         If .Abretabela(Sql) Then
            While Not .RsAux.EOF
               If .RsAux("ESERVICO") = 1 Then
                  sQTDSMES = ValBr(.RsAux("QTD"), 0)
                  sVLSMES = "R$ " & ValBr(.RsAux("VALOR"))
               Else
                  sQTDPMES = ValBr(.RsAux("QTD"), 0)
                  sVLPMES = "R$ " & ValBr(.RsAux("VALOR"))
               End If
               sVLTMES1 = xVal(sVLTMES1) + .RsAux("VALOR")
               .RsAux.MoveNext
            Wend
            
            sVLDescMes = "0"
            Sql = " Select isNull(Sum(V.VLDESC),0) [DESC]" & vbNewLine
            Sql = Sql & " From CVENDA V" & vbNewLine
            Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
            Sql = Sql & " And Month(V.DTVENDA)=" & SqlNum(Month(pDia)) & vbNewLine
            Sql = Sql & " And Year(V.DTVENDA)=" & SqlNum(Year(pDia))
            If .Abretabela(Sql) Then
               sVLDescMes = xVal(.RsAux("DESC") & "")
            End If
            sVLFATMes = "0"
            Sql = "Select isNull(Sum(F.VALOR),0) [FAT]" & vbNewLine
            Sql = Sql & " From FFATURA F " & vbNewLine
            Sql = Sql & " Left Join CVENDA V ON F.IDLOJA=V.IDLOJA And F.IDVENDA=V.IDVENDA" & vbNewLine
            Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
            Sql = Sql & " And Convert(VarChar, F.DTEMISSAO,103)=" & SqlStr(Format(pDia, "dd/mm/yyyy")) & vbNewLine
            Sql = Sql & " And ("
            Sql = Sql & " V.DTVENDA Is Null Or"
            Sql = Sql & " V.DTVENDA>" & SqlStr("01/" & Format(CStr(CDate(Format(pDia, "28/mm/yyyy")) + 5), "mm/yyyy 00:00"))
            Sql = Sql & ")"
            If .Abretabela(Sql) Then
               sVLFATMes = xVal(.RsAux("FAT") & "")
            End If
            
            sVLTMES1 = "R$ " & ValBr(xVal(sVLTMES1) - xVal(sVLDescMes) - xVal(sVLFATMes))
            sVLDescMes = "(R$ " & ValBr(sVLDescMes) & ")"
            sVLFATMes = "(R$ " & ValBr(sVLFATMes) & ")"
         End If
         
         '************
         '* Testes do Mês
         Sql = "Select isNull(COUNT(*),0) [QTD]" & vbNewLine
         Sql = Sql & " From OSESSAO S" & vbNewLine
         Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Month(A.DTATEND) = " & SqlNum(Month(pDia)) & vbNewLine
         Sql = Sql & " And Year(A.DTATEND) = " & SqlNum(Year(pDia)) & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO=1"
         Sql = Sql & " Having Count(*)>0"
         If .Abretabela(Sql) Then
            If xVal(.RsAux("QTD")) > 0 Then
               sQTDTMES = ValBr(.RsAux("QTD"), 0)
            End If
         End If
         Sql = "Select isNull(COUNT(*),0) [QTD]" & vbNewLine
         Sql = Sql & " From OSESSAO S" & vbNewLine
         Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Join OTPTRATAMENTO T On S.IDLOJA=T.IDLOJA And S.IDTPTRATAMENTO=T.IDTPTRATAMENTO"
         Sql = Sql & " Where S.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Month(A.DTATEND) = " & SqlNum(Month(pDia)) & vbNewLine
         Sql = Sql & " And Year(A.DTATEND) = " & SqlNum(Year(pDia)) & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO=2" & vbNewLine
         Sql = Sql & " And S.NUMSESSAO = 1" & vbNewLine
         Sql = Sql & " And T.FLGAVALIACAO=1"
         If .Abretabela(Sql) Then
            sQTDTMES = sQTDTMES & " / " & ValBr(.RsAux("QTD"), 0)
         Else
            sQTDTMES = sQTDTMES & " / 0"
         End If
         
                           
         '************
         '* Forma de Pagamento Mês
'         Sql = "Select P.IDFORMAPGTO, SUM(P.VLPGTO) [Valor]" & vbNewLine
'         Sql = Sql & " From CPGTOSVENDA P" & vbNewLine
'         Sql = Sql & " Where P.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
'         Sql = Sql & " And Month(P.DTPGTO)=" & SqlNum(Month(pDia)) & vbNewLine
'         Sql = Sql & " And Year(P.DTPGTO)=" & SqlNum(Year(pDia)) & vbNewLine
'         Sql = Sql & " Group By P.IDFORMAPGTO"
         Sql = "Select P.IDFORMAPGTO, SUM(P.VLPGTO)-SUM(V.VLTROCO) [Valor]" & vbNewLine
         Sql = Sql & " From CPGTOSVENDA P "
         Sql = Sql & " Join CVENDA V On P.IDLOJA=V.IDLOJA And P.IDVENDA=V.IDVENDA" & vbNewLine
         Sql = Sql & " Where P.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Month(V.DTVENDA)=" & SqlNum(Month(pDia)) & vbNewLine
         Sql = Sql & " And Year(V.DTVENDA)=" & SqlNum(Year(pDia)) & vbNewLine
         Sql = Sql & " And isNull(V.IDFATURA,0)=0"
         Sql = Sql & " And isNull(V.FLGCANCELADA,0)=0"
         Sql = Sql & " And V.SITVENDA=10"
         Sql = Sql & " Group By P.IDFORMAPGTO"
         
         If .Abretabela(Sql) Then
            While Not .RsAux.EOF
               If .RsAux("IDFORMAPGTO") = 1 Then '* Dinheiro
                  sVLDINMES = "R$ " & ValBr(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 2 Then '* Débito
                  sVLDEBMES = "R$ " & ValBr(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 3 Then '* Crédito
                  sVLCREMES = "R$ " & ValBr(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 4 Then '* Cheque
                  sVLCHEMES = "R$ " & ValBr(.RsAux("Valor"))
               End If
               sVLTMES2 = xVal(sVLTMES2) + xVal(.RsAux("Valor"))
               .RsAux.MoveNext
            Wend
            sVLTMES2 = "R$ " & ValBr(sVLTMES2)
         End If
         
         '************
         '* Média de Sessões por dia Trabalhado
         Sql = " Select  Count(Distinct day(DTVENDA)) [QTD]" & vbNewLine
         Sql = Sql & " From CITENSVENDA I " & vbNewLine
         Sql = Sql & " Join CVENDA V On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA" & vbNewLine
         Sql = Sql & " Join SPRODUTO P On I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD AND P.EVENDA=1" & vbNewLine
         Sql = Sql & " Where I.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         Sql = Sql & " And Month(V.DTVENDA) = " & SqlNum(Month(pDia)) & vbNewLine
         Sql = Sql & " And Year(V.DTVENDA) = " & SqlNum(Year(pDia)) & vbNewLine
         Sql = Sql & " And P.ESERVICO=1"
         If .Abretabela(Sql) Then
            If xVal(.RsAux("QTD")) > 0 Then
               sDIATRABMES = xVal(.RsAux("QTD"))
               sMEDMES = ValBr(xVal(sQTDSMES) / xVal(sDIATRABMES), 2)
            End If
         End If
                  
         '************
         '* Previsão do Mês
         Dim nAreaPrevProd As Integer
         Dim nAreaRealProd  As Integer
         Dim nAreaRealVend As Integer
         
         Sql = " Select Count(*) [PREV]" & vbNewLine
         Sql = Sql & " From OEVENTOAGENDA E" & vbNewLine
         Sql = Sql & " Join OSERVICOEVT S On E.IDLOJA=S.IDLOJA and E.IDEVENTO=S.IDEVENTO" & vbNewLine
         Sql = Sql & " Join OCLIENTE C On E.IDLOJA=C.IDLOJA And E.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0" & vbNewLine
         Sql = Sql & " Where E.STARTDATETIME >= Getdate()" & vbNewLine
         Sql = Sql & " And Month(E.STARTDATETIME)=" & Month(pDia) & vbNewLine
         Sql = Sql & " And Year(E.STARTDATETIME)=" & Year(pDia) & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO <>1" & vbNewLine
         Sql = Sql & " And E.FLGCANCELADO<>1" & vbNewLine
         Sql = Sql & " And E.IDLOJA=" & mvarIDLOJA
         If .Abretabela(Sql) Then
            nAreaPrevProd = xVal(.RsAux("PREV") & "")
         End If
         Sql = "Select Count(*) [PROD]" & vbNewLine
         Sql = Sql & " From OATENDIMENTO A " & vbNewLine
         Sql = Sql & " Join OSESSAO S On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Join OCLIENTE C On C.IDLOJA=A.IDLOJA And C.IDCLIENTE=A.IDCLIENTE And C.ISENTO=0" & vbNewLine
         Sql = Sql & " Where S.IDTPSERVICO <>1" & vbNewLine
         Sql = Sql & " And Month(A.DTATEND) =" & Month(pDia) & vbNewLine
         Sql = Sql & " And Year(A.DTATEND) =" & Year(pDia) & vbNewLine
         Sql = Sql & " And A.IDLOJA=" & mvarIDLOJA
         If .Abretabela(Sql) Then
            nAreaRealProd = xVal(.RsAux("PROD") & "")
         End If
         Sql = "Select Sum(I.QTDVENDA*I.VLUNIT)/AVG(I.VLUNIT)- Sum(V.VLDESC)/AVG(I.VLUNIT) as [Qtd]" & vbNewLine
         Sql = Sql & " From CITENSVENDA I" & vbNewLine
         Sql = Sql & " Join CVENDA V On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA" & vbNewLine
         Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And C.IDCLIENTE=V.IDCLIENTE And C.ISENTO=0" & vbNewLine
         Sql = Sql & " Join SPRODUTO P On P.IDLOJA=I.IDLOJA And P.IDPROD=I.IDPROD" & vbNewLine
         Sql = Sql & " Where Month(V.DTVENDA)=" & Month(pDia) & vbNewLine
         Sql = Sql & " And Year(V.DTVENDA)=" & Year(pDia) & vbNewLine
         Sql = Sql & " And V.IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And P.ESERVICO=1" & vbNewLine
         Sql = Sql & " Group by Year(V.DTVENDA), Month(V.DTVENDA)" & vbNewLine
         Sql = Sql & " Order by Year(V.DTVENDA), Month(V.DTVENDA)"
         If .Abretabela(Sql) Then
            nAreaRealVend = xVal(.RsAux("Qtd") & "")
         End If
         sPREVVENDAMES = ValBr(nAreaPrevProd + nAreaRealVend, 0)
         sPREVPRODMES = ValBr(nAreaPrevProd + nAreaRealProd, 0)
         
         '************
         '* Valores Anuais
         Sql = "Select Year(V.DTVENDA) As [Ano], Month(V.DTVENDA) As [Mes], SUM(I.QTDVENDA) As [QTD]"
         Sql = Sql & " From CITENSVENDA I"
         Sql = Sql & " Join CVENDA V On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA"
         Sql = Sql & " Join SPRODUTO P On I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD AND P.EVENDA=1"
         Sql = Sql & " Where I.IDLOJA=" & SqlNum(mvarIDLOJA)
         Sql = Sql & " And Year(V.DTVENDA) >=" & SqlNum(sANO2)
         Sql = Sql & " And P.ESERVICO =1"
         Sql = Sql & " Group By Year(V.DTVENDA), Month(V.DTVENDA)"
         Sql = Sql & " Order By Year(V.DTVENDA), Month(V.DTVENDA)"
         If .Abretabela(Sql) Then
            While Not .RsAux.EOF
               nAno = xVal(sANO0) - .RsAux("Ano")
               nMes = .RsAux("Mes")
               sAux = "@@QTDS" & StrZero(nAno, 2) & StrZero(nMes, 2)
               sHtml = Replace(sHtml, sAux, ValBr(.RsAux("QTD"), 0))
               
               If nAno = 0 Then sQTDT00 = xVal(sQTDT00) + .RsAux("QTD")
               If nAno = 1 Then sQTDT01 = xVal(sQTDT01) + .RsAux("QTD")
               If nAno = 2 Then sQTDT02 = xVal(sQTDT02) + .RsAux("QTD")
               
               .RsAux.MoveNext
            Wend
            If sQTDT00 <> "--" Then sQTDT00 = ValBr(sQTDT00, 0)
            If sQTDT01 <> "--" Then sQTDT01 = ValBr(sQTDT01, 0)
            If sQTDT02 <> "--" Then sQTDT02 = ValBr(sQTDT02, 0)
         End If
         For nAno = 0 To 2
            For nMes = 1 To 12
               sAux = "@@QTDS" & StrZero(nAno, 2) & StrZero(nMes, 2)
               sHtml = Replace(sHtml, sAux, "--")
            Next
         Next
         
         '**************************
         '***** POPULAR HTML *******
         '**************************
         sHtml = Replace(sHtml, "@@NMLOJA", sNMLoja)
         
         sHtml = Replace(sHtml, "@@LINHAPDIA", sLinhaPDia)
         sHtml = Replace(sHtml, "@@LINHAPMES", sLinhaPMes)
         sHtml = Replace(sHtml, "@@VLTPDIA", sVLTPDIA)
         sHtml = Replace(sHtml, "@@VLTPMES", sVLTPMES)
         
         sHtml = Replace(sHtml, "@@ANO0", sANO0)
         sHtml = Replace(sHtml, "@@ANO1", sANO1)
         sHtml = Replace(sHtml, "@@ANO2", sANO2)
         
         sHtml = Replace(sHtml, "@@QTDSDIA", sQTDSDIA)
         sHtml = Replace(sHtml, "@@QTDPDIA", sQTDPDIA)
         sHtml = Replace(sHtml, "@@QTDTDIA", sQTDTDIA)
         sHtml = Replace(sHtml, "@@VLSDIA", sVLSDIA)
         sHtml = Replace(sHtml, "@@VLPDIA", sVLPDIA)
         sHtml = Replace(sHtml, "@@VLDESCDIA", sVLDescDIA)
         sHtml = Replace(sHtml, "@@VLFATDIA", sVLFATDIA)
         sHtml = Replace(sHtml, "@@VLTDIA1", sVLTDIA1)
         sHtml = Replace(sHtml, "@@VLTDIA2", sVLTDIA2)
         sHtml = Replace(sHtml, "@@VLDINDIA", sVLDINDIA)
         sHtml = Replace(sHtml, "@@VLDEBDIA", sVLDEBDIA)
         sHtml = Replace(sHtml, "@@VLCREDIA", sVLCREDIA)
         sHtml = Replace(sHtml, "@@VLCHEDIA", sVLCHEDIA)
         
         sHtml = Replace(sHtml, "@@VLCAIXADIA", sVLCAIXADIA)
         sHtml = Replace(sHtml, "@@QTDPROD", sQTDPROD)
         sHtml = Replace(sHtml, "@@PREVPRODMES", sPREVPRODMES)
         sHtml = Replace(sHtml, "@@PREVVENDAMES", sPREVVENDAMES)
         
         sHtml = Replace(sHtml, "@@QTDSMES", sQTDSMES)
         sHtml = Replace(sHtml, "@@QTDPMES", sQTDPMES)
         sHtml = Replace(sHtml, "@@QTDTMES", sQTDTMES)
         sHtml = Replace(sHtml, "@@VLSMES", sVLSMES)
         sHtml = Replace(sHtml, "@@VLPMES", sVLPMES)
         sHtml = Replace(sHtml, "@@VLDESCMES", sVLDescMes)
         sHtml = Replace(sHtml, "@@VLFATMES", sVLFATMes)
         sHtml = Replace(sHtml, "@@VLTMES1", sVLTMES1)
         sHtml = Replace(sHtml, "@@VLTMES2", sVLTMES2)
         sHtml = Replace(sHtml, "@@VLDINMES", sVLDINMES)
         sHtml = Replace(sHtml, "@@VLDEBMES", sVLDEBMES)
         sHtml = Replace(sHtml, "@@VLCREMES", sVLCREMES)
         sHtml = Replace(sHtml, "@@VLCHEMES", sVLCHEMES)
         sHtml = Replace(sHtml, "@@DIATRABMES", sDIATRABMES)
         sHtml = Replace(sHtml, "@@MEDMES", sMEDMES)
         
         sHtml = Replace(sHtml, "@@QTDT00", sQTDT00)
         sHtml = Replace(sHtml, "@@QTDT01", sQTDT01)
         sHtml = Replace(sHtml, "@@QTDT02", sQTDT02)
         
         sHtml = Replace(sHtml, "@@DIA", sDIA)
         sHtml = Replace(sHtml, "@@MES", sMES)
      
      End With
   End If
   
   '********
   '* Escrever Arquivo
   sPathFile = mvarSys.PathTmp & "Htm\Fechamento_" & Format(pDia, "yymmdd") & ".htm"
   Call CriarDiretorio(mvarSys.PathTmp & "Htm")
   Call ExcluirArquivo(sPathFile)
   Open sPathFile For Output As #1
   Print #1, sHtml
   Close #1
   
   GetHtmleMailDiario = sHtml
End Function
Public Function GetHtmleMailFaturamento(pDia As String) As String
   Dim sFile As String
   Dim sPath As String
   Dim sPathFile As String
   Dim sHtml As String
   Dim Sql As String
   Dim i As Integer
   Dim nAux As Integer
   Dim bExiste As Boolean
   
   sFile = "Faturamento3R.htm"
   sPath = Environ("TEMP") & "\" & mvarSys.CODSIS & "\ "
   sPathFile = sPath & sFile
   
   Call ExtractResData("FATURAMENTO", "HTML", sPathFile)
   If ExisteArquivo(sPathFile) Then
      sHtml = ReadTextFile(sPathFile)
      Call ExcluirArquivo(sPathFile, False)
   End If
   If sHtml = "" Then
      Call CopiarArquivo(App.Path & "\" & sFile, sPathFile)
      sHtml = ReadTextFile(sPathFile)
   End If
   
   If Trim(sHtml) <> "" Then
      Dim sNMLoja As String, sRAZAO As String, sCNPJ As String, sMESREF As String
      Dim sQTDSERV As String, sVLSERV As String
      Dim sQTDPROD As String, sVLPROD As String
      Dim sVLTOTAL As String, sDETALHE As String
      Dim sVLDESC As String, sQTDREAL As String
      Dim sQTDPRODReal As String
             
      sMESREF = " " & MonthName(Month(pDia))
      sMESREF = sMESREF & "/" & Format(pDia, "YY")
      sMESREF = StrConv(sMESREF, vbProperCase)
           
      With mvarSys.xDb
         Sql = "Select NOME, CNPJ, RAZAO, RZABREV "
         Sql = Sql & " From OLOJA"
         Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA
         If .Abretabela(Sql) Then
            sNMLoja = .RsAux("NOME")
            sRAZAO = .RsAux("RAZAO")
            sCNPJ = UnFormat(.RsAux("CNPJ"))
            sCNPJ = Format(Left(sCNPJ, 8), "0#,###,###") & "/" & Format(Right(sCNPJ, 6), "0###-##")
         End If
         
         '************
         '* Valor Serviços Faturados
         Sql = " Select IDLOJA, NUMNOTA, SERIE, IDRECIBO, DTEMISSAO, VLTOTAL" & vbNewLine
         Sql = Sql & " From FNOTAFISCAL" & vbNewLine
         Sql = Sql & " Where IDLOJA =" & mvarSys.IDLOJA & vbNewLine
         Sql = Sql & " And Month(DTEMISSAO)=" & Month(pDia) & vbNewLine
         Sql = Sql & " And Year(DTEMISSAO)=" & Year(pDia) & vbNewLine
         Sql = Sql & " And SERIE=1" & vbNewLine
         Sql = Sql & " And FLGCANCELADA<>1" & vbNewLine
         sDETALHE = ""
         sDETALHE = sDETALHE & "Nota Fiscal Serie Recibo &nbsp;Data Emissão &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Valor" & "<br>"
         sDETALHE = sDETALHE & String(50, "=") & "<br>"
         bExiste = False
         If .Abretabela(Sql) Then
            bExiste = True
            While Not .RsAux.EOF
               sDETALHE = sDETALHE & StrZero(.RsAux("NUMNOTA"), 6) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
               sDETALHE = sDETALHE & .RsAux("SERIE") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
               sDETALHE = sDETALHE & StrZero(.RsAux("IDRECIBO"), 6) & "&nbsp;"
               sDETALHE = sDETALHE & "&nbsp;" & .RsAux("DTEMISSAO") & "&nbsp;&nbsp;&nbsp;"
               nAux = Len(ValBr(.RsAux("VLTOTAL")))
               sDETALHE = sDETALHE & "R$"
               For i = 1 To 9 - nAux
                  sDETALHE = sDETALHE & "&nbsp;"
               Next
               sDETALHE = sDETALHE & ValBr(.RsAux("VLTOTAL")) & "&nbsp;"
               sDETALHE = sDETALHE & "<br>" & vbNewLine
               .RsAux.MoveNext
            Wend
         End If
         Sql = " Select IsNull(Sum(VLTOTAL),0) As [Total]" & vbNewLine
         Sql = Sql & " From FNOTAFISCAL" & vbNewLine
         Sql = Sql & " Where IDLOJA =" & mvarSys.IDLOJA & vbNewLine
         Sql = Sql & " And Month(DTEMISSAO)=" & Month(pDia) & vbNewLine
         Sql = Sql & " And Year(DTEMISSAO)=" & Year(pDia) & vbNewLine
         Sql = Sql & " And SERIE=1" & vbNewLine
         Sql = Sql & " And FLGCANCELADA<>1" & vbNewLine
         If bExiste Then
            Call .Abretabela(Sql)
            sVLSERV = ValBr(.RsAux("TOTAL"), 2)
            
            '************
            '* Quantidade Serviços Faturados
            Sql = " Select P.ESERVICO, Sum(I.QTDVENDA) As [QTD], AVG(I.VLUNIT) As [MEDIA]" & vbNewLine
            Sql = Sql & " From FNOTAFISCAL F" & vbNewLine
            Sql = Sql & " JOIN CVENDA V ON F.IDLOJA=V.IDLOJA AND F.IDVENDA=V.IDVENDA" & vbNewLine
            Sql = Sql & " JOIN CITENSVENDA I ON V.IDLOJA=I.IDLOJA AND V.IDVENDA=I.IDVENDA" & vbNewLine
            Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA AND I.IDPROD=P.IDPROD" & vbNewLine
            Sql = Sql & " Where F.IDLOJA =" & mvarSys.IDLOJA & vbNewLine
            Sql = Sql & " And Month(F.DTEMISSAO)=" & Month(pDia) & vbNewLine
            Sql = Sql & " And Year(F.DTEMISSAO)=" & Year(pDia) & vbNewLine
            Sql = Sql & " And F.SERIE=1" & vbNewLine
            Sql = Sql & " And P.ESERVICO=1" & vbNewLine
            Sql = Sql & " And F.FLGCANCELADA<>1" & vbNewLine
            Sql = Sql & " AND V.FLGCANCELADA<>1" & vbNewLine
            Sql = Sql & " GROUP BY P.ESERVICO"
            If .Abretabela(Sql) Then
               sQTDSERV = ValBr(.RsAux("QTD"), 2)
               sQTDSERV = Int(xVal(sVLSERV) / xVal(.RsAux("MEDIA")))
            End If
         Else
            Sql = Sql & " Select IsNull(Sum(V.VLVENDA-V.VLDESC),0) As [Total]" & vbNewLine
            Sql = Sql & " From CVENDA V" & vbNewLine
            Sql = Sql & " Where V.IDLOJA =" & mvarSys.IDLOJA & vbNewLine
            Sql = Sql & " And Month(V.DTVENDA)=" & Month(pDia) & vbNewLine
            Sql = Sql & " And Year(V.DTVENDA)=" & Year(pDia) & vbNewLine
            Sql = Sql & " AND V.FLGCANCELADA<>1" & vbNewLine
            If .Abretabela(Sql) Then
               sVLSERV = ValBr(.RsAux("Total"), 2)
               '************
               '* Valor Descontos
               Sql = " Select Sum(V.VLDESC) As [Total]" & vbNewLine
               Sql = Sql & " From CVENDA V" & vbNewLine
               Sql = Sql & " Where V.IDLOJA =" & mvarSys.IDLOJA & vbNewLine
               Sql = Sql & " And Month(V.DTVENDA)=" & Month(pDia) & vbNewLine
               Sql = Sql & " And Year(V.DTVENDA)=" & Year(pDia) & vbNewLine
               Sql = Sql & " AND V.FLGCANCELADA<>1" & vbNewLine
               If .Abretabela(Sql) Then
                  sVLDESC = ValBr(.RsAux("Total"), 2)
               End If
               
               '************
               '* Quantidade Serviços Vendidos
               Sql = " Select P.ESERVICO, Sum(I.QTDVENDA*I.VLUNIT) [TOTAL]" & vbNewLine
               Sql = Sql & ", Sum(I.QTDVENDA) As [QTD], AVG(I.VLUNIT) As [MEDIA]" & vbNewLine
               Sql = Sql & " From CVENDA V " & vbNewLine
               Sql = Sql & " JOIN CITENSVENDA I ON V.IDLOJA=I.IDLOJA AND V.IDVENDA=I.IDVENDA" & vbNewLine
               Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA AND I.IDPROD=P.IDPROD" & vbNewLine
               Sql = Sql & " Where V.IDLOJA =" & mvarSys.IDLOJA & vbNewLine
               Sql = Sql & " And Month(V.DTVENDA)=" & Month(pDia) & vbNewLine
               Sql = Sql & " And Year(V.DTVENDA)=" & Year(pDia) & vbNewLine
               Sql = Sql & " And P.ESERVICO=1" & vbNewLine
               Sql = Sql & " AND V.FLGCANCELADA<>1" & vbNewLine
               Sql = Sql & " GROUP BY P.ESERVICO"
               If .Abretabela(Sql) Then
                  sVLSERV = ValBr(xVal(.RsAux("TOTAL") & "") - xVal(sVLDESC), 2)
                  sQTDREAL = ValBr(.RsAux("QTD"), 2)
                  sQTDSERV = Int(xVal(sVLSERV) / xVal(.RsAux("MEDIA")))
               End If
            End If
         End If
         
         '************
         '* Valor Produtos Faturados
         Sql = " Select IDLOJA, NUMNOTA, SERIE" & vbNewLine
         Sql = Sql & " From FNOTAFISCAL" & vbNewLine
         Sql = Sql & " Where IDLOJA =" & mvarSys.IDLOJA & vbNewLine
         Sql = Sql & " And Month(DTEMISSAO)=" & Month(pDia) & vbNewLine
         Sql = Sql & " And Year(DTEMISSAO)=" & Year(pDia) & vbNewLine
         Sql = Sql & " And SERIE=2" & vbNewLine
         Sql = Sql & " And FLGCANCELADA<>1" & vbNewLine
         bExiste = False
         If .Abretabela(Sql) Then
            bExiste = True
            While Not .RsAux.EOF
               sDETALHE = sDETALHE & StrZero(.RsAux("NUMNOTA"), 6) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
               sDETALHE = sDETALHE & .RsAux("SERIE") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
               sDETALHE = sDETALHE & StrZero(.RsAux("IDRECIBO"), 6) & "&nbsp;"
               sDETALHE = sDETALHE & "&nbsp" & .RsAux("DTEMISSAO") & "&nbsp;&nbsp;&nbsp;"
               nAux = Len(ValBr(.RsAux("VLTOTAL")))
               sDETALHE = sDETALHE & "R$"
               For i = 1 To 9 - nAux
                  sDETALHE = sDETALHE & "&nbsp;"
               Next
               sDETALHE = sDETALHE & ValBr(.RsAux("VLTOTAL")) & "&nbsp;"
               sDETALHE = sDETALHE & "<br>" & vbNewLine
               .RsAux.MoveNext
            Wend
         End If
         
         Sql = " Select IsNull(Sum(VLTOTAL),0) As [Total]" & vbNewLine
         Sql = Sql & " From FNOTAFISCAL" & vbNewLine
         Sql = Sql & " Where IDLOJA =" & mvarSys.IDLOJA & vbNewLine
         Sql = Sql & " And Month(DTEMISSAO)=" & Month(pDia) & vbNewLine
         Sql = Sql & " And Year(DTEMISSAO)=" & Year(pDia) & vbNewLine
         Sql = Sql & " And SERIE=2" & vbNewLine
         Sql = Sql & " And FLGCANCELADA<>1" & vbNewLine
         If bExiste Then
            Call .Abretabela(Sql)
            sVLSERV = ValBr(.RsAux("TOTAL"), 2)
            
            '************
            '* Quantidade Produtos Faturados
            Sql = " Select P.ESERVICO, Sum(I.QTDVENDA) As [QTD], AVG(I.VLUNIT) As [MEDIA]"
            Sql = Sql & " From FNOTAFISCAL F"
            Sql = Sql & " JOIN CVENDA V ON F.IDLOJA=V.IDLOJA AND F.IDVENDA=V.IDVENDA"
            Sql = Sql & " JOIN CITENSVENDA I ON V.IDLOJA=I.IDLOJA AND V.IDVENDA=I.IDVENDA"
            Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA AND I.IDPROD=P.IDPROD"
            Sql = Sql & " Where F.IDLOJA =" & mvarSys.IDLOJA
            Sql = Sql & " And Month(F.DTEMISSAO)=" & Month(pDia)
            Sql = Sql & " And Year(F.DTEMISSAO)=" & Year(pDia)
            Sql = Sql & " And F.SERIE=2"
            Sql = Sql & " And P.ESERVICO=0"
            Sql = Sql & " And F.FLGCANCELADA<>1"
            Sql = Sql & " AND V.FLGCANCELADA<>1"
            Sql = Sql & " GROUP BY P.ESERVICO"
            If .Abretabela(Sql) Then
                sQTDPROD = ValBr(.RsAux("QTD"), 2)
            End If
         Else
            '************
            '* Quantidade e Valor Produtos Vendidos
            Sql = " Select P.ESERVICO, Sum(I.QTDVENDA*I.VLUNIT) [TOTAL]" & vbNewLine
            Sql = Sql & ", Sum(I.QTDVENDA) As [QTD], AVG(I.VLUNIT) As [MEDIA]" & vbNewLine
            Sql = Sql & " From CVENDA V " & vbNewLine
            Sql = Sql & " JOIN CITENSVENDA I ON V.IDLOJA=I.IDLOJA AND V.IDVENDA=I.IDVENDA" & vbNewLine
            Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA AND I.IDPROD=P.IDPROD" & vbNewLine
            Sql = Sql & " Where V.IDLOJA =" & mvarSys.IDLOJA & vbNewLine
            Sql = Sql & " And Month(V.DTVENDA)=" & Month(pDia) & vbNewLine
            Sql = Sql & " And Year(V.DTVENDA)=" & Year(pDia) & vbNewLine
            Sql = Sql & " And P.ESERVICO=0" & vbNewLine
            Sql = Sql & " AND V.FLGCANCELADA<>1" & vbNewLine
            Sql = Sql & " GROUP BY P.ESERVICO" & vbNewLine
            If .Abretabela(Sql) Then
               sVLPROD = ValBr(.RsAux("TOTAL"), 2)
               sQTDPRODReal = ValBr(.RsAux("QTD"), 2)
               sQTDPROD = Int(xVal(sVLPROD) / xVal(.RsAux("MEDIA")))
            End If
         End If
      End With
      sVLTOTAL = "R$ " & ValBr(xVal(sVLSERV) + xVal(sVLPROD))
      sVLSERV = "R$ " & sVLSERV
      sVLPROD = "R$ " & sVLPROD
      
      '**************************
      '***** POPULAR HTML *******
      '**************************
      sHtml = Replace(sHtml, "@@NMLOJA", sNMLoja)
      sHtml = Replace(sHtml, "@@RAZAO", sRAZAO)
      sHtml = Replace(sHtml, "@@CNPJ", sCNPJ)
      sHtml = Replace(sHtml, "@@MESREF", sMESREF)

      sHtml = Replace(sHtml, "@@QTDSERV", sQTDSERV)
      sHtml = Replace(sHtml, "@@VLSERV", sVLSERV)
      sHtml = Replace(sHtml, "@@QTDPROD", sQTDPROD)
      sHtml = Replace(sHtml, "@@VLPROD", sVLPROD)
      sHtml = Replace(sHtml, "@@VLTOTAL", sVLTOTAL)
      sHtml = Replace(sHtml, "@@DETALHE", sDETALHE)
      
   End If
   
   GetHtmleMailFaturamento = sHtml
End Function
Private Function ExisteRPS(Optional pBaixar As Boolean = False, Optional pDTFechamento As String) As String
   Dim Sql As String
   Dim nDia As Integer
   Dim sDTFIM As String
   Dim sHoje As String
   Dim sAmanha As String
   Dim sDtAux As String
   Dim QtdDia As Integer
   Dim bOk As Boolean
   
   QtdDia = 15
   sHoje = mvarSys.xDb.SysDate()
   sAmanha = DateAdd("d", CDate(mvarSys.xDb.SysDate()), 1)
   
   nDia = Day(sAmanha)
   sDtAux = Format(sAmanha, "/mm/yyyy")
   If QtdDia = 10 Then
      If nDia <= QtdDia Then
         sDTFIM = DateAdd("d", CDate("01" & sDtAux), -1)
      ElseIf nDia > QtdDia And nDia <= (2 * QtdDia) Then
         sDTFIM = Trim(CStr(QtdDia)) & sDtAux
      Else
         sDTFIM = Trim(CStr(2 * QtdDia)) & sDtAux
      End If
   Else
      If nDia <= QtdDia Then
         sDTFIM = DateAdd("d", CDate("01" & sDtAux), -1)
      ElseIf nDia > QtdDia Then
         sDTFIM = Trim(CStr(QtdDia)) & sDtAux
      End If
   End If
   sDTFIM = sDTFIM & " 23:59:59"
  
   If pBaixar Then
      Sql = "Select Min(R.DTEMISSAO) as [DTINI]"
      Sql = Sql & " From FRECIBO R"
      Sql = Sql & " Left Join CVENDA V On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA AND V.FLGCANCELADA<>1"
      Sql = Sql & " Left Join FNOTAFISCAL F On V.IDLOJA=F.IDLOJA And V.IDVENDA=F.IDVENDA"
      Sql = Sql & " Where R.IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And R.TIPO=0"
      Sql = Sql & " And (R.IDLOTE = 0 Or R.IDLOTE is Null)"
      Sql = Sql & " And (F.NUMNOTA= 0 Or F.NUMNOTA is Null)"
      Sql = Sql & " And R.DTEMISSAO <=" & SqlDate(sDTFIM)
      If mvarSys.xDb.Abretabela(Sql) Then
         If DateDiff("D", mvarSys.xDb.RsAux("DTINI"), sHoje) < 10 Then
            If mvarSys.xDb.SysDate <= CDate(sDTFIM) Then
               ExisteRPS = ""
               Exit Function
            End If
         End If
      End If
   End If
   
   bOk = nDia Mod QtdDia <= 6
   bOk = bOk Or pBaixar
   If pDTFechamento <> "" Then
      bOk = bOk Or Day(pDTFechamento) Mod QtdDia <= 6
   End If
   If bOk Then
      Sql = "Select Count(*) as [QTD]"
      Sql = Sql & " From FRECIBO R"
      Sql = Sql & " Left Join CVENDA V On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA AND V.FLGCANCELADA<>1"
      Sql = Sql & " Left Join FNOTAFISCAL F On V.IDLOJA=F.IDLOJA And V.IDVENDA=F.IDVENDA"
      Sql = Sql & " Where R.IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And R.TIPO=0"
      Sql = Sql & " And (R.IDLOTE = 0 Or R.IDLOTE is Null)"
      Sql = Sql & " And (F.NUMNOTA= 0 Or F.NUMNOTA is Null)"
      Sql = Sql & " And R.DTEMISSAO <=" & SqlDate(sDTFIM)
      If mvarSys.xDb.Abretabela(Sql) Then
         If (xVal(mvarSys.xDb.RsAux("QTD") & "") > 0) Then
            ExisteRPS = sDTFIM
         End If
      End If
   End If
End Function
Public Sub F_BaixarNF_FTP(Optional pBaixar As Boolean = False)
   Dim nMouse  As Integer
   Dim bIsWeb  As Boolean
   Dim oFtp    As Object
   Dim NgFin   As Object
   Dim sArq    As String
   Dim RemPath As String
   Dim RPath   As String
   Dim sUNIDADE As String
   Dim LocPath   As String
   Dim bProducao As Boolean
   Dim bConectado As Boolean
   Dim bBaixar    As Boolean
   Dim sDir       As String
   Dim sDT        As String
   
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   On Error GoTo Saida
      
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUCAO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUCÃO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUÇAO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUÇÃO")
   If Not bProducao Then GoTo Saida
   
   If Trim(ExisteRPS(pBaixar)) <> "" Then
      With mvarSys.MDI
         .ProgressBar2.Visible = True
         .CommandBars.StatusBar.Pane(3).Handle = .ProgressBar2.hwnd
      End With
      
      Set NgFin = CriarObjeto("Financ3R.NG_Financeiro")
      Set NgFin.Sys = mvarSys
      sArq = NgFin.ArquivoRPS
      If Trim(sArq) <> "" Then
         sArq = Mid(sArq, 1, Len(sArq) - 3) & "xml"
         
         bIsWeb = IsWebConnected
         If Err <> 0 Then bIsWeb = True
         bConectado = False
         '* Conectar FTP
         If bIsWeb Then
            Set oFtp = ConectarFTP()
            If Not oFtp Is Nothing Then
               bConectado = oFtp.Conectado
            End If
         End If
         
         LocPath = mvarSys.PathTmp & "RPS\"
         Call CriarDiretorio(LocPath)
         
         sUNIDADE = mvarSys.Propriedades("NMLOJA")
         RPath = Trim(mvarSys.GetParam("FtpBakPath", Default:="", pDescricao:="Caminho do Ftp de Backup"))
         
         If RPath = "" Then RPath = "RPS"
         RPath = IIf(Mid(RPath, 1, 1) = "/", "", "/") & RPath
         RPath = RPath & IIf(Right(RPath, 1) = "/", "", "/")
         
         RemPath = ""
         If RPath = "/RPS/" Then
            RemPath = "/" & sUNIDADE & RPath
         End If
                        
         '* Baixar Arquivo
         bBaixar = bConectado
         If ExisteArquivo(LocPath & sArq) Then
            bBaixar = (ExibirPergunta("Arquivo '" & sArq & "' já existe em seu computador." & vbNewLine & vbNewLine & "Deseja baixar novamente?", "Nota Fiscal", False) = vbYes)
            If bBaixar Then
               Call CopiarArquivo(LocPath & sArq, LocPath & sArq & "bak")
            End If
         End If
         If bBaixar Then
            If oFtp.BaixarArquivo(RemPath, sArq, LocPath, sArq) Then
               RemPath = RemPath
               'Call NgFin.ImportarNF(LocPath & sArq, False)
            Else
               RemPath = "/" & sUNIDADE & IIf(Mid(RPath, 1, 1) = "/", "", "/") & RPath
               If oFtp.BaixarArquivo(RemPath, sArq, LocPath, sArq) Then
               '   Call NgFin.ImportarNF(LocPath & sArq, False)
               End If
            End If
         End If
         If Not ExisteArquivo(LocPath & sArq) Then
            If ExisteArquivo(LocPath & sArq & "bak") Then
               Call CopiarArquivo(LocPath & sArq & "bak", LocPath & sArq)
               Call ExcluirArquivo(LocPath & sArq & "bak")
            End If
         End If
         If ExisteArquivo(LocPath & sArq) Then
            Call NgFin.ImportarNF(LocPath & sArq, False)
         Else
            If ExisteArquivo(mvarSys.ExePath & "RPS\" & sArq) Then
               Call NgFin.ImportarNF(mvarSys.ExePath & "RPS\" & sArq, False)
            Else
               '* Verificar existencia de Arquivos XML não importados
               sDir = Dir$(LocPath & "*.xml")
               sDT = "01/01/1900"
               While sDir <> ""
                  If CDate(sDT) <= CDate(FileDateTime(LocPath & sDir)) Then
                     sArq = sDir
                     sDT = FileDateTime(LocPath & sDir)
                  End If
                  sDir = Dir$
               Wend
               If ExisteArquivo(LocPath & sArq) Then
                  Call NgFin.ImportarNF(LocPath & sArq, False)
               End If
            End If
            If bConectado Then
               oFtp.DesconectarFTP
            End If
         End If
      End If
      
      With mvarSys.MDI
         .ProgressBar2.Visible = False
         .CommandBars.StatusBar.Pane(3).Handle = .ProgressBar.hwnd
      End With
   End If
Saida:
   Screen.MousePointer = nMouse
End Sub
Private Function GerarArquivoRPS(Optional pDTINI As String, Optional pDTFIM As String, Optional pDTFechamento As String)
   Dim sFile As String
   Dim nAux As Integer
   Dim NgFin As Object
   Dim sDTINI As String
   Dim sDTFIM As String
   
   sDTINI = pDTINI
   If pDTFIM = "" Then
      sDTFIM = ExisteRPS(pDTFechamento:=pDTFechamento)
   Else
      sDTFIM = pDTFIM
   End If
   
   If Trim(sDTFIM) <> "" Then
      Set NgFin = CriarObjeto("Financ3R.NG_Financeiro")
      Set NgFin.Sys = mvarSys
      sFile = NgFin.ExportarRPS(pDTINI:=sDTINI, pDTFIM:=sDTFIM, bExibe:=False, pDTFechamento:=pDTFechamento)
      Set NgFin = Nothing
   End If
   GerarArquivoRPS = sFile
End Function
Private Function ConectarFTP() As Object
   Dim oFtp       As Object
   Dim FtpBakURL  As String
   Dim FtpBakUID  As String
   Dim FtpBakPWD  As String
   Dim SenhaNova  As String
   Dim bConectado As Boolean
'   Dim sFtpBak As String
'   Dim sFtpBakUID As String
'   Dim sFtpBakPWD As String
'   Dim mvarFileIni As String
   
   FtpBakURL = mvarSys.GetParam("FTP", Default:="ftp.classeaconsultoria.com.br", pDescricao:="Servidor de Ftp")
   If UCase(Mid(FtpBakURL, 1, 3)) = "FTP" Then
      FtpBakURL = FtpBakURL
   Else
      FtpBakURL = Decrypt2(FtpBakURL)
   End If
   FtpBakUID = mvarSys.GetParam("FtpUID", Default:="clientedpil", pDescricao:="Usuário do Ftp") 'freguesia
   FtpBakPWD = Decrypt2(mvarSys.GetParam("FtpPWD", Default:=Encrypt2(""), pDescricao:="Senha do Ftp")) 'dpil1010 tresrios10 Dpil!010
   
   'mvarFileIni = Environ("programfiles") & "\ClasseA\Admin\Dll\" & "CAVERSAO.ini"
   'If ExisteArquivo(mvarFileIni) Then
   '   sFtpBak = ReadIniFile(mvarFileIni, "Config", "Host", "")
   '   sFtpBakUID = ReadIniFile(mvarFileIni, "Config", "UID", "")
   '   sFtpBakPWD = Decrypt2(ReadIniFile(mvarFileIni, "Config", "PWD", ""))
   'End If
   'FtpBakURL = IIf(sFtpBak = "", FtpBakURL, sFtpBak)
   'FtpBakUID = IIf(sFtpBakUID = "", FtpBakUID, sFtpBakUID)
   'FtpBakPWD = IIf(sFtpBakPWD = "", FtpBakPWD, sFtpBakPWD)
   
   'SenhaNova = "Dpil10!0"
   'If FtpBakPWD <> SenhaNova Then
   '   FtpBakPWD = SenhaNova
   '   Call mvarsys.SaveParam("FtpPWD", Encrypt2(FtpBakPWD))
   'End If
   
   Set oFtp = Nothing
   Set oFtp = CriarObjeto("VersaoFTP.TL_VerifVersao")
   With oFtp
      Set .Sys = mvarSys
      If .ConectarFTP(FtpBakURL, FtpBakUID, FtpBakPWD, False) Then
         bConectado = True
      End If
   End With
   Set ConectarFTP = oFtp
End Function
Private Function EnviarSMS(pTipoMsg As eTipoSMS, Optional pIDCLIENTE As Long) As Boolean
   Dim sMsg       As String
   Dim sMobile    As String
   Dim sUserName  As String
   'Dim sUserID    As String
   'Dim sPwd       As String
   
   Dim sUrl       As String
   Dim nTipoSMS   As eTipoSMS
   Dim sWebServ   As String
   Dim sStatus    As String
      
   '*************
   '* Recuperar parâmetros
   sMsg = GetMessageSMS(nTipoSMS)
   sMobile = GetMobileSMS(nTipoSMS, False)
   sUserName = "Dpil"
   
   '*************
   '* Montar URL que configura API da empresa Webserver
   '* sWebServ = "Human"
   sWebServ = "Comtele"
   sUrl = GetURLSMS(sWebServ, sUserName, "", "", "", sMobile, sMsg)
  
  '*************
  '* Enviar SMS
   sStatus = SendSMS(sUrl)
   
   If sStatus = "True" Then
      MsgBox "Torpedo enviado com sucesso!"
   Else
      MsgBox sStatus
   End If
End Function
Public Function GetURLSMS(pWebService As String, pUserName As String, pUserID As String, pPWD As String, pURLBase As String, pMobile As String, pMsg As String) As String
   Dim sUrl       As String
   
   Select Case UCase(pWebService)
      Case "HUMAN"
         pURLBase = "http://system.human.com.br:8080/GatewayIntegration/msgSms.do?dispatch=send"
         pUserID = "3rios"
         pPWD = "9AsiQPOKH6" '"FsVhV0dX"
         
         sUrl = ""
         sUrl = sUrl + pURLBase
         sUrl = sUrl + "&account=" + pUserID
         sUrl = sUrl + "&code=" + pPWD
         sUrl = sUrl + "&msg=" + pMsg
         sUrl = sUrl + "&to=" + pMobile
   
      Case "COMTELE"
         '* http://www.comtele.com.br/sms/api/api_fuse_connection.php?fuse=get_id&user=60140&pwd=uhbg140
         '* ID = MjE0
         'pURLBase = "http://www.comtele.com.br/sms/api/api_fuse_connection.php?fuse=send_msg"
         'pUserID = "MjE0"
         pMsg = UrlEncode(pMsg)
      
         sUrl = ""
         sUrl = sUrl + pURLBase
         sUrl = sUrl + "&id=" + pUserID
         sUrl = sUrl + "&from=" + pUserName
         sUrl = sUrl + "&msg=" + pMsg
         sUrl = sUrl + "&number=" + pMobile
   End Select
      
   GetURLSMS = sUrl
End Function
Public Function GetMessageSMS(pTipo As eTipoSMS)
   Dim sMsg As String
   
   Select Case pTipo
      Case eTipoSMS.PosicaoDia
         sMsg = "V:R$1.240/23.560" & vbNewLine
         sMsg = sMsg & "A:18/416" & vbNewLine
         sMsg = sMsg & "C:04/25" & vbNewLine
         sMsg = sMsg & "T:04/103"
   
      Case eTipoSMS.LembrarSessao
         sMsg = "Viemos lembra de sua sessao em 26/11 as 08:30h" & vbNewLine
         sMsg = mvarSys.Propriedades("NMLOJA") & vbNewLine
         sMsg = sMsg & "Tel:2427-0821"
         
      Case eTipoSMS.LembrarCreme
         sMsg = "Viemos lembra de sua sessao em 26/11 as 08:30h." & vbNewLine
         sMsg = sMsg & "Não esqueça de raspar a área e usar nosso creme hidratante" & vbNewLine
         sMsg = mvarSys.Propriedades("NMLOJA") & vbNewLine
         sMsg = sMsg & "Tel:2427-0821"
   End Select
   
   GetMessageSMS = UrlEncode(sMsg)
End Function
Public Function GetMobileSMS(pTipo As eTipoSMS, Optional CodPais As Boolean)
   Dim sMobile  As String
   
   Select Case pTipo
      Case eTipoSMS.PosicaoDia
         sMobile = IIf(CodPais, "55", "") & "2178344618"
         'sMobile = sMobile + "," + IIf(CodPais, "55", "") & "2178341814"
         'sMobile = sMobile + "," + IIf(CodPais, "55", "") & "2178340909"
         'sMobile = sMobile + "," + IIf(CodPais, "55", "") & "217834"
   
      Case eTipoSMS.LembrarSessao
         sMobile = IIf(CodPais, "55", "") & "2178344618"
         
      Case eTipoSMS.LembrarCreme
         sMobile = IIf(CodPais, "55", "") & "2178344618"
   End Select
   
   GetMobileSMS = sMobile
End Function
Public Sub F_CONFIG()
   Dim xConfig As TL_FrmMDI
   
   Set xConfig = New TL_FrmMDI
   Set xConfig.Sys = mvarSys
   xConfig.Show
   Set xConfig = Nothing
   
   'Set mvarPane = New TL_PaneParam
   'Set mvarPane.Sys = mvarSys
   'mvarPane.Show
End Sub

