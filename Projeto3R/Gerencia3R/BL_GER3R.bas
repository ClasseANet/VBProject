Attribute VB_Name = "BL_GER3R"
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
'============================================================
Public gDebug As Boolean
Public gIDUSU As String
Public gCODSIS As String
'============================================================
Public gDBTipo As Integer
Public gALIAS  As String
Public gSERVER As String
Public gDBNAME As String
Public gDBUSER As String
Public gDBPWD  As String
'============================================================
Public gCaption1  As String
Public gCaption2  As String
Public gCaption3  As String
'============================================================
Public gLocalReg  As String
Public Const gSetupFile = "SETUP.INI"
'============================================================
Global MyMDI As TL_MDI
Global MdiMenu As Object
'============================================================
'Global Sys     As Object
'Global XDbMaua As Object
'============================================================
Global Splash  As Object
'Global DsAuto  As Object
'Global DsDsr   As Object
'Global DsMsg   As Object
'Global DSLOAD  As Object
Sub Main()
   Call MyLoadgCODSIS
   
   Set MyMDI = New TL_MDI
   MyMDI.Show
End Sub
Public Sub MyLoadgCODSIS(Optional bCODSIS As Boolean)
   Dim sPathSetup As String
   Dim sConn As String
      
   gDBTipo = -1
   gCODSIS = "P3R"
   If bCODSIS Then Exit Sub
   
   gCODSIS = "P3R"
   gLocalReg = App.Path & "\P3R.reg"
   If Not ExisteArquivo(gLocalReg) Then
      gLocalReg = Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\P3R.reg"
   End If
   
   
   'Call WriteIniFile(gLocalReg, "Conections", "Last", "1")
   
   If gLocalReg <> "" Then
      sConn = "Conection " & ReadIniFile(gLocalReg, "Conections", "Last", "0")
      gALIAS = ReadIniFile(gLocalReg, sConn, "Alias", "")
      gDBTipo = ReadIniFile(gLocalReg, sConn, "dbTipo", "-1")
      gSERVER = ReadIniFile(gLocalReg, sConn, "Server", "")
      gDBNAME = ReadIniFile(gLocalReg, sConn, "dbName", "")
      gDBUSER = ReadIniFile(gLocalReg, sConn, "UID", "")
      gDBPWD = Decrypt2(ReadIniFile(gLocalReg, sConn, "Pwd", ""))
      
      sPathSetup = ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "")
      sPathSetup = ResolvePathName(sPathSetup)
   End If
   
   If sPathSetup <> "" Then
      sPathSetup = sPathSetup & gSetupFile
      If gALIAS = "" Then gSERVER = ReadIniFile(sPathSetup, "Database Format", "Alias", "")
      If gDBTipo = -1 Then gDBTipo = ReadIniFile(sPathSetup, "Database Format", "dbTipo", "-1")
      If gSERVER = "" Then gSERVER = ReadIniFile(sPathSetup, "Database Format", "Server", "")
      If gDBNAME = "" Then gDBNAME = ReadIniFile(sPathSetup, "Database Format", "dbName", "")
      If gDBUSER = "" Then
         gDBUSER = ReadIniFile(sPathSetup, "Database Format", "UID", "")
         If gDBPWD = "" Then gDBPWD = Decrypt2(ReadIniFile(sPathSetup, "Database Format", "Pwd", ""))
      End If
   End If
      
'   gIDUSU = "DIO"   '* Para não exibir Splash
   If gALIAS = "" Then gALIAS = "PRODUCAO"
   If gDBTipo = -1 Then gDBTipo = 1
   If gDBNAME = "" Then gDBNAME = "G3R"
   If gDBUSER = "" Then
      gDBUSER = "USU_VERIF"
      If gDBPWD = "" Then gDBPWD = Decrypt2("7B75787D7A776274616A7A")
   End If

   gCaption1 = "Projeto 3R"
   gCaption2 = "Módulo"
   gCaption3 = "Gerencial"
End Sub
Public Function LoadVersion() As String
   Dim sVer    As String
   LoadVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function
Public Function MyLoadPicture() As Object
   On Error GoTo TrataErro

   If ExisteArquivo(App.Path & "\" & gCODSIS & ".ico") Then
      Set MyLoadPicture = LoadPicture(App.Path & "\" & gCODSIS & ".ico")
   ElseIf ExisteArquivo(Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\P3R.ico") Then
      Set MyLoadPicture = LoadPicture(Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\P3R.ico")
   End If
Exit Function
TrataErro:
   MsgBox Err & "-" & Error
End Function
Public Sub AtualizaBD(Optional pSys As Object)
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
      
   On Error Resume Next
      
   '*****
   '* excluir Arquivo de Log
   Call ExcluirArquivo(App.Path & "\" & "ExeScr.log")
      
   '*****
   '* Localiza Script.zia
   sPathSetup = ResolvePathName(ReadIniFile(gLocalReg, "Setup", "PATHSETUP", ""))
   'sPathSql = Environ("TEMP") & "\Sql"
   sPathSql = pSys.PathTmp & "Sql"
   
   sFileZip = pSys.xDb.dbName
   If InStr(sFileZip, "_") <> 0 Then
      sFileZip = Mid(sFileZip, 1, InStr(sFileZip, "_") - 1)
   End If
   sFileZip = sFileZip & "REV.zia"
   bExiste = ExisteArquivo(sPathSetup & sFileZip)
   If Not bExiste Then
      sFileZip = sFileZip & "REV.zip"
      ExisteArquivo (sPathSetup & sFileZip)
   End If
   If bExiste Then
      If ExisteArquivo(sPathSql & "\*.*") Then
         Kill sPathSql & "\*.Sql"
      End If
   
      Call Unzip(sPathSetup, sFileZip, sPathSql & "\", False)
      'Call ExcluirArquivo(sPathSetup & sFileZip)
      
      '*****
      '* Verifica Versão do Banco
      Sql = "Select IDBD, DSCBD, VSBD, ATUBD, DTATU"
      Sql = Sql & ", ARQATU "
      Sql = Sql & " From VERSAOBD"
      If pSys.xDb.AbreTabela(Sql, MyRs) Then
         sArqAtu = MyRs("ARQATU") & ""
         'nAtu = MyRs("VSBD") & ""
         nAtu = MyRs("ATUBD") & ""
      End If
      Set MyRs = Nothing
            
      sArqAtu = IIf(UCase(sArqAtu) = "REV" & StrZero(nAtu, 2) & ".SQL", "Rev" & StrZero(nAtu + 1, 2) & ".sql", sArqAtu)
      nAtu = nAtu + 1
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
            Call ExecuteScript(pSys.xDb, sPathSql & "\" & sArqAtu)
         End If
         
         nAtu = nAtu + 1
         sArqAtu = IIf(sArqAtu = "", "REV00.sql", sArqAtu)
      Wend
      
      '******
      '* Verifica se o Banco é Local ou Remoto
      bAux = (pSys.xDb.Server <> pSys.xDb.ServerName("[Remote]"))
      If gIDUSU = "DIO" Then
'         If ExibirPergunta("Atualiza Menu e Pesquisas?", "Acesso Restrito", False) = vbYes Then
'            bAux = True
'         End If
      End If
      If bAux Then
         '*****
         '* Executa Menu
         sPathSql = pSys.PathTmp & "Sql"
         'If Not pSys.XDB.AbreTabela("Select Distinct ALTERSTAMP From MODULO") Then
         If pSys.xDb.Alias <> "WEB" Then
            DoEvents
            sFileSql = Dir(sPathSql & "\*InsertMenu.Sql")
            If ExisteArquivo(sPathSql & "\" & sFileSql) Then
               If UCase(Mid(pSys.ExePath, 1, 12)) = "C:\SISTEMAS\" Then sPathSql = App.Path & "\Script"
               Call ExecuteScript(pSys.xDb, sPathSql & "\" & sFileSql)
            End If
         End If
         '*****
         '* Executa Pesquisas
         sPathSql = pSys.PathTmp & "Sql"
         'If Not pSys.XDB.AbreTabela("Select Distinct ALTERSTAMP From GPESQUISA") Then
         If pSys.xDb.Alias <> "WEB" Then
            sFileSql = Dir(sPathSql & "\*Pesquisas.Sql")
            If ExisteArquivo(sPathSql & "\" & sFileSql) Then
               If UCase(Mid(pSys.ExePath, 1, 12)) = "C:\SISTEMAS\" Then sPathSql = App.Path & "\Script"
               Call ExecuteScript(pSys.xDb, sPathSql & "\" & sFileSql)
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
