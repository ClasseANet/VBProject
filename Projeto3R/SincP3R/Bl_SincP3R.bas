Attribute VB_Name = "Bl_Sinc3R"
Option Explicit
Global gLocalReg  As String
Global gSetupFile As String
Global gCaption1  As String
Global gCaption2  As String
Global gCaption3  As String
Global gDebug     As Boolean

Global MDI As FrmSincP3R
Global gLocalPath As String
Public Sub Main()
   On Error Resume Next
  
   If AppAtiva(App) Then End
   
   On Error GoTo TrataErro
   gDebug = (InStr(UCase(Command$), "DEBUG") <> 0)
   Screen.MousePointer = vbHourglass
   
   gLocalPath = Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\"

   Set MDI = New FrmSincP3R
   MDI.Show
   Exit Sub
TrataErro:
   MsgBox Err.Number & " - " & Err.Description
   End
End Sub
'*********
'* Testa se já existe uma cópia da aplicação rodando e define formato Data e número.
Public Function AppAtiva(pApp As App) As Boolean
   Dim MyLoad As Object
   Dim bAtiva As Boolean

   bAtiva = False
   Set MyLoad = CriarObjeto("DSACTIVE.DS_LOAD")
   If Not MyLoad Is Nothing Then
      MyLoad.Aplic = App
      If MyLoad.Ativa Then
         bAtiva = True
      End If
   End If
   Set MyLoad = Nothing
   AppAtiva = bAtiva
End Function
Public Function AtualizaBD(pDb As Object) As Long
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
   Dim bLocal     As Boolean
   Dim bExiste    As Boolean
      
   On Error Resume Next
      
   '*****
   '* excluir Arquivo de Log
   Call ExcluirArquivo(App.Path & "\" & "ExeScr.log")
      
   '*****
   '* Localiza Script.zia
   sPathSetup = ResolvePathName(ReadIniFile(gLocalReg, "Setup", "PATHSETUP", ""))
   'sPathSql = Environ("TEMP") & "\Sql"
   sPathSql = Mid(GetSpecialFolder(1), 1, InStr(GetSpecialFolder(1), ":") - 1) & ":\Tmp\" & gCODSIS & "\Sql"
   Call CriarDiretorio(sPathSql)
   
   sFileZip = pDb.DbName 'Sys.xdb.dbName
   If InStr(sFileZip, "_") <> 0 Then
      sFileZip = Mid(sFileZip, 1, InStr(sFileZip, "_") - 1)
   End If
   If Mid(pDb.DbName, 1, 3) = "G3R" Then
      sFileZip = "G3R"
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
      If ExisteArquivo(sPathSql & "*.*") Then
         Kill sPathSql & "*.Sql"
      End If
   
      Call Unzip(sPathSetup, sFileZip, sPathSql, False)
      'Call ExcluirArquivo(sPathSetup & sFileZip)
      
      '*****
      '* Verifica Versão do Banco
      Sql = "Select IDBD, DSCBD, VSBD, ATUBD, DTATU"
      Sql = Sql & ", ARQATU "
      Sql = Sql & " From VERSAOBD"
      If pDb.AbreTabela(Sql, MyRs) Then
         sArqAtu = MyRs("ARQATU") & ""
         'nAtu = MyRs("VSBD") & ""
         nAtu = MyRs("ATUBD") & ""
      End If
      Set MyRs = Nothing
            
      sArqAtu = IIf(UCase(sArqAtu) = "REV" & StrZero(nAtu, 2) & ".SQL", "Rev" & StrZero(nAtu + 1, 2) & ".sql", sArqAtu)
      nAtu = nAtu + 1
      While ExisteArquivo(sPathSql & sArqAtu)
         For i = 1 To Len(sArqAtu)
            bNum = IsNumeric(Mid(sArqAtu, i, 1))
            If bNum Then Exit For
         Next
         If bNum Then
            sArqAtu = Mid(sArqAtu, 1, i - 1) & StrZero(nAtu, 2) & ".sql"
         Else
            Exit Function
         End If
         
         '*****
         '* Executa Atualização
         If ExisteArquivo(sPathSql & sArqAtu) Then
            Call ExecuteScript(pDb, sPathSql & sArqAtu)
         End If
         nAtu = nAtu + 1
         sArqAtu = IIf(sArqAtu = "", "REV00.sql", sArqAtu)
      Wend
      AtualizaBD = nAtu - 1
      '******
      '* Verifica se o Banco é Local ou Remoto
      bLocal = (pDb.SERVER <> pDb.ServerName("[Remote]"))
'      If gIDUSU = "DIO" Then
'         If ExibirPergunta("Atualiza Menu e Pesquisas?", "Acesso Restrito", False) = vbYes Then
'            bAux = True
'         End If
'      End If
      If bLocal Then
         '*****
         '* Executa Menu
         sPathSql = Mid(GetSpecialFolder(1), 1, InStr(GetSpecialFolder(1), ":") - 1) & ":\Tmp\" & gCODSIS & "\Sql\"
         'If Not pDb.AbreTabela("Select Distinct ALTERSTAMP From MODULO") Then
         If pDb.Alias <> "WEB" Then
            DoEvents
            sFileSql = Dir(sPathSql & "*InsertMenu.Sql")
            If UCase(Mid(App.Path, 1, 12)) = "C:\SISTEMAS\" Then sPathSql = "C:\Sistemas\Dsr\Projeto3R\Script\"
            If ExisteArquivo(sPathSql & sFileSql) Then
               Call ExecuteScript(pDb, sPathSql & sFileSql)
            End If
         End If
         '*****
         '* Executa Pesquisas
         sPathSql = Mid(GetSpecialFolder(1), 1, InStr(GetSpecialFolder(1), ":") - 1) & ":\Tmp\" & gCODSIS & "\Sql\"
         'If Not pDb.AbreTabela("Select Distinct ALTERSTAMP From GPESQUISA") Then
         If pDb.Alias <> "WEB" Then
            sFileSql = Dir(sPathSql & "*Pesquisas.Sql")
            If UCase(Mid(App.Path, 1, 12)) = "C:\SISTEMAS\" Then sPathSql = "C:\Sistemas\Dsr\Projeto3R\Script\"
            If ExisteArquivo(sPathSql & sFileSql) Then
               Call ExecuteScript(pDb, sPathSql & sFileSql)
            End If
         End If
         sPathSql = Mid(GetSpecialFolder(1), 1, InStr(GetSpecialFolder(1), ":") - 1) & ":\Tmp\" & gCODSIS & "\Sql\"
      End If
      '*****
      '* Apagar pasta
      Kill sPathSql & "*.Sql"
      'Call ExcluirArquivo(sPathSql & "\*.*")
      'Call ExcluirDiretorio(sPathSql)
   End If
End Function

