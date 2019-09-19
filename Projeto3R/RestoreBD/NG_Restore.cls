VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_RESTORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mvarDtSinc             As Date
Dim mvarDtDelete           As Date
Dim mvarSincTab            As Boolean
Dim mvarSincDelTab         As Boolean
Dim mvarDebug              As Boolean

Private mvarFrmObj         As Object
Private mvarPause          As Boolean
Private mvarImport         As Boolean
Private mvarExport         As Boolean
Private mvarIniFile        As String

Private mvarLocalServer    As String
Private mvarLocalDbName    As String
Private mvarLocalUID       As String
Private mvarLocalPWD       As String
Private mvarRemoteServer   As String
Private mvarRemoteDbName   As String
Private mvarRemoteUID      As String
Private mvarRemotePWD      As String

Private mvarArrayNoSync    As Variant
Private mvarDelScriptTab   As String
Private mvarFieldsOnTab    As String
Private mvarLojasIn        As String
Private mvarLojasLike      As String
Private mvarWebConnected   As Boolean
Private mvarDeletedFilter  As String
Private mvarSincFilter     As String
Public Property Get SincFilter() As String
   SincFilter = mvarSincFilter
End Property
Public Property Let SincFilter(ByVal vData As String)
    mvarSincFilter = vData
End Property
Public Property Get DeletedFilter() As String
   DeletedFilter = mvarDeletedFilter
End Property
Public Property Let DeletedFilter(ByVal vData As String)
    mvarDeletedFilter = vData
End Property
Public Property Get WebConnected() As Boolean
   WebConnected = mvarWebConnected
End Property
Public Property Let WebConnected(ByVal vData As Boolean)
    mvarWebConnected = vData
End Property
Public Property Get LojasLike() As String
   LojasLike = mvarLojasLike
End Property
Public Property Let LojasLike(ByVal vData As String)
    mvarLojasLike = vData
End Property
Public Property Get LojasIn() As String
   LojasIn = mvarLojasIn
End Property
Public Property Let LojasIn(ByVal vData As String)
    mvarLojasIn = vData
End Property
Public Property Get FieldsOnTab() As String
   FieldsOnTab = mvarFieldsOnTab
End Property
Public Property Let FieldsOnTab(ByVal vData As String)
    mvarFieldsOnTab = vData
End Property
Public Property Get DelScriptTab() As String
   DelScriptTab = mvarDelScriptTab
End Property
Public Property Let DelScriptTab(ByVal vData As String)
    mvarDelScriptTab = vData
End Property
Public Property Get ArrayNoSync() As Variant
   ArrayNoSync = mvarArrayNoSync
End Property
Public Property Let ArrayNoSync(ByVal vData As Variant)
    mvarArrayNoSync = vData
End Property
Public Property Get RemotePWD() As String
   RemotePWD = mvarRemotePWD
End Property
Public Property Let RemotePWD(ByVal vData As String)
    mvarRemotePWD = vData
End Property
Public Property Get RemoteUID() As String
   RemoteUID = mvarRemoteUID
End Property
Public Property Let RemoteUID(ByVal vData As String)
    mvarRemoteUID = vData
End Property
Public Property Get RemoteDbName() As String
   RemoteDbName = mvarRemoteDbName
End Property
Public Property Let RemoteDbName(ByVal vData As String)
    mvarRemoteDbName = vData
End Property
Public Property Get RemoteServer() As String
   RemoteServer = mvarRemoteServer
End Property
Public Property Let RemoteServer(ByVal vData As String)
    mvarRemoteServer = vData
End Property
Public Property Get LocalPWD() As String
   LocalPWD = mvarLocalPWD
End Property
Public Property Let LocalPWD(ByVal vData As String)
    mvarLocalPWD = vData
End Property
Public Property Get LocalUID() As String
   LocalUID = mvarLocalUID
End Property
Public Property Let LocalUID(ByVal vData As String)
    mvarLocalUID = vData
End Property
Public Property Get LocalDbName() As String
   LocalDbName = mvarLocalDbName
End Property
Public Property Let LocalDbName(ByVal vData As String)
    mvarLocalDbName = vData
End Property
Public Property Get LocalServer() As String
   LocalServer = mvarLocalServer
End Property
Public Property Let LocalServer(ByVal vData As String)
    mvarLocalServer = vData
End Property
Public Property Get IniFile() As String
   IniFile = mvarIniFile
End Property
Public Property Let IniFile(ByVal vData As String)
    mvarIniFile = vData
End Property
Public Property Get Export() As Boolean
   Export = mvarExport
End Property
Public Property Let Export(ByVal vData As Boolean)
   mvarExport = vData
   If mvarIniFile <> "" Then
      Call WriteIniFile(mvarIniFile, "CONFIG", "MERGE", IIf(mvarExport And mvarImport, "0", IIf(mvarExport, "1", IIf(mvarImport, "1", "0"))))
   End If
End Property
Public Property Get Import() As Boolean
   Import = mvarImport
End Property
Public Property Let Import(ByVal vData As Boolean)
   mvarImport = vData
   If mvarIniFile <> "" Then
      Call WriteIniFile(mvarIniFile, "CONFIG", "MERGE", IIf(mvarExport And mvarImport, "0", IIf(mvarExport, "1", IIf(mvarImport, "1", "0"))))
   End If
End Property

Public Property Set FrmObj(ByRef vData As Object)
   Set mvarFrmObj = vData
   If Not mvarFrmObj Is Nothing Then
      If mvarImport And mvarExport Then
         mvarFrmObj.OptMerge(0).Value = True
      ElseIf mvarExport Then
         mvarFrmObj.OptMerge(1).Value = True
      ElseIf mvarImport Then
         mvarFrmObj.OptMerge(2).Value = True
      End If
   End If
End Property
Public Property Get FrmObj() As Object
    Set FrmObj = mvarFrmObj
End Property
Public Property Get Pause() As Boolean
   Pause = mvarPause
End Property
Public Property Let Pause(ByVal vData As Boolean)
    mvarPause = vData
End Property
Public Sub Run()
   Dim Sql        As String
   Dim sTabelas   As String
   Dim sTab       As String
   Dim xDbLoc     As Object
   Dim xDbRem     As Object
   Dim QtdTAbs    As Integer
   Dim i          As Integer
   Dim sDelFilter As String
   Dim bSincImp   As Boolean
   Dim bSincExp   As Boolean
                           
   On Error GoTo Saida
   
   mvarWebConnected = IsWebConnected
   If Not mvarWebConnected Then GoTo Saida
   
   Call ConectarServidor(xDbLoc, 1)
   Call ConectarServidor(xDbRem, 2)
   
   If Not (xDbLoc.Conectado And xDbRem.Conectado) Then GoTo Saida
   
   sTabelas = GetTables(xDbLoc, "C.NAME In (" & mvarFieldsOnTab & ")")
   QtdTAbs = GetTag(sTabelas, "TABS", 0)
   DoEvents
   Debug.Print "Local: " & xDbLoc.SERVER & "." & xDbLoc.dbname
   Debug.Print "Remoto: " & xDbRem.SERVER & "." & xDbRem.dbname
   mvarSincDelTab = False
   
   Call RefreshLoja(xDbRem, xDbLoc)
   Call RefreshColigada(xDbRem, xDbLoc)
   
   While True
'      For i = 1 To QtdTAbs
'         sTab = GetTag(sTabelas, "TAB" & i, "")
'         If sTab <> "DELETEDROWS" Then
'            Sql = "Update " & sTab & " Set ALTERSTAMP=1"
'            Call xDbLoc.Executa(Sql)
'         End If
'      Next
      Dim lDtSinc As Date
      If mvarDebug Then MsgBox DateDiff("s", CDate(lDtSinc), CDate(xDbRem.SysDate(3)))
      mvarDtSinc = xDbRem.SysDate(3)
      lDtSinc = mvarDtSinc
      If mvarDelScriptTab <> "" Then
         If mvarSincDelTab Then
            sDelFilter = "(" & mvarLojasLike & ") And (TIMESTAMP>" & SqlDate(mvarDtDelete) & ")"
            If mvarExport Then Call SyncDeleted(xDbLoc, xDbRem, mvarDelScriptTab, pDeletedFilter:=sDelFilter)
            If mvarImport Then Call SyncDeleted(xDbRem, xDbLoc, mvarDelScriptTab, pDeletedFilter:=sDelFilter)
         Else
            If mvarExport Then Call SyncDeleted(xDbLoc, xDbRem, mvarDelScriptTab)
            If mvarImport Then Call SyncDeleted(xDbRem, xDbLoc, mvarDelScriptTab)
         End If
         If mvarDtDelete <> "00:00" Then
            'Atualizar GSINC
            If mvarSincDelTab Then
                sDelFilter = "(" & mvarLojasLike & ") And (TIMESTAMP>" & SqlDate(mvarDtDelete) & ")"
                If mvarExport Then Call SyncDeleted(xDbLoc, xDbRem, mvarDelScriptTab, pDeletedFilter:=sDelFilter)
                If mvarImport Then Call SyncDeleted(xDbRem, xDbLoc, mvarDelScriptTab, pDeletedFilter:=sDelFilter)
             Else
                If mvarExport Then Call SyncDeleted(xDbLoc, xDbRem, mvarDelScriptTab)
                If mvarImport Then Call SyncDeleted(xDbRem, xDbLoc, mvarDelScriptTab)
             End If
         Else
            mvarDtDelete = Now()
         End If
      End If

      mvarSincTab = True
      
      If Not FrmObj Is Nothing Then
         DoEvents
         FrmObj.ProgBar.Visible = True
         FrmObj.LblProg.Visible = True
      End If
      
      For i = 1 To QtdTAbs
         While mvarPause: DoEvents: Wend
         
         sTab = GetTag(sTabelas, "TAB" & i, "")
         If Not InArray(sTab, mvarArrayNoSync) Then
            'If sTab = "RFUNCIONARIO" Then
               If mvarDelScriptTab <> "" Then
                  sDelFilter = "(" & mvarLojasLike & ") And (TIMESTAMP>" & SqlDate(mvarDtDelete) & ")"
                  If mvarExport Then Call SyncDeleted(xDbLoc, xDbRem, mvarDelScriptTab, pDeletedFilter:=sDelFilter)
                  If mvarImport Then Call SyncDeleted(xDbRem, xDbLoc, mvarDelScriptTab, pDeletedFilter:=sDelFilter)
                  
               End If
               'mvarDtDelete = xDbRem.SysDate(3)
               mvarDtSinc = mvarDtDelete
               bSincImp = True
               bSincExp = True
               If sTab = "XOEVENTOAGENDA" Then
                  FrmObj.TrayIcon.MaximizeFromTray FrmObj.hwnd
                  sTab = sTab
               End If
               
               If mvarImport Then bSincImp = SyncTable(xDbRem, xDbLoc, sTab, i)
               If bSincImp Then
                  Call AtualizaDTSINC(xDbRem, sTab, mvarDtSinc)
               End If
               
               If mvarExport Then bSincExp = SyncTable(xDbLoc, xDbRem, sTab, i)
               If bSincExp Then
                  Call AtualizaDTSINC(xDbLoc, sTab, mvarDtSinc)
               End If

               
            'End If
         End If
         If Not FrmObj Is Nothing Then
            'FrmObj.Width = 5700
            DoEvents
            FrmObj.ProgBar.Value = CInt(100 * i / QtdTAbs)
            FrmObj.LblProg.Caption = StrZero(CInt(100 * i / QtdTAbs), 2) & "%"
         End If
         DoEvents
      Next
      DoEvents
   Wend
   If Not FrmObj Is Nothing Then
      FrmObj.ProgBar.Visible = False
      FrmObj.LblProg.Visible = False
   End If
   
'   MsgBox xDbLoc.dbname
   Exit Sub
Saida:
   mvarWebConnected = IsWebConnected
   If mvarWebConnected Then Resume Next
End Sub

Function SyncDeleted(xDbLoc As Object, xDbRem As Object, sTab As String, Optional pDeletedFilter As String) As Boolean
   Dim MyTabLoc      As Object
   Dim MyTabRem      As Object
   Dim MyTab         As Object
   Dim MyTabDeleted  As Object
   Dim bRemoteDest   As Boolean
   Dim sFilter       As String
   Dim ColP          As Collection
   Dim Sql           As String
   Dim bExeDel       As Boolean
   Dim sTabela       As String
   Dim nWhere        As Integer
   Dim sWhere        As String
   Dim Queries       As Collection
   Dim bExeDel_0     As Boolean
   Dim bExeSave      As Boolean
   
   On Error GoTo Saida
   
   'bExeDel = xDbRem.Executa("sp_msforeachtable ""ALTER TABLE ? NOCHECK CONSTRAINT ALL""")
   'bExeDel = xDbRem.Executa("sp_msforeachtable ""ALTER TABLE ? DISABLE TRIGGER ALL""")
   'bExeDel = xDbRem.Executa("sp_msforeachtable ""ALTER TABLE ? WITH CHECK CONSTRAINT ALL""")
   'bExeDel = xDbRem.Executa("sp_msforeachtable ""ALTER TABLE ? ENABLE TRIGGER ALL""")
   
   If Trim(sTab) = "" Then Exit Function
   Set MyTabLoc = CriarObjeto("BANCO_3R.TB_" & sTab)
   If MyTabLoc Is Nothing Then Set MyTabLoc = CriarObjeto("BANCO.TB_" & sTab)
   If MyTabLoc Is Nothing Then Exit Function
   
   pDeletedFilter = IIf(pDeletedFilter = "", mvarDeletedFilter, pDeletedFilter)
   mvarSincDelTab = True
   bRemoteDest = LCase(xDbRem.SERVER) = LCase(xDbRem.SERVERNAME("[Remote]"))
   sFilter = pDeletedFilter
   
   If bRemoteDest Then
      sFilter = sFilter & " And (ALTERSTAMP=1)"
   
      Set MyTabDeleted = CriarObjeto("BANCO_3R.TB_" & sTab)
      If MyTabDeleted Is Nothing Then Set MyTabLoc = CriarObjeto("BANCO.TB_" & sTab)
      If Not MyTabDeleted Is Nothing Then Set MyTabDeleted.xDb = xDbRem
   End If
   
   Set MyTabLoc.xDb = xDbLoc
   bExeDel = False
   If MyTabLoc.Pesquisar(Ch_Where:=sFilter, Ch_OrderBy:="TIMESTAMP, IDDELETED") Then
      While Not MyTabLoc.Rs.EOF
         MyTabLoc.Popula
                  
         Set Queries = Nothing
         Set Queries = New Collection
         
         If bRemoteDest Then
            'MyTabDeleted.IDDELETED = 0
            MyTabDeleted.Query = MyTabLoc.Query
            MyTabDeleted.SITQUERY = 0
            MyTabDeleted.Tag = MyTabLoc.Tag
            
            Queries.Add MyTabDeleted.QryInsert
         End If
                                    
         Set ColP = Nothing
         Set ColP = GetWords(MyTabLoc.Query)
         If Not ColP Is Nothing Then sTabela = ColP(3)
         nWhere = InStr(MyTabLoc.Query, " Where")
         If nWhere > 0 And ExisteItem(xDbLoc.Tables, sTabela) Then
            sWhere = Mid(MyTabLoc.Query, nWhere + 8)
            
            Set MyTab = CriarObjeto("BANCO_3R.TB_" & sTabela)
            If MyTab Is Nothing Then Set MyTab = CriarObjeto("BANCO.TB_" & sTabela)
            If Not MyTab Is Nothing Then
               'bExeDel = xDbRem.Executa(MyTabLoc.Query)
               'bExeDel = xDbRem.Executa(Queries)
               
               Set MyTabRem = CriarObjeto("BANCO_3R.TB_" & sTabela)
               If MyTabRem Is Nothing Then Set MyTabRem = CriarObjeto("BANCO.TB_" & sTabela)
               If Not MyTabRem Is Nothing Then Set MyTabRem.xDb = xDbRem
               
               bExeDel = True
               Set MyTab.xDb = xDbLoc
               If MyTab.Pesquisar(Ch_Where:=sWhere) Then
                  While Not MyTab.Rs.EOF
                     MyTab.Popula
                     If Not MyTabRem Is Nothing Then
                        Call MyTabRem.Popula(MyTab.Rs)
                        Queries.Add MyTabRem.QrySave(pAutoId:=False, pSinc:=True)
                     End If
                     DoEvents
                     MyTab.Rs.MoveNext
                  Wend
                  bExeSave = xDbRem.Executa(Queries)
               Else
                  Queries.Add MyTabLoc.Query
               End If
            Else
               Queries.Add MyTabLoc.Query
            End If
         Else
            Queries.Add MyTabLoc.Query
         End If
         bExeDel_0 = bExeDel
         bExeDel = bExeDel And xDbRem.Executa(Queries)
         
         If bExeDel Then
            DoEvents
            If bRemoteDest Then
               MyTabLoc.ALTERSTAMP = 0
            End If
            Call AtualizaDTSINC(xDbLoc, sTab, DateAdd("S", 1, CDate(MyTabLoc.TimeStamp)))
         Else
            mvarDtDelete = IIf(mvarDtDelete <= MyTabLoc.TimeStamp And mvarDtDelete <> 0, mvarDtDelete, MyTabLoc.TimeStamp)
            Set ColP = Nothing
            Set ColP = GetWords(MyTabLoc.Query)
            If Not ColP Is Nothing Then
               sTab = ColP(3)
               sWhere = Mid(MyTabLoc.Query, InStr(UCase(MyTabLoc.Query), " WHERE"))
               Sql = "Select * From " & sTab
               Sql = Sql & sWhere
               If xDbRem.ExisteReg(Sql) Then
                  If xDbLoc.ExisteReg(Sql) Then
                     Sql = "Update " & ColP(3)
                     Sql = Sql & " Set TIMESTAMP=GetDate()"
                     Sql = Sql & " , ALTERSTAMP=1"
                     Sql = Sql & Mid(MyTabLoc.Query, InStr(UCase(MyTabLoc.Query), " WHERE"))
                     If xDbLoc.Executa(Sql) Then
                        DoEvents
                     Else
                        mvarSincDelTab = False
                     End If
                  Else
                     '* Implementa Delete Cascade
                     If DeleteCascade(xDbRem, sTab, sWhere) Then
                        If xDbRem.Executa(Queries) Then
                           bExeDel = bExeDel_0
                           If bRemoteDest Then
                              MyTabLoc.ALTERSTAMP = 0
                           End If
                           Call AtualizaDTSINC(xDbLoc, sTab, DateAdd("S", 1, CDate(MyTabLoc.TimeStamp)))
                        
                        End If
                     Else
                        mvarSincDelTab = False
                     End If
                  End If
               End If
            End If
            Set ColP = Nothing
         End If
         If Not FrmObj Is Nothing Then
            FrmObj.TrayIcon.Text = "[" & sTab & "] - " & ValBr((MyTabLoc.Rs.AbsolutePosition / MyTabLoc.Rs.RecordCount) * 100, 2) & "%"
            FrmObj.Caption = FrmObj.Tag & " - " & MyTabLoc.Rs.AbsolutePosition & " / " & MyTabLoc.Rs.RecordCount
         End If

         DoEvents
         While mvarPause: DoEvents: Wend
         mvarWebConnected = IsWebConnected
         If Not mvarWebConnected Then GoTo Saida
         
         MyTabLoc.Rs.MoveNext
      Wend
      'If mvarSincDelTab Then Call AtualizaDTSINC(xDbLoc, sTab, mvarDtSinc)
      If mvarDtDelete <> 0 Then
         Call AtualizaDTSINC(xDbLoc, sTab, DateAdd("S", 1, CDate(mvarDtDelete)))
      End If
      If Not FrmObj Is Nothing Then
         FrmObj.TrayIcon.Text = FrmObj.Tag
         FrmObj.Caption = FrmObj.Tag
      End If
   End If
   
   Exit Function
Saida:
   mvarWebConnected = IsWebConnected
   If mvarWebConnected Then Resume Next
End Function
Private Function DeleteCascade(pDbRem As Object, pTab As String, pWhere As String) As Boolean
   Dim RsTabs As Object
   Dim Sql As String
   Dim bOk As Boolean
   Dim bResult As Boolean
   
   bResult = True
   
   Sql = "Select Name AS ChaveEstrangeira, OBJECT_NAME(Parent_Object_ID) AS TabelaFilho, OBJECT_NAME(Referenced_Object_ID) As TabelaPai"
   Sql = Sql & " From SYS.FOREIGN_KEYS"
   Sql = Sql & " Where OBJECT_NAME(Referenced_Object_ID)=" & SqlStr(pTab)
   If pDbRem.AbreTabela(Sql, RsTabs) Then
      While Not RsTabs.EOF
         Sql = "Delete From " & RsTabs("TabelaFilho")
         Sql = Sql & pWhere
         
         bOk = pDbRem.Executa(Sql)
         If Not bOk Then
            bOk = DeleteCascade(pDbRem, RsTabs("TabelaFilho"), pWhere)
         End If
         bResult = bResult And bOk
         
         RsTabs.MoveNext
      Wend
   End If
   DeleteCascade = bResult
End Function
Function SyncTable(xDbOrig As Object, xDbDest As Object, sTab As String, pNumTab As Integer) As Boolean
   Dim MyTabOri   As Object
   Dim MyTabDest   As Object
   Dim nErr       As Long
   Dim bImportou  As Boolean
   Dim sFilter    As String
   Dim bSalvou    As Boolean
   Dim bRemoteDest As Boolean
   Dim bAchou     As Boolean
   Dim sTextLog   As String
   Dim sTextSinc  As String
   Dim Queries    As Collection
   Dim Sql        As String
         
   On Error GoTo Saida
   
   If Trim(sTab) = "" Then Exit Function
   Set MyTabOri = CriarObjeto("BANCO_3R.TB_" & sTab)
   If MyTabOri Is Nothing Then Set MyTabOri = CriarObjeto("BANCO.TB_" & sTab)
   If MyTabOri Is Nothing Then Exit Function
   
   Set MyTabDest = CriarObjeto("BANCO_3R.TB_" & sTab)
   If MyTabDest Is Nothing Then Set MyTabDest = CriarObjeto("BANCO.TB_" & sTab)
   If MyTabOri Is Nothing Then Exit Function
     
   Set MyTabOri.xDb = xDbOrig
   Set MyTabDest.xDb = xDbDest
   
   bRemoteDest = LCase(xDbDest.SERVER) = LCase(xDbDest.SERVERNAME("[Remote]"))
   
   mvarSincTab = True
   sFilter = Replace(mvarSincFilter, "@@TABELA", sTab)
   If bRemoteDest Then
      If xDbDest.ExisteReg("Select Top 1 * From " & sTab & " Where IDLOJA In (" & mvarLojasIn & ")") Then
         sFilter = sFilter & " And (ALTERSTAMP=1)"
      End If
   End If
   If sTab = "OSESSAO" Then
      Debug.Print "MyTabOri: " & MyTabOri.xDb.SERVER
      Debug.Print MyTabOri.QrySelect(Ch_Where:=sFilter, Ch_OrderBy:="TIMESTAMP")
      sTab = sTab
   End If
   sTextLog = FrmObj.TxtLog.Text
   bAchou = MyTabOri.Pesquisar(Ch_Where:=sFilter, Ch_OrderBy:="TIMESTAMP")
   If bAchou Then
      nErr = 0
      While Not MyTabOri.Rs.EOF
         MyTabOri.Popula
         Call MyTabDest.Popula(MyTabOri.Rs)
         
         'If MyTabOri.Rs("IDEVENTO") = 1908 Then
         '    bImportou = bImportou
         'End If
         
         bImportou = False
         bSalvou = True
         If mvarImport And mvarExport Then
            If MyTabDest.Pesquisar Then
               '**********
               '* Se Servidor mais atual então atualiza local
               If CDate(MyTabDest.TimeStamp) > CDate(MyTabOri.TimeStamp) Then  'And MyTabDest.ALTERSTAMP = 1
                  Set Queries = New Collection
                  Call MyTabOri.Popula(MyTabDest.Rs)
               
                  Queries.Add MyTabOri.QrySave(pAutoId:=False, pSinc:=True)
                  If Not bRemoteDest Then
                     Sql = "Update GSINC"
                     Sql = Sql & " Set DTSINC = " & SqlDate(CDate(MyTabOri.TimeStamp) - (1 / 24 / 60))
                     Sql = Sql & " Where IDLOJA in (" & mvarLojasIn & ")"
                     Sql = Sql & " And TABELA = " & SqlStr(sTab)
                     Sql = Sql & " And CODMAQ <>" & SqlStr(Environ("COMPUTERNAME"))
                     Queries.Add Sql
                  End If
                  bSalvou = xDbOrig.Executa(Queries)
                  
                  If bSalvou Then
                     If bRemoteDest Then
                        MyTabOri.ALTERSTAMP = 0
                     Else
                        MyTabDest.ALTERSTAMP = 0
                     End If
                     bImportou = True
                  Else
                     mvarSincTab = False
                  End If
               End If
            End If
         End If
         
         If Not bImportou Then
            If bSalvou Then
               Set Queries = New Collection
               Call MyTabDest.Popula(MyTabOri.Rs)
               
               Queries.Add MyTabDest.QrySave(pAutoId:=False, pSinc:=True)
               If bRemoteDest Then
                  Sql = "Update GSINC"
                  Sql = Sql & " Set DTSINC = " & SqlDate(CDate(MyTabDest.TimeStamp) - (1 / 24 / 60))
                  Sql = Sql & " Where IDLOJA in (" & mvarLojasIn & ")"
                  Sql = Sql & " And TABELA = " & SqlStr(sTab)
                  Sql = Sql & " And CODMAQ <>" & SqlStr(Environ("COMPUTERNAME"))
                  Queries.Add Sql
               End If
               bSalvou = xDbDest.Executa(Queries)
            End If
         End If
         If bSalvou Then
            If bRemoteDest Then
               MyTabOri.ALTERSTAMP = 0
            Else
               MyTabDest.ALTERSTAMP = 0
            End If
            nErr = 0
         Else
            mvarSincTab = False
            nErr = nErr + 1
            If nErr >= 10 Then
               MyTabOri.Rs.MoveLast
            End If
         End If
         
         If Not FrmObj Is Nothing Then
            'FrmObj.Width = 5700
            DoEvents
            FrmObj.TrayIcon.Text = IIf(bRemoteDest, ">", "<") & "[" & sTab & "] - " & ValBr((MyTabOri.Rs.AbsolutePosition / MyTabOri.Rs.RecordCount) * 100, 2) & "%"
            FrmObj.Caption = FrmObj.Tag & " - " & IIf(bRemoteDest, ">", "<") & "[" & pNumTab & "].[" & sTab & "] - " & MyTabOri.Rs.AbsolutePosition & " / " & MyTabOri.Rs.RecordCount
            
            FrmObj.TxtLog.Text = Left("[" & Format(Now(), "hh:mm:ss") & "] " & Mid(FrmObj.Caption, Len(FrmObj.Tag) + 1) & vbNewLine & sTextLog, 10000)
            DoEvents
            'FrmObj.TxtLog.SelStart = Len(FrmObj.TxtLog.Text)
            'FrmObj.TxtLog.SelLength = 0
            'Call ScrollText(FrmObj.TxtLog.hwnd, 7)
            'FrmObj.Refresh
         End If
         
         DoEvents
         While mvarPause: DoEvents: Wend
         mvarWebConnected = IsWebConnected
         If Not mvarWebConnected Then GoTo Saida
         
         MyTabOri.Rs.MoveNext
         'For i = 1 To 10600
         '   MyTabOri.Rs.MoveNext
         'Next
      Wend
      If Not FrmObj Is Nothing Then
         FrmObj.TrayIcon.Text = FrmObj.Tag
         FrmObj.Caption = FrmObj.Tag
      End If
   Else
      FrmObj.TxtLog.Text = Left("[" & Format(Now(), "hh:mm:ss") & "] - " & IIf(bRemoteDest, ">", "<") & "[" & pNumTab & "].[" & sTab & "]" & vbNewLine & sTextLog, 10000)
      FrmObj.TxtLog.SelStart = Len(FrmObj.TxtLog.Text)
      FrmObj.TxtLog.SelLength = 0
   
   End If
   'If mvarSincTab Then
   '   If Not bAchou Then
   '      sFilter = Replace(mvarSincFilter, "@@TABELA", sTab)
   '      bAchou = MyTabOri.Pesquisar(Ch_Where:=sFilter, Ch_OrderBy:="TIMESTAMP")
   '   End If
   'End If
   SyncTable = bAchou
   
   Set MyTabOri = Nothing
   Set MyTabDest = Nothing
   
   Exit Function
Saida:
   mvarWebConnected = IsWebConnected
   If mvarWebConnected Then Resume Next
End Function
Public Sub Resinc(pDate As String)
   Dim Sql        As String
   Dim sTabelas   As String
   Dim QtdTAbs    As Integer
   Dim i          As Integer
   Dim xDbLoc     As Object
   
   Call ConectarServidor(xDbLoc, 1)
   sTabelas = GetTables(xDbLoc, "C.NAME In (" & mvarFieldsOnTab & ")")
   QtdTAbs = GetTag(sTabelas, "TABS", 0)
   
   For i = 1 To QtdTAbs
      Sql = "Update " & GetTag(sTabelas, "TAB" & i, "")
      Sql = Sql & " Set ALTERSTAMP=1 "
      Sql = Sql & " Where TIMESTAMP>=" & SqlDate(pDate)
      Sql = Sql & " And IDLOJA in (" & mvarLojasIn & ");"
      Call xDbLoc.Executa(Sql)
   Next
End Sub
Private Sub RefreshLoja(xDbRem As Object, xDbLoc As Object)
   Dim Sql     As String
   Dim sAux    As String
   Dim nAux    As Integer
   Dim MyRs    As Object
   
   'Dim xDBRem  As Object

   If xDbRem.Conectado Then
      Sql = "Delete From GSINC "
      Sql = Sql & " Where TABELA='OLOJA'"
      Sql = Sql & " And CODMAQ = '" & Environ("COMPUTERNAME") & "'"
      Call xDbRem.Executa(Sql)
   
      Sql = "Select Distinct IDLOJA, IDCOLIGADA"
      Sql = Sql & " From OLOJA"
      Sql = Sql & " Where IDCOLIGADA In (Select IDCOLIGADA From OLOJA Where IDLOJA in (" & mvarLojasIn & "))"
      If xDbRem.AbreTabela(Sql, MyRs) Then
         mvarLojasIn = ""
         mvarLojasLike = ""
         While Not MyRs.EOF
            mvarLojasIn = mvarLojasIn & IIf(Trim(mvarLojasIn) = "", "", ",") & MyRs("IDLOJA") & ""
            mvarLojasLike = mvarLojasLike & IIf(Trim(mvarLojasLike) = "", "", " Or ") & "QUERY Like '%IDLOJA = " & MyRs("IDLOJA") & "%'"
            MyRs.MoveNext
         Wend
         If InStr(mvarSincFilter, " IDLOJA In") <> 0 Then
            sAux = Mid(mvarSincFilter, InStr(mvarSincFilter, " IDLOJA In"), Len(" IDLOJA In"))
            nAux = InStr(InStr(mvarSincFilter, " IDLOJA In") + 1, mvarSincFilter, ")")
            
            sAux = sAux & " (" & mvarLojasIn & ")" & Mid(mvarSincFilter, nAux + 1)
            mvarSincFilter = sAux
         End If
      End If
      
      If mvarImport Then Call SyncTable(xDbRem, xDbLoc, "OLOJA", 0)
   End If
End Sub
Private Sub RefreshColigada(xDbRem As Object, xDbLoc As Object)
   Dim Sql     As String
   Dim sAux    As String
   Dim nAux    As Integer
   'Dim xDBRem  As Object

   'If xDbRem.Conectado Then
   '   If mvarImport Then
   '      Call SyncTable(xDbRem, xDbLoc, "COLIGADA")
   '   End If
   'End If
End Sub
Private Sub AtualizaDTSINC(xDbLoc As Object, pTab As String, pDtSinc As Date)
   Dim MyTabOri   As Object
   Dim MyTabDest  As Object
   Dim sIDLOJA    As String
   Dim nAux       As Integer
   Dim sAux       As String
    
   If mvarSincTab And mvarSincDelTab Then
      sAux = Replace(Replace(mvarLojasIn, "(", ""), ")", "")
      nAux = InStr(sAux, ",")
      If nAux = 0 Then
         sIDLOJA = sAux
      Else
         sIDLOJA = Mid(sAux, 1, nAux - 1)
      End If
      While xVal(sIDLOJA) <> 0
         Set MyTabOri = CriarObjeto("BANCO_3R.TB_GSINC")
         If MyTabOri Is Nothing Then Set MyTabOri = CriarObjeto("BANCO.TB_GSINC")
         With MyTabOri
            Set .xDb = xDbLoc
            Call .Pesquisar(Ch_IDLOJA:=sIDLOJA, Ch_CODMAQ:=Environ("COMPUTERNAME"), Ch_TABELA:=pTab)
            .IDLOJA = sIDLOJA
            .CODMAQ = Environ("COMPUTERNAME")
            .TABELA = pTab
            .DTSINC = pDtSinc
            Call .Salvar
         End With
         'Set MyTabDest = CriarObjeto("BANCO_3R.TB_GSINC")
         'If MyTabDest Is Nothing Then Set MyTabDest = CriarObjeto("BANCO.TB_GSINC")
         'With MyTabDest
         '   Set .xDB = xDbDest
         '   Call .Pesquisar(Ch_IDLOJA:=sIDLOJA, Ch_CODMAQ:=Environ("COMPUTERNAME"), Ch_TABELA:=pTab)
         '   .IDLOJA = sIDLOJA
         '   .CODMAQ = Environ("COMPUTERNAME")
         '   .TABELA = pTab
         '   .DTSINC = pDtSinc
         '   If MyTabOri.TimeStamp = "" Then
         '      .TimeStamp = xDbDest.SysDate(3)
         '   Else
         '      .TimeStamp = MyTabOri.TimeStamp
         '   End If
         '   Call .Salvar(pSinc:=True)
         'End With
         
         sAux = Trim(Mid(sAux, nAux + 1))
         nAux = InStr(sAux, ",")
         If nAux = 0 Then
            If sIDLOJA <> sAux Then
               sIDLOJA = sAux
            Else
               sIDLOJA = ""
            End If
            sAux = ""
         Else
            sIDLOJA = Mid(sAux, 1, nAux - 1)
         End If
      Wend
   End If
End Sub
Function GetTables(pxDb As Object, Optional pCampos As String = "") As String
   Dim Sql     As String
   Dim RsTab   As Object
   Dim sAux    As String
   Dim QtdTAbs As Integer
   Dim i       As Integer
   
   sAux = ""
   sAux = sAux & "|TABS=83"
   sAux = sAux & "|TAB1=COLIGADA"
   sAux = sAux & "|TAB2=OLOJA"
   sAux = sAux & "|TAB3=OTPMAQ"
   sAux = sAux & "|TAB4=OTPCONHEC"
   sAux = sAux & "|TAB5=OSALA"
   sAux = sAux & "|TAB6=OMAQUINA"
   sAux = sAux & "|TAB7=OCLIENTE"
   sAux = sAux & "|TAB8=OAGENDA"
   sAux = sAux & "|TAB9=SESTOQUE"
   sAux = sAux & "|TAB10=RFUNCIONARIO"
   sAux = sAux & "|TAB11=OTPMANIPULO"
   sAux = sAux & "|TAB12=OSALA_MAQUINA"
   sAux = sAux & "|TAB13=OMANIPULO"
   sAux = sAux & "|TAB14=OEVENTOREC"
   sAux = sAux & "|TAB15=OEVENTOAGENDA"
   sAux = sAux & "|TAB16=ODIRECAO"
   sAux = sAux & "|TAB17=OATENDIMENTO"
   sAux = sAux & "|TAB18=OAREA"
   sAux = sAux & "|TAB19=GPAIS"
   sAux = sAux & "|TAB20=FMOEDA"
   sAux = sAux & "|TAB21=CVENDA"
   sAux = sAux & "|TAB22=CFORMAPGTO"
   sAux = sAux & "|TAB23=CCONDPGTO"
   sAux = sAux & "|TAB24=SPRODUTO"
   sAux = sAux & "|TAB25=SESTPROD"
   sAux = sAux & "|TAB26=PPN"
   sAux = sAux & "|TAB27=PMETA"
   sAux = sAux & "|TAB28=OTPTRATAMENTO"
   sAux = sAux & "|TAB29=OTPSERVICO"
   sAux = sAux & "|TAB30=OSESSAO"
   sAux = sAux & "|TAB31=OREACAOADV"
   sAux = sAux & "|TAB32=OLAMPADA"
   sAux = sAux & "|TAB33=OCLASSE"
   sAux = sAux & "|TAB34=OCATTAREFA"
   sAux = sAux & "|TAB35=FLOTERPS"
   sAux = sAux & "|TAB36=FGRUPODESP"
   sAux = sAux & "|TAB37=FFATURA"
   sAux = sAux & "|TAB38=FDESPESA"
   sAux = sAux & "|TAB39=FCCORRENTE"
   sAux = sAux & "|TAB40=CPROMOCAO"
   sAux = sAux & "|TAB41=CPGTOSVENDA"
   sAux = sAux & "|TAB42=CCUPOM"
   sAux = sAux & "|TAB43=CBANDC"
   sAux = sAux & "|TAB44=SMOVEST"
   sAux = sAux & "|TAB45=SEST_LOJA"
   sAux = sAux & "|TAB46=RBATIDA"
   sAux = sAux & "|TAB47=PPNITEM"
   sAux = sAux & "|TAB48=PMETAITEM"
   sAux = sAux & "|TAB49=PARAM"
   sAux = sAux & "|TAB50=OTRATAMENTOCLI"
   sAux = sAux & "|TAB51=OTRATAMENTO_PROD"
   sAux = sAux & "|TAB52=OTPMANUT"
   sAux = sAux & "|TAB53=OTIPOTAREFA"
   sAux = sAux & "|TAB54=OTAREFAEVT"
   sAux = sAux & "|TAB55=OTAREFA_CAT"
   sAux = sAux & "|TAB56=OSESSAO_REACAOADV"
   sAux = sAux & "|TAB57=OSERVICOEVT"
   sAux = sAux & "|TAB58=OPENDENCIACLI"
   sAux = sAux & "|TAB59=OMAQMANUT"
   sAux = sAux & "|TAB60=OMAQDISPAROS"
   sAux = sAux & "|TAB61=OFOTOTIPO"
   sAux = sAux & "|TAB62=ODIARIO"
   sAux = sAux & "|TAB63=OCONTATO"
   sAux = sAux & "|TAB64=OCLIENTEFOTOS"
   sAux = sAux & "|TAB65=OCLASSE_CONTATO"
   sAux = sAux & "|TAB66=OCLASSE_CLIENTE"
   sAux = sAux & "|TAB67=OATENDIMENTO_VENDA"
   sAux = sAux & "|TAB68=OATENDIMENTO_PRODUTO"
   sAux = sAux & "|TAB69=OAREA_TRATAMENTO"
   sAux = sAux & "|TAB70=GSINC"
   sAux = sAux & "|TAB71=GFERIADO"
   sAux = sAux & "|TAB72=FRECIBO"
   sAux = sAux & "|TAB73=FORNECEDOR_PRODUTO"
   sAux = sAux & "|TAB74=FORNECEDOR"
   sAux = sAux & "|TAB75=FNOTAFISCAL"
   sAux = sAux & "|TAB76=FLAN"
   sAux = sAux & "|TAB77=FITEMFATURA"
   sAux = sAux & "|TAB78=DELETEDROWS"
   sAux = sAux & "|TAB79=DADOSPARAM"
   sAux = sAux & "|TAB80=CITENSVENDA"
   sAux = sAux & "|TAB81=CCUPOM_VENDA"
   sAux = sAux & "|TAB82=CBANDCTX"
   sAux = sAux & "|TAB83=CPACOTE"
   sAux = sAux & "|TAB84=CPROMO_PROD"
   
   QtdTAbs = GetTag(sAux, "TABS", 0)
   
   Sql = "SELECT DISTINCT O.NAME [TABELA], O.CRDATE"
   Sql = Sql & " FROM SYSOBJECTS O JOIN SYSCOLUMNS C ON C.ID=O.ID"
   Sql = Sql & " WHERE O.XTYPE='U'"
   If pCampos <> "" Then Sql = Sql & " And " & pCampos
   Sql = Sql & " And Not O.NAME in ("
   For i = 1 To QtdTAbs
      Sql = Sql & "'" & GetTag(sAux, "TAB" & i, "") & "',"
   Next
   Sql = Mid(Sql, 1, Len(Sql) - 1)
   Sql = Sql & ")"
   Sql = Sql & " ORDER BY O.CRDATE"
      
   If pxDb.AbreTabela(Sql, RsTab) Then
      Call SetTag(sAux, "TABS", RsTab.RecordCount + QtdTAbs)
      RsTab.MoveFIRST
      While Not RsTab.EOF
         sAux = sAux & "|TAB" & QtdTAbs + RsTab.AbsolutePosition & "=" & RsTab("TABELA")
         RsTab.MoveNext
      Wend
   End If
   Set RsTab = Nothing
   
   sAux = sAux & "|"
   GetTables = sAux
End Function
Private Function ConectarServidor(ByRef xDbRem As Object, LocalRemote As Integer) As Boolean
  
   Set xDbRem = CriarObjeto("XBANCO01.DS_BANCO")
   With xDbRem
      If LocalRemote = 1 Then '* 1=Local Settings
         .SERVER = mvarLocalServer
         .dbname = mvarLocalDbName
         .UID = mvarLocalUID
         .PWD = mvarLocalPWD
      ElseIf LocalRemote = 2 Then '* 2=Remote Settings
         .SERVER = mvarRemoteServer
         .dbname = mvarRemoteDbName
         .UID = mvarRemoteUID
         .PWD = mvarRemotePWD
      End If
      
      .SrvConecta
      If .Conectado Then
         ConectarServidor = True
      Else
         ConectarServidor = False
      End If
   End With
End Function
Private Sub Class_Initialize()
   mvarIniFile = Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\RESTOREBD.ini"
   If ReadIniFile(mvarIniFile, "CONFIG", "MERGE", "0") = "0" Then
      mvarExport = True
      mvarImport = True
   ElseIf ReadIniFile(mvarIniFile, "CONFIG", "MERGE", "1") = "0" Then
      mvarExport = True
      mvarImport = False
   ElseIf ReadIniFile(mvarIniFile, "CONFIG", "MERGE", "2") = "0" Then
      mvarExport = False
      mvarImport = True
   End If
   mvarDelScriptTab = "DELETEDROWS"
   mvarFieldsOnTab = "'TIMESTAMP'"
   
   mvarSincFilter = ""
   'mvarSincFilter = "(ALTERSTAMP=1 Or ALTERSTAMP Is Null) And"
   mvarSincFilter = mvarSincFilter & " IDLOJA In (2,3,4)"
   mvarSincFilter = mvarSincFilter & " And (TIMESTAMP>("
   mvarSincFilter = mvarSincFilter & " Select IsNull(Min(DTSINC),0)"
   mvarSincFilter = mvarSincFilter & " From GSINC"
   mvarSincFilter = mvarSincFilter & " Where IDLOJA In (2,3,4)"
   mvarSincFilter = mvarSincFilter & " And CODMAQ=" & SqlStr(Environ("COMPUTERNAME"))
   mvarSincFilter = mvarSincFilter & " And TABELA='@@TABELA'"
   mvarSincFilter = mvarSincFilter & "))"
   
   
   mvarDeletedFilter = ""
   'mvarDeletedFilter = "(ALTERSTAMP=1 Or ALTERSTAMP Is Null) And"
   mvarDeletedFilter = mvarDeletedFilter & " (QUERY Like '%IDLOJA = 2%' Or QUERY Like '%IDLOJA = 3%' Or QUERY Like '%IDLOJA = 4%')"
   mvarDeletedFilter = mvarDeletedFilter & " And (TIMESTAMP>("
   mvarDeletedFilter = mvarDeletedFilter & " Select IsNull(Min(DTSINC),0)"
   mvarDeletedFilter = mvarDeletedFilter & " From GSINC"
   mvarDeletedFilter = mvarDeletedFilter & " Where IDLOJA In (2,3,4)"
   mvarDeletedFilter = mvarDeletedFilter & " And CODMAQ=" & SqlStr(Environ("COMPUTERNAME"))
   mvarDeletedFilter = mvarDeletedFilter & " And TABELA=" & SqlStr(mvarDelScriptTab)
   mvarDeletedFilter = mvarDeletedFilter & "))"
   
   
   
   If Not mvarFrmObj Is Nothing Then
      mvarFrmObj.OptMerge(Val(ReadIniFile(mvarIniFile, "CONFIG", "MERGE", "0"))).Value = True
   End If
End Sub
