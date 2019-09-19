VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DS_BANCO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DS Banco de Dados"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Private WithEvents rs As ADODB.Recordset
Option Explicit
Private mvarWS             As Workspace
Private mvarWsODBC         As Workspace
Private mvarxDBs           As New Collection
Private mvarKey            As String

'=================
'= Propriedades  =
'=================
Private mvardBase          As Database
Private mvarADOConect      As ADODB.Connection
Private mvarConect         As Connection
Private mvarDbObj          As Object

Private mvardBases         As New Collection
Private mvarConections     As New Collection
Private mvarADOConections  As New Collection
Private mvarTables         As Collection

Private mvarRs             As DAO.Recordset
Private mvarADORs          As New ADODB.Recordset
Private mvarRsAux          As Object

Private mvarisODBC         As Boolean
Private mvarisADO          As Boolean
Private mvarConectado      As Boolean
'Private mvarCodeSql        As Integer
Private mvarMinutosConexao As Integer

Private mvarDbDrive        As String
Private mvarDbName         As String
Private mvarDSN            As String
Private mvarUID            As String
Private mvarPWD            As String
Private mvarDriver         As String
Private mvarServer         As String
Private mvarStrDATABASE    As String
Private mvarStrConect      As String
Private mvarAlias          As String
Private mvarDSNDrive       As String
Private mvarTimeOut        As Long

'****************************
'* Propriedades para suprir *
'* a ausência de DSR100.DLL *
'****************************
Private mvarDrvTmp         As String
Private mvarDrvErro        As String
Private mvarSepDec         As String
Private mvarSepMil         As String
Private mvarSepDt          As String
Private mvarDtMask         As String
Private mvarComputerName   As String
Private mvarRemoteComputerName   As String
Private mvarIdioma         As eIdioma

'=================
'= User Define   =
'=================
Enum eSysDate
   Data = 1
   HORA
   Data_Hora
End Enum
Enum eIdioma
   PORTUGUES = 5000
   INGLES = 6000
   FRANCES = 7000
   ESPANHOL = 8000
End Enum
Public Enum eDbTipo
   Access = 0
   SQL_SERVER
   ORACLE
   MySQL
End Enum
Enum eQry
   eINCLUSAO = 1
   eALTERACAO
   eEXCLUSAO
   eLEITURA
End Enum
Enum eFind
   NotFound = False
   Found = True
   ERRO = 1
End Enum
Enum eTipoDeDados
   NUMERO = 1
   DATAHORA = 2
   Caracter = 3
End Enum
Private Type Tabela
   NAME As String
   OWNER As String
   isSys As Boolean
   FIELDS As Collection
End Type

'=================
'= Constantes    =
'=================
Private mvarDbTipo         As eDbTipo
Private mvarDbVersao       As String

Public SetProperty         As Boolean

Private mvarCodeSql        As eFind
Public Property Let TimeOut(ByVal vData As Long)
   mvarTimeOut = vData

   On Error Resume Next
   mvarDbObj.CommandTimeout = mvarTimeOut
   mvarADOConect.CommandTimeout = mvarTimeOut
   mvarDbObj.QueryTimeout = mvarTimeOut
   mvarConect.CommandTimeout = mvarTimeOut
   mvardBase.QueryTimeout = mvarTimeOut
End Property
Public Property Get TimeOut() As Long
   TimeOut = mvarTimeOut
End Property
Public Property Let dbVersao(ByVal vData As String)
   mvarDbVersao = vData
End Property
Public Property Get dbVersao() As String
   dbVersao = mvarDbVersao
End Property
Public Property Let DrvErro(vData As String)
   mvarDrvErro = vData
End Property
Public Property Get DrvErro() As String
   DrvErro = mvarDrvErro
End Property
Public Property Let Idioma(vData As eIdioma)
   mvarIdioma = vData
   gIdioma = vData
End Property
Public Property Get Idioma() As eIdioma
   Idioma = mvarIdioma
End Property
Public Property Let DtMask(vData As String)
   mvarDtMask = vData
   gDtMask = vData
End Property
Public Property Get DtMask() As String
   DtMask = mvarDtMask
End Property
Public Property Let SepMil(vData As String)
   mvarSepMil = vData
   gSepMil = vData
End Property
Public Property Get SepMil() As String
   SepMil = mvarSepMil
End Property
Public Property Let SepDec(vData As String)
   mvarSepDec = vData
   gSepDec = vData
End Property
Public Property Get SepDec() As String
   SepDec = mvarSepDec
End Property
Public Property Let SepDt(vData As String)
   mvarSepDt = vData
   gSepDt = vData
End Property
Public Property Get SepDt() As String
   SepDt = mvarSepDt
End Property
Public Property Let ComputerName(vData As String)
   mvarComputerName = vData
   gComputerName = vData
End Property
Public Property Get ComputerName() As String
   ComputerName = mvarComputerName
End Property
Public Property Let RemoteComputerName(vData As String)
   mvarRemoteComputerName = vData
End Property
Public Property Get RemoteComputerName() As String
   RemoteComputerName = mvarRemoteComputerName
End Property

Public Property Let DrvTmp(vData As String)
   mvarDrvTmp = vData
   gDrvTmp = vData
End Property
Public Property Get DrvTmp() As String
   DrvTmp = mvarDrvTmp
End Property
'*********************************************
Public Property Let dbTipo(vData As eDbTipo)
   mvarDbTipo = vData
End Property
Public Property Get dbTipo() As eDbTipo
   dbTipo = mvarDbTipo
End Property
Public Property Set DbObj(ByVal vData As Object)
   Dim i%, j%, n
   Set mvarDbObj = vData
If SetProperty Then Exit Property
   If mvarxDBs.Count = 0 Then Exit Property
   i = 0
   j = 0
   For Each n In mvarxDBs
      i = i + 1
      If n Is vData Then
         j = i
         Exit For
      End If
   Next
   Call Set_Property(mvarxDBs(j), Me)
End Property
Public Property Get DbObj() As Object
   Set DbObj = mvarDbObj
End Property
Public Property Set WsODBC(ByVal vData As Workspace)
   Set mvarWsODBC = vData
End Property
Public Property Get WsODBC() As Variant
   If mvarWsODBC Is Nothing Then
      Set DBEngine = Nothing
      Set mvarWsODBC = DBEngine.CreateWorkspace("WsODBC", "admin", "", dbUseODBC)
   End If
   Set WsODBC = mvarWsODBC
End Property
Public Property Set WS(ByVal vData As Workspace)
   Set mvarWS = vData
End Property
Public Property Get WS() As Variant
   If mvarWS Is Nothing Then
      Set DBEngine = Nothing
      Set mvarWS = DBEngine.CreateWorkspace("WsEngine", "admin", "")
   End If
   Set WS = mvarWS
End Property
Public Property Get dBase() As Database
   If mvarisODBC And mvardBase Is Nothing And mvarStrConect$ <> "" Then
      Set mvardBase = WS.OpenDatabase("", False, False, mvarStrConect$)
   End If
   Set dBase = mvardBase
End Property
Public Property Let dBase(ByVal vData As Database)
   Dim n, JaExiste As Boolean
   Dim i%
   Set mvardBase = vData
If SetProperty Then Exit Property
   If mvarxDBs.Count = 0 Then Exit Property
   i% = 0
   For Each n In mvardBases
      i = i + 1
      If TypeName(n) = "Database" Then
         If n.NAME = vData.NAME Then JaExiste = True
      ElseIf TypeName(n) = "Connection" Then
         If n.Connect = vData.Connect Then JaExiste = True
      End If
   Next
   If Not JaExiste Then
      i = i + 1
      dBases.Add vData
   End If
   Call Set_Property(mvarxDBs(i), Me)
   If isODBC Then Call GetODBCConnectParts(mvardBase.Connect)
End Property
Public Property Set dBase(ByVal vData As Database)
   Dim n, JaExiste As Boolean
   Dim i%
   Set mvardBase = vData
If SetProperty Then Exit Property
   If mvarxDBs.Count = 0 Then Exit Property
   
   For Each n In dBases
      i = i + 1
      If TypeName(n) = "Database" Then
         If n.NAME = vData.NAME Then JaExiste = True
      ElseIf TypeName(n) = "Connection" Then
         If n.Connect = vData.Connect Then JaExiste = True
      End If
   Next
   If Not JaExiste Then
      i = i + 1
      dBases.Add vData
   End If
   Call Set_Property(mvarxDBs(i), Me)
   If isODBC Then Call GetODBCConnectParts(mvardBase.Connect)
End Property
Public Property Let dbDrive(vData As String)
   mvarDbDrive = vData
End Property
Public Property Get dbDrive() As String
   dbDrive = mvarDbDrive
End Property
Public Property Get dbName() As String
   dbName = mvarDbName
End Property
Public Property Let dbName(vData As String)
   If LCase(VBA.Right$(vData, 4)) = ".mdb" Then
      mvarDbTipo = 0 'eDbTipo.Access
   End If
   mvarDbName = vData
End Property
Public Property Get StrConect() As String
   StrConect = mvarStrConect
End Property
Public Property Let StrConect(vData As String)
   mvarStrConect = vData
End Property
Public Property Get Driver() As String
   Driver = mvarDriver
End Property
Public Property Let Driver(vData As String)
   mvarDriver = vData
End Property
Public Property Get Server() As String
   Server = mvarServer
End Property
Public Property Let Server(vData As String)
   mvarServer = ServerName(vData)
End Property
Public Function ServerName(ByVal pSERVER As String, Optional ByVal pLocalName As String, Optional ByVal pRemoteName As String)
   pLocalName = IIf(Trim(pLocalName) = "", mvarComputerName, Trim(pLocalName))
   pRemoteName = IIf(Trim(pRemoteName) = "", mvarRemoteComputerName, Trim(pRemoteName))
   
   If InStr(pSERVER, "[Local]") Then pSERVER = Replace(pSERVER, "[Local]", pLocalName)
   If InStr(pSERVER, "[Remote]") Then pSERVER = Replace(pSERVER, "[Remote]", pRemoteName)
   
   ServerName = pSERVER
End Function
Public Property Get DSN() As String
   DSN = mvarDSN
End Property
Public Property Let DSN(vData As String)
   mvarDSN = vData
End Property
Public Property Get DSNDrive() As String
   DSNDrive = mvarDSNDrive
End Property
Public Property Let DSNDrive(vData As String)
   mvarDSNDrive = vData
End Property
Public Property Get UID() As String
   UID = mvarUID
End Property
Public Property Let UID(vData As String)
   mvarUID = vData
End Property
Public Property Get PWD() As String
   PWD = mvarPWD
End Property
Public Property Let PWD(vData As String)
   mvarPWD = vData
End Property
Public Property Get Key() As String
   Key = mvarKey
End Property
Friend Property Let Key(vData As String)
   mvarKey = vData
End Property
Public Property Get StrDATABASE() As String
   StrDATABASE = mvarStrDATABASE
End Property
Public Property Let StrDATABASE(vData As String)
   mvarStrDATABASE = vData
End Property
Public Property Get Alias() As String
   Alias = mvarAlias
End Property
Public Property Let Alias(vData As String)
   mvarAlias = vData
End Property
Public Property Set dBases(ByVal vData As Object)
   Set mvardBases = vData
End Property
Public Property Get dBases() As Collection
   Set dBases = mvardBases
End Property
Public Property Set xDBs(ByVal vData As Object)
   Set mvarxDBs = vData
End Property
Public Property Get xDBs() As Collection
   Set xDBs = mvarxDBs
End Property
Public Property Set Tables(ByVal vData As Object)
   Set mvarTables = vData
End Property
Public Property Get Tables() As Collection
   Dim n          As Variant
   Dim n1         As Variant
   Dim MyDataBase As Database
   Dim WsAux      As Workspace
   Dim dBaseAux   As Database
   Dim MyTab      As DbTabela
   Dim MyField    As DbCampo
   
   Dim RsTables   As New ADODB.Recordset
   Dim RsFields   As New ADODB.Recordset
   
   Dim Achou      As Boolean
   
   On Error GoTo fim
   
   If mvarTables Is Nothing Then
      If mvarisADO Then
         If mvarADOConect Is Nothing Then Exit Sub
         
         Set RsTables = mvarADOConect.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
         
         Set mvarTables = New Collection
         Do Until RsTables.EOF
            'DoEvents
            Set MyTab = New DbTabela
            With MyTab
              Set .ADOConect = mvarADOConect
              .NOME = RsTables("TABLE_NAME") & ""
              .DESCRICAO = RsTables("DESCRIPTION") & ""
              .BANCO = RsTables("TABLE_CATALOG") & ""
              .OWNER = RsTables("TABLE_SCHEMA") & ""
              
              .isSys = False
              mvarTables.Add MyTab, .NOME
            End With
            Set MyTab = Nothing
            RsTables.MoveNext
         Loop
         RsTables.Close
      Else
         Select Case TypeName(mvarDbObj)
            Case "Database"
               Set mvarTables = New Collection
               For Each n In mvarDbObj.TableDefs
                  DoEvents
                  With MyTab
                     .OWNER = ""
                     .isSys = False
                     For Each n1 In n.FIELDS
                        .FIELDS.Add n1.NAME
                     Next
                  End With
                  mvarTables.Add n.NAME
               Next
            Case "Connection"
               Set mvarTables = New Collection
               Set DBEngine = Nothing
               Set WsAux = DBEngine.CreateWorkspace("WsAux", "admin", "")
               If dBaseAux Is Nothing Then
                  Set dBaseAux = WsAux.OpenDatabase("", False, False, mvarStrConect$)
                  
                  Dim Total As Long
                  Total = dBaseAux.TableDefs.Count
                  'Set mvarConect.Database = dBaseAux
                  For Each n In dBaseAux.TableDefs
                     DoEvents
                     Dim Bool As Boolean
                     Bool = ((n.Attributes And dbSystemObject) = 0)
                     If Bool Then
                        Set MyTab = New DbTabela
                        With MyTab
                           .NOME = n.NAME
                           If InStr(n.NAME, ".") > 1 Then
                              .OWNER = Mid(n.NAME, 1, InStr(n.NAME, ".") - 1)
                           End If
                           .isSys = False
                           .StrConect = mvarStrConect
                           '                        For Each n1 In n.FIELDS
                           '                        DoEvents
                           '                           If n1.Attributes < dbSystemField Then
                           '                              Set MyField = New DbCampo
                           '                              With MyField
                           '                                 .NOME = n1.NAME
                           '                                 .TAM = n1.Size
                           '                                 .TIPO = n1.Type
                           '                                 .isSys = False
                           '                              End With
                           '                              .FIELDS.Add MyField
                           '                              Set MyField = Nothing
                           '                           End If
                           '                        Next
                        End With
                        mvarTables.Add MyTab, n.NAME
                        Set MyTab = Nothing
                     End If
                  Next
               End If
               dBaseAux.Close
               WsAux.Close
               Set dBaseAux = Nothing
               Set WsAux = Nothing
         End Select
      End If
   End If
   
   Set Tables = mvarTables
   
   Exit Property
fim:
   Call ShowError
End Property
Public Property Let Conect(vData As Connection)
   Set mvarConect = vData
End Property
Public Property Set Conect(vData As Connection)
   Set mvarConect = vData
   If SetProperty Then Exit Property
   If mvarxDBs.Count = 0 Then Exit Property
End Property
Public Property Get Conect() As Connection
   Set Conect = mvarConect
End Property
Public Property Let ADOConect(vData As ADODB.Connection)
   Set mvarADOConect = vData
End Property
Public Property Set ADOConect(vData As ADODB.Connection)
   Set mvarADOConect = vData
   If SetProperty Then Exit Property
   If mvarxDBs.Count = 0 Then Exit Property
End Property
Public Property Get ADOConect() As ADODB.Connection
   Set ADOConect = mvarADOConect
End Property
Public Property Let MinutosConexao(Num%)
   mvarMinutosConexao = Num%
End Property
Public Property Get MinutosConexao() As Integer
   MinutosConexao = mvarMinutosConexao
End Property
Public Property Let CodeSql(vData As eFind)
   mvarCodeSql = vData
End Property
Public Property Get CodeSql() As eFind
   CodeSql = mvarCodeSql
End Property
Public Property Get Rs() As DAO.Recordset
   Set Rs = mvarRs
End Property
Public Property Let Rs(vData As DAO.Recordset)
   Set mvarRs = vData
End Property
Public Property Get ADORs() As ADODB.Recordset
   Set ADORs = mvarADORs
End Property
Public Property Let ADORs(vData As ADODB.Recordset)
   Set mvarADORs = vData
End Property
Public Property Get RSAux() As Object
   Set RSAux = mvarRsAux
End Property
Public Property Let RSAux(vData As Object)
   Set mvarRsAux = vData
End Property
Public Property Get Conectado() As Boolean
   Conectado = mvarConectado
End Property
Public Property Let Conectado(Bool As Boolean)
   mvarConectado = Bool
End Property
Public Property Get isODBC() As Boolean
   isODBC = mvarisODBC
End Property
Public Property Let isODBC(vData As Boolean)
   mvarisODBC = vData
End Property
Public Property Get isADO() As Boolean
   isADO = mvarisADO
End Property
Public Property Let isADO(vData As Boolean)
   mvarisADO = vData
End Property
Public Sub RetornaTransacao()
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Executar 'Rollback Transaction'.                **
   '**                                                            **
   '** Recebe: SQL$ - Query.                                      **
   '**                                                            **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   Dim Res%, Cod_Erro%, Erro_ODBC$
   On Error GoTo SqlRetBD
   Call Verifica_Estado_DB
   Select Case TypeName(DbObj)
      Case "Database"
         mvarWS.Rollback
      Case "Connection"
         If mvarisADO Then
            mvarADOConect.RollbackTrans
         Else
            mvarWsODBC.Rollback
         End If
   End Select
   Exit Sub
SqlRetBD:
   Screen.MousePointer = vbDefault
   ShowError
End Sub
Public Sub TerminaTransacao()
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Executar 'Commit Transaction'.                  **
   '**                                                            **
   '** Recebe: SQL$ - Query.                                      **
   '**                                                            **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   Dim Res%, Cod_Erro%, Erro_ODBC$
   
   On Error GoTo SqlTerT
   Call Verifica_Estado_DB
   Select Case TypeName(DbObj)
      Case "Database": mvarWS.CommitTrans
      Case "Connection"
         If mvarisADO Then
            mvarADOConect.CommitTrans
         Else
            mvarWsODBC.CommitTrans
         End If
   End Select
   mvarCodeSql = eFind.Found
   GoTo SqlTerTf
   
SqlTerT:
   
   mvarCodeSql = eFind.ERRO
   Call RetornaTransacao
   Screen.MousePointer = vbDefault
   ShowError
   Resume SqlTerTf
SqlTerTf:
   On Error GoTo 0
End Sub
Public Sub ImportarArquivo(ArquivoTexto As String, Optional Tabela_ArqSemExt)
   Dim Arq As String
   Dim nInt As Integer
   Dim TextLine As String
   Dim Campos As New Collection
   Dim Valores As New Collection
   Dim Tamanhos As New Collection
   Dim n As Variant, i As Integer
   Dim sTabela As String
   Dim Sql As String
            
   On Error Resume Next
   
'   Arq = Trim(GetPath(hWnd, ""))
'   If VBA.Right$(Arq, 1) <> "\" Then
'      Arq = Arq & "\"
'   End If
'   Arq = Arq & Tabela & ".txt"
   

   Arq = ArquivoTexto
   If FileExists(Arq) Then
      If IsMissing(Tabela_ArqSemExt) Then
         While InStr(i + 1, Arq, "\")
            i = InStr(i + 1, Arq, "\")
         Wend
         sTabela = Mid(Arq, i + 1)
         sTabela = Mid(sTabela, 1, InStr(sTabela, ".") - 1)
      Else
         sTabela = Tabela_ArqSemExt
      End If

      nInt = FreeFile
      Open Arq For Input As nInt
      Do While Not EOF(nInt)
         Line Input #nInt, TextLine
         If Campos.Count = 0 Then
            While InStr(TextLine, " ") <> 0
               Campos.Add Trim(Mid(TextLine, 1, InStr(TextLine, " ")))
               TextLine = Trim(Mid(TextLine, InStr(TextLine, " "))) & " "
               If TextLine = " " Then TextLine = ""
            Wend
         ElseIf Mid(Trim(TextLine), 1, 1) = "-" Then
            While InStr(TextLine, " ") <> 0
               Tamanhos.Add Len(Trim(Mid(TextLine, 1, InStr(TextLine, " ") - 1)))
               TextLine = Trim(Mid(TextLine, InStr(TextLine, " "))) & " "
               If TextLine = " " Then TextLine = ""
            Wend
         Else
            i = 0
            Set Valores = Nothing
            Set Valores = New Collection
            While InStr(TextLine, " ") <> 0
               i = i + 1
               Valores.Add "'" & Trim(Mid(TextLine, 1, Tamanhos(i))) & "'"
               TextLine = Trim(Mid(TextLine, Tamanhos(i) + 2)) & " "
               If TextLine = " " Then TextLine = ""
            Wend
            Sql = "insert into " & sTabela
            Sql = Sql & " ("
            For Each n In Campos
               Sql = Sql & n & ","
            Next
            Sql = Mid(Sql, 1, Len(Sql) - 1)
            Sql = Sql & ") Values ("
            For Each n In Valores
               Sql = Sql & n & ","
            Next
            Sql = Mid(Sql, 1, Len(Sql) - 1)
            Sql = Sql & ")"
            
            Call IniciaTransacao
            
            If mvarisODBC Then
               Select Case TypeName(mvarDbObj)
                  Case "Connection": mvarDbObj.Execute Sql, dbExecDirect
                  Case "Database":   mvarDbObj.Execute Sql, dbSQLPassThrough
               End Select
               DBEngine.Idle (dbFreeLocks)
            ElseIf mvarisADO Then
               Set mvarADORs = New ADODB.Recordset
               Set mvarADORs = mvarADOConect.Execute(Sql)
               'Call mvarADOConect.Execute(Vet(i%))
            Else
               mvarDbObj.Execute Sql, dbFailOnError
               DBEngine.Idle (dbFreeLocks)
            End If
            
            If Err = 0 Then
               Call TerminaTransacao
            Else
               Call RetornaTransacao
               Call GravaError(mvarADOConect.Errors(0).Description, "ImportarArquivo")
            End If
            Err = 0
         End If
      Loop
      Close nInt
   End If
End Sub
Public Sub IniciaTransacao()
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Executar 'Begin Transaction'.                   **
   '**                                                            **
   '** Recebe: SQL$ - Query.                                      **
   '**                                                            **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   Dim Res%, Erro_ODBC$, Cod_Erro%
   
   On Error GoTo SqlInicT
   Call Verifica_Estado_DB
   Select Case TypeName(DbObj)
      Case "Database"
         mvarWS.BeginTrans
      Case "Connection":
         If mvarisADO Then
            On Error Resume Next
            If mvarADORs.State = 1 Then mvarADORs.Close
            On Error GoTo SqlInicT
            mvarADOConect.BeginTrans
         Else
            mvarWsODBC.BeginTrans
         End If
   End Select
   mvarCodeSql = Found
   GoTo SqlInicTf
   
SqlInicT:
   mvarCodeSql = ERRO
   Screen.MousePointer = vbDefault
   ShowError
   Resume SqlInicTf
SqlInicTf:
   On Error GoTo 0
End Sub
Public Function ExisteChave(Tabela As String, Campo As String, Chave As String, Optional pTipoDeDados = eTipoDeDados.Caracter) As Integer
   '================================================================
   '= Última Alteração : 14/12/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Verificar a existência de uma chave em uma      **
   '**            determinada tebela.                             **
   '**                                                            **
   '** Recebe: Tabela$  - Tabela de Usuarios                      **
   '**         Campo$   - Campo da tabela referênte ao código da  **
   '**                    chave                                   **
   '**         Chave$   - Valor da chave a ser pesquisada.        **
   '**                                                            **
   '** Retorna: True se a chave existir caso contrário False      **
   '**                                                            **
   '****************************************************************
   Dim Sql$
   Dim TipoCampo As Integer
   
   Call Verifica_Estado_DB
   If Not mvarConectado Then
      Exit Function
   End If
   
   On Error GoTo fim
   mvarMinutosConexao = 0
   Sql$ = "select *"
   Sql$ = Sql$ & " from " & Tabela$
   Sql$ = Sql$ & " where " & Campo$ & " = "
   
   If IsMissing(pTipoDeDados) Then
      If mvarisADO Then
         Set mvarADORs = New ADODB.Recordset
         mvarADORs.CacheSize = 30
         Set mvarADORs = mvarADOConect.OpenSchema(adSchemaColumns)
         'dim i as Integer
         'For i = 0 To mvarADORS.FIELDS.Count - 1
         '   Debug.Print mvarADORS(i).NAME
         'Next
         Do While Not mvarADORs.EOF
            If mvarADORs("TABLE_NAME") = Tabela And mvarADORs("COLUMN_NAME") = Campo Then
               TipoCampo = mvarADORs("DATA_TYPE")
               Exit Do
            End If
            mvarADORs.MoveNext
         Loop
      Else
         TipoCampo = mvardBase.TableDefs(Tabela).FIELDS(Campo).Type
      End If
   Else
      Select Case pTipoDeDados
         Case eTipoDeDados.NUMERO: TipoCampo = dbInteger
         Case eTipoDeDados.DATAHORA: TipoCampo = dbDate
         Case eTipoDeDados.Caracter: TipoCampo = dbText
      End Select
   End If
   Select Case GrpTipoCampo(TipoCampo)
      Case eTipoDeDados.NUMERO: Sql$ = Sql$ & Chave$
      Case eTipoDeDados.DATAHORA: Sql$ = Sql$ & SqlDate(Chave$)
      Case eTipoDeDados.Caracter: Sql$ = Sql$ & SqlStr(Chave$)
   End Select
   
   Call AbreTabela(Sql$)
   
   If mvarCodeSql = Found Then
      ExisteChave% = True
   Else
      ExisteChave% = False
   End If
   Exit Function
fim:
   Call ShowError(Sql)
End Function
Public Sub CopyStructure(Orig$, dest$, Bool%)
   Dim Sql$, i%
   Dim Rs  As DAO.Recordset
   Dim Td  As DAO.TableDef
   Dim Fld As DAO.Field
   
   On Error GoTo fim
   Set Rs = mvardBase.OpenRecordset(Orig$)
   Set Td = mvardBase.CreateTableDef(dest$)
   For i = 0 To Rs.FIELDS.Count - 1
      '   For Each Fields In Rs.Fields
      Set Fld = Td.CreateField(Rs.FIELDS(i).NAME, _
      Rs.FIELDS(i).Type, _
      Rs.FIELDS(i).Size)
      Td.FIELDS.Append Fld
   Next
   mvardBase.TableDefs.Append Td
   Rs.Close
   If Bool% Then
      Sql$ = "insert into " + dest$ + " select T.* from " + Orig$ + " T"
      mvardBase.Execute Sql$, dbFailOnError
   End If
   Exit Sub
fim:
   ShowError
End Sub
Public Function CriarBD(Optional pDBase, Optional TIPO = "ACCESS") As Boolean
   '================================================================
   '= Última Alteração : 15/03/99                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '********************************************************************
   '**                                                                **
   '** OBJETIVO : Criar um Banco de Dados ACCESS                      **
   '**                                                                **
   '** Recebe: dBase$   - Caminho e Banco de Dados (Ex.:"C:\DSR\X.MDB)**
   '**         Tipo$    - Tipo de Banco de Dados                      **
   '**                    chave                                       **
   '**                                                                **
   '** Retorna: Banco de Dados Criado no caminho especificado.        **
   '**          e True/False indicando se o Banco foi criado ou não.  **
   '**                                                                **
   '********************************************************************
   'dbVersion30
   Dim NewBd As Database
   Dim pos%, Resto$, Path$
'   Dim mPointer%
   On Error GoTo fim
   mvarMinutosConexao = 0
'   mPointer% = Screen.MousePointer
'   Screen.MousePointer = vbHourglass
   
   If IsMissing(pDBase) Then pDBase = mvarDbName
   
   CriarBD = False
   '* Acaba com BD problema
   '    Kill dBase$
   '* Cria novo BD
   '* Cria caminho
   pos = InStr(pDBase, "\")
   Resto = pDBase
   While pos <> 0
      Resto = Mid(Resto, pos + 1)
      pos = InStr(Resto, "\")
   Wend
   Path$ = Mid(pDBase, 1, Len(pDBase) - Len(Resto))
   Call MakePath(Path)
   'TODO: Obsolete DAO method used. Switch to newer method
   Set NewBd = CreateDatabase(pDBase, dbLangGeneral, 0) '";LANGID=0x0809;CP=1252;COUNTRY=0", 0)
   NewBd.Close
   
   'TODO: Obsolete DAO method used. Switch to newer method
   Set NewBd = OpenDatabase(pDBase, True, False)
   NewBd.Close
   
   CriarBD = True
'   Screen.MousePointer = mPointer%
   
   Exit Function
fim:
   If Errors(0).Number = 53 Then Resume Next
   If Errors(0).Number <> 3204 Then '* Banco já Existe
      ShowError
   End If
'   Screen.MousePointer = mPointer%
End Function
Public Function CriarSqlTabela(Tabela As String, Sql As ADODB.Recordset, Optional ArrInd) As Integer
   '------------------------------------------------------------------------
   ' Funcao     : Cria_Tab
   ' Autor      : Diogenes
   ' Data       :
   ' Parametro  : cTabela - Nome da tabela a ser criada
   ' Retorno    : true/false - se criada com sucesso
   ' Obj.       : cria a tabela para impressao de relatorios do
   '              sistema.
   '------------------------------------------------------------------------
'   Dim Vet_Atrib
'   Dim Vet_Indices
   
   Dim MyRS As DAO.Recordset

   Dim i%, mPointer%
   
   Dim cTab As New DAO.TableDef
   Dim Fd() As New DAO.Field
   Dim Id() As New Index
   
   mvarMinutosConexao = 0
   CriarSqlTabela = False
   mPointer% = Screen.MousePointer
   Screen.MousePointer = vbHourglass
'   On Error GoTo fim
   
   '* Chama Rotina de deleção da tabela
'   If pAPAGAR Then
'      Call ExcluirTabela(Tabela)
'   End If
   
   cTab.NAME = Tabela
   
   ReDim Fd(0 To (Sql.FIELDS.Count - 1))
   
   '*************
   '* Cria Campos
   '*************
   '* Vetor com Duas Dimensões
   On Error GoTo fim
   For i = 0 To Sql.FIELDS.Count - 1
      Fd(i).NAME = Sql.FIELDS(i).NAME
      If Sql.FIELDS(i).Type = adChar Or Sql.FIELDS(i).Type = adVarChar Then
         Fd(i).Type = dbText
      ElseIf InArray(Sql.FIELDS(i).Type, Array(adNumeric, adInteger, vbInteger)) Then
         Fd(i).Type = dbDouble
      ElseIf Sql.FIELDS(i).Type = adLongVarChar Then
         Fd(i).Type = dbMemo
      ElseIf Sql.FIELDS(i).Type = adDBTimeStamp Or Sql.FIELDS(i).Type = adDBTime Or Sql.FIELDS(i).Type = adDBDate Then
         Fd(i).Type = dbDate
      Else
         Fd(i).Type = Sql.FIELDS(i).Type
      End If

      If Fd(i).Type = dbText Or Fd(i).Type = dbMemo Then
         Fd(i).AllowZeroLength = True
         If Fd(i).Type = dbText And Sql.FIELDS(i).DefinedSize > 255 Then
            Fd(i).Type = dbMemo
         End If
      End If
      Fd(i).Size = Sql.FIELDS(i).DefinedSize
      cTab.FIELDS.Append Fd(i)  ' Add field to c ollection.
   Next
   If Not IsMissing(ArrInd) Then
      ReDim Id(0 To IIf(UBound(ArrInd, 1) = 0, 0, (UBound(ArrInd, 1) - 1)))
      If UCase(TypeName(ArrInd(0))) = "VARIANT" Then
         For i = 0 To UBound(ArrInd, 1)
            Id(i).NAME = ArrInd(i)(0)
            Id(i).FIELDS = ArrInd(i)(1)
            If InArray(UCase(Id(i).NAME), Array("PRIMARY KEY", "PRIMARYKEY", "PRIMARY")) Then
               Id(i).Primary = True
               Id(i).Unique = True
            End If
            cTab.Indexes.Append Id(i)  ' Add index to collection.
         Next
      Else
         For i = 0 To UBound(ArrInd) Step 2
            Id(i).NAME = ArrInd(i)
            Id(i).FIELDS = ArrInd(i + 1)
            If InArray(UCase(Id(i).NAME), Array("PRIMARY KEY", "PRIMARYKEY", "PRIMARY")) Then
               Id(i).Primary = True
               Id(i).Unique = True
            End If
            cTab.Indexes.Append Id(i)  ' Add index to collection.
         Next
      End If
   End If
   '* Cria Tabela
   mvardBase.TableDefs.Append cTab
   While Not Sql.EOF
      Set MyRS = mvardBase.OpenRecordset(Tabela, dbOpenTable, dbAppendOnly)
      With MyRS
         While Not Sql.EOF
            .AddNew
            For i = 0 To Sql.FIELDS.Count - 1
               .FIELDS(i).Value = Sql.FIELDS(i).Value
            Next
            .Update
'             .Bookmark = .LastModified
            Sql.MoveNext
         Wend
      End With
      
   Wend
   CriarSqlTabela = True
   Screen.MousePointer = mPointer%
   Exit Function
fim:
   Screen.MousePointer = mPointer%
   If Err = 3010 Then '* Table already exists.
      Resume Next
   Else
      ShowError
   End If
End Function

'************************************************************************************
Public Function CriarTabela(Tabela$, Vet_Atrib(), Vet_Indices(), Optional pAPAGAR = False) As Integer
   '------------------------------------------------------------------------
   ' Funcao     : Cria_Tab
   ' Autor      : Diogenes
   ' Data       :
   ' Parametro  : cTabela - Nome da tabela a ser criada
   ' Retorno    : true/false - se criada com sucesso
   ' Obj.       : cria a tabela para impressao de relatorios do
   '              sistema.
   '------------------------------------------------------------------------
   Dim i%, mPointer%
   Dim cTab As New DAO.TableDef
   Dim Fd() As New DAO.Field
   Dim Id() As New Index
   
   mvarMinutosConexao = 0
   CriarTabela = False
   mPointer% = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   On Error GoTo fim
   
   '* Chama Rotina de deleção da tabela
   If pAPAGAR Then
      Call ExcluirTabela(Tabela)
   End If
   
   cTab.NAME = Tabela
   ReDim Fd(0 To (UBound(Vet_Atrib, 1) - 1))
   ReDim Id(0 To IIf(UBound(Vet_Indices, 1) = 0, 0, (UBound(Vet_Indices, 1) - 1)))
   '*************
   '* Cria Campos
   '*************
   On Error Resume Next
   Fd(0).NAME = Vet_Atrib(0, 1)
   If Err = 9 Then '* Subscript out of range
      '* Vetor com Uma Dimensão
      On Error GoTo fim
      Err = 0
      For i = 0 To UBound(Vet_Atrib) - 1
         Fd(i).NAME = Vet_Atrib(i)(0)
         Fd(i).Type = Vet_Atrib(i)(1)
         If InArray(Fd(i).Type, Array(dbText, dbMemo)) Then 'Text
            Fd(i).AllowZeroLength = True
         End If
         If Fd(i).Type <> dbMemo Then 'Text
            Fd(i).Size = Vet_Atrib(i)(2)
         End If
         cTab.FIELDS.Append Fd(i)  ' Add field to collection.
      Next
   Else
      '* Vetor com Duas Dimensões
      On Error GoTo fim
      For i = 0 To UBound(Vet_Atrib, 1) - 1
         Fd(i).NAME = Vet_Atrib(i, 1)
         Fd(i).Type = Vet_Atrib(i, 2)
         If InArray(Fd(i).Type, Array(dbText, dbMemo)) Then 'Text
            Fd(i).AllowZeroLength = True
         End If
         Fd(i).Size = Vet_Atrib(i, 3)
         cTab.FIELDS.Append Fd(i)  ' Add field to collection.
      Next
   End If
   '**************
   '* Cria Indices
   '**************
   If UBound(Vet_Indices, 1) > 0 Then
      On Error Resume Next
      Id(0).NAME = Vet_Indices(0, 1)
      If Err = 9 Then '* Subscript out of range
         '* Vetor com Uma Dimensão
         Err = 0
         On Error GoTo fim
         For i = 0 To UBound(Vet_Indices, 1) - 1
            Id(i).NAME = Vet_Indices(i)(0)
            Id(i).FIELDS = Vet_Indices(i)(1)
            If InArray(UCase(Id(i).NAME), Array("PRIMARY KEY", "PRIMARYKEY", "PRIMARY")) Then
               Id(i).Primary = True
               Id(i).Unique = True
            End If
            cTab.Indexes.Append Id(i)  ' Add index to collection.
         Next
      Else
         On Error GoTo fim
         For i = 0 To UBound(Vet_Indices, 1) - 1
            Id(i).NAME = Vet_Indices(i, 1)
            Id(i).FIELDS = Vet_Indices(i, 2)
            If InArray(UCase(Id(i).NAME), Array("PRIMARY KEY", "PRIMARYKEY", "PRIMARY")) Then
               Id(i).Primary = True
               Id(i).Unique = True
            End If
            cTab.Indexes.Append Id(i)  ' Add index to collection.
         Next
      End If
   End If
   
   '* Cria Tabela
   mvardBase.TableDefs.Append cTab
   CriarTabela = True
   Screen.MousePointer = mPointer%
   Exit Function
fim:
   Screen.MousePointer = mPointer%
   ShowError
End Function
Public Sub CriarIndice(cTabela As String, nNomeIndice As String, ArrCampos As Variant)
   Dim MyIndex As Index
   Dim i As Integer
   Dim mPointer As Integer
   
   On Error GoTo fim
   mvarMinutosConexao = 0
   mPointer = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   With mvardBase.TableDefs(cTabela)
   
      Set MyIndex = .CreateIndex(nNomeIndice)
      
      For i = LBound(ArrCampos) To UBound(ArrCampos)
         MyIndex.FIELDS.Append MyIndex.CreateField(ArrCampos(i))
      Next
      
      If InArray(UCase(MyIndex.NAME), Array("PRIMARY KEY", "PRIMARYKEY", "PRIMARY")) Then
         MyIndex.Primary = True
         MyIndex.Unique = True
      End If
      
      .Indexes.Append MyIndex
   End With
   
   Exit Sub
fim:
   Screen.MousePointer = mPointer
   ShowError
End Sub
Public Function ConverteSql(pSql As String, Optional pDbTipo As Integer = -1) As String
   pDbTipo = IIf(pDbTipo = -1, mvarDbTipo, pDbTipo)
   
   Select Case pDbTipo
      Case eDbTipo.MySQL
         pSql = SystaxAlias(pSql)
   End Select
   ConverteSql = pSql
End Function
Private Function SystaxAlias(pSql As String) As String
   Dim MyCol As Collection
   Dim n As Variant
   
   Set MyCol = GetWords(pSql)
   If Not MyCol Is Nothing Then
      For Each n In MyCol
         If InStr(n, "[") <> 0 Then
         
         End If
      Next
   End If
   SystaxAlias = pSql
End Function

Public Sub CloseDataBase(Optional pKeyAliasBase = "", Optional pRemoveDb = True)
   '================================================================
   '= Última Alteração : 21/07/2000                                =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Fechar Banco de Dados Corrente.                 **
   '**                                                            **
   '** Recebe: DataBase - Nome do Banco a ser fechado.            **
   '**                    parâmetro opcional se "DataBase" vazio  **
   '**                    o Banco de Dados corrente será fechado. **
   '**                                                            **
   '****************************************************************
   Dim i%, n
   Dim Bool As Boolean
   
   On Error Resume Next
   mvarRs.Close
   mvarRsAux.Close
   mvarADORs.Close
   On Error GoTo 0
   On Error GoTo SqlDesc
   If Trim(pKeyAliasBase) <> "" Then
      Set mvarDbObj = mvarxDBs(pKeyAliasBase).DbObj
   Else
      If TypeName(mvarDbObj) = "Database" Then
         pKeyAliasBase = mvarDbObj.NAME
      ElseIf TypeName(mvarDbObj) = "Connection" Then
         pKeyAliasBase = mvarDbObj.Connect
      End If
      pKeyAliasBase = mvarKey
   End If
   
   i = 0
   For Each n In mvardBases
      i = i + 1
      If n Is mvarDbObj Then
         mvardBases(i).Close
         mvardBases.Remove i
         
         If pRemoveDb Then
            mvarxDBs.Remove i
         Else
            mvarxDBs(i).Conectado = False
         End If
         Exit For
      End If
   Next
   
   Set mvarDbObj = Nothing
   Set mvardBase = Nothing
   Set mvarConect = Nothing
   Set mvarADOConect = Nothing
   
   mvarConectado = False
   i = 0
   For Each n In mvarxDBs
      i = i + 1
      If n.Conectado Then
         If TypeName(n.DbObj) = "Database" Then
            Set mvardBase = n.DbObj
         ElseIf TypeName(n.DbObj) = "Connection" Then
            Set mvarConect = n.DbObj
         End If
         Set mvarDbObj = n.DbObj
         Call Set_Property(mvarxDBs(i), Me)
         Exit For
      End If
   Next
   GoTo SqlDescf
SqlDesc:
   Call ShowError
   Resume SqlDescf
SqlDescf:
   On Error GoTo 0
End Sub
Public Sub CompactDB(Bd_Ori$, Optional ByVal Bd_Dest = "")
   '================================================================
   '= Última Alteração : 10/07/98                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Compactar Banco de Dados Access.                **
   '**                                                            **
   '** Recebe: Bd_Ori$ - Nome do arquivo de extenção "MDB" a ser  **
   '**                   Compactado                               **
   '**         Bd_Dest - Nome do novo arquivo compactado, se for  **
   '**                   omitido o novo banco receberá mesmo nome **
   '**                                                            **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   Dim Men$
   Dim lConectado As Boolean, ExcluiDestino As Boolean
   Screen.MousePointer = vbHourglass
   
   On Error GoTo fim
   
   If Bd_Dest = "" Then
      Bd_Dest = App.Path & "\Compact.mdb"
      ExcluiDestino = True
   End If
   
   '* Se Banco estiver conectado então desconecta
   lConectado = mvarConectado
   If lConectado Then Call SrvDesconecta
   
   '* Exclui arquivo destino se existir
   Call Del(CStr(Bd_Dest))
   Call Del(Mid(Bd_Dest, 1, InStr(Bd_Dest, ".")) + "LDB")
   
   '* Repara e Compacta Banco de Dados
   DBEngine.RepairDatabase Bd_Ori$
   
   DBEngine.CompactDatabase Bd_Ori$, _
   Bd_Dest, _
   dbLangGeneral, _
   dbDecrypt + dbVersion30
   '* Exclui Arquivo Origem
   Call Del(CStr(Bd_Ori$))
   Call Del(Mid(Bd_Ori$, 1, InStr(Bd_Ori$, ".")) + "LDB")
   
   '* Compacta o Banco Destino para Obter
   '* o Origem também compactado
   DBEngine.CompactDatabase Bd_Dest, _
   Bd_Ori$, _
   dbLangGeneral, _
   dbDecrypt + dbVersion30
   
   
   If ExcluiDestino Then
      Call Del(CStr(Bd_Dest))
      Call Del(Mid(Bd_Dest, 1, InStr(Bd_Dest, ".")) + "LDB")
   End If
   
   If lConectado Then Call SrvConecta(mvarDbDrive, mvarDbName, "", "", "", "")
   
   Exit Sub
fim:
   '* You attempted to open a database that is already opened
   '* exclusively by user 'admin' on machine 'TECAPLUS03'.
   '* Try again when the database is available.
   If Err = 3356 Then
      Exit Sub
   Else
      ShowError
   End If
   
End Sub
'Public Sub ConvertQuery(Sql as String , optional TipoBanco as edbtipo)
'      If mvarDbTipo = SQL_SERVER Then
'         Vet(i%) = StrReplace(UCase(Vet(i%)), " CDATE", " ")
'         Vet(i%) = StrReplace(UCase(Vet(i%)), ",CDATE", ",")
'      ElseIf mvarDbTipo = ORACLE Then
'         Vet(i%) = StrReplace(UCase(Vet(i%)), " CDATE", " TO_DATE")
'         Vet(i%) = StrReplace(UCase(Vet(i%)), ",CDATE", ",TO_DATE")
'      End If
'end sub
Public Sub AnexarTabela(DBOrig$, Tabela$, Formato$, Optional ByRef Tab_Anex = "TABANEX")
   Dim Td As New TableDef, i%
   Dim Bc As Database
   On Error GoTo fim
   
   Set Td = mvardBase.CreateTableDef(Tab_Anex)      'Nome da Tabela Destino.
   Td.Connect = Formato$ + "DATABASE=" & DBOrig$ 'Banco de Dados Origem
   Td.SourceTableName = Tabela$                  'Nome da Tabela Origem.
   
   '* Verifica se existe a tabela anexa e apaga se existir
   Call ExcluirTabela("" + Tab_Anex + "")
   mvardBase.TableDefs.Append Td
   mvardBase.Connect = True
   Exit Sub
fim:
   Call ShowError("Anexar : " & Tabela)
End Sub
Public Function ExisteReg(ByVal Sql$, Optional ByRef RecodSet_Retorno) As Integer
   Dim Res%, DyReg As Recordset
   
   On Error GoTo fim
   
   mvarMinutosConexao = 0
   ExisteReg = False
   Call AbreTabela(Sql$)
   If mvarCodeSql = Found Then
      If Not IsMissing(RecodSet_Retorno) Then
         Set RecodSet_Retorno = mvarRsAux
      End If
      '      DyReg.Close
      ExisteReg = True
   End If
   Exit Function
fim:
   Call ShowError(Sql)
End Function
Public Function LastRec(Tabela As String) As Double
   Dim RSAux As Object, Sql$
   
   If Not mvarConectado Then Exit Function
   
   On Error GoTo fim
   If Not ExisteTabela(Tabela$) Then Exit Function
   Sql$ = "select count(*) as [CONT] from " + Tabela
   If AbreTabela(Sql$, RSAux) Then
      LastRec = RSAux("CONT") & ""
      RSAux.Close
   End If
   Set RSAux = Nothing
   Exit Function
fim:
   Set RSAux = Nothing
   ShowError
End Function
Function ExisteTabela(Tabela As String)
   Dim i%
   On Error GoTo fim
   mvarMinutosConexao = 0
   ExisteTabela = False
   '   For Each TableDefs In mvarDBase.TableDefs
   If mvardBase Is Nothing Then
     ExisteTabela = True
      
   Else
      For i = 0 To mvardBase.TableDefs.Count - 1
         If UCase(mvardBase.TableDefs(i).NAME) = UCase(Tabela) Then
            ExisteTabela = True
            Exit For
         End If
      Next
   End If
   Exit Function
fim:
   Call ShowError(Tabela)
End Function
Public Sub ExcluirTabela(Tabela$)
   Dim i%
   On Error GoTo fim
   If Not mvarConectado Then Exit Sub
   
   If mvardBase Is Nothing Then
   Else
      For i = 0 To mvardBase.TableDefs.Count - 1
         If UCase(mvardBase.TableDefs(i).NAME) = UCase(Tabela$) Then
            mvardBase.TableDefs.Delete mvardBase.TableDefs(i).NAME
         End If
         If i = mvardBase.TableDefs.Count - 1 Then Exit For
      Next
   End If
   Exit Sub
fim:
   Call ShowError(Tabela)
End Sub
Public Function DECODE$(Arg$, Comp$, ResultTrue$, ResulteFalse$)
   DECODE = "iif(" & Arg$ & Comp & "," & ResultTrue$ & "," & ResulteFalse$ & ")"
End Function
Public Sub DerrubarSistema()
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '** OBJETIVO : Fechar arquivos e Desconectar Servidor.         **
   '****************************************************************
   Call SrvDesconecta
   Call FecharWorkSpace
End Sub
Public Sub FecharWorkSpace()
   Dim i%
   On Error Resume Next

   For i% = 0 To mvarWS.Databases.Count - 1
      mvarWS.Databases(i).Close
   Next
   For i = 0 To mvarWsODBC.Connections.Count - 1
      mvarWsODBC.Connections(i).Close
   Next
   
End Sub
Public Function RowCount(Optional pRecordSet As Object) As Double
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Obter o número de registros da tabela           **
   '**                                                            **
   '** Recebe: mvarRS   - Tabela('Dynaset').                        **
   '**                                                            **
   '** Retorna: ´Número de registros da tabela.                   **
   '**                                                            **
   '****************************************************************
   Dim Bm      As Variant
   Dim QtdReg  As Double
   
   On Error GoTo fim
   If TypeName(pRecordSet) = "Nothing" Then
      Set pRecordSet = mvarRs
   End If
   QtdReg = 0
   If Not pRecordSet.EOF Then
      Bm = pRecordSet.Bookmark
      pRecordSet.MoveLast
      QtdReg = pRecordSet.RecordCount
      pRecordSet.Bookmark = Bm
   End If
   RowCount = pRecordSet.RecordCount
   Exit Function
fim:
   ShowError
End Function
Public Function High(Tabela As String, Campo As String)
   Dim Sql As String
   Dim RC  As Recordset
   On Error GoTo fim
   Sql$ = "Select MAX([" + Campo$ + "]) as [HIGH] from " + Tabela$
   If AbreTabela(Sql, RC) Then
      High = RC("HIGH")
      RC.Close
   End If
   If isNull(High) Then High = 1
   If Val(High) = 0 Then High = 1
   Exit Function
fim:
   Call ShowError(Tabela)
End Function
Public Function Executa(Query As Variant, Optional ComCOMMIT = True, Optional pExibeMsg As Boolean = True) As Boolean
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Executar Query.                                 **
   '**                                                            **
   '** Recebe: SQL$ - Vetor de Query's.                           **
   '**                                                            **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   Dim i    As Integer
   Dim k    As Integer
   Dim Num  As Double
   Dim bRes As Boolean
   Dim Dsc  As String
   Dim Sql  As String
   Dim Vet()
   
   
   On Error GoTo TrataErro
   If Not mvarConectado Then Exit Function

   mvarMinutosConexao = 0
   Executa = False
   mvarCodeSql = NotFound
   
   Set mvardBase = mvardBase
   
   Select Case TypeName(Query)
      Case "String"
         If Trim(Query) = "" Then
            Executa = True
            GoTo Saida
         End If
         ReDim Vet(1)
         Vet(0) = Query
         '       For i = LBound(Query) To UBound(Query)
         '           If Len(Trim(Query(i))) = 0 Then Exit Sub
         '        Next
      Case "Collection"
         '* Redimensiona Vetor Eliminando os Elementos vazios
         If Query.Count = 0 Then
            Executa = True
            GoTo Saida
         End If
         i = Query.Count
         ReDim Vet(i)
         k = -1
         For i = 1 To Query.Count
            If Len(Trim(Query(i))) > 0 Then
               k = k + 1
               Vet(k) = Query(i)
            End If
         Next
         If k < 0 Then Exit Function
         ReDim Preserve Vet(k)
      Case Else
         '* Redimensiona Vetor Eliminando os Elementos vazios
         If UBound(Query) = 0 And Len(Trim(Query(LBound(Query)))) = 0 Then
            Executa = True
            GoTo Saida
         End If
         
         i = UBound(Query) + 1
         ReDim Vet(i)
         k = -1
         For i = LBound(Query) To UBound(Query)
            If Len(Trim(Query(i))) > 0 Then
               k = k + 1
               Vet(k) = Query(i)
            End If
         Next
         If k < 0 Then Exit Function
         ReDim Preserve Vet(k)
   End Select
   If ComCOMMIT Then
      Call IniciaTransacao
   End If
   mvarCodeSql = Found
   
   Call SET_DATEFORMAT("dmy")
   
   For i% = LBound(Vet) To UBound(Vet)
      'Call ConvertQuery(Vet(i), mvarDbTipo)
      Sql = Vet(i%)
      '**********
      '* Tratar Caracter aspas simples (')
      If mvarDbTipo = Access Or mvarDbTipo = ORACLE Then
         While InStr(Sql, " + CHAR(39)  + ") <> 0
            Sql = StrReplace(Sql, " + CHAR(39)  + ", " + ""'"" + ")
         Wend
      End If
      
      If Trim(Sql) <> "" Then
         If mvarisODBC Then
            Select Case TypeName(mvarDbObj)
               Case "Connection": mvarDbObj.Execute Sql, dbExecDirect
               Case "Database":   mvarDbObj.Execute Sql, dbSQLPassThrough
            End Select
            DBEngine.Idle (dbFreeLocks)
         ElseIf mvarisADO Then
            Set mvarADORs = New ADODB.Recordset
            Set mvarADORs = mvarADOConect.Execute(Sql)
            'Call mvarADOConect.Execute(Sql)
         Else
            mvarDbObj.Execute Sql, dbFailOnError
            DBEngine.Idle (dbFreeLocks)
         End If
         
         bRes = True
         
         If Err <> 0 Then bRes = False
         mvarCodeSql = IIf(bRes, mvarCodeSql, eFind.NotFound)
         If Not mvarCodeSql Then Exit For
      End If
   Next
   If Not bRes Then
      mvarCodeSql = NotFound
      If ComCOMMIT Then
         Call RetornaTransacao
      End If
   Else
      mvarCodeSql = Found
      Executa = True
      If ComCOMMIT Then
         Call TerminaTransacao
      End If
   End If
   Executa = (mvarCodeSql = Found)
   GoTo Saida
   '"Attempt to insert duplicate key"
TrataErro:
   '   Tit$ = LoadMsg(1)
   On Error Resume Next
   Executa = False
   If ComCOMMIT Then
      Call RetornaTransacao
   End If
   If mvarisADO Then
      Num = mvarADOConect.Errors(0).Number
      Dsc = mvarADOConect.Errors(0).Description
   Else
      Num = Errors(0).Number
      Dsc = Errors(0).Description
   End If
   mvarCodeSql = ERRO
   Err = Num
   Screen.MousePointer = vbDefault
   If mvarisADO Then
      'i = mvarADOConect.Errors.Count - 1
      If Trim(Dsc) = "" Then
         Dsc = mvarADOConect.Errors(0).Description
      Else
         Dsc = Dsc & vbNewLine & mvarADOConect.Errors(0).Description
      End If
      If Trim(Dsc) = "" Then Dsc = Error
      Dsc = Dsc & vbNewLine & Sql
      
      Call ShowError(Dsc, pExibeMsg)
      'Call ExibirStop(Dsc, "ERRO")
   Else
      If Num <> 0 Then
         Call ShowError(Vet(i), pExibeMsg)
      End If
   End If

   Resume Saida
Saida:
   On Error GoTo 0
End Function
Public Function AbreTabela(ByVal pSql As String, Optional ByRef RcSet As Variant, Optional pExibeMsg As Boolean = True) As Boolean
   '===================================================================
   '= Última Alteração : 14/01/98                                     =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)              =
   '===================================================================
   '*******************************************************************
   '**                                                               **
   '** OBJETIVO : Executar Query.                                    **
   '**                                                               **
   '** Recebe: pSQL  - Query.                                        **
   '**         RcSet - Nome do RecordSet que será criado(Default=RS) **
   '**                                                               **
   '** Retorna: RecordSet                                            **
   '**                                                               **
   '*******************************************************************
   
   Dim Res        As Integer
   Dim Erro_ODBC  As String
   Dim Cod_Erro   As Integer
   Dim Miss       As Integer
   Dim bConnTeste As Boolean
   
   On Error GoTo Trata_Error

   bConnTeste = False
   mvarMinutosConexao = 0
   
   If Trim(pSql) = "" Then Exit Function
   
   Set mvardBase = mvardBase
   Miss = False
   If IsMissing(RcSet) Then
      If mvarisADO Then
         On Error Resume Next
         If mvarADORs.State = 1 Then mvarADORs.Close
         On Error GoTo Trata_Error
         Set mvarADORs = New ADODB.Recordset
         Set RcSet = mvarADORs
      Else
         Set RcSet = mvarRs
      End If
      Miss = True
   End If
   
   Call Verifica_Estado_DB
   If Not mvarConectado Then
      Exit Function
   End If
   
   Call SET_DATEFORMAT("dmy")
   pSql = ConverteSql(pSql)
   
   If mvarisODBC Then
      If TypeName(mvarDbObj) = "Connection" Then
         Set RcSet = mvarDbObj.OpenRecordset(pSql, adOpenStatic, dbExecDirect)
      ElseIf TypeName(mvarDbObj) = "Database" Then
         Set RcSet = mvarDbObj.OpenRecordset(pSql, adOpenStatic, dbSQLPassThrough)
      End If
   ElseIf mvarisADO Then
      On Error GoTo Trata_ADO_Error
      mvarADOConect.Errors.Clear
      Set mvarADORs = New ADODB.Recordset
      mvarADORs.LockType = adLockOptimistic
      mvarADORs.CursorLocation = adUseClient
      mvarADORs.CursorType = adOpenStatic

      Set RcSet = New ADODB.Recordset
      RcSet.LockType = adLockOptimistic
      RcSet.CursorLocation = adUseClient
      RcSet.CursorType = adOpenStatic
      RcSet.Open pSql, mvarADOConect, adOpenStatic, adLockOptimistic
'      Set RcSet = mvarADOConect.Execute(pSql)

      RcSet.MoveFirst
   Else
      Set RcSet = mvardBase.OpenRecordset(pSql, adOpenStatic, , LockEdit:=adLockOptimistic)
      'Set RcSet = mvardBase.OpenRecordset(pSql)
      DBEngine.Idle dbRefreshCache
      DBEngine.Idle dbFreeLocks
   End If
   If Not RcSet.EOF Then
      mvarCodeSql = eFind.Found
   Else
      mvarCodeSql = eFind.NotFound
   End If
   If Miss Then
      If mvarisADO Then
         Set mvarADORs = RcSet
      Else
         Set mvarRs = RcSet
      End If
   End If
   Set mvarRsAux = RcSet
   AbreTabela = (mvarCodeSql = eFind.Found)
   GoTo SqlInif
   
Trata_Error:
   mvarCodeSql = eFind.ERRO
   AbreTabela = (mvarCodeSql = eFind.Found)
   Screen.MousePointer = vbDefault
   Call ShowError(pSql, pExibeMsg)
   Resume SqlInif

Trata_ADO_Error:
   If mvarADOConect.Errors.Count <> 0 Then
      If pExibeMsg Then
         MsgBox "ERRO" & vbNewLine & "Description - " & mvarADOConect.Errors(0).Description & vbNewLine & "Number - " & CStr(mvarADOConect.Errors(0).Number)
      End If
      'Resume Next
      mvarCodeSql = eFind.NotFound
      AbreTabela = False
   ElseIf mvarADOConect.Errors.Count = 0 Then
      'Set RcSet = New ADODB.Recordset
      'Set RcSet = Nothing
      mvarCodeSql = eFind.NotFound
      Set mvarRsAux = RcSet
      AbreTabela = (mvarCodeSql = eFind.Found)
   End If
SqlInif:
   On Error GoTo 0
End Function
Public Sub ShowDbError(Sql)
   If mvarADOConect.Errors.Count <> 0 Then
    
   Else
      Call ShowError(Sql)
   End If
End Sub
Public Sub SqlToTable(Sql$, Tabela$, ByRef DbDest As DS_BANCO, Optional ByVal pIndices)
   Dim RcToTable As DAO.Recordset
   Dim Atributos(), Indices()
   Dim i%, j%, Campos$, Valores$
   On Error GoTo fim
   Call AbreTabela(Sql$, RcToTable)
   If mvarCodeSql <> Found Then
      DbDest.ExcluirTabela (Tabela$)
      '      If DbDest.DExisteTabela(Tabela) Then
      '         DbDest.Executa ("delete * from " & Tabela)
      '      End If
      Exit Sub
   End If
   ReDim Atributos(RcToTable.FIELDS.Count, 3)
   For i = 0 To RcToTable.FIELDS.Count - 1
      Campos = Campos & IIf(i = 0, "", ",") & RcToTable.FIELDS(i).NAME
      Atributos(i, 1) = CStr(RcToTable.FIELDS(i).NAME)
      Atributos(i, 2) = CStr(RcToTable.FIELDS(i).Type)
      Atributos(i, 3) = CStr(RcToTable.FIELDS(i).Size)
   Next
   If IsMissing(pIndices) Then
      ReDim Indices(0)
   Else
      ReDim Indices(UBound(pIndices), 2)
      For i = LBound(pIndices) To UBound(pIndices) - 1
         For j = 1 To 2
            Indices(i, j) = pIndices(i, j)
         Next
      Next
   End If
   
   Call DbDest.CriarTabela(Tabela, Atributos(), Indices(), True)
   RcToTable.MoveFirst
   While Not RcToTable.EOF
      Sql = "INSERT INTO " & Tabela$ & "(" & Campos & ")"
      Sql = Sql & " values ("
      Valores = ""
      For i = 0 To RcToTable.FIELDS.Count - 1
         Valores = Valores & IIf(i = 0, "", ", ")
         Select Case GrpTipoCampo(RcToTable.FIELDS(i).Type)
            Case eTipoDeDados.NUMERO
               If RcToTable.FIELDS(i).Value & "" = "" Then
                  Valores = Valores & "0"
               Else
                  Valores = Valores & CStr(RcToTable.FIELDS(i).Value & "")
               End If
            Case eTipoDeDados.Caracter
               Valores = Valores & SqlStr(RcToTable.FIELDS(i).Value & "")
            Case eTipoDeDados.DATAHORA
               Valores = Valores & SqlDate(CStr(RcToTable.FIELDS(i).Value & ""))
         End Select
      Next
      Sql = Sql & Valores & ")"
      Call DbDest.Executa(Sql)
      RcToTable.MoveNext
   Wend
   Exit Sub
fim:
   Call ShowError(Sql)
End Sub
Public Sub SrvConecta(Optional pDbDrive, Optional pDbName, Optional pDSN, Optional pUID, Optional pPWD, Optional pStrDataBase, Optional pAlias, Optional pSERVER, Optional ExibeErro = True, Optional bDrive51 As Boolean = True, Optional bRemote As Boolean = False)
   '**********************************************************************
   '* Desenvolvedor    : Diogenes S. Ramos
   '* Data / Hora      : 10/02/2001 23:19:03
   '* Nome do Módulo   : DS_BANCO (DS_BANCO.cls)
   '* Nome da Função   : SrvConecta
   '* Parâmetros       :
   '* Comentário       :
   '*
   '*
   '**********************************************************************
   Dim mPointer As Integer, i As Integer
   Dim StrName As String
   Dim Filtro As String
   Dim Con As Connection
   Dim n, JaExiste As Boolean
   Dim Reconectar As Boolean
   Dim MyDb As DS_BANCO
   
   Set MyDb = Nothing
   Set MyDb = New DS_BANCO
   
'   If mvarConectado Then
'      If vbNo = ExibirPergunta("Baco de Dados já está conectado." & vbNewLine & vbNewLine & "Deseja Reconectar?", "Acesso") Then
'         Exit Sub
'      End If
'   End If
   Reconectar = (xDBs.Count = 0)
   For Each n In xDBs
      If n.Conectado = False And n.Key = mvarKey Then
         Reconectar = True
         Exit For
      End If
   Next
   
   If MyDb.dbTipo < 0 And Not Reconectar Then
      MyDb.DSNDrive = mvarDSNDrive
      If InStr(LCase(mvarDSNDrive), "oracle") <> 0 Then
         MyDb.dbTipo = eDbTipo.ORACLE
      ElseIf InStr(LCase(mvarDSNDrive), "sql") <> 0 Then
         MyDb.dbTipo = eDbTipo.SQL_SERVER
      ElseIf InStr(LCase(mvarDSNDrive), "access") <> 0 Then
         MyDb.dbTipo = eDbTipo.Access
      ElseIf InStr(LCase(mvarDSNDrive), "mysql") <> 0 Then
         MyDb.dbTipo = eDbTipo.MySQL
      End If
   End If
   
   If Reconectar Or Not mvarConectado Then
      Call Set_Property(Me, MyDb)
   Else
      Call Set_Property(MyDb, Me)
   End If
   
   If Not IsMissing(pDbDrive) Then
      If Right(pDbDrive, 1) <> "\" Then pDbDrive = pDbDrive & "\"
      mvarDbDrive = pDbDrive
   End If
   
   If Not IsMissing(pAlias) Then mvarAlias = pAlias
   If Not IsMissing(pDbName) Then mvarDbName = pDbName
   If Not IsMissing(pDSN) Then mvarDSN = pDSN  'UNB01_32
   If Not IsMissing(pUID) Then mvarUID = pUID  'USU_VERIF - ORDSR
   If Not IsMissing(pPWD) Then mvarPWD = pPWD  'USU_VERIF - P678694694
   If Not IsMissing(pStrDataBase) Then mvarStrDATABASE = pStrDataBase
   If Not IsMissing(pSERVER) Then mvarServer = ServerName(pSERVER)
   
   On Error GoTo SqlCon
   
   Set mvardBase = mvardBase
   Set mvarConect = mvarConect
   Set mvarADOConect = mvarADOConect
   Set mvarDbObj = mvarDbObj
   
   mPointer% = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   'mvarisODBC = (mvarDSN <> "")
   
   If mvarisODBC Then
      StrName = ""
      If Not IsMissing(pDSN) Then mvarDSN = pDSN                '* UNB01_32
      If Not IsMissing(pUID) Then mvarUID = pUID                '* USU_VERIF - ORDSR
      If Not IsMissing(pPWD) Then mvarPWD = pPWD                '* USU_VERIF - P678694694
      If Not IsMissing(pStrDataBase) Then mvarStrDATABASE = pStrDataBase '* RIO_TST
      mvarStrConect$ = "ODBC;DSN=" & mvarDSN & ";UID=" & mvarUID & ";PWD=" & mvarPWD
      If Trim(mvarStrDATABASE) <> "" Then
         mvarStrConect$ = mvarStrConect$ & ";DATABASE=" & mvarStrDATABASE
      End If
      Set mvarConect = WsODBC.OpenConnection("WsODBC", dbDriverNoPrompt, , mvarStrConect$)
      Set mvarDbObj = mvarConect
      
   ElseIf mvarisADO Then
      ' Open connection.
       Select Case mvarDbTipo
          Case SQL_SERVER
             mvarStrConect$ = "Provider=SQLOLEDB;"
             mvarStrConect$ = mvarStrConect$ & "Initial Catalog=" & mvarDbName & ";"
             mvarStrConect$ = mvarStrConect$ & "Data Source=" & mvarServer & ";"
             mvarStrConect$ = mvarStrConect$ & "User Id=" & mvarUID & ";"
             mvarStrConect$ = mvarStrConect$ & "Password=" & mvarPWD & ";"
             
   mvarStrConect$ = mvarStrConect$ & "Use Encryption for Data=False;"
   mvarStrConect$ = mvarStrConect$ & "Tag with column collation when possible=False;"
'008156
          Case ORACLE
             mvarStrConect$ = "Provider=MSDAORA.1;"
             mvarStrConect$ = mvarStrConect$ & "Data Source=" & mvarServer & ";"
             mvarStrConect$ = mvarStrConect$ & "User Id=" & mvarUID & ";"
             mvarStrConect$ = mvarStrConect$ & "Password=" & mvarPWD & ";"
             
          Case Access
             mvarStrConect$ = "Provider=Microsoft.Jet.OLEDB.4.0;"
             mvarStrConect$ = mvarStrConect$ & "Data Source=" & mvarDbDrive & mvarDbName
             
          Case MySQL
'            mvarDbTipo = 3
'            mvarDbName = "g3r"
'            mvarServer = "localhost"
'            mvarUID = "root"
'            mvarPWD = ""
'            bDrive51 = True
            
            If bDrive51 Then
               mvarStrConect$ = "DRIVER={MySQL ODBC 5.1 Driver};"
            Else
               mvarStrConect$ = "DRIVER={MySQL};"
            End If
            mvarStrConect$ = mvarStrConect$ & "Server=" & mvarServer & ";"
            mvarStrConect$ = mvarStrConect$ & "Database=" & mvarDbName & ";"
            mvarStrConect$ = mvarStrConect$ & "User=" & mvarUID & ";"
            mvarStrConect$ = mvarStrConect$ & "Password=" & mvarPWD & ";"
            mvarStrConect$ = mvarStrConect$ & "Port=3306;"
            mvarStrConect$ = mvarStrConect$ & "STT=;"
            mvarStrConect$ = mvarStrConect$ & "Option=2048;"
            
'            If bDrive351 Then
'               mvarStrConect$ = mvarStrConect$ & "Option=3;"
'            Else
               'mvarStrConect$ = mvarStrConect$ & "Option=2048;"
'               If bRemote Then
               '   mvarStrConect$ = mvarStrConect$ & "Option=131072;"
                  
'               Else
               '   mvarStrConect$ = mvarStrConect$ & "Option=16834;"
'               End If
'            End If
'          "DRIVER={MySQL};Database=Lista;Server=192.168.0.1;Port=3306;Uid=root;Pwd=Senha;STT=;Option=2048;"
'          "DRIVER={MySQL ODBC 3.51 Driver};Database=Lista;Server=192.168.0.1;Port=3306;Uid=root;Pwd=Senha;STMT=;Opton=2048;"
'             mvarStrConect$ = "DRIVER={MySQL ODBC 3.51 Driver};"
'             mvarStrConect$ = mvarStrConect$ & "Database=" & mvarDbName & ";"
'             mvarStrConect$ = mvarStrConect$ & "Server=" & mvarServer & ";"
'             mvarStrConect$ = mvarStrConect$ & "Port=3306;"
'             mvarStrConect$ = mvarStrConect$ & "Uid=" & mvarUID & ";"
'             mvarStrConect$ = mvarStrConect$ & "Pwd=" & mvarPWD & ";"
'             mvarStrConect$ = mvarStrConect$ & "STMT=;Option=2048;"
          
'             mvarStrConect$ = "Provider=SQLOLEDB;"
'             mvarStrConect$ = mvarStrConect$ & "Initial Catalog=" & mvarDbName & ";"
'             mvarStrConect$ = mvarStrConect$ & "Data Source=" & mvarServer & ";"
'             mvarStrConect$ = mvarStrConect$ & "User Id=" & mvarUID & ";"
'             mvarStrConect$ = mvarStrConect$ & "Password=" & mvarPWD & ";"
             
'MyODBC 2.50 Local database:
'"Driver={mySQL};Server=localhost;Option=16834;Database=mydatabase;"
'MyODBC 3.51 Local database:
'DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=myDatabase;USER=myUsername;PASSWORD=myPassword;OPTION=3;"

'MyODBC 2.50 Remote database:
'"Driver={mySQL};Server=data.domain.com;Port=3306;Option=131072;Stmt=;Database=my-database;Uid=myUsername;Pwd=myPassword;"
'MyODBC 3.51 Remote database:
'"DRIVER={MySQL ODBC 3.51 Driver};SERVER=data.domain.com;PORT=3306;DATABASE=myDatabase; USER=myUsername;PASSWORD=myPassword;OPTION=3;"
'Driver={MySQL ODBC 5.1 Driver}
'       mvarStrConect$ = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=g3r;User=root;Password=;Option=3;"
'       mvarStrConect$ = "Driver={MySQL ODBC 5.1 Driver};Server=localhost;Database=g3r;Uid=root;Pwd=;Port=3306;STMT=;Option=2048;"
'       mvarStrConect$ = "Driver={MySQL ODBC 5.1 Driver};Server=localhost;Database=g3r;User=root;Password=;Port=3306;STMT=;Option=2048;"
       
       End Select
       
       
       Set mvarADOConect = New ADODB.Connection
       mvarADOConect.CommandTimeout = mvarTimeOut
       mvarADOConect.CursorLocation = adUseClient
       mvarADOConect.ConnectionString = mvarStrConect$
       mvarADOConect.Open
       
       Set mvarDbObj = mvarADOConect
       
   ElseIf mvarDbTipo = Access Then
         mvarStrConect$ = ";"
         StrName = mvarDbDrive & mvarDbName
         Set mvardBase = WS.OpenDatabase(StrName, False, False, mvarStrConect$)
         Set mvarDbObj = mvardBase
   Else
      MsgBox "DataBase Type not defined."
      Exit Sub
   End If
   
   Set mvarTables = Nothing
   '* Adiciona Database na Coleção
   For Each n In mvardBases
      If n Is mvarDbObj Then
         JaExiste = True
         Exit For
      End If
   Next
   
   If mvarisODBC Then
      mvarKey = Trim(mvarDbObj.Connect)
   ElseIf mvarisADO Then
      If mvarDbTipo = Access Then
         mvarKey = Trim(mvarDbObj.Properties("Data Source Name").Value)
      ElseIf mvarDbTipo = SQL_SERVER Or mvarDbTipo = ORACLE Or mvarDbTipo = MySQL Then
         mvarKey = Trim(mvarDbObj.Properties("Current Catalog").Value)
      End If
   Else
     mvarKey = Trim(mvarDbObj.NAME)
   End If
   
   On Error Resume Next
   mvarAlias = IIf(mvarAlias = "", mvarKey, mvarAlias)
   If Not JaExiste Then
      mvardBases.Add mvarDbObj, mvarAlias
   End If
   
   On Error GoTo SqlCon
   If mvarisODBC Then Call GetODBCConnectParts(mvarDbObj.Connect)
   mvarConectado = True
   If mvarisADO Then
      mvarDbObj.CommandTimeout = mvarTimeOut
   Else
      mvarDbObj.QueryTimeout = mvarTimeOut
   End If
   '* Adiciona DB
   For Each n In mvarxDBs
      If n.Key = Me.Key Then JaExiste = True
      If n.Key = mvarAlias Then JaExiste = True
   Next
   If Not JaExiste Then
      Set MyDb = Nothing
      Set MyDb = New DS_BANCO
      Call Set_Property(Me, MyDb)
      mvarxDBs.Add MyDb, mvarAlias
   Else
      i = 0
      For Each n In mvarxDBs
         i = i + 1
         If n.Key = mvarKey Then
            Call Set_Property(Me, mvarxDBs(i))
            Exit For
         End If
      Next
   End If
   
   Set MyDb = Nothing
   Screen.MousePointer = mPointer%
   GoTo SqlConf
SqlCon:
   
   mvarCodeSql = ERRO
   Do While Err = 3024
      Filtro$ = "Access Files (*.mdb)|*.mdb|Report (*.rpt)|*.rpt"
      Exit Do
      '      If ProcurarArquivo(SysMdi.CmDialog, "Find File", UCase(mvardbDrive$ & mvardbName), Filtro$, 1) <> "" Then
      '         mvardbDrive = UCase(SysMdi.CmDialog.Tag)
      '         mvardbName = UCase(SysMdi.CmDialog.FileTitle)
      '         Set dBase = WS.OpenDatabase(mvardbDrive$ & mvardbName, False, False, conect$)
      '       Else
      '         Exit Do
      '       End If
   Loop
   If Err = 0 Then
      mvardBase.QueryTimeout = mvarTimeOut
      Screen.MousePointer = mPointer%
      GoTo SqlConf
   Else
      If ExibeErro Then
         ShowError
      End If
   End If
   Resume SqlConf
SqlConf:
   On Error GoTo 0
End Sub
Public Sub SET_DATEFORMAT(Formato As String)
   Dim Sql As String
   
   If mvarDbTipo = SQL_SERVER Then
      Sql = "SET DATEFORMAT " & Formato
      If mvarisODBC Then
         Select Case TypeName(mvarDbObj)
            Case "Connection": mvarDbObj.Execute Sql, dbExecDirect
            Case "Database":   mvarDbObj.Execute Sql, dbSQLPassThrough
         End Select
         DBEngine.Idle (dbFreeLocks)
      ElseIf mvarisADO Then
         Set mvarADORs = New ADODB.Recordset
         Set mvarADORs = mvarADOConect.Execute(Sql)
         'Call mvarADOConect.Execute(Sql)
      Else
         mvarDbObj.Execute Sql, dbFailOnError
         DBEngine.Idle (dbFreeLocks)
      End If
'      If mvarisADO Then
'         Sql = ""
'         For i% = LBound(Vet) To UBound(Vet)
'            Sql = Vet(i%) & " " & vbNewLine
'         Next
'         Set mvarADORs = New ADODB.Recordset
'         Set mvarADORs = mvarADOConect.Execute(Sql)
'      End If
   End If
End Sub
Private Sub Set_Property(DbOri As DS_BANCO, DbDest As DS_BANCO)
   SetProperty = True
   With DbDest
      .Alias = DbOri.Alias
      .dbTipo = DbOri.dbTipo
      .CodeSql = DbOri.CodeSql
      .Conectado = DbOri.Conectado
      If Not DbOri.ADOConect Is Nothing Then
         Set .ADOConect = DbOri.ADOConect
      End If
      If Not DbOri.Conect Is Nothing Then
         Set .Conect = DbOri.Conect
      End If
      If Not DbOri.dBase Is Nothing Then
         Set .dBase = DbOri.dBase
      End If
      '      Set .dBases = DbOri.dBases
      .dbDrive = DbOri.dbDrive
      .dbName = DbOri.dbName
      Set .DbObj = DbOri.DbObj
      .isODBC = DbOri.isODBC
      '      Set .xDBs = DbOri.xDBs
      .Driver = DbOri.Driver
      .DSN = DbOri.DSN
      .DSNDrive = DbOri.DSNDrive
      .Rs = DbOri.Rs
      .RSAux = DbOri.RSAux
      .Key = DbOri.Key
      .MinutosConexao = DbOri.MinutosConexao
      .PWD = DbOri.PWD
      .Server = DbOri.Server
      .StrConect = DbOri.StrConect
      .StrDATABASE = DbOri.StrDATABASE
      If Not mvarTables Is Nothing Then
         If DbOri.Tables Is Nothing Then
            Set .Tables = Nothing
         Else
            Set .Tables = DbOri.Tables
         End If
      End If
      .UID = DbOri.UID
      .isADO = DbOri.isADO
   End With
SetProperty = False
End Sub
Public Sub SrvDesconecta()
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Desfazer conexão com o servidor.                **
   '**                                                            **
   '** Recebe: SQL$ - Query.                                      **
   '**                                                            **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   Dim i%
   Dim n As Variant
   
   On Error Resume Next
   If mvarConectado Then
      If Not mvarRs Is Nothing Then mvarRs.Close
      If Not mvarRsAux Is Nothing Then mvarRsAux.Close
      If Not mvarADORs Is Nothing Then mvarADORs.Close
   End If
   Set mvarRs = Nothing
   Set mvarRsAux = Nothing
   Set mvarADORs = Nothing
   
   mvarADOConect.Close
   
   On Error GoTo SqlDesc
    

   Set mvardBase = mvardBase
   '   If mvarConectado Then
   For Each n In xDBs
      For i = 1 To n.dBases.Count
         '         If mvarDBases(i).Name = mvardBase.Name Then
         n.dBases(i).Close
         '             mvarDBases.Remove i
         '            Exit For
         '         End If
      Next
      Set n.dBases = Nothing
      Set n.dBase = Nothing
      n.Conectado = False
   Next
   mvarConectado = False
   '   End If
   GoTo SqlDescf
SqlDesc:
   
   Call ShowError
   Resume SqlDescf
SqlDescf:
   On Error GoTo 0
End Sub
Public Function SqlDate(ByVal DT As String, Optional Format_Date As eSysDate, Optional InsereNull As Boolean = True) As String
   Dim MyFormat As String
   Dim pos As Integer
   Dim sData As String
   Dim dData As Date
   Dim sAux As String
   Dim nTipoAux As Integer
   
   On Error GoTo TrataErro
'* em Sql 103 - dd/mm/yy e 101 - mm/dd/yy
   If Val(DT) = 0 Then
      SqlDate = IIf(InsereNull, "Null", "")
   Else
      nTipoAux = mvarDbTipo
      nTipoAux = IIf(nTipoAux = eDbTipo.MySQL, eDbTipo.SQL_SERVER, nTipoAux)
      
      Select Case nTipoAux
         Case eDbTipo.Access:     SqlDate = "CDATE('" & Format$(DT, mvarDtMask) & "')"
         Case eDbTipo.SQL_SERVER
            'mvarDtMask = "mm/dd/yyyy"
            Select Case Format_Date
               Case eSysDate.Data:      MyFormat = mvarDtMask
               Case eSysDate.Data_Hora: MyFormat = mvarDtMask & " hh:mm:ss"
               Case eSysDate.HORA:      MyFormat = "hh:mm:ss"
               Case Else:               MyFormat = mvarDtMask
            End Select
            DT = Trim(DT)
            sData = StrZero(Mid(DT, 1, InStr(DT, "/") - 1), 2)
            pos = InStr(DT, "/") + 1
            If InStr(pos, DT, "/") - pos >= 0 Then
               sData = sData & "/" & StrZero(Mid(DT, pos, InStr(pos, DT, "/") - pos), 2)
            End If
            pos = InStr(pos, DT, "/") + 1
            
            sAux = DT
            If InStr(DT, " ") <> 0 Then
               sAux = Mid(DT, 1, InStr(DT, " ") - 1)
            End If
            If Len(Mid(sAux, pos)) = 2 Then
               sData = sData & "/" & VBA.Right$(Mid(sAux, pos), 2)
            ElseIf Len(Mid(sAux, pos)) = 4 Then
               If Mid(Mid(sAux, pos), 1, 2) > 18 Then
                  sData = sData & "/" & Mid(sAux, pos)
               Else
                  sData = sData & "/" & VBA.Right$(Mid(sAux, pos), 2)
               End If
            Else
               sData = sData & "/" & VBA.Right$(Mid(sAux, pos), 2)
            End If
            dData = VBA.CDate(sData)
            sData = ""
            If Format_Date = eSysDate.Data Then
               sData = StrZero(Day(dData), 2) & "/" & StrZero(Month(dData), 2) & "/" & Year(dData)
            ElseIf Format_Date = eSysDate.HORA Then
               sData = StrZero(Hour(DT), 2) & ":" & StrZero(Minute(DT), 2) & ":" & StrZero(Second(DT), 2)
            ElseIf Format_Date = eSysDate.Data_Hora Then
               sData = StrZero(Day(dData), 2) & "/" & StrZero(Month(dData), 2) & "/" & Year(dData)
               sData = sData & " " & StrZero(Hour(DT), 2) & ":" & StrZero(Minute(DT), 2) & ":" & StrZero(Second(DT), 2)
            Else
               sData = StrZero(Day(dData), 2) & "/" & StrZero(Month(dData), 2) & "/" & Year(dData)
            End If
            
            
            'SqlDate = " CONVERT(DATETIME, '" & Format$(DT, MyFormat) & "', 103) "
            SqlDate = "CAST('" & sData & "' As DATETIME) "

         Case eDbTipo.ORACLE:     SqlDate = "TO_DATE('" & Format$(DT, mvarDtMask) & "')"
         'Case Else:               SqlDate = "CDATE('" & Format$(DT, mvarDtMask) & "')"
      End Select
   End If
Exit Function
TrataErro:
   Call ExibirStop("Formato de Data Inválido", "Atenção")
End Function
Function SqlNum(ByVal Num As String, Optional InsereNull As Boolean = False) As String
   If InsereNull And xVal(Num) = 0 Then
      SqlNum = "Null"
   Else
      If Trim(Num$) = "" Then
          SqlNum = "0"
      Else
          'SqlNum = StrReplace(Tiraponto(Num$), ",", ".")
          Num$ = xVal(Num$)
          SqlNum = StrReplace(Tiraponto(Num$), ",", ".")
      End If
   End If
End Function
Public Function SqlStr(ByVal Txt As String, Optional InsereNull As Boolean = False) As String
   Dim sAux_Antes  As String
   Dim sAux_Depois As String
   
   If InsereNull And Trim(Txt) = "" Then
      SqlStr = "Null"
   Else
      If InStr(Txt, "'") <> 0 Then
         sAux_Antes = Mid(Txt, 1, InStr(Txt, "'") - 1)
         sAux_Depois = Mid(Txt, InStr(Txt, "'") + 1)
         If mvarDbTipo = SQL_SERVER Or mvarDbTipo = MySQL Then
            Txt = SqlStr(sAux_Antes) & " + CHAR(39)  + " & SqlStr(sAux_Depois)
         ElseIf mvarDbTipo = ORACLE Then
            Txt = SqlStr(sAux_Antes) & " + ""'"" + " & SqlStr(sAux_Depois)
         ElseIf mvarDbTipo = Access Then
            Txt = SqlStr(sAux_Antes) & " + ""'"" + " & SqlStr(sAux_Depois)
         Else
            Txt = SqlStr(sAux_Antes) & " + ""'"" + " & SqlStr(sAux_Depois)
         End If
         SqlStr = Txt
      Else
         SqlStr = "'" & Txt & "'"
      End If
   End If
End Function
Public Function Verifica_Estado_DB() As Boolean
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Verificar se banco de dados está desconectado,  **
   '**            se estiver ele conecta.                         **
   '**                                                            **
   '** Recebe:                                                    **
   '**                                                            **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   'se o banco de dados estiver desconectado, conecta
   Dim Aux
   On Error Resume Next
   Set mvardBase = mvardBase
   
   If mvarDbTipo = Access Then
      Aux = dBase.NAME
   ElseIf mvarDbTipo = SQL_SERVER Then
   
   ElseIf mvarDbTipo = ORACLE Then
   
   End If
   
'   For i = 0 To mvarDbObj.Properties.Count - 1
'      Debug.Print mvarDbObj.Properties(i).NAME & " = " & mvarDbObj.Properties(i).Value
'   Next

   If mvarisADO Then
      mvarConectado = (mvarADOConect.State = adStateOpen)
      mvarConectado = mvarConectado And (mvarADOConect.Properties("Connection Status") = adStateOpen)
   End If
   
   
   'mvarConectado = (Err = 0)
   If Not mvarConectado Then
      Call SrvConecta(mvarDbDrive$, mvarDbName$, mvarDSN$, mvarUID$, mvarPWD$, mvarStrDATABASE$, mvarAlias, mvarServer)
   End If
   
   Verifica_Estado_DB = mvarConectado
   mvarMinutosConexao = 0
   
End Function
Public Function Sysdate(Optional ByVal Data_Ou_Hora As eSysDate = Data) As String
   '----------------------------------------------------------
   ' Função     : SYSDATE
   ' Autor      : Diogenes
   ' Data       : 14/12/98
   ' Objetivo   : Retornar a data do dia do servidor ( SYSDATE )
   ' Parametros : lHora - Lógico - True - Retorna data
   '                             - False - Retorna hora
   '
   ' Retorno    : Data/Hora do Dia
   '----------------------------------------------------------
   Dim Sql$, Mascara$
   Dim Dy_Date As Object
      
   On Error GoTo fim
   mvarMinutosConexao = 0
   Screen.MousePointer = vbHourglass
   
   If IsMissing(Data_Ou_Hora) Then Data_Ou_Hora = 0

   If Data_Ou_Hora = Data Then
      Mascara = "dd/mm/yyyy"
   ElseIf Data_Ou_Hora = HORA Then
      Mascara = "hh:mm:ss"
   ElseIf Data_Ou_Hora = Data_Hora Then
      Mascara = "dd/mm/yyyy hh:mm:ss"
   Else
      Mascara = "dd/mm/yyyy"
   End If

   Select Case mvarDbTipo
      Case eDbTipo.SQL_SERVER
         Sql = "SELECT GETDATE()"
         Call AbreTabela(Sql$, Dy_Date)
         Sysdate = Format(Dy_Date(0), Mascara$)
'         Dy_Date.Close
      Case eDbTipo.ORACLE
         Sql$ = "SELECT SYSDATE FROM DUAL"
         Call AbreTabela(Sql$, Dy_Date)
         Sysdate = Format(Dy_Date(0), Mascara$)
'         Dy_Date.Close
      Case eDbTipo.Access
         Sysdate = Format$(Now, Mascara$)
      Case eDbTipo.MySQL
         Sql = "SELECT Now()"
         Call AbreTabela(Sql$, Dy_Date)
         Sysdate = Format(Dy_Date(0), Mascara$)
   End Select
   If mvarDbTipo <> eDbTipo.Access Then
      If Data_Ou_Hora = Data Then
         Sysdate = StrZero(Day(Dy_Date(0)), 2) & "/" & StrZero(Month(Dy_Date(0)), 2) & "/" & Year(Dy_Date(0))
      ElseIf Data_Ou_Hora = HORA Then
         Sysdate = StrZero(Hour(Dy_Date(0)), 2) & ":" & StrZero(Minute(Dy_Date(0)), 2) & ":" & StrZero(Second(Dy_Date(0)), 2)
      ElseIf Data_Ou_Hora = Data_Hora Then
         Sysdate = StrZero(Day(Dy_Date(0)), 2) & "/" & StrZero(Month(Dy_Date(0)), 2) & "/" & Year(Dy_Date(0))
         Sysdate = Sysdate & " " & StrZero(Hour(Dy_Date(0)), 2) & ":" & StrZero(Minute(Dy_Date(0)), 2) & ":" & StrZero(Second(Dy_Date(0)), 2)
      Else
         Sysdate = StrZero(Day(Dy_Date(0)), 2) & "/" & StrZero(Month(Dy_Date(0)), 2) & "/" & Year(Dy_Date(0))
      End If
   End If
   Screen.MousePointer = vbDefault
   Exit Function
fim:
   Set Dy_Date = Nothing
   ShowError
   Screen.MousePointer = vbDefault
End Function
Public Sub GetOwners(ByRef DB As Object, ByRef Owners As Collection)
   '================================================================
   '= Última Alteração : 28/12/99                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Recuparar todos os Owners do Banco de Dados.    **
   '**                                                            **
   '** Recebe: DB     - Objeto de Banco de Dados ( DS_BANCO ).    **
   '**         Owners - Coleção de Todos os Owners                **
   '**                                                            **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   Dim i%, Txt$, pos%
   Dim Sql$, n As Variant
   Dim MyRc As Recordset
   Txt$ = Chr(255)
   With DB.dBase
      On Error Resume Next
      Set Owners = New Collection
      If mvarDbTipo = eDbTipo.ORACLE Then
         Sql$ = "select distinct OWNER "
         Sql$ = Sql$ & " from ALL_TABLES"
         Call AbreTabela(Sql$, MyRc)
         If Not MyRc.EOF Then
            MyRc.MoveFirst
            While Not MyRc.EOF
               Owners.Add MyRc("OWNER") & "", MyRc("OWNER") & ""
               MyRc.MoveNext
            Wend
         End If
         GoTo Saida
      End If
      
      
      '      For Each n In Tables
      '         Owners.Add Mid(n.NOME, 1, InStr(n.NOME, ".") - 1), Mid(n.NOME, 1, InStr(n.NOME, ".") - 1)
      '      Next
      For i = 0 To .TableDefs.Count - 1
         pos% = InStr(.TableDefs(i).NAME, ".")
         If Txt <> Mid(.TableDefs(i).NAME, 1, IIf(pos = 0, 1, pos) - 1) Then
            If InStr(.TableDefs(i).NAME, ".") <> 0 Then
               Owners.Add Mid(.TableDefs(i).NAME, 1, InStr(.TableDefs(i).NAME, ".") - 1), Mid(.TableDefs(i).NAME, 1, InStr(.TableDefs(i).NAME, ".") - 1)
            Else
               Owners.Add " ", " "
            End If
         End If
         Txt = Mid(.TableDefs(i).NAME, 1, IIf(pos = 0, 1, pos) - 1)
         'Debug.Print .TableDefs(i).Name
         If i + 2 < .TableDefs.Count - 1 Then
            If Txt = Mid(.TableDefs(i + 2).NAME, 1, IIf(pos = 0, 1, pos) - 1) Then
               i = i + 2
            End If
         End If
      Next
   End With
Saida:
   Set MyRc = Nothing
End Sub
Public Function GravarTabela%(Sql$, Oper%, Optional Confirm = "")
   '================================================================
   '= Última Alteração : 28/11/97                                  =
   '= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
   '================================================================
   '****************************************************************
   '**                                                            **
   '** OBJETIVO : Gravar alterações na tabela.                    **
   '**                                                            **
   '** Recebe: SQL$   - Query de atualização da tabela.           **
   '**         Oper$  - Operação de atualização                   **
   '**                  (INCLUSÃO, ALTERAÇÃO, EXCLUSÃO)           **
   '**         Confirm - Título da Confirmação de atualização da  **
   '**                  tabela (Opcional)                         **
   '** Retorna:                                                   **
   '**                                                            **
   '****************************************************************
   Dim Resp%, mPointer%
   Dim Gresp%
   
   On Error GoTo fim
   mvarMinutosConexao = 0
   Resp% = vbNo
   If Confirm = "" Then Confirm = LoadMsg(1)
   If Confirm = "NOTSHOW" Then Resp% = vbYes
   mPointer% = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   If Oper% = eINCLUSAO Then
      DoEvents
      Gresp% = False
      Call Executa(Sql$)
      If mvarCodeSql <> Found Then
         Call ExibirAviso(LoadMsg(2), CStr(Confirm))
      Else
         Gresp% = True
      End If
   End If
   If Oper% = eALTERACAO Then
      If Resp% = vbNo Then Resp% = ExibirPergunta(LoadMsg(3), CStr(Confirm))
      Gresp% = Resp%
      DoEvents
      If Resp% = vbYes Then
         Call Executa(Sql$)
         If mvarCodeSql = NotFound Then
            Call ExibirStop(LoadMsg(4), CStr(Confirm))
         End If
      End If
   End If
   If Oper% = eEXCLUSAO Then
      If Resp% = vbNo Then Resp% = ExibirPergunta(LoadMsg(5), CStr(Confirm))
      DoEvents
      Gresp% = Resp%
      If Resp% = vbYes Then
         Call Executa(Sql$)
         If mvarCodeSql = NotFound Then
            Call ExibirStop(LoadMsg(6), CStr(Confirm))
         End If
         Gresp% = vbYes
      Else
         Gresp% = vbNo
      End If
   End If
   GravarTabela = Gresp
   Screen.MousePointer = mPointer%
   Exit Function
fim:
   Call ShowError(Sql)
End Function
Public Function GrpTipoCampo(Tip As Integer) As eTipoDeDados
   Select Case Tip
      Case dbBigInt, dbBinary, dbBoolean, dbByte, dbCurrency, dbDecimal, _
         dbDouble, dbInteger, dbLong, dbNumeric, dbSingle, dbLongBinary, _
         dbGUID, dbVarBinary, adNumeric, adSmallInt, adSingle, adInteger, _
         adDouble, adDecimal
         GrpTipoCampo = 1 '* Numérico
         
      Case dbDate, dbTime, dbTimeStamp, adDBTimeStamp
         GrpTipoCampo = 2 '* Data
         
      Case dbChar, dbMemo, dbText, adChar, adVarChar, adWChar, adVarWChar
         GrpTipoCampo = 3 '* Caracter
      Case adBinary, adGUID, adIDispatch, adIUnknown
         GrpTipoCampo = 4 '* object
   End Select
End Function
Public Sub MergeTabelas(ByRef DbAux As Database, TabOrig$, TabDest$, Chave$)
   Dim Sql        As String
   Dim nCampo     As String
   Dim nSetCampos As String
   Dim i          As Integer
   Dim j          As Integer
   Dim Tmp        As Recordset
   
   On Error GoTo fim
   
   Sql$ = "Insert into " & TabDest
   Sql$ = Sql$ & " select distinctrow " + TabOrig + ".* "
   Sql$ = Sql$ & " From " + TabOrig
   Sql$ = Sql$ & " where 0=(select count(*) from " + TabDest
   Sql$ = Sql$ & " where " + Chave$ + ")"
   DbAux.Execute Sql$, dbFailOnError
   
   Set Tmp = DbAux.OpenRecordset(TabDest$)
   j = 0
   For i = 0 To Tmp.FIELDS.Count - 1
      nCampo = Tmp.FIELDS(i).NAME
      If Tmp.FIELDS(i).Attributes <> 49 Then
         If j = 0 Then
            j = 1
            nSetCampos = nSetCampos & TabDest$ + "." & nCampo & " = " + TabOrig$ + "." & nCampo
         Else
            nSetCampos = nSetCampos & " , " + TabDest$ + "." & nCampo & " = " + TabOrig$ + "." & nCampo
         End If
      End If
   Next
   Sql$ = "update " + TabDest$ + " inner join " + TabOrig$ + " on "
   Sql$ = Sql$ + Chave$
   Sql$ = Sql$ + " set " + nSetCampos
   DbAux.Execute Sql$, dbFailOnError
   Exit Sub
fim:
   ShowError
End Sub
Public Sub GetODBCConnectParts(rsConnect As String)
   On Error Resume Next
   
   Dim i As Integer
   Dim sTmp As String
   
   'process the connect string just in case the
   'values came from the ODBC dialogs
   If InStr(rsConnect, "=") Then
      i = 1
      While i <= Len(rsConnect) + 1
         If Mid(rsConnect, i, 1) = ";" Or i = Len(rsConnect) + 1 Then
            If Len(sTmp) > 0 And InStr(sTmp, "=") > 0 Then
               Select Case UCase(Mid(sTmp, 1, InStr(1, sTmp, "=") - 1))
                  Case "DSN"
                     mvarDSN = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
                  Case "DATABASE"
                     mvarStrDATABASE = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
                  Case "DBQ"
                     mvarStrDATABASE = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
                  Case "UID"
                     mvarUID = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
                  Case "PWD"
                     mvarPWD = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
                  Case "DRIVER"
                     mvarDriver = Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp))
                  Case "SERVER"
                     mvarServer = ServerName(Mid(sTmp, InStr(1, sTmp, "=") + 1, Len(sTmp)))
                  Case Else
                     'nothing
               End Select
            End If
            sTmp = vbNullString
         Else
            sTmp = sTmp + Mid(rsConnect, i, 1)
         End If
         i = i + 1
      Wend
   End If
End Sub
Public Sub FileToBinary(ByVal Arquivo As String, Optional Conexao As Object, Optional pFld As Object, Optional Tabela As String, Optional Campo As String, Optional Chave As String, Optional Threshold As Long = 1048576)
' Assumes file exists
' Assumes calling routine does the UPDATE
' File cannot exceed approx. 2Gb in size
   Dim Fld     As ADODB.Field
   Dim Rs         As ADODB.Recordset
   Dim f          As Long
   Dim Data()     As Byte
   Dim FileSize   As Long
   Dim Sql        As String
   Dim LimpaRs    As Boolean
   
   If IsMissing(pFld) Or pFld Is Nothing Then
      LimpaRs = True
      Sql = "Select * From " & Tabela & " Where " & Chave
      Set Rs = New ADODB.Recordset
      If UCase(TypeName(Conexao)) = "DS_BANCO" Then
         Rs.Open Sql, Conexao.ADOConect, adOpenKeyset, adLockOptimistic
      Else
         Rs.Open Sql, Conexao, adOpenKeyset, adLockOptimistic
      End If
      Set Fld = Rs(Campo)
      If Arquivo = "" Then
         Fld.Value = Null
         Rs.Update
         Rs.Close
         Exit Sub
      End If
   Else
      Set Fld = pFld
   End If
   
   f = FreeFile
   Open Arquivo For Binary As #f
      FileSize = LOF(f)
      Select Case Fld.Type
         Case adBinary
            If FileSize > Threshold Then
               ReadToBinary f, Fld, FileSize
            Else
               Data = InputB(FileSize, f)
               Fld.Value = Data
            End If
         
         Case adLongVarBinary
            If FileSize > Threshold Then
               ReadToBinary f, Fld, FileSize
            Else
               Data = InputB(FileSize, f)
               Fld.Value = Data
            End If
         Case adLongVarChar, adLongVarWChar
            If FileSize > Threshold Then
               ReadToText f, Fld, FileSize
            Else
               Fld.Value = Input(FileSize, f)
            End If
      End Select
   Close #f
   If LimpaRs Then
      Rs.Update
      Rs.Close
   End If
End Sub
Public Sub BinaryToFile(Optional pConexao As Object, Optional pFld As Object, Optional Arquivo As String, Optional ByVal Tabela As String, Optional Campo As String, Optional Chave As String, Optional FieldSize As Long = -1, Optional Threshold As Long = 1048576)
' Assumes file does not exist
' Data cannot exceed approx. 2Gb in size
   
   Dim Fld     As ADODB.Field
   Dim Rs      As ADODB.Recordset
   Dim f       As Long
   Dim bData() As Byte
   Dim sData   As String
   Dim LimpaRs As Boolean
   Dim Sql     As String
   
   If IsMissing(pConexao) Then Set pConexao = Me
   If IsMissing(pFld) Or pFld Is Nothing Then
      LimpaRs = True
      Sql = "Select * From " & Tabela & " Where " & Chave
      Set Rs = New ADODB.Recordset
      If UCase(TypeName(pConexao)) = "DS_BANCO" Then
         Rs.Open Sql, pConexao.ADOConect, adOpenStatic, adLockReadOnly
      Else
         Rs.Open Sql, pConexao, adOpenStatic, adLockReadOnly
      End If
      Set Fld = Rs(Campo)
   Else
      Set Fld = pFld
   End If
   
   f = FreeFile
   Open Arquivo For Binary As #f
      Select Case Fld.Type
         Case adBinary
            If FieldSize = -1 Then   ' blob field is of unknown size
               WriteFromUnsizedBinary f, Fld
            Else                     ' blob field is of known size
               If FieldSize > Threshold Then   ' very large actual data
                  WriteFromBinary f, Fld, FieldSize
               Else                            ' smallish actual data
                  bData = Fld.Value
                  Put #f, , bData  ' PUT tacks on overhead if use fld.Value
               End If
            End If
         
         Case adLongVarBinary
            If FieldSize = -1 Then   ' blob field is of unknown size
               WriteFromUnsizedBinary f, Fld
            Else                     ' blob field is of known size
               If FieldSize > Threshold Then   ' very large actual data
                  WriteFromBinary f, Fld, FieldSize
               Else                            ' smallish actual data
                  bData = Fld.Value
                  Put #f, , bData  ' PUT tacks on overhead if use fld.Value
               End If
            End If
         Case adLongVarChar, adLongVarWChar
            If FieldSize = -1 Then
               WriteFromUnsizedText f, Fld
            Else
               If FieldSize > Threshold Then
                  WriteFromText f, Fld, FieldSize
               Else
                  sData = Fld.Value
                  Put #f, , sData  ' PUT tacks on overhead if use fld.Value
               End If
            End If
      End Select
   Close #f
   If LimpaRs Then
      Rs.Close
   End If
End Sub
Public Sub WriteFromBinary(ByVal f As Long, Fld As ADODB.Field, ByVal FieldSize As Long)
   Dim Data() As Byte
   Dim BytesRead As Long
   
   Do While FieldSize <> BytesRead
      If FieldSize - BytesRead < BLOCK_SIZE Then
         Data = Fld.GetChunk(FieldSize - BLOCK_SIZE)
         BytesRead = FieldSize
      Else
         Data = Fld.GetChunk(BLOCK_SIZE)
         BytesRead = BytesRead + BLOCK_SIZE
      End If
      Put #f, , Data
   Loop
End Sub
Public Sub WriteFromUnsizedBinary(ByVal f As Long, Fld As ADODB.Field)
   Dim Data() As Byte
   Dim Temp As Variant
   Do
      Temp = Fld.GetChunk(BLOCK_SIZE)
      If isNull(Temp) Then Exit Do
      Data = Temp
      Put #f, , Data
   Loop While LenB(Temp) = BLOCK_SIZE
End Sub
Public Sub WriteFromText(ByVal f As Long, Fld As ADODB.Field, ByVal FieldSize As Long)
   Dim Data As String
   Dim CharsRead As Long
   
   Do While FieldSize <> CharsRead
      If FieldSize - CharsRead < BLOCK_SIZE Then
         Data = Fld.GetChunk(FieldSize - BLOCK_SIZE)
         CharsRead = FieldSize
      Else
         Data = Fld.GetChunk(BLOCK_SIZE)
         CharsRead = CharsRead + BLOCK_SIZE
      End If
      Put #f, , Data
   Loop
End Sub
Public Sub WriteFromUnsizedText(ByVal f As Long, Fld As ADODB.Field)
   Dim Data As String, Temp As Variant
   Do
      Temp = Fld.GetChunk(BLOCK_SIZE)
      If isNull(Temp) Then Exit Do
      Data = Temp
      Put #f, , Data
   Loop While Len(Temp) = BLOCK_SIZE
End Sub
Public Sub ReadToBinary(ByVal f As Long, Fld As ADODB.Field, ByVal FileSize As Long)
   Dim Data() As Byte
   Dim BytesRead As Long
   Do While FileSize <> BytesRead
      If FileSize - BytesRead < BLOCK_SIZE Then
        Data = InputB(FileSize - BytesRead, f)
        BytesRead = FileSize
      Else
        Data = InputB(BLOCK_SIZE, f)
        BytesRead = BytesRead + BLOCK_SIZE
      End If
      Fld.AppendChunk Data
   Loop
End Sub
Public Sub ReadToText(ByVal f As Long, Fld As ADODB.Field, ByVal FileSize As Long)
   Dim Data As String
   Dim CharsRead As Long
   
   Do While FileSize <> CharsRead
      If FileSize - CharsRead < BLOCK_SIZE Then
         Data = Input(FileSize - CharsRead, f)
         CharsRead = FileSize
      Else
         Data = Input(BLOCK_SIZE, f)
         CharsRead = CharsRead + BLOCK_SIZE
      End If
      Fld.AppendChunk Data
   Loop
End Sub
Private Sub Class_Initialize()
   mvarisODBC = False
   mvarisADO = True
   mvarConectado = False
   mvarMinutosConexao = 0
   mvarDSNDrive = "Banco de Dados MS Access"
   mvarDbTipo = GetSetting("DSR", "General Format", "dbTipo", 1)

   'mvarDbTipo = ACCESS
   mvarDbVersao = "3.0"
   mvarTimeOut = 300
   
   '****************************
   '* Propriedades para suprir *
   '* a ausência de DSR100.DLL *
   '****************************
   SepDec = ","
   SepMil = "."
   SepDt = "/"
   DtMask = "dd" + mvarSepDt + "mm" + mvarSepDt + "yyyy"
   DrvTmp = "C:\TMP\"
   DrvErro = mvarDrvTmp
   Idioma = PORTUGUES '* 5000
   mvarComputerName = Environ("ComputerName")
   mvarRemoteComputerName = "mssql.classeaconsultoria.com.br"
End Sub
Private Sub Class_Terminate()
   Dim n
   On Error Resume Next
   If Not mvardBases Is Nothing Then
      For Each n In mvardBases
         n.Close
      Next
   End If
   Set mvardBases = Nothing
   Set mvarConections = Nothing
   Set mvarADOConections = Nothing
   Set mvarTables = Nothing
End Sub

