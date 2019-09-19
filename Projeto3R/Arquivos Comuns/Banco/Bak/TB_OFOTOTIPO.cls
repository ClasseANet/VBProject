VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OFOTOTIPO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb         As Object 
Private mvarRS          As Object 

Private mvarQryInsert   As String
Private mvarQryUpDate   As String
Private mvarQryDelete   As String
Private mvarQrySelect   As String
Private mvarQrySave     As String

Private mvarisDirt      As Boolean

Private mvarFOTOTIPO   As Double
Private mvarPELE       As String
Private mvarSARDAS     As String
Private mvarPELO       As String
Private mvarOLHOS      As String
Private mvarDEFINICAO  As String
Private mvarQUEIMADURA As String
Private mvarBRONZEADO  As String
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let FOTOTIPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFOTOTIPO <> vData)
   mvarFOTOTIPO = vData
End Property
Public Property Get FOTOTIPO() As Double
   FOTOTIPO = mvarFOTOTIPO
End Property
Public Property Let PELE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPELE <> vData)
   mvarPELE = vData
End Property
Public Property Get PELE() As String
   PELE = mvarPELE
End Property
Public Property Let SARDAS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSARDAS <> vData)
   mvarSARDAS = vData
End Property
Public Property Get SARDAS() As String
   SARDAS = mvarSARDAS
End Property
Public Property Let PELO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPELO <> vData)
   mvarPELO = vData
End Property
Public Property Get PELO() As String
   PELO = mvarPELO
End Property
Public Property Let OLHOS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOLHOS <> vData)
   mvarOLHOS = vData
End Property
Public Property Get OLHOS() As String
   OLHOS = mvarOLHOS
End Property
Public Property Let DEFINICAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDEFINICAO <> vData)
   mvarDEFINICAO = vData
End Property
Public Property Get DEFINICAO() As String
   DEFINICAO = mvarDEFINICAO
End Property
Public Property Let QUEIMADURA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarQUEIMADURA <> vData)
   mvarQUEIMADURA = vData
End Property
Public Property Get QUEIMADURA() As String
   QUEIMADURA = mvarQUEIMADURA
End Property
Public Property Let BRONZEADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarBRONZEADO <> vData)
   mvarBRONZEADO = vData
End Property
Public Property Get BRONZEADO() As String
   BRONZEADO = mvarBRONZEADO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OFOTOTIPO Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " FOTOTIPO     = " & SqlNum(mvarFOTOTIPO) & VbNewLine 
   If Not mvarXDb Is Nothing Then
      If mvarXDb.Conectado Then
         If mvarXDb.Executa(Sql, True) Then
            mvarALTERSTAMP = vData
         End If
      End If
   End If
End Property
Public Property Get ALTERSTAMP() As Double
   ALTERSTAMP = mvarALTERSTAMP
End Property
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
End Property
Public Property Get isDirt() As Boolean
   isDirt = mvarisDirt
End Property
Public Property Set XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Let XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Get XDb() As Object
   Set XDb = mvarXDb
End Property
Public Property Get RS() As Object
   Set RS = mvarRS
End Property
Public Property Get QryInsert(Optional pAutoId as Boolean = True) As String
   Dim Sql As String
 
   Sql = "Insert Into OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO" & VbNewLine 
   Sql = Sql & ", OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarFOTOTIPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPELE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSARDAS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPELO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOLHOS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDEFINICAO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarQUEIMADURA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarBRONZEADO) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_FOTOTIPO) As String
   Dim Sql As String
 
   Sql = "Delete From OFOTOTIPO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_FOTOTIPO) Then Sql = Sql & " FOTOTIPO = " & SqlNum(Cstr(Ch_FOTOTIPO)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   MyDelRow.Query = Sql & ";"
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update OFOTOTIPO Set " & VbNewLine 
   Sql = Sql & " FOTOTIPO   = " & SqlNum(mvarFOTOTIPO) & VbNewLine 
   Sql = Sql & " , PELE       = " & SqlStr(mvarPELE) & VbNewLine 
   Sql = Sql & " , SARDAS     = " & SqlStr(mvarSARDAS) & VbNewLine 
   Sql = Sql & " , PELO       = " & SqlStr(mvarPELO) & VbNewLine 
   Sql = Sql & " , OLHOS      = " & SqlStr(mvarOLHOS) & VbNewLine 
   Sql = Sql & " , DEFINICAO  = " & SqlStr(mvarDEFINICAO) & VbNewLine 
   Sql = Sql & " , QUEIMADURA = " & SqlStr(mvarQUEIMADURA) & VbNewLine 
   Sql = Sql & " , BRONZEADO  = " & SqlStr(mvarBRONZEADO) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP  = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OFOTOTIPO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " FOTOTIPO     = " & SqlNum(mvarFOTOTIPO) & VbNewLine 

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId as Boolean = True) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_FOTOTIPO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select FOTOTIPO, PELE, SARDAS, PELO" & VbNewLine 
   Sql = Sql & ", OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OFOTOTIPO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_FOTOTIPO)  Then 
      Sql = Sql & " FOTOTIPO = " & SqlNum(mvarFOTOTIPO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_FOTOTIPO) Then Sql = Sql & " FOTOTIPO = " & SqlNum(Cstr(Ch_FOTOTIPO)) & " AND " & VbNewLine 
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & VbNewLine
      End If
   End If 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY 

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar( Optional Ch_FOTOTIPO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_FOTOTIPO, Ch_WHERE, Ch_ORDERBY)
   bExiste = mvarXDb.AbreTabela(Sql, mvarRS)
   mvarisDirt = False
   With mvarRS
      If bExiste Then bExiste = Not .EOF
      If bExiste Then
         Me.Popula
         Pesquisar = True
      Else
         Pesquisar = False
      End If
   End With
   Exit Function
PesquisarErr:
    call ShowError(Sql)
    Pesquisar = False
End Function
Public Sub Popula(Optional pRcSet)
   If IsMissing(pRcSet) Then Set pRcSet = mvarRS
   With pRcSet
      mvarFOTOTIPO = XVal(!FOTOTIPO & "")
      mvarPELE = !PELE & ""
      mvarSARDAS = !SARDAS & ""
      mvarPELO = !PELO & ""
      mvarOLHOS = !OLHOS & ""
      mvarDEFINICAO = !DEFINICAO & ""
      mvarQUEIMADURA = !QUEIMADURA & ""
      mvarBRONZEADO = !BRONZEADO & ""
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True, Optional pAutoId as Boolean = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId as Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarFOTOTIPO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_FOTOTIPO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OFOTOTIPO Set "
   Sql = Sql & " FOTOTIPO = " & SqlNum(Ch_FOTOTIPO)
   Sql = Sql & " Where 
   Sql = Sql & " FOTOTIPO = " & SqlNum(mvarFOTOTIPO)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
