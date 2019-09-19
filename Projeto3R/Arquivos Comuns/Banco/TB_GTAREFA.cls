VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GTAREFA"
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

Private mvarIDTAREFA  As Double
Private mvarNOME      As String
Private mvarDESCRICAO As String
Private mvarCODSIS    As String
Private mvarCONEXAO   As String
Private mvarSQL       As String
Private mvarSCRIPT    As String
Public Property Get IDTAREFA() As Double
   IDTAREFA = mvarIDTAREFA
End Property
Public Property Let NOME(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNOME <> vData)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Public Property Let DESCRICAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDESCRICAO <> vData)
   mvarDESCRICAO = vData
End Property
Public Property Get DESCRICAO() As String
   DESCRICAO = mvarDESCRICAO
End Property
Public Property Let CODSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIS <> vData)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let CONEXAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCONEXAO <> vData)
   mvarCONEXAO = vData
End Property
Public Property Get CONEXAO() As String
   CONEXAO = mvarCONEXAO
End Property
Public Property Let SQL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSQL <> vData)
   mvarSQL = vData
End Property
Public Property Get SQL() As String
   SQL = mvarSQL
End Property
Public Property Let SCRIPT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSCRIPT <> vData)
   mvarSCRIPT = vData
End Property
Public Property Get SCRIPT() As String
   SCRIPT = mvarSCRIPT
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into GTAREFA (NOME, DESCRICAO, CODSIS" & VbNewLine 
   Sql = Sql & ", CONEXAO, SQL, SCRIPT) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDESCRICAO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCONEXAO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSQL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSCRIPT) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDTAREFA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From GTAREFA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDTAREFA)  Then 
      Sql = Sql & " IDTAREFA = " & SqlNum(mvarIDTAREFA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDTAREFA) Then Sql = Sql & " IDTAREFA = " & SqlNum(Cstr(Ch_IDTAREFA)) & " AND " & VbNewLine 
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
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update GTAREFA Set " & VbNewLine 
   Sql = Sql & " NOME      = " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & " , DESCRICAO = " & SqlStr(mvarDESCRICAO) & VbNewLine 
   Sql = Sql & " , CODSIS    = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , CONEXAO   = " & SqlStr(mvarCONEXAO) & VbNewLine 
   Sql = Sql & " , SQL       = " & SqlStr(mvarSQL) & VbNewLine 
   Sql = Sql & " , SCRIPT    = " & SqlStr(mvarSCRIPT) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GTAREFA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDTAREFA    = " & SqlNum(mvarIDTAREFA) & VbNewLine 

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate(pSinc:=pSinc)
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId, pSinc:=pSinc)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_IDTAREFA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDTAREFA, NOME, DESCRICAO, CODSIS" & VbNewLine 
   Sql = Sql & ", CONEXAO, SQL, SCRIPT" & VbNewLine 
   Sql = Sql &" From GTAREFA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDTAREFA)  Then 
      Sql = Sql & " IDTAREFA = " & SqlNum(mvarIDTAREFA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDTAREFA) Then Sql = Sql & " IDTAREFA = " & SqlNum(Cstr(Ch_IDTAREFA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDTAREFA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDTAREFA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDTAREFA = XVal(!IDTAREFA & "")
      mvarNOME = !NOME & ""
      mvarDESCRICAO = !DESCRICAO & ""
      mvarCODSIS = !CODSIS & ""
      mvarCONEXAO = !CONEXAO & ""
      mvarSQL = !SQL & ""
      mvarSCRIPT = !SCRIPT & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDTAREFA = 0
   mvarNOME = ""
   mvarDESCRICAO = ""
   mvarCODSIS = ""
   mvarCONEXAO = ""
   mvarSQL = ""
   mvarSCRIPT = ""

   On Error Resume Next
   Call Class_Initialize
End Sub
Public Function Salvar(Optional ComCOMMIT As Boolean = True, Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId, pSinc:=pSinc), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId as Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDTAREFA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Private Sub Class_Initialize()
   Set mvarRS = Nothing
   Set mvarXDb = Nothing
   mvarQryInsert = ""
   mvarQryUpDate = ""
   mvarQryDelete = ""
   mvarQrySelect = ""
   mvarQrySave = ""
   mvarisDirt = False
End Sub
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
