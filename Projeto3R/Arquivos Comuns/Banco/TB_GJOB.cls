VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GJOB"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb            As Object 
Private mvarRS             As Object 

Private mvarQryInsert      As String
Private mvarQryUpDate      As String
Private mvarQryDelete      As String
Private mvarQrySelect      As String
Private mvarQrySave        As String

Private mvarisDirt         As Boolean

Private mvarIDJOB        As Double
Private mvarCODSIS       As String
Private mvarNOME         As String
Private mvarDSCJOB       As String
Private mvarDTINICIO     As String
Private mvarDTFIM        As String
Private mvarHORAEXECUCAO As String
Private mvarDURACAO      As String
Private mvarDIASSEMANA   As String
Public Property Get IDJOB() As Double
   IDJOB = mvarIDJOB
End Property
Public Property Let CODSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIS <> vData)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let NOME(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNOME <> vData)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Public Property Let DSCJOB(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCJOB <> vData)
   mvarDSCJOB = vData
End Property
Public Property Get DSCJOB() As String
   DSCJOB = mvarDSCJOB
End Property
Public Property Let DTINICIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTINICIO <> vData)
   mvarDTINICIO = vData
End Property
Public Property Get DTINICIO() As String
   DTINICIO = mvarDTINICIO
End Property
Public Property Let DTFIM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTFIM <> vData)
   mvarDTFIM = vData
End Property
Public Property Get DTFIM() As String
   DTFIM = mvarDTFIM
End Property
Public Property Let HORAEXECUCAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHORAEXECUCAO <> vData)
   mvarHORAEXECUCAO = vData
End Property
Public Property Get HORAEXECUCAO() As String
   HORAEXECUCAO = mvarHORAEXECUCAO
End Property
Public Property Let DURACAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDURACAO <> vData)
   mvarDURACAO = vData
End Property
Public Property Get DURACAO() As String
   DURACAO = mvarDURACAO
End Property
Public Property Let DIASSEMANA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDIASSEMANA <> vData)
   mvarDIASSEMANA = vData
End Property
Public Property Get DIASSEMANA() As String
   DIASSEMANA = mvarDIASSEMANA
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
 
   Sql = "Insert Into GJOB (CODSIS, NOME, DSCJOB" & VbNewLine 
   Sql = Sql & ", DTINICIO, DTFIM, HORAEXECUCAO, DURACAO, DIASSEMANA) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCJOB) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTINICIO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarHORAEXECUCAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDURACAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDIASSEMANA) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From GJOB" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update GJOB Set " & VbNewLine 
   Sql = Sql & " CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , NOME         = " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & " , DSCJOB       = " & SqlStr(mvarDSCJOB) & VbNewLine 
   Sql = Sql & " , DTINICIO     = " & SqlDate(mvarDTINICIO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTFIM        = " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , HORAEXECUCAO = " & SqlDate(mvarHORAEXECUCAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DURACAO      = " & SqlDate(mvarDURACAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DIASSEMANA   = " & SqlStr(mvarDIASSEMANA) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GJOB'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 

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
Public Property Get QrySelect(Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDJOB, CODSIS, NOME, DSCJOB" & VbNewLine 
   Sql = Sql & ", DTINICIO, DTFIM, HORAEXECUCAO, DURACAO, DIASSEMANA" & VbNewLine 
   Sql = Sql &" From GJOB" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY 

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar(Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_WHERE, Ch_ORDERBY)
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
      mvarIDJOB = XVal(!IDJOB & "")
      mvarCODSIS = !CODSIS & ""
      mvarNOME = !NOME & ""
      mvarDSCJOB = !DSCJOB & ""
      mvarDTINICIO = xDate(!DTINICIO & "", True)
      mvarDTFIM = xDate(!DTFIM & "", True)
      mvarHORAEXECUCAO = xDate(!HORAEXECUCAO & "", True)
      mvarDURACAO = xDate(!DURACAO & "", True)
      mvarDIASSEMANA = !DIASSEMANA & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDJOB = 0
   mvarCODSIS = ""
   mvarNOME = ""
   mvarDSCJOB = ""
   mvarDTINICIO = ""
   mvarDTFIM = ""
   mvarHORAEXECUCAO = ""
   mvarDURACAO = ""
   mvarDIASSEMANA = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete(), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GJOB Set "
   Sql = Sql & " Where 
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
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
