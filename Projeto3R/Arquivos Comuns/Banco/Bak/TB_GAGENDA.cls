VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GAGENDA"
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

Private mvarIDAGENDA       As Double
Private mvarIDTAREFA       As Double
Private mvarOCORRENCIA     As String
Private mvarINTERVALO      As Double
Private mvarFREQSEMANA     As String
Private mvarDIAMES         As Double
Private mvarORDEMMES       As Double
Private mvarFREQMES        As Double
Private mvarHRINI          As String
Private mvarHRFIM          As String
Private mvarDTINI          As String
Private mvarDTFIM          As String
Private mvarDTHULTEXECUCAO As String
Private mvarSTATUS         As String
Public Property Get IDAGENDA() As Double
   IDAGENDA = mvarIDAGENDA
End Property
Public Property Let IDTAREFA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTAREFA <> vData)
   mvarIDTAREFA = vData
End Property
Public Property Get IDTAREFA() As Double
   IDTAREFA = mvarIDTAREFA
End Property
Public Property Let OCORRENCIA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOCORRENCIA <> vData)
   mvarOCORRENCIA = vData
End Property
Public Property Get OCORRENCIA() As String
   OCORRENCIA = mvarOCORRENCIA
End Property
Public Property Let INTERVALO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarINTERVALO <> vData)
   mvarINTERVALO = vData
End Property
Public Property Get INTERVALO() As Double
   INTERVALO = mvarINTERVALO
End Property
Public Property Let FREQSEMANA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFREQSEMANA <> vData)
   mvarFREQSEMANA = vData
End Property
Public Property Get FREQSEMANA() As String
   FREQSEMANA = mvarFREQSEMANA
End Property
Public Property Let DIAMES(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarDIAMES <> vData)
   mvarDIAMES = vData
End Property
Public Property Get DIAMES() As Double
   DIAMES = mvarDIAMES
End Property
Public Property Let ORDEMMES(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarORDEMMES <> vData)
   mvarORDEMMES = vData
End Property
Public Property Get ORDEMMES() As Double
   ORDEMMES = mvarORDEMMES
End Property
Public Property Let FREQMES(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFREQMES <> vData)
   mvarFREQMES = vData
End Property
Public Property Get FREQMES() As Double
   FREQMES = mvarFREQMES
End Property
Public Property Let HRINI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHRINI <> vData)
   mvarHRINI = vData
End Property
Public Property Get HRINI() As String
   HRINI = mvarHRINI
End Property
Public Property Let HRFIM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHRFIM <> vData)
   mvarHRFIM = vData
End Property
Public Property Get HRFIM() As String
   HRFIM = mvarHRFIM
End Property
Public Property Let DTINI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTINI <> vData)
   mvarDTINI = vData
End Property
Public Property Get DTINI() As String
   DTINI = mvarDTINI
End Property
Public Property Let DTFIM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTFIM <> vData)
   mvarDTFIM = vData
End Property
Public Property Get DTFIM() As String
   DTFIM = mvarDTFIM
End Property
Public Property Let DTHULTEXECUCAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTHULTEXECUCAO <> vData)
   mvarDTHULTEXECUCAO = vData
End Property
Public Property Get DTHULTEXECUCAO() As String
   DTHULTEXECUCAO = mvarDTHULTEXECUCAO
End Property
Public Property Let STATUS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSTATUS <> vData)
   mvarSTATUS = vData
End Property
Public Property Get STATUS() As String
   STATUS = mvarSTATUS
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
 
   Sql = "Insert Into GAGENDA (IDTAREFA, OCORRENCIA, INTERVALO" & VbNewLine 
   Sql = Sql & ", FREQSEMANA, DIAMES, ORDEMMES, FREQMES, HRINI" & VbNewLine 
   Sql = Sql & ", HRFIM, DTINI, DTFIM, DTHULTEXECUCAO, STATUS) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDTAREFA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOCORRENCIA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarINTERVALO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFREQSEMANA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarDIAMES) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarORDEMMES) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFREQMES) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarHRINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarHRFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTHULTEXECUCAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSTATUS) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDAGENDA, Optional Ch_IDTAREFA) As String
   Dim Sql As String
 
   Sql = "Delete From GAGENDA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDAGENDA) Then Sql = Sql & " IDAGENDA = " & SqlNum(Cstr(Ch_IDAGENDA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDTAREFA) Then Sql = Sql & " IDTAREFA = " & SqlNum(Cstr(Ch_IDTAREFA)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GAGENDA Set " & VbNewLine 
   Sql = Sql & " IDTAREFA       = " & SqlNum(mvarIDTAREFA) & VbNewLine 
   Sql = Sql & " , OCORRENCIA     = " & SqlStr(mvarOCORRENCIA) & VbNewLine 
   Sql = Sql & " , INTERVALO      = " & SqlNum(mvarINTERVALO) & VbNewLine 
   Sql = Sql & " , FREQSEMANA     = " & SqlStr(mvarFREQSEMANA) & VbNewLine 
   Sql = Sql & " , DIAMES         = " & SqlNum(mvarDIAMES) & VbNewLine 
   Sql = Sql & " , ORDEMMES       = " & SqlNum(mvarORDEMMES) & VbNewLine 
   Sql = Sql & " , FREQMES        = " & SqlNum(mvarFREQMES) & VbNewLine 
   Sql = Sql & " , HRINI          = " & SqlDate(mvarHRINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , HRFIM          = " & SqlDate(mvarHRFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTINI          = " & SqlDate(mvarDTINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTFIM          = " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTHULTEXECUCAO = " & SqlDate(mvarDTHULTEXECUCAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , STATUS         = " & SqlStr(mvarSTATUS) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GAGENDA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDAGENDA         = " & SqlNum(mvarIDAGENDA) & VbNewLine 
   Sql = Sql & " And IDTAREFA     = " & SqlNum(mvarIDTAREFA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDAGENDA, Optional Ch_IDTAREFA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDAGENDA, IDTAREFA, OCORRENCIA, INTERVALO" & VbNewLine 
   Sql = Sql & ", FREQSEMANA, DIAMES, ORDEMMES, FREQMES, HRINI" & VbNewLine 
   Sql = Sql & ", HRFIM, DTINI, DTFIM, DTHULTEXECUCAO, STATUS" & VbNewLine 
   Sql = Sql &" From GAGENDA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDAGENDA)  And IsMissing(Ch_IDTAREFA)  Then 
      Sql = Sql & " IDAGENDA = " & SqlNum(mvarIDAGENDA) & " AND " & VbNewLine 
      Sql = Sql & " IDTAREFA = " & SqlNum(mvarIDTAREFA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDAGENDA) Then Sql = Sql & " IDAGENDA = " & SqlNum(Cstr(Ch_IDAGENDA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDAGENDA, Optional Ch_IDTAREFA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDAGENDA, Ch_IDTAREFA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDAGENDA = XVal(!IDAGENDA & "")
      mvarIDTAREFA = XVal(!IDTAREFA & "")
      mvarOCORRENCIA = !OCORRENCIA & ""
      mvarINTERVALO = XVal(!INTERVALO & "")
      mvarFREQSEMANA = !FREQSEMANA & ""
      mvarDIAMES = XVal(!DIAMES & "")
      mvarORDEMMES = XVal(!ORDEMMES & "")
      mvarFREQMES = XVal(!FREQMES & "")
      mvarHRINI = xDate(!HRINI & "", True)
      mvarHRFIM = xDate(!HRFIM & "", True)
      mvarDTINI = xDate(!DTINI & "", True)
      mvarDTFIM = xDate(!DTFIM & "", True)
      mvarDTHULTEXECUCAO = xDate(!DTHULTEXECUCAO & "", True)
      mvarSTATUS = !STATUS & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDAGENDA, mvarIDTAREFA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDTAREFA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GAGENDA Set "
   Sql = Sql & " IDTAREFA = " & SqlNum(Ch_IDTAREFA)
   Sql = Sql & " Where 
   Sql = Sql & " IDAGENDA = " & SqlNum(mvarIDAGENDA)
   Sql = Sql & " and IDTAREFA = " & SqlNum(mvarIDTAREFA)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
