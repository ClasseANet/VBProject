VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OTRATAMENTOCLI"
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

Private mvarIDAREA         As Double
Private mvarIDCLIENTE      As Double
Private mvarIDLOJA         As Double
Private mvarIDTPTRATAMENTO As Double
Private mvarDTINI          As String
Private mvarDTFIM          As String
Private mvarSITTRATAMENTO  As Double
Private mvarFOTOTIPO       As Double
Private mvarOBS            As String
Private mvarALTERSTAMP     As Double
Private mvarTIMESTAMP      As String
Public Property Let IDAREA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAREA <> vData)
   mvarIDAREA = vData
End Property
Public Property Get IDAREA() As Double
   IDAREA = mvarIDAREA
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDTPTRATAMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPTRATAMENTO <> vData)
   mvarIDTPTRATAMENTO = vData
End Property
Public Property Get IDTPTRATAMENTO() As Double
   IDTPTRATAMENTO = mvarIDTPTRATAMENTO
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
Public Property Let SITTRATAMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSITTRATAMENTO <> vData)
   mvarSITTRATAMENTO = vData
End Property
Public Property Get SITTRATAMENTO() As Double
   SITTRATAMENTO = mvarSITTRATAMENTO
End Property
Public Property Let FOTOTIPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFOTOTIPO <> vData)
   mvarFOTOTIPO = vData
End Property
Public Property Get FOTOTIPO() As Double
   FOTOTIPO = mvarFOTOTIPO
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OTRATAMENTOCLI Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDAREA           = " & SqlNum(mvarIDAREA) & VbNewLine 
   Sql = Sql & " And IDCLIENTE    = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & " And IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
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
Public Property Let TIMESTAMP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIMESTAMP <> vData)
   mvarTIMESTAMP = vData
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into OTRATAMENTOCLI (IDAREA, IDCLIENTE, IDLOJA, IDTPTRATAMENTO" & VbNewLine 
   Sql = Sql & ", DTINI, DTFIM, SITTRATAMENTO, FOTOTIPO, OBS" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDAREA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSITTRATAMENTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFOTOTIPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDAREA, Optional Ch_IDCLIENTE, Optional Ch_IDLOJA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OTRATAMENTOCLI" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDAREA)  And IsMissing(Ch_IDCLIENTE)  And IsMissing(Ch_IDLOJA)  Then 
      Sql = Sql & " IDAREA = " & SqlNum(mvarIDAREA) & " AND " & VbNewLine 
      Sql = Sql & " IDCLIENTE = " & SqlNum(mvarIDCLIENTE) & " AND " & VbNewLine 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDAREA) Then Sql = Sql & " IDAREA = " & SqlNum(Cstr(Ch_IDAREA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCLIENTE) Then Sql = Sql & " IDCLIENTE = " & SqlNum(Cstr(Ch_IDCLIENTE)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
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

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = "
   sTag = sTag & "|IDAREA = " & mvarIDAREA
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDTPTRATAMENTO = " & mvarIDTPTRATAMENTO
   sTag = sTag & "|DTINI = " & mvarDTINI
   sTag = sTag & "|DTFIM = " & mvarDTFIM
   sTag = sTag & "|SITTRATAMENTO = " & mvarSITTRATAMENTO
   sTag = sTag & "|FOTOTIPO = " & mvarFOTOTIPO
   sTag = sTag & "|OBS = " & mvarOBS
   sTag = sTag & "|Where = " & IIf(IsMissing(Ch_WHERE), "", Ch_WHERE)
   sTag = sTag & "|"

   MyDelRow.Query = Sql & ";"
   MyDelRow.Tag = sTag
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update OTRATAMENTOCLI Set " & VbNewLine 
   Sql = Sql & " IDAREA         = " & SqlNum(mvarIDAREA) & VbNewLine 
   Sql = Sql & " , IDCLIENTE      = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & " , IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDTPTRATAMENTO = " & IIf(mvarIDTPTRATAMENTO = 0, "Null", SqlNum(mvarIDTPTRATAMENTO)) & VbNewLine 
   Sql = Sql & " , DTINI          = " & SqlDate(mvarDTINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTFIM          = " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , SITTRATAMENTO  = " & SqlNum(mvarSITTRATAMENTO) & VbNewLine 
   Sql = Sql & " , FOTOTIPO       = " & SqlNum(mvarFOTOTIPO) & VbNewLine 
   Sql = Sql & " , OBS            = " & SqlStr(mvarOBS) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP     = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP     = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP      = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP      = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OTRATAMENTOCLI'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDAREA           = " & SqlNum(mvarIDAREA) & VbNewLine 
   Sql = Sql & " And IDCLIENTE    = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & " And IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDAREA, Optional Ch_IDCLIENTE, Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDAREA, IDCLIENTE, IDLOJA, IDTPTRATAMENTO" & VbNewLine 
   Sql = Sql & ", DTINI, DTFIM, SITTRATAMENTO, FOTOTIPO, OBS" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OTRATAMENTOCLI" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDAREA)  And IsMissing(Ch_IDCLIENTE)  And IsMissing(Ch_IDLOJA)  Then 
      Sql = Sql & " IDAREA = " & SqlNum(mvarIDAREA) & " AND " & VbNewLine 
      Sql = Sql & " IDCLIENTE = " & SqlNum(mvarIDCLIENTE) & " AND " & VbNewLine 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDAREA) Then Sql = Sql & " IDAREA = " & SqlNum(Cstr(Ch_IDAREA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCLIENTE) Then Sql = Sql & " IDCLIENTE = " & SqlNum(Cstr(Ch_IDCLIENTE)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDAREA, Optional Ch_IDCLIENTE, Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDAREA, Ch_IDCLIENTE, Ch_IDLOJA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDAREA = XVal(!IDAREA & "")
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
      mvarIDLOJA = XVal(!IDLOJA & "")
      mvarIDTPTRATAMENTO = XVal(!IDTPTRATAMENTO & "")
      mvarDTINI = xDate(!DTINI & "", True)
      mvarDTFIM = xDate(!DTFIM & "", True)
      mvarSITTRATAMENTO = XVal(!SITTRATAMENTO & "")
      mvarFOTOTIPO = XVal(!FOTOTIPO & "")
      mvarOBS = !OBS & ""
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDAREA = 0
   mvarIDCLIENTE = 0
   mvarIDLOJA = 0
   mvarIDTPTRATAMENTO = 0
   mvarDTINI = ""
   mvarDTFIM = ""
   mvarSITTRATAMENTO = 0
   mvarFOTOTIPO = 0
   mvarOBS = ""
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDAREA, mvarIDCLIENTE, mvarIDLOJA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDAREA As String, Ch_IDCLIENTE As String, Ch_IDLOJA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OTRATAMENTOCLI Set "
   Sql = Sql & " IDAREA = " & SqlNum(Ch_IDAREA)
   Sql = Sql & " , IDCLIENTE = " & SqlNum(Ch_IDCLIENTE)
   Sql = Sql & " , IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " Where 
   Sql = Sql & " IDAREA = " & SqlNum(mvarIDAREA)
   Sql = Sql & " and IDCLIENTE = " & SqlNum(mvarIDCLIENTE)
   Sql = Sql & " and IDLOJA = " & SqlNum(mvarIDLOJA)
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
