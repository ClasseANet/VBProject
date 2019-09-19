VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OATENDIMENTO"
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

Private mvarIDLOJA        As Double
Private mvarIDATENDIMENTO As Double
Private mvarIDMAQUINA     As Double
Private mvarIDSALA        As Double
Private mvarIDMANIPULO    As Double
Private mvarIDEVENTO      As Double
Private mvarIDCLIENTE     As Double
Private mvarIDFUNCIONARIO As Double
Private mvarOBS           As String
Private mvarDTATEND       As String
Private mvarHHINI         As String
Private mvarHHFIM         As String
Private mvarSITATEND      As String
Private mvarALTERSTAMP    As Double
Private mvarTIMESTAMP     As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDATENDIMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDATENDIMENTO <> vData)
   mvarIDATENDIMENTO = vData
End Property
Public Property Get IDATENDIMENTO() As Double
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDMAQUINA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMAQUINA <> vData)
   mvarIDMAQUINA = vData
End Property
Public Property Get IDMAQUINA() As Double
   IDMAQUINA = mvarIDMAQUINA
End Property
Public Property Let IDSALA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSALA <> vData)
   mvarIDSALA = vData
End Property
Public Property Get IDSALA() As Double
   IDSALA = mvarIDSALA
End Property
Public Property Let IDMANIPULO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMANIPULO <> vData)
   mvarIDMANIPULO = vData
End Property
Public Property Get IDMANIPULO() As Double
   IDMANIPULO = mvarIDMANIPULO
End Property
Public Property Let IDEVENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEVENTO <> vData)
   mvarIDEVENTO = vData
End Property
Public Property Get IDEVENTO() As Double
   IDEVENTO = mvarIDEVENTO
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDFUNCIONARIO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFUNCIONARIO <> vData)
   mvarIDFUNCIONARIO = vData
End Property
Public Property Get IDFUNCIONARIO() As Double
   IDFUNCIONARIO = mvarIDFUNCIONARIO
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
End Property
Public Property Let DTATEND(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTATEND <> vData)
   mvarDTATEND = vData
End Property
Public Property Get DTATEND() As String
   DTATEND = mvarDTATEND
End Property
Public Property Let HHINI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHHINI <> vData)
   mvarHHINI = vData
End Property
Public Property Get HHINI() As String
   HHINI = mvarHHINI
End Property
Public Property Let HHFIM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHHFIM <> vData)
   mvarHHFIM = vData
End Property
Public Property Get HHFIM() As String
   HHFIM = mvarHHFIM
End Property
Public Property Let SITATEND(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITATEND <> vData)
   mvarSITATEND = vData
End Property
Public Property Get SITATEND() As String
   SITATEND = mvarSITATEND
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OATENDIMENTO Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & VbNewLine 
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
 
   Sql = "Insert Into OATENDIMENTO (IDLOJA, IDATENDIMENTO, IDMAQUINA, IDSALA" & VbNewLine 
   Sql = Sql & ", IDMANIPULO, IDEVENTO, IDCLIENTE, IDFUNCIONARIO, OBS" & VbNewLine 
   Sql = Sql & ", DTATEND, HHINI, HHFIM, SITATEND, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDATENDIMENTO),0)+1 From OATENDIMENTO Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDATENDIMENTO) & VbNewLine 
   End If
   Sql = Sql & ", " & IIf(mvarIDMAQUINA = 0, "Null", SqlNum(mvarIDMAQUINA)) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDSALA = 0, "Null", SqlNum(mvarIDSALA)) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDMANIPULO = 0, "Null", SqlNum(mvarIDMANIPULO)) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDEVENTO = 0, "Null", SqlNum(mvarIDEVENTO)) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDCLIENTE = 0, "Null", SqlNum(mvarIDCLIENTE)) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDFUNCIONARIO = 0, "Null", SqlNum(mvarIDFUNCIONARIO)) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTATEND, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarHHINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarHHFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSITATEND) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", " &  SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", " &  SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDATENDIMENTO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OATENDIMENTO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDATENDIMENTO)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDATENDIMENTO) Then Sql = Sql & " IDATENDIMENTO = " & SqlNum(Cstr(Ch_IDATENDIMENTO)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDATENDIMENTO = " & mvarIDATENDIMENTO
   sTag = sTag & "|IDMAQUINA = " & mvarIDMAQUINA
   sTag = sTag & "|IDSALA = " & mvarIDSALA
   sTag = sTag & "|IDMANIPULO = " & mvarIDMANIPULO
   sTag = sTag & "|IDEVENTO = " & mvarIDEVENTO
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|IDFUNCIONARIO = " & mvarIDFUNCIONARIO
   sTag = sTag & "|OBS = " & mvarOBS
   sTag = sTag & "|DTATEND = " & mvarDTATEND
   sTag = sTag & "|HHINI = " & mvarHHINI
   sTag = sTag & "|HHFIM = " & mvarHHFIM
   sTag = sTag & "|SITATEND = " & mvarSITATEND
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
 
   Sql = "Update OATENDIMENTO Set " & VbNewLine 
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & VbNewLine 
   Sql = Sql & " , IDMAQUINA     = " & IIf(mvarIDMAQUINA = 0, "Null", SqlNum(mvarIDMAQUINA)) & VbNewLine 
   Sql = Sql & " , IDSALA        = " & IIf(mvarIDSALA = 0, "Null", SqlNum(mvarIDSALA)) & VbNewLine 
   Sql = Sql & " , IDMANIPULO    = " & IIf(mvarIDMANIPULO = 0, "Null", SqlNum(mvarIDMANIPULO)) & VbNewLine 
   Sql = Sql & " , IDEVENTO      = " & IIf(mvarIDEVENTO = 0, "Null", SqlNum(mvarIDEVENTO)) & VbNewLine 
   Sql = Sql & " , IDCLIENTE     = " & IIf(mvarIDCLIENTE = 0, "Null", SqlNum(mvarIDCLIENTE)) & VbNewLine 
   Sql = Sql & " , IDFUNCIONARIO = " & IIf(mvarIDFUNCIONARIO = 0, "Null", SqlNum(mvarIDFUNCIONARIO)) & VbNewLine 
   Sql = Sql & " , OBS           = " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & " , DTATEND       = " & SqlDate(mvarDTATEND, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , HHINI         = " & SqlDate(mvarHHINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , HHFIM         = " & SqlDate(mvarHHFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , SITATEND      = " & SqlStr(mvarSITATEND) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP    = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP    = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP     = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP     = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OATENDIMENTO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDATENDIMENTO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDATENDIMENTO, IDMAQUINA, IDSALA" & VbNewLine 
   Sql = Sql & ", IDMANIPULO, IDEVENTO, IDCLIENTE, IDFUNCIONARIO, OBS" & VbNewLine 
   Sql = Sql & ", DTATEND, HHINI, HHFIM, SITATEND, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OATENDIMENTO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDATENDIMENTO)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDATENDIMENTO) Then Sql = Sql & " IDATENDIMENTO = " & SqlNum(Cstr(Ch_IDATENDIMENTO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDATENDIMENTO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDATENDIMENTO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDLOJA = XVal(!IDLOJA & "")
      mvarIDATENDIMENTO = XVal(!IDATENDIMENTO & "")
      mvarIDMAQUINA = XVal(!IDMAQUINA & "")
      mvarIDSALA = XVal(!IDSALA & "")
      mvarIDMANIPULO = XVal(!IDMANIPULO & "")
      mvarIDEVENTO = XVal(!IDEVENTO & "")
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
      mvarIDFUNCIONARIO = XVal(!IDFUNCIONARIO & "")
      mvarOBS = !OBS & ""
      mvarDTATEND = xDate(!DTATEND & "", True)
      mvarHHINI = xDate(!HHINI & "", True)
      mvarHHFIM = xDate(!HHFIM & "", True)
      mvarSITATEND = !SITATEND & ""
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDATENDIMENTO = 0
   mvarIDMAQUINA = 0
   mvarIDSALA = 0
   mvarIDMANIPULO = 0
   mvarIDEVENTO = 0
   mvarIDCLIENTE = 0
   mvarIDFUNCIONARIO = 0
   mvarOBS = ""
   mvarDTATEND = ""
   mvarHHINI = ""
   mvarHHFIM = ""
   mvarSITATEND = ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDATENDIMENTO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDATENDIMENTO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OATENDIMENTO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDATENDIMENTO = " & SqlNum(Ch_IDATENDIMENTO)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO)
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
