VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OTAREFAEVT"
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

Private mvarIDLOJA      As Double
Private mvarIDTAREFA    As Double
Private mvarIDEVENTO    As Double
Private mvarDSCTAREFA   As String
Private mvarTITULO      As String
Private mvarDTTAREFA    As String
Private mvarSITTAREFA   As Double
Private mvarIDCLIENTE   As Double
Private mvarPRIORIDADE  As Double
Private mvarFLGLEMBRETE As Double
Private mvarDTLEMBRETE  As String
Private mvarIDTPTAREFA  As Double
Private mvarFLGDELETE   As Double
Private mvarFLGAUTO     As Double
Private mvarALTERSTAMP  As Double
Private mvarTIMESTAMP   As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDTAREFA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTAREFA <> vData)
   mvarIDTAREFA = vData
End Property
Public Property Get IDTAREFA() As Double
   IDTAREFA = mvarIDTAREFA
End Property
Public Property Let IDEVENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEVENTO <> vData)
   mvarIDEVENTO = vData
End Property
Public Property Get IDEVENTO() As Double
   IDEVENTO = mvarIDEVENTO
End Property
Public Property Let DSCTAREFA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCTAREFA <> vData)
   mvarDSCTAREFA = vData
End Property
Public Property Get DSCTAREFA() As String
   DSCTAREFA = mvarDSCTAREFA
End Property
Public Property Let TITULO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTITULO <> vData)
   mvarTITULO = vData
End Property
Public Property Get TITULO() As String
   TITULO = mvarTITULO
End Property
Public Property Let DTTAREFA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTTAREFA <> vData)
   mvarDTTAREFA = vData
End Property
Public Property Get DTTAREFA() As String
   DTTAREFA = mvarDTTAREFA
End Property
Public Property Let SITTAREFA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSITTAREFA <> vData)
   mvarSITTAREFA = vData
End Property
Public Property Get SITTAREFA() As Double
   SITTAREFA = mvarSITTAREFA
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let PRIORIDADE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPRIORIDADE <> vData)
   mvarPRIORIDADE = vData
End Property
Public Property Get PRIORIDADE() As Double
   PRIORIDADE = mvarPRIORIDADE
End Property
Public Property Let FLGLEMBRETE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGLEMBRETE <> vData)
   mvarFLGLEMBRETE = vData
End Property
Public Property Get FLGLEMBRETE() As Double
   FLGLEMBRETE = mvarFLGLEMBRETE
End Property
Public Property Let DTLEMBRETE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTLEMBRETE <> vData)
   mvarDTLEMBRETE = vData
End Property
Public Property Get DTLEMBRETE() As String
   DTLEMBRETE = mvarDTLEMBRETE
End Property
Public Property Let IDTPTAREFA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPTAREFA <> vData)
   mvarIDTPTAREFA = vData
End Property
Public Property Get IDTPTAREFA() As Double
   IDTPTAREFA = mvarIDTPTAREFA
End Property
Public Property Let FLGDELETE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGDELETE <> vData)
   mvarFLGDELETE = vData
End Property
Public Property Get FLGDELETE() As Double
   FLGDELETE = mvarFLGDELETE
End Property
Public Property Let FLGAUTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGAUTO <> vData)
   mvarFLGAUTO = vData
End Property
Public Property Get FLGAUTO() As Double
   FLGAUTO = mvarFLGAUTO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OTAREFAEVT Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDTAREFA  = " & SqlNum(mvarIDTAREFA) & VbNewLine 
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
 
   Sql = "Insert Into OTAREFAEVT (IDLOJA, IDTAREFA, IDEVENTO, DSCTAREFA" & VbNewLine 
   Sql = Sql & ", TITULO, DTTAREFA, SITTAREFA, IDCLIENTE, PRIORIDADE" & VbNewLine 
   Sql = Sql & ", FLGLEMBRETE, DTLEMBRETE, IDTPTAREFA, FLGDELETE, FLGAUTO" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDTAREFA),0)+1 From OTAREFAEVT Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDTAREFA) & VbNewLine 
   End If
   Sql = Sql & ", " & IIf(mvarIDEVENTO = 0, "Null", SqlNum(mvarIDEVENTO)) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCTAREFA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTITULO) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTTAREFA, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSITTAREFA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPRIORIDADE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGLEMBRETE) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTLEMBRETE, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTPTAREFA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGDELETE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGAUTO) & VbNewLine 
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
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDTAREFA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OTAREFAEVT" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDTAREFA)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDTAREFA = " & SqlNum(mvarIDTAREFA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
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

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = "
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDTAREFA = " & mvarIDTAREFA
   sTag = sTag & "|IDEVENTO = " & mvarIDEVENTO
   sTag = sTag & "|DSCTAREFA = " & mvarDSCTAREFA
   sTag = sTag & "|TITULO = " & mvarTITULO
   sTag = sTag & "|DTTAREFA = " & mvarDTTAREFA
   sTag = sTag & "|SITTAREFA = " & mvarSITTAREFA
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|PRIORIDADE = " & mvarPRIORIDADE
   sTag = sTag & "|FLGLEMBRETE = " & mvarFLGLEMBRETE
   sTag = sTag & "|DTLEMBRETE = " & mvarDTLEMBRETE
   sTag = sTag & "|IDTPTAREFA = " & mvarIDTPTAREFA
   sTag = sTag & "|FLGDELETE = " & mvarFLGDELETE
   sTag = sTag & "|FLGAUTO = " & mvarFLGAUTO
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
 
   Sql = "Update OTAREFAEVT Set " & VbNewLine 
   Sql = Sql & " IDLOJA      = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDTAREFA    = " & SqlNum(mvarIDTAREFA) & VbNewLine 
   Sql = Sql & " , IDEVENTO    = " & IIf(mvarIDEVENTO = 0, "Null", SqlNum(mvarIDEVENTO)) & VbNewLine 
   Sql = Sql & " , DSCTAREFA   = " & SqlStr(mvarDSCTAREFA) & VbNewLine 
   Sql = Sql & " , TITULO      = " & SqlStr(mvarTITULO) & VbNewLine 
   Sql = Sql & " , DTTAREFA    = " & SqlDate(mvarDTTAREFA, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , SITTAREFA   = " & SqlNum(mvarSITTAREFA) & VbNewLine 
   Sql = Sql & " , IDCLIENTE   = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & " , PRIORIDADE  = " & SqlNum(mvarPRIORIDADE) & VbNewLine 
   Sql = Sql & " , FLGLEMBRETE = " & SqlNum(mvarFLGLEMBRETE) & VbNewLine 
   Sql = Sql & " , DTLEMBRETE  = " & SqlDate(mvarDTLEMBRETE, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , IDTPTAREFA  = " & SqlNum(mvarIDTPTAREFA) & VbNewLine 
   Sql = Sql & " , FLGDELETE   = " & SqlNum(mvarFLGDELETE) & VbNewLine 
   Sql = Sql & " , FLGAUTO     = " & SqlNum(mvarFLGAUTO) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP  = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP  = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP   = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP   = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OTAREFAEVT'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDTAREFA  = " & SqlNum(mvarIDTAREFA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDTAREFA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDTAREFA, IDEVENTO, DSCTAREFA" & VbNewLine 
   Sql = Sql & ", TITULO, DTTAREFA, SITTAREFA, IDCLIENTE, PRIORIDADE" & VbNewLine 
   Sql = Sql & ", FLGLEMBRETE, DTLEMBRETE, IDTPTAREFA, FLGDELETE, FLGAUTO" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OTAREFAEVT" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDTAREFA)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDTAREFA = " & SqlNum(mvarIDTAREFA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDTAREFA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDTAREFA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDTAREFA = XVal(!IDTAREFA & "")
      mvarIDEVENTO = XVal(!IDEVENTO & "")
      mvarDSCTAREFA = !DSCTAREFA & ""
      mvarTITULO = !TITULO & ""
      mvarDTTAREFA = xDate(!DTTAREFA & "", True)
      mvarSITTAREFA = XVal(!SITTAREFA & "")
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
      mvarPRIORIDADE = XVal(!PRIORIDADE & "")
      mvarFLGLEMBRETE = XVal(!FLGLEMBRETE & "")
      mvarDTLEMBRETE = xDate(!DTLEMBRETE & "", True)
      mvarIDTPTAREFA = XVal(!IDTPTAREFA & "")
      mvarFLGDELETE = XVal(!FLGDELETE & "")
      mvarFLGAUTO = XVal(!FLGAUTO & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDTAREFA = 0
   mvarIDEVENTO = 0
   mvarDSCTAREFA = ""
   mvarTITULO = ""
   mvarDTTAREFA = ""
   mvarSITTAREFA = 0
   mvarIDCLIENTE = 0
   mvarPRIORIDADE = 0
   mvarFLGLEMBRETE = 0
   mvarDTLEMBRETE = ""
   mvarIDTPTAREFA = 0
   mvarFLGDELETE = 0
   mvarFLGAUTO = 0
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDTAREFA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDTAREFA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OTAREFAEVT Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDTAREFA = " & SqlNum(Ch_IDTAREFA)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDTAREFA = " & SqlNum(mvarIDTAREFA)
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
