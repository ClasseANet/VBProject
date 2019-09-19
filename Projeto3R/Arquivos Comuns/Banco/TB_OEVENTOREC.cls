VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OEVENTOREC"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb                        As Object 
Private mvarRS                         As Object 

Private mvarQryInsert                  As String
Private mvarQryUpDate                  As String
Private mvarQryDelete                  As String
Private mvarQrySelect                  As String
Private mvarQrySave                    As String

Private mvarisDirt                     As Boolean

Private mvarIDLOJA                     As Double
Private mvarIDEVENTOREC                As Double
Private mvarIDAGENDA                   As Double
Private mvarMasterEventID              As Double
Private mvarPatternStartDate           As String
Private mvarPatternEndMethod           As Double
Private mvarPatternEndDate             As String
Private mvarPatternEndAfterOccurrences As Double
Private mvarEventStartTime             As String
Private mvarEventDuration              As Double
Private mvarOptionsData1               As Double
Private mvarOptionsData2               As Double
Private mvarOptionsData3               As Double
Private mvarOptionsData4               As Double
Private mvarCustomPropertiesXMLData    As String
Private mvarALTERSTAMP                 As Double
Private mvarTIMESTAMP                  As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDEVENTOREC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEVENTOREC <> vData)
   mvarIDEVENTOREC = vData
End Property
Public Property Get IDEVENTOREC() As Double
   IDEVENTOREC = mvarIDEVENTOREC
End Property
Public Property Let IDAGENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAGENDA <> vData)
   mvarIDAGENDA = vData
End Property
Public Property Get IDAGENDA() As Double
   IDAGENDA = mvarIDAGENDA
End Property
Public Property Let MasterEventID(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarMasterEventID <> vData)
   mvarMasterEventID = vData
End Property
Public Property Get MasterEventID() As Double
   MasterEventID = mvarMasterEventID
End Property
Public Property Let PatternStartDate(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPatternStartDate <> vData)
   mvarPatternStartDate = vData
End Property
Public Property Get PatternStartDate() As String
   PatternStartDate = mvarPatternStartDate
End Property
Public Property Let PatternEndMethod(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPatternEndMethod <> vData)
   mvarPatternEndMethod = vData
End Property
Public Property Get PatternEndMethod() As Double
   PatternEndMethod = mvarPatternEndMethod
End Property
Public Property Let PatternEndDate(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPatternEndDate <> vData)
   mvarPatternEndDate = vData
End Property
Public Property Get PatternEndDate() As String
   PatternEndDate = mvarPatternEndDate
End Property
Public Property Let PatternEndAfterOccurrences(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPatternEndAfterOccurrences <> vData)
   mvarPatternEndAfterOccurrences = vData
End Property
Public Property Get PatternEndAfterOccurrences() As Double
   PatternEndAfterOccurrences = mvarPatternEndAfterOccurrences
End Property
Public Property Let EventStartTime(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEventStartTime <> vData)
   mvarEventStartTime = vData
End Property
Public Property Get EventStartTime() As String
   EventStartTime = mvarEventStartTime
End Property
Public Property Let EventDuration(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarEventDuration <> vData)
   mvarEventDuration = vData
End Property
Public Property Get EventDuration() As Double
   EventDuration = mvarEventDuration
End Property
Public Property Let OptionsData1(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarOptionsData1 <> vData)
   mvarOptionsData1 = vData
End Property
Public Property Get OptionsData1() As Double
   OptionsData1 = mvarOptionsData1
End Property
Public Property Let OptionsData2(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarOptionsData2 <> vData)
   mvarOptionsData2 = vData
End Property
Public Property Get OptionsData2() As Double
   OptionsData2 = mvarOptionsData2
End Property
Public Property Let OptionsData3(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarOptionsData3 <> vData)
   mvarOptionsData3 = vData
End Property
Public Property Get OptionsData3() As Double
   OptionsData3 = mvarOptionsData3
End Property
Public Property Let OptionsData4(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarOptionsData4 <> vData)
   mvarOptionsData4 = vData
End Property
Public Property Get OptionsData4() As Double
   OptionsData4 = mvarOptionsData4
End Property
Public Property Let CustomPropertiesXMLData(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCustomPropertiesXMLData <> vData)
   mvarCustomPropertiesXMLData = vData
End Property
Public Property Get CustomPropertiesXMLData() As String
   CustomPropertiesXMLData = mvarCustomPropertiesXMLData
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OEVENTOREC Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA                       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDEVENTOREC              = " & SqlNum(mvarIDEVENTOREC) & VbNewLine 
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
 
   Sql = "Insert Into OEVENTOREC (IDLOJA, IDEVENTOREC, IDAGENDA, MasterEventID" & VbNewLine 
   Sql = Sql & ", PatternStartDate, PatternEndMethod, PatternEndDate, PatternEndAfterOccurrences, EventStartTime" & VbNewLine 
   Sql = Sql & ", EventDuration, OptionsData1, OptionsData2, OptionsData3, OptionsData4" & VbNewLine 
   Sql = Sql & ", CustomPropertiesXMLData, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDEVENTOREC),0)+1 From OEVENTOREC Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDEVENTOREC) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDAGENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarMasterEventID) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarPatternStartDate, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPatternEndMethod) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarPatternEndDate, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPatternEndAfterOccurrences) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEventStartTime) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarEventDuration) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarOptionsData1) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarOptionsData2) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarOptionsData3) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarOptionsData4) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCustomPropertiesXMLData) & VbNewLine 
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
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDEVENTOREC, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OEVENTOREC" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDEVENTOREC)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDEVENTOREC = " & SqlNum(mvarIDEVENTOREC) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDEVENTOREC) Then Sql = Sql & " IDEVENTOREC = " & SqlNum(Cstr(Ch_IDEVENTOREC)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDEVENTOREC = " & mvarIDEVENTOREC
   sTag = sTag & "|IDAGENDA = " & mvarIDAGENDA
   sTag = sTag & "|MasterEventID = " & mvarMasterEventID
   sTag = sTag & "|PatternStartDate = " & mvarPatternStartDate
   sTag = sTag & "|PatternEndMethod = " & mvarPatternEndMethod
   sTag = sTag & "|PatternEndDate = " & mvarPatternEndDate
   sTag = sTag & "|PatternEndAfterOccurrences = " & mvarPatternEndAfterOccurrences
   sTag = sTag & "|EventStartTime = " & mvarEventStartTime
   sTag = sTag & "|EventDuration = " & mvarEventDuration
   sTag = sTag & "|OptionsData1 = " & mvarOptionsData1
   sTag = sTag & "|OptionsData2 = " & mvarOptionsData2
   sTag = sTag & "|OptionsData3 = " & mvarOptionsData3
   sTag = sTag & "|OptionsData4 = " & mvarOptionsData4
   sTag = sTag & "|CustomPropertiesXMLData = " & mvarCustomPropertiesXMLData
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
 
   Sql = "Update OEVENTOREC Set " & VbNewLine 
   Sql = Sql & " IDLOJA                     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDEVENTOREC                = " & SqlNum(mvarIDEVENTOREC) & VbNewLine 
   Sql = Sql & " , IDAGENDA                   = " & SqlNum(mvarIDAGENDA) & VbNewLine 
   Sql = Sql & " , MasterEventID              = " & SqlNum(mvarMasterEventID) & VbNewLine 
   Sql = Sql & " , PatternStartDate           = " & SqlDate(mvarPatternStartDate, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , PatternEndMethod           = " & SqlNum(mvarPatternEndMethod) & VbNewLine 
   Sql = Sql & " , PatternEndDate             = " & SqlDate(mvarPatternEndDate, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , PatternEndAfterOccurrences = " & SqlNum(mvarPatternEndAfterOccurrences) & VbNewLine 
   Sql = Sql & " , EventStartTime             = " & SqlStr(mvarEventStartTime) & VbNewLine 
   Sql = Sql & " , EventDuration              = " & SqlNum(mvarEventDuration) & VbNewLine 
   Sql = Sql & " , OptionsData1               = " & SqlNum(mvarOptionsData1) & VbNewLine 
   Sql = Sql & " , OptionsData2               = " & SqlNum(mvarOptionsData2) & VbNewLine 
   Sql = Sql & " , OptionsData3               = " & SqlNum(mvarOptionsData3) & VbNewLine 
   Sql = Sql & " , OptionsData4               = " & SqlNum(mvarOptionsData4) & VbNewLine 
   Sql = Sql & " , CustomPropertiesXMLData    = " & SqlStr(mvarCustomPropertiesXMLData) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP                 = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP                 = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP                  = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP                  = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OEVENTOREC'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA                       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDEVENTOREC              = " & SqlNum(mvarIDEVENTOREC) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDEVENTOREC, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDEVENTOREC, IDAGENDA, MasterEventID" & VbNewLine 
   Sql = Sql & ", PatternStartDate, PatternEndMethod, PatternEndDate, PatternEndAfterOccurrences, EventStartTime" & VbNewLine 
   Sql = Sql & ", EventDuration, OptionsData1, OptionsData2, OptionsData3, OptionsData4" & VbNewLine 
   Sql = Sql & ", CustomPropertiesXMLData, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OEVENTOREC" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDEVENTOREC)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDEVENTOREC = " & SqlNum(mvarIDEVENTOREC) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDEVENTOREC) Then Sql = Sql & " IDEVENTOREC = " & SqlNum(Cstr(Ch_IDEVENTOREC)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDEVENTOREC, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDEVENTOREC, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDEVENTOREC = XVal(!IDEVENTOREC & "")
      mvarIDAGENDA = XVal(!IDAGENDA & "")
      mvarMasterEventID = XVal(!MasterEventID & "")
      mvarPatternStartDate = xDate(!PatternStartDate & "", True)
      mvarPatternEndMethod = XVal(!PatternEndMethod & "")
      mvarPatternEndDate = xDate(!PatternEndDate & "", True)
      mvarPatternEndAfterOccurrences = XVal(!PatternEndAfterOccurrences & "")
      mvarEventStartTime = !EventStartTime & ""
      mvarEventDuration = XVal(!EventDuration & "")
      mvarOptionsData1 = XVal(!OptionsData1 & "")
      mvarOptionsData2 = XVal(!OptionsData2 & "")
      mvarOptionsData3 = XVal(!OptionsData3 & "")
      mvarOptionsData4 = XVal(!OptionsData4 & "")
      mvarCustomPropertiesXMLData = !CustomPropertiesXMLData & ""
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDEVENTOREC = 0
   mvarIDAGENDA = 0
   mvarMasterEventID = 0
   mvarPatternStartDate = ""
   mvarPatternEndMethod = 0
   mvarPatternEndDate = ""
   mvarPatternEndAfterOccurrences = 0
   mvarEventStartTime = ""
   mvarEventDuration = 0
   mvarOptionsData1 = 0
   mvarOptionsData2 = 0
   mvarOptionsData3 = 0
   mvarOptionsData4 = 0
   mvarCustomPropertiesXMLData = ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDEVENTOREC), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDEVENTOREC As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OEVENTOREC Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDEVENTOREC = " & SqlNum(Ch_IDEVENTOREC)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDEVENTOREC = " & SqlNum(mvarIDEVENTOREC)
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
