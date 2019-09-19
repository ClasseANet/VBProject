VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OEVENTOAGENDA"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb                           As Object 
Private mvarRS                            As Object 

Private mvarQryInsert                     As String
Private mvarQryUpDate                     As String
Private mvarQryDelete                     As String
Private mvarQrySelect                     As String
Private mvarQrySave                       As String

Private mvarisDirt                        As Boolean

Private mvarIDEVENTO                     As Double
Private mvarIDLOJA                       As Double
Private mvarIDEVENTOREC                  As Double
Private mvarIDAGENDA                     As Double
Private mvarIDCLIENTE                    As Double
Private mvarStartDateTime                As String
Private mvarEndDateTime                  As String
Private mvarRecurrenceState              As Double
Private mvarSubject                      As String
Private mvarLocation                     As String
Private mvarBody                         As String
Private mvarBusyStatus                   As Double
Private mvarImportanceLevel              As Double
Private mvarLabelID                      As Double
Private mvarScheduleID                   As Double
Private mvarIsRecurrenceExceptionDeleted As Double
Private mvarRExceptionStartTimeOrig      As String
Private mvarRExceptionEndTimeOrig        As String
Private mvarIsAllDayEvent                As Double
Private mvarIsMeeting                    As Double
Private mvarIsPrivate                    As Double
Private mvarIsReminder                   As Double
Private mvarReminderMinutesBeforeStart   As Double
Private mvarRemainderSoundFile           As String
Private mvarCustomPropertiesXMLData      As String
Private mvarCustomIconsIDs               As String
Private mvarCreated                      As String
Private mvarModified                     As String
Private mvarFLGCONFIRMADO                As Double
Private mvarFLGCANCELADO                 As Double
Private mvarFLGREMARCADO                 As Double
Private mvarALTERSTAMP                   As Double
Private mvarTIMESTAMP                    As String
Public Property Let IDEVENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEVENTO <> vData)
   mvarIDEVENTO = vData
End Property
Public Property Get IDEVENTO() As Double
   IDEVENTO = mvarIDEVENTO
End Property
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
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let StartDateTime(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarStartDateTime <> vData)
   mvarStartDateTime = vData
End Property
Public Property Get StartDateTime() As String
   StartDateTime = mvarStartDateTime
End Property
Public Property Let EndDateTime(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEndDateTime <> vData)
   mvarEndDateTime = vData
End Property
Public Property Get EndDateTime() As String
   EndDateTime = mvarEndDateTime
End Property
Public Property Let RecurrenceState(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarRecurrenceState <> vData)
   mvarRecurrenceState = vData
End Property
Public Property Get RecurrenceState() As Double
   RecurrenceState = mvarRecurrenceState
End Property
Public Property Let Subject(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSubject <> vData)
   mvarSubject = vData
End Property
Public Property Get Subject() As String
   Subject = mvarSubject
End Property
Public Property Let Location(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLocation <> vData)
   mvarLocation = vData
End Property
Public Property Get Location() As String
   Location = mvarLocation
End Property
Public Property Let Body(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarBody <> vData)
   mvarBody = vData
End Property
Public Property Get Body() As String
   Body = mvarBody
End Property
Public Property Let BusyStatus(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarBusyStatus <> vData)
   mvarBusyStatus = vData
End Property
Public Property Get BusyStatus() As Double
   BusyStatus = mvarBusyStatus
End Property
Public Property Let ImportanceLevel(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarImportanceLevel <> vData)
   mvarImportanceLevel = vData
End Property
Public Property Get ImportanceLevel() As Double
   ImportanceLevel = mvarImportanceLevel
End Property
Public Property Let LabelID(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarLabelID <> vData)
   mvarLabelID = vData
End Property
Public Property Get LabelID() As Double
   LabelID = mvarLabelID
End Property
Public Property Let ScheduleID(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarScheduleID <> vData)
   mvarScheduleID = vData
End Property
Public Property Get ScheduleID() As Double
   ScheduleID = mvarScheduleID
End Property
Public Property Let IsRecurrenceExceptionDeleted(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIsRecurrenceExceptionDeleted <> vData)
   mvarIsRecurrenceExceptionDeleted = vData
End Property
Public Property Get IsRecurrenceExceptionDeleted() As Double
   IsRecurrenceExceptionDeleted = mvarIsRecurrenceExceptionDeleted
End Property
Public Property Let RExceptionStartTimeOrig(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRExceptionStartTimeOrig <> vData)
   mvarRExceptionStartTimeOrig = vData
End Property
Public Property Get RExceptionStartTimeOrig() As String
   RExceptionStartTimeOrig = mvarRExceptionStartTimeOrig
End Property
Public Property Let RExceptionEndTimeOrig(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRExceptionEndTimeOrig <> vData)
   mvarRExceptionEndTimeOrig = vData
End Property
Public Property Get RExceptionEndTimeOrig() As String
   RExceptionEndTimeOrig = mvarRExceptionEndTimeOrig
End Property
Public Property Let IsAllDayEvent(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIsAllDayEvent <> vData)
   mvarIsAllDayEvent = vData
End Property
Public Property Get IsAllDayEvent() As Double
   IsAllDayEvent = mvarIsAllDayEvent
End Property
Public Property Let IsMeeting(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIsMeeting <> vData)
   mvarIsMeeting = vData
End Property
Public Property Get IsMeeting() As Double
   IsMeeting = mvarIsMeeting
End Property
Public Property Let IsPrivate(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIsPrivate <> vData)
   mvarIsPrivate = vData
End Property
Public Property Get IsPrivate() As Double
   IsPrivate = mvarIsPrivate
End Property
Public Property Let IsReminder(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIsReminder <> vData)
   mvarIsReminder = vData
End Property
Public Property Get IsReminder() As Double
   IsReminder = mvarIsReminder
End Property
Public Property Let ReminderMinutesBeforeStart(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarReminderMinutesBeforeStart <> vData)
   mvarReminderMinutesBeforeStart = vData
End Property
Public Property Get ReminderMinutesBeforeStart() As Double
   ReminderMinutesBeforeStart = mvarReminderMinutesBeforeStart
End Property
Public Property Let RemainderSoundFile(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRemainderSoundFile <> vData)
   mvarRemainderSoundFile = vData
End Property
Public Property Get RemainderSoundFile() As String
   RemainderSoundFile = mvarRemainderSoundFile
End Property
Public Property Let CustomPropertiesXMLData(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCustomPropertiesXMLData <> vData)
   mvarCustomPropertiesXMLData = vData
End Property
Public Property Get CustomPropertiesXMLData() As String
   CustomPropertiesXMLData = mvarCustomPropertiesXMLData
End Property
Public Property Let CustomIconsIDs(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCustomIconsIDs <> vData)
   mvarCustomIconsIDs = vData
End Property
Public Property Get CustomIconsIDs() As String
   CustomIconsIDs = mvarCustomIconsIDs
End Property
Public Property Let Created(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCreated <> vData)
   mvarCreated = vData
End Property
Public Property Get Created() As String
   Created = mvarCreated
End Property
Public Property Let Modified(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarModified <> vData)
   mvarModified = vData
End Property
Public Property Get Modified() As String
   Modified = mvarModified
End Property
Public Property Let FLGCONFIRMADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGCONFIRMADO <> vData)
   mvarFLGCONFIRMADO = vData
End Property
Public Property Get FLGCONFIRMADO() As Double
   FLGCONFIRMADO = mvarFLGCONFIRMADO
End Property
Public Property Let FLGCANCELADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGCANCELADO <> vData)
   mvarFLGCANCELADO = vData
End Property
Public Property Get FLGCANCELADO() As Double
   FLGCANCELADO = mvarFLGCANCELADO
End Property
Public Property Let FLGREMARCADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGREMARCADO <> vData)
   mvarFLGREMARCADO = vData
End Property
Public Property Get FLGREMARCADO() As Double
   FLGREMARCADO = mvarFLGREMARCADO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OEVENTOAGENDA Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDEVENTO                       = " & SqlNum(mvarIDEVENTO) & VbNewLine 
   Sql = Sql & " And IDLOJA                     = " & SqlNum(mvarIDLOJA) & VbNewLine 
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
 
   Sql = "Insert Into OEVENTOAGENDA (IDEVENTO, IDLOJA, IDEVENTOREC, IDAGENDA" & VbNewLine 
   Sql = Sql & ", IDCLIENTE, StartDateTime, EndDateTime, RecurrenceState, Subject" & VbNewLine 
   Sql = Sql & ", Location, Body, BusyStatus, ImportanceLevel, LabelID" & VbNewLine 
   Sql = Sql & ", ScheduleID, IsRecurrenceExceptionDeleted, RExceptionStartTimeOrig, RExceptionEndTimeOrig, IsAllDayEvent" & VbNewLine 
   Sql = Sql & ", IsMeeting, IsPrivate, IsReminder, ReminderMinutesBeforeStart, RemainderSoundFile" & VbNewLine 
   Sql = Sql & ", CustomPropertiesXMLData, CustomIconsIDs, Created, Modified, FLGCONFIRMADO" & VbNewLine 
   Sql = Sql & ", FLGCANCELADO, FLGREMARCADO, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   If pAutoId Then
      Sql = Sql &  "(Select isNull(Max(IDEVENTO),0)+1 From OEVENTOAGENDA Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql &  SqlNum(mvarIDEVENTO) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDEVENTOREC) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDAGENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarStartDateTime, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarEndDateTime, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarRecurrenceState) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSubject) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLocation) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarBody) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarBusyStatus) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarImportanceLevel) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarLabelID) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarScheduleID) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIsRecurrenceExceptionDeleted) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarRExceptionStartTimeOrig, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarRExceptionEndTimeOrig, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIsAllDayEvent) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIsMeeting) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIsPrivate) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIsReminder) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarReminderMinutesBeforeStart) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarRemainderSoundFile) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCustomPropertiesXMLData) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCustomIconsIDs) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarCreated, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarModified, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGCONFIRMADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGCANCELADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGREMARCADO) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDEVENTO, Optional Ch_IDLOJA) As String
   Dim Sql As String
 
   Sql = "Delete From OEVENTOAGENDA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDEVENTO) Then Sql = Sql & " IDEVENTO = " & SqlNum(Cstr(Ch_IDEVENTO)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
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
 
   Sql = "Update OEVENTOAGENDA Set " & VbNewLine 
   Sql = Sql & " IDEVENTO                     = " & SqlNum(mvarIDEVENTO) & VbNewLine 
   Sql = Sql & " , IDLOJA                       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDEVENTOREC                  = " & SqlNum(mvarIDEVENTOREC) & VbNewLine 
   Sql = Sql & " , IDAGENDA                     = " & SqlNum(mvarIDAGENDA) & VbNewLine 
   Sql = Sql & " , IDCLIENTE                    = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & " , StartDateTime                = " & SqlDate(mvarStartDateTime, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , EndDateTime                  = " & SqlDate(mvarEndDateTime, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , RecurrenceState              = " & SqlNum(mvarRecurrenceState) & VbNewLine 
   Sql = Sql & " , Subject                      = " & SqlStr(mvarSubject) & VbNewLine 
   Sql = Sql & " , Location                     = " & SqlStr(mvarLocation) & VbNewLine 
   Sql = Sql & " , Body                         = " & SqlStr(mvarBody) & VbNewLine 
   Sql = Sql & " , BusyStatus                   = " & SqlNum(mvarBusyStatus) & VbNewLine 
   Sql = Sql & " , ImportanceLevel              = " & SqlNum(mvarImportanceLevel) & VbNewLine 
   Sql = Sql & " , LabelID                      = " & SqlNum(mvarLabelID) & VbNewLine 
   Sql = Sql & " , ScheduleID                   = " & SqlNum(mvarScheduleID) & VbNewLine 
   Sql = Sql & " , IsRecurrenceExceptionDeleted = " & SqlNum(mvarIsRecurrenceExceptionDeleted) & VbNewLine 
   Sql = Sql & " , RExceptionStartTimeOrig      = " & SqlDate(mvarRExceptionStartTimeOrig, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , RExceptionEndTimeOrig        = " & SqlDate(mvarRExceptionEndTimeOrig, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , IsAllDayEvent                = " & SqlNum(mvarIsAllDayEvent) & VbNewLine 
   Sql = Sql & " , IsMeeting                    = " & SqlNum(mvarIsMeeting) & VbNewLine 
   Sql = Sql & " , IsPrivate                    = " & SqlNum(mvarIsPrivate) & VbNewLine 
   Sql = Sql & " , IsReminder                   = " & SqlNum(mvarIsReminder) & VbNewLine 
   Sql = Sql & " , ReminderMinutesBeforeStart   = " & SqlNum(mvarReminderMinutesBeforeStart) & VbNewLine 
   Sql = Sql & " , RemainderSoundFile           = " & SqlStr(mvarRemainderSoundFile) & VbNewLine 
   Sql = Sql & " , CustomPropertiesXMLData      = " & SqlStr(mvarCustomPropertiesXMLData) & VbNewLine 
   Sql = Sql & " , CustomIconsIDs               = " & SqlStr(mvarCustomIconsIDs) & VbNewLine 
   Sql = Sql & " , Created                      = " & SqlDate(mvarCreated, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , Modified                     = " & SqlDate(mvarModified, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , FLGCONFIRMADO                = " & SqlNum(mvarFLGCONFIRMADO) & VbNewLine 
   Sql = Sql & " , FLGCANCELADO                 = " & SqlNum(mvarFLGCANCELADO) & VbNewLine 
   Sql = Sql & " , FLGREMARCADO                 = " & SqlNum(mvarFLGREMARCADO) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP                   = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP                    = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OEVENTOAGENDA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDEVENTO                       = " & SqlNum(mvarIDEVENTO) & VbNewLine 
   Sql = Sql & " And IDLOJA                     = " & SqlNum(mvarIDLOJA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDEVENTO, Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDEVENTO, IDLOJA, IDEVENTOREC, IDAGENDA" & VbNewLine 
   Sql = Sql & ", IDCLIENTE, StartDateTime, EndDateTime, RecurrenceState, Subject" & VbNewLine 
   Sql = Sql & ", Location, Body, BusyStatus, ImportanceLevel, LabelID" & VbNewLine 
   Sql = Sql & ", ScheduleID, IsRecurrenceExceptionDeleted, RExceptionStartTimeOrig, RExceptionEndTimeOrig, IsAllDayEvent" & VbNewLine 
   Sql = Sql & ", IsMeeting, IsPrivate, IsReminder, ReminderMinutesBeforeStart, RemainderSoundFile" & VbNewLine 
   Sql = Sql & ", CustomPropertiesXMLData, CustomIconsIDs, Created, Modified, FLGCONFIRMADO" & VbNewLine 
   Sql = Sql & ", FLGCANCELADO, FLGREMARCADO, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OEVENTOAGENDA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDEVENTO)  And IsMissing(Ch_IDLOJA)  Then 
      Sql = Sql & " IDEVENTO = " & SqlNum(mvarIDEVENTO) & " AND " & VbNewLine 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDEVENTO) Then Sql = Sql & " IDEVENTO = " & SqlNum(Cstr(Ch_IDEVENTO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDEVENTO, Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDEVENTO, Ch_IDLOJA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDEVENTO = XVal(!IDEVENTO & "")
      mvarIDLOJA = XVal(!IDLOJA & "")
      mvarIDEVENTOREC = XVal(!IDEVENTOREC & "")
      mvarIDAGENDA = XVal(!IDAGENDA & "")
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
      mvarStartDateTime = xDate(!StartDateTime & "", True)
      mvarEndDateTime = xDate(!EndDateTime & "", True)
      mvarRecurrenceState = XVal(!RecurrenceState & "")
      mvarSubject = !Subject & ""
      mvarLocation = !Location & ""
      mvarBody = !Body & ""
      mvarBusyStatus = XVal(!BusyStatus & "")
      mvarImportanceLevel = XVal(!ImportanceLevel & "")
      mvarLabelID = XVal(!LabelID & "")
      mvarScheduleID = XVal(!ScheduleID & "")
      mvarIsRecurrenceExceptionDeleted = XVal(!IsRecurrenceExceptionDeleted & "")
      mvarRExceptionStartTimeOrig = xDate(!RExceptionStartTimeOrig & "", True)
      mvarRExceptionEndTimeOrig = xDate(!RExceptionEndTimeOrig & "", True)
      mvarIsAllDayEvent = XVal(!IsAllDayEvent & "")
      mvarIsMeeting = XVal(!IsMeeting & "")
      mvarIsPrivate = XVal(!IsPrivate & "")
      mvarIsReminder = XVal(!IsReminder & "")
      mvarReminderMinutesBeforeStart = XVal(!ReminderMinutesBeforeStart & "")
      mvarRemainderSoundFile = !RemainderSoundFile & ""
      mvarCustomPropertiesXMLData = !CustomPropertiesXMLData & ""
      mvarCustomIconsIDs = !CustomIconsIDs & ""
      mvarCreated = xDate(!Created & "", True)
      mvarModified = xDate(!Modified & "", True)
      mvarFLGCONFIRMADO = XVal(!FLGCONFIRMADO & "")
      mvarFLGCANCELADO = XVal(!FLGCANCELADO & "")
      mvarFLGREMARCADO = XVal(!FLGREMARCADO & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDEVENTO, mvarIDLOJA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDEVENTO As String, Ch_IDLOJA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OEVENTOAGENDA Set "
   Sql = Sql & " IDEVENTO = " & SqlNum(Ch_IDEVENTO)
   Sql = Sql & " , IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " Where 
   Sql = Sql & " IDEVENTO = " & SqlNum(mvarIDEVENTO)
   Sql = Sql & " and IDLOJA = " & SqlNum(mvarIDLOJA)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
