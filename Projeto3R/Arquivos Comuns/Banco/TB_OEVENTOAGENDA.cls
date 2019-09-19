VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
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
Private mvarFLGAVALIACAO                 As Double
Private mvarIDSALA                       As Double
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
 
   Sql = "Update OEVENTOAGENDA Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDEVENTO                       = " & SqlNum(mvarIDEVENTO) & vbNewLine
   Sql = Sql & " And IDLOJA                     = " & SqlNum(mvarIDLOJA) & vbNewLine
   If Not mvarXDb Is Nothing Then
      If mvarXDb.Conectado Then
         If mvarXDb.Executa(Sql, True) Then
            mvarALTERSTAMP = vData
         End If
      End If
   End If
End Property
Public Property Let FLGAVALIACAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGAVALIACAO <> vData)
   mvarFLGAVALIACAO = vData
End Property
Public Property Get FLGAVALIACAO() As Double
   FLGAVALIACAO = mvarFLGAVALIACAO
End Property
Public Property Let IDSALA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSALA <> vData)
   mvarIDSALA = vData
End Property
Public Property Get IDSALA() As Double
   IDSALA = mvarIDSALA
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
Public Property Get QryInsert(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into OEVENTOAGENDA (IDEVENTO, IDLOJA, IDEVENTOREC, IDAGENDA" & vbNewLine
   Sql = Sql & ", IDCLIENTE, StartDateTime, EndDateTime, RecurrenceState, Subject" & vbNewLine
   Sql = Sql & ", Location, Body, BusyStatus, ImportanceLevel, LabelID" & vbNewLine
   Sql = Sql & ", ScheduleID, IsRecurrenceExceptionDeleted, RExceptionStartTimeOrig, RExceptionEndTimeOrig, IsAllDayEvent" & vbNewLine
   Sql = Sql & ", IsMeeting, IsPrivate, IsReminder, ReminderMinutesBeforeStart, RemainderSoundFile" & vbNewLine
   Sql = Sql & ", CustomPropertiesXMLData, CustomIconsIDs, Created, Modified, FLGCONFIRMADO" & vbNewLine
   Sql = Sql & ", FLGCANCELADO, FLGREMARCADO, FLGAVALIACAO, IDSALA, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   If pAutoId Then
      Sql = Sql & "(Select isNull(Max(IDEVENTO),0)+1 From OEVENTOAGENDA Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & SqlNum(mvarIDEVENTO) & vbNewLine
   End If
   Sql = Sql & ", " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDEVENTOREC = 0, "Null", SqlNum(mvarIDEVENTOREC)) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDAGENDA = 0, "Null", SqlNum(mvarIDAGENDA)) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDCLIENTE = 0, "Null", SqlNum(mvarIDCLIENTE)) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarStartDateTime, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarEndDateTime, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarRecurrenceState) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarSubject) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarLocation) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarBody) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarBusyStatus) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarImportanceLevel) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarLabelID) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarScheduleID) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIsRecurrenceExceptionDeleted) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarRExceptionStartTimeOrig, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarRExceptionEndTimeOrig, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIsAllDayEvent) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIsMeeting) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIsPrivate) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIsReminder) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarReminderMinutesBeforeStart) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarRemainderSoundFile) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCustomPropertiesXMLData) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCustomIconsIDs) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarCreated, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarModified, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGCONFIRMADO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGCANCELADO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGREMARCADO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGAVALIACAO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDSALA) & vbNewLine
   If pSinc Then
      Sql = Sql & ", " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If
   

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDEVENTO, Optional Ch_IDLOJA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OEVENTOAGENDA" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDEVENTO) And IsMissing(Ch_IDLOJA) Then
      Sql = Sql & " IDEVENTO = " & SqlNum(mvarIDEVENTO) & " AND " & vbNewLine
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDEVENTO) Then Sql = Sql & " IDEVENTO = " & SqlNum(CStr(Ch_IDEVENTO)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = ""
   sTag = sTag & "|IDEVENTO = " & mvarIDEVENTO
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDEVENTOREC = " & mvarIDEVENTOREC
   sTag = sTag & "|IDAGENDA = " & mvarIDAGENDA
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|StartDateTime = " & mvarStartDateTime
   sTag = sTag & "|EndDateTime = " & mvarEndDateTime
   sTag = sTag & "|RecurrenceState = " & mvarRecurrenceState
   sTag = sTag & "|Subject = " & mvarSubject
   sTag = sTag & "|Location = " & mvarLocation
   sTag = sTag & "|Body = " & mvarBody
   sTag = sTag & "|BusyStatus = " & mvarBusyStatus
   sTag = sTag & "|ImportanceLevel = " & mvarImportanceLevel
   sTag = sTag & "|LabelID = " & mvarLabelID
   sTag = sTag & "|ScheduleID = " & mvarScheduleID
   sTag = sTag & "|IsRecurrenceExceptionDeleted = " & mvarIsRecurrenceExceptionDeleted
   sTag = sTag & "|RExceptionStartTimeOrig = " & mvarRExceptionStartTimeOrig
   sTag = sTag & "|RExceptionEndTimeOrig = " & mvarRExceptionEndTimeOrig
   sTag = sTag & "|IsAllDayEvent = " & mvarIsAllDayEvent
   sTag = sTag & "|IsMeeting = " & mvarIsMeeting
   sTag = sTag & "|IsPrivate = " & mvarIsPrivate
   sTag = sTag & "|IsReminder = " & mvarIsReminder
   sTag = sTag & "|ReminderMinutesBeforeStart = " & mvarReminderMinutesBeforeStart
   sTag = sTag & "|RemainderSoundFile = " & mvarRemainderSoundFile
   sTag = sTag & "|CustomPropertiesXMLData = " & mvarCustomPropertiesXMLData
   sTag = sTag & "|CustomIconsIDs = " & mvarCustomIconsIDs
   sTag = sTag & "|Created = " & mvarCreated
   sTag = sTag & "|Modified = " & mvarModified
   sTag = sTag & "|FLGCONFIRMADO = " & mvarFLGCONFIRMADO
   sTag = sTag & "|FLGCANCELADO = " & mvarFLGCANCELADO
   sTag = sTag & "|FLGREMARCADO = " & mvarFLGREMARCADO
   sTag = sTag & "|FLGAVALIACAO = " & mvarFLGAVALIACAO
   sTag = sTag & "|IDSALA = " & mvarIDSALA
   sTag = sTag & "|Where = " & IIf(IsMissing(Ch_WHERE), "", Ch_WHERE)
   sTag = sTag & "|"

   MyDelRow.QUERY = Sql & ";"
   MyDelRow.TAG = sTag
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update OEVENTOAGENDA Set " & vbNewLine
   Sql = Sql & " IDEVENTO                     = " & SqlNum(mvarIDEVENTO) & vbNewLine
   Sql = Sql & " , IDLOJA                       = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDEVENTOREC                  = " & IIf(mvarIDEVENTOREC = 0, "Null", SqlNum(mvarIDEVENTOREC)) & vbNewLine
   Sql = Sql & " , IDAGENDA                     = " & IIf(mvarIDAGENDA = 0, "Null", SqlNum(mvarIDAGENDA)) & vbNewLine
   Sql = Sql & " , IDCLIENTE                    = " & IIf(mvarIDCLIENTE = 0, "Null", SqlNum(mvarIDCLIENTE)) & vbNewLine
   Sql = Sql & " , StartDateTime                = " & SqlDate(mvarStartDateTime, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , EndDateTime                  = " & SqlDate(mvarEndDateTime, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , RecurrenceState              = " & SqlNum(mvarRecurrenceState) & vbNewLine
   Sql = Sql & " , Subject                      = " & SqlStr(mvarSubject) & vbNewLine
   Sql = Sql & " , Location                     = " & SqlStr(mvarLocation) & vbNewLine
   Sql = Sql & " , Body                         = " & SqlStr(mvarBody) & vbNewLine
   Sql = Sql & " , BusyStatus                   = " & SqlNum(mvarBusyStatus) & vbNewLine
   Sql = Sql & " , ImportanceLevel              = " & SqlNum(mvarImportanceLevel) & vbNewLine
   Sql = Sql & " , LabelID                      = " & SqlNum(mvarLabelID) & vbNewLine
   Sql = Sql & " , ScheduleID                   = " & SqlNum(mvarScheduleID) & vbNewLine
   Sql = Sql & " , IsRecurrenceExceptionDeleted = " & SqlNum(mvarIsRecurrenceExceptionDeleted) & vbNewLine
   Sql = Sql & " , RExceptionStartTimeOrig      = " & SqlDate(mvarRExceptionStartTimeOrig, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , RExceptionEndTimeOrig        = " & SqlDate(mvarRExceptionEndTimeOrig, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , IsAllDayEvent                = " & SqlNum(mvarIsAllDayEvent) & vbNewLine
   Sql = Sql & " , IsMeeting                    = " & SqlNum(mvarIsMeeting) & vbNewLine
   Sql = Sql & " , IsPrivate                    = " & SqlNum(mvarIsPrivate) & vbNewLine
   Sql = Sql & " , IsReminder                   = " & SqlNum(mvarIsReminder) & vbNewLine
   Sql = Sql & " , ReminderMinutesBeforeStart   = " & SqlNum(mvarReminderMinutesBeforeStart) & vbNewLine
   Sql = Sql & " , RemainderSoundFile           = " & SqlStr(mvarRemainderSoundFile) & vbNewLine
   Sql = Sql & " , CustomPropertiesXMLData      = " & SqlStr(mvarCustomPropertiesXMLData) & vbNewLine
   Sql = Sql & " , CustomIconsIDs               = " & SqlStr(mvarCustomIconsIDs) & vbNewLine
   Sql = Sql & " , Created                      = " & SqlDate(mvarCreated, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , Modified                     = " & SqlDate(mvarModified, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , FLGCONFIRMADO                = " & SqlNum(mvarFLGCONFIRMADO) & vbNewLine
   Sql = Sql & " , FLGCANCELADO                 = " & SqlNum(mvarFLGCANCELADO) & vbNewLine
   Sql = Sql & " , FLGREMARCADO                 = " & SqlNum(mvarFLGREMARCADO) & vbNewLine
   Sql = Sql & " , FLGAVALIACAO                 = " & SqlNum(mvarFLGAVALIACAO) & vbNewLine
   Sql = Sql & " , IDSALA                       = " & SqlNum(mvarIDSALA) & vbNewLine
   If pSinc Then
      Sql = Sql & ", ALTERSTAMP                   = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP                   = 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", TIMESTAMP                    = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP                    = GetDate()" & vbNewLine
   End If
   
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OEVENTOAGENDA'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDEVENTO                       = " & SqlNum(mvarIDEVENTO) & vbNewLine
   Sql = Sql & " And IDLOJA                     = " & SqlNum(mvarIDLOJA) & vbNewLine

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & vbNewLine
   Sql = Sql & Me.QryUpDate(pSinc:=pSinc)
   Sql = Sql & " Else " & vbNewLine
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId, pSinc:=pSinc)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect(Optional Ch_IDEVENTO, Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDEVENTO, IDLOJA, IDEVENTOREC, IDAGENDA" & vbNewLine
   Sql = Sql & ", IDCLIENTE, StartDateTime, EndDateTime, RecurrenceState, Subject" & vbNewLine
   Sql = Sql & ", Location, Body, BusyStatus, ImportanceLevel, LabelID" & vbNewLine
   Sql = Sql & ", ScheduleID, IsRecurrenceExceptionDeleted, RExceptionStartTimeOrig, RExceptionEndTimeOrig, IsAllDayEvent" & vbNewLine
   Sql = Sql & ", IsMeeting, IsPrivate, IsReminder, ReminderMinutesBeforeStart, RemainderSoundFile" & vbNewLine
   Sql = Sql & ", CustomPropertiesXMLData, CustomIconsIDs, Created, Modified, FLGCONFIRMADO" & vbNewLine
   Sql = Sql & ", FLGCANCELADO, FLGREMARCADO, FLGAVALIACAO, IDSALA, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From OEVENTOAGENDA" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDEVENTO) And IsMissing(Ch_IDLOJA) Then
      Sql = Sql & " IDEVENTO = " & SqlNum(mvarIDEVENTO) & " AND " & vbNewLine
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDEVENTO) Then Sql = Sql & " IDEVENTO = " & SqlNum(CStr(Ch_IDEVENTO)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar(Optional Ch_IDEVENTO, Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDEVENTO, Ch_IDLOJA, Ch_WHERE, Ch_ORDERBY)
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
    Call ShowError(Sql)
    Pesquisar = False
End Function
Public Sub Popula(Optional pRcSet)
   If IsMissing(pRcSet) Then Set pRcSet = mvarRS
   With pRcSet
      mvarIDEVENTO = xVal(!IDEVENTO & "")
      mvarIDLOJA = xVal(!IDLOJA & "")
      mvarIDEVENTOREC = xVal(!IDEVENTOREC & "")
      mvarIDAGENDA = xVal(!IDAGENDA & "")
      mvarIDCLIENTE = xVal(!IDCLIENTE & "")
      mvarStartDateTime = xDate(!StartDateTime & "", True)
      mvarEndDateTime = xDate(!EndDateTime & "", True)
      mvarRecurrenceState = xVal(!RecurrenceState & "")
      mvarSubject = !Subject & ""
      mvarLocation = !Location & ""
      mvarBody = !Body & ""
      mvarBusyStatus = xVal(!BusyStatus & "")
      mvarImportanceLevel = xVal(!ImportanceLevel & "")
      mvarLabelID = xVal(!LabelID & "")
      mvarScheduleID = xVal(!ScheduleID & "")
      mvarIsRecurrenceExceptionDeleted = xVal(!IsRecurrenceExceptionDeleted & "")
      mvarRExceptionStartTimeOrig = xDate(!RExceptionStartTimeOrig & "", True)
      mvarRExceptionEndTimeOrig = xDate(!RExceptionEndTimeOrig & "", True)
      mvarIsAllDayEvent = xVal(!IsAllDayEvent & "")
      mvarIsMeeting = xVal(!IsMeeting & "")
      mvarIsPrivate = xVal(!IsPrivate & "")
      mvarIsReminder = xVal(!IsReminder & "")
      mvarReminderMinutesBeforeStart = xVal(!ReminderMinutesBeforeStart & "")
      mvarRemainderSoundFile = !RemainderSoundFile & ""
      mvarCustomPropertiesXMLData = !CustomPropertiesXMLData & ""
      mvarCustomIconsIDs = !CustomIconsIDs & ""
      mvarCreated = xDate(!Created & "", True)
      mvarModified = xDate(!Modified & "", True)
      mvarFLGCONFIRMADO = xVal(!FLGCONFIRMADO & "")
      mvarFLGCANCELADO = xVal(!FLGCANCELADO & "")
      mvarFLGREMARCADO = xVal(!FLGREMARCADO & "")
      mvarFLGAVALIACAO = xVal(!FLGAVALIACAO & "")
      mvarIDSALA = xVal(!IDSALA & "")
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDEVENTO = 0
   mvarIDLOJA = 0
   mvarIDEVENTOREC = 0
   mvarIDAGENDA = 0
   mvarIDCLIENTE = 0
   mvarStartDateTime = ""
   mvarEndDateTime = ""
   mvarRecurrenceState = 0
   mvarSubject = ""
   mvarLocation = ""
   mvarBody = ""
   mvarBusyStatus = 0
   mvarImportanceLevel = 0
   mvarLabelID = 0
   mvarScheduleID = 0
   mvarIsRecurrenceExceptionDeleted = 0
   mvarRExceptionStartTimeOrig = ""
   mvarRExceptionEndTimeOrig = ""
   mvarIsAllDayEvent = 0
   mvarIsMeeting = 0
   mvarIsPrivate = 0
   mvarIsReminder = 0
   mvarReminderMinutesBeforeStart = 0
   mvarRemainderSoundFile = ""
   mvarCustomPropertiesXMLData = ""
   mvarCustomIconsIDs = ""
   mvarCreated = ""
   mvarModified = ""
   mvarFLGCONFIRMADO = 0
   mvarFLGCANCELADO = 0
   mvarFLGREMARCADO = 0
   mvarFLGAVALIACAO = 0
   mvarIDSALA = 0
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
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId As Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDEVENTO, mvarIDLOJA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDEVENTO As String, Ch_IDLOJA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OEVENTOAGENDA Set "
   Sql = Sql & " IDEVENTO = " & SqlNum(Ch_IDEVENTO)
   Sql = Sql & " , IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " Where "
   Sql = Sql & " IDEVENTO = " & SqlNum(mvarIDEVENTO)
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
