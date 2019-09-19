VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLDataHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CodeJockCalendarDataType
    cjCalendarData_Unknown = 0
    cjCalendarData_Memory = 1
    cjCalendarData_Access = 2
    cjCalendarData_MAPI = 3
    cjCalendarData_SQLServer = 4
    cjCalendarData_MySQL = 5
End Enum

Dim m_pCalendar As CalendarControl
Attribute m_pCalendar.VB_VarHelpID = -1

'======================================================================
Public Sub SetCalendar(ByVal pCalendar As CalendarControl)
    Set m_pCalendar = pCalendar
End Sub

Private Function RoundNum2D(nNumber As Long) As String
    Debug.Assert nNumber < 100
    RoundNum2D = IIf(nNumber < 10, "0", "") & nNumber
End Function


Public Function MakeSQLDateTime(dtDate As Date) As String
    Dim strDate As String
    
    ' "{ts'%Y-%m-%d %H:%M:%S'}"
    strDate = "{ts'" & Year(dtDate) & "-"
    strDate = strDate + RoundNum2D(Month(dtDate)) & "-"
    strDate = strDate + RoundNum2D(Day(dtDate)) & " "
    
    strDate = strDate + RoundNum2D(Hour(dtDate)) & ":"
    strDate = strDate + RoundNum2D(Minute(dtDate)) & ":"
    strDate = strDate + RoundNum2D(Second(dtDate)) & "'}"
    
    MakeSQLDateTime = strDate
End Function

Public Function MakeSQLstr(ByVal strVB As String) As String
    strVB = Replace(strVB, "\", "\\")
    strVB = Replace(strVB, """", "\""")
    
    strVB = """" & strVB & """"
    
    MakeSQLstr = strVB
End Function
    
Public Function CreateEventFromRS(pEventRS As ADODB.Recordset, ByVal bRException) As CalendarEvent
    On Error GoTo err1
    
    Set CreateEventFromRS = Nothing
    
    Dim pEvent As CalendarEvent
    Dim nEventID As Long
        
    If pEventRS.EOF Then
        Exit Function
    End If
        
    nEventID = pEventRS("EventID")
    Set pEvent = m_pCalendar.DataProvider.CreateEventEx(nEventID)
    
    If pEvent Is Nothing Then
        Exit Function
    End If
        
    pEvent.Subject = pEventRS("Subject")
    pEvent.Location = pEventRS("Location")
    pEvent.Body = pEventRS("Body")
    
    pEvent.MeetingFlag = IIf(pEventRS("IsMeeting") <> 0, True, False)
    pEvent.PrivateFlag = IIf(pEventRS("IsPrivate") <> 0, True, False)
    
    pEvent.Label = pEventRS("LabelID")
    pEvent.BusyStatus = pEventRS("BusyStatus")
    pEvent.Importance = pEventRS("ImportanceLevel")
    
    pEvent.StartTime = pEventRS("StartDateTime")
    pEvent.EndTime = pEventRS("EndDateTime")
        
    pEvent.AllDayEvent = IIf(pEventRS("IsAllDayEvent") <> 0, True, False)

    pEvent.Reminder = IIf(pEventRS("IsReminder") <> 0, True, False)
    pEvent.ReminderMinutesBeforeStart = pEventRS("ReminderMinutesBeforeStart")
    pEvent.ReminderSoundFile = pEventRS("RemainderSoundFile")
    
    pEvent.CustomProperties.LoadFromString pEventRS("CustomPropertiesXMLData")
    pEvent.CustomIcons.LoadFromString pEventRS("CustomIconsIDs")
    
    pEvent.ScheduleID = pEventRS("ScheduleID")
    
    If bRException Then
        pEvent.MakeAsRException
        
        pEvent.RExceptionStartTimeOrig = pEventRS("RExceptionStartTimeOrig")
        pEvent.RExceptionEndTimeOrig = pEventRS("RExceptionEndTimeOrig")
        pEvent.RExceptionDeleted = IIf(pEventRS("ISRecurrenceExceptionDeleted") <> 0, True, False)
    End If
    
    If Not bRException Then
        
        ' "process_RecurrenceState" and "process_RecurrencePatternID" properties
        ' are used to process master events.
        '
        ' If they are set and RecurrenceStaie is Master Data provider will
        ' fier DoReadRPattern event and make event as Master.
        ' And it will also generate ocurrences for RetrieveDayEvents method.
        '
        ' Thise properties are temporary and they will be removed by data provider.
        '
        ' If these properties are not set data provider expect that master event
        ' is already compleated - CreateRecurrence method is called and
        ' Recurrence pattern is set.
        '
        ' This mechanism is usefull for DB data providers, when events and patterns
        ' are stored separately (in tables).
        ' But if events stored in some memory collection or array
        ' it should not be used because master event store recurrence pattern inside.
        '
        pEvent.CustomProperties("process_RecurrenceState") = pEventRS("RecurrenceState").Value
        pEvent.CustomProperties("process_RecurrencePatternID") = pEventRS("RecurrencePatternID").Value
    End If
    
    Set CreateEventFromRS = pEvent
    Exit Function
    
err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot read event from recordset: " & Err.Description
        Debug.Assert False
    End If
End Function

Public Sub PutEventToRS(pEvent As CalendarEvent, pEventRS As ADODB.Recordset)

    On Error GoTo err1
    
    pEventRS("Subject") = pEvent.Subject
    pEventRS("Location") = pEvent.Location
    pEventRS("Body") = pEvent.Body
    
    pEventRS("IsMeeting") = IIf(pEvent.MeetingFlag, 1, 0)
    pEventRS("IsPrivate") = IIf(pEvent.PrivateFlag, 1, 0)
    
    pEventRS("LabelID") = pEvent.Label
    pEventRS("BusyStatus") = pEvent.BusyStatus
    pEventRS("ImportanceLevel") = pEvent.Importance
    
    pEventRS("StartDateTime") = pEvent.StartTime
    pEventRS("EndDateTime") = pEvent.EndTime
        
    pEventRS("IsAllDayEvent") = IIf(pEvent.AllDayEvent, 1, 0)

    pEventRS("IsReminder") = IIf(pEvent.Reminder, 1, 0)
    pEventRS("ReminderMinutesBeforeStart") = pEvent.ReminderMinutesBeforeStart
    pEventRS("RemainderSoundFile") = pEvent.ReminderSoundFile
            
    pEventRS("RecurrenceState") = pEvent.RecurrenceState
    
    If pEvent.RecurrenceState = xtpCalendarRecurrenceMaster Or _
       pEvent.RecurrenceState = xtpCalendarRecurrenceException Then
        
        pEventRS("RecurrencePatternID") = pEvent.RecurrencePattern.Id
    Else
        pEventRS("RecurrencePatternID") = 0
    End If
    
    pEventRS("RExceptionStartTimeOrig") = pEvent.RExceptionStartTimeOrig
    pEventRS("RExceptionEndTimeOrig") = pEvent.RExceptionEndTimeOrig
    pEventRS("ISRecurrenceExceptionDeleted") = IIf(pEvent.RExceptionDeleted, 1, 0)
        
    pEventRS("CustomPropertiesXMLData") = pEvent.CustomProperties.SaveToString
    pEventRS("CustomIconsIDs") = pEvent.CustomIcons.SaveToString
            
    pEventRS("ScheduleID") = pEvent.ScheduleID
    
err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot put event to recordset: " & Err.Description
        Debug.Assert False
    End If
End Sub

Public Function CreateRPatternFromRS(pRPatternRS As ADODB.Recordset) As CalendarRecurrencePattern
    On Error GoTo err1
    
    Set CreateRPatternFromRS = Nothing
    
    Dim pRPattern As CalendarRecurrencePattern
    Dim nPatternID As Long
        
    If pRPatternRS.EOF Then
        Exit Function
    End If
        
    nPatternID = pRPatternRS("RecurrencePatternID")
    Set pRPattern = m_pCalendar.DataProvider.CreateRecurrencePattern(nPatternID)
    
    If pRPattern Is Nothing Then
        Exit Function
    End If
        
    pRPattern.MasterEventId = pRPatternRS("MasterEventID")
    
    pRPattern.StartTime = pRPatternRS("EventStartTime")
    pRPattern.DurationMinutes = pRPatternRS("EventDuration")
    
    pRPattern.StartDate = pRPatternRS("PatternStartDate")
    
    pRPattern.EndMethod = pRPatternRS("PatternEndMethod")
    
    If pRPattern.EndMethod = xtpCalendarPatternEndDate Then
        pRPattern.EndDate = pRPatternRS("PatternEndDate")
    ElseIf pRPattern.EndMethod = xtpCalendarPatternEndAfterOccurrences Then
        pRPattern.EndAfterOccurrences = pRPatternRS("PatternEndAfterOccurrences")
    Else
        Debug.Assert pRPattern.EndMethod = xtpCalendarPatternEndNoDate
    End If

    pRPattern.Options.Data1 = pRPatternRS("OptionsData1")
    pRPattern.Options.Data2 = pRPatternRS("OptionsData2")
    pRPattern.Options.Data3 = pRPatternRS("OptionsData3")
    pRPattern.Options.Data4 = pRPatternRS("OptionsData4")
    
    pRPattern.CustomProperties.LoadFromString pRPatternRS("CustomPropertiesXMLData")
        
    ReadRPatternExceptions pRPattern, pRPatternRS.ActiveConnection
        
    Set CreateRPatternFromRS = pRPattern
    Exit Function
    
err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot read recurrence pattern from recordset: " & Err.Description
        Debug.Assert False
    End If
End Function

Public Sub PutRPatternToRS(pRPattern As CalendarRecurrencePattern, pRPatternRS As ADODB.Recordset)
    On Error GoTo err1
    
    pRPatternRS("MasterEventID") = pRPattern.MasterEventId
    
    pRPatternRS("EventStartTime") = pRPattern.StartTime
    pRPatternRS("EventDuration") = pRPattern.DurationMinutes
    
    pRPatternRS("PatternStartDate") = pRPattern.StartDate
    
    pRPatternRS("PatternEndMethod") = pRPattern.EndMethod
    pRPatternRS("PatternEndDate") = pRPattern.EndDate
    pRPatternRS("PatternEndAfterOccurrences") = pRPattern.EndAfterOccurrences
    
    pRPatternRS("OptionsData1") = pRPattern.Options.Data1
    pRPatternRS("OptionsData2") = pRPattern.Options.Data2
    pRPatternRS("OptionsData3") = pRPattern.Options.Data3
    pRPatternRS("OptionsData4") = pRPattern.Options.Data4
    
    pRPatternRS("CustomPropertiesXMLData") = pRPattern.CustomProperties.SaveToString
    
    Exit Sub
    
err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot read recurrence pattern from recordset: " & Err.Description
        Debug.Assert False
    End If
End Sub

Public Sub ReadRPatternExceptions(pPattern As CalendarRecurrencePattern, pConnection As ADODB.Connection)

    Dim nPatternID As Long
    Dim strSQL As String
    
    nPatternID = pPattern.Id
        
    strSQL = "SELECT * FROM CalendarEvents WHERE " & vbCrLf
    
    strSQL = strSQL + " RecurrenceState = " & xtpCalendarRecurrenceException & " AND " & vbCrLf
    strSQL = strSQL + " RecurrencePatternID = " & nPatternID
    
    Dim rsEvents As ADODB.Recordset
    Set rsEvents = pConnection.Execute(strSQL)
    
    Dim pEvent As CalendarEvent
        
    While Not rsEvents.EOF
        Set pEvent = CreateEventFromRS(rsEvents, True)
        If Not pEvent Is Nothing Then
            pPattern.SetException pEvent
        End If
        
        rsEvents.MoveNext
    Wend
       
End Sub

Public Function MakeUpdateEventSQL(pEvent As CalendarEvent) As String
    Dim strSQL As String
  
    strSQL = "UPDATE CalendarEvents SET " & vbCrLf
    strSQL = strSQL + "Subject = " & MakeSQLstr(pEvent.Subject) & ", " & vbCrLf
    strSQL = strSQL + "Location = " & MakeSQLstr(pEvent.Location) & ", " & vbCrLf
    strSQL = strSQL + "Body = " & MakeSQLstr(pEvent.Body) & ", " & vbCrLf
       
    strSQL = strSQL + "IsMeeting = " & IIf(pEvent.MeetingFlag, 1, 0) & ", " & vbCrLf
    strSQL = strSQL + "IsPrivate = " & IIf(pEvent.PrivateFlag, 1, 0) & ", " & vbCrLf
    
    strSQL = strSQL + "LabelID = " & pEvent.Label & ", " & vbCrLf
    strSQL = strSQL + "BusyStatus = " & pEvent.BusyStatus & ", " & vbCrLf
    strSQL = strSQL + "ImportanceLevel = " & pEvent.Importance & ", " & vbCrLf
    
    strSQL = strSQL + "StartDateTime = " & MakeSQLDateTime(pEvent.StartTime) & ", " & vbCrLf
    strSQL = strSQL + "EndDateTime = " & MakeSQLDateTime(pEvent.EndTime) & ", " & vbCrLf
        
    strSQL = strSQL + "IsAllDayEvent = " & IIf(pEvent.AllDayEvent, 1, 0) & ", " & vbCrLf

    strSQL = strSQL + "IsReminder = " & IIf(pEvent.Reminder, 1, 0) & ", " & vbCrLf
    strSQL = strSQL + "ReminderMinutesBeforeStart = " & pEvent.ReminderMinutesBeforeStart & ", " & vbCrLf
    strSQL = strSQL + "RemainderSoundFile = " & MakeSQLstr(pEvent.ReminderSoundFile) & ", " & vbCrLf
            
    strSQL = strSQL + "RecurrenceState = " & pEvent.RecurrenceState & ", " & vbCrLf
    
    If pEvent.RecurrenceState = xtpCalendarRecurrenceMaster Or _
       pEvent.RecurrenceState = xtpCalendarRecurrenceException Then
        
        strSQL = strSQL + "RecurrencePatternID = " & pEvent.RecurrencePattern.Id & ", " & vbCrLf
    Else
        strSQL = strSQL + "RecurrencePatternID = 0, " & vbCrLf
    End If
    
    strSQL = strSQL + "RExceptionStartTimeOrig = " & MakeSQLDateTime(pEvent.RExceptionStartTimeOrig) & ", " & vbCrLf
    strSQL = strSQL + "RExceptionEndTimeOrig = " & MakeSQLDateTime(pEvent.RExceptionEndTimeOrig) & ", " & vbCrLf
    strSQL = strSQL + "ISRecurrenceExceptionDeleted = " & IIf(pEvent.RExceptionDeleted, 1, 0) & ", " & vbCrLf
    
    strSQL = strSQL + "CustomPropertiesXMLData = " & MakeSQLstr(pEvent.CustomProperties.SaveToString) & " " & vbCrLf
    strSQL = strSQL + "CustomIconsIDs = " & MakeSQLstr(pEvent.CustomIcons.SaveToString) & ", " & vbCrLf
    
    strSQL = strSQL + "ScheduleID = " & pEvent.ScheduleID & " " & vbCrLf
   
    strSQL = strSQL + "WHERE EventID = " & pEvent.Id & ";"
    
    ''--------------------------------------------
    MakeUpdateEventSQL = strSQL
End Function

Public Function MakeCreateEventSQL(pEvent As CalendarEvent) As String
    Dim strSQL As String
  
    strSQL = "INSERT INTO CalendarEvents ("
    strSQL = strSQL + "Subject, "
    strSQL = strSQL + "Location, "
    strSQL = strSQL + "Body, "
       
    strSQL = strSQL + "IsMeeting, "
    strSQL = strSQL + "IsPrivate, "
    
    strSQL = strSQL + "LabelID, "
    strSQL = strSQL + "BusyStatus, "
    strSQL = strSQL + "ImportanceLevel, "
    
    strSQL = strSQL + "StartDateTime, "
    strSQL = strSQL + "EndDateTime, "
        
    strSQL = strSQL + "IsAllDayEvent, "

    strSQL = strSQL + "IsReminder, "
    strSQL = strSQL + "ReminderMinutesBeforeStart, "
    strSQL = strSQL + "RemainderSoundFile, "
            
    strSQL = strSQL + "RecurrenceState, "
    strSQL = strSQL + "RecurrencePatternID, "
    
    strSQL = strSQL + "RExceptionStartTimeOrig, "
    strSQL = strSQL + "RExceptionEndTimeOrig, "
    strSQL = strSQL + "ISRecurrenceExceptionDeleted, "
    
    strSQL = strSQL + "CustomPropertiesXMLData, "
    strSQL = strSQL + "CustomIconsIDs, "
    strSQL = strSQL + "ScheduleID) " & vbCrLf

    strSQL = strSQL + "VALUES("
    
    strSQL = strSQL + MakeSQLstr(pEvent.Subject) & ", " & vbCrLf
    strSQL = strSQL + MakeSQLstr(pEvent.Location) & ", " & vbCrLf
    strSQL = strSQL + MakeSQLstr(pEvent.Body) & ", " & vbCrLf
       
    strSQL = strSQL & IIf(pEvent.MeetingFlag, 1, 0) & ", " & vbCrLf
    strSQL = strSQL & IIf(pEvent.PrivateFlag, 1, 0) & ", " & vbCrLf
    
    strSQL = strSQL & pEvent.Label & ", " & vbCrLf
    strSQL = strSQL & pEvent.BusyStatus & ", " & vbCrLf
    strSQL = strSQL & pEvent.Importance & ", " & vbCrLf
    
    strSQL = strSQL & MakeSQLDateTime(pEvent.StartTime) & ", " & vbCrLf
    strSQL = strSQL & MakeSQLDateTime(pEvent.EndTime) & ", " & vbCrLf
        
    strSQL = strSQL & IIf(pEvent.AllDayEvent, 1, 0) & ", " & vbCrLf

    strSQL = strSQL & IIf(pEvent.Reminder, 1, 0) & ", " & vbCrLf
    strSQL = strSQL & pEvent.ReminderMinutesBeforeStart & ", " & vbCrLf
    strSQL = strSQL & MakeSQLstr(pEvent.ReminderSoundFile) & ", " & vbCrLf
            
    strSQL = strSQL & pEvent.RecurrenceState & ", " & vbCrLf
    
    If pEvent.RecurrenceState = xtpCalendarRecurrenceMaster Or _
       pEvent.RecurrenceState = xtpCalendarRecurrenceException Then
        
        strSQL = strSQL & pEvent.RecurrencePattern.Id & ", " & vbCrLf
    Else
        strSQL = strSQL + "0, " & vbCrLf
    End If
    
    strSQL = strSQL & MakeSQLDateTime(pEvent.RExceptionStartTimeOrig) & ", " & vbCrLf
    strSQL = strSQL & MakeSQLDateTime(pEvent.RExceptionEndTimeOrig) & ", " & vbCrLf
    strSQL = strSQL & IIf(pEvent.RExceptionDeleted, 1, 0) & ", " & vbCrLf
    
    strSQL = strSQL & MakeSQLstr(pEvent.CustomProperties.SaveToString) & ", " & vbCrLf
    strSQL = strSQL & MakeSQLstr(pEvent.CustomIcons.SaveToString) & ", " & vbCrLf
    
    strSQL = strSQL & pEvent.ScheduleID & ");" & vbCrLf
       
    ''--------------------------------------------
    MakeCreateEventSQL = strSQL
End Function


Public Function MakeUpdatePatternSQL(pRPattern As CalendarRecurrencePattern) As String

    Dim strSQL As String
  
    strSQL = "UPDATE CalendarRecurrencePatterns SET " & vbCrLf
    
    strSQL = strSQL + "MasterEventID = " & pRPattern.MasterEventId & ", " & vbCrLf
    
    strSQL = strSQL + "EventStartTime = " & MakeSQLDateTime(pRPattern.StartTime) & ", " & vbCrLf
    strSQL = strSQL + "EventDuration = " & pRPattern.DurationMinutes & ", " & vbCrLf
    
    strSQL = strSQL + "PatternStartDate = " & MakeSQLDateTime(pRPattern.StartDate) & ", " & vbCrLf
    
    strSQL = strSQL + "PatternEndMethod = " & pRPattern.EndMethod & ", " & vbCrLf
    strSQL = strSQL + "PatternEndDate = " & MakeSQLDateTime(pRPattern.EndDate) & ", " & vbCrLf
    strSQL = strSQL + "PatternEndAfterOccurrences = " & pRPattern.EndAfterOccurrences & ", " & vbCrLf
    
    strSQL = strSQL + "OptionsData1 = " & pRPattern.Options.Data1 & ", " & vbCrLf
    strSQL = strSQL + "OptionsData2 = " & pRPattern.Options.Data2 & ", " & vbCrLf
    strSQL = strSQL + "OptionsData3 = " & pRPattern.Options.Data3 & ", " & vbCrLf
    strSQL = strSQL + "OptionsData4 = " & pRPattern.Options.Data4 & ", " & vbCrLf
    
    strSQL = strSQL + "CustomPropertiesXMLData = " & MakeSQLstr(pRPattern.CustomProperties.SaveToString) & " " & vbCrLf
    
    strSQL = strSQL + " WHERE RecurrencePatternID = " & pRPattern.Id
    
    MakeUpdatePatternSQL = strSQL
End Function

Public Function MakeCreatePatternSQL(pRPattern As CalendarRecurrencePattern) As String

    Dim strSQL As String
      
    strSQL = "INSERT INTO CalendarRecurrencePatterns ("
    
    strSQL = strSQL + "MasterEventID, " & vbCrLf
    
    strSQL = strSQL + "EventStartTime, " & vbCrLf
    strSQL = strSQL + "EventDuration, " & vbCrLf
    
    strSQL = strSQL + "PatternStartDate, " & vbCrLf
    
    strSQL = strSQL + "PatternEndMethod, " & vbCrLf
    strSQL = strSQL + "PatternEndDate, " & vbCrLf
    strSQL = strSQL + "PatternEndAfterOccurrences, " & vbCrLf
    
    strSQL = strSQL + "OptionsData1, " & vbCrLf
    strSQL = strSQL + "OptionsData2, " & vbCrLf
    strSQL = strSQL + "OptionsData3, " & vbCrLf
    strSQL = strSQL + "OptionsData4, " & vbCrLf
    
    strSQL = strSQL + "CustomPropertiesXMLData) " & vbCrLf
    
    '--------
    strSQL = strSQL + "VALUES("
    
    strSQL = strSQL & pRPattern.MasterEventId & ", " & vbCrLf
    
    strSQL = strSQL & MakeSQLDateTime(pRPattern.StartTime) & ", " & vbCrLf
    strSQL = strSQL & pRPattern.DurationMinutes & ", " & vbCrLf
    
    strSQL = strSQL & MakeSQLDateTime(pRPattern.StartDate) & ", " & vbCrLf
    
    strSQL = strSQL & pRPattern.EndMethod & ", " & vbCrLf
    strSQL = strSQL & MakeSQLDateTime(pRPattern.EndDate) & ", " & vbCrLf
    strSQL = strSQL & pRPattern.EndAfterOccurrences & ", " & vbCrLf
    
    strSQL = strSQL & pRPattern.Options.Data1 & ", " & vbCrLf
    strSQL = strSQL & pRPattern.Options.Data2 & ", " & vbCrLf
    strSQL = strSQL & pRPattern.Options.Data3 & ", " & vbCrLf
    strSQL = strSQL & pRPattern.Options.Data4 & ", " & vbCrLf
    
    strSQL = strSQL & MakeSQLstr(pRPattern.CustomProperties.SaveToString) & "); "
        
    MakeCreatePatternSQL = strSQL
End Function

Public Function MakeRetrieveDayEventsSQL(ByVal dtDay As Date) As String

    Dim nYear As Long, nMonth As Long, nDay As Long
    Dim strSQL As String
    
    nYear = Year(dtDay)
    nMonth = Month(dtDay)
    nDay = Day(dtDay)
    
    strSQL = "SELECT * FROM CalendarEvents WHERE " & vbCrLf
    
    strSQL = strSQL + " ( RecurrenceState = " & xtpCalendarRecurrenceNotRecurring & " OR " & vbCrLf
    strSQL = strSQL + "   RecurrenceState = " & xtpCalendarRecurrenceMaster & ") AND " & vbCrLf
        
    strSQL = strSQL + "( YEAR(StartDateTime) < " & nYear & vbCrLf
    strSQL = strSQL + "  OR ( YEAR(StartDateTime) = " & nYear & " AND " & vbCrLf
    strSQL = strSQL + "       ( MONTH(StartDateTime) < " & nMonth & " OR "
    strSQL = strSQL + "         MONTH(StartDateTime) = " & nMonth & " AND " & vbCrLf
    strSQL = strSQL + "         DAY(StartDateTime) <= " & nDay & vbCrLf
    
    strSQL = strSQL + "     ) ) )AND " & vbCrLf
    
    strSQL = strSQL + "( YEAR(EndDateTime) > " & nYear & vbCrLf
    strSQL = strSQL + "  OR ( YEAR(EndDateTime) = " & nYear & " AND " & vbCrLf
    strSQL = strSQL + "       (MONTH(EndDateTime) > " & nMonth & " OR "
    strSQL = strSQL + "        MONTH(EndDateTime) = " & nMonth & " AND " & vbCrLf
    strSQL = strSQL + "        DAY(EndDateTime) >= " & nDay & vbCrLf
    strSQL = strSQL + "     ) ) ) " & vbCrLf
    
    'Debug.Print strSQL

    MakeRetrieveDayEventsSQL = strSQL
End Function

Public Function MakeGetUpcomingEventsSQL(ByVal eDBtype As CodeJockCalendarDataType, ByVal dtFrom As Date, ByVal PeriodMinutes As Long, bOptimized As Boolean)
    
    ' The SQL script below is read all events from DB
    ' RemindersManager will select only reminder which will be Fiered
    ' untill dtFrom + PeriodMinutes. (generally dtFrom is Now)
    '
    ' Of cause you can write this script more optimized.
    '
    
    Dim strSQL As String
    
    strSQL = "SELECT * FROM CalendarEvents WHERE " & vbCrLf
    
    strSQL = strSQL + "(RecurrenceState = " & xtpCalendarRecurrenceNotRecurring & " OR " & vbCrLf
    strSQL = strSQL + " RecurrenceState = " & xtpCalendarRecurrenceMaster & ") " & vbCrLf
    
    '----------------------------------------------------------
    If bOptimized Then
        
        Dim dtUntil As Date, strUntil As String
        
        dtUntil = DateAdd("n", PeriodMinutes, dtFrom)
        strUntil = MakeSQLDateTime(dtUntil)
        
        Dim strReminderStartDateSQL As String
        If eDBtype = cjCalendarData_SQLServer Then
            strReminderStartDateSQL = " DATEADD(""n"", -1 * ReminderMinutesBeforeStart, StartDateTime) "
        ElseIf eDBtype = cjCalendarData_MySQL Then
            strReminderStartDateSQL = " ADDDATE(StartDateTime, INTERVAL -1*ReminderMinutesBeforeStart MINUTE) "
        Else
            MakeGetUpcomingEventsSQL = "err"
            Debug.Assert False
            Exit Function
        End If
        
        strSQL = strSQL + " AND " & vbCrLf
        strSQL = strSQL + " ( IsReminder <> 0 OR " & vbCrLf
        strSQL = strSQL + "     ( RecurrenceState = " & xtpCalendarRecurrenceMaster & " AND " & vbCrLf
        strSQL = strSQL + "       RecurrencePatternID IN " & vbCrLf
        strSQL = strSQL + "         ( SELECT RecurrencePatternID  FROM CalendarEvents " & vbCrLf
        strSQL = strSQL + "           WHERE " & vbCrLf
        strSQL = strSQL + "             RecurrenceState = " & xtpCalendarRecurrenceException & " AND " & vbCrLf
        strSQL = strSQL + "             IsReminder <> 0 AND " & vbCrLf
        strSQL = strSQL + "         " & strReminderStartDateSQL & " <= " & strUntil & vbCrLf
        strSQL = strSQL + "         ) " & vbCrLf
        strSQL = strSQL + "     ) " & vbCrLf
        strSQL = strSQL + " ) AND " & vbCrLf
        'strSQL = strSQL + " DATEADD(""n"", -1 * ReminderMinutesBeforeStart, StartDateTime) <= " & strUntil
        strSQL = strSQL & strReminderStartDateSQL & " <= " & strUntil
    End If
    
    MakeGetUpcomingEventsSQL = strSQL
    
End Function

Public Function GetDSN(strConnection As String)
    ' Remove "Provider=Custom;" string. Note: this is hardcoded value!
    GetDSN = Right(strConnection, Len(strConnection) - Len("Provider=Custom;"))
End Function

Private Sub Class_Initialize()

End Sub
