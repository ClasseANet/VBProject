VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "providerSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents m_pCalendar As CalendarControl
Attribute m_pCalendar.VB_VarHelpID = -1

Dim m_Connection As New ADODB.Connection
Dim m_bOpened As Boolean
Dim m_pHelper As New SQLDataHelper

'======================================================================
Public Sub SetCalendar(ByVal pCalendar As CalendarControl)
    Set m_pCalendar = pCalendar
    
    m_pHelper.SetCalendar pCalendar
End Sub

Public Sub OpenDB(ByVal strConnection As String)
On Error GoTo err1
    ' Open DSN connection
    m_Connection.Open m_pHelper.GetDSN(strConnection)
    m_bOpened = True

err1:
    If Err.Number <> 0 Then
        m_bOpened = False
        MsgBox "Cannot open SQL Server DB: " & Err.Description
    End If
End Sub

Public Sub CloseDB()
    If m_bOpened Then
        m_Connection.Close
        m_bOpened = False
    End If
End Sub


Private Sub Class_Initialize()
    m_bOpened = False
End Sub

Private Sub Class_Terminate()
    CloseDB
End Sub


Private Sub m_pCalendar_DoCreateEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, NewEventID As Long, bResult As Boolean)
    
    bResult = False
    
    If Not m_bOpened Then
        Exit Sub
    End If

On Error GoTo err1

    Dim strSQL As String
    strSQL = "SELECT * FROM CalendarEvents WHERE EventID = 0"
        
    Dim rsEvent As New ADODB.Recordset
    rsEvent.Open strSQL, m_Connection, adOpenStatic, adLockOptimistic
        
    m_Connection.BeginTrans
    '-----------------------------------------------------
    Debug.Assert rsEvent.BOF And rsEvent.EOF
    
    rsEvent.AddNew
        
    m_pHelper.PutEventToRS pEvent, rsEvent

    rsEvent.Update
    ''--------------------------------------------
    rsEvent.Close
    m_Connection.CommitTrans
    
    '-- update Auto EventID
    Dim rsNewID As ADODB.Recordset
    strSQL = "SELECT MAX(EventID) AS NewEventID FROM CalendarEvents"
    
    Set rsNewID = m_Connection.Execute(strSQL)
    NewEventID = rsNewID("NewEventID")
            
    rsNewID.Close
    
    bResult = True
    Exit Sub

err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot CreateEvent in DB: " & Err.Description
        Debug.Assert False
    End If
End Sub

Private Sub m_pCalendar_DoCreateRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, NewPatternID As Long, bResult As Boolean)
    bResult = False
    
    If Not m_bOpened Then
        Exit Sub
    End If

On Error GoTo err1

    Dim strSQL As String
    strSQL = "SELECT * FROM CalendarRecurrencePatterns "
    strSQL = strSQL + " WHERE RecurrencePatternID = 0 "
        
    Dim rsPattern As New ADODB.Recordset
    rsPattern.Open strSQL, m_Connection, adOpenStatic, adLockOptimistic
        
    Dim nTransLevel As Long
    nTransLevel = m_Connection.BeginTrans
    
    '-----------------------------------------------------
    Debug.Assert rsPattern.BOF And rsPattern.EOF
    
    rsPattern.AddNew
    
    m_pHelper.PutRPatternToRS pPattern, rsPattern

    rsPattern.Update
    ''--------------------------------------------
    
    rsPattern.Close
    m_Connection.CommitTrans

    '-- update Auto EventID
    Dim rsNewID As ADODB.Recordset
    strSQL = "SELECT MAX(RecurrencePatternID) AS NewNewPatternID FROM CalendarRecurrencePatterns"
    
    Set rsNewID = m_Connection.Execute(strSQL)
    NewPatternID = rsNewID("NewNewPatternID")
    
    rsNewID.Close
    
    bResult = True
    Exit Sub
err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot UpdateEvent in DB: " & Err.Description
        Debug.Assert False
    End If

End Sub

Private Sub m_pCalendar_DoDeleteEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, bResult As Boolean)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "DELETE FROM CalendarEvents WHERE EventID = " & pEvent.Id
    
    m_Connection.Execute strSQL, , adExecuteNoRecords
    
    bResult = True
End Sub

Private Sub m_pCalendar_DoDeleteRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, bResult As Boolean)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "DELETE FROM CalendarRecurrencePatterns WHERE RecurrencePatternID = " & pPattern.Id
    
    m_Connection.Execute strSQL, , adExecuteNoRecords
    
    bResult = True
End Sub

Private Sub m_pCalendar_DoGetUpcomingEvents(ByVal dtFrom As Date, ByVal PeriodMinutes As Long, ByVal pEvents As XtremeCalendarControl.CalendarEvents)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    '// The SQL script below can be created in 2 ways - simple or optimized.
    '//
    '// * The simple script is read all events from DB.
    '//   RemindersManager will filter events and use only events with reminder
    '//   which will be Fired untill dtFrom + PeriodMinutes.
    '//   This can be very slow operation,
    '//   but it is called no so often - once per 2 hours by default.
    '//
    '// * The optimized script is filter events instead of reminder manager.
    '//   This can significantly improve performance.
    '//
    '// The simple variant can be useful for some debug purposes.
    '//
    '// NOTE: generally dtFrom is Now.
    
    Dim bOptimizeSQL As Boolean
    bOptimizeSQL = True
    
    Dim strSQL As String
    strSQL = m_pHelper.MakeGetUpcomingEventsSQL(cjCalendarData_SQLServer, dtFrom, PeriodMinutes, bOptimizeSQL)
    'Debug.Print strSQL
    '----------------------------------------------------------
    
    Dim rsEvents As ADODB.Recordset
    Set rsEvents = m_Connection.Execute(strSQL)
    
    Dim pEvent As CalendarEvent
        
    While Not rsEvents.EOF
        Set pEvent = m_pHelper.CreateEventFromRS(rsEvents, False)
        If Not pEvent Is Nothing Then
            pEvents.Add pEvent
        End If
        
        rsEvents.MoveNext
    Wend

End Sub

Private Sub m_pCalendar_DoReadEvent(ByVal EventID As Long, pEvent As XtremeCalendarControl.CalendarEvent)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "SELECT * FROM CalendarEvents WHERE EventID = " & EventID
        
    Dim rsEvent As ADODB.Recordset
    Set rsEvent = m_Connection.Execute(strSQL)
    
    Set pEvent = m_pHelper.CreateEventFromRS(rsEvent, False)
End Sub

Private Sub m_pCalendar_DoReadRPattern(ByVal PatternID As Long, pPattern As XtremeCalendarControl.CalendarRecurrencePattern)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "SELECT * FROM CalendarRecurrencePatterns "
    strSQL = strSQL + " WHERE RecurrencePatternID = " & PatternID
        
    Dim rsRPattern As ADODB.Recordset
    Set rsRPattern = m_Connection.Execute(strSQL)
    
    Set pPattern = m_pHelper.CreateRPatternFromRS(rsRPattern)
End Sub

Private Sub m_pCalendar_DoRetrieveDayEvents(ByVal dtDay As Date, ByVal pEvents As XtremeCalendarControl.CalendarEvents)

    If Not m_bOpened Then
        Exit Sub
    End If

    Dim strSQL As String
    strSQL = m_pHelper.MakeRetrieveDayEventsSQL(dtDay)

    Dim rsEvents As ADODB.Recordset
    Set rsEvents = m_Connection.Execute(strSQL)
    
    Dim pEvent As CalendarEvent
        
    While Not rsEvents.EOF
        Set pEvent = m_pHelper.CreateEventFromRS(rsEvents, False)
        If Not pEvent Is Nothing Then
            pEvents.Add pEvent
        End If
        
        rsEvents.MoveNext
    Wend
       
End Sub

Private Sub m_pCalendar_DoUpdateEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, bResult As Boolean)
    
    bResult = False
    
    If Not m_bOpened Then
        Exit Sub
    End If

On Error GoTo err1

    Dim strSQL As String
    strSQL = "SELECT * FROM CalendarEvents WHERE EventID = " & pEvent.Id
        
    Dim rsEvent As New ADODB.Recordset
    rsEvent.Open strSQL, m_Connection, adOpenStatic, adLockOptimistic
    
    m_Connection.BeginTrans
    
    '-----------------------------------------------------
    m_pHelper.PutEventToRS pEvent, rsEvent

    rsEvent.Update
   
    ''--------------------------------------------
    rsEvent.Close
    
    m_Connection.CommitTrans
    
    bResult = True
    Exit Sub
err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot UpdateEvent in DB: " & Err.Description
        Debug.Assert False
    End If
    
End Sub

Private Sub m_pCalendar_DoUpdateRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, bResult As Boolean)
    
    bResult = False
    
    If Not m_bOpened Then
        Exit Sub
    End If

On Error GoTo err1

    Dim strSQL As String
    strSQL = "SELECT * FROM CalendarRecurrencePatterns "
    strSQL = strSQL + " WHERE RecurrencePatternID = " & pPattern.Id
        
    Dim rsPattern As New ADODB.Recordset
    rsPattern.Open strSQL, m_Connection, adOpenStatic, adLockOptimistic
        
    Debug.Assert Not rsPattern.BOF And Not rsPattern.EOF
    
    Dim nTransLevel As Long
    nTransLevel = m_Connection.BeginTrans
    
    '-----------------------------------------------------
    m_pHelper.PutRPatternToRS pPattern, rsPattern

    rsPattern.Update
    ''--------------------------------------------
    rsPattern.Close
    
    m_Connection.CommitTrans
    
    bResult = True
    Exit Sub
err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot UpdateEvent in DB: " & Err.Description
        Debug.Assert False
    End If
    
End Sub
