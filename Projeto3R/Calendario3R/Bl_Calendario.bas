Attribute VB_Name = "Bl_Calendario"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CB_SETDROPPEDWIDTH = &H160

Public Enum CodeJockCalendarDataType
    cjCalendarData_Unknown = 0
    cjCalendarData_Memory = 1
    cjCalendarData_Access = 2
    cjCalendarData_MAPI = 3
    cjCalendarData_SQLServer = 4
    cjCalendarData_MySQL = 5
End Enum

Public ModalFormsRunningCounter              As Long
Public DisableDragging_ForRecurrenceEvents   As Boolean
Public DisableInPlaceCreateEvents_ForSaSu    As Boolean
Public EnableScrollV_DayView                 As Boolean
Public EnableScrollH_DayView                 As Boolean
Public EnableScrollV_WeekView                As Boolean
Public EnableScrollV_MonthView               As Boolean
Public g_DataResourcesMan                    As CalendarResourcesManager
Public g_bUseBuiltInCalendarDialogs          As Boolean
Public g_dlgCalendarReminders                As New CalendarDialogs
Public Sub CarregaComboTempo(CmbTempo As ComboBox, bSnoozeBox As Boolean)
   With CmbTempo
      .Clear
      If Not bSnoozeBox Then
         .AddItem "0 minuto"
         .AddItem "1 minuto"
      End If
      
      .AddItem "5 minutos"
      .AddItem "10 minutos"
      .AddItem "15 minutos"
      .AddItem "30 minutos"
      
      .AddItem "1 hora"
      .AddItem "2 horas"
      .AddItem "4 horas"
      .AddItem "8 horas"
      
      .AddItem "0.5 dia"
      .AddItem "1 dia"
      .AddItem "2 dias"
      .AddItem "3 dias"
      .AddItem "4 dias"
      
      .AddItem "1 semana"
      .AddItem "2 semanas"
   End With
End Sub
Public Function CalcStandardDurations_0m_2wString(lDuration As Long) As String
    Select Case lDuration
        Case 0:
            CalcStandardDurations_0m_2wString = "0 minutes"
        Case 1:
            CalcStandardDurations_0m_2wString = "1 minute"
        Case 5:
            CalcStandardDurations_0m_2wString = "5 minutes"
        Case 10:
            CalcStandardDurations_0m_2wString = "10 minutes"
        Case 15:
            CalcStandardDurations_0m_2wString = "15 minutes"
        Case 30:
            CalcStandardDurations_0m_2wString = "30 minutes"
        
        Case 60:
            CalcStandardDurations_0m_2wString = "1 hour"
        Case (60 * 2):
            CalcStandardDurations_0m_2wString = "2 hours"
        Case (60 * 4):
            CalcStandardDurations_0m_2wString = "4 hours"
        Case (60 * 8):
            CalcStandardDurations_0m_2wString = "8 hours"
        
        Case (60 * 12):
            CalcStandardDurations_0m_2wString = "0.5 day"
        Case (60 * 24):
            CalcStandardDurations_0m_2wString = "1 day"
        Case (60 * 24 * 2):
            CalcStandardDurations_0m_2wString = "2 days"
        Case (60 * 24 * 3):
            CalcStandardDurations_0m_2wString = "3 days"
        Case (60 * 24 * 4):
            CalcStandardDurations_0m_2wString = "4 days"
        
        Case (60 * 24 * 7):
            CalcStandardDurations_0m_2wString = "1 week"
        Case (60 * 24 * 7 * 2):
            CalcStandardDurations_0m_2wString = "2 weeks"
    End Select
End Function
Public Function CalcStandardDurations_0m_2wLong(sDuration As String) As Long
    Select Case sDuration
        Case "0 minutes":
            CalcStandardDurations_0m_2wLong = 0
        Case "1 minute":
            CalcStandardDurations_0m_2wLong = 1
        Case "5 minutes":
            CalcStandardDurations_0m_2wLong = 5
        Case "10 minutes":
            CalcStandardDurations_0m_2wLong = 10
        Case "15 minutes":
            CalcStandardDurations_0m_2wLong = 15
        Case "30 minutes":
            CalcStandardDurations_0m_2wLong = 30
        
        Case "1 hour":
            CalcStandardDurations_0m_2wLong = 60
        Case "2 hours":
            CalcStandardDurations_0m_2wLong = 60 * 2
        Case "4 hours":
            CalcStandardDurations_0m_2wLong = 60 * 4
        Case "8 hours":
            CalcStandardDurations_0m_2wLong = 60 * 8
        
        Case "0.5 day":
            CalcStandardDurations_0m_2wLong = 60 * 12
        Case "1 day":
            CalcStandardDurations_0m_2wLong = 60 * 24
        Case "2 days":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 2
        Case "3 days":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 3
        Case "4 days":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 4
        
        Case "1 week":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 7
        Case "2 weeks":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 7 * 2
    End Select
End Function
Public Function DateFromString(DatePart As String, TimePart As String) As Date
    Dim dtDatePart As Date
    Dim dtTimePart As Date
    
   If TimePart = "" Then
      TimePart = "00:00:00"
   End If
   
   dtDatePart = Mid(DatePart, 1, 10)
   dtTimePart = Mid(TimePart, 1, 8)
    
   DateFromString = dtDatePart + dtTimePart
End Function
Public Function DefineShowAs(pIDTPSERVICO As Integer) As Integer
   Select Case pIDTPSERVICO
      Case 1: DefineShowAs = 0 '* Teste.
      Case 2: DefineShowAs = 2 '* Tratamento.
      Case 3: DefineShowAs = 1 '* Manutenção.
   End Select
End Function
Public Function DefineLabelID(pIDTPTRATAMENTO As Integer, pIDTPSERVICO As Integer, Optional pFLGCANCELADO As Integer) As Long
   If pFLGCANCELADO = 0 Then
      If pIDTPSERVICO = 0 Then      'Compromisso Simples
         DefineLabelID = 0
      ElseIf pIDTPSERVICO = 1 Then  'Avaliação
         DefineLabelID = 1001
      Else
         DefineLabelID = pIDTPTRATAMENTO
      End If
   Else
      DefineLabelID = 9999          'Cancelado
   End If
End Function
Public Function FormatTimeDuration(ByVal nMinutes As Long, ByVal bAprox As Boolean) As String
    Dim nWeeks As Long, nDays As Long, nHours As Long
    
    nWeeks = nMinutes / (7 * 24 * 60)
    nDays = nMinutes / (24 * 60)
    nHours = nMinutes / 60

    Dim strDuration As String
    
    If (bAprox Or (nMinutes Mod (7 * 24 * 60)) = 0) And nWeeks > 0 Then
        strDuration = nWeeks & " week" & IIf(nWeeks > 1, "s", "")
    
    ElseIf (bAprox Or (nMinutes Mod (24 * 60)) = 0) And nDays > 0 Then
        strDuration = nDays & " day" & IIf(nDays > 1, "s", "")
        
    ElseIf (bAprox Or (nMinutes Mod 60) = 0) And nHours > 0 Then
        strDuration = nHours & " hour" & IIf(nHours > 1, "s", "")
        
    Else
        strDuration = nMinutes & " minute" & IIf(nMinutes > 1, "s", "")
    End If

    FormatTimeDuration = strDuration
End Function
Public Function ParseTimeDuration(ByVal strTime As String, ByRef pnMinutes As Long) As Boolean
    pnMinutes = 0
    ParseTimeDuration = False
        
    Dim nI As Long, nLen As Long
    Dim nMeasureStart As Long, nFIdx As Long
    Dim strChI As String
        
    strTime = Trim(strTime)
    nLen = Len(strTime)
    
    If nLen = 0 Then
        Exit Function
    End If
    
    '------------------------------------------
    nMeasureStart = -1
    For nI = 1 To nLen
        strChI = Mid(strTime, nI, 1)
        nFIdx = InStr(1, "-+.,0123456789", strChI)
        If nFIdx <= 0 Then
            nMeasureStart = nI
            Exit For
        End If
    Next
    
    Dim strNumber As String, strMeasure As String
    Dim nMultiplier As Long
            
    If nMeasureStart > 0 Then
        strNumber = Left(strTime, nMeasureStart - 1)
        strMeasure = Mid(strTime, nMeasureStart)
        strMeasure = Trim(strMeasure)
    Else
        strNumber = strTime
    End If
    
    If Len(strNumber) = 0 Then
        Exit Function
    End If

    Dim strM0 As String
    strM0 = Left(strMeasure, 1)
    
    nMultiplier = 1
    If strM0 = "m" Or strM0 = "M" Then
        nMultiplier = 1
    ElseIf strM0 = "h" Or strM0 = "H" Then
        nMultiplier = 60
    ElseIf strM0 = "d" Or strM0 = "D" Then
        nMultiplier = 60 * 24
    ElseIf strM0 = "w" Or strM0 = "W" Then
        nMultiplier = 60 * 24 * 7
    End If

    Dim dblTime As Double
    dblTime = Val(strNumber)
    
    pnMinutes = dblTime * nMultiplier
    ParseTimeDuration = True
End Function
