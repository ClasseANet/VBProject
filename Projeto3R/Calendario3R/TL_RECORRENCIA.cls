VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Recorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmRecorrencia
Attribute mvarMe.VB_VarHelpID = -1

Private m_bWasNotRecur As Boolean

Private mvarisDirt   As Boolean
Private mvarCalEvent As CalendarEvent 'mvarCalEvent
Private mvarRecEvent As CalendarRecurrencePattern 'mvarRecEvent
Private mvarRecState As CalendarEventRecurrenceState
Private mvarCalBegin As Date
Public Property Get RecState() As CalendarEventRecurrenceState
   RecState = mvarRecState
End Property
Public Property Let CalBegin(ByVal vData As Date)
   mvarCalBegin = vData
End Property
Public Property Get CalBegin() As Date
   CalBegin = mvarCalBegin
End Property
Public Property Set CalEvent(ByVal vData As CalendarEvent)
   Set mvarCalEvent = vData
End Property
Public Property Get CalEvent() As CalendarEvent
   Set CalEvent = mvarCalEvent
End Property
Public Property Let isDirt(ByVal vData As Boolean)
   mvarisDirt = vData
End Property
Public Property Get isDirt() As Boolean
   mvarisDirt = IIf(mvarisDirt, mvarisDirt, (mvarCalEvent.RecurrenceState <> mvarRecState))
   
   isDirt = mvarisDirt
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmRecorrencia
   Set mvarCalEvent = Nothing
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
   Set mvarCalEvent = Nothing
End Sub
Private Sub mvarMe_BtnCancelClick()
   mvarisDirt = False
   Unload mvarMe
End Sub
Private Sub mvarMe_BtnOKClick()
   Dim nRes
   
   Call UpdatePatternFromControls
   
   If mvarRecEvent.Exceptions.Count > 0 Then
       
       nRes = MsgBox("Any exceptions associated with this recurring appointment will be lost. Is this OK?", vbOKCancel)
       If nRes = vbCancel Then
           Exit Sub
       End If
       mvarRecEvent.RemoveAllExceptions
   End If
   mvarCalEvent.UpdateRecurrence mvarRecEvent
   ''=====================================================
   Unload mvarMe
End Sub
Private Sub mvarMe_BtnRemoveRecurrenceClick()
    mvarCalEvent.RemoveRecurrence
    Unload mvarMe
End Sub

Private Sub mvarMe_CmbEventDurationChange()
   Dim ArrHora
   Dim dHoraFim   As Date
   Dim nDuracao   As Long
   Dim sEscala    As String
   Dim i          As Integer
   
   ArrHora = Split(mvarMe.cmbEventStartTime, ":")
   mvarRecEvent.StartTime = TimeSerial(ArrHora(0), ArrHora(1), 0)
   
   nDuracao = Val(mvarMe.cmbEventDuration.Text)
   If Trim(mvarMe.cmbEventDuration.Text) <> "" Then
      For i = 1 To Len(Trim(mvarMe.cmbEventDuration.Text))
         sEscala = UCase(Trim(Mid(Trim(mvarMe.cmbEventDuration.Text), i, 1)))
         If Not IsNumeric(sEscala) And sEscala <> "" Then
            Exit For
         End If
      Next
      'ArrDur = Split(Trim(mvarMe.cmbEventDuration.Text), " ")
      'sEscala = UCase(Left(Trim(ArrDur(1)), 1))
      
      Select Case sEscala
         Case "H"
            mvarMe.cmbEventDuration.Text = nDuracao & " hora" & IIf(nDuracao = 1, "", "s")
            SendKeys "{End}"
            nDuracao = nDuracao * 60
            
         Case "D"
            mvarMe.cmbEventDuration.Text = nDuracao & " dia" & IIf(nDuracao = 1, "", "s")
            SendKeys "{End}"
            nDuracao = nDuracao * 60 * 24
            
         Case "S"
            mvarMe.cmbEventDuration.Text = nDuracao & " semana" & IIf(nDuracao = 1, "", "s")
            SendKeys "{End}"
            nDuracao = nDuracao * 60 * 24 * 7
      End Select
      
   End If
   dHoraFim = DateAdd("n", nDuracao, mvarRecEvent.StartTime)
   mvarMe.txtEventEndTime.Caption = FormatDateTime(dHoraFim, vbShortTime)

End Sub

Private Sub mvarMe_Load()
'    On Error GoTo skip1
   
   Dim dtStartTime   As Date
   Dim dtEndTime     As Date
   Dim nDuartionMinutes As Long
   Dim i             As Integer
   Dim sAux          As String
   
   m_bWasNotRecur = False
   mvarisDirt = False
    
   mvarRecState = mvarCalEvent.RecurrenceState

   If mvarRecState = xtpCalendarRecurrenceMaster Then
      Set mvarRecEvent = mvarCalEvent.RecurrencePattern
   Else
      Set mvarCalEvent = mvarCalEvent.CloneEvent

      If mvarRecState = xtpCalendarRecurrenceNotRecurring Then
         dtStartTime = mvarCalEvent.StartTime
         dtEndTime = mvarCalEvent.EndTime
         
         On Error Resume Next
         
         nDuartionMinutes = 0
         nDuartionMinutes = Abs(DateDiff("n", dtEndTime, dtStartTime))
         
         mvarCalEvent.CreateRecurrence
         m_bWasNotRecur = True
         
         Set mvarRecEvent = mvarCalEvent.RecurrencePattern
         With mvarRecEvent
            .StartTime = TimeValue(dtStartTime)
            .DurationMinutes = nDuartionMinutes
            .StartDate = DateValue(dtStartTime)
         End With
         mvarCalEvent.UpdateRecurrence mvarRecEvent
         
      Else
         'Debug.Assert nRState = xtpCalendarRecurrenceOccurrence Or nRState = xtpCalendarRecurrenceException
         
         mvarisDirt = True
         Set mvarRecEvent = mvarCalEvent.RecurrencePattern
         Set mvarCalEvent = mvarRecEvent.MasterEvent
      End If
   End If
      
skip1:
    '================================
   With mvarMe
      .pageDaily.BorderStyle = 0 ' 0 - none
      .pageWeekly.BorderStyle = 0 ' 0 - none
      .pageMonthly.BorderStyle = 0 ' 0 - none
      .pageYearly.BorderStyle = 0 ' 0 - none
        
      .frameRecurrencePatterm.Width = .frameApointmentTime.Width
      .pageWeekly.Move .pageDaily.Left, .pageDaily.Top
      .pageMonthly.Move .pageDaily.Left, .pageDaily.Top
      .pageYearly.Move .pageDaily.Left, .pageDaily.Top

      Call SetActivePage(1)
    
      .optPatternNoEnd.Value = True
        
      .btnRemoveRecurrence.Enabled = Not m_bWasNotRecur
  
      .cmbEventStartTime.Clear
      For i = 0 To 24 * 60 Step 30
         sAux = Format(Int(i / 60), "00") + ":" + Format(i Mod 60, "00")
         .cmbEventStartTime.AddItem sAux
      Next
    
      .cmbEventDuration.AddItem "05 minutos"
      .cmbEventDuration.AddItem "10 minutos"
      .cmbEventDuration.AddItem "15 minutos"
      .cmbEventDuration.AddItem "30 minutos"
      .cmbEventDuration.AddItem "01 hora"
      .cmbEventDuration.AddItem "02 horas"
      .cmbEventDuration.AddItem "03 horas"
      .cmbEventDuration.AddItem "04 horas"
      .cmbEventDuration.AddItem "01 dia"
      .cmbEventDuration.AddItem "02 dias"
      
    
      Call MonthlyInitialization
      Call YearlyInitialization
    
   
'      oprPatternNoEnd.OptionValue = xtpCalendarPatternEndNoDate
'      oprPatternEndAfter.OptionValue = xtpCalendarPatternEndAfterOccurrences
'      oprPatternEndByDate.OptionValue = xtpCalendarPatternEndDate

      UpdateControlsFromEvent
   End With
   Screen.MousePointer = vbDefault
End Sub
Private Sub MonthlyInitialization()
    On Error Resume Next
    
    If mvarMe.cmbMonthlyDayOfMonth.ListCount = 0 Then
        Dim cnt As Integer
        For cnt = 1 To 31
            mvarMe.cmbMonthlyDayOfMonth.AddItem cnt, cnt - 1
        Next cnt

        If mvarCalBegin <> "00:00:00" Then
            mvarMe.cmbMonthlyDayOfMonth.ListIndex = Day(mvarCalBegin) - 1
        End If
    End If

    If mvarMe.cmbMonthlyDay.ListCount = 0 Then
        mvarMe.cmbMonthlyDay.AddItem "1º", 0
        mvarMe.cmbMonthlyDay.AddItem "2º", 1
        mvarMe.cmbMonthlyDay.AddItem "3º", 2
        mvarMe.cmbMonthlyDay.AddItem "4º", 3
        mvarMe.cmbMonthlyDay.AddItem "Último", 4

        mvarMe.cmbMonthlyDay.ListIndex = 0
    End If

    If mvarMe.cmbMonthlyDay.ListCount = 0 Then
        mvarMe.cmbMonthlyDay.AddItem "Dia", 0
        mvarMe.cmbMonthlyDay.AddItem "Dia da semana", 1
        mvarMe.cmbMonthlyDay.AddItem "Dia do fim de semana", 2
        mvarMe.cmbMonthlyDay.AddItem "Domingo", 3
        mvarMe.cmbMonthlyDay.AddItem "Segunda", 4
        mvarMe.cmbMonthlyDay.AddItem "Terça", 5
        mvarMe.cmbMonthlyDay.AddItem "Quarta", 6
        mvarMe.cmbMonthlyDay.AddItem "Quinta", 7
        mvarMe.cmbMonthlyDay.AddItem "Sexta", 8
        mvarMe.cmbMonthlyDay.AddItem "Sábado", 9

        If mvarCalBegin <> "00:00:00" Then
            mvarMe.cmbMonthlyDay.ListIndex = Weekday(mvarCalBegin) + 2
        End If
    End If

    If mvarMe.txtMonthlyEveryMonth.Text = "" Then
        mvarMe.txtMonthlyEveryMonth.Text = "1"
    End If

    If mvarMe.txtMonthlyOfEveryTheMonths.Text = "" Then
        mvarMe.txtMonthlyOfEveryTheMonths.Text = "1"
    End If
End Sub

Private Sub YearlyInitialization()
    On Error Resume Next
    
    If mvarMe.cmbYearlyDate.ListCount = 0 Then
        Dim cnt As Integer
        For cnt = 1 To 31
            mvarMe.cmbYearlyDate.AddItem cnt, cnt - 1
        Next cnt

        If mvarCalBegin <> "00:00:00" Then
            mvarMe.cmbYearlyDate.ListIndex = Day(mvarCalBegin) - 1
        End If
    End If

    If mvarMe.cmdYearlyThePartOfWeek.ListCount = 0 Then
        mvarMe.cmdYearlyThePartOfWeek.AddItem "1º", 0
        mvarMe.cmdYearlyThePartOfWeek.AddItem "2º", 1
        mvarMe.cmdYearlyThePartOfWeek.AddItem "3º", 2
        mvarMe.cmdYearlyThePartOfWeek.AddItem "4º", 3
        mvarMe.cmdYearlyThePartOfWeek.AddItem "Último", 4

        mvarMe.cmdYearlyThePartOfWeek.ListIndex = 0
    End If

    If mvarMe.cmbYearlyTheDay.ListCount = 0 Then
        mvarMe.cmbYearlyTheDay.AddItem "Dia", 0
        mvarMe.cmbYearlyTheDay.AddItem "Dia da semana", 1
        mvarMe.cmbYearlyTheDay.AddItem "Dia do fim de semana", 2
        mvarMe.cmbYearlyTheDay.AddItem "Domingo", 3
        mvarMe.cmbYearlyTheDay.AddItem "Segunda", 4
        mvarMe.cmbYearlyTheDay.AddItem "Terça", 5
        mvarMe.cmbYearlyTheDay.AddItem "Quarta", 6
        mvarMe.cmbYearlyTheDay.AddItem "Quinta", 7
        mvarMe.cmbYearlyTheDay.AddItem "Sexta", 8
        mvarMe.cmbYearlyTheDay.AddItem "Sábado", 9
        
        If mvarCalBegin <> "00:00:00" Then
            mvarMe.cmbYearlyTheDay.ListIndex = Weekday(mvarCalBegin) + 2
        End If
    End If

    If mvarMe.cmbYearlyEveryDate.ListCount = 0 Then
        
        mvarMe.cmbYearlyEveryDate.AddItem "Janeiro", 0
        mvarMe.cmbYearlyEveryDate.AddItem "Fevereiro", 1
        mvarMe.cmbYearlyEveryDate.AddItem "Março", 2
        mvarMe.cmbYearlyEveryDate.AddItem "Abril", 3
        mvarMe.cmbYearlyEveryDate.AddItem "Maio", 4
        mvarMe.cmbYearlyEveryDate.AddItem "Junhoi", 5
        mvarMe.cmbYearlyEveryDate.AddItem "Julho", 6
        mvarMe.cmbYearlyEveryDate.AddItem "Agosto", 7
        mvarMe.cmbYearlyEveryDate.AddItem "Setembro", 8
        mvarMe.cmbYearlyEveryDate.AddItem "Outobro", 9
        mvarMe.cmbYearlyEveryDate.AddItem "Novembro", 10
        mvarMe.cmbYearlyEveryDate.AddItem "Dezembro", 11
        
        Dim k As Integer
        For k = 0 To mvarMe.cmbYearlyEveryDate.ListCount - 1
            mvarMe.cmbYearlyTheMonth.AddItem mvarMe.cmbYearlyEveryDate.List(k), k
        Next
        
        If mvarCalBegin <> "00:00:00" Then
            mvarMe.cmbYearlyEveryDate.ListIndex = Month(mvarCalBegin) - 1
            mvarMe.cmbYearlyTheMonth.ListIndex = Month(mvarCalBegin) - 1
        End If
    End If

End Sub

Private Sub SetActivePage(nPage As Long)
   With mvarMe
      .optRecDaily.Value = (nPage = 1)
      .optRecWeekly.Value = (nPage = 2)
      .optRecMonthly.Value = (nPage = 3)
      .optRecYearly.Value = (nPage = 4)
      
      .pageDaily.Visible = (nPage = 1)
      .pageWeekly.Visible = (nPage = 2)
      .pageMonthly.Visible = (nPage = 3)
      .pageYearly.Visible = (nPage = 4)
   End With
End Sub

Function WhichDayMask2index(ByVal nWDay As Long) As Long

    Select Case (nWDay And xtpCalendarDayAllWeek)
        
        Case xtpCalendarDayAllWeek:
            WhichDayMask2index = 0
        
        Case xtpCalendarDayMo_Fr:
            WhichDayMask2index = 1
        
        Case xtpCalendarDaySaSu:
            WhichDayMask2index = 2
        
        Case xtpCalendarDaySunday:
            WhichDayMask2index = 3
        
        Case xtpCalendarDayMonday:
            WhichDayMask2index = 4
        
        Case xtpCalendarDayTuesday:
            WhichDayMask2index = 5
        
        Case xtpCalendarDayWednesday:
            WhichDayMask2index = 6
        
        Case xtpCalendarDayThursday:
            WhichDayMask2index = 7
        
        Case xtpCalendarDayFriday:
            WhichDayMask2index = 8
        
        Case xtpCalendarDaySaturday:
            WhichDayMask2index = 9
    End Select
End Function

Function index2WhichDayMask(ByVal nIndex As Long) As Long

    Select Case (nIndex)
        Case 0:
            index2WhichDayMask = xtpCalendarDayAllWeek
        Case 1:
            index2WhichDayMask = xtpCalendarDayMo_Fr
        Case 2:
            index2WhichDayMask = xtpCalendarDaySaSu
        Case 3:
            index2WhichDayMask = xtpCalendarDaySunday
        Case 4:
            index2WhichDayMask = xtpCalendarDayMonday
        Case 5:
            index2WhichDayMask = xtpCalendarDayTuesday
        Case 6:
            index2WhichDayMask = xtpCalendarDayWednesday
        Case 7:
            index2WhichDayMask = xtpCalendarDayThursday
        Case 8:
            index2WhichDayMask = xtpCalendarDayFriday
        Case 9:
            index2WhichDayMask = xtpCalendarDaySaturday
    End Select
End Function

Private Sub UpdateControlsFromEvent()

    mvarMe.cmbEventStartTime.Text = FormatDateTime(mvarRecEvent.StartTime, vbShortTime)
    mvarMe.cmbEventDuration.Text = mvarRecEvent.DurationMinutes & " minutos"
    Call mvarMe_CmbEventDurationChange  ' Update EndTime control
    
    SetActivePage 1
    '
    If mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceDaily Then
        SetActivePage 1
        
        If mvarRecEvent.Options.DailyEveryWeekDayOnly Then
            mvarMe.optDailyEveryNdays = False
            mvarMe.optDailyEveryWorkDay = True
            mvarMe.txtDailyEveryNdays = 1
        Else
            mvarMe.optDailyEveryWorkDay = False
            mvarMe.optDailyEveryNdays = True
            mvarMe.txtDailyEveryNdays = mvarRecEvent.Options.DailyIntervalDays
        End If
    ElseIf mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceWeekly Then
        SetActivePage 2
        
        mvarMe.txtWeeklyNWeeks = mvarRecEvent.Options.WeeklyIntervalWeeks
        Dim nWDays As Long
        nWDays = mvarRecEvent.Options.WeeklyDayOfWeekMask
        
        mvarMe.chkWeeklyMonday = IIf((nWDays And xtpCalendarDayMonday) <> 0, 1, 0)
        mvarMe.chkWeeklyTusday = IIf((nWDays And xtpCalendarDayTuesday) <> 0, 1, 0)
        mvarMe.chkWeeklyWednesday = IIf((nWDays And xtpCalendarDayWednesday) <> 0, 1, 0)
        mvarMe.chkWeeklyThursday = IIf((nWDays And xtpCalendarDayThursday) <> 0, 1, 0)
        mvarMe.chkWeeklyFriday = IIf((nWDays And xtpCalendarDayFriday) <> 0, 1, 0)
        
        mvarMe.chkWeeklySaturday = IIf((nWDays And xtpCalendarDaySaturday) <> 0, 1, 0)
        mvarMe.chkWeeklySunday = IIf((nWDays And xtpCalendarDaySunday) <> 0, 1, 0)
                
    ElseIf mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceMonthly Or _
        mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceMonthNth _
    Then
        SetActivePage 3
        
        If mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceMonthly Then
            mvarMe.optMonthlyDay = True
            
            mvarMe.cmbMonthlyDayOfMonth.ListIndex = mvarRecEvent.Options.MonthlyDayOfMonth - 1
            mvarMe.txtMonthlyEveryMonth.Text = mvarRecEvent.Options.MonthlyIntervalMonths
        Else
            mvarMe.optMonthlyThe = True
            
            mvarMe.cmbMonthlyDay.ListIndex = mvarRecEvent.Options.MonthNthWhichDay - 1
            mvarMe.cmbMonthlyDay.ListIndex = WhichDayMask2index(mvarRecEvent.Options.MonthNthWhichDayMask)
            
            mvarMe.txtMonthlyOfEveryTheMonths.Text = mvarRecEvent.Options.MonthNthIntervalMonths
        End If
    
    ElseIf mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceYearly Or _
        mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceYearNth _
    Then
        SetActivePage 4
        
        If mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceYearly Then
            mvarMe.optYearlyDay = True
            
            mvarMe.cmbYearlyEveryDate.ListIndex = mvarRecEvent.Options.YearlyMonthOfYear - 1
            mvarMe.cmbYearlyDate.ListIndex = mvarRecEvent.Options.YearlyDayOfMonth - 1
       
        Else
            mvarMe.optYearlyThe = True
            
            mvarMe.cmdYearlyThePartOfWeek.ListIndex = mvarRecEvent.Options.YearNthWhichDay - 1
            mvarMe.cmbYearlyTheDay.ListIndex = WhichDayMask2index(mvarRecEvent.Options.YearNthWhichDayMask)
            mvarMe.cmbYearlyTheMonth.ListIndex = mvarRecEvent.Options.YearNthMonthOfYear - 1
        End If
    End If
    
    'Start-End pattern
    mvarMe.ddPatternStartDate = mvarRecEvent.StartDate
    
    mvarMe.optPatternNoEnd = mvarRecEvent.EndMethod = xtpCalendarPatternEndNoDate
    mvarMe.optPatternEndAfter = mvarRecEvent.EndMethod = xtpCalendarPatternEndAfterOccurrences
    mvarMe.optPatternEndByDate = mvarRecEvent.EndMethod = xtpCalendarPatternEndDate
        
    mvarMe.txtPatternEndAfter = "10"
    mvarMe.ddPatternEndDate = DateValue(Now() + 5)
    
    If mvarRecEvent.EndMethod = xtpCalendarPatternEndAfterOccurrences Then
        mvarMe.txtPatternEndAfter = mvarRecEvent.EndAfterOccurrences
    ElseIf mvarRecEvent.EndMethod = xtpCalendarPatternEndDate Then
        mvarMe.ddPatternEndDate = mvarRecEvent.EndDate
    End If
    
End Sub

Private Sub UpdatePatternFromControls()
    
    Dim arTimes
    arTimes = Split(mvarMe.cmbEventStartTime.Text, ":")
    mvarRecEvent.StartTime = TimeSerial(arTimes(0), arTimes(1), 0)
    mvarRecEvent.DurationMinutes = Val(mvarMe.cmbEventDuration.Text)

    If mvarMe.optRecDaily Then
        mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceDaily
                
        If mvarMe.optDailyEveryWorkDay Then
            mvarRecEvent.Options.DailyEveryWeekDayOnly = True
        Else
            mvarRecEvent.Options.DailyEveryWeekDayOnly = False
            mvarRecEvent.Options.DailyIntervalDays = Val(mvarMe.txtDailyEveryNdays)
        End If
        
    ElseIf mvarMe.optRecWeekly Then
        mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceWeekly

        mvarRecEvent.Options.WeeklyIntervalWeeks = Val(mvarMe.txtWeeklyNWeeks)
        Dim nWDays As Long
        nWDays = 0
        
        nWDays = nWDays + IIf(mvarMe.chkWeeklyMonday, xtpCalendarDayMonday, 0)
        nWDays = nWDays + IIf(mvarMe.chkWeeklyTusday, xtpCalendarDayTuesday, 0)
        nWDays = nWDays + IIf(mvarMe.chkWeeklyWednesday, xtpCalendarDayWednesday, 0)
        nWDays = nWDays + IIf(mvarMe.chkWeeklyThursday, xtpCalendarDayThursday, 0)
        nWDays = nWDays + IIf(mvarMe.chkWeeklyFriday, xtpCalendarDayFriday, 0)
        
        nWDays = nWDays + IIf(mvarMe.chkWeeklySaturday, xtpCalendarDaySaturday, 0)
        nWDays = nWDays + IIf(mvarMe.chkWeeklySunday, xtpCalendarDaySunday, 0)
                
        mvarRecEvent.Options.WeeklyDayOfWeekMask = nWDays
        
    ElseIf mvarMe.optRecMonthly Then
        
        If mvarMe.optMonthlyDay Then
            mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceMonthly
            
            mvarRecEvent.Options.MonthlyDayOfMonth = mvarMe.cmbMonthlyDayOfMonth.ListIndex + 1
            mvarRecEvent.Options.MonthlyIntervalMonths = CLng(Val(mvarMe.txtMonthlyEveryMonth.Text))
            If mvarRecEvent.Options.MonthlyIntervalMonths < 1 Then mvarRecEvent.Options.MonthlyIntervalMonths = 1
        
        Else
            mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceMonthNth
            
            mvarRecEvent.Options.MonthNthWhichDay = mvarMe.cmbMonthlyDay.ListIndex + 1
            mvarRecEvent.Options.MonthNthWhichDayMask = index2WhichDayMask(mvarMe.cmbMonthlyDay.ListIndex)
            
            mvarRecEvent.Options.MonthNthIntervalMonths = CLng(Val(mvarMe.txtMonthlyOfEveryTheMonths.Text))
            If mvarRecEvent.Options.MonthNthIntervalMonths < 1 Then mvarRecEvent.Options.MonthNthIntervalMonths = 1
        End If

    ElseIf mvarMe.optRecYearly Then
        '
        If mvarMe.optYearlyDay Then
            mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceYearly
            
            mvarRecEvent.Options.YearlyMonthOfYear = mvarMe.cmbYearlyEveryDate.ListIndex + 1
            mvarRecEvent.Options.YearlyDayOfMonth = mvarMe.cmbYearlyDate.ListIndex + 1
        Else
            mvarRecEvent.Options.RecurrenceType = xtpCalendarRecurrenceYearNth
            
            mvarRecEvent.Options.YearNthWhichDay = mvarMe.cmdYearlyThePartOfWeek.ListIndex + 1
            mvarRecEvent.Options.YearNthWhichDayMask = index2WhichDayMask(mvarMe.cmbYearlyTheDay.ListIndex)
            mvarRecEvent.Options.YearNthMonthOfYear = mvarMe.cmbYearlyTheMonth.ListIndex + 1
        End If
    End If
            
    'Start-End pattern
    mvarRecEvent.StartDate = DateValue(CDate(mvarMe.ddPatternStartDate))
    
    If mvarMe.optPatternNoEnd Then
        mvarRecEvent.EndMethod = xtpCalendarPatternEndNoDate
        
    ElseIf mvarMe.optPatternEndAfter Then
        mvarRecEvent.EndMethod = xtpCalendarPatternEndAfterOccurrences
        mvarRecEvent.EndAfterOccurrences = Val(mvarMe.txtPatternEndAfter)
        
    ElseIf mvarMe.optPatternEndByDate Then
        mvarRecEvent.EndMethod = xtpCalendarPatternEndDate
        mvarRecEvent.EndDate = DateValue(CDate(mvarMe.ddPatternEndDate))
    End If
    
End Sub
Private Sub mvarMe_OptRecDailyClick()
   SetActivePage 1
End Sub
Private Sub mvarMe_OptRecMonthlyClick()
   SetActivePage 3
End Sub
Private Sub mvarMe_OptRecWeeklyClick()
   SetActivePage 2
End Sub
Private Sub mvarMe_OptRecYearlyClick()
   SetActivePage 4
End Sub
