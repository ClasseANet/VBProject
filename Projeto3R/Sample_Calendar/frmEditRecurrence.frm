VERSION 5.00
Begin VB.Form frmEditRecurrence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Recurrence"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameRangeOfRecurrence 
      Caption         =   "Range of recurrence"
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   5520
      Width           =   7575
      Begin VB.ComboBox ddPatternEndDate 
         Height          =   315
         Left            =   4440
         TabIndex        =   39
         Top             =   1080
         Width           =   1635
      End
      Begin VB.TextBox txtPatternEndAfter 
         Height          =   315
         Left            =   4440
         TabIndex        =   37
         Text            =   "10"
         Top             =   660
         Width           =   675
      End
      Begin VB.OptionButton optPatternEndByDate 
         Caption         =   "End by:"
         Height          =   195
         Left            =   3360
         TabIndex        =   36
         Top             =   1140
         Width           =   1035
      End
      Begin VB.OptionButton optPatternEndAfter 
         Caption         =   "End after:"
         Height          =   195
         Left            =   3360
         TabIndex        =   35
         Top             =   720
         Width           =   1035
      End
      Begin VB.OptionButton optPatternNoEnd 
         Caption         =   "No end date"
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   300
         Width           =   1755
      End
      Begin VB.ComboBox ddPatternStartDate 
         Height          =   315
         Left            =   720
         TabIndex        =   33
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "ocurrences"
         Height          =   195
         Left            =   5160
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Start:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   30
      Top             =   7440
      Width           =   1275
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      Top             =   7440
      Width           =   1275
   End
   Begin VB.CommandButton btnRemoveRecurrence 
      Caption         =   "Remove Recurrence"
      Height          =   375
      Left            =   5760
      TabIndex        =   28
      Top             =   7440
      Width           =   1875
   End
   Begin VB.Frame frameRecurrencePatterm 
      Caption         =   "Recurrence Pattern"
      Height          =   4035
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   7575
      Begin VB.Frame pageYearly 
         Caption         =   "Yearly"
         Height          =   1095
         Left            =   1200
         TabIndex        =   52
         Top             =   2760
         Width           =   6015
         Begin VB.ComboBox cmbYearlyTheMonth 
            Height          =   315
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   650
            Width           =   1815
         End
         Begin VB.ComboBox cmbYearlyDate 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optYearlyThe 
            Caption         =   "The"
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   660
            Width           =   675
         End
         Begin VB.OptionButton optYearlyDay 
            Caption         =   "Day"
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.ComboBox cmbYearlyEveryDate 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox cmbYearlyTheDay 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   650
            Width           =   1815
         End
         Begin VB.ComboBox cmdYearlyThePartOfWeek 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   650
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "of"
            Height          =   255
            Left            =   3840
            TabIndex        =   58
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame pageMonthly 
         Caption         =   "Monthly"
         Height          =   1095
         Left            =   1200
         TabIndex        =   40
         Top             =   1560
         Width           =   6015
         Begin VB.TextBox txtMonthlyOfEveryTheMonths 
            Height          =   315
            Left            =   4200
            TabIndex        =   49
            Text            =   "1"
            Top             =   650
            Width           =   615
         End
         Begin VB.ComboBox cmbMonthlyDay 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   650
            Width           =   1095
         End
         Begin VB.ComboBox cmbMonthlyDayOfWeek 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   650
            Width           =   1455
         End
         Begin VB.ComboBox cmbMonthlyDayOfMonth 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optMonthlyDay 
            Caption         =   "Day"
            Height          =   315
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.TextBox txtMonthlyEveryMonth 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   42
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optMonthlyThe 
            Caption         =   "The"
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   660
            Width           =   675
         End
         Begin VB.Label Label11 
            Caption         =   "of Every"
            Height          =   255
            Left            =   3480
            TabIndex        =   51
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Month(s)"
            Height          =   195
            Left            =   4920
            TabIndex        =   50
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label9 
            Caption         =   "of Every"
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Month(s)"
            Height          =   195
            Left            =   3120
            TabIndex        =   44
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame pageWeekly 
         Caption         =   "Weekly"
         Height          =   1395
         Left            =   3480
         TabIndex        =   13
         Top             =   120
         Width           =   4935
         Begin VB.CheckBox chkWeeklySunday 
            Caption         =   "Sunday"
            Height          =   255
            Left            =   2400
            TabIndex        =   27
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklySaturday 
            Caption         =   "Saturday"
            Height          =   255
            Left            =   1260
            TabIndex        =   26
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklyFriday 
            Caption         =   "Friday"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklyThursday 
            Caption         =   "Thursday"
            Height          =   255
            Left            =   3720
            TabIndex        =   24
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklyWednesday 
            Caption         =   "Wednesday"
            Height          =   255
            Left            =   2400
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkWeeklyTusday 
            Caption         =   "Tusday"
            Height          =   255
            Left            =   1260
            TabIndex        =   22
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklyMonday 
            Caption         =   "Monday"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.TextBox txtWeeklyNWeeks 
            Height          =   315
            Left            =   1140
            TabIndex        =   19
            Text            =   "1"
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label6 
            Caption         =   "week(s) on:"
            Height          =   255
            Left            =   1980
            TabIndex        =   20
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Recurr 
            Caption         =   "Recur Every"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame pageDaily 
         Caption         =   "Daily"
         Height          =   1275
         Left            =   1200
         TabIndex        =   12
         Top             =   180
         Width           =   2235
         Begin VB.OptionButton optDailyEveryWorkDay 
            Caption         =   "Every Work Day"
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1635
         End
         Begin VB.TextBox txtDailyEveryNdays 
            Height          =   285
            Left            =   900
            TabIndex        =   15
            Text            =   "1"
            Top             =   180
            Width           =   615
         End
         Begin VB.OptionButton optDailyEveryNdays 
            Caption         =   "Every"
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   180
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.Label Label5 
            Caption         =   "day(s)"
            Height          =   195
            Left            =   1560
            TabIndex        =   16
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.OptionButton optRecYearly 
         Caption         =   "Yearly"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Width           =   855
      End
      Begin VB.OptionButton optRecMonthly 
         Caption         =   "Monthly"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   915
      End
      Begin VB.OptionButton optRecWeekly 
         Caption         =   "Weekly"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   915
      End
      Begin VB.OptionButton optRecDaily 
         Caption         =   "Daily"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Line lnSeparator1 
         BorderColor     =   &H80000003&
         X1              =   1140
         X2              =   1140
         Y1              =   300
         Y2              =   1860
      End
   End
   Begin VB.Frame frameApointmentTime 
      Caption         =   "Apointment Time"
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.ComboBox cmbEventDuration 
         Height          =   315
         Left            =   5160
         TabIndex        =   6
         Top             =   300
         Width           =   2235
      End
      Begin VB.ComboBox cmbEventStartTime 
         Height          =   315
         Left            =   780
         TabIndex        =   2
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Duration:"
         Height          =   255
         Left            =   4260
         TabIndex        =   5
         Top             =   360
         Width           =   795
      End
      Begin VB.Label txtEventEndTime 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   315
         Left            =   3000
         TabIndex        =   4
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "End:"
         Height          =   255
         Left            =   2100
         TabIndex        =   3
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmEditRecurrence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_pMasterEvent As CalendarEvent
Public m_bUpdateFromEvent As Boolean

Private m_pRPattern As CalendarRecurrencePattern
Private m_bWasNotRecur As Boolean

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()

    UpdatePatternFromControls
    
    If m_pRPattern.Exceptions.Count > 0 Then
        Dim nRes
        nRes = MsgBox("Any exceptions associated with this recurring appointment will be lost. Is this OK?", vbOKCancel)
        If nRes = vbCancel Then
            Exit Sub
        End If
        
        m_pRPattern.RemoveAllExceptions
    End If
    
    m_pMasterEvent.UpdateRecurrence m_pRPattern

    ''=====================================================
    Unload Me
End Sub

Private Sub btnRemoveRecurrence_Click()

    m_pMasterEvent.RemoveRecurrence
    
    Unload Me
End Sub

Private Sub cmbEventDuration_Change()
    Dim arTimes
    arTimes = Split(cmbEventStartTime, ":")
    m_pRPattern.StartTime = TimeSerial(arTimes(0), arTimes(1), 0)
    
    Dim dtEndTime As Date
    dtEndTime = DateAdd("n", Val(cmbEventDuration), m_pRPattern.StartTime)
                
    txtEventEndTime.Caption = FormatDateTime(dtEndTime, vbShortTime)
End Sub

Private Sub cmbEventDuration_Click()
    cmbEventDuration_Change
End Sub

Private Sub cmbEventStartTime_Change()
'    UpdateEndTimeCombo
End Sub

Private Sub cmbEventStartTime_Click()
    cmbEventDuration_Change
End Sub

Private Sub cmbEventStartTime_LostFocus()
    cmbEventDuration_Change
End Sub

Private Sub Form_Load()
    
'    On Error GoTo skip1

    m_bWasNotRecur = False
    m_bUpdateFromEvent = False
    
    Dim nRState As CalendarEventRecurrenceState
    nRState = m_pMasterEvent.RecurrenceState

    If nRState = xtpCalendarRecurrenceMaster Then
        
        Set m_pRPattern = m_pMasterEvent.RecurrencePattern
        
    Else
        Set m_pMasterEvent = m_pMasterEvent.CloneEvent

        If nRState = xtpCalendarRecurrenceNotRecurring Then
        
            Dim dtStartTime As Date, dtEndTime As Date
            Dim nDuartionMinutes As Long
            
            dtStartTime = m_pMasterEvent.StartTime
            dtEndTime = m_pMasterEvent.EndTime
            
            On Error Resume Next
            
            nDuartionMinutes = 0
            nDuartionMinutes = Abs(DateDiff("n", dtEndTime, dtStartTime))

            m_pMasterEvent.CreateRecurrence
            m_bWasNotRecur = True
            
            Set m_pRPattern = m_pMasterEvent.RecurrencePattern
                      
            m_pRPattern.StartTime = TimeValue(dtStartTime)
            m_pRPattern.DurationMinutes = nDuartionMinutes

            m_pRPattern.StartDate = DateValue(dtStartTime)

            m_pMasterEvent.UpdateRecurrence m_pRPattern
        Else
            Debug.Assert nRState = xtpCalendarRecurrenceOccurrence Or nRState = xtpCalendarRecurrenceException
        
            m_bUpdateFromEvent = True

            Set m_pRPattern = m_pMasterEvent.RecurrencePattern
            Set m_pMasterEvent = m_pRPattern.MasterEvent
        End If
    End If
      
skip1:
    '================================
    pageDaily.BorderStyle = 0 ' 0 - none
    pageWeekly.BorderStyle = 0 ' 0 - none
    pageMonthly.BorderStyle = 0 ' 0 - none
    pageYearly.BorderStyle = 0 ' 0 - none
        
    frameRecurrencePatterm.Width = frameApointmentTime.Width
    pageWeekly.Move pageDaily.Left, pageDaily.Top
    pageMonthly.Move pageDaily.Left, pageDaily.Top
    pageYearly.Move pageDaily.Left, pageDaily.Top

    SetActivePage 1
    
    optPatternNoEnd.Value = True
        
    btnRemoveRecurrence.Enabled = Not m_bWasNotRecur
  
    Dim i As Long, strTmp As String
    For i = 0 To 24 * 60 Step 30
        strTmp = Format(i / 60, "00") + ":" + Format(i Mod 60, "00")
        cmbEventStartTime.AddItem strTmp
    Next
    
    cmbEventDuration.AddItem "5 minutes"
    cmbEventDuration.AddItem "10 minutes"
    cmbEventDuration.AddItem "15 minutes"
    cmbEventDuration.AddItem "30 minutes"
    cmbEventDuration.AddItem "60 minutes (1 hour)"
    
    cmbEventDuration.AddItem "120 minutes (2 hours)"
    cmbEventDuration.AddItem "180 minutes (3 hours)"
    cmbEventDuration.AddItem "240 minutes (4 hours)"
    cmbEventDuration.AddItem 24 * 60 & " minutes (1 day)"
    cmbEventDuration.AddItem 2 * 24 * 60 & " minutes (2 days)"
    
    MonthlyInitialization
    
    YearlyInitialization
    
   
'    oprPatternNoEnd.OptionValue = xtpCalendarPatternEndNoDate
'    oprPatternEndAfter.OptionValue = xtpCalendarPatternEndAfterOccurrences
'    oprPatternEndByDate.OptionValue = xtpCalendarPatternEndDate
    '===
    UpdateControlsFromEvent
End Sub

Private Sub MonthlyInitialization()
    On Error Resume Next
    
    If cmbMonthlyDayOfMonth.ListCount = 0 Then
        Dim cnt As Integer
        For cnt = 1 To 31
            cmbMonthlyDayOfMonth.AddItem cnt, cnt - 1
        Next cnt

        If frmMain.CalendarControl.ActiveView.Selection.IsValid Then
            cmbMonthlyDayOfMonth.ListIndex = Day(frmMain.CalendarControl.ActiveView.Selection.Begin) - 1
        End If
    End If

    If cmbMonthlyDay.ListCount = 0 Then
        cmbMonthlyDay.AddItem "first", 0
        cmbMonthlyDay.AddItem "second", 1
        cmbMonthlyDay.AddItem "third", 2
        cmbMonthlyDay.AddItem "fourth", 3
        cmbMonthlyDay.AddItem "last", 4

        cmbMonthlyDay.ListIndex = 0
    End If

    If cmbMonthlyDayOfWeek.ListCount = 0 Then
        cmbMonthlyDayOfWeek.AddItem "Day", 0
        cmbMonthlyDayOfWeek.AddItem "WeekDay", 1
        cmbMonthlyDayOfWeek.AddItem "WeekendDay", 2
        cmbMonthlyDayOfWeek.AddItem "Sunday", 3
        cmbMonthlyDayOfWeek.AddItem "Monday", 4
        cmbMonthlyDayOfWeek.AddItem "Tuesday", 5
        cmbMonthlyDayOfWeek.AddItem "Wednesday", 6
        cmbMonthlyDayOfWeek.AddItem "Thursday", 7
        cmbMonthlyDayOfWeek.AddItem "Friday", 8
        cmbMonthlyDayOfWeek.AddItem "Saturday", 9

        If frmMain.CalendarControl.ActiveView.Selection.IsValid Then
            cmbMonthlyDayOfWeek.ListIndex = Weekday(frmMain.CalendarControl.ActiveView.Selection.Begin) + 2
        End If
    End If

    If txtMonthlyEveryMonth.Text = "" Then
        txtMonthlyEveryMonth.Text = "1"
    End If

    If txtMonthlyOfEveryTheMonths.Text = "" Then
        txtMonthlyOfEveryTheMonths.Text = "1"
    End If
End Sub

Private Sub YearlyInitialization()
    On Error Resume Next
    
    If cmbYearlyDate.ListCount = 0 Then
        Dim cnt As Integer
        For cnt = 1 To 31
            cmbYearlyDate.AddItem cnt, cnt - 1
        Next cnt

        If frmMain.CalendarControl.ActiveView.Selection.IsValid Then
            cmbYearlyDate.ListIndex = Day(frmMain.CalendarControl.ActiveView.Selection.Begin) - 1
        End If
    End If

    If cmdYearlyThePartOfWeek.ListCount = 0 Then
        cmdYearlyThePartOfWeek.AddItem "first", 0
        cmdYearlyThePartOfWeek.AddItem "second", 1
        cmdYearlyThePartOfWeek.AddItem "third", 2
        cmdYearlyThePartOfWeek.AddItem "fourth", 3
        cmdYearlyThePartOfWeek.AddItem "last", 4

        cmdYearlyThePartOfWeek.ListIndex = 0
    End If

    If cmbYearlyTheDay.ListCount = 0 Then
        cmbYearlyTheDay.AddItem "Day", 0
        cmbYearlyTheDay.AddItem "WeekDay", 1
        cmbYearlyTheDay.AddItem "WeekendDay", 2
        cmbYearlyTheDay.AddItem "Sunday", 3
        cmbYearlyTheDay.AddItem "Monday", 4
        cmbYearlyTheDay.AddItem "Tuesday", 5
        cmbYearlyTheDay.AddItem "Wednesday", 6
        cmbYearlyTheDay.AddItem "Thursday", 7
        cmbYearlyTheDay.AddItem "Friday", 8
        cmbYearlyTheDay.AddItem "Saturday", 9
        
        If frmMain.CalendarControl.ActiveView.Selection.IsValid Then
            cmbYearlyTheDay.ListIndex = Weekday(frmMain.CalendarControl.ActiveView.Selection.Begin) + 2
        End If
    End If

    If cmbYearlyEveryDate.ListCount = 0 Then
        
        cmbYearlyEveryDate.AddItem "January", 0
        cmbYearlyEveryDate.AddItem "February", 1
        cmbYearlyEveryDate.AddItem "March", 2
        cmbYearlyEveryDate.AddItem "April", 3
        cmbYearlyEveryDate.AddItem "May", 4
        cmbYearlyEveryDate.AddItem "June", 5
        cmbYearlyEveryDate.AddItem "July", 6
        cmbYearlyEveryDate.AddItem "August", 7
        cmbYearlyEveryDate.AddItem "September", 8
        cmbYearlyEveryDate.AddItem "October", 9
        cmbYearlyEveryDate.AddItem "November", 10
        cmbYearlyEveryDate.AddItem "December", 11
        
        Dim k As Integer
        For k = 0 To cmbYearlyEveryDate.ListCount - 1
            cmbYearlyTheMonth.AddItem cmbYearlyEveryDate.List(k), k
        Next
        
        If frmMain.CalendarControl.ActiveView.Selection.IsValid Then
            cmbYearlyEveryDate.ListIndex = Month(frmMain.CalendarControl.ActiveView.Selection.Begin) - 1
            cmbYearlyTheMonth.ListIndex = Month(frmMain.CalendarControl.ActiveView.Selection.Begin) - 1
        End If
    End If

End Sub

Private Sub SetActivePage(nPage As Long)
    optRecDaily.Value = (nPage = 1)
    optRecWeekly.Value = (nPage = 2)
    optRecMonthly.Value = (nPage = 3)
    optRecYearly.Value = (nPage = 4)
    
    pageDaily.Visible = (nPage = 1)
    pageWeekly.Visible = (nPage = 2)
    pageMonthly.Visible = (nPage = 3)
    pageYearly.Visible = (nPage = 4)
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

    cmbEventStartTime.Text = FormatDateTime(m_pRPattern.StartTime, vbShortTime)
    cmbEventDuration.Text = m_pRPattern.DurationMinutes & " minutes"
    cmbEventDuration_Change ' Update EndTime control
    
    SetActivePage 1
    '
    If m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceDaily Then
        SetActivePage 1
        
        If m_pRPattern.Options.DailyEveryWeekDayOnly Then
            optDailyEveryNdays = False
            optDailyEveryWorkDay = True
            txtDailyEveryNdays = 1
        Else
            optDailyEveryWorkDay = False
            optDailyEveryNdays = True
            txtDailyEveryNdays = m_pRPattern.Options.DailyIntervalDays
        End If
    ElseIf m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceWeekly Then
        SetActivePage 2
        
        txtWeeklyNWeeks = m_pRPattern.Options.WeeklyIntervalWeeks
        Dim nWDays As Long
        nWDays = m_pRPattern.Options.WeeklyDayOfWeekMask
        
        chkWeeklyMonday = IIf((nWDays And xtpCalendarDayMonday) <> 0, 1, 0)
        chkWeeklyTusday = IIf((nWDays And xtpCalendarDayTuesday) <> 0, 1, 0)
        chkWeeklyWednesday = IIf((nWDays And xtpCalendarDayWednesday) <> 0, 1, 0)
        chkWeeklyThursday = IIf((nWDays And xtpCalendarDayThursday) <> 0, 1, 0)
        chkWeeklyFriday = IIf((nWDays And xtpCalendarDayFriday) <> 0, 1, 0)
        
        chkWeeklySaturday = IIf((nWDays And xtpCalendarDaySaturday) <> 0, 1, 0)
        chkWeeklySunday = IIf((nWDays And xtpCalendarDaySunday) <> 0, 1, 0)
                
    ElseIf m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceMonthly Or _
        m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceMonthNth _
    Then
        SetActivePage 3
        
        If m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceMonthly Then
            optMonthlyDay = True
            
            cmbMonthlyDayOfMonth.ListIndex = m_pRPattern.Options.MonthlyDayOfMonth - 1
            txtMonthlyEveryMonth.Text = m_pRPattern.Options.MonthlyIntervalMonths
        Else
            optMonthlyThe = True
            
            cmbMonthlyDay.ListIndex = m_pRPattern.Options.MonthNthWhichDay - 1
            cmbMonthlyDayOfWeek.ListIndex = WhichDayMask2index(m_pRPattern.Options.MonthNthWhichDayMask)
            
            txtMonthlyOfEveryTheMonths.Text = m_pRPattern.Options.MonthNthIntervalMonths
        End If
    
    ElseIf m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceYearly Or _
        m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceYearNth _
    Then
        SetActivePage 4
        
        If m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceYearly Then
            optYearlyDay = True
            
            cmbYearlyEveryDate.ListIndex = m_pRPattern.Options.YearlyMonthOfYear - 1
            cmbYearlyDate.ListIndex = m_pRPattern.Options.YearlyDayOfMonth - 1
       
        Else
            optYearlyThe = True
            
            cmdYearlyThePartOfWeek.ListIndex = m_pRPattern.Options.YearNthWhichDay - 1
            cmbYearlyTheDay.ListIndex = WhichDayMask2index(m_pRPattern.Options.YearNthWhichDayMask)
            cmbYearlyTheMonth.ListIndex = m_pRPattern.Options.YearNthMonthOfYear - 1
        End If
    End If
    
    'Start-End pattern
    ddPatternStartDate = m_pRPattern.StartDate
    
    optPatternNoEnd = m_pRPattern.EndMethod = xtpCalendarPatternEndNoDate
    optPatternEndAfter = m_pRPattern.EndMethod = xtpCalendarPatternEndAfterOccurrences
    optPatternEndByDate = m_pRPattern.EndMethod = xtpCalendarPatternEndDate
        
    txtPatternEndAfter = "10"
    ddPatternEndDate = DateValue(Now() + 5)
    
    If m_pRPattern.EndMethod = xtpCalendarPatternEndAfterOccurrences Then
        txtPatternEndAfter = m_pRPattern.EndAfterOccurrences
    ElseIf m_pRPattern.EndMethod = xtpCalendarPatternEndDate Then
        ddPatternEndDate = m_pRPattern.EndDate
    End If
    
End Sub

Private Sub UpdatePatternFromControls()
    
    Dim arTimes
    arTimes = Split(cmbEventStartTime.Text, ":")
    m_pRPattern.StartTime = TimeSerial(arTimes(0), arTimes(1), 0)
    m_pRPattern.DurationMinutes = Val(cmbEventDuration.Text)

    If optRecDaily Then
        m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceDaily
                
        If optDailyEveryWorkDay Then
            m_pRPattern.Options.DailyEveryWeekDayOnly = True
        Else
            Debug.Assert optDailyEveryNdays
            m_pRPattern.Options.DailyEveryWeekDayOnly = False
            m_pRPattern.Options.DailyIntervalDays = Val(txtDailyEveryNdays)
        End If
        
    ElseIf optRecWeekly Then
        m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceWeekly

        m_pRPattern.Options.WeeklyIntervalWeeks = Val(txtWeeklyNWeeks)
        Dim nWDays As Long
        nWDays = 0
        
        nWDays = nWDays + IIf(chkWeeklyMonday, xtpCalendarDayMonday, 0)
        nWDays = nWDays + IIf(chkWeeklyTusday, xtpCalendarDayTuesday, 0)
        nWDays = nWDays + IIf(chkWeeklyWednesday, xtpCalendarDayWednesday, 0)
        nWDays = nWDays + IIf(chkWeeklyThursday, xtpCalendarDayThursday, 0)
        nWDays = nWDays + IIf(chkWeeklyFriday, xtpCalendarDayFriday, 0)
        
        nWDays = nWDays + IIf(chkWeeklySaturday, xtpCalendarDaySaturday, 0)
        nWDays = nWDays + IIf(chkWeeklySunday, xtpCalendarDaySunday, 0)
                
        m_pRPattern.Options.WeeklyDayOfWeekMask = nWDays
        
    ElseIf optRecMonthly Then
        
        If optMonthlyDay Then
            m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceMonthly
            
            m_pRPattern.Options.MonthlyDayOfMonth = cmbMonthlyDayOfMonth.ListIndex + 1
            m_pRPattern.Options.MonthlyIntervalMonths = CLng(Val(txtMonthlyEveryMonth.Text))
            If m_pRPattern.Options.MonthlyIntervalMonths < 1 Then m_pRPattern.Options.MonthlyIntervalMonths = 1
        
        Else
            Debug.Assert optMonthlyThe
            
            m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceMonthNth
            
            m_pRPattern.Options.MonthNthWhichDay = cmbMonthlyDay.ListIndex + 1
            m_pRPattern.Options.MonthNthWhichDayMask = index2WhichDayMask(cmbMonthlyDayOfWeek.ListIndex)
            
            m_pRPattern.Options.MonthNthIntervalMonths = CLng(Val(txtMonthlyOfEveryTheMonths.Text))
            If m_pRPattern.Options.MonthNthIntervalMonths < 1 Then m_pRPattern.Options.MonthNthIntervalMonths = 1
        End If

    ElseIf optRecYearly Then
        '
        If optYearlyDay Then
            m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceYearly
            
            m_pRPattern.Options.YearlyMonthOfYear = cmbYearlyEveryDate.ListIndex + 1
            m_pRPattern.Options.YearlyDayOfMonth = cmbYearlyDate.ListIndex + 1
        Else
            Debug.Assert optYearlyThe
            m_pRPattern.Options.RecurrenceType = xtpCalendarRecurrenceYearNth
            
            m_pRPattern.Options.YearNthWhichDay = cmdYearlyThePartOfWeek.ListIndex + 1
            m_pRPattern.Options.YearNthWhichDayMask = index2WhichDayMask(cmbYearlyTheDay.ListIndex)
            m_pRPattern.Options.YearNthMonthOfYear = cmbYearlyTheMonth.ListIndex + 1
        End If
    End If
            
    'Start-End pattern
    m_pRPattern.StartDate = DateValue(CDate(ddPatternStartDate))
    
    If optPatternNoEnd Then
        m_pRPattern.EndMethod = xtpCalendarPatternEndNoDate
        
    ElseIf optPatternEndAfter Then
        m_pRPattern.EndMethod = xtpCalendarPatternEndAfterOccurrences
        m_pRPattern.EndAfterOccurrences = Val(txtPatternEndAfter)
        
    ElseIf optPatternEndByDate Then
        m_pRPattern.EndMethod = xtpCalendarPatternEndDate
        m_pRPattern.EndDate = DateValue(CDate(ddPatternEndDate))
    End If
    
End Sub
    
    


Private Sub optRecDaily_Click()
    SetActivePage 1
End Sub

Private Sub optRecMonthly_Click()
    SetActivePage 3
End Sub

Private Sub optRecWeekly_Click()
    SetActivePage 2
End Sub

Private Sub optRecYearly_Click()
    SetActivePage 4
End Sub
