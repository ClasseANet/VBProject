VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Format Day/Week/Month View"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmAdditionalOpt 
      Caption         =   "Additional Options"
      Height          =   2055
      Left            =   3360
      TabIndex        =   32
      Top             =   2640
      Width           =   2775
      Begin VB.CheckBox chkMVShowEndTimeAlways 
         Caption         =   "Show End time Always"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox chkMVShowStartTimeAlways 
         Caption         =   "Show Start time Always"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox chkWVShowEndTimeAlways 
         Caption         =   "Show End time Always"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox chkWVShowStartTimeAlways 
         Caption         =   "Show Start time Always"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Month View:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Week View:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Other"
      Height          =   855
      Left            =   3360
      TabIndex        =   29
      Top             =   4800
      Width           =   2715
      Begin VB.ComboBox cmbToolTipsMode 
         Height          =   315
         Left            =   1140
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ToolTips:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame frmMonth 
      Caption         =   "Month view"
      Height          =   1815
      Left            =   60
      TabIndex        =   23
      Top             =   3840
      Width           =   3135
      Begin VB.ComboBox cmbWeeksCount 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1260
         Width           =   1455
      End
      Begin VB.CheckBox chkComperssWeekendDays 
         Caption         =   "Compress &weekend days"
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox chkShowEndTimeMonth 
         Caption         =   "Show end time"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkShowTimeAsClockMonth 
         Caption         =   "Show time as clock"
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblWeeksCount 
         Caption         =   "Weeks count"
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame frmWeek 
      Caption         =   "Week view"
      Height          =   1035
      Left            =   60
      TabIndex        =   20
      Top             =   2640
      Width           =   3135
      Begin VB.CheckBox chkShowEndTimeWeek 
         Caption         =   "Show end time"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkShowTimeAsClockWeek 
         Caption         =   "Show time as clock"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Frame frmWorkWeek 
      Caption         =   "Work week"
      Height          =   1575
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   6015
      Begin VB.ComboBox cmbEndTime 
         Height          =   315
         Left            =   4440
         TabIndex        =   16
         Text            =   "cmbEndTime"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cmbStartTime 
         Height          =   315
         ItemData        =   "frmSettings.frx":0000
         Left            =   4440
         List            =   "frmSettings.frx":0002
         TabIndex        =   15
         Text            =   "cmbStartTime"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbFirstDayOfWeek 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Sat"
         Height          =   195
         Index           =   6
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Fri"
         Height          =   195
         Index           =   5
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Thu"
         Height          =   195
         Index           =   4
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Wed"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Tue"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Mon"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "Sun"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblEndTime 
         Caption         =   "E&nd time:"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblStartTime 
         Caption         =   "Sta&rt time:"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label lblFirstDayOfWeek 
         Caption         =   "First day of w&eek:"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   765
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3203
      TabIndex        =   2
      Top             =   6300
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1703
      TabIndex        =   1
      Top             =   6300
      Width           =   1335
   End
   Begin VB.Frame frmDay 
      Caption         =   "Day view"
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   1800
      Width           =   6015
      Begin VB.CommandButton cmdTimeZone 
         Caption         =   "Time zone ..."
         Height          =   315
         Left            =   4440
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbTimeScale 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTimeScale 
         AutoSize        =   -1  'True
         Caption         =   "Time scale:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Property Get CalendarControl() As CalendarControl
    Set CalendarControl = frmMain.CalendarControl
End Property

Private Sub chkMVShowStartTimeAlways_Click()
    chkMVShowEndTimeAlways.Enabled = chkMVShowStartTimeAlways.Value <> 0
    If Not chkMVShowEndTimeAlways.Enabled Then
        chkMVShowEndTimeAlways.Value = 0
    End If
    
End Sub

Private Sub chkWVShowStartTimeAlways_Click()
     chkWVShowEndTimeAlways.Enabled = chkWVShowStartTimeAlways.Value <> 0
     If Not chkWVShowEndTimeAlways.Enabled Then
        chkWVShowEndTimeAlways.Value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ApplySettings
    CalendarControl.Populate
    
    Unload Me
End Sub

Sub AddTimeScale(TimeScale As Long)
    cmbTimeScale.AddItem TimeScale & " minutes"
    cmbTimeScale.ItemData(cmbTimeScale.ListCount - 1) = TimeScale

    If CalendarControl.DayView.TimeScale = TimeScale Then cmbTimeScale.ListIndex = cmbTimeScale.ListCount - 1
End Sub

Sub AddCalendarDay(Index As Long, Day As CalendarWeekDay, Caption As String, FirstDayOfTheWeek As Long)
    chkWorkDay(Index).Value = IIf(CalendarControl.Options.WorkWeekMask And Day, 1, 0)
    
    cmbFirstDayOfWeek.AddItem Caption
    If (CalendarControl.Options.FirstDayOfTheWeek = FirstDayOfTheWeek) Then cmbFirstDayOfWeek.ListIndex = Index
    
End Sub

Private Sub cmdTimeZone_Click()
    If frmMain.g_bUseBuiltInCalendarDialogs Then
        Dim dlgCalendar As New CalendarDialogs
        dlgCalendar.ParentHWND = Me.hwnd
        dlgCalendar.Calendar = CalendarControl
        
        dlgCalendar.ShowTimeScaleProperties
        Exit Sub
    End If
    
    frmTimeZone.Show vbModal, Me
End Sub

Private Sub Form_Load()
    AddTimeScale 5
    AddTimeScale 6
    AddTimeScale 10
    AddTimeScale 15
    AddTimeScale 30
    AddTimeScale 60
    
    Dim WorkWeekMask As CalendarWeekDay
    WorkWeekMask = CalendarControl.Options.WorkWeekMask
    
    AddCalendarDay 0, xtpCalendarDaySunday, "Sunday", 1
    AddCalendarDay 1, xtpCalendarDayMonday, "Monday", 2
    AddCalendarDay 2, xtpCalendarDayTuesday, "Tuesday", 3
    AddCalendarDay 3, xtpCalendarDayWednesday, "Wednesday", 4
    AddCalendarDay 4, xtpCalendarDayThursday, "Thursday", 5
    AddCalendarDay 5, xtpCalendarDayFriday, "Friday", 6
    AddCalendarDay 6, xtpCalendarDaySaturday, "Saturday", 7
    
    cmbStartTime.Text = CalendarControl.Options.WorkDayStartTime
    cmbEndTime.Text = CalendarControl.Options.WorkDayEndTime
    
    chkShowTimeAsClockWeek.Value = IIf(CalendarControl.Options.WeekViewShowTimeAsClocks, 1, 0)
    chkShowEndTimeWeek.Value = IIf(CalendarControl.Options.WeekViewShowEndDate, 1, 0)
    
    chkShowTimeAsClockMonth.Value = IIf(CalendarControl.Options.MonthViewShowTimeAsClocks, 1, 0)
    chkShowEndTimeMonth.Value = IIf(CalendarControl.Options.MonthViewShowEndDate, 1, 0)
    chkComperssWeekendDays.Value = IIf(CalendarControl.Options.MonthViewCompressWeekendDays, 1, 0)
    
    cmbWeeksCount.AddItem "2"
    cmbWeeksCount.AddItem "3"
    cmbWeeksCount.AddItem "4"
    cmbWeeksCount.AddItem "5"
    cmbWeeksCount.AddItem "6"
    cmbWeeksCount.ListIndex = CalendarControl.MonthView.WeeksCount - 2
    
    cmbToolTipsMode.AddItem "Standard"
    cmbToolTipsMode.ItemData(cmbToolTipsMode.NewIndex) = 0
    
    cmbToolTipsMode.AddItem "Custom"
    cmbToolTipsMode.ItemData(cmbToolTipsMode.NewIndex) = 1
    
    cmbToolTipsMode.AddItem "Disabled"
    cmbToolTipsMode.ItemData(cmbToolTipsMode.NewIndex) = 2
    
    Dim i As Long
    For i = 0 To 2
        If cmbToolTipsMode.ItemData(i) = frmMain.ToolTips_Mode Then
            cmbToolTipsMode.ListIndex = i
            Exit For
        End If
    Next
    
    '---------------------------------------------------------------
    chkWVShowStartTimeAlways.Value = IIf(CalendarControl.Options.AdditionalOptionsFlags.IsFlagSet( _
                                     xtpCalendarOptWeekViewShowStartTimeAlways), 1, 0)
    chkWVShowEndTimeAlways.Value = IIf(CalendarControl.Options.AdditionalOptionsFlags.IsFlagSet( _
                                     xtpCalendarOptWeekViewShowEndTimeAlways), 1, 0)
    
    chkMVShowStartTimeAlways.Value = IIf(CalendarControl.Options.AdditionalOptionsFlags.IsFlagSet( _
                                     xtpCalendarOptMonthViewShowStartTimeAlways), 1, 0)
    chkMVShowEndTimeAlways.Value = IIf(CalendarControl.Options.AdditionalOptionsFlags.IsFlagSet( _
                                     xtpCalendarOptMonthViewShowEndTimeAlways), 1, 0)
                                    
    chkMVShowStartTimeAlways_Click
    chkWVShowStartTimeAlways_Click
        
    frmMain.ModalFormsRunningCounter = frmMain.ModalFormsRunningCounter + 1
End Sub


Sub ApplyCalendarDay(Index As Long, Day As CalendarWeekDay, FirstDayOfTheWeek As Long)
    
    If (chkWorkDay(Index).Value) Then CalendarControl.Options.WorkWeekMask = CalendarControl.Options.WorkWeekMask Or Day
    
    If (cmbFirstDayOfWeek.ListIndex = Index) Then CalendarControl.Options.FirstDayOfTheWeek = FirstDayOfTheWeek
    
End Sub


Sub ApplySettings()
    Dim eViewType As Long
    eViewType = CalendarControl.ViewType
    
    CalendarControl.Options.WorkWeekMask = 0

    ApplyCalendarDay 0, xtpCalendarDaySunday, 1
    ApplyCalendarDay 1, xtpCalendarDayMonday, 2
    ApplyCalendarDay 2, xtpCalendarDayTuesday, 3
    ApplyCalendarDay 3, xtpCalendarDayWednesday, 4
    ApplyCalendarDay 4, xtpCalendarDayThursday, 5
    ApplyCalendarDay 5, xtpCalendarDayFriday, 6
    ApplyCalendarDay 6, xtpCalendarDaySaturday, 7
    
    
    CalendarControl.DayView.TimeScale = cmbTimeScale.ItemData(cmbTimeScale.ListIndex)
    
    CalendarControl.Options.WeekViewShowTimeAsClocks = chkShowTimeAsClockWeek.Value
    CalendarControl.Options.WeekViewShowEndDate = chkShowEndTimeWeek.Value
    
    CalendarControl.Options.MonthViewShowTimeAsClocks = chkShowTimeAsClockMonth.Value
    CalendarControl.Options.MonthViewShowEndDate = chkShowEndTimeMonth.Value
    CalendarControl.Options.MonthViewCompressWeekendDays = chkComperssWeekendDays.Value
    
    CalendarControl.MonthView.WeeksCount = cmbWeeksCount.ListIndex + 2
    
    CalendarControl.Options.WorkDayStartTime = TimeValue(cmbStartTime.Text)
    CalendarControl.Options.WorkDayEndTime = TimeValue(cmbEndTime.Text)
    
    '---------------------------------------------------------------
    CalendarControl.Options.AdditionalOptionsFlags.Flags = 0
    
    If chkWVShowStartTimeAlways.Value <> 0 Then
        CalendarControl.Options.AdditionalOptionsFlags.SetFlag xtpCalendarOptWeekViewShowStartTimeAlways
    End If
    
    If chkWVShowEndTimeAlways.Value <> 0 Then
        CalendarControl.Options.AdditionalOptionsFlags.SetFlag xtpCalendarOptWeekViewShowEndTimeAlways
    End If
    
    If chkMVShowStartTimeAlways.Value <> 0 Then
        CalendarControl.Options.AdditionalOptionsFlags.SetFlag xtpCalendarOptMonthViewShowStartTimeAlways
    End If
        
    If chkMVShowEndTimeAlways.Value <> 0 Then
        CalendarControl.Options.AdditionalOptionsFlags.SetFlag xtpCalendarOptMonthViewShowEndTimeAlways
    End If

    'to apply WorkWeekMask changes
    CalendarControl.ViewType = eViewType
        
    frmMain.ToolTips_Mode = cmbToolTipsMode.ItemData(cmbToolTipsMode.ListIndex)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.ModalFormsRunningCounter = frmMain.ModalFormsRunningCounter - 1
End Sub
