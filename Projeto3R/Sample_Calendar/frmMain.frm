VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#11.2#0"; "Codejock.Calendar.v11.2.2.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000010&
   Caption         =   "CalendarControl Sample"
   ClientHeight    =   6300
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Day"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Work Week"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Week"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Month"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Open provider"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Print Page Setup"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Print Preview"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin XtremeCalendarControl.DatePicker wndDatePicker 
      Height          =   4215
      Left            =   6600
      TabIndex        =   5
      Top             =   840
      Width           =   2535
      _Version        =   720898
      _ExtentX        =   4471
      _ExtentY        =   7435
      _StockProps     =   64
      RowCount        =   2
   End
   Begin XtremeCalendarControl.CalendarControl CalendarControl 
      Height          =   4155
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   6555
      _Version        =   720898
      _ExtentX        =   11562
      _ExtentY        =   7329
      _StockProps     =   64
   End
   Begin VB.Timer timerRMDForm 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1500
      Top             =   420
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6045
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   6135
   End
   Begin ComctlLib.ImageList ilToolBar 
      Left            =   1980
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0552
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1548
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1652
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":175C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1866
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1970
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      Caption         =   " Calendar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenDataProvider 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page &Setup"
      End
      Begin VB.Menu mnuPrintCalendar 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMultiSchedulesSimple 
         Caption         =   "Load sample MultiSchedules (simple)"
      End
      Begin VB.Menu mnuMultiSchedulesExtended 
         Caption         =   "Load sample MultiSchedules (Extended)"
      End
      Begin VB.Menu mnuResourcesManager 
         Caption         =   "Open Resources Manager form"
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuCalendar 
      Caption         =   "Calendar"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuAdvancedOptions 
         Caption         =   "&Advanced Options"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnableTheme2007 
         Caption         =   "Enable Office 2007 Theme"
      End
      Begin VB.Menu mnuCustomizeTheme2007 
         Caption         =   "Customize Office 2007 Theme"
      End
      Begin VB.Menu mnuEmpty1 
         Caption         =   "---------------------------------------------"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCustomIcons 
         Caption         =   "Show custom icons example"
      End
      Begin VB.Menu mnuShowDynamicCustomization 
         Caption         =   "Show dynamic customization example"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReminders 
         Caption         =   "&Reminders Window"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuContextMenu 
      Caption         =   "Context Menu"
      Begin VB.Menu mnuContexEditEvent 
         Caption         =   "Edit Event Context Menu"
         Begin VB.Menu mnuOpenEvent 
            Caption         =   "&Open"
         End
         Begin VB.Menu mnuDeleteEvent 
            Caption         =   "&Delete"
         End
      End
      Begin VB.Menu mnuContextNewEvent 
         Caption         =   "New Event Context Menu"
         Begin VB.Menu mnuNewEvent 
            Caption         =   "&New Event"
         End
      End
      Begin VB.Menu mnuContextTimeScale 
         Caption         =   "Time Scale Context Menu"
         Begin VB.Menu mnuNewEvent2 
            Caption         =   "&New Event"
         End
         Begin VB.Menu mnuChangeTimeZone 
            Caption         =   "Change Time Zone"
         End
         Begin VB.Menu mnuSeparator 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTimeScale 
            Caption         =   "6&0 Minutes"
            HelpContextID   =   60
            Index           =   1
         End
         Begin VB.Menu mnuTimeScale 
            Caption         =   "&30 Minutes"
            HelpContextID   =   30
            Index           =   2
         End
         Begin VB.Menu mnuTimeScale 
            Caption         =   "&15 Minutes"
            HelpContextID   =   15
            Index           =   3
         End
         Begin VB.Menu mnuTimeScale 
            Caption         =   "10 &Minutes"
            HelpContextID   =   10
            Index           =   4
         End
         Begin VB.Menu mnuTimeScale 
            Caption         =   "&5 Minutes"
            HelpContextID   =   5
            Index           =   5
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_pCustomDataHandler  As Object
Public m_eActiveDataProvider As CodeJockCalendarDataType

Dim ContextEvent As CalendarEvent

Public ModalFormsRunningCounter As Long

Public DisableDragging_ForRecurrenceEvents As Boolean
Public DisableInPlaceCreateEvents_ForSaSu As Boolean

Public EnableScrollV_DayView As Boolean
Public EnableScrollH_DayView As Boolean

Public EnableScrollV_WeekView As Boolean

Public EnableScrollV_MonthView As Boolean

Public g_DataResourcesMan As New CalendarResourcesManager

Public g_bUseBuiltInCalendarDialogs As Boolean
Public g_dlgCalendarReminders As New CalendarDialogs
'--------------------------
' 0 - Standard
' 1 - Custom
' 2 - Disabled
'--------------------------
Private nToolTips_Mode As Long

Public Property Get ToolTips_Mode() As Long
   ToolTips_Mode = nToolTips_Mode
End Property

Public Property Let ToolTips_Mode(ByRef nMode As Long)

   nToolTips_Mode = nMode
   CalendarControl.EnableToolTips (nMode < 2)
   
    If nMode = 1 Then
        CalendarControl.AskItemTextFlags.SetFlag xtpCalendarItemText_EventToolTipText
    Else
        CalendarControl.AskItemTextFlags.ResetFlag xtpCalendarItemText_EventToolTipText
    End If
   
End Property

Private Sub CalendarControl_BeforeDrawDayViewCell(ByVal CellParams As XtremeCalendarControl.CalendarDayViewCellParams)
    ' standard colors are
    ' non-work cell Bk = RGB(255, 244, 188)
    '     work cell Bk = RGB(255, 255, 213)

    If CellParams.Selected Then
        'CellParams.BackgroundColor = RGB(20, 250, 50)
        Exit Sub
    End If
    
    If TimeValue(CellParams.BeginTime) >= #1:00:00 PM# And TimeValue(CellParams.BeginTime) < #2:00:00 PM# _
       And Weekday(CellParams.BeginTime) <> 1 And Weekday(CellParams.BeginTime) <> 7 Then
        CellParams.BackgroundColor = RGB(198, 198, 198)
    End If
End Sub

Private Sub CalendarControl_BeforeDrawThemeObject(ByVal eObjType As XtremeCalendarControl.CalendarBeforeDrawThemeObject, ByVal Params As Variant)
    Dim pTheme2007 As CalendarThemeOffice2007
    Set pTheme2007 = CalendarControl.Theme
    
    If eObjType = xtpCalendarBeforeDraw_DayViewDay Then
        Debug.Print "BitmapID=" & pTheme2007.DayView.Day.Header.BackgroundNormal.BitmapID
        pTheme2007.DayView.Day.Header.BackgroundNormal.BitmapID = 0
    End If
    
    If eObjType = xtpCalendarBeforeDraw_MonthViewWeekDayHeader Then
        pTheme2007.MonthView.WeekDayHeader.BackgroundNormal.BitmapID = 0
    End If
    
    If eObjType = xtpCalendarBeforeDraw_MonthViewWeekHeader Then
        pTheme2007.MonthView.WeekHeader.BaseColor = RGB((Params + 1) * 11, (Params + 1) * (Params + 1) * 7 Mod 256, (Params + 1) * 311 Mod 256)
        pTheme2007.MonthView.WeekHeader.BackgroundNormal.BitmapID = 0
    End If
    
    If eObjType = xtpCalendarBeforeDraw_DayViewTimeScale Then
        If Params = 2 Then
            pTheme2007.DayView.TimeScale.TimeTextBigBase.Color = RGB(200, 0, 0)
            pTheme2007.DayView.TimeScale.TimeTextSmall.Color = RGB(0, 170, 0)
            pTheme2007.DayView.TimeScale.BackgroundColor = RGB(170, 170, 55)
        End If
    ElseIf eObjType = xtpCalendarBeforeDraw_DayViewTimeScaleCell Then
        'CalendarControl.Theme.DayView.TimeScale.TimeTextBigBase.Font.Strikethrough = (Second(Now) Mod 2) = 0
        Dim nSubCellIndex
        nSubCellIndex = Params.Minutes / CalendarControl.DayView.TimeScale
        
        If Params.Minutes = -1 Then
            pTheme2007.DayView.TimeScale.TimeTextBigBase.Font.Bold = (Params.Index Mod 2) = 0
            pTheme2007.DayView.TimeScale.TimeTextSmall.Font.Bold = (Params.Index Mod 2) = 0
            pTheme2007.DayView.TimeScale.TimeTextSmall.Color = 20000 + 1111 * Params.Index
        Else
            If (nSubCellIndex Mod 2) = 0 And CalendarControl.DayView.TimeScale < 15 Then
                pTheme2007.DayView.TimeScale.ShowMinutes = False
            Else
                pTheme2007.DayView.TimeScale.TimeTextBigBase.Font.Bold = nSubCellIndex Mod 4 = 0
                pTheme2007.DayView.TimeScale.TimeTextBigBase.Color = _
                    RGB(33 * nSubCellIndex Mod 200, _
                        77 * Params.Minutes Mod 55, _
                        17 * Params.Minutes Mod 155)
                                
                pTheme2007.DayView.TimeScale.AmPmText.Color = RGB(255, 255, 255)
                pTheme2007.DayView.TimeScale.AmPmText.Font.Bold = True
                
                
                pTheme2007.DayView.TimeScale.TimeTextSmall.Font.Bold = False
                pTheme2007.DayView.TimeScale.TimeTextSmall.Color = 20000 + 111 * Params.Index
                
            End If
        
        End If
        
        If CalendarControl.DayView.TimeScale > 15 Then
            CalendarControl.Theme.DayView.TimeScale.TimeTextBigBase.Color = 20000 + 100 * Params.Index
        End If
        CalendarControl.Theme.DayView.TimeScale.LineColor = CalendarControl.Theme.DayView.TimeScale.TimeTextBigBase.Color
        
    ElseIf eObjType = xtpCalendarBeforeDraw_DayViewTimeScaleCaption Then
        CalendarControl.Theme.DayView.TimeScale.Caption.Font.Italic = (Second(Now) Mod 2) = 0
        CalendarControl.Theme.DayView.TimeScale.Caption.Color = RGB(10, 200, 100)
        CalendarControl.Theme.DayView.TimeScale.LineColor = RGB(10, 200, 100)
        
    ElseIf eObjType = xtpCalendarBeforeDraw_DayViewCell Then
        Dim pCell As CalendarThemeDayViewCellParams
        Set pCell = Params
        
        If pCell.Index Mod 3 = 0 Then
            If pCell.WorkCell Then
                pTheme2007.DayView.Day.Group.Cell.WorkCell.BackgroundColor = RGB(200, 255, 0)
            Else
                pTheme2007.DayView.Day.Group.Cell.NonWorkCell.BackgroundColor = RGB(255, 200, 0)
            End If
        End If
        
    ElseIf eObjType = xtpCalendarBeforeDraw_DayViewDay Then
        pTheme2007.DayView.Day.Header.BaseColor = 20000 + 100 * Weekday(Params.Date)
        pTheme2007.DayView.Day.Header.TodayBaseColor = RGB(0, 0, 0)
    
    ElseIf eObjType = xtpCalendarBeforeDraw_MonthViewDay Then
        pTheme2007.MonthView.Day.Header.BaseColor = RGB(Day(Params.Date) * 11 Mod 255, Day(Params.Date) * 257 Mod 255, Day(Params.Date) * 1001 Mod 255)
        pTheme2007.MonthView.Day.Header.TodayBaseColor = RGB(255, 0, 0)
        pTheme2007.MonthView.Day.TodayBorderColor = pTheme2007.MonthView.Day.Header.TodayBaseColor
    
    ElseIf eObjType = xtpCalendarBeforeDraw_MonthViewWeekDayHeader Then
        
        If Params >= 0 Then
            pTheme2007.MonthView.WeekDayHeader.BaseColor = _
                RGB(Params * 55 Mod 155 + 100, _
                    Params * 7 Mod 55 + 200, _
                    Params * 101 Mod 200 + 55)
        End If
        
    End If
    
End Sub

Private Sub CalendarControl_BeforeEditOperation(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, CancelOperation As Boolean)
'    If OpParams.Operation = xtpCalendarEO_DragCopy Or _
'        OpParams.Operation = xtpCalendarEO_DragMove _
'    Then
'        If OpParams.DraggingEvent.AllDayEvent <> OpParams.DraggingEventNew.AllDayEvent Then
'            CancelOperation = True
'        End If
'    End If
End Sub

Public Sub OpenProvider(ByVal eDataProviderType As CodeJockCalendarDataType, ByVal strConnectionString As String)
    Set m_pCustomDataHandler = Nothing
    
    ' SQL Server provider
    If eDataProviderType = cjCalendarData_SQLServer Then
        Set m_pCustomDataHandler = New providerSQLServer
        '' Create DSN "Calendar_SQLServer" to connect to SQL Server Calendar DB
        m_pCustomDataHandler.OpenDB strConnectionString
        
        m_pCustomDataHandler.SetCalendar CalendarControl
    End If
    
    ' MySQL provider
    If eDataProviderType = cjCalendarData_MySQL Then
        Set m_pCustomDataHandler = New providerMySQL
        m_pCustomDataHandler.OpenDB strConnectionString
        
        m_pCustomDataHandler.SetCalendar CalendarControl
    End If
                
    CalendarControl.SetDataProvider strConnectionString
        
    If eDataProviderType = cjCalendarData_SQLServer Or eDataProviderType = cjCalendarData_MySQL Then
        CalendarControl.DataProvider.CacheMode = xtpCalendarDPCacheModeOnRepeat
    End If
    
    If Not CalendarControl.DataProvider.Open Then
        CalendarControl.DataProvider.Create
    End If
    
    m_eActiveDataProvider = eDataProviderType
        
    CalendarControl.Populate
    wndDatePicker.RedrawControl

End Sub

Private Sub CalendarControl_ContextMenu(ByVal X As Single, ByVal Y As Single)

    Debug.Print "On context menu"
    
    Dim HitTest As CalendarHitTestInfo
    Set HitTest = CalendarControl.ActiveView.HitTest
    
    If Not HitTest.ViewEvent Is Nothing Then
        Set ContextEvent = HitTest.ViewEvent.Event
        Me.PopupMenu mnuContexEditEvent
        Set ContextEvent = Nothing
    ElseIf (HitTest.HitCode = xtpCalendarHitTestDayViewTimeScale) Then
        Me.PopupMenu mnuContextTimeScale
    Else
        Me.PopupMenu mnuContextNewEvent
    End If

End Sub

Private Sub CalendarControl_DblClick()
    Dim HitTest As CalendarHitTestInfo
    Set HitTest = CalendarControl.ActiveView.HitTest
    
    Dim Events As CalendarEvents
    If Not HitTest.HitCode = xtpCalendarHitTestUnknown Then
     '   Set Events = CalendarControl.DataProvider.RetrieveDayEvents(HitTest.ViewDay.Date)
    End If
    
    If HitTest.ViewEvent Is Nothing Then
        mnuNewEvent_Click
    Else
        ModifyEvent HitTest.ViewEvent.Event
    End If
End Sub

Private Sub CalendarControl_EventChanged(ByVal EventID As Long)
    Dim pEvent As CalendarEvent
    Set pEvent = CalendarControl.DataProvider.GetEvent(EventID)
    
    ' pEvent Is Nothing for recurrence Ocurrence and Exception.
    ' See CalendarControl_PatternChanged for recurrence events changes.
    If Not pEvent Is Nothing Then
        
    End If
End Sub

Private Sub CalendarControl_GetItemText(ByVal Params As XtremeCalendarControl.CalendarGetItemTextParams)
    If Params.Item = xtpCalendarItemText_EventSubject Then
        Params.Text = "*Custom subject* " & Params.Text
    
    ElseIf Params.Item = xtpCalendarItemText_EventLocation Then
        Params.Text = "ScheduleID = " & Params.ViewEvent.Event.ScheduleID
        
    ElseIf Params.Item = xtpCalendarItemText_EventBody Then
        Params.Text = "custom BODY text."
        
    ElseIf Params.Item = xtpCalendarItemText_DayViewDayHeader Or _
           Params.Item = xtpCalendarItemText_WeekViewDayHeader Or _
           Params.Item = xtpCalendarItemText_MonthViewDayHeader _
        Then
        Params.Text = "Date:" & Params.ViewDay.Date
    
    ElseIf Params.Item = xtpCalendarItemText_DayViewDayHeaderLeft Or _
            Params.Item = xtpCalendarItemText_WeekViewDayHeaderLeft Or _
            Params.Item = xtpCalendarItemText_MonthViewDayHeaderLeft _
        Then
        Params.Text = Day(Params.ViewDay.Date)
    
    ElseIf Params.Item = xtpCalendarItemText_DayViewDayHeaderCenter Or _
           Params.Item = xtpCalendarItemText_WeekViewDayHeaderCenter Or _
           Params.Item = xtpCalendarItemText_MonthViewDayHeaderCenter _
        Then
        Params.Text = WeekdayName(Weekday(Params.ViewDay.Date), , vbSunday)
    
    ElseIf Params.Item = xtpCalendarItemText_DayViewDayHeaderRight Or _
           Params.Item = xtpCalendarItemText_WeekViewDayHeaderRight Or _
           Params.Item = xtpCalendarItemText_MonthViewDayHeaderRight _
        Then
        Params.Text = Year(Params.ViewDay.Date)
        
    ElseIf Params.Item = xtpCalendarItemText_MonthViewWeekDayHeader Then
        Params.Text = Params.Text & " - " & Params.Weekday
        
    ElseIf Params.Item = xtpCalendarItemText_EventToolTipText Then
        Params.Text = "ID = [" & Params.ViewEvent.Event.Id & "]  " & vbCrLf & Params.ViewEvent.Event.Subject & _
                vbCrLf & Params.ViewEvent.Event.Location & vbCrLf & Params.ViewEvent.Event.Body
    End If
    
End Sub

Private Sub CalendarControl_GotFocus()
    Debug.Print "GotFocus"
End Sub

Private Sub CalendarControl_IsEditOperationDisabled(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, DisableOperation As Boolean)

    If DisableDragging_ForRecurrenceEvents Then
        If OpParams.Operation = xtpCalendarEO_DragCopy Or _
            OpParams.Operation = xtpCalendarEO_DragMove Or _
            OpParams.Operation = xtpCalendarEO_DragResizeBegin Or _
            OpParams.Operation = xtpCalendarEO_DragResizeEnd _
        Then
            If OpParams.EventViews(0).Event.RecurrenceState <> xtpCalendarRecurrenceNotRecurring Then
                DisableOperation = True
            End If
        End If
    End If
        
    If DisableInPlaceCreateEvents_ForSaSu Then
        If OpParams.Operation = xtpCalendarEO_InPlaceCreateEvent Then
            Dim dtBegin As Date, dtEnd As Date, bAllDay As Boolean
            Dim nSelDays As Long, nSelWDay As Long
                                   
            If CalendarControl.ActiveView.GetSelection(dtBegin, dtEnd, bAllDay) = False Then
                Exit Sub
            End If
            
            nSelDays = Abs(DateDiff("d", dtEnd, dtBegin))
            If dtBegin < dtEnd Then
                nSelWDay = Weekday(dtBegin)
            Else
                nSelWDay = Weekday(dtEnd)
            End If
            
            If bAllDay And nSelDays > 0 Then
                nSelDays = nSelDays - 1
            End If
            
            If nSelWDay = 1 Or (nSelWDay + nSelDays) >= 7 Then
                DisableOperation = True
            End If
        End If
    End If

End Sub

Private Sub CalendarControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Debug.Print "KeyDown"
    Dim BeginSelection As Date, EndSelection As Date, AllDay As Boolean

    If CalendarControl.ActiveView.GetSelection(BeginSelection, EndSelection, AllDay) Then
        Debug.Print "Selection: "; BeginSelection; " - "; EndSelection
    End If

End Sub

Private Sub CalendarControl_LostFocus()
Debug.Print "LostFocus"
End Sub

Private Sub CalendarControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim HitTest As CalendarHitTestInfo
    Set HitTest = CalendarControl.ActiveView.HitTest
    
    If (Not HitTest.ViewEvent Is Nothing) Then
       Debug.Print "MouseMove. HitTest = "; HitTest.ViewEvent.Event.Subject
    End If
    
    UpdateToolbar
        
'    Debug.Print "HitTest.HitCode = " & Hex(HitTest.HitCode)
    
'   If HitTest.HitCode And xtpCalendarHitTestDayArea Then
'        Debug.Print "HitTest DayArea"
'    End If
'    If HitTest.HitCode And xtpCalendarHitTestDayHeader Then
'        Debug.Print "HitTest DayHEADER"
'    End If
    
'    If HitTest.HitCode And xtpCalendarHitTestGroupArea Then
'        Debug.Print "HitTest GroupArea"
'    End If
'    If HitTest.HitCode And xtpCalendarHitTestGroupHeader Then
'        Debug.Print "HitTest GroupHeader"
'    End If
    
'    If HitTest.HitCode And xtpCalendarHitTestDayViewAllDayEvent Then
'        Debug.Print "HitTest AllDayEvent"
'    End If
'    If HitTest.HitCode And xtpCalendarHitTestDayViewCell Then
'        Debug.Print "HitTest DayViewCell"
'    End If
'    If HitTest.HitCode And xtpCalendarHitTestDayViewTimeScale Then
'        Debug.Print "HitTest TimeScale"
'    End If
        
End Sub

Private Sub CalendarControl_OnReminders(ByVal Action As XtremeCalendarControl.CalendarRemindersAction, ByVal Reminder As XtremeCalendarControl.CalendarReminder)
    
    If g_bUseBuiltInCalendarDialogs Then
        g_dlgCalendarReminders.ShowRemindersWindow
        Exit Sub
    End If
    
    frmReminders.OnReminders Action, Reminder
    
    If Action = xtpCalendarRemindersFire Then
        If ModalFormsRunningCounter = 0 Then
            'frmReminders.Visible = True
            frmReminders.Show vbModeless, Me
        Else
            timerRMDForm.Enabled = True
        End If
    End If
End Sub

Private Sub CalendarControl_PrePopulate(ByVal ViewGroup As XtremeCalendarControl.CalendarViewGroup, ByVal Events As XtremeCalendarControl.CalendarEvents)
    'If Events.Count = 0 Then
    '    Events.Add CalendarControl.DataProvider.CreateEvent
    '    Events(0).Label = 5
    '
    '    Events(0).StartTime = ViewGroup.ViewDay.Date
    '    Events(0).EndTime = ViewGroup.ViewDay.Date
    'End If
    '
    'If Events.Count > 0 Then
    '    Events(0).Subject = "EVENT-0. Dynamically added/changed in PrePopulate."
    'End If
    
    Dim pEvent As CalendarEvent
    Dim strData As String
    
    For Each pEvent In Events
        
        pEvent.CustomIcons.RemoveAll
                
        If mnuCustomIcons.Checked Then
            
            ' customize standard icons
            If pEvent.PrivateFlag Then
                pEvent.CustomIcons.Add xtpCalendarEventIconIDPrivate
            End If
            
            If pEvent.Reminder Then
                pEvent.CustomIcons.Add xtpCalendarEventIconIDReminder
            End If
            
            If pEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence Then
                pEvent.CustomIcons.Add xtpCalendarEventIconIDOccurrence
            End If
            
            If pEvent.RecurrenceState = xtpCalendarRecurrenceException Then
                pEvent.CustomIcons.Add xtpCalendarEventIconIDException
            End If
        End If
    Next
    
    If mnuCustomIcons.Checked Then
        If Events.Count = 1 Then
            Events(0).CustomIcons.Add 1
            Events(0).CustomIcons.Add 4
            Events(0).CustomIcons.Add 5
                    
        ElseIf Events.Count >= 3 Then
            Events(0).CustomIcons.Add 3
            
            Events(1).CustomIcons.Add 2
            Events(1).CustomIcons.Add 6

            Events(2).CustomIcons.Add 6
            Events(2).CustomIcons.Add 6
        
        ElseIf Events.Count >= 2 Then
            Events(0).CustomIcons.Add 2
            Events(0).CustomIcons.Add 5
            
            Events(1).CustomIcons.Add 4
            Events(1).CustomIcons.Add 6
        End If
    End If
    
End Sub

Private Sub CalendarControl_PrePopulateDay(ByVal ViewDay As XtremeCalendarControl.CalendarViewDay)
    
'    If CalendarControl.MultipleResources.Count > 1 Then
'
'        Dim arRes As New CalendarResources
'        'arRes.RemoveAll
            
'        If Day(ViewDay.Date) Mod 2 = 0 Then
'            arRes.Add CalendarControl.MultipleResources.Item(0)
'            ViewDay.SetMultipleResources arRes
'        ElseIf Day(ViewDay.Date) Mod 3 = 0 Then
'            arRes.Add CalendarControl.MultipleResources.Item(CalendarControl.MultipleResources.Count - 1)
'            ViewDay.SetMultipleResources arRes
'        End If
'    End If

End Sub

Private Sub CalendarControl_SelectionChanged(ByVal SelType As XtremeCalendarControl.CalendarSelectionChanged)
    If SelType = xtpCalendarSelectionDays Then
        Debug.Print "SelectionChanged. Day(s)."
        
        If CalendarControl.ActiveView.Selection.IsValid Then
            Debug.Print CalendarControl.ActiveView.Selection.Begin
            Debug.Print CalendarControl.ActiveView.Selection.End
        End If
    End If
    If SelType = xtpCalendarSelectionEvents Then
        Debug.Print "SelectionChanged. Event(s)."
    End If
    
    UpdateToolbar
End Sub

Private Sub CalendarControl_ViewChanged()
    Dim DaysCount As Long
    DaysCount = CalendarControl.ActiveView.DaysCount
        
    Debug.Print "Number of Days: " & DaysCount
        
    If (DaysCount = 1) Then
        lblDate = Format(CalendarControl.ActiveView.Days(0).Date, "Long Date")
    ElseIf (DaysCount > 1) Then
        lblDate = Format(CalendarControl.ActiveView.Days(0).Date, "Long Date") & " - " & _
            Format(CalendarControl.ActiveView.Days(DaysCount - 1).Date, "Long Date")
    End If
    
    UpdateToolbar
End Sub

Private Sub Form_Load()
    DisableDragging_ForRecurrenceEvents = False
    DisableInPlaceCreateEvents_ForSaSu = False
    
    EnableScrollV_DayView = True
    EnableScrollH_DayView = True

    EnableScrollV_WeekView = True
    
    EnableScrollV_MonthView = True
    g_bUseBuiltInCalendarDialogs = False
    
    m_eActiveDataProvider = cjCalendarData_Unknown
    '---
    OpenProvider cjCalendarData_Memory, "Provider=XML;Data Source=" & App.Path & "\Events.xml"
    '=============================================================================
    Dim bAddRecurrenceEvent As Boolean
    
    bAddRecurrenceEvent = False
    
    If bAddRecurrenceEvent Then
        Dim NewEvent As CalendarEvent, Recurrence As CalendarRecurrencePattern
        Set NewEvent = CalendarControl.DataProvider.CreateEvent
        
        NewEvent.Subject = "RecEv"
        NewEvent.Location = "1"
        NewEvent.Body = "."
        NewEvent.ReminderSoundFile = ".."
        
        Set Recurrence = NewEvent.CreateRecurrence
        
        Recurrence.StartTime = #3:00:00 PM#
        Recurrence.DurationMinutes = 90
        
        Recurrence.StartDate = Now - 2 '#4/11/2005#
        Recurrence.EndDate = Now + 9 '#4/20/2005#
    
        Recurrence.Options.RecurrenceType = xtpCalendarRecurrenceWeekly
        Recurrence.Options.WeeklyIntervalWeeks = 1
        Recurrence.Options.WeeklyDayOfWeekMask = xtpCalendarDayMo_Fr
        NewEvent.UpdateRecurrence Recurrence
    
        CalendarControl.DataProvider.AddEvent NewEvent
    End If
    '=============================================================================
    
    Dim Today As Date
    Today = Now
    CalendarControl.ViewType = xtpCalendarDayView
    CalendarControl.DayView.ShowDays Today - 2, Today + 2
    CalendarControl.ViewType = xtpCalendarWorkWeekView
       
    CalendarControl.Populate
    CalendarControl.DayView.ScrollToWorkDayBegin

    Dim bReminders As Boolean
    bReminders = GetSetting("Codejock Calendar VB Sample", "AdvancedOptions", "RemindersManEnabled", True)
    
    CalendarControl.EnableReminders bReminders
    
    wndDatePicker.AttachToCalendar CalendarControl
        
    If g_bUseBuiltInCalendarDialogs Then
        Set g_dlgCalendarReminders.Calendar = CalendarControl
        g_dlgCalendarReminders.ParentHWND = Me.hwnd
        g_dlgCalendarReminders.RemindersWindowShowInTaskBar = False
    
        g_dlgCalendarReminders.CreateRemindersWindow
    End If
        
    mnuEnableTheme2007_Click
    
    '//***  Using Themes example
    '
    
    'mnuEnableTheme2007_Click
    
    'CalendarControl.BeforeDrawThemeObjectFlags = -1 ' set all
    
    '** customize texts ---------------------------------------------
    '
    ' can use -1 to set all flags. Like:
    'CalendarControl.AskItemTextFlags = -1
    
    ' or use separate flasg as:
    'CalendarControl.AskItemTextFlags.SetFlag xtpCalendarItemText_EventSubject
    'CalendarControl.AskItemTextFlags.SetFlag xtpCalendarItemText_EventLocation
    
   '---------------------------------------------
   ' extended recurrence example:
   '
   ' Dim dtTooday As Date
   ' dtTooday = Date
   '
   ' Dim masterEv As CalendarEvent
   ' Dim recPattern As CalendarRecurrencePattern
   '
   ' Set masterEv = CalendarControl.DataProvider.CreateEvent
   ' Set recPattern = masterEv.CreateRecurrence
   '
   ' masterEv.Subject = "Very Spesific Recurrence"
   '
   ' recPattern.StartTime = #3:00:00 PM#
   ' recPattern.DurationMinutes = 60
   ' recPattern.StartDate = dtTooday
   ' recPattern.EndAfterOccurrences = 10
   '
   ' recPattern.Options.RecurrenceType = xtpCalendarRecurrenceDaily
   ' recPattern.Options.DailyIntervalDays = 1
   ' recPattern.Options.DailyEveryWeekDayOnly = False
   '
   ' masterEv.UpdateRecurrence recPattern
   '
   ' CalendarControl.DataProvider.AddEvent masterEv
   '
   ' '**************************
   ' Dim excepEv As CalendarEvent
   ' Set excepEv = CalendarControl.DataProvider.CreateEvent
   '
   ' excepEv.StartTime = DateAdd("d", 1, dtTooday) + #3:00:00 PM#
   ' excepEv.EndTime = DateAdd("n", 60, excepEv.StartTime)
   '
   ' excepEv.MakeAsRExceptionEx recPattern.Id
   '
   ' excepEv.StartTime = DateAdd("d", 1, dtTooday) + #2:00:00 PM#
   ' excepEv.EndTime = DateAdd("n", 260, excepEv.StartTime)
   '
   ' excepEv.Subject = "Excep Event (very specific)"
   '
   ' CalendarControl.DataProvider.ChangeEvent excepEv
   '
   ' CalendarControl.Populate
   ' CalendarControl.RedrawControl
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    CalendarControl.DataProvider.Save
    CalendarControl.DataProvider.Close
    
    Dim bReminders As Boolean
    bReminders = CalendarControl.IsRemindersEnabled
    SaveSetting "Codejock Calendar VB Sample", "AdvancedOptions", "RemindersEnabled", bReminders
    
    Unload frmReminders
    Unload frmTheme2007
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim nHeight As Long, LabelWidth As Long
    
    nHeight = Height - StatusBar.Height * 3 - CalendarControl.Top - 12 * Screen.TwipsPerPixelY
    
    LabelWidth = Me.ScaleWidth - lblCaption.Width - 100
    
    If nHeight < 0 Then Height = 0
    If LabelWidth < 0 Then LabelWidth = 0

    
    wndDatePicker.Left = Me.ScaleWidth - wndDatePicker.Width - 3
    wndDatePicker.Height = nHeight
        
    'wndDatePicker.Move Me.ScaleWidth - wndDatePicker.Width, CalendarControl.Top, wndDatePicker.Left, nHeight
    'CalendarControl.Move 0, CalendarControl.Top, Me.ScaleWidth, nHeight
    CalendarControl.Move 0, CalendarControl.Top, wndDatePicker.Left, nHeight
        
    lblDate.Move lblCaption.Width, lblDate.Top, LabelWidth, lblDate.Height
End Sub


Private Sub lblCaption_Click()
    Dim Today As Date
    Today = Now
    CalendarControl.ActiveView.ShowDay Today + 1, False
End Sub

Private Sub mnuAdvancedOptions_Click()
    frmAdvancedOptions.Show vbModal, Me
End Sub

Private Sub mnuCalendar_Click()
    mnuReminders.Enabled = CalendarControl.IsRemindersEnabled
    
    Dim objThemeOfice2007 As CalendarThemeOffice2007
    Set objThemeOfice2007 = CalendarControl.Theme
    
    mnuEnableTheme2007.Checked = Not objThemeOfice2007 Is Nothing
    mnuCustomizeTheme2007.Enabled = Not objThemeOfice2007 Is Nothing
    mnuCustomIcons.Enabled = mnuCustomizeTheme2007.Enabled
    mnuShowDynamicCustomization.Enabled = mnuCustomizeTheme2007.Enabled
    
End Sub

Private Sub mnuChangeTimeZone_Click()

    If g_bUseBuiltInCalendarDialogs Then
        
        Dim dlgCalendar As New CalendarDialogs
        
        dlgCalendar.ParentHWND = Me.hwnd
        dlgCalendar.Calendar = CalendarControl
        
        dlgCalendar.ShowTimeScaleProperties
        
        Exit Sub
        
    End If
    
    frmTimeZone.Show vbModal, Me
End Sub

Private Sub mnuCustomIcons_Click()
    Dim objThemeOfice2007 As CalendarThemeOffice2007
    Set objThemeOfice2007 = CalendarControl.Theme
    
    If objThemeOfice2007 Is Nothing Then Exit Sub
        
    If Not mnuCustomIcons.Checked Then
        mnuCustomIcons.Checked = True
        
        ' add custom icons with special IDs to use them instead of standard
        ' see also PrePopulate event handler
        
        objThemeOfice2007.CustomIcons.LoadBitmap App.Path & "\Icons\Reminder.bmp", xtpCalendarEventIconIDReminder, xtpImageNormal
        objThemeOfice2007.CustomIcons.LoadBitmap App.Path & "\Icons\Occ.bmp", xtpCalendarEventIconIDOccurrence, xtpImageNormal
        objThemeOfice2007.CustomIcons.LoadBitmap App.Path & "\Icons\Exc.bmp", xtpCalendarEventIconIDException, xtpImageNormal
        objThemeOfice2007.CustomIcons.LoadBitmap App.Path & "\Icons\Private.bmp", xtpCalendarEventIconIDPrivate, xtpImageNormal
        
        'objThemeOfice2007.CustomIcons.AddIcon ImageListCustomIcons.ListImages.Item(1).ExtractIcon.Handle, xtpCalendarEventIconIDReminder, xtpImageNormal
        'objThemeOfice2007.CustomIcons.AddIcon ImageListCustomIcons.ListImages.Item(2).ExtractIcon.Handle , xtpCalendarEventIconIDOccurrence, xtpImageNormal
        'objThemeOfice2007.CustomIcons.AddIcon ImageListCustomIcons.ListImages.Item(3).ExtractIcon.Handle, xtpCalendarEventIconIDException, xtpImageNormal
        'objThemeOfice2007.CustomIcons.AddIcon ImageListCustomIcons.ListImages.Item(4).ExtractIcon.Handle, xtpCalendarEventIconIDPrivate, xtpImageNormal
                        
        '' increase event height for 4 pixels to have enough space to draw custom icons.
        'objThemeOfice2007.MonthView.Event.HeightFormula.Constant = 5
        'objThemeOfice2007.WeekView.Event.HeightFormula.Constant = 5
        'objThemeOfice2007.RefreshMetrics
        
        ' custom icons
        
        'objThemeOfice2007.SetCustomIcons ImageListCustomIcons
        
        Dim arCustIconsIDs(5) As Long
        arCustIconsIDs(0) = 1 ' unread mail
        arCustIconsIDs(1) = 2 ' read mail
        arCustIconsIDs(2) = 3 ' replyed mail
        arCustIconsIDs(3) = 4 ' attachment
        arCustIconsIDs(4) = 5 ' Low priority
        arCustIconsIDs(5) = 6 ' HIGH priority
        
        objThemeOfice2007.CustomIcons.LoadBitmap App.Path & "\Icons\EventCustomIcons.bmp", arCustIconsIDs, xtpImageNormal
    Else
        mnuCustomIcons.Checked = False
        
        objThemeOfice2007.CustomIcons.RemoveAll
    End If
    
    CalendarControl.Populate
    
End Sub

Private Sub mnuCustomizeTheme2007_Click()
    frmTheme2007.Show vbModeless, Me
    
    
End Sub

Private Sub mnuDeleteEvent_Click()
    Dim bDeleted As Boolean
    bDeleted = False
    
    If ContextEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence _
        Or ContextEvent.RecurrenceState = xtpCalendarRecurrenceException _
    Then
        frmOccurrenceSeriesChooser.m_bOcurrence = True
        frmOccurrenceSeriesChooser.m_bDeleteRequest = True
        frmOccurrenceSeriesChooser.m_strEventSubject = ContextEvent.Subject
        
        frmOccurrenceSeriesChooser.Show vbModal
        
        If frmOccurrenceSeriesChooser.m_bOK = False Then
            Exit Sub
        ElseIf Not frmOccurrenceSeriesChooser.m_bOcurrence Then
            ' Series
            CalendarControl.DataProvider.DeleteEvent ContextEvent.RecurrencePattern.MasterEvent
            bDeleted = True
        End If
    End If
        
    If Not bDeleted Then
        CalendarControl.DataProvider.DeleteEvent ContextEvent
    End If
    
    CalendarControl.Populate
    'CalendarControl.RedrawControl
End Sub

Private Sub mnuEnableTheme2007_Click()
    Dim objThemeOfice2007 As CalendarThemeOffice2007
    Set objThemeOfice2007 = CalendarControl.Theme
    
    If objThemeOfice2007 Is Nothing Then
        
        ' Create and set Theme
        Set objThemeOfice2007 = New CalendarThemeOffice2007
        CalendarControl.SetTheme objThemeOfice2007
                        
        ' Load customized theme options
        Dim px As PropExchange
        Set px = XtremeCalendarControl.CreatePropExchange
        If px.CreateAsXML(True, "CalendarThemeOffice2007") Then
            If px.LoadFromFile(App.Path & "\cfgCalendarThemeOffice2007.xml") Then
                frmMain.CalendarControl.Theme.DoPropExchange px
            End If
        End If
        
        ' Set theme for DatePicker
        
        Dim objDPTheme2007 As New DatePickerThemeOffice2007
        wndDatePicker.SetTheme objDPTheme2007
    Else
        Unload frmTheme2007
        
        If mnuCustomIcons.Checked Then
            mnuCustomIcons_Click
        End If
        CalendarControl.SetTheme Nothing
        wndDatePicker.SetTheme Nothing
    End If
    
    CalendarControl.Populate
    wndDatePicker.RedrawControl
        
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuMultiSchedulesExtended_Click()
    
    '// Prepare resource manager
    Dim strCfgFile As String
    strCfgFile = App.Path & "\" & "CalendarMultipleSchedulesSample.xml"

    ' try to load previously saved configuration
    'g_DataResourcesMan.LoadCfg strCfgFile
    
    '// Setup sample configuration
    '// ** data provider
    Dim strConnectionString As String
    strConnectionString = "Provider=XML;Data Source=" & App.Path & "\" & "CalendarMultipleSchedulesExt.xtp_cal"
    
    Dim bResult As Boolean
    
    bResult = g_DataResourcesMan.AddDataProvider( _
            strConnectionString, xtpCalendarDPF_CreateIfNotExists + _
            xtpCalendarDPF_SaveOnDestroy + xtpCalendarDPF_CloseOnDestroy)
    
    If Not bResult Then
        Exit Sub
    End If
        
    Dim pData As CalendarDataProvider
    Set pData = g_DataResourcesMan.DataProvider(0)
    
    '// ** schedules
    Dim pSchedules As CalendarSchedules
    Set pSchedules = pData.Schedules
    
    If pSchedules Is Nothing Then
        Exit Sub
    End If
    
    If pSchedules.Count < 1 Then
        pSchedules.AddNewSchedule "John"
        pSchedules.AddNewSchedule "Peter"
        pSchedules.AddNewSchedule "Room N1"
        pSchedules.AddNewSchedule "Room N2"
        
        pData.Save
    End If
    
    '// ** resources
    If g_DataResourcesMan.ResourcesCount = 0 Then
        g_DataResourcesMan.AddResource "John", True
        g_DataResourcesMan.AddResource "Peter", True
        g_DataResourcesMan.AddResource "Rooms", True

        Dim pRCDesc As CalendarResourceDescription

        Set pRCDesc = g_DataResourcesMan.Resource(0)
        pRCDesc.Resource.SetDataProvider pData, False
        pRCDesc.Resource.ScheduleIDs.Add pSchedules(0).Id
        pRCDesc.GenerateName = True

        Set pRCDesc = g_DataResourcesMan.Resource(1)
        pRCDesc.Resource.SetDataProvider pData, False
        pRCDesc.Resource.ScheduleIDs.Add pSchedules(1).Id
        pRCDesc.GenerateName = True

        Set pRCDesc = g_DataResourcesMan.Resource(2)
        pRCDesc.Resource.SetDataProvider pData, False
        pRCDesc.Resource.ScheduleIDs.Add pSchedules(2).Id
        pRCDesc.Resource.ScheduleIDs.Add pSchedules(3).Id
        pRCDesc.GenerateName = False

        '// Save changed resources configuration
        g_DataResourcesMan.SaveCfg strCfgFile
    End If
    
    '// Apply resources configuration
    g_DataResourcesMan.ApplyToCalendar CalendarControl
    
    CalendarControl.Populate
    CalendarControl.RedrawControl
End Sub

Private Sub mnuMultiSchedulesSimple_Click()

    '// Setup sample configuration
    '// ** data provider
    Dim strConnectionString As String
    strConnectionString = "Provider=XML;Data Source=" & App.Path & "\" & "CalendarMultipleSchedulesSample.xtp_cal"
    
    Dim arResources As New CalendarResources
    Dim pRes0 As New CalendarResource
    Dim pRes1 As New CalendarResource
    
    pRes0.SetDataProvider2 strConnectionString, True
      
    If pRes0.DataProvider Is Nothing Then
        Debug.Assert False
        Exit Sub
    End If
    
    If Not pRes0.DataProvider.Open Then
        If Not pRes0.DataProvider.Create Then
            Debug.Assert False
            Exit Sub
        End If
    End If
    
    '// ** schedules
    Dim pSchedules As CalendarSchedules
    Set pSchedules = pRes0.DataProvider.Schedules
    
    If pSchedules Is Nothing Then
        Exit Sub
    End If
    
    If pSchedules.Count < 1 Then
        pSchedules.AddNewSchedule "John"
        pSchedules.AddNewSchedule "Jane"
        
        pRes0.DataProvider.Save
    End If
    
    '// ** resources
    
    pRes0.Name = pSchedules.Item(0).Name
    pRes0.ScheduleIDs.Add pSchedules.Item(0).Id
    
    pRes1.SetDataProvider pRes0.DataProvider, False
    pRes1.Name = "Jane (+8 hr)" 'pSchedules.Item(1).Name
    pRes1.ScheduleIDs.Add pSchedules.Item(1).Id
      
    arResources.Add pRes0
    arResources.Add pRes1
        
    CalendarControl.SetMultipleResources arResources
    
    CalendarControl.Populate
    CalendarControl.RedrawControl
    
End Sub

Private Sub mnuNewEvent2_Click()
    mnuNewEvent_Click
End Sub
Private Sub mnuNewEvent_Click()
    If g_bUseBuiltInCalendarDialogs Then
        Dim dlgCalendar As New CalendarDialogs
        dlgCalendar.ParentHWND = Me.hwnd
        dlgCalendar.Calendar = CalendarControl
        
        dlgCalendar.ShowNewEvent
        Exit Sub
    End If
    
    frmEditEvent.NewEvent
    frmEditEvent.Show vbModal, Me
End Sub


Private Sub mnuOpenDataProvider_Click()
    frmCalendarDataChooser.Show vbModal, frmMain
    
    If Not frmCalendarDataChooser.Cancelled Then
        Screen.MousePointer = vbHourglass
        frmMain.OpenProvider frmCalendarDataChooser.ProviderType, frmCalendarDataChooser.ConnectionString
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuOpenEvent_Click()
    ModifyEvent ContextEvent
End Sub

Private Sub mnuPageSetup_Click()
    CalendarControl.ShowPrintPageSetup
End Sub

Private Sub mnuPrintCalendar_Click()
    ' Header ---------------------------------------------------
    
    'CalendarControl.PrintOptions.Header.Font.Name = "Courier New"
    'CalendarControl.PrintOptions.Header.Font.Bold = True
    'CalendarControl.PrintOptions.Header.Font.SIZE = 10
        
    CalendarControl.PrintOptions.Header.TextLeft = "(c)1998-2007 Codejock Software, All Rights Reserved."
    CalendarControl.PrintOptions.Header.TextCenter = "Calendar Control"
    CalendarControl.PrintOptions.Header.TextRight = "Page 1 of 1 "
    
    ' Footer ---------------------------------------------------
    
    'CalendarControl.PrintOptions.Header.Font.Italic = True
    
    CalendarControl.PrintOptions.Footer.TextLeft = "Date: " & DateValue(Now)
    CalendarControl.PrintOptions.Footer.TextCenter = "Codejock Software, " & vbLf & " Print calendar example "
    CalendarControl.PrintOptions.Footer.TextRight = "Time: " & TimeValue(Now)
    
    ' Other fonts
    
    'CalendarControl.PrintOptions.DateHeaderFont.SIZE = 12
    'CalendarControl.PrintOptions.DateHeaderCalendarFont.SIZE = 10
            
    'If CalendarControl.ShowPrintPageSetup() Then
        CalendarControl.PrintCalendar 0
    'End If
End Sub

Private Sub mnuPrintPreview_Click()
    CalendarControl.PrintPreviewOptions.Title = "Calendar Control VB 6.0 Sample application"
    CalendarControl.PrintPreview True
End Sub

Private Sub mnuReminders_Click()
    If g_bUseBuiltInCalendarDialogs Then
        g_dlgCalendarReminders.ShowRemindersWindow
        Exit Sub
    End If

    If ModalFormsRunningCounter = 0 Then
        If Not frmReminders.Visible Then
            frmReminders.OnReminders xtpCalendarRemindersFire, Nothing
        End If
        
        frmReminders.Show vbModeless, Me
        'frmReminders.Visible = True
    End If
End Sub

Private Sub mnuResourcesManager_Click()
    If ModalFormsRunningCounter = 0 Then
        Dim strCFG As String
        strCFG = App.Path & "\" & "CalMulSchSampleVB.xml"
    
        frmResourcesManager.Show vbModal, frmMain
        If Not frmResourcesManager.Cancelled Then
            g_DataResourcesMan.SaveCfg strCFG
            g_DataResourcesMan.ApplyToCalendar CalendarControl
            CalendarControl.Populate
            CalendarControl.RedrawControl
        End If
    End If
End Sub

Private Sub mnuSettings_Click()
    frmSettings.Show vbModal, Me
End Sub

Private Sub mnuShowDynamicCustomization_Click()
    
    Dim pTheme2007 As CalendarThemeOffice2007
    Set pTheme2007 = CalendarControl.Theme
        
    If CalendarControl.BeforeDrawThemeObjectFlags <> 0 Then
        CalendarControl.BeforeDrawThemeObjectFlags = 0 ' reset all flags
        
        pTheme2007.DayView.TimeScale.ShowMinutes = CalendarControl.Options.DayViewTimeScaleShowMinutes
    Else
        CalendarControl.BeforeDrawThemeObjectFlags = -1 ' sett all flags
        
        pTheme2007.DayView.TimeScale.ShowMinutes = True
                        
        'to improve performance set only flags which you need
        '
        'CalendarControl.BeforeDrawThemeObjectFlags.SetFlag xtpCalendarBeforeDraw_DayViewTimeScale
        'CalendarControl.BeforeDrawThemeObjectFlags.SetFlag xtpCalendarBeforeDraw_DayViewTimeScaleCell
        'CalendarControl.BeforeDrawThemeObjectFlags.SetFlag xtpCalendarBeforeDraw_DayViewTimeScaleCaption
        'CalendarControl.BeforeDrawThemeObjectFlags.SetFlag xtpCalendarBeforeDraw_DayViewDay
        'CalendarControl.BeforeDrawThemeObjectFlags.SetFlag xtpCalendarBeforeDraw_MonthViewDay
    
   End If
   
   mnuShowDynamicCustomization.Checked = CalendarControl.BeforeDrawThemeObjectFlags <> 0
   
   CalendarControl.RedrawControl
End Sub

Private Sub mnuTimeScale_Click(Index As Integer)
    CalendarControl.DayView.TimeScale = mnuTimeScale(Index).HelpContextID
    
End Sub


Private Sub ModifyEvent(ModEvent As CalendarEvent)
        
    If g_bUseBuiltInCalendarDialogs Then
        Dim dlgCalendar As New CalendarDialogs
        dlgCalendar.ParentHWND = Me.hwnd
        dlgCalendar.Calendar = CalendarControl
        
        dlgCalendar.ShowEditEvent ModEvent
        
        Exit Sub
    End If
        
    If ModEvent.RecurrenceState <> xtpCalendarRecurrenceNotRecurring Then
        
        frmOccurrenceSeriesChooser.m_bOcurrence = True
        frmOccurrenceSeriesChooser.m_bDeleteRequest = False
        frmOccurrenceSeriesChooser.m_strEventSubject = ModEvent.Subject
        
        frmOccurrenceSeriesChooser.Show vbModal
        
        If frmOccurrenceSeriesChooser.m_bOK = False Then
            Exit Sub
        ElseIf frmOccurrenceSeriesChooser.m_bOcurrence Then
            If ModEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence Then
                Set ModEvent = ModEvent.CloneEvent
                ModEvent.MakeAsRException
            End If
        Else
            ' Series
            Set ModEvent = ModEvent.RecurrencePattern.MasterEvent
        End If
    End If
    

    frmEditEvent.ModifyEvent ModEvent
    frmEditEvent.Show vbModal, Me
End Sub

Function GetMonday(dtDate As Date) As Date
    Dim Day As Long
    Day = Weekday(dtDate, vbMonday)
    If (Day < 5) Then
        GetMonday = dtDate - Day + 1
    Else
        GetMonday = dtDate - 2
    End If
        
End Function


Private Sub timerRMDForm_Timer()
    If ModalFormsRunningCounter = 0 Then
        timerRMDForm.Enabled = False
        frmReminders.Show vbModeless, Me
        'frmReminders.Visible = True
    End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim BeginSelection As Date, EndSelection As Date, AllDay As Boolean

    Select Case (Button.Index)
        Case 1:
            CalendarControl.ViewType = xtpCalendarDayView
        Case 2:
            CalendarControl.ViewType = xtpCalendarWorkWeekView
        Case 3:
            CalendarControl.ViewType = xtpCalendarWeekView
        Case 4:
            CalendarControl.ViewType = xtpCalendarMonthView
        
        Case 6:
            CalendarControl.ActiveView.Cut
        Case 7:
            CalendarControl.ActiveView.Copy
        Case 8:
            CalendarControl.ActiveView.Paste
        
        Case 12:
            mnuOpenDataProvider_Click
            
        Case 14:
            mnuPageSetup_Click
        Case 15:
            mnuPrintPreview_Click
        Case 16:
            mnuPrintCalendar_Click
    End Select
    
    UpdateToolbar

    'CalendarControl.ActiveView.DayHeaderFormatLong = "'(1)' ddd, dd MMMM yyyy"
    'CalendarControl.ActiveView.DayHeaderFormatMiddle = "'(2)' ddd, dd MMM yy"
    'CalendarControl.ActiveView.DayHeaderFormatShort = "'(3)' dd MMM "
    'CalendarControl.ActiveView.DayHeaderFormatShortest = "'(4)' d.MM"
    'CalendarControl.Populate
End Sub

Private Sub UpdateToolbar()
    
    Toolbar.Buttons(6).Enabled = CalendarControl.ActiveView.CanCut
    Toolbar.Buttons(7).Enabled = CalendarControl.ActiveView.CanCopy
    Toolbar.Buttons(8).Enabled = CalendarControl.ActiveView.CanPaste
    
End Sub
