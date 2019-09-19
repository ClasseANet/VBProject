VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#11.2#0"; "Codejock.Calendar.v11.2.2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmCalendario 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "CalendarControl Sample"
   ClientHeight    =   6300
   ClientLeft      =   4965
   ClientTop       =   750
   ClientWidth     =   9195
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerCal 
      Interval        =   1000
      Left            =   3720
      Top             =   5760
   End
   Begin VB.Timer timerRMDForm 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1500
      Top             =   -60
   End
   Begin XtremeCalendarControl.CalendarControl CalendarControl 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8835
      _Version        =   720898
      _ExtentX        =   15584
      _ExtentY        =   7329
      _StockProps     =   64
      ViewType        =   1
   End
   Begin XtremeSuiteControls.TabControl TabSchedule 
      Height          =   360
      Left            =   5760
      TabIndex        =   2
      Top             =   180
      Width           =   3015
      _Version        =   720898
      _ExtentX        =   5318
      _ExtentY        =   635
      _StockProps     =   68
      AllowReorder    =   -1  'True
      Appearance      =   2
      Color           =   8
      PaintManager.BoldSelected=   -1  'True
      PaintManager.MinTabWidth=   60
      ItemCount       =   3
      Item(0).Caption =   "Sala 01"
      Item(0).ControlCount=   0
      Item(1).Caption =   "Sala 02"
      Item(1).ControlCount=   0
      Item(2).Caption =   "Todos"
      Item(2).ControlCount=   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   1200
      Top             =   5040
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageEvento 
      Left            =   8400
      Top             =   4680
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmCalendario.frx":0000
   End
   Begin XtremeShortcutBar.ShortcutCaption SccCalendario 
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8760
      _Version        =   720898
      _ExtentX        =   15452
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "Calendario"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin ComctlLib.ImageList ilToolBar0 
      Left            =   7800
      Top             =   4440
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
            Picture         =   "FrmCalendario.frx":275A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":2CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":31FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":3750
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":3CA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":3DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":3EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":3FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":40CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":42A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCalendario.frx":447E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event Resize()
Event Unload(Cancel As Integer)
Event CalendarControlBeforeDrawDayViewCell(ByVal CellParams As XtremeCalendarControl.CalendarDayViewCellParams)
Event CalendarControlBeforeEditOperation(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, CancelOperation As Boolean)
Event CalendarControlContextMenu(ByVal x As Single, ByVal y As Single)
Event CalendarControlDblClick()
Event CalendarControlEventChanged(ByVal EventID As Long)
Event CalendarControlKeyDown(KeyCode As Integer, Shift As Integer)
Event CalendarControlMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event CalendarControlOnReminders(ByVal Action As XtremeCalendarControl.CalendarRemindersAction, ByVal Reminder As XtremeCalendarControl.CalendarReminder)
Event CalendarControlPrePopulate(ByVal ViewGroup As XtremeCalendarControl.CalendarViewGroup, ByVal Events As XtremeCalendarControl.CalendarEvents)
Event CalendarControlPrePopulateDay(ByVal ViewDay As XtremeCalendarControl.CalendarViewDay)
Event CalendarControlSelectionChanged(ByVal SelType As XtremeCalendarControl.CalendarSelectionChanged)
Event CalendarControlViewChanged()
Event TabScheduleSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Event TimerCalTimer()
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
   Me.CalendarControl.EnableToolTips (nMode < 2)
   
    If nMode = 1 Then
        Me.CalendarControl.AskItemTextFlags.SetFlag xtpCalendarItemText_EventToolTipText
    Else
        Me.CalendarControl.AskItemTextFlags.ResetFlag xtpCalendarItemText_EventToolTipText
    End If
   
End Property
Private Sub CalendarControl_BeforeDrawDayViewCell(ByVal CellParams As XtremeCalendarControl.CalendarDayViewCellParams)
   RaiseEvent CalendarControlBeforeDrawDayViewCell(CellParams)
End Sub

Private Sub CalendarControl_BeforeDrawThemeObject(ByVal eObjType As XtremeCalendarControl.CalendarBeforeDrawThemeObject, ByVal Params As Variant)
Exit Sub
    Dim pTheme2007 As CalendarThemeOffice2007
    Set pTheme2007 = Me.CalendarControl.Theme
    
    If eObjType = xtpCalendarBeforeDraw_DayViewDay Then
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
        'me.CalendarControl.Theme.DayView.TimeScale.TimeTextBigBase.Font.Strikethrough = (Second(Now) Mod 2) = 0
        Dim nSubCellIndex
        nSubCellIndex = Params.Minutes / Me.CalendarControl.DayView.TimeScale
        
        If Params.Minutes = -1 Then
            pTheme2007.DayView.TimeScale.TimeTextBigBase.Font.Bold = (Params.Index Mod 2) = 0
            pTheme2007.DayView.TimeScale.TimeTextSmall.Font.Bold = (Params.Index Mod 2) = 0
            pTheme2007.DayView.TimeScale.TimeTextSmall.Color = 20000 + 1111 * Params.Index
        Else
            If (nSubCellIndex Mod 2) = 0 And Me.CalendarControl.DayView.TimeScale < 15 Then
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
        
        If Me.CalendarControl.DayView.TimeScale > 15 Then
            Me.CalendarControl.Theme.DayView.TimeScale.TimeTextBigBase.Color = 20000 + 100 * Params.Index
        End If
        Me.CalendarControl.Theme.DayView.TimeScale.LineColor = Me.CalendarControl.Theme.DayView.TimeScale.TimeTextBigBase.Color
        
    ElseIf eObjType = xtpCalendarBeforeDraw_DayViewTimeScaleCaption Then
        Me.CalendarControl.Theme.DayView.TimeScale.Caption.Font.Italic = (Second(Now) Mod 2) = 0
        Me.CalendarControl.Theme.DayView.TimeScale.Caption.Color = RGB(10, 200, 100)
        Me.CalendarControl.Theme.DayView.TimeScale.LineColor = RGB(10, 200, 100)
        
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
   RaiseEvent CalendarControlBeforeEditOperation(OpParams, CancelOperation)
End Sub
Private Sub CalendarControl_ContextMenu(ByVal x As Single, ByVal y As Single)
   RaiseEvent CalendarControlContextMenu(x, y)
End Sub
Private Sub CalendarControl_DblClick()
   RaiseEvent CalendarControlDblClick
End Sub
Private Sub CalendarControl_EventChanged(ByVal EventID As Long)
   RaiseEvent CalendarControlEventChanged(EventID)
End Sub
Private Sub CalendarControl_GetItemText(ByVal Params As XtremeCalendarControl.CalendarGetItemTextParams)
   Select Case Params.Item
      Case xtpCalendarItemText_EventSubject:             Params.Text = "*Custom subject* " & Params.Text
      Case xtpCalendarItemText_EventLocation:            Params.Text = "ScheduleID = " & Params.ViewEvent.Event.ScheduleID
      Case xtpCalendarItemText_EventBody:                Params.Text = "custom BODY text."
      Case xtpCalendarItemText_DayViewDayHeader:         Params.Text = "Date:" & Params.ViewDay.Date
      Case xtpCalendarItemText_WeekViewDayHeader:        Params.Text = "Date:" & Params.ViewDay.Date
      Case xtpCalendarItemText_MonthViewDayHeader:       Params.Text = "Date:" & Params.ViewDay.Date
      Case xtpCalendarItemText_DayViewDayHeaderLeft:     Params.Text = Day(Params.ViewDay.Date)
      Case xtpCalendarItemText_WeekViewDayHeaderLeft:    Params.Text = Day(Params.ViewDay.Date)
      Case xtpCalendarItemText_MonthViewDayHeaderLeft:   Params.Text = Day(Params.ViewDay.Date)
      Case xtpCalendarItemText_DayViewDayHeaderCenter:   Params.Text = WeekdayName(Weekday(Params.ViewDay.Date), , vbSunday)
      Case xtpCalendarItemText_WeekViewDayHeaderCenter:  Params.Text = WeekdayName(Weekday(Params.ViewDay.Date), , vbSunday)
      Case xtpCalendarItemText_MonthViewDayHeaderCenter: Params.Text = WeekdayName(Weekday(Params.ViewDay.Date), , vbSunday)
      Case xtpCalendarItemText_DayViewDayHeaderRight:    Params.Text = Year(Params.ViewDay.Date)
      Case xtpCalendarItemText_WeekViewDayHeaderRight:   Params.Text = Year(Params.ViewDay.Date)
      Case xtpCalendarItemText_MonthViewDayHeaderRight:  Params.Text = Year(Params.ViewDay.Date)
      Case xtpCalendarItemText_MonthViewWeekDayHeader:   Params.Text = Params.Text & " - " & Params.Weekday
      Case xtpCalendarItemText_EventToolTipText
         Params.Text = "ID = [" & Params.ViewEvent.Event.Id & "]  " & vbCrLf
         Params.Text = Params.Text & Params.ViewEvent.Event.Subject & vbCrLf
         Params.Text = Params.Text & Params.ViewEvent.Event.Location & vbCrLf
         'Params.Text = Params.Text & Params.ViewEvent.Event.Body
    End Select
End Sub
Private Sub CalendarControl_GotFocus()
'    Debug.Print "GotFocus"
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
                                   
            If Me.CalendarControl.ActiveView.GetSelection(dtBegin, dtEnd, bAllDay) = False Then
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
   RaiseEvent CalendarControlKeyDown(KeyCode, Shift)
End Sub
Private Sub CalendarControl_LostFocus()
   'Debug.Print "LostFocus"
End Sub
Private Sub CalendarControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent CalendarControlMouseDown(Button, Shift, x, y)
End Sub
Private Sub CalendarControl_OnReminders(ByVal Action As XtremeCalendarControl.CalendarRemindersAction, ByVal Reminder As XtremeCalendarControl.CalendarReminder)
   RaiseEvent CalendarControlOnReminders(Action, Reminder)
End Sub
Private Sub CalendarControl_PrePopulate(ByVal ViewGroup As XtremeCalendarControl.CalendarViewGroup, ByVal Events As XtremeCalendarControl.CalendarEvents)
   RaiseEvent CalendarControlPrePopulate(ViewGroup, Events)
End Sub
Private Sub CalendarControl_PrePopulateDay(ByVal ViewDay As XtremeCalendarControl.CalendarViewDay)
    RaiseEvent CalendarControlPrePopulateDay(ViewDay)
End Sub
Private Sub CalendarControl_SelectionChanged(ByVal SelType As XtremeCalendarControl.CalendarSelectionChanged)
   RaiseEvent CalendarControlSelectionChanged(SelType)
End Sub
Private Sub CalendarControl_ViewChanged()
   RaiseEvent CalendarControlViewChanged
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub mnuSettings_Click()
    FrmOpcao.Show vbModal, Me
End Sub
Private Sub SccCalendario_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim Today As Date
   Dim nType As Integer
    
   With Me.CalendarControl
      nType = .ViewType
      Today = Now
      .ActiveView.ShowDay Today, True
      .ViewType = nType
   End With
    
    'Dim pTheme2007 As CalendarThemeOffice2007
    'Set pTheme2007 = Me.CalendarControl.Theme
    'pTheme2007.DayView.TimeScale.HeightFormula.Multiplier = 17
    'pTheme2007.DayView.TimeScale.HeightFormula.Divisor = 13
    'pTheme2007.DayView.TimeScale.HeightFormula.Constant = 5
    
    'pTheme2007.DayView.Day.Header.HeightFormula.Multiplier = 14
    'pTheme2007.RefreshMetrics
    'Me.CalendarControl.DayView.Days
    'Me.CalendarControl.Populate
End Sub
Private Sub TabSchedule_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   RaiseEvent TabScheduleSelectedChanged(Item)
End Sub
Private Sub TimerCal_Timer()
   RaiseEvent TimerCalTimer
End Sub
Private Sub timerRMDForm_Timer()
    If ModalFormsRunningCounter = 0 Then
        timerRMDForm.Enabled = False
        FrmLembrete.Show vbModeless, Me
        'FrmLembrete.Visible = True
    End If
End Sub


