VERSION 5.00
Begin VB.Form frmAdvancedOptions 
   Caption         =   "Advanced Calendar Control options"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Current Time Mark (Day View)"
      Height          =   675
      Left            =   120
      TabIndex        =   23
      Top             =   3180
      Width           =   4815
      Begin VB.OptionButton optTimeMarkNone 
         Caption         =   "None"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox chkTimeMarkPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optTimeMarkAlways 
         Caption         =   "Always"
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optTimeMarkForToday 
         Caption         =   "For Today only"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   300
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   5040
      TabIndex        =   21
      Top             =   2760
      Width           =   2895
      Begin VB.CheckBox chkEnableReminders 
         Caption         =   "Enable Reminders"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   4815
      Begin VB.CheckBox chkShowMinutes 
         Caption         =   "Show minutes on TimeScale (DayView)"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   4215
      End
      Begin VB.CheckBox chkUseOutlookFontGlyphs 
         Caption         =   "Use ""MS Outlook"" font glyphs"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Scrolling possibilities"
      Height          =   2535
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox chkScrollingV_MonthView 
         Caption         =   "Enable Vertical scrolling"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CheckBox chkScrollingV_WeekView 
         Caption         =   "Enable Vertical scrolling"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkScrollingH_DayView 
         Caption         =   "Enable Horizontal scrolling"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox chkScrollingV_DayView 
         Caption         =   "Enable Vertical scrolling"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month View:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Week View:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Day View:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Customization demo options"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   7815
      Begin VB.CheckBox chkUseBuiltInDialogs 
         Caption         =   "Use built-in calendar dialogs"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   1260
         Width           =   2835
      End
      Begin VB.CheckBox chkDisableInPlaceCreateEvents_ForSaSu 
         Caption         =   "Disable in-place event creation for Sundays and Saturdays"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   4575
      End
      Begin VB.CheckBox chkDisableDragging_ForRecurrenceEvents 
         Caption         =   "Disable Dragging operations for Recurrence events "
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Standard Event editing options"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CheckBox chkInplaceCreateEvent 
         Caption         =   "Enable in-place Create Event"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CheckBox chkEditSubject_ByTab 
         Caption         =   "Enable in-place edit event subject - by TAB"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   4215
      End
      Begin VB.CheckBox chkEditSubject_ByF2 
         Caption         =   "Enable in-place edit event subject - by F2"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   4215
      End
      Begin VB.CheckBox chkEditSubject_AfterResize 
         Caption         =   "Enable in-place edit event subject - After Event Resize"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4335
      End
      Begin VB.CheckBox chkEditSubject_ByMouseClick 
         Caption         =   "Enable in-place edit event subject - by Mouse Click"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7920
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frmAdvancedOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Property Get CalendarControl() As CalendarControl
    Set CalendarControl = frmMain.CalendarControl
End Property

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    CalendarControl.Options.EnableInPlaceEditEventSubject_ByMouseClick = BinToBoolean(chkEditSubject_ByMouseClick.Value)
    CalendarControl.Options.EnableInPlaceEditEventSubject_AfterEventResize = BinToBoolean(chkEditSubject_AfterResize.Value)
    
    CalendarControl.Options.EnableInPlaceEditEventSubject_ByF2 = BinToBoolean(chkEditSubject_ByF2.Value)
    CalendarControl.Options.EnableInPlaceEditEventSubject_ByTab = BinToBoolean(chkEditSubject_ByTab.Value)
    
    CalendarControl.Options.EnableInPlaceCreateEvent = BinToBoolean(chkInplaceCreateEvent.Value)
    CalendarControl.Options.UseOutlookFontGlyphs = BinToBoolean(chkUseOutlookFontGlyphs.Value)
    CalendarControl.Options.DayViewTimeScaleShowMinutes = BinToBoolean(chkShowMinutes.Value)
    
        
    If optTimeMarkNone.Value Then
        CalendarControl.Options.DayViewCurrentTimeMarkVisible = xtpCalendarCurrentTimeMarkNone
    Else
        If optTimeMarkForToday.Value Then
            CalendarControl.Options.DayViewCurrentTimeMarkVisible = xtpCalendarCurrentTimeMarkVisibleForToday
        ElseIf optTimeMarkAlways.Value Then
            CalendarControl.Options.DayViewCurrentTimeMarkVisible = xtpCalendarCurrentTimeMarkVisibleAlways
        End If
        
        If chkTimeMarkPrint.Value <> 0 Then
            CalendarControl.Options.DayViewCurrentTimeMarkVisible = CalendarControl.Options.DayViewCurrentTimeMarkVisible + xtpCalendarCurrentTimeMarkPrinted
        End If
    End If
        
    frmMain.DisableDragging_ForRecurrenceEvents = BinToBoolean(chkDisableDragging_ForRecurrenceEvents.Value)
    frmMain.DisableInPlaceCreateEvents_ForSaSu = BinToBoolean(chkDisableInPlaceCreateEvents_ForSaSu.Value)
    
    frmMain.EnableScrollV_DayView = BinToBoolean(chkScrollingV_DayView.Value)
    frmMain.EnableScrollH_DayView = BinToBoolean(chkScrollingH_DayView.Value)
    
    frmMain.EnableScrollV_WeekView = BinToBoolean(chkScrollingV_WeekView.Value)
        
    frmMain.EnableScrollV_MonthView = BinToBoolean(chkScrollingV_MonthView.Value)
        
    CalendarControl.DayView.EnableVScroll frmMain.EnableScrollV_DayView
    CalendarControl.DayView.EnableHScroll frmMain.EnableScrollH_DayView
    
    CalendarControl.WeekView.EnableVScroll frmMain.EnableScrollV_WeekView
    CalendarControl.MonthView.EnableVScroll frmMain.EnableScrollV_MonthView
    
    CalendarControl.EnableReminders BinToBoolean(chkEnableReminders.Value)

    Dim bPrev As Boolean
    bPrev = frmMain.g_bUseBuiltInCalendarDialogs
    frmMain.g_bUseBuiltInCalendarDialogs = BinToBoolean(chkUseBuiltInDialogs.Value)
    
    If bPrev <> frmMain.g_bUseBuiltInCalendarDialogs Then
    
        If frmMain.g_bUseBuiltInCalendarDialogs Then
            Set frmMain.g_dlgCalendarReminders.Calendar = CalendarControl
            frmMain.g_dlgCalendarReminders.ParentHWND = frmMain.hwnd
            frmMain.g_dlgCalendarReminders.RemindersWindowShowInTaskBar = True
        
            frmMain.g_dlgCalendarReminders.CreateRemindersWindow
            frmReminders.Visible = False
        Else
            frmMain.g_dlgCalendarReminders.CloseRemindersWindow
        End If
        
    End If
    
    CalendarControl.RedrawControl
        
    Unload Me
End Sub


Private Sub Form_Load()
    chkEditSubject_ByMouseClick.Value = BooleanToBin(CalendarControl.Options.EnableInPlaceEditEventSubject_ByMouseClick)
    chkEditSubject_AfterResize.Value = BooleanToBin(CalendarControl.Options.EnableInPlaceEditEventSubject_AfterEventResize)
    
    chkEditSubject_ByF2.Value = BooleanToBin(CalendarControl.Options.EnableInPlaceEditEventSubject_ByF2)
    chkEditSubject_ByTab.Value = BooleanToBin(CalendarControl.Options.EnableInPlaceEditEventSubject_ByTab)
    
    chkInplaceCreateEvent.Value = BooleanToBin(CalendarControl.Options.EnableInPlaceCreateEvent)
    chkUseOutlookFontGlyphs.Value = BooleanToBin(CalendarControl.Options.UseOutlookFontGlyphs)
    chkShowMinutes.Value = BooleanToBin(CalendarControl.Options.DayViewTimeScaleShowMinutes)
            
    chkDisableDragging_ForRecurrenceEvents.Value = BooleanToBin(frmMain.DisableDragging_ForRecurrenceEvents)
    chkDisableInPlaceCreateEvents_ForSaSu.Value = BooleanToBin(frmMain.DisableInPlaceCreateEvents_ForSaSu)
    
    chkScrollingV_DayView.Value = BooleanToBin(frmMain.EnableScrollV_DayView)
    chkScrollingH_DayView.Value = BooleanToBin(frmMain.EnableScrollH_DayView)
    
    chkScrollingV_WeekView.Value = BooleanToBin(frmMain.EnableScrollV_WeekView)
    
    chkScrollingV_MonthView.Value = BooleanToBin(frmMain.EnableScrollV_MonthView)
    
    chkEnableReminders.Value = BooleanToBin(CalendarControl.IsRemindersEnabled)
    
    optTimeMarkForToday.Value = (CalendarControl.Options.DayViewCurrentTimeMarkVisible And xtpCalendarCurrentTimeMarkVisibleForToday) <> 0
    optTimeMarkAlways.Value = (CalendarControl.Options.DayViewCurrentTimeMarkVisible And xtpCalendarCurrentTimeMarkVisibleAlways) <> 0
    optTimeMarkNone.Value = Not (optTimeMarkForToday.Value Or optTimeMarkAlways.Value)
    chkTimeMarkPrint.Value = BooleanToBin((CalendarControl.Options.DayViewCurrentTimeMarkVisible And xtpCalendarCurrentTimeMarkPrinted) <> 0)
        
    chkUseBuiltInDialogs.Value = BooleanToBin(frmMain.g_bUseBuiltInCalendarDialogs)
        
    frmMain.ModalFormsRunningCounter = frmMain.ModalFormsRunningCounter + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.ModalFormsRunningCounter = frmMain.ModalFormsRunningCounter - 1
End Sub

Private Sub optTimeMarkAlways_Click()
    chkTimeMarkPrint.Enabled = Not optTimeMarkNone.Value
End Sub

Private Sub optTimeMarkForToday_Click()
    chkTimeMarkPrint.Enabled = Not optTimeMarkNone.Value
End Sub

Private Sub optTimeMarkNone_Click()

    chkTimeMarkPrint.Enabled = Not optTimeMarkNone.Value

End Sub
