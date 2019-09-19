VERSION 5.00
Begin VB.Form frmTimeZone 
   Caption         =   "Time Zone"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmAdditionalTimeZone 
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   5895
      Begin VB.CheckBox chkAutoAdjustDaylight2 
         Caption         =   "A&utomatically adjust clock for daylight saving changes"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   5535
      End
      Begin VB.TextBox txtLabel2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbTimeZone2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblLabel2 
         Caption         =   "&Label:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   405
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Time &Zone:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CheckBox chkShowAdditionalTimeZone 
      Caption         =   "Show an Additional Time Zone"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame frmCurrentTimeZone 
      Caption         =   "Current time zone"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.CheckBox chkAutoAdjustDaylight1 
         Caption         =   "Automatically adjust clock for daylight saving changes"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   5535
      End
      Begin VB.ComboBox cmbTimeZone 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "cmbTimeZone"
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtLabel 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Time &Zone:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblLabel 
         Caption         =   "&Label:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   400
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmTimeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private g_pTimeZones As CalendarTimeZones

Private Function IsAutoAdjustDaylight_Exists(pTimeZone As CalendarTimeZone)
    If pTimeZone.DaylightDate.wMonth <> 0 And _
        pTimeZone.StandardDate.wMonth <> 0 Then
        IsAutoAdjustDaylight_Exists = True
    Else
        IsAutoAdjustDaylight_Exists = False
    End If
End Function

Private Function IsAutoAdjustDaylight_Checked(pTimeZone As CalendarTimeZone)
    If pTimeZone.DaylightBias <> 0 Or pTimeZone.StandardBias <> 0 Then
        IsAutoAdjustDaylight_Checked = True
    Else
        IsAutoAdjustDaylight_Checked = False
    End If
End Function

Property Get CalendarControl() As CalendarControl
    Set CalendarControl = FrmCalendario.CalendarControl
End Property


Private Sub chkShowAdditionalTimeZone_Click()
    cmbTimeZone2.Enabled = chkShowAdditionalTimeZone.Value
    txtLabel2.Enabled = chkShowAdditionalTimeZone.Value
    chkAutoAdjustDaylight2.Enabled = chkShowAdditionalTimeZone.Value
End Sub


Private Sub cmbTimeZone2_Change()
    Dim tziScale2 As CalendarTimeZone
    Set tziScale2 = g_pTimeZones.Item(cmbTimeZone2.ListIndex)
    
    chkAutoAdjustDaylight2.Enabled = IsAutoAdjustDaylight_Exists(tziScale2)
    
    If IsAutoAdjustDaylight_Checked(tziScale2) Then
        chkAutoAdjustDaylight2.Value = 1
    Else
        chkAutoAdjustDaylight2.Value = 0
    End If
End Sub

Private Sub cmbTimeZone2_Click()
    cmbTimeZone2_Change
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error Resume Next
    'mvarme.CalendarControl.DayView.ScaleText = txtLabel.Text
    'mvarme.CalendarControl.DayView.Scale2Text = txtLabel2.Text
    mvarme.CalendarControl.Options.DayViewScaleLabel = txtLabel.Text
    mvarme.CalendarControl.Options.DayViewScale2Label = txtLabel2.Text
    
    If chkShowAdditionalTimeZone.Value = 0 Then
        'mvarme.CalendarControl.DayView.Scale2Visible = False
        mvarme.CalendarControl.Options.DayViewScale2Visible = False
    Else
        'mvarme.CalendarControl.DayView.Scale2Visible = True
        mvarme.CalendarControl.Options.DayViewScale2Visible = True
    End If
    
    Dim tziScale2 As CalendarTimeZone
    Set tziScale2 = g_pTimeZones.Item(cmbTimeZone2.ListIndex)
    
    If chkAutoAdjustDaylight2.Value = 0 Then
        tziScale2.StandardBias = 0
        tziScale2.DaylightBias = 0
    End If
        
    'mvarme.CalendarControl.DayView.SetScale2TimeZone tziScale2
    mvarme.CalendarControl.Options.SetScale2TimeZone tziScale2
             
    mvarme.CalendarControl.Populate
    
    Unload Me
End Sub

Private Sub Form_Load()
  On Error Resume Next
    Dim tziCurrent As CalendarTimeZone
    Dim tziTmp As CalendarTimeZone
    Dim tzi2 As CalendarTimeZone
    Dim bIsScale2Visible As Boolean
              
    Set tziCurrent = mvarme.CalendarControl.Options.GetCurrentTimeZone
    Set tzi2 = mvarme.CalendarControl.Options.GetScale2TimeZone
    Set g_pTimeZones = mvarme.CalendarControl.Options.EnumAllTimeZones
    bIsScale2Visible = mvarme.CalendarControl.Options.DayViewScale2Visible
        
    txtLabel.Text = mvarme.CalendarControl.Options.DayViewScaleLabel
    txtLabel2.Text = mvarme.CalendarControl.Options.DayViewScale2Label
    
    cmbTimeZone.Text = tziCurrent.DisplayString
        
    chkAutoAdjustDaylight1.Visible = IsAutoAdjustDaylight_Exists(tziCurrent)
    If IsAutoAdjustDaylight_Checked(tziCurrent) Then
        chkAutoAdjustDaylight1.Value = 1
    Else
        chkAutoAdjustDaylight1.Value = 0
    End If
    
    cmbTimeZone.Enabled = False
    chkAutoAdjustDaylight1.Enabled = False
    
    '==========================================================
    Dim nIndex As Long
    Dim nCurrentIndex As Long
       
    nIndex = 0
    nCurrentIndex = g_pTimeZones.Count / 2
    For Each tziTmp In g_pTimeZones
    
        cmbTimeZone2.AddItem tziTmp.DisplayString, nIndex
                
        If Not tzi2 Is Nothing Then
            If tzi2.IsEqual(tziTmp) Then
                nCurrentIndex = nIndex
            End If
        End If
                
        nIndex = nIndex + 1
    Next
        
    cmbTimeZone2.ListIndex = nCurrentIndex
    
    If bIsScale2Visible Then
        chkShowAdditionalTimeZone.Value = 1
    Else
        chkShowAdditionalTimeZone.Value = 0
    End If
        
    chkAutoAdjustDaylight2.Enabled = IsAutoAdjustDaylight_Exists(tzi2) And _
                                     bIsScale2Visible
    If IsAutoAdjustDaylight_Checked(tzi2) Then
        chkAutoAdjustDaylight2.Value = 1
    Else
        chkAutoAdjustDaylight2.Value = 0
    End If
    
    ModalFormsRunningCounter = ModalFormsRunningCounter + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ModalFormsRunningCounter = ModalFormsRunningCounter - 1
End Sub
