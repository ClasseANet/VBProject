VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTheme2007 
   Caption         =   "Calendar Theme: Office 2007 "
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10020
   StartUpPosition =   1  'CenterOwner
   Begin CalendarSample.ctrlThemeHeader ctrlThemeHeader3day 
      Height          =   855
      Left            =   5820
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      ShowHeightFormula=   -1  'True
      ShowToodayColor =   -1  'True
   End
   Begin CalendarSample.ctrlThemeHeader ctrlThemeHeader2ex 
      Height          =   855
      Left            =   4500
      TabIndex        =   21
      Top             =   4140
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ShowHeightFormula=   -1  'True
      ShowToodayColor =   0   'False
   End
   Begin CalendarSample.ctrlThemeHeader ctrlThemeHeader1 
      Height          =   375
      Left            =   3180
      TabIndex        =   20
      Top             =   4140
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ShowHeightFormula=   0   'False
      ShowToodayColor =   -1  'True
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   5220
      Width           =   1035
   End
   Begin CalendarSample.ctrlThemeMVHeaderW ctrlThemeMVHeaderW1 
      Height          =   375
      Left            =   8400
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
   End
   Begin CalendarSample.ctrlTheme2007MVEventSD ctrlTheme2007MVEventSD1 
      Height          =   795
      Left            =   6240
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1402
   End
   Begin CalendarSample.ctrlThemeHeightFormula ctrlThemeHeightFormula1 
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   1500
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin CalendarSample.ctrlTheme2007MVDay ctrlTheme2007MonthViewDay1 
      Height          =   375
      Left            =   5580
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
   End
   Begin CalendarSample.ctrlTheme2007EventEx ctrlTheme2007EventEx1 
      Height          =   435
      Left            =   5640
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlTheme2007EventDVMD ctrlTheme2007EventDayViewMultiDay1 
      Height          =   435
      Left            =   4680
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlThemeDayViewCell ctrlThemeDayViewCell1 
      Height          =   435
      Left            =   4200
      TabIndex        =   12
      Top             =   3180
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlThemeDVAllDayEv ctrlThemeDayViewAlldayEvents1 
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   2700
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   979
   End
   Begin CalendarSample.ctrlTheme2007DVDayGroup ctrlTheme2007DayViewDayGroup1 
      Height          =   375
      Left            =   3540
      TabIndex        =   10
      Top             =   2220
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   2461
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlTheme2007DayViewDay ctrlTheme2007DayViewDay1 
      Height          =   315
      Left            =   3300
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1720
   End
   Begin CalendarSample.ctrlTheme2007Event ctrlTheme2007Event1 
      Height          =   375
      Left            =   3300
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   3519
      _ExtentY        =   1296
   End
   Begin CalendarSample.ctrlTheme2007EventFC ctrlTheme2007EventFontColors1 
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2778
      _ExtentY        =   1085
   End
   Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor1 
      Height          =   435
      Left            =   5460
      TabIndex        =   6
      Top             =   300
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   2355
      _ExtentY        =   1296
   End
   Begin CalendarSample.ctrlThemeTimeScale ctrlThemeTimeScale1 
      Height          =   555
      Left            =   8220
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   4260
      _ExtentY        =   2355
   End
   Begin CalendarSample.ctrlThemeBase ctrlThemeBase1 
      Height          =   435
      Left            =   3180
      TabIndex        =   4
      Top             =   300
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   7964
      _ExtentY        =   2672
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   7740
      TabIndex        =   3
      Top             =   5220
      Width           =   1035
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8820
      TabIndex        =   2
      Top             =   5220
      Width           =   1035
   End
   Begin ComctlLib.TreeView ctrlSettingsTree 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9763
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   443
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin CalendarSample.ctrlThemeHeaderText ctrlThemeHeaderText1 
      Height          =   435
      Left            =   4500
      TabIndex        =   1
      Top             =   300
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   2990
      _ExtentY        =   1085
   End
   Begin CalendarSample.ctrlTheme2007MVDay ctrlTheme2007WeekVDay 
      Height          =   495
      Left            =   7920
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3836
      _ExtentY        =   661
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Height          =   5055
      Left            =   2880
      Top             =   60
      Width           =   7095
   End
   Begin VB.Shape ctrlPlaceHolder 
      Height          =   4935
      Left            =   2940
      Top             =   120
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "frmTheme2007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ThemePropPageData
    pData As Object
    pPage As Object
End Type

Private m_arPages(100) As ThemePropPageData
Private m_nPagesCount As Long

Private m_nActivePageIndex As Long

Private m_nMinWidth As Long, m_nMinHeight As Long

'====================================================================
Private Function GetNextPageIndex() As Long
    m_nPagesCount = m_nPagesCount + 1
    GetNextPageIndex = m_nPagesCount
End Function

Private Sub btnApply_Click()
    If m_nActivePageIndex > 0 And Not m_arPages(m_nActivePageIndex).pPage Is Nothing Then
        m_arPages(m_nActivePageIndex).pPage.UpdateData
    End If
    
    frmMain.CalendarControl.Theme.RefreshMetrics
    frmMain.CalendarControl.Populate
    
    SaveCfg
End Sub

Private Sub SaveCfg()
    Dim px As PropExchange
    Set px = XtremeCalendarControl.CreatePropExchange
    
    If px.CreateAsXML(False, "CalendarThemeOffice2007") Then
        frmMain.CalendarControl.Theme.DoPropExchange px
        px.SaveToFile App.Path & "\cfgCalendarThemeOffice2007.xml"
    End If
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnReset_Click()
    Dim objThemeOfice2007 As New CalendarThemeOffice2007
    
    Dim pPX As PropExchange
    Set pPX = XtremeCalendarControl.CreatePropExchange
    Dim strData As String
        
    ' create PX to store
    If Not pPX.CreateAsXML(False, "CalendarThemeOffice2007") Then
        Exit Sub
    End If
    
    ' save default settings
    objThemeOfice2007.DoPropExchange pPX
    
    ' store configuration data to string
    strData = pPX.Value
        
    ' create PX to load
    If Not pPX.CreateAsXML(True, "CalendarThemeOffice2007") Then
        Exit Sub
    End If
        
    ' set configuration data from string
    pPX.Value = strData
        
    ' is data valid
    If pPX.Valid Then
    
        ' load default settings to current calendar theme object
        frmMain.CalendarControl.Theme.DoPropExchange pPX
        frmMain.CalendarControl.Theme.RefreshMetrics
        
        ' update this form controls
        If m_nActivePageIndex > 0 And Not m_arPages(m_nActivePageIndex).pPage Is Nothing Then
            m_arPages(m_nActivePageIndex).pPage.SetData m_arPages(m_nActivePageIndex).pData
        End If
    End If
            
    '=====================================================
    ' Other way
    '
    '' create and set new theme object
    'Dim objThemeOfice2007 As New CalendarThemeOffice2007
    'frmMain.CalendarControl.SetTheme objThemeOfice2007
    '
    'Dim nX, nY
    'nX = Left
    'nY = Top
    '
    ''reconnect this form to new theme
    'Unload Me
    'frmTheme2007.Show vbModeless, frmMain
    '
    'frmTheme2007.Left = nX
    'frmTheme2007.Top = nY
    '''''''''''''''''''''''''''''''''
    
    frmMain.CalendarControl.Populate
    SaveCfg
End Sub

Private Sub ctrlSettingsTree_NodeClick(ByVal pNode As ComctlLib.Node)
    Debug.Assert m_nActivePageIndex >= 0 And m_nActivePageIndex <= m_nPagesCount
    
    If m_nActivePageIndex > 0 And Not m_arPages(m_nActivePageIndex).pPage Is Nothing Then
        m_arPages(m_nActivePageIndex).pPage.UpdateData
        m_arPages(m_nActivePageIndex).pPage.Visible = False
        
        frmMain.CalendarControl.Theme.RefreshMetrics
        frmMain.CalendarControl.Populate
    End If
    
    m_nActivePageIndex = pNode.Tag
    
    If Not m_arPages(pNode.Tag).pPage Is Nothing Then
        m_arPages(pNode.Tag).pPage.SetData m_arPages(pNode.Tag).pData
        m_arPages(pNode.Tag).pPage.Visible = True
                
        m_arPages(pNode.Tag).pPage.Left = ctrlPlaceHolder.Left
        m_arPages(pNode.Tag).pPage.Top = ctrlPlaceHolder.Top
        m_arPages(pNode.Tag).pPage.Width = ctrlPlaceHolder.Width
        m_arPages(pNode.Tag).pPage.Height = ctrlPlaceHolder.Height
    End If

End Sub

Private Sub Form_Load()
    
    m_nMinWidth = Width
    m_nMinHeight = Height

    m_nPagesCount = 0
    m_nActivePageIndex = 0
    
    Dim objTheme2007 As CalendarThemeOffice2007
    Set objTheme2007 = frmMain.CalendarControl.Theme
    
    If objTheme2007 Is Nothing Then
        Debug.Assert False
        Exit Sub
    End If
        
    Dim pNodeStart As Node
    
    Dim pNode0 As Node, pNode1 As Node, pNode2 As Node, pNode3 As Node, pNode4 As Node
                
    Set m_arPages(0).pData = Nothing
    Set m_arPages(0).pPage = Nothing
        
    ' Base
    Set pNode0 = ctrlSettingsTree.Nodes.Add(, , , "Base")
    pNode0.Expanded = True
    pNode0.Tag = GetNextPageIndex
    Set m_arPages(pNode0.Tag).pData = objTheme2007
    Set m_arPages(pNode0.Tag).pPage = ctrlThemeBase1
            
    Set pNodeStart = pNode0
    
    ' Header
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Header")
    'pNode1.Expanded = True
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.Header
    Set m_arPages(pNode1.Tag).pPage = ctrlThemeHeader1
            
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Text Center")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.Header.TextCenter
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeaderText1
    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Text Left/Right")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.Header.TextLeftRight
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeaderText1

    ' Event
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Event")
    'pNode1.Expanded = True
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.Event
    Set m_arPages(pNode1.Tag).pPage = ctrlTheme2007Event1

    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Normal")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.Event.Normal
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventFontColors1

    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Selected")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.Event.Selected
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventFontColors1
    
    ' Day View
    Set pNode0 = ctrlSettingsTree.Nodes.Add(, , , "Day View")
    pNode0.Expanded = True
    pNode0.Tag = 0
    
    ' Day View: Header
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Header")
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.DayView.Header
    Set m_arPages(pNode1.Tag).pPage = ctrlThemeHeader2ex
        
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Text Center")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.Header.TextCenter
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeaderText1
    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Text Left/Right")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.Header.TextLeftRight
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeaderText1

    ' Day View: Event
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Event")
    'pNode1.Expanded = True
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.DayView.Event
    Set m_arPages(pNode1.Tag).pPage = ctrlTheme2007EventEx1

    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Normal")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.Event.Normal
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventFontColors1
    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Selected")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.Event.Selected
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventFontColors1
        
    ' Day View: TimeScale
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "TimeScale")
    'pNode1.Expanded = True
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.DayView.TimeScale
    Set m_arPages(pNode1.Tag).pPage = ctrlThemeTimeScale1
    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Caption Text")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.TimeScale.Caption
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeFontColor1
    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Am-Pm Text")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.TimeScale.AmPmText
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeFontColor1
    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Time Text small")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.TimeScale.TimeTextSmall
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeFontColor1
    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Time Text Big base")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.TimeScale.TimeTextBigBase
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeFontColor1

    ' Day View: Day
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Day")
    pNode1.Expanded = True
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.DayView.Day
    Set m_arPages(pNode1.Tag).pPage = ctrlTheme2007DayViewDay1
    
    ' Day View: Day : Header
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Header")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.Day.Header
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeader3day
        
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Text Center")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.DayView.Day.Header.TextCenter
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeHeaderText1
    
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Text Left/Right")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.DayView.Day.Header.TextLeftRight
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeHeaderText1

    ' Day View: Day : Group
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Group")
    pNode2.Expanded = True
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.DayView.Day.Group
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007DayViewDayGroup1
    
    ' Day View: Day : Group : Header
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Header")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.DayView.Day.Group.Header
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeHeader3day 'ctrlThemeHeader2ex
               
    Set pNode4 = ctrlSettingsTree.Nodes.Add(pNode3.Index, tvwChild, , "Text Center")
    pNode4.Tag = GetNextPageIndex
    Set m_arPages(pNode4.Tag).pData = objTheme2007.DayView.Day.Group.Header.TextCenter
    Set m_arPages(pNode4.Tag).pPage = ctrlThemeHeaderText1
    
    Set pNode4 = ctrlSettingsTree.Nodes.Add(pNode3.Index, tvwChild, , "Text Left/Right")
    pNode4.Tag = GetNextPageIndex
    Set m_arPages(pNode4.Tag).pData = objTheme2007.DayView.Day.Group.Header.TextLeftRight
    Set m_arPages(pNode4.Tag).pPage = ctrlThemeHeaderText1
    
    ' Day View: Day : Group : AllDayEvents
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "All Day Events")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.DayView.Day.Group.AllDayEvents
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeDayViewAlldayEvents1
    
    ' Day View: Day : Group : Cell
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Cell")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.DayView.Day.Group.Cell
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeDayViewCell1
    
    ' Day View: Day : Group : Single Day Event
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Single Day Event")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.DayView.Day.Group.SingleDayEvent
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventEx1

    Set pNode4 = ctrlSettingsTree.Nodes.Add(pNode3.Index, tvwChild, , "Normal")
    pNode4.Tag = GetNextPageIndex
    Set m_arPages(pNode4.Tag).pData = objTheme2007.DayView.Day.Group.SingleDayEvent.Normal
    Set m_arPages(pNode4.Tag).pPage = ctrlTheme2007EventFontColors1
    
    Set pNode4 = ctrlSettingsTree.Nodes.Add(pNode3.Index, tvwChild, , "Selected")
    pNode4.Tag = GetNextPageIndex
    Set m_arPages(pNode4.Tag).pData = objTheme2007.DayView.Day.Group.SingleDayEvent.Selected
    Set m_arPages(pNode4.Tag).pPage = ctrlTheme2007EventFontColors1
    
    ' Day View: Day : Group : Multi Day Event
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Multi Day Event")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.DayView.Day.Group.MultiDayEvent
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventDayViewMultiDay1

    Set pNode4 = ctrlSettingsTree.Nodes.Add(pNode3.Index, tvwChild, , "Normal")
    pNode4.Tag = GetNextPageIndex
    Set m_arPages(pNode4.Tag).pData = objTheme2007.DayView.Day.Group.MultiDayEvent.Normal
    Set m_arPages(pNode4.Tag).pPage = ctrlTheme2007EventFontColors1
    
    Set pNode4 = ctrlSettingsTree.Nodes.Add(pNode3.Index, tvwChild, , "Selected")
    pNode4.Tag = GetNextPageIndex
    Set m_arPages(pNode4.Tag).pData = objTheme2007.DayView.Day.Group.MultiDayEvent.Selected
    Set m_arPages(pNode4.Tag).pPage = ctrlTheme2007EventFontColors1
    
    '//=====================================================
    ' Month View
    Set pNode0 = ctrlSettingsTree.Nodes.Add(, , , "Month View")
    pNode0.Expanded = True
    pNode0.Tag = 0
    
    ' Month View: Header
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Header")
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.MonthView.Header
    Set m_arPages(pNode1.Tag).pPage = ctrlThemeHeader2ex
        
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Text Center")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.MonthView.Header.TextCenter
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeaderText1
    
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Text Left/Right")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.MonthView.Header.TextLeftRight
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeaderText1

    ' Month View: WeekDayHeader
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "WeekDayHeader")
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.MonthView.WeekDayHeader
    Set m_arPages(pNode1.Tag).pPage = ctrlThemeMVHeaderW1
        
    ' Month View: WeekHeader
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "WeekHeader")
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.MonthView.WeekHeader
    Set m_arPages(pNode1.Tag).pPage = ctrlThemeMVHeaderW1
        
    ' Month View: Event
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Event")
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.MonthView.Event
    Set m_arPages(pNode1.Tag).pPage = ctrlTheme2007EventEx1

    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Normal")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.MonthView.Event.Normal
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventFontColors1

    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Selected")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.MonthView.Event.Selected
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventFontColors1
       
    ' Month View: Day
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Day")
    pNode1.Expanded = True
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.MonthView.Day
    Set m_arPages(pNode1.Tag).pPage = ctrlTheme2007MonthViewDay1
    
    ' Month View: Day : Header
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Header")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.MonthView.Day.Header
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeader3day
        
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Text Center")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.MonthView.Day.Header.TextCenter
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeHeaderText1
    
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Text Left/Right")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.MonthView.Day.Header.TextLeftRight
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeHeaderText1
                          
    ' Month View: Day : Single Day Event
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Single Day Event")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.MonthView.Day.SingleDayEvent
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007MVEventSD1

    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Normal")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.MonthView.Day.SingleDayEvent.Normal
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventFontColors1
    
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Selected")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.MonthView.Day.SingleDayEvent.Selected
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventFontColors1
    
    ' Month View: Day : Multi Day Event
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Multi Day Event")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.MonthView.Day.MultiDayEvent
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventEx1

    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Normal")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.MonthView.Day.MultiDayEvent.Normal
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventFontColors1
    
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Selected")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.MonthView.Day.MultiDayEvent.Selected
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventFontColors1
    
    '//=====================================================
    ' Week View
    Set pNode0 = ctrlSettingsTree.Nodes.Add(, , , "Week View")
    pNode0.Expanded = True
    pNode0.Tag = 0
    
    ' Week View: Event
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Event")
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.WeekView.Event
    Set m_arPages(pNode1.Tag).pPage = ctrlTheme2007EventEx1

    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Normal")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.WeekView.Event.Normal
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventFontColors1

    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Selected")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.WeekView.Event.Selected
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventFontColors1
       
    ' Week View: Day
    Set pNode1 = ctrlSettingsTree.Nodes.Add(pNode0.Index, tvwChild, , "Day")
    pNode1.Expanded = True
    pNode1.Tag = GetNextPageIndex
    Set m_arPages(pNode1.Tag).pData = objTheme2007.WeekView.Day
    Set m_arPages(pNode1.Tag).pPage = ctrlTheme2007WeekVDay 'ctrlTheme2007MonthViewDay1
    ctrlTheme2007WeekVDay.UseOffice2003HeaderFormatVisible = True
    
    ' Week View: Day : Header
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Header")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.WeekView.Day.Header
    Set m_arPages(pNode2.Tag).pPage = ctrlThemeHeader3day
        
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Text Center")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.WeekView.Day.Header.TextCenter
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeHeaderText1
    
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Text Left/Right")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.WeekView.Day.Header.TextLeftRight
    Set m_arPages(pNode3.Tag).pPage = ctrlThemeHeaderText1
                          
    ' Week View: Day : Single Day Event
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Single Day Event")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.WeekView.Day.SingleDayEvent
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007MVEventSD1

    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Normal")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.WeekView.Day.SingleDayEvent.Normal
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventFontColors1
    
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Selected")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.WeekView.Day.SingleDayEvent.Selected
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventFontColors1
    
    ' Week View: Day : Multi Day Event
    Set pNode2 = ctrlSettingsTree.Nodes.Add(pNode1.Index, tvwChild, , "Multi Day Event")
    pNode2.Tag = GetNextPageIndex
    Set m_arPages(pNode2.Tag).pData = objTheme2007.WeekView.Day.MultiDayEvent
    Set m_arPages(pNode2.Tag).pPage = ctrlTheme2007EventEx1

    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Normal")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.WeekView.Day.MultiDayEvent.Normal
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventFontColors1
    
    Set pNode3 = ctrlSettingsTree.Nodes.Add(pNode2.Index, tvwChild, , "Selected")
    pNode3.Tag = GetNextPageIndex
    Set m_arPages(pNode3.Tag).pData = objTheme2007.WeekView.Day.MultiDayEvent.Selected
    Set m_arPages(pNode3.Tag).pPage = ctrlTheme2007EventFontColors1
    
    '------------------------------------
    ctrlSettingsTree_NodeClick pNodeStart
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Width = m_nMinWidth
'    If Width < m_nMinWidth Then
 '       Width = m_nMinWidth
  '  End If
    
    If Height < m_nMinHeight Then
        Height = m_nMinHeight
    End If
    
    ctrlSettingsTree.Height = ScaleHeight - 120

End Sub

Private Sub Form_Unload(Cancel As Integer)
    btnApply_Click
End Sub
