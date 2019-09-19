VERSION 5.00
Begin VB.Form frmResourcesManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar Resources Manager"
   ClientHeight    =   7530
   ClientLeft      =   6555
   ClientTop       =   6120
   ClientWidth     =   7845
   Icon            =   "frmResourcesManager.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameResProps 
      Caption         =   "Selected resource properties"
      Height          =   1575
      Left            =   3600
      TabIndex        =   12
      Top             =   3240
      Width           =   4095
      Begin VB.ComboBox cmbResDP 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkNameAuto 
         Caption         =   "Auto"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtResourceName 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Text            =   "Resource name"
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblResDP 
         Caption         =   "Data provider:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4005
      TabIndex        =   11
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2625
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame frameSchedules 
      Caption         =   "Schedules"
      Height          =   2055
      Left            =   3600
      TabIndex        =   2
      Top             =   4920
      Width           =   4095
      Begin VB.CommandButton btnScheduleRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton btnScheduleChange 
         Caption         =   "Change"
         Height          =   375
         Left            =   3000
         TabIndex        =   22
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton btnScheduleAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtScheduleName 
         Height          =   405
         Left            =   120
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CheckBox chkSchedulesShowAll 
         Caption         =   "Show All"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox lstSchedules 
         Height          =   1185
         ItemData        =   "frmResourcesManager.frx":000C
         Left            =   120
         List            =   "frmResourcesManager.frx":000E
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame frameResources 
      Caption         =   "Resources"
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   3375
      Begin VB.CommandButton btnResRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton btnResAdd 
         Caption         =   "Add..."
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   855
      End
      Begin VB.ListBox lstResources 
         Height          =   2535
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame frameDP 
      Caption         =   "Data Providers"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton btnDPRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton btnDPAdd 
         Caption         =   "Add..."
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton btnDPChange 
         Caption         =   "Change..."
         Height          =   375
         Left            =   6120
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox lstDataProviders 
         Height          =   2205
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmResourcesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nRcNameAuto_StoredState As Integer
Private bProcessing  As Boolean
Public Cancelled As Boolean

Property Get CalendarControl() As CalendarControl
    Set CalendarControl = frmMain.CalendarControl
End Property

Property Get ResourceManager() As CalendarResourcesManager
    Set ResourceManager = frmMain.g_DataResourcesMan
End Property

Property Get GetSelRCDesc() As CalendarResourceDescription
    ' Find selected resource
    Dim i As Integer
    If lstResources.ListIndex < lstResources.ListCount And lstResources.ListIndex >= 0 Then
        Set GetSelRCDesc = ResourceManager.Resource(lstResources.ItemData(lstResources.ListIndex))
        Exit Property
    End If
    
    Set GetSelRCDesc = Nothing
End Property



Private Sub btnCancel_Click()
    Cancelled = True
    Unload Me
End Sub

Private Sub UpdateDataProvidersList()
    Dim nSel As Integer
    nSel = lstDataProviders.ListIndex
    
    ' Clear old list
    lstDataProviders.Clear
    
    ' Iterate all data providers
    Dim pDP As CalendarDataProvider
    Dim i As Integer
    For i = 0 To ResourceManager.DataProvidersCount - 1
        Set pDP = ResourceManager.DataProvider(i)
        lstDataProviders.AddItem (CStr(i) + ": " + pDP.ConnectionString)
    Next i
    
    ' adjust list scrolling
    mListBoxAdjustHScroll lstDataProviders
    
    ' restore list index
    If nSel >= 0 And nSel < lstDataProviders.ListCount Then lstDataProviders.ListIndex = nSel
    
    ' proceed controls populating
    UpdateDataProviders_RCCombo
End Sub

Private Sub UpdateDataProviders_RCCombo()
    ' Clear old contents
    cmbResDP.Clear
    
    ' Iterate all data providers
    Dim pDP As CalendarDataProvider
    Dim i As Integer
    For i = 0 To ResourceManager.DataProvidersCount - 1
        Set pDP = ResourceManager.DataProvider(i)
        cmbResDP.AddItem (CStr(i) + ": " + pDP.ConnectionString)
        cmbResDP.ItemData(i) = i
    Next i
    
    UpdateSchedulesList
End Sub

Private Sub UpdateSchedulesList()
    Dim nSel As Integer
    nSel = lstSchedules.ListIndex

    lstSchedules.Clear
    txtScheduleName.Text = ""

    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then
        EnableSchedulesInfoControls False
        Exit Sub
    End If
    
    If pResDesc.Resource.DataProvider Is Nothing Then
        EnableSchedulesInfoControls False
        Exit Sub
    End If

    ' Iterate schedules
    Dim i As Integer
    Dim pSchedules As CalendarSchedules
    Set pSchedules = pResDesc.Resource.DataProvider.Schedules
    If pSchedules Is Nothing Then
        EnableSchedulesInfoControls False
        Exit Sub
    End If
    
    Dim pSchedule As CalendarSchedule
    Dim bVisible As Boolean
    
    EnableSchedulesInfoControls True
    bProcessing = True
    
    For i = 0 To pSchedules.Count - 1
        Set pSchedule = pSchedules.Item(i)
        lstSchedules.AddItem (pSchedule.Name)
        lstSchedules.ItemData(i) = pSchedule.Id
        bVisible = pResDesc.Resource.ExistsScheduleID(pSchedule.Id, False)
        lstSchedules.Selected(i) = bVisible
    Next i
    
    ' restore list index
    If nSel >= 0 And nSel < lstSchedules.ListCount Then lstSchedules.ListIndex = nSel
    
    ' continue
    chkSchedulesShowAll.Value = BooleanToBin(pResDesc.Resource.IsSchedulesSetEmpty)
    UpdateSchedulesAll_DependsCtrls
    
    bProcessing = False
End Sub

Private Sub UpdateSchedulesAll_DependsCtrls()
    'lstSchedules.Enabled = BinToBoolean(chkSchedulesShowAll.Value)
End Sub

Private Sub UpdateResourcesList()
    Dim nSel As Integer
    nSel = lstResources.ListIndex
    
    lstResources.Clear
    ' add resources
    Dim i As Integer
    Dim pResDesc As CalendarResourceDescription
    For i = 0 To ResourceManager.ResourcesCount - 1
        Set pResDesc = ResourceManager.Resource(i)
        If pResDesc Is Nothing Or pResDesc.Resource Is Nothing Then Exit Sub
        lstResources.AddItem (pResDesc.Resource.Name)
        lstResources.ItemData(lstResources.ListCount - 1) = i
        lstResources.Selected(lstResources.ListCount - 1) = pResDesc.Enabled
    Next i
    
    ' restore list index
    If nSel >= 0 And nSel < lstResources.ListCount Then lstResources.ListIndex = nSel
    
End Sub

Private Sub UpdateResourceInfoPane()
    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then
        EnableResourceInfoPane (False)
        Exit Sub
    End If
    EnableResourceInfoPane (True)
   
    bProcessing = True
    ' Update properties
    txtResourceName.Text = pResDesc.Resource.Name
    chkNameAuto.Value = BooleanToBin(pResDesc.GenerateName)
    txtScheduleName.Enabled = Not pResDesc.GenerateName
    
    ' Set corresponding data provider in combobox
    Dim pDP As CalendarDataProvider
    Set pDP = pResDesc.Resource.DataProvider
    If Not pDP Is Nothing Then
        Dim nDPIndex As Integer
        nDPIndex = ResourceManager.GetDataProviderIndex(pDP.ConnectionString)
        If nDPIndex >= 0 And nDPIndex < cmbResDP.ListCount Then cmbResDP.ListIndex = nDPIndex
    End If
    
    chkSchedulesShowAll.Value = 0

    ' continue
    UpdateSchedulesList
    ProcessAutoName
    
    bProcessing = False
End Sub

Private Sub ProcessAutoName()
    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then Exit Sub
    If pResDesc.Resource Is Nothing Then Exit Sub
    If pResDesc.Resource.ScheduleIDs Is Nothing Then Exit Sub

    ' process
    Dim bAuto As Boolean
    bAuto = True
    If pResDesc.Resource.DataProvider Is Nothing Then
        bAuto = False
    Else
        If pResDesc.Resource.DataProvider.Schedules Is Nothing Then bAuto = False
    End If
    
    If Not bAuto Then
        chkNameAuto.Value = 0
        ApplyUpdateRcNameAuto
        chkNameAuto.Enabled = False
    End If

    ' Calculate automatic name
    Dim strNameNew As String
    strNameNew = CalcAutoRCName
    If strNameNew = "" Then bAuto = False
    chkNameAuto.Enabled = bAuto
    If bAuto Then
        If BinToBoolean(chkNameAuto.Value) Or (nRcNameAuto_StoredState = 1) Then
            nRcNameAuto_StoredState = -1
            If chkNameAuto.Value = 0 Then
                chkNameAuto.Value = 1
                ApplyUpdateRcNameAuto
            End If
            
            Dim strNamePrev As String
            strNamePrev = txtResourceName.Text
            If strNameNew <> strNamePrev Then
                txtResourceName.Text = strNameNew
            End If
        End If
    Else
        If nRcNameAuto_StoredState = -1 Then
            nRcNameAuto_StoredState = chkNameAuto.Value
        End If
        chkNameAuto.Value = 0
        ApplyUpdateRcNameAuto
    End If
End Sub

Function CalcAutoRCName() As String
    If BinToBoolean(chkSchedulesShowAll.Value) Then
        CalcAutoRCName = ""
        Exit Function
    End If
    
    ' Automatic name calculation algorithm could be any of your choice.
    ' There is implemented a simplest one with some redundancy
    Dim i As Integer
    Dim strName As String
    For i = 0 To lstSchedules.ListCount - 1
        If lstSchedules.Selected(i) Then
            strName = lstSchedules.List(i)
        End If
    Next i
    ' Finalize return value
    If lstSchedules.SelCount = 1 Then
        CalcAutoRCName = strName
    Else
        CalcAutoRCName = ""
    End If
End Function

Private Sub ApplyUpdateRcNameAuto()
    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then Exit Sub

    ' process
    txtResourceName.Enabled = Not BinToBoolean(chkNameAuto.Value)
    pResDesc.GenerateName = BinToBoolean(chkNameAuto.Value)

End Sub

Private Sub EnableResourceInfoPane(bEnable As Boolean)
    txtResourceName.Enabled = bEnable
    chkNameAuto.Enabled = bEnable
    cmbResDP.Enabled = bEnable
    
    EnableSchedulesInfoControls (bEnable)

End Sub

Private Sub EnableSchedulesInfoControls(bEnable As Boolean)
    chkSchedulesShowAll.Enabled = bEnable
    lstSchedules.Enabled = bEnable

    btnScheduleRemove.Enabled = bEnable
    btnScheduleAdd.Enabled = bEnable
    btnScheduleChange.Enabled = bEnable

End Sub

Private Sub btnDPAdd_Click()
    frmCalendarDataChooser.Show vbModal, frmResourcesManager
    
    If Not frmCalendarDataChooser.Cancelled Then
        Screen.MousePointer = vbHourglass
        
        Dim strConnectionString As String
        strConnectionString = frmCalendarDataChooser.ConnectionString
        
        Dim pDP As CalendarDataProvider
        Set pDP = ResourceManager.GetDataProvider(strConnectionString)
        If Not pDP Is Nothing Then Exit Sub
        
        If ResourceManager.AddDataProvider(strConnectionString, xtpCalendarDPF_CreateIfNotExists Or xtpCalendarDPF_SaveOnDestroy Or xtpCalendarDPF_CloseOnDestroy) Then
            UpdateDataProvidersList
        End If
        
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub btnDPRemove_Click()
    If lstDataProviders.ListIndex >= 0 And lstDataProviders.ListIndex < lstDataProviders.ListCount Then
        ResourceManager.RemoveDataProvider (lstDataProviders.ListIndex)
        UpdateDataProvidersList
    End If
End Sub

Private Sub btnOK_Click()
    Dim i As Integer
    Dim pRDesc As CalendarResourceDescription
    For i = 0 To lstResources.ListCount - 1
        Set pRDesc = ResourceManager.Resource(lstResources.ItemData(i))
        If Not pRDesc Is Nothing Then
            pRDesc.Enabled = lstResources.Selected(i)
        End If
    Next i
    
    Cancelled = False
    Unload Me
End Sub

Private Sub btnResAdd_Click()
    ResourceManager.AddResource "new resource", True
    UpdateResourcesList
    
    lstResources.ListIndex = lstResources.ListCount - 1
    UpdateResourceInfoPane
End Sub

Private Sub btnResRemove_Click()
    If lstResources.ListIndex >= 0 And lstResources.ListIndex < lstResources.ListCount Then
        ResourceManager.RemoveResource lstResources.ListIndex
        UpdateResourcesList
    End If
    UpdateResourceInfoPane
End Sub

Private Sub btnScheduleAdd_Click()
    If txtScheduleName.Text = "" Then Exit Sub
    
    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then Exit Sub
    If pResDesc.Resource Is Nothing Then Exit Sub

    Dim pDP As CalendarDataProvider
    Set pDP = pResDesc.Resource.DataProvider
    If pDP Is Nothing Then Exit Sub
    
    Dim pSchedules As CalendarSchedules
    Set pSchedules = pDP.Schedules
    If pSchedules Is Nothing Then Exit Sub
    
    ' add new one
    pSchedules.AddNewSchedule txtScheduleName.Text
    UpdateSchedulesList_saveState

    ' update
    If lstSchedules.ListCount > 0 Then
        lstSchedules.ListIndex = lstSchedules.ListCount - 1
        txtScheduleName = pSchedules.GetScheduleName(lstSchedules.ListCount - 1)
    End If
    
    ProcessAutoName
End Sub

Private Sub UpdateSchedulesList_saveState()
    Dim nSel As Integer
    nSel = lstSchedules.ListIndex
    
    Dim bAll As Boolean
    bAll = chkSchedulesShowAll.Enabled
    
    UpdateSchedulesList
    
    chkSchedulesShowAll.Enabled = bAll
    UpdateSchedulesAll_DependsCtrls
    
    If nSel >= 0 And nSel < lstSchedules.ListCount Then
        lstSchedules.ListIndex = nSel
        txtScheduleName.Text = lstSchedules.List(nSel)
    End If
    

End Sub

Private Sub btnScheduleChange_Click()
    If bProcessing Then Exit Sub
    
    Dim nSel As Integer
    nSel = lstSchedules.ListIndex
    If nSel < 0 Or nSel >= lstSchedules.ListCount Then Exit Sub

    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then Exit Sub
    If pResDesc.Resource Is Nothing Then Exit Sub

    Dim pDP As CalendarDataProvider
    Set pDP = pResDesc.Resource.DataProvider
    If pDP Is Nothing Then Exit Sub
    
    Dim pSchedules As CalendarSchedules
    Set pSchedules = pDP.Schedules
    If pSchedules Is Nothing Then Exit Sub
    
    ' update name
    pSchedules.SetScheduleName lstSchedules.ItemData(nSel), txtScheduleName.Text

    UpdateSchedulesList_saveState
    ProcessAutoName
End Sub

Private Sub btnScheduleRemove_Click()
    If lstSchedules.ListIndex < 0 Or lstSchedules.ListIndex >= lstSchedules.ListCount Then Exit Sub
    
    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then Exit Sub
    If pResDesc.Resource Is Nothing Then Exit Sub
    
    Dim pDP As CalendarDataProvider
    Set pDP = pResDesc.Resource.DataProvider
    If pDP Is Nothing Then Exit Sub
    
    Dim pSchedules As CalendarSchedules
    Set pSchedules = pDP.Schedules
    If pSchedules Is Nothing Then Exit Sub
    
    ' remove a schedule
    pSchedules.RemoveSchedule lstSchedules.ItemData(lstSchedules.ListIndex)
    
    UpdateSchedulesList_saveState
    ProcessAutoName
    
End Sub

Private Sub chkNameAuto_Click()
    If bProcessing Then Exit Sub
    
    ApplyUpdateRcNameAuto
    ProcessAutoName
End Sub

Private Sub chkSchedulesShowAll_Click()
    If bProcessing Then Exit Sub
    
    UpdateSchedulesAll_DependsCtrls
    ApplySchedules
    ProcessAutoName
End Sub

Private Sub cmbResDP_Change()
    If bProcessing Then Exit Sub
    
    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then
        UpdateResourceInfoPane
        Exit Sub
    End If
    If pResDesc.Resource Is Nothing Then
        UpdateResourceInfoPane
        Exit Sub
    End If
    
    If cmbResDP.ListIndex >= 0 And cmbResDP.ListIndex < cmbResDP.ListCount Then
        Dim pDP As CalendarDataProvider
        Set pDP = ResourceManager.DataProvider(cmbResDP.ListIndex)
        
        pResDesc.Resource.SetDataProvider pDP, False
        UpdateSchedulesList
    End If
     
End Sub

Private Sub cmbResDP_Click()
    cmbResDP_Change
End Sub

Private Sub Form_Load()
    nRcNameAuto_StoredState = -1
    bProcessing = False
    
    UpdateDataProvidersList
    UpdateResourcesList
    UpdateResourceInfoPane
End Sub

Private Sub lstResources_Click()
    If bProcessing Then Exit Sub
    
    UpdateResourceInfoPane
End Sub

Private Sub lstSchedules_Click()
    If bProcessing Then Exit Sub
    
    bProcessing = True
    ApplySchedules
    ProcessAutoName
    
    If lstSchedules.ListIndex < 0 Or lstSchedules.ListIndex >= lstSchedules.ListCount Then Exit Sub
    
    txtScheduleName.Text = lstSchedules.List(lstSchedules.ListIndex)
    
    bProcessing = False
End Sub

Private Sub txtResourceName_Change()
    If bProcessing Then Exit Sub
    
    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then Exit Sub
    If pResDesc.Resource Is Nothing Then Exit Sub

    ' set changed text
    pResDesc.Resource.Name = txtResourceName.Text
    lstResources.List(lstResources.ListIndex) = txtResourceName.Text
    
End Sub

Private Sub ApplySchedules()
    ' Get current resource description
    Dim pResDesc As CalendarResourceDescription
    Set pResDesc = GetSelRCDesc
    If pResDesc Is Nothing Then Exit Sub
    If pResDesc.Resource Is Nothing Then Exit Sub
    If pResDesc.Resource.ScheduleIDs Is Nothing Then Exit Sub
    
    pResDesc.Resource.ScheduleIDs.RemoveAll
    
    ' check show all button
    If BinToBoolean(chkSchedulesShowAll.Value) Then
        Exit Sub
    End If
    
    ' add necessary schedules
    Dim i As Integer
    For i = 0 To lstSchedules.ListCount - 1
        If lstSchedules.Selected(i) Then
            pResDesc.Resource.ScheduleIDs.Add lstSchedules.ItemData(i)
        End If
    Next i

End Sub
