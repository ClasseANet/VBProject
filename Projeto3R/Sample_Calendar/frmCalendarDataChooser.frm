VERSION 5.00
Begin VB.Form frmCalendarDataChooser 
   Caption         =   "Select data provider and data file"
   ClientHeight    =   3630
   ClientLeft      =   2775
   ClientTop       =   6855
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton rdDP_MAPI 
      Caption         =   "Use MAPI data provider (Outlook)"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   3555
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2588
      TabIndex        =   12
      Top             =   3120
      Width           =   1155
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3968
      TabIndex        =   11
      Top             =   3120
      Width           =   1155
   End
   Begin VB.OptionButton rdDP_MySQL 
      Caption         =   "Use MySQL data provider"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3555
   End
   Begin VB.OptionButton rdDP_SQLServer 
      Caption         =   "Use SQL Server data provider"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1380
      Width           =   3555
   End
   Begin VB.TextBox txtConnectionStr 
      Height          =   315
      Left            =   420
      TabIndex        =   3
      Text            =   "events.mdb"
      Top             =   2400
      Width           =   7095
   End
   Begin VB.OptionButton rdDP_Access 
      Caption         =   "Use DataBase data provider (MS Access)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3555
   End
   Begin VB.OptionButton rdDP_Memory 
      Caption         =   "Use Memory data provider"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3555
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Built-in"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000004&
      Caption         =   "Custom"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3780
      TabIndex        =   10
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000004&
      Caption         =   "Custom"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3780
      TabIndex        =   9
      Top             =   1380
      Width           =   675
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000004&
      Caption         =   "Built-in"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   675
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000004&
      Caption         =   "Built-in"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   180
      Width           =   675
   End
   Begin VB.Label txtLabelBottom 
      Caption         =   "Load/Save events on the fly (when we need show or change data)"
      Height          =   255
      Left            =   420
      TabIndex        =   4
      Top             =   2760
      Width           =   7095
   End
   Begin VB.Label txtLabelTop 
      Caption         =   "DataBase file (*.mdb):"
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   2100
      Width           =   7095
   End
End
Attribute VB_Name = "frmCalendarDataChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type CJProviderInfo
    eTypeID         As CodeJockCalendarDataType
    
    strLabelTop     As String
    strLabelBottom  As String
    
    strConnectionPath     As String
End Type

Private m_pProviders(5)   As CJProviderInfo
Private m_eActiveProvider As CodeJockCalendarDataType
Public Cancelled As Boolean

Public Property Get ProviderType() As CodeJockCalendarDataType
   ProviderType = GetProviderInfo(m_eActiveProvider).eTypeID
End Property

Public Property Get ConnectionString() As String
    Dim strConnEx, strPath As String
    Dim eDPType As CodeJockCalendarDataType
    strPath = GetProviderInfo(m_eActiveProvider).strConnectionPath
    eDPType = ProviderType
    
    ' default, Memory provider (+ save to Bin or XML file)
    If eDPType = cjCalendarData_Memory Then
        strConnEx = "Provider=XML;Data Source='" & App.Path & "\" & strPath & "';Encoding=iso-8859-1;"
    End If
    
    ' MS Access provider
    If eDPType = cjCalendarData_Access Then
        strConnEx = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & App.Path & "\" & strPath & "';"
    End If
    
    ' MAPI provider
    If eDPType = cjCalendarData_MAPI Then
        strConnEx = "Provider=MAPI;"
    End If
    
    ' SQL Server provider
    If eDPType = cjCalendarData_SQLServer Then
        strConnEx = "Provider=Custom;" & strPath
    End If
    
    ' MySQL provider
    If eDPType = cjCalendarData_MySQL Then
        strConnEx = "Provider=Custom;" & strPath
    End If
    
    ConnectionString = strConnEx
    
End Property

Private Function GetProviderInfo(eTypeID As CodeJockCalendarDataType) As CJProviderInfo
    Dim i As Long
    
    For i = 1 To UBound(m_pProviders)
        If m_pProviders(i).eTypeID = eTypeID Then
            GetProviderInfo = m_pProviders(i)
            Exit Function
        End If
    Next
    
    GetProviderInfo.eTypeID = cjCalendarData_Unknown
End Function

Private Sub SaveActiveProviderInfo()
    Dim i As Long
    
    For i = 1 To UBound(m_pProviders)
        If m_pProviders(i).eTypeID = m_eActiveProvider Then
            m_pProviders(i).strConnectionPath = txtConnectionStr.Text
            Exit Sub
        End If
    Next
    
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Unload Me
End Sub

Private Sub btnOK_Click()
    
    SaveActiveProviderInfo
    
    '----------------------------------------------------------
    SaveSetting "Codejock Calendar VB Sample", "Provider", "Memory", m_pProviders(1).strConnectionPath
    SaveSetting "Codejock Calendar VB Sample", "Provider", "Access", m_pProviders(2).strConnectionPath
    SaveSetting "Codejock Calendar VB Sample", "Provider", "MAPI", m_pProviders(3).strConnectionPath
    SaveSetting "Codejock Calendar VB Sample", "Provider", "SQLServer", m_pProviders(4).strConnectionPath
    SaveSetting "Codejock Calendar VB Sample", "Provider", "MySQL", m_pProviders(5).strConnectionPath
        
    '============================================================
    btnOK.Enabled = False
    btnCancel.Enabled = False
    
    SaveSetting "Codejock Calendar VB Sample", "Provider", "Active", m_eActiveProvider
    SaveSetting "Codejock Calendar VB Sample", "Provider", "ActivePath", GetProviderInfo(m_eActiveProvider).strConnectionPath
    
    Cancelled = False
    Unload Me
End Sub

Private Sub Form_Load()
    Cancelled = False

    Dim provInf As CJProviderInfo
    
    '----------------------------------------------------------
    provInf.eTypeID = cjCalendarData_Memory
    provInf.strLabelTop = "XML (binary) data file:"
    provInf.strLabelBottom = "Load/Save events when app Start/Exit."
    provInf.strConnectionPath = GetSetting("Codejock Calendar VB Sample", "Provider", "Memory", "events.xml")
    
    m_pProviders(1) = provInf
    
    '----------------------------------------------------------
    provInf.eTypeID = cjCalendarData_Access
    provInf.strLabelTop = "DataBase file (*.mdb):"
    provInf.strLabelBottom = "Load/Save events on the fly (when we need show or change data)."
    provInf.strConnectionPath = GetSetting("Codejock Calendar VB Sample", "Provider", "Access", "events.mdb")
    
    m_pProviders(2) = provInf
    
    '----------------------------------------------------------
    provInf.eTypeID = cjCalendarData_MAPI
    provInf.strLabelTop = ""
    provInf.strLabelBottom = "Work with events from your Outlook calendar."
    provInf.strConnectionPath = GetSetting("Codejock Calendar VB Sample", "Provider", "MAPI", "")
    
    m_pProviders(3) = provInf
    
    '----------------------------------------------------------
    provInf.eTypeID = cjCalendarData_SQLServer
    provInf.strLabelTop = "Connection string:"
    'provInf.strLabelBottom = "Load/Save events on the fly (when we need show or change data)."
    provInf.strConnectionPath = GetSetting("Codejock Calendar VB Sample", "Provider", "SQLServer", "DSN=Calendar_SQLServer")
    
    m_pProviders(4) = provInf
    
    '----------------------------------------------------------
    provInf.eTypeID = cjCalendarData_MySQL
    'provInf.strLabelTop = "Connection string:"
    'provInf.strLabelBottom = "Load/Save events on the fly (when we need show or change data)."
    provInf.strConnectionPath = GetSetting("Codejock Calendar VB Sample", "Provider", "MySQL", "DSN=Calendar_MySQL")
    
    m_pProviders(5) = provInf
    
    '============================================================
    'm_eActiveProvider = GetSetting("Codejock Calendar VB Sample", "Provider", "Active", cjCalendarData_Unknown)
    m_eActiveProvider = frmMain.m_eActiveDataProvider
    UpdateControls m_eActiveProvider
    
    frmMain.ModalFormsRunningCounter = frmMain.ModalFormsRunningCounter + 1
End Sub

Private Sub UpdateControls(eTypeID As CodeJockCalendarDataType)

    Dim provInf As CJProviderInfo
    provInf = GetProviderInfo(eTypeID)
    
    rdDP_Memory.Value = CBool(eTypeID = cjCalendarData_Memory)
    rdDP_Access.Value = CBool(eTypeID = cjCalendarData_Access)
    rdDP_MAPI.Value = CBool(eTypeID = cjCalendarData_MAPI)
    rdDP_SQLServer.Value = CBool(eTypeID = cjCalendarData_SQLServer)
    rdDP_MySQL.Value = CBool(eTypeID = cjCalendarData_MySQL)
    
    txtLabelTop.Enabled = provInf.eTypeID <> cjCalendarData_Unknown
    txtLabelBottom.Enabled = provInf.eTypeID <> cjCalendarData_Unknown
    txtConnectionStr.Enabled = provInf.eTypeID <> cjCalendarData_Unknown
    
    txtLabelTop.Caption = provInf.strLabelTop
    txtLabelBottom.Caption = provInf.strLabelBottom
    txtConnectionStr.Text = provInf.strConnectionPath
    
    m_eActiveProvider = eTypeID

End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMain.ModalFormsRunningCounter = frmMain.ModalFormsRunningCounter - 1
End Sub

Private Sub rdDP_Access_Click()
    SaveActiveProviderInfo
    UpdateControls cjCalendarData_Access
End Sub

Private Sub rdDP_MAPI_Click()
    SaveActiveProviderInfo
    UpdateControls cjCalendarData_MAPI
End Sub

Private Sub rdDP_Memory_Click()
    SaveActiveProviderInfo
    UpdateControls cjCalendarData_Memory
End Sub

Private Sub rdDP_MySQL_Click()
    SaveActiveProviderInfo
    UpdateControls cjCalendarData_MySQL
End Sub

Private Sub rdDP_SQLServer_Click()
    SaveActiveProviderInfo
    UpdateControls cjCalendarData_SQLServer
End Sub

