VERSION 5.00
Begin VB.UserControl ctrlTheme2007EventDVMD 
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ScaleHeight     =   4935
   ScaleWidth      =   6975
   Begin CalendarSample.ctrlTheme2007EventEx ctrlTheme2007Event1 
      Height          =   3495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6165
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   3540
      Width           =   4515
      Begin VB.TextBox txtString2 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   2355
      End
      Begin VB.TextBox txtString1 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   " 'To' Date format string"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   " 'From' Date format string"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "ctrlTheme2007EventDVMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pDVDEventMultiDay As CalendarThemeOffice2007DayViewEventMultiDay

Public Sub SetData(pDVDEventMultiDay As CalendarThemeOffice2007DayViewEventMultiDay)
    Debug.Assert Not pDVDEventMultiDay Is Nothing
    
    Set m_pDVDEventMultiDay = pDVDEventMultiDay
    
    ctrlTheme2007Event1.SetData m_pDVDEventMultiDay
    
    txtString1.Text = m_pDVDEventMultiDay.DateFormatFrom
    txtString2.Text = m_pDVDEventMultiDay.DateFormatTo
           
End Sub

Public Sub UpdateData()
    
    ctrlTheme2007Event1.UpdateData
    
    If StrComp(m_pDVDEventMultiDay.DateFormatFrom, txtString1.Text) <> 0 Then
        m_pDVDEventMultiDay.DateFormatFrom = txtString1.Text
    End If

    If StrComp(m_pDVDEventMultiDay.DateFormatTo, txtString2.Text) <> 0 Then
        m_pDVDEventMultiDay.DateFormatTo = txtString2.Text
    End If
End Sub


Private Sub UserControl_Initialize()
    txtString1.ToolTipText = "d, dd - Day of month; " & _
                             "ddd, dddd - Day of week; " & _
                             "M , MM, MMM, MMMM - Month " & _
                             "Y , yy, yyyy - Year "
                             
    txtString2.ToolTipText = txtString1.ToolTipText
    
    Label1.ToolTipText = txtString1.ToolTipText
    Label2.ToolTipText = txtString1.ToolTipText

End Sub
