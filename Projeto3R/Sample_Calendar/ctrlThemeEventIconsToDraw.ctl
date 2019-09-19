VERSION 5.00
Begin VB.UserControl ctrlThemeEventIcons 
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ScaleHeight     =   1665
   ScaleWidth      =   4065
   Begin VB.Frame Frame1 
      Caption         =   "Event Icons to draw"
      Height          =   1335
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3615
      Begin VB.CheckBox chkShowException 
         Caption         =   "Show Exception"
         Height          =   255
         Left            =   1860
         TabIndex        =   5
         Top             =   540
         Width           =   1455
      End
      Begin VB.CheckBox chkShowMeeting 
         Caption         =   "Show Meeting"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkShowPrivate 
         Caption         =   "Show Private"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   1695
      End
      Begin VB.CheckBox chkShowOccurrence 
         Caption         =   "Show Occurrence"
         Height          =   255
         Left            =   1860
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
      Begin VB.CheckBox chkShowReminder 
         Caption         =   "Show Reminder"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1695
      End
   End
End
Attribute VB_Name = "ctrlThemeEventIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pIcons As CalendarThemeEventIconsToDraw

Public Sub SetData(pIcons As CalendarThemeEventIconsToDraw)
    Debug.Assert Not pIcons Is Nothing
    
    Set m_pIcons = pIcons
    
    chkShowReminder.Value = BooleanToBin(m_pIcons.ShowReminder)
    chkShowOccurrence.Value = BooleanToBin(m_pIcons.ShowOccurrence)
    chkShowException.Value = BooleanToBin(m_pIcons.ShowException)
    chkShowMeeting.Value = BooleanToBin(m_pIcons.ShowMeeting)
    chkShowPrivate.Value = BooleanToBin(m_pIcons.ShowPrivate)
           
End Sub

Public Sub UpdateData()
    If m_pIcons.ShowReminder <> BinToBoolean(chkShowReminder.Value) Then
        m_pIcons.ShowReminder = BinToBoolean(chkShowReminder.Value)
    End If
    
    If m_pIcons.ShowOccurrence <> BinToBoolean(chkShowOccurrence.Value) Then
        m_pIcons.ShowOccurrence = BinToBoolean(chkShowOccurrence.Value)
    End If
    
    If m_pIcons.ShowException <> BinToBoolean(chkShowException.Value) Then
        m_pIcons.ShowException = BinToBoolean(chkShowException.Value)
    End If
    
    If m_pIcons.ShowMeeting <> BinToBoolean(chkShowMeeting.Value) Then
        m_pIcons.ShowMeeting = BinToBoolean(chkShowMeeting.Value)
    End If
    
    If m_pIcons.ShowPrivate <> BinToBoolean(chkShowPrivate.Value) Then
        m_pIcons.ShowPrivate = BinToBoolean(chkShowPrivate.Value)
    End If
    
End Sub

