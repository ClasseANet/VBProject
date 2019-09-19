VERSION 5.00
Begin VB.UserControl ctrlThemeDVAllDayEv 
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   ScaleHeight     =   1530
   ScaleWidth      =   4410
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor1 
      Height          =   435
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor2 
      Height          =   435
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor3 
      Height          =   435
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   767
   End
   Begin VB.Label Label3 
      Caption         =   "Selected Background Color"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "Bottom Border Color"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Background Color"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1995
   End
End
Attribute VB_Name = "ctrlThemeDVAllDayEv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private m_pAllDayEvents As CalendarThemeDayViewAllDayEvents

Public Sub SetData(pAllDayEvents As CalendarThemeDayViewAllDayEvents)
    Debug.Assert Not pAllDayEvents Is Nothing
    
    Set m_pAllDayEvents = pAllDayEvents
    
    ctrlThemeColor1.Color = m_pAllDayEvents.BottomBorderColor
    ctrlThemeColor2.Color = m_pAllDayEvents.BackgroundColor
    ctrlThemeColor3.Color = m_pAllDayEvents.SelectedBackgroundColor
           
End Sub

Public Sub UpdateData()
    
    If m_pAllDayEvents.BottomBorderColor <> ctrlThemeColor1.Color Then
        m_pAllDayEvents.BottomBorderColor = ctrlThemeColor1.Color
    End If

    If m_pAllDayEvents.BackgroundColor <> ctrlThemeColor2.Color Then
        m_pAllDayEvents.BackgroundColor = ctrlThemeColor2.Color
    End If
    
    If m_pAllDayEvents.SelectedBackgroundColor <> ctrlThemeColor3.Color Then
        m_pAllDayEvents.SelectedBackgroundColor = ctrlThemeColor3.Color
    End If
End Sub


