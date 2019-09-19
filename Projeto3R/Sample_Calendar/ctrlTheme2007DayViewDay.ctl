VERSION 5.00
Begin VB.UserControl ctrlTheme2007DayViewDay 
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   ScaleHeight     =   1845
   ScaleWidth      =   4380
   Begin VB.CheckBox chkUseOffice2003HeaderFormat 
      Caption         =   "Use Office 2003 Header Date Format"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   4215
   End
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
   Begin VB.Label Label3 
      Caption         =   "Today Border Color"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "Border Color"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "ctrlTheme2007DayViewDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pDayViewDay As CalendarThemeOffice2007DayViewDay

Public Sub SetData(pDayViewDay As CalendarThemeOffice2007DayViewDay)
    Debug.Assert Not pDayViewDay Is Nothing
    
    Set m_pDayViewDay = pDayViewDay
    
    ctrlThemeColor1.Color = m_pDayViewDay.BorderColor
    ctrlThemeColor2.Color = m_pDayViewDay.TodayBorderColor
           
    chkUseOffice2003HeaderFormat.Value = IIf(m_pDayViewDay.UseOffice2003HeaderFormat, 1, 0)
           
End Sub

Public Sub UpdateData()
    
    If m_pDayViewDay.BorderColor <> ctrlThemeColor1.Color Then
        m_pDayViewDay.BorderColor = ctrlThemeColor1.Color
    End If

    If m_pDayViewDay.TodayBorderColor <> ctrlThemeColor2.Color Then
        m_pDayViewDay.TodayBorderColor = ctrlThemeColor2.Color
    End If
    
    If m_pDayViewDay.UseOffice2003HeaderFormat <> (chkUseOffice2003HeaderFormat.Value = 1) Then
        m_pDayViewDay.UseOffice2003HeaderFormat = IIf(chkUseOffice2003HeaderFormat.Value = 1, True, False)
    End If
End Sub

