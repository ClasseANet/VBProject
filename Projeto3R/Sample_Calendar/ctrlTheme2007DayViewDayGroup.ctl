VERSION 5.00
Begin VB.UserControl ctrlTheme2007DVDayGroup 
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   ScaleHeight     =   540
   ScaleWidth      =   4395
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor1 
      Height          =   435
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   767
   End
   Begin VB.Label Label2 
      Caption         =   "Left Border Color"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "ctrlTheme2007DVDayGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pDayViewDayGroup As CalendarThemeOffice2007DayViewDayGroup

Public Sub SetData(pDayViewDayGroup As CalendarThemeOffice2007DayViewDayGroup)
    Debug.Assert Not pDayViewDayGroup Is Nothing
    
    Set m_pDayViewDayGroup = pDayViewDayGroup
    
    ctrlThemeColor1.Color = m_pDayViewDayGroup.BorderLeftColor
               
End Sub

Public Sub UpdateData()
    
    If m_pDayViewDayGroup.BorderLeftColor <> ctrlThemeColor1.Color Then
        m_pDayViewDayGroup.BorderLeftColor = ctrlThemeColor1.Color
    End If

End Sub


