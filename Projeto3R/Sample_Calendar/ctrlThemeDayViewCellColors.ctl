VERSION 5.00
Begin VB.UserControl ctrlThemeDVCellColors 
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   ScaleHeight     =   1995
   ScaleWidth      =   4350
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
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor4 
      Height          =   435
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
   End
   Begin VB.Label Label4 
      Caption         =   "Border Bottom Hour Color "
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "Border Bottom In-hour Color "
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "Background Color"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Selected Color"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1995
   End
End
Attribute VB_Name = "ctrlThemeDVCellColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pDayViewCellColors As CalendarThemeDayViewCellColors

Public Sub SetData(pDayViewCellColors As CalendarThemeDayViewCellColors)
    Debug.Assert Not pDayViewCellColors Is Nothing
    
    Set m_pDayViewCellColors = pDayViewCellColors
    
    ctrlThemeColor1.Color = m_pDayViewCellColors.BackgroundColor
    ctrlThemeColor2.Color = m_pDayViewCellColors.SelectedColor
    ctrlThemeColor3.Color = m_pDayViewCellColors.BorderBottomHourColor
    ctrlThemeColor4.Color = m_pDayViewCellColors.BorderBottomInHourColor
           
End Sub

Public Sub UpdateData()
    
    If m_pDayViewCellColors.BackgroundColor <> ctrlThemeColor1.Color Then
        m_pDayViewCellColors.BackgroundColor = ctrlThemeColor1.Color
    End If

    If m_pDayViewCellColors.SelectedColor <> ctrlThemeColor2.Color Then
        m_pDayViewCellColors.SelectedColor = ctrlThemeColor2.Color
    End If
    
    If m_pDayViewCellColors.BorderBottomHourColor <> ctrlThemeColor3.Color Then
        m_pDayViewCellColors.BorderBottomHourColor = ctrlThemeColor3.Color
    End If
    
    If m_pDayViewCellColors.BorderBottomInHourColor <> ctrlThemeColor4.Color Then
        m_pDayViewCellColors.BorderBottomInHourColor = ctrlThemeColor4.Color
    End If
    
End Sub


