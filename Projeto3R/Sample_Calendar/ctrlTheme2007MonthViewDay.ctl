VERSION 5.00
Begin VB.UserControl ctrlTheme2007MVDay 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CheckBox chkUseOffice2003HeaderFormat 
      Caption         =   "Use Office 2003 Header Date Format"
      Enabled         =   0   'False
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
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
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor5 
      Height          =   435
      Left            =   2040
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
   End
   Begin VB.Label Label5 
      Caption         =   "Background Selected Color"
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   1995
   End
   Begin VB.Label Label4 
      Caption         =   "Background Light Color "
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "Background Dark Color "
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "Border Color"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Today Border Color"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1995
   End
End
Attribute VB_Name = "ctrlTheme2007MVDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pMVDay As Object ' CalendarThemeOffice2007MonthViewDay

Public Sub SetData(pMVDay As Object)  'CalendarThemeOffice2007MonthViewDay)
    Debug.Assert Not pMVDay Is Nothing
    
    Set m_pMVDay = pMVDay
    
    ctrlThemeColor1.Color = m_pMVDay.BorderColor
    ctrlThemeColor2.Color = m_pMVDay.TodayBorderColor
    ctrlThemeColor3.Color = m_pMVDay.BackgroundLightColor
    ctrlThemeColor4.Color = m_pMVDay.BackgroundDarkColor
    ctrlThemeColor5.Color = m_pMVDay.BackgroundSelectedColor
               
    If chkUseOffice2003HeaderFormat.Enabled Then
        chkUseOffice2003HeaderFormat.Value = IIf(m_pMVDay.UseOffice2003HeaderFormat, 1, 0)
    End If
               
End Sub

Public Sub UpdateData()
    
    If m_pMVDay.BorderColor <> ctrlThemeColor1.Color Then
        m_pMVDay.BorderColor = ctrlThemeColor1.Color
    End If

    If m_pMVDay.TodayBorderColor <> ctrlThemeColor2.Color Then
        m_pMVDay.TodayBorderColor = ctrlThemeColor2.Color
    End If
    
    If m_pMVDay.BackgroundLightColor <> ctrlThemeColor3.Color Then
        m_pMVDay.BackgroundLightColor = ctrlThemeColor3.Color
    End If
    If m_pMVDay.BackgroundDarkColor <> ctrlThemeColor4.Color Then
        m_pMVDay.BackgroundDarkColor = ctrlThemeColor4.Color
    End If
    If m_pMVDay.BackgroundSelectedColor <> ctrlThemeColor5.Color Then
        m_pMVDay.BackgroundSelectedColor = ctrlThemeColor5.Color
    End If
    
    If chkUseOffice2003HeaderFormat.Enabled Then
        If m_pMVDay.UseOffice2003HeaderFormat <> (chkUseOffice2003HeaderFormat.Value = 1) Then
              m_pMVDay.UseOffice2003HeaderFormat = IIf(chkUseOffice2003HeaderFormat.Value = 1, True, False)
        End If
    End If
    
End Sub

Public Property Get UseOffice2003HeaderFormatVisible() As Boolean
    UseOffice2003HeaderFormatVisible = chkUseOffice2003HeaderFormat.Enabled
End Property

Public Property Let UseOffice2003HeaderFormatVisible(bValue As Boolean)
    chkUseOffice2003HeaderFormat.Enabled = bValue
    chkUseOffice2003HeaderFormat.Visible = bValue
End Property


