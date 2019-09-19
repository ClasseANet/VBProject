VERSION 5.00
Begin VB.UserControl ctrlThemeColorGradient 
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   ScaleHeight     =   900
   ScaleWidth      =   2745
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor1 
      Height          =   435
      Left            =   540
      TabIndex        =   0
      Top             =   0
      Width           =   2235
      _extentx        =   3942
      _extenty        =   767
   End
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor2 
      Height          =   435
      Left            =   540
      TabIndex        =   2
      Top             =   420
      Width           =   2235
      _extentx        =   3942
      _extenty        =   767
   End
   Begin VB.Label Label2 
      Caption         =   "Dark"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Light"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "ctrlThemeColorGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pColorGradient As CalendarThemeColorGradient

Public Sub SetData(pColorGradient As CalendarThemeColorGradient)
    Debug.Assert Not pColorGradient Is Nothing
    
    Set m_pColorGradient = pColorGradient
    
    ctrlThemeColor1.Color = m_pColorGradient.ColorLight
    ctrlThemeColor2.Color = m_pColorGradient.ColorDark
End Sub

Public Sub UpdateData()
    If m_pColorGradient.ColorLight <> ctrlThemeColor1.Color Then
        m_pColorGradient.ColorLight = ctrlThemeColor1.Color
    End If
    If m_pColorGradient.ColorDark <> ctrlThemeColor2.Color Then
        m_pColorGradient.ColorDark = ctrlThemeColor2.Color
    End If
End Sub


