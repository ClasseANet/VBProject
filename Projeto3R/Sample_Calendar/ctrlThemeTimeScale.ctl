VERSION 5.00
Begin VB.UserControl ctrlThemeTimeScale 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ScaleHeight     =   4830
   ScaleWidth      =   4710
   Begin VB.Frame Frame1 
      Caption         =   "Cel Height formula"
      Height          =   1035
      Left            =   60
      TabIndex        =   9
      Top             =   3660
      Width           =   4515
      Begin CalendarSample.ctrlThemeHeightFormula ctrlThemeHeightFormula1 
         Height          =   675
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1191
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Now line"
      Height          =   2115
      Left            =   60
      TabIndex        =   4
      Top             =   1500
      Width           =   3135
      Begin CalendarSample.ctrlThemeColor ctrlThemeColor3 
         Height          =   435
         Left            =   900
         TabIndex        =   6
         Top             =   1560
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin CalendarSample.ctrlThemeColorGradient ctrlThemeColorGradient1 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   120
         X2              =   3000
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         X1              =   120
         X2              =   3000
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Backgrount Gradient color"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Line color"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   795
      End
   End
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor1 
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor2 
      Height          =   435
      Left            =   180
      TabIndex        =   3
      Top             =   1020
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
   End
   Begin VB.Label Label2 
      Caption         =   "Lines color"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Background color"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "ctrlThemeTimeScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pTmeScale As CalendarThemeDayViewTimeScale

Public Sub SetData(pTmeScale As CalendarThemeDayViewTimeScale)
    Debug.Assert Not pTmeScale Is Nothing
    
    Set m_pTmeScale = pTmeScale
    
    ctrlThemeColor1.Color = m_pTmeScale.BackgroundColor
    ctrlThemeColor2.Color = m_pTmeScale.LineColor
    ctrlThemeColor3.Color = m_pTmeScale.NowLineColor
    
    ctrlThemeColorGradient1.SetData m_pTmeScale.NowLineBackground
    ctrlThemeHeightFormula1.SetData m_pTmeScale.HeightFormula
            
End Sub

Public Sub UpdateData()
    If m_pTmeScale.BackgroundColor <> ctrlThemeColor1.Color Then
        m_pTmeScale.BackgroundColor = ctrlThemeColor1.Color
    End If
    
    If m_pTmeScale.LineColor <> ctrlThemeColor2.Color Then
        m_pTmeScale.LineColor = ctrlThemeColor2.Color
    End If
    
    If m_pTmeScale.NowLineColor <> ctrlThemeColor3.Color Then
        m_pTmeScale.NowLineColor = ctrlThemeColor3.Color
    End If
    
    ctrlThemeColorGradient1.UpdateData
    ctrlThemeHeightFormula1.UpdateData
    
End Sub


