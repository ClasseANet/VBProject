VERSION 5.00
Begin VB.UserControl ctrlTheme2007EventFC 
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ScaleHeight     =   4755
   ScaleWidth      =   6915
   Begin VB.Frame Frame6 
      Caption         =   "Location Text"
      Height          =   1095
      Left            =   60
      TabIndex        =   10
      Top             =   1260
      Width           =   3795
      Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor2 
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1508
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Body Text"
      Height          =   1095
      Left            =   60
      TabIndex        =   8
      Top             =   2460
      Width           =   3795
      Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor3 
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1508
      End
   End
   Begin VB.Frame ctrlFrameStartEnd 
      Caption         =   "Start/End Text"
      Height          =   1095
      Left            =   60
      TabIndex        =   6
      Top             =   3660
      Visible         =   0   'False
      Width           =   3795
      Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor4 
         Height          =   795
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1402
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Subject Text"
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   3795
      Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor1 
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1508
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Backgroung Gradient Color"
      Height          =   1095
      Left            =   3960
      TabIndex        =   2
      Top             =   60
      Width           =   2895
      Begin CalendarSample.ctrlThemeColorGradient ctrlThemeColorGradient1 
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Border color"
      Height          =   675
      Left            =   3960
      TabIndex        =   0
      Top             =   1260
      Width           =   2895
      Begin CalendarSample.ctrlThemeColor ctrlThemeColor1 
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
      End
   End
End
Attribute VB_Name = "ctrlTheme2007EventFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pEventFC As CalendarThemeOffice2007EventFontsColors

Public Sub SetData(pEventFC As CalendarThemeOffice2007EventFontsColors)
    Debug.Assert Not pEventFC Is Nothing
    
    Set m_pEventFC = pEventFC
    
    ctrlThemeFontColor1.SetData m_pEventFC.Subject
    ctrlThemeFontColor2.SetData m_pEventFC.Location
    ctrlThemeFontColor3.SetData m_pEventFC.Body
    
    If Not m_pEventFC.StartEnd.Font Is Nothing Then
        ctrlThemeFontColor4.SetData m_pEventFC.StartEnd
        VisibleStartEnd = True
    Else
        VisibleStartEnd = False
    End If
    
    ctrlThemeColorGradient1.SetData m_pEventFC.Background
    
    ctrlThemeColor1.Color = m_pEventFC.BorderColor
    
       
End Sub

Public Sub UpdateData()
    
    ctrlThemeFontColor1.UpdateData
    ctrlThemeFontColor2.UpdateData
    ctrlThemeFontColor3.UpdateData
    
    If VisibleStartEnd Then
        ctrlThemeFontColor4.UpdateData
    End If
    
    ctrlThemeColorGradient1.UpdateData
    
    If m_pEventFC.BorderColor <> ctrlThemeColor1.Color Then
        m_pEventFC.BorderColor = ctrlThemeColor1.Color
    End If

End Sub

Property Get VisibleStartEnd() As Boolean
    VisibleStartEnd = ctrlFrameStartEnd.Visible
End Property

Property Let VisibleStartEnd(bVisible As Boolean)
    ctrlFrameStartEnd.Visible = bVisible
End Property
