VERSION 5.00
Begin VB.UserControl ctrlTheme2007Event 
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   ScaleHeight     =   4920
   ScaleWidth      =   4575
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
      TabIndex        =   2
      Top             =   420
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
   End
   Begin VB.Label Label2 
      Caption         =   "Gripper Border Color"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Gripper Background  Color"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   1995
   End
End
Attribute VB_Name = "ctrlTheme2007Event"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pEvent As CalendarThemeOffice2007Event

Public Sub SetData(pEvent As CalendarThemeOffice2007Event)
    Debug.Assert Not pEvent Is Nothing
    
    Set m_pEvent = pEvent
    
    ctrlThemeColor1.Color = m_pEvent.GripperBorderColor
    ctrlThemeColor2.Color = m_pEvent.GripperBackgroundColor
    
    
   
End Sub

Public Sub UpdateData()
    
    If m_pEvent.GripperBorderColor <> ctrlThemeColor1.Color Then
        m_pEvent.GripperBorderColor = ctrlThemeColor1.Color
    End If

    If m_pEvent.GripperBackgroundColor <> ctrlThemeColor2.Color Then
        m_pEvent.GripperBackgroundColor = ctrlThemeColor2.Color
    End If
        
End Sub
