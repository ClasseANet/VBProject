VERSION 5.00
Begin VB.UserControl ctrlThemeHeaderText 
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ScaleHeight     =   5010
   ScaleWidth      =   3990
   Begin VB.Frame Frame4 
      Caption         =   "Today Selected"
      Height          =   1155
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   3780
      Width           =   3855
      Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor4 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1508
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Today"
      Height          =   1155
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   2580
      Width           =   3855
      Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor3 
         Height          =   855
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1508
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selected"
      Height          =   1155
      Left            =   60
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
      Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor2 
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1508
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Normal"
      Height          =   1155
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3855
      Begin CalendarSample.ctrlThemeFontColor ctrlThemeFontColor1 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1508
      End
   End
End
Attribute VB_Name = "ctrlThemeHeaderText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pHeaderText As CalendarThemeHeaderText

Public Sub SetData(pHeaderText As CalendarThemeHeaderText)
    Debug.Assert Not pHeaderText Is Nothing
    
    Set m_pHeaderText = pHeaderText
    
    ctrlThemeFontColor1.SetData m_pHeaderText.Normal
    ctrlThemeFontColor2.SetData m_pHeaderText.Selected
    ctrlThemeFontColor3.SetData m_pHeaderText.Today
    ctrlThemeFontColor4.SetData m_pHeaderText.TodaySelected
        
End Sub

Public Sub UpdateData()
    ctrlThemeFontColor1.UpdateData
    ctrlThemeFontColor2.UpdateData
    ctrlThemeFontColor3.UpdateData
    ctrlThemeFontColor4.UpdateData
End Sub

