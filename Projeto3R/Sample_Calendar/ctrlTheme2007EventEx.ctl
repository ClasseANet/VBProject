VERSION 5.00
Begin VB.UserControl ctrlTheme2007EventEx 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Frame Frame1 
      Caption         =   "Event height formula"
      Height          =   1035
      Left            =   0
      TabIndex        =   2
      Top             =   1020
      Width           =   4515
      Begin CalendarSample.ctrlThemeHeightFormula ctrlThemeHeightFormula1 
         Height          =   675
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1191
      End
   End
   Begin CalendarSample.ctrlThemeEventIcons ctrlThemeEventIconsToDraw1 
      Height          =   1395
      Left            =   -60
      TabIndex        =   1
      Top             =   2040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2461
   End
   Begin CalendarSample.ctrlTheme2007Event ctrlTheme2007Event1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _extentx        =   8070
      _extenty        =   1720
   End
End
Attribute VB_Name = "ctrlTheme2007EventEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pEvent As Object

Public Sub SetData(pEvent As Object)
    Debug.Assert Not pEvent Is Nothing
    
    Set m_pEvent = pEvent
        
    ctrlTheme2007Event1.SetData m_pEvent
    
    ctrlThemeEventIconsToDraw1.SetData m_pEvent.EventIconsToDraw
    ctrlThemeHeightFormula1.SetData m_pEvent.HeightFormula
       
End Sub

Public Sub UpdateData()
    
    ctrlTheme2007Event1.UpdateData
    
    ctrlThemeEventIconsToDraw1.UpdateData
    ctrlThemeHeightFormula1.UpdateData
    
End Sub

