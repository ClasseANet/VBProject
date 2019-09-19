VERSION 5.00
Begin VB.UserControl ctrlThemeMVHeaderW 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Frame Frame1 
      Height          =   1035
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   1740
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
   Begin VB.Frame Frame1 
      Caption         =   "Text"
      Height          =   1155
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   480
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
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor1 
      Height          =   435
      Left            =   1200
      TabIndex        =   4
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor2 
      Height          =   435
      Left            =   1500
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   767
   End
   Begin VB.Label lblFreeSpace 
      Caption         =   "Free Space Color"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   3060
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Base Color"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "ctrlThemeMVHeaderW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pMVHeaderW As Object

Private Function IsWeekHeader(pMVHeaderW As Object) As Boolean
    On Error Resume Next
    IsWeekHeader = False
      
    Dim pHdrWeek As CalendarThemeOffice2007MonthViewWeekHeader
    Set pHdrWeek = pMVHeaderW
    
    IsWeekHeader = Not pHdrWeek Is Nothing
    
End Function

Public Sub SetData(pMVHeaderW As Object)
    Debug.Assert Not pMVHeaderW Is Nothing
        
    Set m_pMVHeaderW = pMVHeaderW
    
    ctrlThemeColor1.Color = pMVHeaderW.BaseColor
    
    ctrlThemeFontColor1.SetData pMVHeaderW.TextCenter.Normal
    
    If IsWeekHeader(pMVHeaderW) Then
        ctrlThemeHeightFormula1.label1text = "Width"
        ctrlThemeHeightFormula1.SetData pMVHeaderW.WidthFormula
        
        ctrlThemeFontColor1.FontChangeEnabled = False
        
        ctrlThemeColor2.Color = pMVHeaderW.FreeSpaceBackgroundColor
        ctrlThemeColor2.Visible = True
        lblFreeSpace.Visible = True
        
        
        
    Else
        ctrlThemeHeightFormula1.label1text = "Height"
        ctrlThemeHeightFormula1.SetData pMVHeaderW.HeightFormula
        ctrlThemeFontColor1.FontChangeEnabled = True
        
        ctrlThemeColor2.Visible = False
        lblFreeSpace.Visible = False
    End If
            
End Sub

Public Sub UpdateData()
    
    If m_pMVHeaderW.BaseColor <> ctrlThemeColor1.Color Then
        m_pMVHeaderW.BaseColor = ctrlThemeColor1.Color
    End If
    
    If ctrlThemeColor2.Visible Then
        If m_pMVHeaderW.FreeSpaceBackgroundColor <> ctrlThemeColor2.Color Then
            m_pMVHeaderW.FreeSpaceBackgroundColor = ctrlThemeColor2.Color
        End If
    End If

    ctrlThemeFontColor1.UpdateData
    ctrlThemeHeightFormula1.UpdateData
    
End Sub

