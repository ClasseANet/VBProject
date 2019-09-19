VERSION 5.00
Begin VB.UserControl ctrlThemeBase 
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   ScaleHeight     =   1530
   ScaleWidth      =   4530
   Begin CalendarSample.ctrlThemeFont ctrlThemeFont1 
      Height          =   435
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   3135
      _extentx        =   5530
      _extenty        =   767
   End
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor1 
      Height          =   435
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _extentx        =   4683
      _extenty        =   767
   End
   Begin CalendarSample.ctrlThemeFont ctrlThemeFont2 
      Height          =   435
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   3135
      _extentx        =   5530
      _extenty        =   767
   End
   Begin VB.Label Label3 
      Caption         =   "Base Font"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Base Bold Font "
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Base Color"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "ctrlThemeBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pTheme As CalendarThemeOffice2007

Public Sub SetData(pTheme As CalendarThemeOffice2007)
    Debug.Assert Not pTheme Is Nothing
    
    Set m_pTheme = pTheme
    
    ctrlThemeColor1.Color = m_pTheme.BaseColor
    
    ctrlThemeFont1.SetData m_pTheme.BaseFont
    ctrlThemeFont2.SetData m_pTheme.BaseFontBold
       
End Sub

Public Sub UpdateData()
    If m_pTheme.BaseColor <> ctrlThemeColor1.Color Then
        m_pTheme.BaseColor = ctrlThemeColor1.Color
    End If
    
    ctrlThemeFont1.UpdateData m_pTheme.BaseFont
    ctrlThemeFont2.UpdateData m_pTheme.BaseFontBold
End Sub

