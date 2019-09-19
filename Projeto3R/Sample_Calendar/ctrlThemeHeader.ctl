VERSION 5.00
Begin VB.UserControl ctrlThemeHeader 
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ScaleHeight     =   2550
   ScaleWidth      =   4830
   Begin VB.Frame frameHeight 
      Height          =   1035
      Left            =   60
      TabIndex        =   2
      Top             =   840
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
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor1 
      Height          =   435
      Left            =   1380
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   767
   End
   Begin CalendarSample.ctrlThemeColor ctrlThemeColor2 
      Height          =   435
      Left            =   1380
      TabIndex        =   4
      Top             =   420
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   767
   End
   Begin VB.Label lblColor2 
      Caption         =   "Today Base Color"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Base Color"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "ctrlThemeHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ShowHeightFormula As Boolean
Public ShowToodayColor As Boolean

Private m_pHeader As Object

Private Sub UserControl_Show()
    
    lblColor2.Visible = ShowToodayColor
    ctrlThemeColor2.Visible = ShowToodayColor
        
    frameHeight.Visible = ShowHeightFormula
    
    If ShowToodayColor Then
        frameHeight.Top = ctrlThemeColor2.Top + ctrlThemeColor2.Height + 10
    Else
        frameHeight.Top = ctrlThemeColor1.Top + ctrlThemeColor1.Height + 10
    End If
    
End Sub

Public Sub SetData(pHeader As Object)
    Debug.Assert Not pHeader Is Nothing
        
    Set m_pHeader = pHeader
    
    ctrlThemeColor1.Color = m_pHeader.BaseColor
    
    If ShowToodayColor Then
        ctrlThemeColor2.Color = m_pHeader.TodayBaseColor
    End If
        
    If ShowHeightFormula Then
        ctrlThemeHeightFormula1.SetData m_pHeader.HeightFormula
    End If
End Sub

Public Sub UpdateData()
    
    If m_pHeader.BaseColor <> ctrlThemeColor1.Color Then
        m_pHeader.BaseColor = ctrlThemeColor1.Color
    End If
    
    If ShowToodayColor Then
        If m_pHeader.TodayBaseColor <> ctrlThemeColor2.Color Then
            m_pHeader.TodayBaseColor = ctrlThemeColor2.Color
        End If
    End If
    
    If ShowHeightFormula Then
        ctrlThemeHeightFormula1.UpdateData
    End If
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ShowHeightFormula", ShowHeightFormula
    PropBag.WriteProperty "ShowToodayColor", ShowToodayColor
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ShowHeightFormula = PropBag.ReadProperty("ShowHeightFormula", False)
    ShowToodayColor = PropBag.ReadProperty("ShowToodayColor", False)
End Sub

