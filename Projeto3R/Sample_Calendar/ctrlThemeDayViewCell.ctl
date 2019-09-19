VERSION 5.00
Begin VB.UserControl ctrlThemeDayViewCell 
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   ScaleHeight     =   4620
   ScaleWidth      =   4605
   Begin VB.Frame Frame2 
      Caption         =   "Work cell"
      Height          =   2175
      Left            =   60
      TabIndex        =   2
      Top             =   2340
      Width           =   4455
      Begin CalendarSample.ctrlThemeDVCellColors ctrlThemeDayViewCellColors2 
         Height          =   1875
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3307
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Non-work cell"
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      Begin CalendarSample.ctrlThemeDVCellColors ctrlThemeDayViewCellColors1 
         Height          =   1875
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3307
      End
   End
End
Attribute VB_Name = "ctrlThemeDayViewCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pDayViewCell As CalendarThemeDayViewCell

Public Sub SetData(pDayViewCell As CalendarThemeDayViewCell)
    Debug.Assert Not pDayViewCell Is Nothing
    
    Set m_pDayViewCell = pDayViewCell
    
    ctrlThemeDayViewCellColors1.SetData m_pDayViewCell.NonWorkCell
    ctrlThemeDayViewCellColors2.SetData m_pDayViewCell.WorkCell
       
End Sub

Public Sub UpdateData()
    
    ctrlThemeDayViewCellColors1.UpdateData
    ctrlThemeDayViewCellColors2.UpdateData
    
End Sub

