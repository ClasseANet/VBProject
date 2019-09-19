VERSION 5.00
Begin VB.Form frmOccurrenceSeriesChooser 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2220
      TabIndex        =   6
      Top             =   2580
      Width           =   1275
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   780
      TabIndex        =   5
      Top             =   2580
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   4335
      Begin VB.OptionButton optSeries 
         Caption         =   "Open the series"
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   3795
      End
      Begin VB.OptionButton optOccurrence 
         Caption         =   "Open the Occurrence"
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   3975
      End
   End
   Begin VB.Label txtAction 
      Caption         =   "Do you want to open only this occurrence or the series?"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label txtDescription 
      Caption         =   """Recurrence event - lunch"" is a recurring event."
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmOccurrenceSeriesChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_strEventSubject As String
Public m_bOcurrence As Boolean
Public m_bDeleteRequest As Boolean
Public m_bOK As Boolean


Private Sub btnCancel_Click()
    
    m_bOK = False
    
    Unload Me
End Sub

Private Sub btnOK_Click()
    
    m_bOK = True
    m_bOcurrence = optOccurrence
    
    Unload Me
End Sub

Private Sub Form_Initialize()
    m_bDeleteRequest = False
    m_bOcurrence = True
End Sub

Private Sub Form_Load()
    
    If m_bDeleteRequest Then
        txtAction = "Do you want to delete only this occurrence or the series?"
        optOccurrence.Caption = "Delete the Occurrence"
        optSeries.Caption = "Delete the Series"
    
    Else
        txtAction = "Do you want to open only this occurrence or the series?"
        optOccurrence.Caption = "Open the Occurrence"
        optSeries.Caption = "Open the Series"
    End If
    
    m_bOK = False
    
    txtDescription.Caption = "'" & m_strEventSubject & "'" & " is a recurring event."
    
    optOccurrence = m_bOcurrence
    optSeries = Not m_bOcurrence

End Sub

