VERSION 5.00
Begin VB.Form FrmSelItemSerie 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Abrir Item Recorrente"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optOccurrence 
      Caption         =   "&Avbrir esta ocorrência."
      Height          =   315
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.OptionButton optSeries 
      Caption         =   "Abrir série."
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   2595
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   1275
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label txtAction 
      Caption         =   "Deseja abrir somente esta ocorrência ou série de compromissos?"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label txtDescription 
      Caption         =   "Este item é um compromisso recorrente."
      Height          =   675
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FrmSelItemSerie"
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
        Me.txtAction.Caption = "Do you want to delete only this occurrence or the series?"
        Me.optOccurrence.Caption = "Delete the Occurrence"
        Me.optSeries.Caption = "Delete the Series"
    
        Me.txtAction.Caption = "Deseja excluir somente esta ocorrência ou a série de compromissos?"
        Me.optOccurrence.Caption = "Excluir esta ocorrência."
        Me.optSeries.Caption = "Excluir série."
    Else
        Me.txtAction.Caption = "Do you want to open only this occurrence or the series?"
        Me.optOccurrence.Caption = "Open the Occurrence"
        Me.optSeries.Caption = "Open the Series"
        
        Me.txtAction.Caption = "Deseja abrir somente esta ocorrência ou a série de compromissos?"
        Me.optOccurrence.Caption = "Abrir esta ocorrência."
        Me.optSeries.Caption = "Abrir série."
    End If
    
    m_bOK = False
    
    Me.txtDescription.Caption = "'" & m_strEventSubject & "'" & " is a recurring event."
    Me.txtDescription.Caption = "'" & m_strEventSubject & "'" & " é um compromisso recorrente."
    
    Me.optOccurrence = m_bOcurrence
    Me.optSeries = Not m_bOcurrence
End Sub
