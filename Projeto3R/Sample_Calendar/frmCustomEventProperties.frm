VERSION 5.00
Begin VB.Form frmCustomEventProperties 
   Caption         =   "Custom Event Properties"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   LinkTopic       =   "Form2"
   ScaleHeight     =   2970
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox editPropName 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox editPropValue 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox editPropName 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox editPropValue 
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox editPropName 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox editPropValue 
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox editPropName 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox editPropValue 
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox editPropName 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox editPropValue 
      Height          =   285
      Index           =   4
      Left            =   2760
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Property name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmCustomEventProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ModifiedEvent As CalendarEvent

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Long
    Dim propName
    
    ModifiedEvent.CustomProperties.RemoveAll
    
    For i = 0 To 4
        propName = Trim(editPropName(i).Text)
                
        If Len(propName) > 0 Then
            ModifiedEvent.dsr.Property(propName) = editPropValue(i).Text
        End If
    Next
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 0 To 4
        editPropName(i).Text = ""
        editPropValue(i).Text = ""
    Next
    
    frmMain.ModalFormsRunningCounter = frmMain.ModalFormsRunningCounter + 1
End Sub

Public Sub SetEvent(ModEvent As CalendarEvent)
    If ModEvent Is Nothing Then
        Exit Sub
    End If
    
    Set ModifiedEvent = ModEvent
    
    Dim strPropName
    Dim i As Long
    i = 0
    For Each strPropName In ModEvent.CustomProperties
        editPropName(i).Text = strPropName
        editPropValue(i).Text = ModEvent.CustomProperties(strPropName)
        
        i = i + 1
        If i > 4 Then
            Exit For
        End If
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.ModalFormsRunningCounter = frmMain.ModalFormsRunningCounter - 1
End Sub
