VERSION 5.00
Begin VB.UserControl ctrlThemeHeightFormula 
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   ScaleHeight     =   720
   ScaleWidth      =   4425
   Begin VB.TextBox txtConstant 
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   300
      Width           =   675
   End
   Begin VB.TextBox txtDivisor 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   300
      Width           =   615
   End
   Begin VB.TextBox txtMultiplier 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "FontHeight *"
      Height          =   195
      Left            =   900
      TabIndex        =   7
      Top             =   60
      Width           =   975
   End
   Begin VB.Label labelHeight 
      Caption         =   "Height ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Constant"
      Height          =   195
      Left            =   3600
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Divisor  +"
      Height          =   195
      Left            =   2820
      TabIndex        =   2
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Multiplier  /"
      Height          =   195
      Left            =   1920
      TabIndex        =   0
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "ctrlThemeHeightFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pHFormula As CalendarThemeFormulaMulDivC

Public Sub SetData(pHFormula As CalendarThemeFormulaMulDivC)

    Debug.Assert Not pHFormula Is Nothing
    
    Set m_pHFormula = pHFormula
    
    txtMultiplier = m_pHFormula.Multiplier
    txtDivisor = m_pHFormula.Divisor
    txtConstant = m_pHFormula.Constant
    
End Sub

Public Sub UpdateData()

    Dim nMul As Long, nDiv As Long, nConst As Long
    
    nMul = Val(txtMultiplier)
    nDiv = IIf(Val(txtDivisor) = 0, 1, Val(txtDivisor))
    nConst = Val(txtConstant)
    
    If nMul <> m_pHFormula.Multiplier Then
        m_pHFormula.Multiplier = nMul
    End If
    
    If nDiv <> m_pHFormula.Divisor Then
        m_pHFormula.Divisor = nDiv
    End If
    
    If nConst <> m_pHFormula.Constant Then
        m_pHFormula.Constant = nConst
    End If

End Sub

Public Property Get label1text() As String
    label1text = labelHeight
End Property

Public Property Let label1text(newText As String)
    labelHeight = newText
End Property
