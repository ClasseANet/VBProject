VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl ctrlThemeFontColor 
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   ScaleHeight     =   855
   ScaleWidth      =   3600
   Begin VB.CommandButton btnFont 
      Caption         =   "..."
      Height          =   315
      Left            =   3180
      TabIndex        =   6
      Top             =   60
      Width           =   375
   End
   Begin VB.TextBox txtColor 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   480
      Width           =   1275
   End
   Begin VB.CommandButton btnColor 
      Caption         =   "..."
      Height          =   315
      Left            =   2220
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   2640
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   3120
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Font"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   435
   End
   Begin VB.Label txtFont 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   60
      Width           =   2655
   End
   Begin VB.Label ctrlColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Color"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   435
   End
End
Attribute VB_Name = "ctrlThemeFontColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pFontColor As CalendarThemeFontColor

Private m_pFont As New StdFont

Public Sub SetData(pFontColor As CalendarThemeFontColor)
    Debug.Assert Not pFontColor Is Nothing
    
    Set m_pFontColor = pFontColor
    
    CopyFont m_pFont, m_pFontColor.Font
    txtFont.Caption = CLng(m_pFont.SIZE) & " pt. " & m_pFont.Name
    
    ctrlColor.BackColor = m_pFontColor.Color
    txtColor.Text = ColorToStr(ctrlColor.BackColor)
    txtFont.Font.Bold = m_pFont.Bold
    txtFont.Font.Italic = m_pFont.Italic
    txtFont.Font.Strikethrough = m_pFont.Strikethrough
    txtFont.Font.Underline = m_pFont.Underline
       
End Sub

Public Sub UpdateData()
    If AreFontsDifferent(m_pFont, m_pFontColor.Font) Then
        CopyFont m_pFontColor.Font, m_pFont
    End If
    
    If ctrlColor.BackColor <> m_pFontColor.Color Then
        m_pFontColor.Color = ctrlColor.BackColor
    End If
End Sub


Private Sub btnColor_Click()
    dlgColor.Flags = cdlCCRGBInit
    dlgColor.Color = ctrlColor.BackColor
    
    dlgColor.ShowColor
    
    ctrlColor.BackColor = dlgColor.Color
    txtColor.Text = ColorToStr(ctrlColor.BackColor)

End Sub

Private Sub btnFont_Click()
    dlgFont.Flags = cdlCFBoth + cdlCFEffects + cdlCFForceFontExist
    
    dlgFont.Color = ctrlColor.BackColor
    
    dlgFont.FontBold = m_pFont.Bold
    dlgFont.FontItalic = m_pFont.Italic
    dlgFont.FontName = m_pFont.Name
    dlgFont.FontSize = m_pFont.SIZE
    dlgFont.FontStrikethru = m_pFont.Strikethrough
    dlgFont.FontUnderline = m_pFont.Underline
    
    On Error GoTo mmExit:
    dlgFont.ShowFont
        
    m_pFont.Bold = dlgFont.FontBold
    m_pFont.Italic = dlgFont.FontItalic
    m_pFont.Name = dlgFont.FontName
    m_pFont.SIZE = dlgFont.FontSize
    m_pFont.Strikethrough = dlgFont.FontStrikethru
    m_pFont.Underline = dlgFont.FontUnderline

    txtFont.Caption = CLng(m_pFont.SIZE) & " pt. " & m_pFont.Name
    txtFont.Font.Bold = m_pFont.Bold
    txtFont.Font.Italic = m_pFont.Italic
    txtFont.Font.Strikethrough = m_pFont.Strikethrough
    txtFont.Font.Underline = m_pFont.Underline
    
    'ctrlColor.BackColor = dlgFont.Color
    'txtColor.Text = ColorToStr(ctrlColor.BackColor)
mmExit:

End Sub

Public Property Get FontChangeEnabled() As Boolean
    FontChangeEnabled = btnFont.Visible

End Property

Public Property Let FontChangeEnabled(bEnabled As Boolean)
    btnFont.Visible = bEnabled
    
End Property

