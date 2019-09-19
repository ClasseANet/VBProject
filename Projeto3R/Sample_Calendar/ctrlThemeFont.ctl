VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl ctrlThemeFont 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   ScaleHeight     =   435
   ScaleWidth      =   3120
   Begin VB.CommandButton btnFont 
      Caption         =   "..."
      Height          =   315
      Left            =   2700
      TabIndex        =   0
      Top             =   60
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   1680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label txtFont 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   2655
   End
End
Attribute VB_Name = "ctrlThemeFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pFont As New StdFont

Public Sub SetData(ByRef pFontRef As StdFont)
    Debug.Assert Not pFontRef Is Nothing
    
    CopyFont m_pFont, pFontRef
        
    txtFont.Caption = CLng(m_pFont.SIZE) & " pt. " & m_pFont.Name
    txtFont.Font.Bold = m_pFont.Bold
    txtFont.Font.Italic = m_pFont.Italic
    txtFont.Font.Strikethrough = m_pFont.Strikethrough
    txtFont.Font.Underline = m_pFont.Underline
       
End Sub

Public Sub UpdateData(ByRef pFontRef As StdFont)
    If AreFontsDifferent(m_pFont, pFontRef) Then
        CopyFont pFontRef, m_pFont
    End If
    
End Sub

Private Sub btnFont_Click()
    dlgFont.Flags = cdlCFBoth + cdlCFEffects + cdlCFForceFontExist
    
'   dlgFont.Color = ctrlColor.BackColor
    
    dlgFont.FontBold = m_pFont.Bold
    dlgFont.FontItalic = m_pFont.Italic
    dlgFont.FontName = m_pFont.Name
    dlgFont.FontSize = m_pFont.SIZE
    dlgFont.FontStrikethru = m_pFont.Strikethrough
    dlgFont.FontUnderline = m_pFont.Underline
    
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

End Sub
