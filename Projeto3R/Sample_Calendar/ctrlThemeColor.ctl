VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl ctrlThemeColor 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ScaleHeight     =   450
   ScaleWidth      =   2160
   Begin VB.CommandButton btnColor 
      Caption         =   "..."
      Height          =   315
      Left            =   1740
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.TextBox txtColor 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   1200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label ctrlColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "ctrlThemeColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  


Property Get Color() As OLE_COLOR
    Color = ctrlColor.BackColor
End Property

Property Let Color(clrColor As OLE_COLOR)
    ctrlColor.BackColor = clrColor
    txtColor.Text = ColorToStr(ctrlColor.BackColor)
End Property

Private Sub btnColor_Click()
    dlgColor.Flags = cdlCCRGBInit
    dlgColor.Color = ctrlColor.BackColor
    
    dlgColor.ShowColor
    
    ctrlColor.BackColor = dlgColor.Color
    txtColor.Text = ColorToStr(ctrlColor.BackColor)

End Sub

