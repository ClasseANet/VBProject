VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{71971158-2541-45FB-B54B-CB029D892011}#3.0#0"; "HtmlEditor.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmEditorH 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   855
      Left            =   6600
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
      ExtentX         =   2143
      ExtentY         =   1508
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin RichTextLib.RichTextBox rtfSource 
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1296
      _Version        =   393217
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmEditorH.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VisualHtmlEditor.HtmlEditor HtmlEditor1 
      Height          =   1095
      Left            =   2820
      TabIndex        =   7
      Top             =   1500
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1931
      LiveResize      =   0   'False
      MultipleSelection=   0   'False
      ShowBorders     =   0   'False
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   5475
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16757
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":0084
            Key             =   "Bold"
            Object.Tag             =   "Bold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":025E
            Key             =   "Underline"
            Object.Tag             =   "Underline"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":0438
            Key             =   "Italic"
            Object.Tag             =   "Italic"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":0612
            Key             =   "LeftJustify"
            Object.Tag             =   "LeftJustify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":07EC
            Key             =   "RightJustify"
            Object.Tag             =   "RightJustify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":09C6
            Key             =   "CenterJustify"
            Object.Tag             =   "CenterJustify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":0BA0
            Key             =   "FullJustify"
            Object.Tag             =   "FullJustify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":0D7A
            Key             =   "Bullets"
            Object.Tag             =   "Bullets"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":0F54
            Key             =   "Numbers"
            Object.Tag             =   "Numbers"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":112E
            Key             =   "Indent"
            Object.Tag             =   "Indent"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1308
            Key             =   "Outdent"
            Object.Tag             =   "Outdent"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":14E2
            Key             =   "LTR"
            Object.Tag             =   "LTR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":16BC
            Key             =   "SubScript"
            Object.Tag             =   "SubScript"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1816
            Key             =   "SuperScript"
            Object.Tag             =   "SuperScript"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1970
            Key             =   "StrikeThrough"
            Object.Tag             =   "StrikeThrough"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1B4A
            Key             =   "RTL"
            Object.Tag             =   "RTL"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1D24
            Key             =   "ForeColor1"
            Object.Tag             =   "ForeColor1"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1E7E
            Key             =   "ForeColor"
            Object.Tag             =   "ForeColor"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":2572
            Key             =   "BackColor"
            Object.Tag             =   "BackColor"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   540
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":2B0C
            Key             =   "WebFile"
            Object.Tag             =   "WebFile"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":8DA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1020
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   600
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   63
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":8F00
            Key             =   "New2"
            Object.Tag             =   "New2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":90DA
            Key             =   "Open2"
            Object.Tag             =   "Open2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":92B4
            Key             =   "Save2"
            Object.Tag             =   "Save2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":948E
            Key             =   "SaveAs1"
            Object.Tag             =   "SaveAs1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":99E0
            Key             =   "Print"
            Object.Tag             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":9BBA
            Key             =   "Preview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":9D94
            Key             =   "Spell"
            Object.Tag             =   "Spell"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":9F6E
            Key             =   "Cut1"
            Object.Tag             =   "Cut1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":A148
            Key             =   "Copy1"
            Object.Tag             =   "Copy1"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":A322
            Key             =   "Paste1"
            Object.Tag             =   "Paste1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":A4FC
            Key             =   "Undo1"
            Object.Tag             =   "Undo1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":A656
            Key             =   "Redo1"
            Object.Tag             =   "Redo1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":A7B0
            Key             =   "Table1"
            Object.Tag             =   "Table1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":A98A
            Key             =   "Image2"
            Object.Tag             =   "Image2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":AB64
            Key             =   "Link"
            Object.Tag             =   "Link"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":AD3E
            Key             =   "ShowAll"
            Object.Tag             =   "ShowAll"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":AF18
            Key             =   "DeleteCells"
            Object.Tag             =   "DeleteCells"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":B0F2
            Key             =   "InsertColumns"
            Object.Tag             =   "InsertColumns"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":B2CC
            Key             =   "InsertRows"
            Object.Tag             =   "InsertRows"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":B4A6
            Key             =   "DeleteColumns"
            Object.Tag             =   "DeleteColumns"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":B680
            Key             =   "ShowBorders"
            Object.Tag             =   "ShowBorders"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":B85A
            Key             =   "HideBorders"
            Object.Tag             =   "HideBorders"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":BA34
            Key             =   "ColsEven"
            Object.Tag             =   "ColsEven"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":BC0E
            Key             =   "RowsEven"
            Object.Tag             =   "RowsEven1"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":BDE8
            Key             =   "Download"
            Object.Tag             =   "Download"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":BF42
            Key             =   "MergeCells"
            Object.Tag             =   "MergeCells"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":C11C
            Key             =   "SplitCells"
            Object.Tag             =   "SplitCells"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":C2F6
            Key             =   "Video"
            Object.Tag             =   "Video"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":C4D0
            Key             =   "PageSetup"
            Object.Tag             =   "PageSetup"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":C6AA
            Key             =   "PrintPreview"
            Object.Tag             =   "PrintPreview"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":C884
            Key             =   "Properties"
            Object.Tag             =   "Properties"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":CA5E
            Key             =   "Publish"
            Object.Tag             =   "Publish"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":CC38
            Key             =   "WebTransfer"
            Object.Tag             =   "WebTransfer"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":CD92
            Key             =   "Find"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":CEEC
            Key             =   "AlignBottom"
            Object.Tag             =   "AlignBottom"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":D0C6
            Key             =   "AlignTop."
            Object.Tag             =   "AlignTop."
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":D2A0
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":D47A
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":D654
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":D82E
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":DA08
            Key             =   "TextDirection"
            Object.Tag             =   "TextDirection"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":DBE2
            Key             =   "AutoFit"
            Object.Tag             =   "AutoFit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":DDBC
            Key             =   "Comment"
            Object.Tag             =   "Comment"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":DF96
            Key             =   "Website"
            Object.Tag             =   "Website"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":E3E8
            Key             =   "TableAutoFormat"
            Object.Tag             =   "TableAutoFormat"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":E5C2
            Key             =   "Form"
            Object.Tag             =   "Form"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":E79C
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":EE6E
            Key             =   "New"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":F540
            Key             =   "Save1"
            Object.Tag             =   "Save1"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":F71A
            Key             =   "SnapToGrid"
            Object.Tag             =   "SnapToGrid"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":F8F4
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":FFC6
            Key             =   "SaveAs"
            Object.Tag             =   "SaveAs"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":10698
            Key             =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":10D6A
            Key             =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":10EC4
            Key             =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":11596
            Key             =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":116F0
            Key             =   "Redo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1184A
            Key             =   "Replace"
            Object.Tag             =   "Replace"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":11F1C
            Key             =   "Find1"
            Object.Tag             =   "Find1"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":12076
            Key             =   "Image"
            Object.Tag             =   "Image"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":12748
            Key             =   "Table"
            Object.Tag             =   "Table"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":128A2
            Key             =   "FindNext"
            Object.Tag             =   "FindNext"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":129FC
            Key             =   "GoToLine"
            Object.Tag             =   "GoToLine"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1380
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":12B56
            Key             =   "Normal"
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":12D30
            Key             =   "HTML"
            Object.Tag             =   "HTML"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":12F0A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":130E4
            Key             =   "Preview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1323E
            Key             =   "Refresh"
            Object.Tag             =   "Refresh"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":13418
            Key             =   "Back"
            Object.Tag             =   "Back"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":135F2
            Key             =   "Forword"
            Object.Tag             =   "Forword"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":137CC
            Key             =   "Stop"
            Object.Tag             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":139A6
            Key             =   "InsertCells"
            Object.Tag             =   "InsertCells"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":13B80
            Key             =   "InsertColumns"
            Object.Tag             =   "InsertColumns"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":13D5A
            Key             =   "InsertRows2"
            Object.Tag             =   "InsertRows2"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":13F34
            Key             =   "MergeCells"
            Object.Tag             =   "MergeCells"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1410E
            Key             =   "SplitCells"
            Object.Tag             =   "SplitCells"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":142E8
            Key             =   "DeleteCells"
            Object.Tag             =   "DeleteCells"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":144C2
            Key             =   "DeleteColumns"
            Object.Tag             =   "DeleteColumns"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1469C
            Key             =   "DeleteRows"
            Object.Tag             =   "DeleteRows"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":147F6
            Key             =   "InsertRows"
            Object.Tag             =   "InsertRows"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":14950
            Key             =   "PositionAbsolutely"
            Object.Tag             =   "PositionAbsolutely"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":14B2A
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":14D04
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":14EDE
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":150B8
            Key             =   "SendBackward1-delete"
            Object.Tag             =   "SendBackward1"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":15292
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1546C
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1597E
            Key             =   "Textbox"
            Object.Tag             =   "Textbox"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":15B8B
            Key             =   "Textarea"
            Object.Tag             =   "Textarea"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":15D90
            Key             =   "Checkbox"
            Object.Tag             =   "Checkbox"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":15F95
            Key             =   "OptionButton"
            Object.Tag             =   "OptionButton"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1618F
            Key             =   "DropDown"
            Object.Tag             =   "DropDown"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":16398
            Key             =   "PushButton"
            Object.Tag             =   "PushButton"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":16596
            Key             =   "HiddenData"
            Object.Tag             =   "HiddenData"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":16770
            Key             =   "Password"
            Object.Tag             =   "Password"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1694A
            Key             =   "SubmitButton"
            Object.Tag             =   "SubmitButton"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":16B24
            Key             =   "ResetButton"
            Object.Tag             =   "ResetButton"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":16CFE
            Key             =   "ImageButton"
            Object.Tag             =   "ImageButton"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":16ED8
            Key             =   "Form"
            Object.Tag             =   "Form"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":170B2
            Key             =   "BringAboveText"
            Object.Tag             =   "BringAboveText"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":1728C
            Key             =   "SendBelowText"
            Object.Tag             =   "SendBelowText"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":17466
            Key             =   "SnapToGrid"
            Object.Tag             =   "SnapToGrid"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditorH.frx":17640
            Key             =   "ListBox"
            Object.Tag             =   "ListBox"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   34
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Normal"
            Description     =   "Normal"
            Object.ToolTipText     =   "Normal"
            Object.Tag             =   "Normal"
            ImageKey        =   "Normal"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HTML"
            Description     =   "HTML"
            Object.ToolTipText     =   "HTML"
            Object.Tag             =   "HTML"
            ImageKey        =   "HTML"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Description     =   "Preview"
            Object.ToolTipText     =   "Preview"
            Object.Tag             =   "Preview"
            ImageKey        =   "Preview"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   "Refresh"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            Object.Tag             =   "Stop"
            ImageKey        =   "Stop"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertRows"
            Object.ToolTipText     =   "Insert Rows"
            Object.Tag             =   "Insert Rows"
            ImageKey        =   "InsertRows"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertColumns"
            Object.ToolTipText     =   "Insert Columns"
            Object.Tag             =   "Insert Columns"
            ImageKey        =   "InsertColumns"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertCells"
            Object.ToolTipText     =   "Insert Cells"
            Object.Tag             =   "Insert Cells"
            ImageKey        =   "InsertCells"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteCells"
            Object.ToolTipText     =   "Delete Cells"
            Object.Tag             =   "Delete Cells"
            ImageKey        =   "DeleteCells"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteRows"
            Object.ToolTipText     =   "Delete Rows"
            Object.Tag             =   "DeleteRows"
            ImageKey        =   "DeleteRows"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteColumns"
            Object.ToolTipText     =   "Delete Columns"
            Object.Tag             =   "Delete Columns"
            ImageKey        =   "DeleteColumns"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MergeCells"
            Object.ToolTipText     =   "Merge Cells"
            Object.Tag             =   "Merge Cells"
            ImageKey        =   "MergeCells"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SplitCells"
            Object.ToolTipText     =   "Split Cells"
            Object.Tag             =   "Split Cells"
            ImageKey        =   "SplitCells"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PositionAbsolutely"
            Object.ToolTipText     =   "Position Absolutely"
            Object.Tag             =   "Position Absolutely"
            ImageKey        =   "PositionAbsolutely"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BringToFront"
            Object.ToolTipText     =   "Bring To Front"
            Object.Tag             =   "Bring To Front"
            ImageKey        =   "BringToFront"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SendToBack"
            Object.ToolTipText     =   "Send To Back"
            Object.Tag             =   "Send To Back"
            ImageKey        =   "SendToBack"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Textbox"
            Object.ToolTipText     =   "Textbox"
            Object.Tag             =   "Textbox"
            ImageKey        =   "Textbox"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Textarea"
            Object.ToolTipText     =   "Textarea"
            Object.Tag             =   "Textarea"
            ImageKey        =   "Textarea"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Checkbox"
            Object.ToolTipText     =   "Checkbox"
            Object.Tag             =   "Checkbox"
            ImageKey        =   "Checkbox"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OptionButton"
            Object.ToolTipText     =   "Option Button"
            Object.Tag             =   "Option Button"
            ImageKey        =   "OptionButton"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ListBox"
            Object.ToolTipText     =   "ListBox"
            Object.Tag             =   "ListBox"
            ImageKey        =   "ListBox"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DropDownBox"
            Object.ToolTipText     =   "Drop Down Box"
            Object.Tag             =   "Drop Down Box"
            ImageKey        =   "DropDown"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PushButton"
            Object.ToolTipText     =   "Push Button"
            Object.Tag             =   "Push Button"
            ImageKey        =   "PushButton"
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HiddenData"
            Object.ToolTipText     =   "Hidden Data"
            Object.Tag             =   "Hidden Data"
            ImageKey        =   "HiddenData"
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Password"
            Object.ToolTipText     =   "Password"
            Object.Tag             =   "Password"
            ImageKey        =   "Password"
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SubmitButton"
            Object.ToolTipText     =   "Submit Button"
            Object.Tag             =   "Submit Button"
            ImageKey        =   "SubmitButton"
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ResetButton"
            Object.ToolTipText     =   "Reset Button"
            Object.Tag             =   "Reset Button"
            ImageKey        =   "ResetButton"
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ImageButton"
            Object.ToolTipText     =   "Image Button"
            Object.Tag             =   "Image Button"
            ImageKey        =   "ImageButton"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Style"
            Description     =   "Style"
            Object.ToolTipText     =   "Style"
            Object.Tag             =   "Style"
            Style           =   4
            Object.Width           =   5000
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold"
            Object.ToolTipText     =   "Bold"
            Object.Tag             =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic"
            Object.ToolTipText     =   "Italic"
            Object.Tag             =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline"
            Object.ToolTipText     =   "Underline"
            Object.Tag             =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LeftJustify"
            Description     =   "Left Justify"
            Object.ToolTipText     =   "Left Justify"
            ImageKey        =   "LeftJustify"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "CenterJustify"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RightJustify"
            Description     =   "Right Justify"
            Object.ToolTipText     =   "Right Justify"
            ImageKey        =   "RightJustify"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "JustifyFull"
            Object.ToolTipText     =   "Justify Full"
            Object.Tag             =   "Justify Full"
            ImageKey        =   "FullJustify"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SuperScript"
            Object.ToolTipText     =   "Superscript"
            Object.Tag             =   "Superscript"
            ImageKey        =   "SuperScript"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SubScript"
            Object.ToolTipText     =   "Subscript"
            Object.Tag             =   "Subscript"
            ImageKey        =   "SubScript"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StrikeThrough"
            Object.ToolTipText     =   "Strike through"
            Object.Tag             =   "Strike through"
            ImageKey        =   "StrikeThrough"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Numbers"
            Description     =   "Numbers"
            Object.ToolTipText     =   "Numbers"
            Object.Tag             =   "Numbers"
            ImageKey        =   "Numbers"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bullets"
            Description     =   "Bullets"
            Object.ToolTipText     =   "Bullets"
            ImageKey        =   "Bullets"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Outdent"
            Description     =   "Outdent"
            ImageKey        =   "Outdent"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Indent"
            Description     =   "Indent"
            Object.ToolTipText     =   "Indent"
            ImageKey        =   "Indent"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ForeColor"
            Description     =   "ForeColor"
            Object.ToolTipText     =   "Font Color"
            ImageKey        =   "ForeColor"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BackColor"
            Description     =   "BackColor"
            Object.ToolTipText     =   "Background Color"
            ImageKey        =   "BackColor"
         EndProperty
      EndProperty
      Begin VB.ComboBox StyleCombo 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   0
         Width           =   1650
      End
      Begin VB.ComboBox FontSizeCombo 
         Height          =   315
         Left            =   4140
         TabIndex        =   3
         Text            =   "Combo1"
         ToolTipText     =   "Font size"
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox FontCombo 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Text            =   "FontCombo"
         ToolTipText     =   "Font"
         Top             =   0
         Width           =   2355
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveAs"
            Object.ToolTipText     =   "Save As"
            Object.Tag             =   "Save As"
            ImageKey        =   "SaveAs"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Download"
            Object.ToolTipText     =   "Open URL"
            Object.Tag             =   "Download"
            ImageKey        =   "Download"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print"
            Object.Tag             =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View"
            Description     =   "View"
            Object.ToolTipText     =   "View in Browser"
            Object.Tag             =   "View"
            ImageKey        =   "Preview"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SpellCheck"
            Description     =   "Spell Check"
            Object.ToolTipText     =   "Spell Check"
            Object.Tag             =   "Spell"
            ImageKey        =   "Spell"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            Object.Tag             =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Description     =   "Redo"
            Object.ToolTipText     =   "Redo"
            Object.Tag             =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Table"
            Description     =   "Table"
            Object.ToolTipText     =   "Insert Table"
            Object.Tag             =   "Table"
            ImageKey        =   "Table"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Image"
            Description     =   "Image"
            Object.ToolTipText     =   "Insert Image"
            Object.Tag             =   "Image"
            ImageKey        =   "Image"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Hyperlink"
            Description     =   "Hyperlink"
            Object.ToolTipText     =   "Insert Hyperlink"
            Object.Tag             =   "Hyperlink"
            ImageKey        =   "Link"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Form"
            Object.ToolTipText     =   "Insert Form"
            Object.Tag             =   "Form"
            ImageKey        =   "Form"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowAll"
            Description     =   "ShowAll"
            Object.ToolTipText     =   "Show All"
            Object.Tag             =   "ShowAll"
            ImageKey        =   "ShowAll"
            Style           =   1
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowBorders"
            Description     =   "ShowBorders"
            Object.ToolTipText     =   "Show Borders"
            Object.Tag             =   "ShowBorders"
            ImageKey        =   "ShowBorders"
            Style           =   1
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            Object.Tag             =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmEditorH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Mewsoft Visual Html Editor
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================
Option Explicit
Event Load()
Event mnuFileOpenClick()
Event mnuFileSaveAsClick()
Event mnuFileSaveClick()
Private Sub Command10_Click()
    HtmlEditor1.InsertTable 4, 4
End Sub

Private Sub Command11_Click()
    HtmlEditor1.SetZoom -3
End Sub

Private Sub Command3_Click()
    
    HtmlEditor1.BlockDirLTR
    
End Sub

Private Sub Command4_Click()
    HtmlEditor1.BlockDirRTL
End Sub

Private Sub Command5_Click()
    
    HtmlEditor1.CreateBookmark
    
End Sub

Private Sub Command6_Click()
    'Debug.Print HtmlEditor1.GetDirty
    'HtmlEditor1.SetDirty 1
    'Debug.Print HtmlEditor1.GetDirty
    HtmlEditor1.ShowGrid
End Sub

Private Sub Command7_Click()
    
    'HtmlEditor1.SendBackward
'    Dim cEvent As clsHTMLEvent ' Object to set Element Event to reference
    ' Create new instance of the event object:
'    Set cEvent = New clsHTMLEvent
'    cEvent.Event_Details Me, "btnRegion_Click"
'    frmItemMainInfo.WB.Document.All("btnRegion").onclick = cEvent
'    Set cEvent = Nothing
    
    'HtmlEditor1.UnderlineSelected
    Dim el As IHTMLElement
    Dim elements As New Collection
    Set elements = HtmlEditor1.GetSelectedElements
    For Each el In elements
        Debug.Print "Selected: "; el.tagName
    Next el
    
    
End Sub

Public Sub btnRegion_Click()
    Debug.Print "btnRegion_Click"
End Sub

Private Sub Command8_Click()
    'HtmlEditor1.MergeCells
    'HtmlEditor1.SelectedCells
    HtmlEditor1.SendBackward
    
End Sub

Private Sub Command9_Click()

    HtmlEditor1.InsertHTMLCode "<b>Hello Ahmed</b>"
    
End Sub

Private Sub FontCombo_Change()
    'FontSizeCombo.Text = HtmlEditor1.GetFontSize

End Sub

Private Sub FontCombo_Click()
    HtmlEditor1.SetFontName FontCombo.Text
End Sub

Private Sub FontSizeCombo_Change()
    'FontSizeCombo.Text = HtmlEditor1.GetFontSize
End Sub

Private Sub FontSizeCombo_Click()
    HtmlEditor1.SetFontSize FontSizeCombo.Text
End Sub

Private Sub Form_Load()
   RaiseEvent Load
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    DrawWidth = 8   ' Set DrawWidth.
'    Me.DrawMode = 3
'    Me.Line (1000, 1)-(1000, 50000), vbRed
'    Debug.Print "mouse"
End Sub

Private Sub Form_Resize()

    HtmlEditor1.Move Me.ScaleLeft, 1080, Me.ScaleWidth, Me.ScaleHeight - sbStatusBar.Height - Toolbar1.Height - Toolbar2.Height - Toolbar3.Height - 350
    rtfSource.Move Me.ScaleLeft, HtmlEditor1.Top, Me.ScaleWidth, Me.ScaleHeight - sbStatusBar.Height - Toolbar1.Height - Toolbar2.Height - Toolbar3.Height
    WebBrowser1.Move Me.ScaleLeft, HtmlEditor1.Top, Me.ScaleWidth, Me.ScaleHeight - sbStatusBar.Height - Toolbar1.Height - Toolbar2.Height - Toolbar3.Height
End Sub

Private Sub Command2_Click()
    Debug.Print HtmlEditor1.DocumentHtml
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If

End Sub

Private Sub HtmlEditor1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)

    On Error Resume Next
    '----------------------------------------------------------------
    'Toolbar 1
    UpdateToolBar1Button "Copy", HtmlEditor1.IsCopy
    UpdateToolBar1Button "Cut", HtmlEditor1.IsCut
    UpdateToolBar1Button "Paste", HtmlEditor1.IsPaste
    
    UpdateToolBar1Button "Undo", HtmlEditor1.IsUndo
    UpdateToolBar1Button "Redo", HtmlEditor1.IsRedo
    
    UpdateToolBar1Button "Hyperlink", HtmlEditor1.IsCreateLink
    UpdateToolBar1Button "Image", HtmlEditor1.IsInsertImage
    
    '----------------------------------------------------------------
    If HtmlEditor1.ShowDetails Then
        Toolbar2.buttons("ShowAll").Value = tbrPressed
    Else
        Toolbar2.buttons("ShowAll").Value = tbrUnpressed
    End If
    If HtmlEditor1.ShowDetails Then
        Toolbar2.buttons("ShowBorders").Value = tbrPressed
    Else
        Toolbar2.buttons("ShowBorders").Value = tbrUnpressed
    End If
    '----------------------------------------------------------------
    'Toolbar 2
    
    If HtmlEditor1.GetFormatBlock <> "" Then
        StyleCombo.Text = HtmlEditor1.GetFormatBlock
    End If
    If HtmlEditor1.GetFontName <> "" Then
        FontCombo.Text = HtmlEditor1.GetFontName
    End If
    If HtmlEditor1.GetFontSize > 0 Then
        FontSizeCombo.Text = HtmlEditor1.GetFontSize
    End If
    
    UpdateToolBar2Button "Bold", HtmlEditor1.IsBold
    UpdateToolBar2Button "Italic", HtmlEditor1.IsItalic
    UpdateToolBar2Button "Italic", HtmlEditor1.IsItalic
    
    UpdateToolBar2Button "Center", HtmlEditor1.IsJustifyCenter
    UpdateToolBar2Button "LeftJustify", HtmlEditor1.IsJustifyLeft
    UpdateToolBar2Button "RightJustify", HtmlEditor1.IsJustifyRight
    UpdateToolBar2Button "JustifyFull", HtmlEditor1.IsJustifyFull
    
    UpdateToolBar2Button "Outdent", HtmlEditor1.IsOutdent
    UpdateToolBar2Button "Indent", HtmlEditor1.IsIndent
    
    UpdateToolBar2Button "SuperScript", HtmlEditor1.IsSuperScript
    UpdateToolBar2Button "SubScript", HtmlEditor1.IsSubScript
    '----------------------------------------------------------------
    'Toolbar 3
    UpdateToolBar3Button "PositionAbsolutely", HtmlEditor1.IsAbsolutePosition
    UpdateToolBar3ButtonEnabled "PositionAbsolutely", HtmlEditor1.IsAbsolutePositionEnabled
    
    UpdateToolBar3ButtonEnabled "PushButton", HtmlEditor1.IsInsertInputButton
    UpdateToolBar3ButtonEnabled "Textbox", HtmlEditor1.IsInsertInputText
    'IsInsertInputSubmit
    UpdateToolBar3ButtonEnabled "SubmitButton", HtmlEditor1.IsInsertInputSubmit
    'IsInsertInputReset
    UpdateToolBar3ButtonEnabled "ResetButton", HtmlEditor1.IsInsertInputReset
    'IsInsertInputRadio
    UpdateToolBar3ButtonEnabled "OptionButton", HtmlEditor1.IsInsertInputRadio
    'IsInsertInputPassword
    UpdateToolBar3ButtonEnabled "Password", HtmlEditor1.IsInsertInputPassword
    'IsInsertInputImage
    UpdateToolBar3ButtonEnabled "ImageButton", HtmlEditor1.IsInsertInputImage
    'IsInsertInputHidden
    UpdateToolBar3ButtonEnabled "HiddenData", HtmlEditor1.IsInsertInputHidden
    'IsInsertInputFileUpload
    UpdateToolBar3ButtonEnabled "Textbox", HtmlEditor1.IsInsertInputFileUpload
    'IsInsertInputCheckbox
    UpdateToolBar3ButtonEnabled "Checkbox", HtmlEditor1.IsInsertInputCheckbox
    'IsInsertSelectListbox
    UpdateToolBar3ButtonEnabled "ListBox", HtmlEditor1.IsInsertSelectListbox
    'IsInsertTextArea
    UpdateToolBar3ButtonEnabled "Textarea", HtmlEditor1.IsInsertTextArea
    'IsInsertSelectDropdown
    UpdateToolBar3ButtonEnabled "DropDownBox", HtmlEditor1.IsInsertSelectDropdown
    
    
    '----------------------------------------------------------------

End Sub

Private Sub UpdateToolBar1Button(ButtonName As String, Status As Boolean)

    If Status Then
        Toolbar1.buttons(ButtonName).Enabled = True
    Else
        Toolbar1.buttons(ButtonName).Enabled = False
    End If

End Sub

Private Sub UpdateToolBar2Button(ButtonName As String, Status As Boolean)

    If Status Then
        Toolbar2.buttons(ButtonName).Value = tbrPressed
    Else
        Toolbar2.buttons(ButtonName).Value = tbrUnpressed
    End If

End Sub


Private Sub UpdateToolBar3Button(ButtonName As String, Status As Boolean)

    If Status Then
        Toolbar3.buttons(ButtonName).Value = tbrPressed
    Else
        Toolbar3.buttons(ButtonName).Value = tbrUnpressed
    End If

End Sub

Private Sub UpdateToolBar3ButtonEnabled(ButtonName As String, Status As Boolean)

    If Status Then
        Toolbar3.buttons(ButtonName).Enabled = True
    Else
        Toolbar3.buttons(ButtonName).Enabled = False
    End If

End Sub

Private Sub HtmlEditor1_UpdatePageStatus(ByVal pDisp As Object, nPage As Variant, fDone As Variant)

    Debug.Print "UpdatePageStatus: "

End Sub

Private Sub StyleCombo_Change()
    HtmlEditor1.FormatBlock StyleCombo.Text
End Sub

Private Sub StyleCombo_Click()
    HtmlEditor1.FormatBlock StyleCombo.Text
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    
    Dim State As Boolean
    ' Handle toolbar commands
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
            
        Case "Open"
            mnuFileOpen_Click
            
        Case "Save"
            mnuFileSave_Click
        
        Case "SaveAs"
            mnuFileSaveAs_Click
            
        Case "Download"
            'mnuFileNew_Click
            
            CenterForm frmOpenURL, Me
            frmOpenURL.Show vbModal, Me
            
        Case "SpellCheck"
            
        Case "Undo"
            HtmlEditor1.Undo
        
        Case "Redo"
            HtmlEditor1.Redo
        
        Case "Cut"
            HtmlEditor1.Cut
        
        Case "Copy"
            HtmlEditor1.Copy
        
        Case "Paste"
            HtmlEditor1.Paste
        
        Case "Find"
        
        ' Insert Table
        Case "Table"
            HtmlEditor1.InsertTable 3, 4
        
        Case "Image"
            HtmlEditor1.InsertImage
            
        Case "Hyperlink"
            'HtmlEditor1.insert
    
        Case "ShowAll"
            HtmlEditor1.ShowDetails = Not HtmlEditor1.ShowDetails
            
        Case "ShowBorders"
            HtmlEditor1.ShowBorders = Not HtmlEditor1.ShowBorders
            
        Case "Print"
            HtmlEditor1.PrintDocument
            
        Case "Form"
            'HtmlEditor1
            HtmlEditor1.InsertForm
        
        Case "Properties"
        
            
    End Select
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    
    Select Case Button.Key
        Case "Bold"
            HtmlEditor1.Bold
            
        Case "Italic"
            HtmlEditor1.Italic
            
        Case "Underline"
            HtmlEditor1.Underline
            
        Case "Numbers"
            HtmlEditor1.InsertOrderedList
            
        Case "Bullets"
            HtmlEditor1.InsertUnorderedList
            
        Case "Outdent"
            HtmlEditor1.Indent
            
        Case "Indent"
            HtmlEditor1.Outdent
            
        Case "LeftJustify"
            HtmlEditor1.JustifyLeft
            
        Case "Center"
            HtmlEditor1.JustifyCenter
            
        Case "RightJustify"
            HtmlEditor1.JustifyRight
            
        Case "JustifyFull"
            HtmlEditor1.JustifyFull
                    
        Case "SuperScript"
            HtmlEditor1.superScript
        
        Case "SubScript"
            HtmlEditor1.subScript
                    
        Case "StrikeThrough"
            HtmlEditor1.Strikethrough
        
        ' Fore color
        Case "ForeColor"
           
            Dim FColor As String
            'On Error GoTo cleanup
            CommonDialog1.Color = 0
            CommonDialog1.CancelError = True
            CommonDialog1.ShowColor
            FColor = ""
            FColor = FormatRGBString(CommonDialog1.Color)
            HtmlEditor1.SetForeColor FColor
            
        ' Back color
        Case "BackColor"
           
            Dim BColor As String
            'On Error GoTo cleanup
            CommonDialog1.Color = 0
            CommonDialog1.CancelError = True
            CommonDialog1.ShowColor
            BColor = ""
            BColor = FormatRGBString(CommonDialog1.Color)
        
            HtmlEditor1.SetBackColor BColor
            
        End Select
        
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
        
    ' Handle toolbar commands
    Select Case Button.Key
        Case "Normal"
            SwitchToNormal
            
        Case "HTML"
            SwitchToHTML
            
        Case "Preview"
            SwitchToPreview
    
        Case "Refresh"
            
        Case "Stop"
    
        Case "InsertRows"
                HtmlEditor1.InsertRows
        
        Case "InsertColumns":
                HtmlEditor1.InsertColumns
                
        Case "InsertCells"
                HtmlEditor1.InsertCells
                
        Case "DeleteCells"
                HtmlEditor1.DeleteCells
                
        Case "DeleteColumns"
                HtmlEditor1.DeleteColumns
                
        Case "DeleteRows"
                HtmlEditor1.DeleteRows
                
        Case "MergeCells"
            HtmlEditor1.MergeCells
        
        Case "SplitCells"
            HtmlEditor1.SplitSelectedCell
            
        '------------------------------------------------------------
        'PositionAbsolutely
        Case "PositionAbsolutely"
            HtmlEditor1.AbsolutePosition2D
                    
        'BringForward
        Case "BringForward"
            HtmlEditor1.BringForward
        
        'SendBackward
        Case "SendBackward"
        
        'BringToFront
        Case "BringToFront"
            HtmlEditor1.BringToFront
        
        'SendToBack
        Case "SendToBack"
            HtmlEditor1.SendToBack
        
        'Bring Above Text
        Case "BringAboveText"
        
        'Send Below Text
        Case "SendBelowText"
        
        Case "SnapToGrid"
            Debug.Print "SnapToGrid"
            
        'Textbox
        Case "Textbox"
            HtmlEditor1.InsertInputText
        Case "Textarea"
            HtmlEditor1.InsertTextArea
       
        Case "Checkbox"
            HtmlEditor1.InsertInputCheckbox

'InsertFieldset
'InsertHorizontalRule
'InsertIFrame
'InsertImage
'InsertMarquee
'InsertOrderedList
'InsertParagraph
'InsertUnorderedList
        Case "OptionButton"
            HtmlEditor1.InsertInputRadio
        
        Case "ListBox"
            HtmlEditor1.InsertSelectListbox
            
        Case "DropDownBox"
            HtmlEditor1.InsertSelectDropdown

        Case "PushButton"
            'HtmlEditor1.InsertButton
            HtmlEditor1.InsertInputButton
            
        Case "HiddenData"
            HtmlEditor1.InsertInputHidden
            
        Case "Password"
            HtmlEditor1.InsertInputPassword
            
        Case "SubmitButton"
            HtmlEditor1.InsertInputSubmit
            
        Case "ResetButton"
            HtmlEditor1.InsertInputReset
       
        Case "ImageButton"
            HtmlEditor1.InsertInputImage
        
        Case "FileUpload"
            HtmlEditor1.InsertInputFileUpload
            
    End Select

End Sub

Private Sub SwitchToNormal()
    'If HtmlEditor1.Visible = True Then Exit Sub
    If CurrentMode = 0 Then
    ElseIf CurrentMode = 1 Then
        HtmlEditor1.DocumentHtml = rtfSource.Text
    ElseIf CurrentMode = 2 Then
    End If
    CurrentMode = 0
    rtfSource.Visible = False
    WebBrowser1.Visible = False
    HtmlEditor1.Visible = True
End Sub

Private Sub SwitchToHTML()
    'If rtfSource.Visible = True Then Exit Sub
    rtfSource.Text = HtmlEditor1.DocumentHtml
    HtmlEditor1.Visible = False
    WebBrowser1.Visible = False
    rtfSource.Visible = True
    CurrentMode = 1
End Sub

Private Sub SwitchToPreview()
    If CurrentMode = 2 Then Exit Sub
    
    HtmlEditor1.Visible = False
    rtfSource.Visible = False
    WebBrowser1.Visible = True
    
    If CurrentMode = 0 Then
        WebBrowser1.Navigate2 "about: " & HtmlEditor1.DocumentHtml
        
    ElseIf CurrentMode = 1 Then
        HtmlEditor1.DocumentHtml = rtfSource.Text
        WebBrowser1.Navigate2 "about: " & rtfSource.Text
    End If
    
    CurrentMode = 2
End Sub

Private Function FormatRGBString(val As Long) As String
    
    Dim Color As String
    Dim pad As Long
    Dim R As String
    Dim g As String
    Dim b As String
    
    ' This function formats a long consisting of rgb values
    ' taken from the CommonDialog color dialog
    ' to a string in the form of "#RRGGBB" where RRGGBB are
    ' hex values
    
    ' convert to hex
    Color = Hex(val)
    'determine how many zeros to pad in front of converted value
    pad = 6 - Len(Color)
    
    If pad Then
        Color = String(pad, "0") & Color
    End If
        
    'Extract the rgb components
    R = Right(Color, 2)
    g = Mid(Color, 3, 2)
    b = Left(Color, 2)
    
    ' Swab r and b position, color dialog returns
    ' bgr instead of rgb
    Color = "#" & R & g & b
    
    FormatRGBString = Color
End Function

Private Sub HtmlEditor1_DocumentMouseMove(Element As MSHTMLCtl.IHTMLElement, oEvent As MSHTMLCtl.IHTMLEventObj)
    'Debug.Print "DocumentMouseMove, Element: "; Element.tagName
End Sub

Private Sub HtmlEditor1_DocumentMouseDown(Element As MSHTMLCtl.IHTMLElement, oEvent As MSHTMLCtl.IHTMLEventObj)
    
    'Debug.Print "DocumentMouseDown, Element: "; Element.tagName
    
    'Debug.Print "oEvent "; oEvent.Button
    'Element.offsetWidth
    If oEvent.Button = vbLeftButton Then
        'Debug.Print "DocumentMouseDown Left button"
    End If
End Sub

Private Sub HtmlEditor1_DocumentMouseUp(Element As MSHTMLCtl.IHTMLElement, oEvent As MSHTMLCtl.IHTMLEventObj)
    'Debug.Print "DocumentMouseUp, Element: "; Element.tagName
End Sub

Private Sub HtmlEditor1_DocumentMouseOut(Element As MSHTMLCtl.IHTMLElement, oEvent As MSHTMLCtl.IHTMLEventObj)
    'Debug.Print "DocumentMouseOut, Element: "; Element.tagName
End Sub

Private Sub HtmlEditor1_DocumentMouseOver(Element As MSHTMLCtl.IHTMLElement, oEvent As MSHTMLCtl.IHTMLEventObj)
    'Debug.Print "DocumentMouseOver, Element: "; Element.tagName
End Sub

Private Sub mnuFileOpen_Click()
   RaiseEvent mnuFileOpenClick
End Sub
Private Sub mnuFileSaveAs_Click()
   RaiseEvent mnuFileSaveAsClick
End Sub
Private Sub mnuFileSave_Click()
    RaiseEvent mnuFileSaveClick
End Sub

Private Sub mnuFileNew_Click()
    
    Dim DocText As String
    
    DocText = ""
    HtmlEditor1.DocumentHtml = DocText
    
End Sub
