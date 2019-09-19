VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3282702A-52D1-4AF2-9564-7B938A77CDE1}#2.0#0"; "HtmlEditor.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15328
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
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin VisualHtmlEditor.HtmlEditor HtmlEditor1 
      Height          =   1515
      Left            =   2940
      TabIndex        =   6
      Top             =   1560
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2672
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
            Picture         =   "frmMain.frx":0000
            Key             =   "Bold"
            Object.Tag             =   "Bold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01DA
            Key             =   "Underline"
            Object.Tag             =   "Underline"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03B4
            Key             =   "Italic"
            Object.Tag             =   "Italic"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":058E
            Key             =   "LeftJustify"
            Object.Tag             =   "LeftJustify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0768
            Key             =   "RightJustify"
            Object.Tag             =   "RightJustify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0942
            Key             =   "CenterJustify"
            Object.Tag             =   "CenterJustify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B1C
            Key             =   "FullJustify"
            Object.Tag             =   "FullJustify"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CF6
            Key             =   "Bullets"
            Object.Tag             =   "Bullets"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ED0
            Key             =   "Numbers"
            Object.Tag             =   "Numbers"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AA
            Key             =   "Indent"
            Object.Tag             =   "Indent"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1284
            Key             =   "Outdent"
            Object.Tag             =   "Outdent"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":145E
            Key             =   "LTR"
            Object.Tag             =   "LTR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1638
            Key             =   "SubScript"
            Object.Tag             =   "SubScript"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1792
            Key             =   "SuperScript"
            Object.Tag             =   "SuperScript"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18EC
            Key             =   "StrikeThrough"
            Object.Tag             =   "StrikeThrough"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AC6
            Key             =   "RTL"
            Object.Tag             =   "RTL"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CA0
            Key             =   "ForeColor1"
            Object.Tag             =   "ForeColor1"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DFA
            Key             =   "ForeColor"
            Object.Tag             =   "ForeColor"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24EE
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
            Picture         =   "frmMain.frx":2A88
            Key             =   "WebFile"
            Object.Tag             =   "WebFile"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D22
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
            Picture         =   "frmMain.frx":8E7C
            Key             =   "New2"
            Object.Tag             =   "New2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9056
            Key             =   "Open2"
            Object.Tag             =   "Open2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9230
            Key             =   "Save2"
            Object.Tag             =   "Save2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":940A
            Key             =   "SaveAs1"
            Object.Tag             =   "SaveAs1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":995C
            Key             =   "Print"
            Object.Tag             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B36
            Key             =   "Preview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D10
            Key             =   "Spell"
            Object.Tag             =   "Spell"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9EEA
            Key             =   "Cut1"
            Object.Tag             =   "Cut1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A0C4
            Key             =   "Copy1"
            Object.Tag             =   "Copy1"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A29E
            Key             =   "Paste1"
            Object.Tag             =   "Paste1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A478
            Key             =   "Undo1"
            Object.Tag             =   "Undo1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A5D2
            Key             =   "Redo1"
            Object.Tag             =   "Redo1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A72C
            Key             =   "Table1"
            Object.Tag             =   "Table1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A906
            Key             =   "Image2"
            Object.Tag             =   "Image2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AAE0
            Key             =   "Link"
            Object.Tag             =   "Link"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ACBA
            Key             =   "ShowAll"
            Object.Tag             =   "ShowAll"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AE94
            Key             =   "DeleteCells"
            Object.Tag             =   "DeleteCells"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B06E
            Key             =   "InsertColumns"
            Object.Tag             =   "InsertColumns"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B248
            Key             =   "InsertRows"
            Object.Tag             =   "InsertRows"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B422
            Key             =   "DeleteColumns"
            Object.Tag             =   "DeleteColumns"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B5FC
            Key             =   "ShowBorders"
            Object.Tag             =   "ShowBorders"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B7D6
            Key             =   "HideBorders"
            Object.Tag             =   "HideBorders"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B9B0
            Key             =   "ColsEven"
            Object.Tag             =   "ColsEven"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB8A
            Key             =   "RowsEven"
            Object.Tag             =   "RowsEven1"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BD64
            Key             =   "Download"
            Object.Tag             =   "Download"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BEBE
            Key             =   "MergeCells"
            Object.Tag             =   "MergeCells"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C098
            Key             =   "SplitCells"
            Object.Tag             =   "SplitCells"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C272
            Key             =   "Video"
            Object.Tag             =   "Video"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C44C
            Key             =   "PageSetup"
            Object.Tag             =   "PageSetup"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C626
            Key             =   "PrintPreview"
            Object.Tag             =   "PrintPreview"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C800
            Key             =   "Properties"
            Object.Tag             =   "Properties"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C9DA
            Key             =   "Publish"
            Object.Tag             =   "Publish"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CBB4
            Key             =   "WebTransfer"
            Object.Tag             =   "WebTransfer"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CD0E
            Key             =   "Find"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CE68
            Key             =   "AlignBottom"
            Object.Tag             =   "AlignBottom"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D042
            Key             =   "AlignTop."
            Object.Tag             =   "AlignTop."
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D21C
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D3F6
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D5D0
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D7AA
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D984
            Key             =   "TextDirection"
            Object.Tag             =   "TextDirection"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DB5E
            Key             =   "AutoFit"
            Object.Tag             =   "AutoFit"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD38
            Key             =   "Comment"
            Object.Tag             =   "Comment"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF12
            Key             =   "Website"
            Object.Tag             =   "Website"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E364
            Key             =   "TableAutoFormat"
            Object.Tag             =   "TableAutoFormat"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E53E
            Key             =   "Form"
            Object.Tag             =   "Form"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E718
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EDEA
            Key             =   "New"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F4BC
            Key             =   "Save1"
            Object.Tag             =   "Save1"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F696
            Key             =   "SnapToGrid"
            Object.Tag             =   "SnapToGrid"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F870
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FF42
            Key             =   "SaveAs"
            Object.Tag             =   "SaveAs"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10614
            Key             =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10CE6
            Key             =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E40
            Key             =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11512
            Key             =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1166C
            Key             =   "Redo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":117C6
            Key             =   "Replace"
            Object.Tag             =   "Replace"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11E98
            Key             =   "Find1"
            Object.Tag             =   "Find1"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11FF2
            Key             =   "Image"
            Object.Tag             =   "Image"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":126C4
            Key             =   "Table"
            Object.Tag             =   "Table"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1281E
            Key             =   "FindNext"
            Object.Tag             =   "FindNext"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12978
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
            Picture         =   "frmMain.frx":12AD2
            Key             =   "Normal"
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12CAC
            Key             =   "HTML"
            Object.Tag             =   "HTML"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12E86
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13060
            Key             =   "Preview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":131BA
            Key             =   "Refresh"
            Object.Tag             =   "Refresh"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13394
            Key             =   "Back"
            Object.Tag             =   "Back"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1356E
            Key             =   "Forword"
            Object.Tag             =   "Forword"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13748
            Key             =   "Stop"
            Object.Tag             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13922
            Key             =   "InsertCells"
            Object.Tag             =   "InsertCells"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13AFC
            Key             =   "InsertColumns"
            Object.Tag             =   "InsertColumns"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13CD6
            Key             =   "InsertRows2"
            Object.Tag             =   "InsertRows2"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13EB0
            Key             =   "MergeCells"
            Object.Tag             =   "MergeCells"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1408A
            Key             =   "SplitCells"
            Object.Tag             =   "SplitCells"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14264
            Key             =   "DeleteCells"
            Object.Tag             =   "DeleteCells"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1443E
            Key             =   "DeleteColumns"
            Object.Tag             =   "DeleteColumns"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14618
            Key             =   "DeleteRows"
            Object.Tag             =   "DeleteRows"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14772
            Key             =   "InsertRows"
            Object.Tag             =   "InsertRows"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":148CC
            Key             =   "PositionAbsolutely"
            Object.Tag             =   "PositionAbsolutely"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14AA6
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C80
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E5A
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15034
            Key             =   "SendBackward1-delete"
            Object.Tag             =   "SendBackward1"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1520E
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153E8
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158FA
            Key             =   "Textbox"
            Object.Tag             =   "Textbox"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15B07
            Key             =   "Textarea"
            Object.Tag             =   "Textarea"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15D0C
            Key             =   "Checkbox"
            Object.Tag             =   "Checkbox"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15F11
            Key             =   "OptionButton"
            Object.Tag             =   "OptionButton"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1610B
            Key             =   "DropDown"
            Object.Tag             =   "DropDown"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16314
            Key             =   "PushButton"
            Object.Tag             =   "PushButton"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16512
            Key             =   "HiddenData"
            Object.Tag             =   "HiddenData"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":166EC
            Key             =   "Password"
            Object.Tag             =   "Password"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":168C6
            Key             =   "SubmitButton"
            Object.Tag             =   "SubmitButton"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16AA0
            Key             =   "ResetButton"
            Object.Tag             =   "ResetButton"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16C7A
            Key             =   "ImageButton"
            Object.Tag             =   "ImageButton"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16E54
            Key             =   "Form"
            Object.Tag             =   "Form"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1702E
            Key             =   "BringAboveText"
            Object.Tag             =   "BringAboveText"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17208
            Key             =   "SendBelowText"
            Object.Tag             =   "SendBelowText"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":173E2
            Key             =   "SnapToGrid"
            Object.Tag             =   "SnapToGrid"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":175BC
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
      Width           =   11400
      _ExtentX        =   20108
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
      Width           =   11400
      _ExtentX        =   20108
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
      Width           =   11400
      _ExtentX        =   20108
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
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowBorders"
            Description     =   "ShowBorders"
            Object.ToolTipText     =   "Show Borders"
            Object.Tag             =   "ShowBorders"
            ImageKey        =   "ShowBorders"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    
    On Error Resume Next
    
    Dim X As Long
    '---------------------------------------
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    '---------------------------------------
    
    HtmlEditor1.Visible = True
    'Do While Not HtmlEditor1.ReadyState: DoEvents: Loop
    'Debug.Print HtmlEditor1.ReadyState
    '----------------------------------------------------------------
    'HtmlEditor1.DesignMode = True
    HtmlEditor1.LiveResize = True
    'HtmlEditor1.MultipleSelection = True
    'HtmlEditor1.ShowBorders = True
    'HtmlEditor1.Content = "hello <table><tr><td>ahmed amin</td></tr></table>"
    HtmlEditor1.SetDocument "hello <p>I am here</p><br><table width=450 cellspacing=0 cellpadding=2><tr><td width=350>ahmed amin</td><td>cell  2</td></tr><tr><td>elsheshtawy</td><td>cell 4</td></tr><tr><td>elsheshtawy</td><td>cell 4</td></tr><tr><td>elsheshtawy</td><td>cell 4</td></tr></table> <!--comment here--><br><script>var x;</script> <style>Shesh</style>"
    
    
    Dim Document As HTMLDocument
    
    Set Document = HtmlEditor1.Document
    'Document.body.innerHTML = "ahmed"
    'Document.execCommand "bold"
    '----------------------------------------------------------------
    Dim Formats() As String
    
    Formats() = HtmlEditor1.GetBlockFormats
    For X = LBound(Formats) To UBound(Formats) - 1
        StyleCombo.AddItem Formats(X)
    Next X
    '----------------------------------------------------------------
    Dim fontNames() As String
    
    fontNames = HtmlEditor1.GetFontNames
    For X = LBound(fontNames) To UBound(fontNames)
        FontCombo.AddItem fontNames(X)
    Next X

    FontCombo.ListIndex = 0
    '----------------------------------------------------------------
    ' Font size menu
    For X = 1 To 7
        FontSizeCombo.AddItem CStr(X)
    Next X
    
    FontSizeCombo.ListIndex = 0
    '----------------------------------------------------------------
    'HtmlEditor1.ShowBreakGlyph True
    'HtmlEditor1.ShowAllGlyph True
    'HtmlEditor1.ShowAreaGlyph True
    'HtmlEditor1.ShowCommentGlyph True
    'HtmlEditor1.ShowStyleGlyph True
    'HtmlEditor1.DisableEditFocus True
    'HtmlEditor1.KeepSelection True
    'HtmlEditor1.OverrideCursor False
    'HtmlEditor1.InsertObject
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    DrawWidth = 8   ' Set DrawWidth.
'    Me.DrawMode = 3
'    Me.Line (1000, 1)-(1000, 50000), vbRed
'    Debug.Print "mouse"
End Sub

Private Sub Form_Resize()

    HtmlEditor1.Move Me.ScaleLeft, Me.Top + 200, Me.ScaleWidth, Me.ScaleHeight - 200

End Sub

Private Sub Command2_Click()
    Debug.Print HtmlEditor1.GetDocument
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
    'Toolbar 2
    
    StyleCombo.Text = HtmlEditor1.GetFormatBlock
    FontCombo.Text = HtmlEditor1.GetFontName
    FontSizeCombo.Text = HtmlEditor1.GetFontSize
    
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
        Toolbar1.Buttons(ButtonName).Enabled = True
    Else
        Toolbar1.Buttons(ButtonName).Enabled = False
    End If

End Sub

Private Sub UpdateToolBar2Button(ButtonName As String, Status As Boolean)

    If Status Then
        Toolbar2.Buttons(ButtonName).Value = tbrPressed
    Else
        Toolbar2.Buttons(ButtonName).Value = tbrUnpressed
    End If

End Sub


Private Sub UpdateToolBar3Button(ButtonName As String, Status As Boolean)

    If Status Then
        Toolbar3.Buttons(ButtonName).Value = tbrPressed
    Else
        Toolbar3.Buttons(ButtonName).Value = tbrUnpressed
    End If

End Sub

Private Sub UpdateToolBar3ButtonEnabled(ButtonName As String, Status As Boolean)

    If Status Then
        Toolbar3.Buttons(ButtonName).Enabled = True
    Else
        Toolbar3.Buttons(ButtonName).Enabled = False
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
            'FileSave_Click
        
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
    
    'object.SourceCodePreservation [ = enablePreservation ]
    'object.FilterSourceCode sourceCodeIn
    'object.UseDivOnCarriageReturn [ = div ]
    

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
            Debug.Print "FColor:"; FColor
            
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
            Debug.Print "BackColor: "; BColor
        End Select
        
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
    
        
    ' Handle toolbar commands
    Select Case Button.Key
        Case "Normal"
            'SwitchToNormal
            
        Case "HTML"
            'SwitchToHTML
            
        Case "Preview"
            'SwitchToPreview
    
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
    Dim sFile As String
    Dim FileNumber As Long
    
    With CommonDialog1
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Html Files (*.html)|*.html"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    
    'ActiveForm.Caption = sFile
    'ActiveForm.rtfText.SaveFile sFile
    'ActiveForm.SourceEditor1.SetFocus
    Dim DocText As String
    
    DocText = HtmlEditor1.GetDocument
     
    FileNumber = FreeFile
    Open sFile For Input As FileNumber
    DocText = Input(LOF(FileNumber), FileNumber)
    Close FileNumber
    
    HtmlEditor1.SetDocument DocText
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    Dim FileNumber As Long
        
    With CommonDialog1
        .DialogTitle = "Save As"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Html Files (*.html)|*.html"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    
    'ActiveForm.Caption = sFile
    'ActiveForm.rtfText.SaveFile sFile
    'ActiveForm.SourceEditor1.SetFocus
    Dim DocText As String
    
    DocText = HtmlEditor1.GetDocument
    
    FileNumber = FreeFile
    Open sFile For Output As FileNumber
    Print FileNumber, DocText
    Close FileNumber
    
End Sub

Private Sub mnuFileNew_Click()
    
    Dim DocText As String
    
    DocText = ""
    HtmlEditor1.SetDocument DocText
    
End Sub
