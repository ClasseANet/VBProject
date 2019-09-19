VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmOpenURL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open URL"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   2040
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   300
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenURL 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4140
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   780
      Width           =   6615
   End
   Begin VB.Label lblProgressInfo 
      Alignment       =   2  'Center
      Caption         =   "lblProgressInfo"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "URL:"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Type the internet address of the document to open."
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7575
   End
End
Attribute VB_Name = "frmOpenURL"
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

'this variable stores size of the file
'we are going to retrieve from the Web
Private m_lngDocSize As Long

Private Sub cboAddress_Click()
     cmdOpenURL_Click
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cmdOpenURL_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpenURL_Click()
    
    On Error GoTo Errors
    
    Dim URL As String
    
    URL = QualifyURL(cboAddress.Text)
    
    Screen.MousePointer = vbHourglass
    
    'reset file size value
    m_lngDocSize = 0
    
    'reset the ProgressBar control
    ProgressBar1.Value = 0.001
    '
    'clear the label control
    lblProgressInfo.Caption = ""
    '
    'define protocol for the ITC
    Inet1.protocol = icHTTP
    '
    'call the Execute method to send
    'HTTP request to the webserver
    If Len(URL) > 0 Then
        Inet1.Execute Trim$(URL), "GET"
    End If
    Exit Sub

Errors:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    '
    Dim strText As String
    Dim strBuffer As String
    Dim sngProgerssValue As Single
    '
    On Error Resume Next
    '
    Select Case State
        '
        Case icResponseCompleted
            '
            Do  'retrieve data from the buffer
                '
                DoEvents
                '
                strBuffer = Inet1.GetChunk(512)
                strText = strText & strBuffer
                '
                If m_lngDocSize > 0 Then
                    If Len(strBuffer) > 0 Then
                        'get percent value
                        sngProgerssValue = Int((Len(strText) / m_lngDocSize) * 100)
                    End If
                    'update the label control with new caption
                    lblProgressInfo.Caption = "Downloaded " & CStr(Len(strText)) & " bytes (" & CStr(sngProgerssValue) & "%)"
                    'update the PregressBar control with new value
                    ProgressBar1.Value = sngProgerssValue
                End If
            Loop Until Len(strBuffer) = 0
            
            'put retrieved HTML source into the RichTextBox control
            FrmEditorH.HtmlEditor1.DocumentHtml = strText
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
            
        Case icResponseReceived
            If m_lngDocSize = 0 Then
                'retrieve size of the document
                If Len(Inet1.GetHeader("Content-Length")) > 0 Then
                    m_lngDocSize = CLng(Inet1.GetHeader("Content-Length"))
                End If
                '
            End If
            '
    End Select
    
End Sub

