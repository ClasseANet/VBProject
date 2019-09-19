VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CodeGuru Zip/Unzip Test Client"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox zz 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2250
      ScaleWidth      =   8385
      TabIndex        =   0
      Top             =   0
      Width           =   8385
      Begin VB.Frame Frame1 
         Height          =   1005
         Left            =   2505
         TabIndex        =   5
         Top             =   -105
         Width           =   4410
         Begin VB.Label Label1 
            Caption         =   $"Form1.frx":10D2
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   75
            TabIndex        =   6
            Top             =   135
            Width           =   4245
         End
      End
      Begin VB.CommandButton cmdUnZip 
         Caption         =   "UnZip it"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   2
         Top             =   1650
         Width           =   1215
      End
      Begin VB.CommandButton cmdZip 
         Caption         =   "Zip Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblTempDir 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1545
         TabIndex        =   4
         Top             =   1740
         Width           =   2910
      End
      Begin VB.Label lblCurDir 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1545
         TabIndex        =   3
         Top             =   1095
         Width           =   2580
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Unzip/Zip Client program for the CGZipLibrary ActiveXDLL
'
'
' Chris Eastwood, July 1999

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Sub cmdUnZip_Click()

On Error GoTo vbErrorHandler

   'Call RegServer(App.Path & "\Unzip.dll")

'
' Unzip the ZIPTEST.ZIP file to the Windows Temp Directory
'
    Dim oUnZip As CGUnzipFiles
    
    Set oUnZip = New CGUnzipFiles
    With oUnZip
'
' What Zip File ?
'
        .ZipFileName = App.Path & "\MSDE.zip" '"C:\ZIPTEST.ZIP"
'
' Where are we zipping to ?
'
        .ExtractDir = GetTempPathName
'
' Keep Directory Structure of Zip ?
'
        .HonorDirectories = False
'
' Unzip and Display any errors as required
'
        If .Unzip <> 0 Then
            MsgBox .GetLastMessage
        End If
    End With
    
    Set oUnZip = Nothing
    MsgBox "\ZIPTEST.ZIP Extracted Successfully to " & GetTempPathName

    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Form1::cmdUnZip_Click" & " " & Err.Description

End Sub

Private Sub cmdZip_Click()
    Dim oZip As CGZipFiles

On Error GoTo vbErrorHandler

   
    Set oZip = New CGZipFiles
    
    With oZip
'
' Give Zip File a Name / Path
'
        .ZipFileName = "\ZIPTEST.ZIP"
'
' Are we updating a Zip File ?
' - This doesn't seem to work - check InfoZip
' homepage for more info.
'
        .UpdatingZip = False ' ensures a new zip is created
'
' Add in the files to the zip - in this case, we
' want all the ones in the current directory
'
        .AddFile App.Path & "\*.*"
'
' Make the zip file & display any errors
'
        If .MakeZipFile <> 0 Then
            MsgBox .GetLastMessage ' any errors
        End If
    End With
    
    Set oZip = Nothing
    
    MsgBox "\ZIPTEST.ZIP Created Successfully"
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Form1::cmdZip_Click" & " " & Err.Description

End Sub

Private Sub Form_Load()
    lblTempDir.Caption = GetTempPathName
    lblCurDir.Caption = App.Path
End Sub

Private Function GetTempPathName() As String
    Dim sBuffer As String
    Dim lRet As Long
    
    sBuffer = String$(255, vbNullChar)
    
    lRet = GetTempPath(255, sBuffer)
    
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    GetTempPathName = sBuffer
    
End Function

