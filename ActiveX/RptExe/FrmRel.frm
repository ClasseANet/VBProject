VERSION 5.00
Begin VB.Form FrmRel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extrair Relatórios"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "FrmRel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LstRel 
      Height          =   5130
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton CmdOper 
      Caption         =   "Extrair"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Lbl 
      Caption         =   "Relatórios"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOper_Click()
   Dim sNmRel   As String
   Dim sCommand As String
   Dim sItem   As String

   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   
   sItem = Me.LstRel.List(Me.LstRel.ListIndex)
   If sItem = "" Then
      Call MsgBox("Selecione um item.", Title:="Arquivos")
   Else
      sCommand = Mid(sItem, 1, InStr(sItem, "-") - 2)
      sNmRel = sCommand & ".rpt"

      Call ExtractResData(sCommand, "RPT", App.Path & "\" & sNmRel)
   End If
   MsgBox "Arquivo extraído com sucesso!", Title:="Arquivos"
   
   Screen.MousePointer = vbDefault
   Exit Sub
TrataErro:
   MsgBox Err & " - " & Error
   Resume Next
End Sub
Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   With Me.LstRel
      .Clear
      .AddItem "PM001 - PM's para Área"
      
      If .ListCount > 0 Then
         .ListIndex = 0
      End If
   End With
   
   Screen.MousePointer = vbDefault
End Sub
