VERSION 5.00
Object = "{912C475C-AE70-4EAF-98BD-40407863D247}#1.0#0"; "MCIControls.ocx"
Begin VB.Form FrmTestGrid 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin MCIControls.PctAzulEscuro MCIGrid1 
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      _extentx        =   13785
      _extenty        =   5530
      backcolor       =   -2147483633
      font            =   "FrmTestGrid.frx":0000
      scalewidth      =   7815
      scalemode       =   0
      scaleheight     =   3135
      columns         =   2
      rowheight0      =   240
      mouseicon       =   "FrmTestGrid.frx":002C
      gridcolorfixed  =   16761024
      gridcolor       =   16761024
      forecolorsel    =   -2147483634
      forecolorfixed  =   -2147483630
      fontsize        =   8,25
      fontname        =   "MS Sans Serif"
      fontfixed       =   "FrmTestGrid.frx":004A
      colwidth0       =   -1
      cols0           =   2
      backcolorsel    =   16777152
      backcolorfixed  =   -2147483633
   End
End
Attribute VB_Name = "FrmTestGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   Dim i As Integer
   Dim j As Integer
   With Me.MCIGrid1
      .Columns = 10
      .Rows = 10
      '.SelectionMode = flexSelectionByRow
      For i = 0 To .Rows - 1
         For j = 0 To .Columns - 1
            .TextMatrix(i, j) = i & j
         Next
      Next
   End With
End Sub

