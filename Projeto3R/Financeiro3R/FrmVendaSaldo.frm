VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmVendaSaldo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FrmVendaSaldo"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin iGrid251_75B4A91C.iGrid GrdAtendimento 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9551
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin XtremeSuiteControls.FlatEdit TxtIDVENDA 
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   14737632
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTVENDA 
      Height          =   345
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   14737632
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtNome 
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5055
      _Version        =   720898
      _ExtentX        =   8916
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   14737632
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   7440
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label LblCLIENTE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   480
   End
   Begin VB.Label LblDTVENDA 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   240
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   435
   End
   Begin VB.Label LblVenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Venda"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   600
   End
   Begin VB.Label LblAtendimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Atendimentos / Faturas"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "FrmVendaSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event CmdOkClick()
Event Activate()
Event DblClick()
Event KeyUp(KeyCode, Shift)
Event Load()
Event GrdAtendimentoAfterCommitEdit(lRow, lCol)
Event GrdAtendimentoBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
Event GrdAtendimentoColHeaderClick(lCol, bDoDefault, Shift, x, y)
Event GrdAtendimentoColHeaderDblClick(lCol)
Event GrdAtendimentoCustomDrawCell(lRow, lCol, hdc, lLeft, lTop, lRight, lBottom, bSelected)
Event GrdAtendimentoMouseDown(Button, Shift, x, y, lRow, lCol, bDoDefault, bUnderControl)
Event GrdAtendimentoMouseEnter(lRow, lCol)
Event GrdAtendimentoMouseLeave(lRow, lCol)
Event GrdAtendimentoMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
Event GrdAtendimentoRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
Event GrdAtendimentoDblClick(lRow, lCol, bRequestEdit)
Event GrdAtendimentoValidate(Cancel)
Event GrdAtendimentoLostFocus()
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_DblClick()
   RaiseEvent DblClick
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub GrdAtendimento_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendimentoAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdAtendimento_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdAtendimentoBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdAtendimento_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdAtendimentoColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdAtendimento_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdAtendimentoColHeaderDblClick(lCol)
End Sub
Private Sub GrdAtendimento_CustomDrawCell(ByVal lRow As Long, ByVal lCol As Long, ByVal hdc As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal bSelected As Boolean)
   RaiseEvent GrdAtendimentoCustomDrawCell(lRow, lCol, hdc, lLeft, lTop, lRight, lBottom, bSelected)
End Sub
Private Sub GrdAtendimento_LostFocus()
   RaiseEvent GrdAtendimentoLostFocus
End Sub
Private Sub GrdAtendimento_MouseDown(ByVal Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean, ByVal bUnderControl As Boolean)
   RaiseEvent GrdAtendimentoMouseDown(Button, Shift, x, y, lRow, lCol, bDoDefault, bUnderControl)
End Sub
Private Sub GrdAtendimento_MouseEnter(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendimentoMouseEnter(lRow, lCol)
End Sub
Private Sub GrdAtendimento_MouseLeave(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendimentoMouseLeave(lRow, lCol)
End Sub
Private Sub GrdAtendimento_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   RaiseEvent GrdAtendimentoMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdAtendimento_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdAtendimentoRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdAtendimento_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   With Me.GrdAtendimento
      .RowMode = True '(lRow = .RowCount)
      If .RowCount > 0 And lRow > 0 Then
         For i = 1 To .ColCount
            If .ColVisible(i) Then
               If lRow = .RowCount And Mid(.CellValue(.RowCount, 1), 1, 6) = "Clique" Then
                  .CellForeColor(lRow, i) = vbGrayText
               Else
                  .CellForeColor(lRow, i) = vbHighlightText
               End If
               Exit For
            End If
         Next
      End If
   End With
End Sub
Private Sub GrdAtendimento_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   RaiseEvent GrdAtendimentoDblClick(lRow, lCol, bRequestEdit)
End Sub
Private Sub GrdAtendimento_Validate(Cancel As Boolean)
   RaiseEvent GrdAtendimentoValidate(Cancel)
End Sub

