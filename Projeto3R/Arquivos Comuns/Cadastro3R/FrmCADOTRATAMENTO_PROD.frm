VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmCADOTRATAMENTO_PROD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Produtos X Tratamentos"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sai&r"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin iGrid251_75B4A91C.iGrid GrdProd 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   6376
      BorderStyle     =   1
      DefaultRowHeight=   17
      HighlightBackColorNoFocus=   14737632
   End
   Begin XtremeSuiteControls.PushButton CmdTratamento 
      Height          =   280
      Left            =   120
      TabIndex        =   3
      Top             =   3760
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   494
      _StockProps     =   79
      Caption         =   "&Tratamento"
      ForeColor       =   16384
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdProduto 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   3760
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   494
      _StockProps     =   79
      Caption         =   "&Produto"
      ForeColor       =   16384
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "FrmCADOTRATAMENTO_PROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CmdSalvarClick()
Event CmdSairClick()
Event CmdNovoClick()
Event CmdExcluirClick()
Event CmdProdutoClick()
Event CmdTratamentoClick()
Event GrdProdAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdProdBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdProdColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdProdColHeaderDblClick(ByVal lCol As Long)
Event GrdProdMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdProdLostFocus()
Event GrdProdRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdProdValidate(Cancel As Boolean)
Private Sub CmdProduto_Click()
   RaiseEvent CmdProdutoClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub
Private Sub CmdTratamento_Click()
   RaiseEvent CmdTratamentoClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub GrdProd_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdProdAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdProd_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdProdBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdProd_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdProdColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdProd_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdProdColHeaderDblClick(lCol)
End Sub
Private Sub GrdProd_LostFocus()
   RaiseEvent GrdProdLostFocus
End Sub
Private Sub GrdProd_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdProdMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdProd_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdProdRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdProd_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   
   With Me.GrdProd
      .RowMode = (lRow = .RowCount)
      If .RowCount > 0 Then
         For i = 1 To .ColCount
            If .ColVisible(i) Then
               .CellForeColor(.RowCount, i) = IIf(lRow = .RowCount, vbHighlightText, vbGrayText)
               Exit For
            End If
         Next
      End If
   End With
End Sub
Private Sub GrdProd_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdProd.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdProd_Validate(Cancel As Boolean)
   RaiseEvent GrdProdValidate(Cancel)
End Sub
