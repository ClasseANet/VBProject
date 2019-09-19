VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDisparos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Disparos"
   ClientHeight    =   9390
   ClientLeft      =   3660
   ClientTop       =   870
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ComboBox CmbIDTPTRATAMENTO 
      Height          =   315
      Left            =   3960
      TabIndex        =   8
      Top             =   360
      Width           =   1815
      _Version        =   720898
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTDisparo 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDisparos.frx":0000
            Key             =   "UP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDisparos.frx":015A
            Key             =   "UP_D"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDisparos.frx":02B4
            Key             =   "DOWN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDisparos.frx":040E
            Key             =   "DOWN_D"
         EndProperty
      EndProperty
   End
   Begin iGrid251_75B4A91C.iGrid GrdDisparos 
      Height          =   7815
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   13785
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   8880
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   8880
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox CmbIDMAQUINA 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   1695
      _Version        =   720898
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   8880
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Excluir"
      ForeColor       =   64
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmDisparos.frx":0568
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   8880
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   32768
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmDisparos.frx":1032
   End
   Begin XtremeSuiteControls.PushButton CmdUp 
      Height          =   300
      Left            =   7815
      TabIndex        =   14
      Top             =   7635
      Width           =   300
      _Version        =   720898
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   79
      ForeColor       =   12582912
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdDown 
      Height          =   300
      Left            =   7815
      TabIndex        =   15
      Top             =   7920
      Width           =   300
      _Version        =   720898
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   79
      ForeColor       =   12582912
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmDisparos.frx":28FC
      TextImageRelation=   0
   End
   Begin XtremeSuiteControls.ComboBox CmbIDMANIPULO 
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
      _Version        =   720898
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTINI 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tratamento"
      Height          =   195
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Manípulo"
      Height          =   195
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label LblMaquina 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Máquina"
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmDisparos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Load()
Event Unload(Cancel As Integer)
Event Resize()
Event CmbIDMAQUINAClick()
Event CmbIDMANIPULOClick()
Event CmbDTDisparoChange()
Event CmbDTDisparoValidate(Cancel As Boolean)
Event CmbDTDisparoLostFocus()
Event CmbIDTPTRATAMENTOClick()

Event CmbDTINIChange()
Event CmbDTINICloseUp()
Event CmbDTINIValidate(Cancel As Boolean)
Event CmbDTINILostFocus()

Event CmdOkClick()
Event CmdCancelClick()
Event CmdSalvarClick()
Event CmdExcluirClick()
Event CmdUpClick()
Event CmdDownClick()

Event GrdDisparosAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdDisparosBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdDisparosCellSelectionChange(ByVal lRow As Long, ByVal lCol As Long, ByVal bSelected As Boolean)
Event GrdDisparosColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdDisparosColHeaderDblClick(ByVal lCol As Long)
Event GrdDisparosDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
Event GrdDisparosKeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Event GrdDisparosLostFocus()
Event GrdDisparosMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdDisparosQuitCustomEdit()
Event GrdDisparosRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdDisparosValidate(Cancel As Boolean)
Private Sub CmbDTDisparo_Change()
   RaiseEvent CmbDTDisparoChange
End Sub
Private Sub CmbDTDisparo_LostFocus()
   RaiseEvent CmbDTDisparoLostFocus
End Sub
Private Sub CmbDTDisparo_Validate(Cancel As Boolean)
   RaiseEvent CmbDTDisparoValidate(Cancel)
End Sub
Private Sub CmbDTINI_Change()
   RaiseEvent CmbDTINIChange
End Sub
Private Sub CmbDTINIo_LostFocus()
   RaiseEvent CmbDTINILostFocus
End Sub
Private Sub CmbDTINI_CloseUp()
   RaiseEvent CmbDTINICloseUp
End Sub
Private Sub CmbDTINI_Validate(Cancel As Boolean)
   RaiseEvent CmbDTINIValidate(Cancel)
End Sub
Private Sub CmbIDMANIPULO_Click()
   RaiseEvent CmbIDMANIPULOClick
End Sub
Private Sub CmbIDMAQUINA_Click()
   RaiseEvent CmbIDMAQUINAClick
End Sub
Private Sub CmbIDTPTRATAMENTO_Click()
   RaiseEvent CmbIDTPTRATAMENTOClick
End Sub
Private Sub cmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdDown_Click()
   RaiseEvent CmdDownClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub cmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub
Private Sub CmdUp_Click()
   RaiseEvent CmdUpClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
Private Sub GrdDisparos_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdDisparosAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdDisparos_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdDisparosBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdDisparos_CellSelectionChange(ByVal lRow As Long, ByVal lCol As Long, ByVal bSelected As Boolean)
   RaiseEvent GrdDisparosCellSelectionChange(lRow, lCol, bSelected)
End Sub
Private Sub GrdDisparos_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdDisparosColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdSESSAO_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdDisparosColHeaderDblClick(lCol)
End Sub
Private Sub GrdDisparos_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   
   With Me.GrdDisparos
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
Private Sub GrdDisparos_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   RaiseEvent GrdDisparosDblClick(lRow, ByVal lCol, bRequestEdit)
End Sub
Private Sub GrdDisparos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
   RaiseEvent GrdDisparosKeyDown(KeyCode, Shift, bDoDefault)
End Sub
Private Sub GrdDisparos_LostFocus()
   RaiseEvent GrdDisparosLostFocus
End Sub
Private Sub GrdDisparos_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   RaiseEvent GrdDisparosMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdDisparos_QuitCustomEdit()
   RaiseEvent GrdDisparosQuitCustomEdit
End Sub
Private Sub GrdDisparos_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdDisparosRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdDisparos_TextEditKeyUp(ByVal lRow As Long, ByVal lCol As Long, ByVal KeyCode As Integer, ByVal Shift As Integer)
   If KeyCode = vbKeyUp Then
      If lRow >= 2 Then
         Me.GrdDisparos.CommitEdit
         Call Me.GrdDisparos.SetCurCell(lRow - 1, lCol)
      End If
   ElseIf KeyCode = vbKeyDown Then
      If lRow < Me.GrdDisparos.RowCount - 1 Then
         Me.GrdDisparos.CommitEdit
         Call Me.GrdDisparos.SetCurCell(lRow + 1, lCol)
      End If
   End If
End Sub

Private Sub GrdDisparos_Validate(Cancel As Boolean)
   RaiseEvent GrdDisparosValidate(Cancel)
End Sub

Private Sub VScroll1_Change()
End Sub
