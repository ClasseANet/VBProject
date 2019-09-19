VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjComboBox"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Função que ajusta a combo ao abrir com o tamanho da maior string.
'Só funciona para combos do tipo XtremeSuiteControls
Public Sub AjustaTextoComboCodeJock(ByRef pCmb As Object, ByVal pFrm As Object)

   Dim i As Integer
   Dim j As Long
   
   On Error GoTo Fim
   
   pFrm.FontName = pCmb.Font.Name
   pFrm.FontSize = pCmb.Font.Size
   
   For i = 0 To pCmb.ListCount - 1
        If pFrm.TextWidth(pCmb.List(i)) > j Then
            j = pFrm.TextWidth(pCmb.List(i))
        End If
   Next i

   pCmb.DropDownWidth = j + 400
   
Fim:
   'Resume Next
End Sub
Public Function LocalizarCombo(Cmb, Chave As String, Optional SetCombo = True, Optional PorItemData As Boolean = False) As Integer
   Dim Ind  As Integer
   Dim i    As Integer
   If PorItemData Then
      Ind = -1
      For i = 0 To Cmb.ListCount - 1
         If Cmb.ItemData(i) = Chave Then
            Ind = i
            Exit For
         End If
      Next
   Else
      Ind = SendMessageAny(Cmb.hwnd, CB_FINDSTRING, -1, ByVal Chave)
   End If
   If SetCombo Then
      If Cmb.ListCount <> 0 Then
         Cmb.ListIndex = Ind
         'If Cmb.Style <> 2 Then
         '   Cmb = Cmb.List(Ind)
         'End If
      End If
      If Cmb.Style <> 2 And Ind = -1 Then
         Cmb.Text = ""
      End If
   End If
   
   LocalizarCombo = Ind
End Function
