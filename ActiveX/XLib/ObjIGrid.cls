VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjIGrid"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function iGridToRecordset(ByVal pIGrid As Object, Optional pSomenteSelecao, Optional pRsDados As Object) As Object
'Public Function iGridToRecordset(ByVal pIGrid As iGrid, Optional pSomenteSelecao, Optional pRsDados As Recordset) As Recordset
   Dim x             As Long
   Dim Y             As Long
   Dim Rs            As New Recordset
   Dim bSoSelecao    As Boolean
   Dim bRsCarregado  As Boolean
   On Error GoTo ERRO
   'Verifica se os parâmentros de seleção e Rsdados foram utilizados
   If Not IsMissing(pSomenteSelecao) Then
      If pSomenteSelecao = "S" Then
         bSoSelecao = True
      Else
         bSoSelecao = False
      End If
   Else
      bSoSelecao = False
   End If
   If Not IsMissing(pRsDados) Then
      If Not pRsDados Is Nothing Then
         bRsCarregado = True
      Else
         bRsCarregado = False
      End If
   Else
      bRsCarregado = False
   End If
   
   Rs.ActiveConnection = Nothing
   Rs.CursorLocation = adUseClient
   Rs.LockType = adLockBatchOptimistic
   
   If bRsCarregado = True Then
      For x = 1 To pIGrid.ColCount - 1
         Y = 0
         Do While Y <= pRsDados.Fields.Count
            If pRsDados.Fields(Y).Name = pIGrid.ColKey(x) Then
               Select Case pRsDados.Fields(Y).Type
                  Case DataTypeEnum.adChar, DataTypeEnum.adVarChar, DataTypeEnum.adChapter, DataTypeEnum.adLongVarChar
                     Rs.Fields.Append pIGrid.ColKey(x), adVarChar, 200  'Texto
                  Case DataTypeEnum.adBigInt, DataTypeEnum.adInteger, DataTypeEnum.adBinary, DataTypeEnum.adCurrency, DataTypeEnum.adDecimal, DataTypeEnum.adDouble, DataTypeEnum.adNumeric
                     Rs.Fields.Append pIGrid.ColKey(x), adNumeric, 20 'Número
                  Case DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTimeStamp
                     Rs.Fields.Append pIGrid.ColKey(x), adVarChar, 15 'Data
               End Select
               Exit Do
            End If
            Y = Y + 1
         Loop
      Next x
   Else
      For x = 1 To pIGrid.ColCount
         Rs.Fields.Append pIGrid.ColKey(x), adVarChar, 200
      Next x
   End If
   Rs.Open
   
   If bSoSelecao = False Then
      For x = 1 To pIGrid.RowCount
         Rs.AddNew
         For Y = 1 To pIGrid.ColCount
            Rs.Fields(pIGrid.ColKey(Y)) = CStr(pIGrid.CellValue(x, Y) & "")
         Next Y
      Next x
   Else
      For x = 1 To pIGrid.RowCount
         If pIGrid.CellValue(x, pIGrid.ColCount) = 1 Or pIGrid.CellValue(x, pIGrid.ColCount) = True Then
            Rs.AddNew
            For Y = 1 To pIGrid.ColCount - 1
               Rs.Fields(pIGrid.ColKey(Y)) = CStr(pIGrid.CellValue(x, Y) & "")
            Next Y
         End If
      Next x
   End If
   
   Set iGridToRecordset = Rs

   GoTo Fim
   
ERRO:
   Set iGridToRecordset = New Recordset
Fim:
   Set Rs = Nothing
End Function
