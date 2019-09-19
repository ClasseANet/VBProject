VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetTransactions(Optional sArq As String) As Collection
   Dim Queries As Collection
   Dim Trans   As Collection
   Dim i       As Long
   
   Set Trans = New Collection
   For i = 1 To 2
      Set Queries = New Collection
      Set Queries = GetQueries(CStr(i))
      
      Trans.Add Queries, CStr(i)
   Next
   
   
   Set GetTransactions = Trans
End Function
Public Function GetQueries(Optional sArq As String) As Collection
   Dim Sql     As String
   Dim Queries As Collection

   Set Queries = New Collection
   
   If sArq = "1" Then
      Sql = ""
      Sql = Sql & "IF NOT EXISTS("
      Sql = Sql & "              SELECT COL.NAME "
      Sql = Sql & "              FROM SYSOBJECTS OBJ "
      Sql = Sql & "                  INNER JOIN SYSCOLUMNS COL ON OBJ.ID = COL.ID "
      Sql = Sql & "              WHERE  OBJ.NAME LIKE 'PEDCOTACAO' "
      Sql = Sql & "              AND COL.NAME LIKE 'SITCOTACAO'"
      Sql = Sql & "              )"
      Sql = Sql & " ALTER TABLE PEDCOTACAO ADD SITCOTACAO VARCHAR(3) DEFAULT('33A')"
         
      Queries.Add Sql
         
      Sql = ""
      Sql = Sql & "IF NOT EXISTS("
      Sql = Sql & "              SELECT COL.NAME "
      Sql = Sql & "              FROM SYSOBJECTS OBJ "
      Sql = Sql & "                  INNER JOIN SYSCOLUMNS COL ON OBJ.ID = COL.ID "
      Sql = Sql & "              WHERE  OBJ.NAME LIKE 'ITEMCOTACAO' "
      Sql = Sql & "              AND COL.NAME LIKE 'VLESTIMADO'"
      Sql = Sql & "              )"
      Sql = Sql & " ALTER TABLE ITEMCOTACAO ADD VLESTIMADO DECIMAL(15,4)"
      Queries.Add Sql
   
'      Sql = ""
'      Sql = Sql & "IF NOT EXISTS("
'      Sql = Sql & "              SELECT COL.NAME "
'      Sql = Sql & "              FROM SYSOBJECTS OBJ "
'      Sql = Sql & "                  INNER JOIN SYSCOLUMNS COL ON OBJ.ID = COL.ID "
'      Sql = Sql & "              WHERE  OBJ.NAME LIKE 'ITEMCOTACAO' "
'      Sql = Sql & "              AND COL.NAME LIKE 'QTDOC'"
'      Sql = Sql & "              )"
'      Sql = Sql & " ALTER TABLE ITEMCOTACAO ADD QTDOC DECIMAL(15,4)"
'      Queries.Add Sql
   
   ElseIf sArq = "2" Then
      Sql = "Update PEDCOTACAO Set SITCOTACAO= '33A' Where SITCOTACAO is null"
      Queries.Add Sql
   End If
   
   Set GetQueries = Queries
End Function
