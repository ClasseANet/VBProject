VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DbTabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarNOME        As String
Private mvarBANCO       As String
Private mvarOWNER       As String
Private mvarDESCRICAO   As String

Private mvarFIELDS      As Collection
Private mvarPrimaryKey  As Collection

Private mvarisSys       As Boolean
Private mvarStrConect   As String

Private mvarADOConect As ADODB.Connection
Public Property Set ADOConect(vData As ADODB.Connection)
   Set mvarADOConect = vData
'   If SetProperty Then Exit Property
'   If mvarxDBs.Count = 0 Then Exit Property
End Property
Public Property Get ADOConect() As ADODB.Connection
   Set ADOConect = mvarADOConect
End Property

Public Property Set FIELDS(ByVal vData As Object)
   Set mvarFIELDS = vData
End Property
Public Property Get FIELDS() As Collection
   Dim dBaseAux As Database
   Dim WsAux As Workspace
   Dim MyField As DbCampo
     
   Dim RsFields As New ADODB.Recordset
   Dim RsPrimaryKeys As New ADODB.Recordset
   
   Dim Achou As Boolean
   
   If mvarFIELDS Is Nothing Then
      If True Then
         On Error Resume Next
         Set RsFields = mvarADOConect.OpenSchema(adSchemaColumns, Array(mvarBANCO, mvarOWNER, mvarNOME, Empty))
         If Err <> 0 Then
            Err = 0
            Set RsFields = mvarADOConect.OpenSchema(adSchemaColumns)
         End If
         
         Achou = False
         Set mvarFIELDS = New Collection
         
         'For i = 1 To RsFields.FIELDS.Count
         '   Debug.Print RsFields.FIELDS(i).NAME & " = " & RsFields.FIELDS(i).Value
         'Next
         Do Until RsFields.EOF
            If RsFields("TABLE_NAME") & "" = mvarNOME Then
               Achou = True
               Set MyField = New DbCampo

               MyField.NOME = RsFields("COLUMN_NAME") & ""
               MyField.DESCRICAO = RsFields("DESCRIPTION") & ""
               MyField.TIPO = RsFields("DATA_TYPE") & ""
               MyField.TAMANHO = RsFields("CHARACTER_MAXIMUM_LENGTH") & ""
               MyField.isNull = RsFields("IS_NULLABLE") & ""
               MyField.isSys = False
               

               mvarFIELDS.Add MyField, MyField.NOME
               
               Set MyField = Nothing
               
            Else
               If Achou Then Exit Do
            End If
            RsFields.MoveNext
         Loop
'         Set RsPrimaryKeys = mvarADOConect.OpenSchema(adSchemaPrimaryKeys, Array(mvarBANCO, mvarOWNER, mvarNOME))
'         Set mvarPrimaryKey = New Collection
'         If Err <> 0 Then
'            Err = 0
'            Set RsPrimaryKeys = mvarADOConect.OpenSchema(adSchemaPrimaryKeys)
'         End If
'         Do Until RsPrimaryKeys.EOF
'            If RsPrimaryKeys("TABLE_NAME") = mvarNOME Then
'               mvarPrimaryKey.Add mvarFIELDS(RsPrimaryKeys("COLUMN_NAME")), RsPrimaryKeys("COLUMN_NAME")
'            End If
'            RsPrimaryKeys.MoveNext
'         Loop
         RsFields.Close
'         RsPrimaryKeys.Close
      Else
         Set mvarFIELDS = New Collection
         Set DBEngine = Nothing
         Set WsAux = DBEngine.CreateWorkspace("WsAux", "admin", "")
      
         Set dBaseAux = WsAux.OpenDatabase("", False, False, mvarStrConect)
         For Each n In dBaseAux.TableDefs(mvarNOME).FIELDS
            DoEvents
            If n.Attributes < dbSystemField Then
               Set MyField = New DbCampo
      '            With MyField
                  MyField.NOME = n.NAME
                  MyField.DESCRICAO = n.Description
                  MyField.TAMANHO = n.Size
                  MyField.TIPO = n.Type
                  MyField.isSys = False
      '            End With
               mvarFIELDS.Add MyField
               Set MyField = Nothing
            End If
         Next
         dBaseAux.Close
         WsAux.Close
         Set dBaseAux = Nothing
         Set WsAux = Nothing
      End If
   End If
   Set FIELDS = mvarFIELDS
End Property
Public Property Set PrimaryKey(ByVal vData As Object)
   Set mvarPrimaryKey = vData
End Property
Public Property Get PrimaryKey() As Collection
   Dim MyField As DbCampo
   
   Dim RsFields As New ADODB.Recordset
   Dim Achou As Boolean
   
   If mvarPrimaryKey Is Nothing Then
      On Error Resume Next
      Set RsFields = mvarADOConect.OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, mvarNOME))
      If Err <> 0 Then
         Err = 0
         Set RsFields = mvarADOConect.OpenSchema(adSchemaPrimaryKeys)
      End If
      
      Achou = False
      Set mvarPrimaryKey = New Collection
      Do Until RsFields.EOF
         If RsFields("TABLE_NAME") = mvarNOME Then
            Achou = True
            Set MyField = New DbCampo
      
            MyField.NOME = FIELDS(RsFields("COLUMN_NAME") & "").NOME
            MyField.DESCRICAO = mvarFIELDS(RsFields("DESCRIPTION") & "").DESCRICAO
            MyField.TIPO = mvarFIELDS(RsFields("COLUMN_NAME") & "").TIPO
            MyField.TAMANHO = mvarFIELDS(RsFields("COLUMN_NAME") & "").TAMANHO
            MyField.isNull = mvarFIELDS(RsFields("COLUMN_NAME") & "").isNull
            MyField.isSys = mvarFIELDS(RsFields("COLUMN_NAME") & "").isSys
      
            mvarPrimaryKey.Add MyField, MyField.NOME
            
            Set MyField = Nothing
            
         Else
            If Achou Then Exit Do
         End If
         RsFields.MoveNext
      Loop
      RsFields.Close
   End If
   Set PrimaryKey = mvarPrimaryKey
End Property
Public Property Let BANCO(ByVal vData As String)
   mvarBANCO = vData
End Property
Public Property Get BANCO() As String
   BANCO = mvarBANCO
End Property
Public Property Let OWNER(ByVal vData As String)
   mvarOWNER = vData
End Property
Public Property Get OWNER() As String
   OWNER = mvarOWNER
End Property
Public Property Let NOME(ByVal vData As String)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Public Property Let DESCRICAO(ByVal vData As String)
   mvarDESCRICAO = vData
End Property
Public Property Get DESCRICAO() As String
   DESCRICAO = mvarDESCRICAO
End Property
Public Property Let isSys(ByVal vData As Boolean)
   mvarisSys = vData
End Property
Public Property Get isSys() As Boolean
   isSys = mvarisSys
End Property
Public Property Let StrConect(ByVal vData As String)
   mvarStrConect = vData
End Property
Public Property Get StrConect() As String
   StrConect = mvarStrConect
End Property

