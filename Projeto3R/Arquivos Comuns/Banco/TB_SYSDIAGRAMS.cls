VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_SYSDIAGRAMS"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb            As DS_BANCO
Private mvarRS             As Object 

Private mvarQryInsert      As String
Private mvarQryUpDate      As String
Private mvarQryDelete      As String
Private mvarQrySelect      As String
Private mvarQrySave        As String

Private mvarisDirt         As Boolean

Private mvarname         As String
Private mvarprincipal_id As Double
Private mvardiagram_id   As Double
Private mvarversion      As Double
Private mvardefinition   As Double
Public Property Let name(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarname <> vData)
   mvarname = vData
End Property
Public Property Get name() As String
   name = mvarname
End Property
Public Property Let principal_id(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarprincipal_id <> vData)
   mvarprincipal_id = vData
End Property
Public Property Get principal_id() As Double
   principal_id = mvarprincipal_id
End Property
Public Property Get diagram_id() As Double
   diagram_id = mvardiagram_id
End Property
Public Property Let version(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarversion <> vData)
   mvarversion = vData
End Property
Public Property Get version() As Double
   version = mvarversion
End Property
Public Property Let definition(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvardefinition <> vData)
   mvardefinition = vData
End Property
Public Property Get definition() As Double
   definition = mvardefinition
End Property
Public Property Get isDirt() As Boolean
   isDirt = mvarisDirt
End Property
Public Property Set XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Let XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Get XDb() As Object
   Set XDb = mvarXDb
End Property
Public Property Get RS() As Object
   Set RS = mvarRS
End Property
Public Property Get QryInsert() As String
   Dim Sql As String
 
   Sql = "Insert Into SYSDIAGRAMS (name, principal_id, version" & VbNewLine 
   Sql = Sql & ", definition) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarname) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarprincipal_id) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarversion) & VbNewLine 
   Sql = Sql & ", " &  & VbNewLine 
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_diagram_id) As String
   Dim Sql As String
 
   Sql = "Delete From SYSDIAGRAMS" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_diagram_id) Then Sql = Sql & " diagram_id = " & SqlNum(Cstr(Ch_diagram_id)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update SYSDIAGRAMS Set " & VbNewLine 
   Sql = Sql & " name         = " & SqlStr(mvarname) & VbNewLine 
   Sql = Sql & " , principal_id = " & SqlNum(mvarprincipal_id) & VbNewLine 
   Sql = Sql & " , version      = " & SqlNum(mvarversion) & VbNewLine 
   Sql = Sql & " , definition   = " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " diagram_id     = " & SqlNum(mvardiagram_id) & VbNewLine 
   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave() As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_diagram_id, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select name, principal_id, diagram_id, version" & VbNewLine 
   Sql = Sql & ", definition" & VbNewLine 
   Sql = Sql &" From SYSDIAGRAMS" & VbNewLine 
   If IsMissing(Ch_WHERE) Then
      Sql = Sql & " Where " & VbNewLine 
      If IsMissing(Ch_diagram_id)  Then 
         Sql = Sql & " diagram_id = " & SqlNum(mvardiagram_id) & VbNewLine 
      Else 
         If Not isMissing(Ch_diagram_id) Then Sql = Sql & " diagram_id = " & SqlNum(Cstr(Ch_diagram_id)) & " AND " & VbNewLine 
         Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
      End If 
   Else
      If Trim(Ch_WHERE) <> "" Then
         Sql = Sql & " Where " & Ch_WHERE & vbNewLine
      End If
   End If
   If Not IsMissing(Ch_ORDERBY) Then
      If Trim(Ch_ORDERBY) <> "" Then
         Sql = Sql & " Order By " & Ch_ORDERBY & vbNewLine
      End If
   End If

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar( Optional Ch_diagram_id, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_diagram_id, Ch_WHERE, Ch_ORDERBY)
   bExiste = mvarXDb.AbreTabela(Sql, mvarRS)
   mvarisDirt = False
   With mvarRS
      If bExiste Then bExiste = Not .EOF
      If bExiste Then
         Me.Popula
         Pesquisar = True
      Else
         Pesquisar = False
      End If
   End With
   Exit Function
PesquisarErr:
    call ShowError(Sql)
    Pesquisar = False
End Function
Public Sub Popula(Optional pRcSet)
   If IsMissing(pRcSet) Then Set pRcSet = mvarRS
   With pRcSet
      mvarname = !name & ""
      mvarprincipal_id = XVal(!principal_id & "")
      mvardiagram_id = XVal(!diagram_id & "")
      mvarversion = XVal(!version & "")
      mvarversion = XVal(!version & "")
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvardiagram_id), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
