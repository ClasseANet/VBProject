VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_CLASSEMAT"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb         As DS_BANCO
Private mvarRS          As Object 

Private mvarQryInsert   As String
Private mvarQryUpDate   As String
Private mvarQryDelete   As String
Private mvarQrySelect   As String
Private mvarQrySave     As String

Private mvarisDirt      As Boolean

Private mvarIDCLASSE   As Double
Private mvarDESCCLASSE As String
Public Property Let IDCLASSE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLASSE <> vData)
   mvarIDCLASSE = vData
End Property
Public Property Get IDCLASSE() As Double
   IDCLASSE = mvarIDCLASSE
End Property
Public Property Let DESCCLASSE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDESCCLASSE <> vData)
   mvarDESCCLASSE = vData
End Property
Public Property Get DESCCLASSE() As String
   DESCCLASSE = mvarDESCCLASSE
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
 
   Sql = "Insert Into CLASSEMAT (IDCLASSE, DESCCLASSE) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  "isNull(Max(IDCLASSE),0)+1" & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDESCCLASSE) & VbNewLine 
   Sql = Sql & " From CLASSEMAT" & VbNewLine 
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDCLASSE) As String
   Dim Sql As String
 
   Sql = "Delete From CLASSEMAT" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDCLASSE) Then Sql = Sql & " IDCLASSE = " & SqlNum(Cstr(Ch_IDCLASSE)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update CLASSEMAT Set " & VbNewLine 
   Sql = Sql & " IDCLASSE   = " & SqlNum(mvarIDCLASSE) & VbNewLine 
   Sql = Sql & " , DESCCLASSE = " & SqlStr(mvarDESCCLASSE) & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDCLASSE     = " & SqlNum(mvarIDCLASSE) & VbNewLine 
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
Public Property Get QrySelect( Optional Ch_IDCLASSE, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDCLASSE, DESCCLASSE" & VbNewLine 
   Sql = Sql &" From CLASSEMAT" & VbNewLine 
   If IsMissing(Ch_WHERE) Then
      Sql = Sql & " Where " & VbNewLine 
      If IsMissing(Ch_IDCLASSE)  Then 
         Sql = Sql & " IDCLASSE = " & SqlNum(mvarIDCLASSE) & VbNewLine 
      Else 
         If Not isMissing(Ch_IDCLASSE) Then Sql = Sql & " IDCLASSE = " & SqlNum(Cstr(Ch_IDCLASSE)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDCLASSE, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDCLASSE, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCLASSE = XVal(!IDCLASSE & "")
      mvarDESCCLASSE = !DESCCLASSE & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDCLASSE), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDCLASSE As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CLASSEMAT Set "
   Sql = Sql & " IDCLASSE = " & SqlNum(Ch_IDCLASSE)
   Sql = Sql & " Where 
   Sql = Sql & " IDCLASSE = " & SqlNum(mvarIDCLASSE)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
