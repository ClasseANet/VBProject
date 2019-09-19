VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_COMPOS_PROD"
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

Private mvarIDPAI   As Double
Private mvarIDFILHO As Double
Public Property Let IDPAI(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAI <> vData)
   mvarIDPAI = vData
End Property
Public Property Get IDPAI() As Double
   IDPAI = mvarIDPAI
End Property
Public Property Let IDFILHO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFILHO <> vData)
   mvarIDFILHO = vData
End Property
Public Property Get IDFILHO() As Double
   IDFILHO = mvarIDFILHO
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
 
   Sql = "Insert Into COMPOS_PROD (IDPAI, IDFILHO) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFILHO) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPAI, Optional Ch_IDFILHO) As String
   Dim Sql As String
 
   Sql = "Delete From COMPOS_PROD" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDPAI) Then Sql = Sql & " IDPAI = " & SqlNum(Cstr(Ch_IDPAI)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDFILHO) Then Sql = Sql & " IDFILHO = " & SqlNum(Cstr(Ch_IDFILHO)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update COMPOS_PROD Set " & VbNewLine 
   Sql = Sql & " IDPAI   = " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & " , IDFILHO = " & SqlNum(mvarIDFILHO) & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPAI     = " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & " And IDFILHO = " & SqlNum(mvarIDFILHO) & VbNewLine 
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
Public Property Get QrySelect( Optional Ch_IDPAI, Optional Ch_IDFILHO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPAI, IDFILHO" & VbNewLine 
   Sql = Sql &" From COMPOS_PROD" & VbNewLine 
   If IsMissing(Ch_WHERE) Then
      Sql = Sql & " Where " & VbNewLine 
      If IsMissing(Ch_IDPAI)  And IsMissing(Ch_IDFILHO)  Then 
         Sql = Sql & " IDPAI = " & SqlNum(mvarIDPAI) & VbNewLine 
         Sql = Sql & " And IDFILHO = " & SqlNum(mvarIDFILHO) & VbNewLine 
      Else 
         If Not isMissing(Ch_IDPAI) Then Sql = Sql & " IDPAI = " & SqlNum(Cstr(Ch_IDPAI)) & " AND " & VbNewLine 
         If Not isMissing(Ch_IDFILHO) Then Sql = Sql & " IDFILHO = " & SqlNum(Cstr(Ch_IDFILHO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDPAI, Optional Ch_IDFILHO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPAI, Ch_IDFILHO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPAI = XVal(!IDPAI & "")
      mvarIDFILHO = XVal(!IDFILHO & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPAI, mvarIDFILHO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDPAI As String, Ch_IDFILHO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update COMPOS_PROD Set "
   Sql = Sql & " IDPAI = " & SqlNum(Ch_IDPAI)
   Sql = Sql & " , IDFILHO = " & SqlNum(Ch_IDFILHO)
   Sql = Sql & " Where 
   Sql = Sql & " IDPAI = " & SqlNum(mvarIDPAI)
   Sql = Sql & " and IDFILHO = " & SqlNum(mvarIDFILHO)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
