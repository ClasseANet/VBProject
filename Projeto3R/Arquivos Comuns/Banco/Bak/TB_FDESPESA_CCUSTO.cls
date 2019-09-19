VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_FDESPESA_CCUSTO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb         As Object 
Private mvarRS          As Object 

Private mvarQryInsert   As String
Private mvarQryUpDate   As String
Private mvarQryDelete   As String
Private mvarQrySelect   As String
Private mvarQrySave     As String

Private mvarisDirt      As Boolean

Private mvarIDCCUSTO As Double
Private mvarIDDESP   As Double
Public Property Let IDCCUSTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCCUSTO <> vData)
   mvarIDCCUSTO = vData
End Property
Public Property Get IDCCUSTO() As Double
   IDCCUSTO = mvarIDCCUSTO
End Property
Public Property Let IDDESP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDESP <> vData)
   mvarIDDESP = vData
End Property
Public Property Get IDDESP() As Double
   IDDESP = mvarIDDESP
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True) As String
   Dim Sql As String
 
   Sql = "Insert Into FDESPESA_CCUSTO (IDCCUSTO, IDDESP) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDCCUSTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDDESP) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDCCUSTO, Optional Ch_IDDESP) As String
   Dim Sql As String
 
   Sql = "Delete From FDESPESA_CCUSTO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDCCUSTO) Then Sql = Sql & " IDCCUSTO = " & SqlNum(Cstr(Ch_IDCCUSTO)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDDESP) Then Sql = Sql & " IDDESP = " & SqlNum(Cstr(Ch_IDDESP)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update FDESPESA_CCUSTO Set " & VbNewLine 
   Sql = Sql & " IDCCUSTO = " & SqlNum(mvarIDCCUSTO) & VbNewLine 
   Sql = Sql & " , IDDESP   = " & SqlNum(mvarIDDESP) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FDESPESA_CCUSTO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDCCUSTO   = " & SqlNum(mvarIDCCUSTO) & VbNewLine 
   Sql = Sql & " And IDDESP = " & SqlNum(mvarIDDESP) & VbNewLine 

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId as Boolean = True) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_IDCCUSTO, Optional Ch_IDDESP, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDCCUSTO, IDDESP" & VbNewLine 
   Sql = Sql &" From FDESPESA_CCUSTO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDCCUSTO)  And IsMissing(Ch_IDDESP)  Then 
      Sql = Sql & " IDCCUSTO = " & SqlNum(mvarIDCCUSTO) & " AND " & VbNewLine 
      Sql = Sql & " IDDESP = " & SqlNum(mvarIDDESP) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDCCUSTO) Then Sql = Sql & " IDCCUSTO = " & SqlNum(Cstr(Ch_IDCCUSTO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDDESP) Then Sql = Sql & " IDDESP = " & SqlNum(Cstr(Ch_IDDESP)) & " AND " & VbNewLine 
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & VbNewLine
      End If
   End If 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY 

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar( Optional Ch_IDCCUSTO, Optional Ch_IDDESP, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDCCUSTO, Ch_IDDESP, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCCUSTO = XVal(!IDCCUSTO & "")
      mvarIDDESP = XVal(!IDDESP & "")
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True, Optional pAutoId as Boolean = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId as Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDCCUSTO, mvarIDDESP), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDCCUSTO As String, Ch_IDDESP As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FDESPESA_CCUSTO Set "
   Sql = Sql & " IDCCUSTO = " & SqlNum(Ch_IDCCUSTO)
   Sql = Sql & " , IDDESP = " & SqlNum(Ch_IDDESP)
   Sql = Sql & " Where 
   Sql = Sql & " IDCCUSTO = " & SqlNum(mvarIDCCUSTO)
   Sql = Sql & " and IDDESP = " & SqlNum(mvarIDDESP)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
