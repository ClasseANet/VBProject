VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GPESQUISA_PESQFERR"
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

Private mvarIDPESQUISA As Double
Private mvarIDFERR     As Double
Public Property Let IDPESQUISA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPESQUISA <> vData)
   mvarIDPESQUISA = vData
End Property
Public Property Get IDPESQUISA() As Double
   IDPESQUISA = mvarIDPESQUISA
End Property
Public Property Let IDFERR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFERR <> vData)
   mvarIDFERR = vData
End Property
Public Property Get IDFERR() As Double
   IDFERR = mvarIDFERR
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
 
   Sql = "Insert Into GPESQUISA_PESQFERR (IDPESQUISA, IDFERR) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   If pAutoId Then
      Sql = Sql &  "(Select isNull(Max(IDPESQUISA),0)+1 From GPESQUISA_PESQFERR Where IDFERR = " & SqlNum(mvarIDFERR) & ")" & vbNewLine
   Else
      Sql = Sql &  SqlNum(mvarIDPESQUISA) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDFERR) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPESQUISA, Optional Ch_IDFERR) As String
   Dim Sql As String
 
   Sql = "Delete From GPESQUISA_PESQFERR" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDPESQUISA) Then Sql = Sql & " IDPESQUISA = " & SqlNum(Cstr(Ch_IDPESQUISA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDFERR) Then Sql = Sql & " IDFERR = " & SqlNum(Cstr(Ch_IDFERR)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GPESQUISA_PESQFERR Set " & VbNewLine 
   Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA) & VbNewLine 
   Sql = Sql & " , IDFERR     = " & SqlNum(mvarIDFERR) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GPESQUISA_PESQFERR'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPESQUISA   = " & SqlNum(mvarIDPESQUISA) & VbNewLine 
   Sql = Sql & " And IDFERR   = " & SqlNum(mvarIDFERR) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDPESQUISA, Optional Ch_IDFERR, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPESQUISA, IDFERR" & VbNewLine 
   Sql = Sql &" From GPESQUISA_PESQFERR" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPESQUISA)  And IsMissing(Ch_IDFERR)  Then 
      Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA) & " AND " & VbNewLine 
      Sql = Sql & " IDFERR = " & SqlNum(mvarIDFERR) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDPESQUISA) Then Sql = Sql & " IDPESQUISA = " & SqlNum(Cstr(Ch_IDPESQUISA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFERR) Then Sql = Sql & " IDFERR = " & SqlNum(Cstr(Ch_IDFERR)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDPESQUISA, Optional Ch_IDFERR, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPESQUISA, Ch_IDFERR, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPESQUISA = XVal(!IDPESQUISA & "")
      mvarIDFERR = XVal(!IDFERR & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPESQUISA, mvarIDFERR), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDPESQUISA As String, Ch_IDFERR As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GPESQUISA_PESQFERR Set "
   Sql = Sql & " IDPESQUISA = " & SqlNum(Ch_IDPESQUISA)
   Sql = Sql & " , IDFERR = " & SqlNum(Ch_IDFERR)
   Sql = Sql & " Where 
   Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA)
   Sql = Sql & " and IDFERR = " & SqlNum(mvarIDFERR)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
