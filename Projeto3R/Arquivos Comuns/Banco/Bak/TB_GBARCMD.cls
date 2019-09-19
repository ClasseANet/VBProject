VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GBARCMD"
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

Private mvarIDMODU  As String
Private mvarCODSIS  As String
Private mvarDSCMODU As String
Private mvarGRUPO   As String
Private mvarORDEM   As Double
Private mvarIMAGEM  As Object
Public Property Let IDMODU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMODU <> vData)
   mvarIDMODU = vData
End Property
Public Property Get IDMODU() As String
   IDMODU = mvarIDMODU
End Property
Public Property Let CODSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIS <> vData)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let DSCMODU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCMODU <> vData)
   mvarDSCMODU = vData
End Property
Public Property Get DSCMODU() As String
   DSCMODU = mvarDSCMODU
End Property
Public Property Let GRUPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarGRUPO <> vData)
   mvarGRUPO = vData
End Property
Public Property Get GRUPO() As String
   GRUPO = mvarGRUPO
End Property
Public Property Let ORDEM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarORDEM <> vData)
   mvarORDEM = vData
End Property
Public Property Get ORDEM() As Double
   ORDEM = mvarORDEM
End Property
Public Property Set IMAGEM(ByVal vData As Object)
   If Not mvarisDirt Then mvarisDirt = (mvarIMAGEM <> vData)
   Set mvarIMAGEM = vData
End Property
Public Property Get IMAGEM() As Object
   Set IMAGEM = mvarIMAGEM
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
 
   Sql = "Insert Into GBARCMD (IDMODU, CODSIS, DSCMODU, GRUPO" & VbNewLine 
   Sql = Sql & ", ORDEM) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCMODU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarGRUPO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarORDEM) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDMODU, Optional Ch_CODSIS) As String
   Dim Sql As String
 
   Sql = "Delete From GBARCMD" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDMODU) Then Sql = Sql & " IDMODU = " & SqlStr(Cstr(Ch_IDMODU)) & " AND " & VbNewLine 
   if Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GBARCMD Set " & VbNewLine 
   Sql = Sql & " IDMODU  = " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & " , CODSIS  = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , DSCMODU = " & SqlStr(mvarDSCMODU) & VbNewLine 
   Sql = Sql & " , GRUPO   = " & SqlStr(mvarGRUPO) & VbNewLine 
   Sql = Sql & " , ORDEM   = " & SqlNum(mvarORDEM) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GBARCMD'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDMODU    = " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & " And CODSIS = " & SqlStr(mvarCODSIS) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDMODU, Optional Ch_CODSIS, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDMODU, CODSIS, DSCMODU, GRUPO" & VbNewLine 
   Sql = Sql & ", ORDEM" & VbNewLine 
   Sql = Sql &" From GBARCMD" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDMODU)  And IsMissing(Ch_CODSIS)  Then 
      Sql = Sql & " IDMODU = " & SqlStr(mvarIDMODU) & " AND " & VbNewLine 
      Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDMODU) Then Sql = Sql & " IDMODU = " & SqlStr(Cstr(Ch_IDMODU)) & " AND " & VbNewLine 
      If Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDMODU, Optional Ch_CODSIS, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDMODU, Ch_CODSIS, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDMODU = !IDMODU & ""
      mvarCODSIS = !CODSIS & ""
      mvarDSCMODU = !DSCMODU & ""
      mvarGRUPO = !GRUPO & ""
      mvarORDEM = XVal(!ORDEM & "")
      mvarORDEM = XVal(!ORDEM & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDMODU, mvarCODSIS), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDMODU As String, Ch_CODSIS As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GBARCMD Set "
   Sql = Sql & " IDMODU = " & SqlStr(Ch_IDMODU)
   Sql = Sql & " , CODSIS = " & SqlStr(Ch_CODSIS)
   Sql = Sql & " Where 
   Sql = Sql & " IDMODU = " & SqlStr(mvarIDMODU)
   Sql = Sql & " and CODSIS = " & SqlStr(mvarCODSIS)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
