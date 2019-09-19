VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_MODULO_SISTEMA"
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

Private mvarID        As Double
Private mvarIDMODU    As String
Private mvarCODSIS    As String
Private mvarIDPAI     As Double
Private mvarMODUPAI   As String
Private mvarINDICE    As Double
Private mvarMENU      As String
Private mvarGRUPOMENU As Double
Public Property Let ID(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarID <> vData)
   mvarID = vData
End Property
Public Property Get ID() As Double
   ID = mvarID
End Property
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
Public Property Let IDPAI(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAI <> vData)
   mvarIDPAI = vData
End Property
Public Property Get IDPAI() As Double
   IDPAI = mvarIDPAI
End Property
Public Property Let MODUPAI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMODUPAI <> vData)
   mvarMODUPAI = vData
End Property
Public Property Get MODUPAI() As String
   MODUPAI = mvarMODUPAI
End Property
Public Property Let INDICE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarINDICE <> vData)
   mvarINDICE = vData
End Property
Public Property Get INDICE() As Double
   INDICE = mvarINDICE
End Property
Public Property Let MENU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMENU <> vData)
   mvarMENU = vData
End Property
Public Property Get MENU() As String
   MENU = mvarMENU
End Property
Public Property Let GRUPOMENU(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarGRUPOMENU <> vData)
   mvarGRUPOMENU = vData
End Property
Public Property Get GRUPOMENU() As Double
   GRUPOMENU = mvarGRUPOMENU
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into MODULO_SISTEMA (ID, IDMODU, CODSIS, IDPAI" & VbNewLine 
   Sql = Sql & ", MODUPAI, INDICE, MENU, GRUPOMENU) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarID) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMODUPAI) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarINDICE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMENU) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarGRUPOMENU) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDMODU, Optional Ch_CODSIS, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From MODULO_SISTEMA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
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
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update MODULO_SISTEMA Set " & VbNewLine 
   Sql = Sql & " ID        = " & SqlNum(mvarID) & VbNewLine 
   Sql = Sql & " , IDMODU    = " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & " , CODSIS    = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , IDPAI     = " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & " , MODUPAI   = " & SqlStr(mvarMODUPAI) & VbNewLine 
   Sql = Sql & " , INDICE    = " & SqlNum(mvarINDICE) & VbNewLine 
   Sql = Sql & " , MENU      = " & SqlStr(mvarMENU) & VbNewLine 
   Sql = Sql & " , GRUPOMENU = " & SqlNum(mvarGRUPOMENU) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'MODULO_SISTEMA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDMODU      = " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & " And CODSIS  = " & SqlStr(mvarCODSIS) & VbNewLine 

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate(pSinc:=pSinc)
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId, pSinc:=pSinc)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_IDMODU, Optional Ch_CODSIS, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select ID, IDMODU, CODSIS, IDPAI" & VbNewLine 
   Sql = Sql & ", MODUPAI, INDICE, MENU, GRUPOMENU" & VbNewLine 
   Sql = Sql &" From MODULO_SISTEMA" & VbNewLine 
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
      mvarID = XVal(!ID & "")
      mvarIDMODU = !IDMODU & ""
      mvarCODSIS = !CODSIS & ""
      mvarIDPAI = XVal(!IDPAI & "")
      mvarMODUPAI = !MODUPAI & ""
      mvarINDICE = XVal(!INDICE & "")
      mvarMENU = !MENU & ""
      mvarGRUPOMENU = XVal(!GRUPOMENU & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarID = 0
   mvarIDMODU = ""
   mvarCODSIS = ""
   mvarIDPAI = 0
   mvarMODUPAI = ""
   mvarINDICE = 0
   mvarMENU = ""
   mvarGRUPOMENU = 0

   On Error Resume Next
   Call Class_Initialize
End Sub
Public Function Salvar(Optional ComCOMMIT As Boolean = True, Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId, pSinc:=pSinc), ComCOMMIT)

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
 
   Sql = "Update MODULO_SISTEMA Set "
   Sql = Sql & " IDMODU = " & SqlStr(Ch_IDMODU)
   Sql = Sql & " , CODSIS = " & SqlStr(Ch_CODSIS)
   Sql = Sql & " Where 
   Sql = Sql & " IDMODU = " & SqlStr(mvarIDMODU)
   Sql = Sql & " and CODSIS = " & SqlStr(mvarCODSIS)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Initialize()
   Set mvarRS = Nothing
   Set mvarXDb = Nothing
   mvarQryInsert = ""
   mvarQryUpDate = ""
   mvarQryDelete = ""
   mvarQrySelect = ""
   mvarQrySave = ""
   mvarisDirt = False
End Sub
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
