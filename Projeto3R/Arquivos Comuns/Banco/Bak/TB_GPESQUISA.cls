VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GPESQUISA"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb            As Object 
Private mvarRS             As Object 

Private mvarQryInsert      As String
Private mvarQryUpDate      As String
Private mvarQryDelete      As String
Private mvarQrySelect      As String
Private mvarQrySave        As String

Private mvarisDirt         As Boolean

Private mvarIDPESQUISA   As Double
Private mvarCODSIS       As String
Private mvarIDMODU       As String
Private mvarNOMEPESQUISA As String
Private mvarTIPOPESQUISA As String
Private mvarDSCPESQUISA  As String
Private mvarESCOPO       As Double
Private mvarPESQDEFAULT  As Double
Private mvarPESQSQL      As String
Private mvarIDUSU        As String
Private mvarIDCONEXAO    As Double
Private mvarTAGCAMPOS    As String
Public Property Get IDPESQUISA() As Double
   IDPESQUISA = mvarIDPESQUISA
End Property
Public Property Let CODSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIS <> vData)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let IDMODU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMODU <> vData)
   mvarIDMODU = vData
End Property
Public Property Get IDMODU() As String
   IDMODU = mvarIDMODU
End Property
Public Property Let NOMEPESQUISA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNOMEPESQUISA <> vData)
   mvarNOMEPESQUISA = vData
End Property
Public Property Get NOMEPESQUISA() As String
   NOMEPESQUISA = mvarNOMEPESQUISA
End Property
Public Property Let TIPOPESQUISA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPOPESQUISA <> vData)
   mvarTIPOPESQUISA = vData
End Property
Public Property Get TIPOPESQUISA() As String
   TIPOPESQUISA = mvarTIPOPESQUISA
End Property
Public Property Let DSCPESQUISA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCPESQUISA <> vData)
   mvarDSCPESQUISA = vData
End Property
Public Property Get DSCPESQUISA() As String
   DSCPESQUISA = mvarDSCPESQUISA
End Property
Public Property Let ESCOPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarESCOPO <> vData)
   mvarESCOPO = vData
End Property
Public Property Get ESCOPO() As Double
   ESCOPO = mvarESCOPO
End Property
Public Property Let PESQDEFAULT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPESQDEFAULT <> vData)
   mvarPESQDEFAULT = vData
End Property
Public Property Get PESQDEFAULT() As Double
   PESQDEFAULT = mvarPESQDEFAULT
End Property
Public Property Let PESQSQL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPESQSQL <> vData)
   mvarPESQSQL = vData
End Property
Public Property Get PESQSQL() As String
   PESQSQL = mvarPESQSQL
End Property
Public Property Let IDUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSU <> vData)
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let IDCONEXAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCONEXAO <> vData)
   mvarIDCONEXAO = vData
End Property
Public Property Get IDCONEXAO() As Double
   IDCONEXAO = mvarIDCONEXAO
End Property
Public Property Let TAGCAMPOS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTAGCAMPOS <> vData)
   mvarTAGCAMPOS = vData
End Property
Public Property Get TAGCAMPOS() As String
   TAGCAMPOS = mvarTAGCAMPOS
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
 
   Sql = "Insert Into GPESQUISA (CODSIS, IDMODU, NOMEPESQUISA" & VbNewLine 
   Sql = Sql & ", TIPOPESQUISA, DSCPESQUISA, ESCOPO, PESQDEFAULT, PESQSQL" & VbNewLine 
   Sql = Sql & ", IDUSU, IDCONEXAO, TAGCAMPOS) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNOMEPESQUISA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPOPESQUISA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCPESQUISA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarESCOPO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPESQDEFAULT) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPESQSQL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCONEXAO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTAGCAMPOS) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPESQUISA) As String
   Dim Sql As String
 
   Sql = "Delete From GPESQUISA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDPESQUISA) Then Sql = Sql & " IDPESQUISA = " & SqlNum(Cstr(Ch_IDPESQUISA)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GPESQUISA Set " & VbNewLine 
   Sql = Sql & " CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , IDMODU       = " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & " , NOMEPESQUISA = " & SqlStr(mvarNOMEPESQUISA) & VbNewLine 
   Sql = Sql & " , TIPOPESQUISA = " & SqlStr(mvarTIPOPESQUISA) & VbNewLine 
   Sql = Sql & " , DSCPESQUISA  = " & SqlStr(mvarDSCPESQUISA) & VbNewLine 
   Sql = Sql & " , ESCOPO       = " & SqlNum(mvarESCOPO) & VbNewLine 
   Sql = Sql & " , PESQDEFAULT  = " & SqlNum(mvarPESQDEFAULT) & VbNewLine 
   Sql = Sql & " , PESQSQL      = " & SqlStr(mvarPESQSQL) & VbNewLine 
   Sql = Sql & " , IDUSU        = " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & " , IDCONEXAO    = " & SqlNum(mvarIDCONEXAO) & VbNewLine 
   Sql = Sql & " , TAGCAMPOS    = " & SqlStr(mvarTAGCAMPOS) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GPESQUISA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPESQUISA     = " & SqlNum(mvarIDPESQUISA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDPESQUISA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPESQUISA, CODSIS, IDMODU, NOMEPESQUISA" & VbNewLine 
   Sql = Sql & ", TIPOPESQUISA, DSCPESQUISA, ESCOPO, PESQDEFAULT, PESQSQL" & VbNewLine 
   Sql = Sql & ", IDUSU, IDCONEXAO, TAGCAMPOS" & VbNewLine 
   Sql = Sql &" From GPESQUISA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPESQUISA)  Then 
      Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDPESQUISA) Then Sql = Sql & " IDPESQUISA = " & SqlNum(Cstr(Ch_IDPESQUISA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDPESQUISA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPESQUISA, Ch_WHERE, Ch_ORDERBY)
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
      mvarCODSIS = !CODSIS & ""
      mvarIDMODU = !IDMODU & ""
      mvarNOMEPESQUISA = !NOMEPESQUISA & ""
      mvarTIPOPESQUISA = !TIPOPESQUISA & ""
      mvarDSCPESQUISA = !DSCPESQUISA & ""
      mvarESCOPO = XVal(!ESCOPO & "")
      mvarPESQDEFAULT = XVal(!PESQDEFAULT & "")
      mvarPESQSQL = !PESQSQL & ""
      mvarIDUSU = !IDUSU & ""
      mvarIDCONEXAO = XVal(!IDCONEXAO & "")
      mvarTAGCAMPOS = !TAGCAMPOS & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPESQUISA), ComCOMMIT)
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
