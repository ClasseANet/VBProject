VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_FRECIBO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb               As Object 
Private mvarRS                As Object 

Private mvarQryInsert         As String
Private mvarQryUpDate         As String
Private mvarQryDelete         As String
Private mvarQrySelect         As String
Private mvarQrySave           As String

Private mvarisDirt            As Boolean

Private mvarIDLOJA          As Double
Private mvarSERIE           As String
Private mvarIDRECIBO        As Double
Private mvarIDVENDA         As Double
Private mvarDTEMISSAO       As String
Private mvarVLTOTAL         As Double
Private mvarTIPO            As Double
Private mvarIDLOTE          As Double
Private mvarSTATUS          As Double
Private mvarNATUREZAOP      As Double
Private mvarASSINATURA      As String
Private mvarREGESPECIALTRIB As Double
Private mvarSIMPLES         As Double
Private mvarIDRECIBO0       As Double
Private mvarSERIE0          As String
Private mvarTRIBUTACAO      As Double
Private mvarINCENTIVOCULT   As Double
Private mvarCODSERVFEDERAL  As String
Private mvarCODSERVMUNIC    As String
Private mvarALTERSTAMP      As Double
Private mvarTIMESTAMP       As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let SERIE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSERIE <> vData)
   mvarSERIE = vData
End Property
Public Property Get SERIE() As String
   SERIE = mvarSERIE
End Property
Public Property Let IDRECIBO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDRECIBO <> vData)
   mvarIDRECIBO = vData
End Property
Public Property Get IDRECIBO() As Double
   IDRECIBO = mvarIDRECIBO
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let DTEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEMISSAO <> vData)
   mvarDTEMISSAO = vData
End Property
Public Property Get DTEMISSAO() As String
   DTEMISSAO = mvarDTEMISSAO
End Property
Public Property Let VLTOTAL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLTOTAL <> vData)
   mvarVLTOTAL = vData
End Property
Public Property Get VLTOTAL() As Double
   VLTOTAL = mvarVLTOTAL
End Property
Public Property Let TIPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPO <> vData)
   mvarTIPO = vData
End Property
Public Property Get TIPO() As Double
   TIPO = mvarTIPO
End Property
Public Property Let IDLOTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOTE <> vData)
   mvarIDLOTE = vData
End Property
Public Property Get IDLOTE() As Double
   IDLOTE = mvarIDLOTE
End Property
Public Property Let STATUS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSTATUS <> vData)
   mvarSTATUS = vData
End Property
Public Property Get STATUS() As Double
   STATUS = mvarSTATUS
End Property
Public Property Let NATUREZAOP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNATUREZAOP <> vData)
   mvarNATUREZAOP = vData
End Property
Public Property Get NATUREZAOP() As Double
   NATUREZAOP = mvarNATUREZAOP
End Property
Public Property Let ASSINATURA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarASSINATURA <> vData)
   mvarASSINATURA = vData
End Property
Public Property Get ASSINATURA() As String
   ASSINATURA = mvarASSINATURA
End Property
Public Property Let REGESPECIALTRIB(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarREGESPECIALTRIB <> vData)
   mvarREGESPECIALTRIB = vData
End Property
Public Property Get REGESPECIALTRIB() As Double
   REGESPECIALTRIB = mvarREGESPECIALTRIB
End Property
Public Property Let SIMPLES(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSIMPLES <> vData)
   mvarSIMPLES = vData
End Property
Public Property Get SIMPLES() As Double
   SIMPLES = mvarSIMPLES
End Property
Public Property Let IDRECIBO0(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDRECIBO0 <> vData)
   mvarIDRECIBO0 = vData
End Property
Public Property Get IDRECIBO0() As Double
   IDRECIBO0 = mvarIDRECIBO0
End Property
Public Property Let SERIE0(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSERIE0 <> vData)
   mvarSERIE0 = vData
End Property
Public Property Get SERIE0() As String
   SERIE0 = mvarSERIE0
End Property
Public Property Let TRIBUTACAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTRIBUTACAO <> vData)
   mvarTRIBUTACAO = vData
End Property
Public Property Get TRIBUTACAO() As Double
   TRIBUTACAO = mvarTRIBUTACAO
End Property
Public Property Let INCENTIVOCULT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarINCENTIVOCULT <> vData)
   mvarINCENTIVOCULT = vData
End Property
Public Property Get INCENTIVOCULT() As Double
   INCENTIVOCULT = mvarINCENTIVOCULT
End Property
Public Property Let CODSERVFEDERAL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSERVFEDERAL <> vData)
   mvarCODSERVFEDERAL = vData
End Property
Public Property Get CODSERVFEDERAL() As String
   CODSERVFEDERAL = mvarCODSERVFEDERAL
End Property
Public Property Let CODSERVMUNIC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSERVMUNIC <> vData)
   mvarCODSERVMUNIC = vData
End Property
Public Property Get CODSERVMUNIC() As String
   CODSERVMUNIC = mvarCODSERVMUNIC
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update FRECIBO Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA            = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And SERIE         = " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & " And IDRECIBO      = " & SqlNum(mvarIDRECIBO) & VbNewLine 
   If Not mvarXDb Is Nothing Then
      If mvarXDb.Conectado Then
         If mvarXDb.Executa(Sql, True) Then
            mvarALTERSTAMP = vData
         End If
      End If
   End If
End Property
Public Property Get ALTERSTAMP() As Double
   ALTERSTAMP = mvarALTERSTAMP
End Property
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
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
 
   Sql = "Insert Into FRECIBO (IDLOJA, SERIE, IDRECIBO, IDVENDA" & VbNewLine 
   Sql = Sql & ", DTEMISSAO, VLTOTAL, TIPO, IDLOTE, STATUS" & VbNewLine 
   Sql = Sql & ", NATUREZAOP, ASSINATURA, REGESPECIALTRIB, SIMPLES, IDRECIBO0" & VbNewLine 
   Sql = Sql & ", SERIE0, TRIBUTACAO, INCENTIVOCULT, CODSERVFEDERAL, CODSERVMUNIC" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSERIE) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDRECIBO),0)+1 From FRECIBO Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And SERIE = " & SqlStr(mvarSERIE) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDRECIBO) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLTOTAL) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTIPO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDLOTE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSTATUS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarNATUREZAOP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarASSINATURA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarREGESPECIALTRIB) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSIMPLES) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDRECIBO0) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSERIE0) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTRIBUTACAO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarINCENTIVOCULT) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODSERVFEDERAL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODSERVMUNIC) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_IDRECIBO) As String
   Dim Sql As String
 
   Sql = "Delete From FRECIBO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_SERIE) Then Sql = Sql & " SERIE = " & SqlStr(Cstr(Ch_SERIE)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDRECIBO) Then Sql = Sql & " IDRECIBO = " & SqlNum(Cstr(Ch_IDRECIBO)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   MyDelRow.Query = Sql & ";"
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update FRECIBO Set " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , SERIE           = " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & " , IDRECIBO        = " & SqlNum(mvarIDRECIBO) & VbNewLine 
   Sql = Sql & " , IDVENDA         = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " , DTEMISSAO       = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VLTOTAL         = " & SqlNum(mvarVLTOTAL) & VbNewLine 
   Sql = Sql & " , TIPO            = " & SqlNum(mvarTIPO) & VbNewLine 
   Sql = Sql & " , IDLOTE          = " & SqlNum(mvarIDLOTE) & VbNewLine 
   Sql = Sql & " , STATUS          = " & SqlNum(mvarSTATUS) & VbNewLine 
   Sql = Sql & " , NATUREZAOP      = " & SqlNum(mvarNATUREZAOP) & VbNewLine 
   Sql = Sql & " , ASSINATURA      = " & SqlStr(mvarASSINATURA) & VbNewLine 
   Sql = Sql & " , REGESPECIALTRIB = " & SqlNum(mvarREGESPECIALTRIB) & VbNewLine 
   Sql = Sql & " , SIMPLES         = " & SqlNum(mvarSIMPLES) & VbNewLine 
   Sql = Sql & " , IDRECIBO0       = " & SqlNum(mvarIDRECIBO0) & VbNewLine 
   Sql = Sql & " , SERIE0          = " & SqlStr(mvarSERIE0) & VbNewLine 
   Sql = Sql & " , TRIBUTACAO      = " & SqlNum(mvarTRIBUTACAO) & VbNewLine 
   Sql = Sql & " , INCENTIVOCULT   = " & SqlNum(mvarINCENTIVOCULT) & VbNewLine 
   Sql = Sql & " , CODSERVFEDERAL  = " & SqlStr(mvarCODSERVFEDERAL) & VbNewLine 
   Sql = Sql & " , CODSERVMUNIC    = " & SqlStr(mvarCODSERVMUNIC) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP      = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP       = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FRECIBO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA            = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And SERIE         = " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & " And IDRECIBO      = " & SqlNum(mvarIDRECIBO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_IDRECIBO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, SERIE, IDRECIBO, IDVENDA" & VbNewLine 
   Sql = Sql & ", DTEMISSAO, VLTOTAL, TIPO, IDLOTE, STATUS" & VbNewLine 
   Sql = Sql & ", NATUREZAOP, ASSINATURA, REGESPECIALTRIB, SIMPLES, IDRECIBO0" & VbNewLine 
   Sql = Sql & ", SERIE0, TRIBUTACAO, INCENTIVOCULT, CODSERVFEDERAL, CODSERVMUNIC" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From FRECIBO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_SERIE)  And IsMissing(Ch_IDRECIBO)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " SERIE = " & SqlStr(mvarSERIE) & " AND " & VbNewLine 
      Sql = Sql & " IDRECIBO = " & SqlNum(mvarIDRECIBO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_SERIE) Then Sql = Sql & " SERIE = " & SqlStr(Cstr(Ch_SERIE)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDRECIBO) Then Sql = Sql & " IDRECIBO = " & SqlNum(Cstr(Ch_IDRECIBO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_IDRECIBO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_SERIE, Ch_IDRECIBO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDLOJA = XVal(!IDLOJA & "")
      mvarSERIE = !SERIE & ""
      mvarIDRECIBO = XVal(!IDRECIBO & "")
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarVLTOTAL = XVal(!VLTOTAL & "")
      mvarTIPO = XVal(!TIPO & "")
      mvarIDLOTE = XVal(!IDLOTE & "")
      mvarSTATUS = XVal(!STATUS & "")
      mvarNATUREZAOP = XVal(!NATUREZAOP & "")
      mvarASSINATURA = !ASSINATURA & ""
      mvarREGESPECIALTRIB = XVal(!REGESPECIALTRIB & "")
      mvarSIMPLES = XVal(!SIMPLES & "")
      mvarIDRECIBO0 = XVal(!IDRECIBO0 & "")
      mvarSERIE0 = !SERIE0 & ""
      mvarTRIBUTACAO = XVal(!TRIBUTACAO & "")
      mvarINCENTIVOCULT = XVal(!INCENTIVOCULT & "")
      mvarCODSERVFEDERAL = !CODSERVFEDERAL & ""
      mvarCODSERVMUNIC = !CODSERVMUNIC & ""
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarSERIE, mvarIDRECIBO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_SERIE As String, Ch_IDRECIBO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FRECIBO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , SERIE = " & SqlStr(Ch_SERIE)
   Sql = Sql & " , IDRECIBO = " & SqlNum(Ch_IDRECIBO)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and SERIE = " & SqlStr(mvarSERIE)
   Sql = Sql & " and IDRECIBO = " & SqlNum(mvarIDRECIBO)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
