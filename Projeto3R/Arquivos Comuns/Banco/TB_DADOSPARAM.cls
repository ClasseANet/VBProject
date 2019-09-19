VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_DADOSPARAM"
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

Private mvarCODSIS     As String
Private mvarIDLOJA     As Double
Private mvarCODPARAM   As String
Private mvarCODDADO    As String
Private mvarDSCDADO    As String
Private mvarTIPODADO   As String
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let CODSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIS <> vData)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let CODPARAM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODPARAM <> vData)
   mvarCODPARAM = vData
End Property
Public Property Get CODPARAM() As String
   CODPARAM = mvarCODPARAM
End Property
Public Property Let CODDADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODDADO <> vData)
   mvarCODDADO = vData
End Property
Public Property Get CODDADO() As String
   CODDADO = mvarCODDADO
End Property
Public Property Let DSCDADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCDADO <> vData)
   mvarDSCDADO = vData
End Property
Public Property Get DSCDADO() As String
   DSCDADO = mvarDSCDADO
End Property
Public Property Let TIPODADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPODADO <> vData)
   mvarTIPODADO = vData
End Property
Public Property Get TIPODADO() As String
   TIPODADO = mvarTIPODADO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update DADOSPARAM Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " And IDLOJA   = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And CODPARAM = " & SqlStr(mvarCODPARAM) & VbNewLine 
   Sql = Sql & " And CODDADO  = " & SqlStr(mvarCODDADO) & VbNewLine 
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
Public Property Let TIMESTAMP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIMESTAMP <> vData)
   mvarTIMESTAMP = vData
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into DADOSPARAM (CODSIS, IDLOJA, CODPARAM, CODDADO" & VbNewLine 
   Sql = Sql & ", DSCDADO, TIPODADO, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODPARAM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODDADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCDADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPODADO) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_CODSIS, Optional Ch_IDLOJA, Optional Ch_CODPARAM, Optional Ch_CODDADO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From DADOSPARAM" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_CODSIS)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_CODPARAM)  And IsMissing(Ch_CODDADO)  Then 
      Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS) & " AND " & VbNewLine 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " CODPARAM = " & SqlStr(mvarCODPARAM) & " AND " & VbNewLine 
      Sql = Sql & " CODDADO = " & SqlStr(mvarCODDADO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_CODPARAM) Then Sql = Sql & " CODPARAM = " & SqlStr(Cstr(Ch_CODPARAM)) & " AND " & VbNewLine 
      If Not isMissing(Ch_CODDADO) Then Sql = Sql & " CODDADO = " & SqlStr(Cstr(Ch_CODDADO)) & " AND " & VbNewLine 
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

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = "
   sTag = sTag & "|CODSIS = " & mvarCODSIS
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|CODPARAM = " & mvarCODPARAM
   sTag = sTag & "|CODDADO = " & mvarCODDADO
   sTag = sTag & "|DSCDADO = " & mvarDSCDADO
   sTag = sTag & "|TIPODADO = " & mvarTIPODADO
   sTag = sTag & "|Where = " & IIf(IsMissing(Ch_WHERE), "", Ch_WHERE)
   sTag = sTag & "|"

   MyDelRow.Query = Sql & ";"
   MyDelRow.Tag = sTag
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update DADOSPARAM Set " & VbNewLine 
   Sql = Sql & " CODSIS     = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , CODPARAM   = " & SqlStr(mvarCODPARAM) & VbNewLine 
   Sql = Sql & " , CODDADO    = " & SqlStr(mvarCODDADO) & VbNewLine 
   Sql = Sql & " , DSCDADO    = " & SqlStr(mvarDSCDADO) & VbNewLine 
   Sql = Sql & " , TIPODADO   = " & SqlStr(mvarTIPODADO) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP  = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP  = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'DADOSPARAM'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " And IDLOJA   = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And CODPARAM = " & SqlStr(mvarCODPARAM) & VbNewLine 
   Sql = Sql & " And CODDADO  = " & SqlStr(mvarCODDADO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_CODSIS, Optional Ch_IDLOJA, Optional Ch_CODPARAM, Optional Ch_CODDADO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select CODSIS, IDLOJA, CODPARAM, CODDADO" & VbNewLine 
   Sql = Sql & ", DSCDADO, TIPODADO, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From DADOSPARAM" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_CODSIS)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_CODPARAM)  And IsMissing(Ch_CODDADO)  Then 
      Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS) & " AND " & VbNewLine 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " CODPARAM = " & SqlStr(mvarCODPARAM) & " AND " & VbNewLine 
      Sql = Sql & " CODDADO = " & SqlStr(mvarCODDADO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_CODPARAM) Then Sql = Sql & " CODPARAM = " & SqlStr(Cstr(Ch_CODPARAM)) & " AND " & VbNewLine 
      If Not isMissing(Ch_CODDADO) Then Sql = Sql & " CODDADO = " & SqlStr(Cstr(Ch_CODDADO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_CODSIS, Optional Ch_IDLOJA, Optional Ch_CODPARAM, Optional Ch_CODDADO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_CODSIS, Ch_IDLOJA, Ch_CODPARAM, Ch_CODDADO, Ch_WHERE, Ch_ORDERBY)
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
      mvarCODSIS = !CODSIS & ""
      mvarIDLOJA = XVal(!IDLOJA & "")
      mvarCODPARAM = !CODPARAM & ""
      mvarCODDADO = !CODDADO & ""
      mvarDSCDADO = !DSCDADO & ""
      mvarTIPODADO = !TIPODADO & ""
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarCODSIS = ""
   mvarIDLOJA = 0
   mvarCODPARAM = ""
   mvarCODDADO = ""
   mvarDSCDADO = ""
   mvarTIPODADO = ""
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarCODSIS, mvarIDLOJA, mvarCODPARAM, mvarCODDADO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_CODSIS As String, Ch_IDLOJA As String, Ch_CODPARAM As String, Ch_CODDADO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update DADOSPARAM Set "
   Sql = Sql & " CODSIS = " & SqlStr(Ch_CODSIS)
   Sql = Sql & " , IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , CODPARAM = " & SqlStr(Ch_CODPARAM)
   Sql = Sql & " , CODDADO = " & SqlStr(Ch_CODDADO)
   Sql = Sql & " Where 
   Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS)
   Sql = Sql & " and IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and CODPARAM = " & SqlStr(mvarCODPARAM)
   Sql = Sql & " and CODDADO = " & SqlStr(mvarCODDADO)
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