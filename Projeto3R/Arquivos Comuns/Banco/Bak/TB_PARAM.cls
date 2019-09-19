VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_PARAM"
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
Private mvarDSCPARAM   As String
Private mvarVLPARAM    As String
Private mvarTIPOPARAM  As String
Private mvarCATPARAM   As String
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
Public Property Let DSCPARAM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCPARAM <> vData)
   mvarDSCPARAM = vData
End Property
Public Property Get DSCPARAM() As String
   DSCPARAM = mvarDSCPARAM
End Property
Public Property Let VLPARAM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVLPARAM <> vData)
   mvarVLPARAM = vData
End Property
Public Property Get VLPARAM() As String
   VLPARAM = mvarVLPARAM
End Property
Public Property Let TIPOPARAM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPOPARAM <> vData)
   mvarTIPOPARAM = vData
End Property
Public Property Get TIPOPARAM() As String
   TIPOPARAM = mvarTIPOPARAM
End Property
Public Property Let CATPARAM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCATPARAM <> vData)
   mvarCATPARAM = vData
End Property
Public Property Get CATPARAM() As String
   CATPARAM = mvarCATPARAM
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update PARAM Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " And IDLOJA   = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And CODPARAM = " & SqlStr(mvarCODPARAM) & VbNewLine 
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
 
   Sql = "Insert Into PARAM (CODSIS, IDLOJA, CODPARAM, DSCPARAM" & VbNewLine 
   Sql = Sql & ", VLPARAM, TIPOPARAM, CATPARAM, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarCODSIS) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDLOJA),0)+1 From PARAM Where CODSIS = " & SqlStr(mvarCODSIS) & " And CODPARAM = " & SqlStr(mvarCODPARAM) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDLOJA) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarCODPARAM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCPARAM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVLPARAM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPOPARAM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCATPARAM) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_CODSIS, Optional Ch_IDLOJA, Optional Ch_CODPARAM) As String
   Dim Sql As String
 
   Sql = "Delete From PARAM" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_CODPARAM) Then Sql = Sql & " CODPARAM = " & SqlStr(Cstr(Ch_CODPARAM)) & " AND " & VbNewLine 
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
 
   Sql = "Update PARAM Set " & VbNewLine 
   Sql = Sql & " CODSIS     = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , CODPARAM   = " & SqlStr(mvarCODPARAM) & VbNewLine 
   Sql = Sql & " , DSCPARAM   = " & SqlStr(mvarDSCPARAM) & VbNewLine 
   Sql = Sql & " , VLPARAM    = " & SqlStr(mvarVLPARAM) & VbNewLine 
   Sql = Sql & " , TIPOPARAM  = " & SqlStr(mvarTIPOPARAM) & VbNewLine 
   Sql = Sql & " , CATPARAM   = " & SqlStr(mvarCATPARAM) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP  = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'PARAM'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " And IDLOJA   = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And CODPARAM = " & SqlStr(mvarCODPARAM) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_CODSIS, Optional Ch_IDLOJA, Optional Ch_CODPARAM, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select CODSIS, IDLOJA, CODPARAM, DSCPARAM" & VbNewLine 
   Sql = Sql & ", VLPARAM, TIPOPARAM, CATPARAM, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From PARAM" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_CODSIS)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_CODPARAM)  Then 
      Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS) & " AND " & VbNewLine 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " CODPARAM = " & SqlStr(mvarCODPARAM) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_CODPARAM) Then Sql = Sql & " CODPARAM = " & SqlStr(Cstr(Ch_CODPARAM)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_CODSIS, Optional Ch_IDLOJA, Optional Ch_CODPARAM, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_CODSIS, Ch_IDLOJA, Ch_CODPARAM, Ch_WHERE, Ch_ORDERBY)
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
      mvarDSCPARAM = !DSCPARAM & ""
      mvarVLPARAM = !VLPARAM & ""
      mvarTIPOPARAM = !TIPOPARAM & ""
      mvarCATPARAM = !CATPARAM & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarCODSIS, mvarIDLOJA, mvarCODPARAM), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_CODSIS As String, Ch_IDLOJA As String, Ch_CODPARAM As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update PARAM Set "
   Sql = Sql & " CODSIS = " & SqlStr(Ch_CODSIS)
   Sql = Sql & " , IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , CODPARAM = " & SqlStr(Ch_CODPARAM)
   Sql = Sql & " Where 
   Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS)
   Sql = Sql & " and IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and CODPARAM = " & SqlStr(mvarCODPARAM)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
