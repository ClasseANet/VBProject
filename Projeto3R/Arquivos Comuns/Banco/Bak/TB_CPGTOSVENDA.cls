VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_CPGTOSVENDA"
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

Private mvarIDLOJA      As Double
Private mvarIDVENDA     As Double
Private mvarIDPGTO      As Double
Private mvarIDCONDPGTO  As Double
Private mvarIDFORMAPGTO As Double
Private mvarVLPGTO      As Double
Private mvarDTPGTO      As String
Private mvarTIMESTAMP   As String
Private mvarALTERSTAMP  As Double
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDPGTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPGTO <> vData)
   mvarIDPGTO = vData
End Property
Public Property Get IDPGTO() As Double
   IDPGTO = mvarIDPGTO
End Property
Public Property Let IDCONDPGTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCONDPGTO <> vData)
   mvarIDCONDPGTO = vData
End Property
Public Property Get IDCONDPGTO() As Double
   IDCONDPGTO = mvarIDCONDPGTO
End Property
Public Property Let IDFORMAPGTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFORMAPGTO <> vData)
   mvarIDFORMAPGTO = vData
End Property
Public Property Get IDFORMAPGTO() As Double
   IDFORMAPGTO = mvarIDFORMAPGTO
End Property
Public Property Let VLPGTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLPGTO <> vData)
   mvarVLPGTO = vData
End Property
Public Property Get VLPGTO() As Double
   VLPGTO = mvarVLPGTO
End Property
Public Property Let DTPGTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTPGTO <> vData)
   mvarDTPGTO = vData
End Property
Public Property Get DTPGTO() As String
   DTPGTO = mvarDTPGTO
End Property
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update CPGTOSVENDA Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDVENDA   = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " And IDPGTO    = " & SqlNum(mvarIDPGTO) & VbNewLine 
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
 
   Sql = "Insert Into CPGTOSVENDA (IDLOJA, IDVENDA, IDPGTO, IDCONDPGTO" & VbNewLine 
   Sql = Sql & ", IDFORMAPGTO, VLPGTO, DTPGTO, TIMESTAMP, ALTERSTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDPGTO),0)+1 From CPGTOSVENDA Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDVENDA = " & SqlNum(mvarIDVENDA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDPGTO) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDCONDPGTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFORMAPGTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLPGTO) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTPGTO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDPGTO) As String
   Dim Sql As String
 
   Sql = "Delete From CPGTOSVENDA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(Cstr(Ch_IDVENDA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDPGTO) Then Sql = Sql & " IDPGTO = " & SqlNum(Cstr(Ch_IDPGTO)) & " AND " & VbNewLine 
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
 
   Sql = "Update CPGTOSVENDA Set " & VbNewLine 
   Sql = Sql & " IDLOJA      = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDVENDA     = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " , IDPGTO      = " & SqlNum(mvarIDPGTO) & VbNewLine 
   Sql = Sql & " , IDCONDPGTO  = " & SqlNum(mvarIDCONDPGTO) & VbNewLine 
   Sql = Sql & " , IDFORMAPGTO = " & SqlNum(mvarIDFORMAPGTO) & VbNewLine 
   Sql = Sql & " , VLPGTO      = " & SqlNum(mvarVLPGTO) & VbNewLine 
   Sql = Sql & " , DTPGTO      = " & SqlDate(mvarDTPGTO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , TIMESTAMP   = " & "GetDate()" & VbNewLine 
   Sql = Sql & " , ALTERSTAMP  = " & "1" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CPGTOSVENDA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDVENDA   = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " And IDPGTO    = " & SqlNum(mvarIDPGTO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDPGTO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDVENDA, IDPGTO, IDCONDPGTO" & VbNewLine 
   Sql = Sql & ", IDFORMAPGTO, VLPGTO, DTPGTO, TIMESTAMP, ALTERSTAMP" & VbNewLine 
   Sql = Sql &" From CPGTOSVENDA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDVENDA)  And IsMissing(Ch_IDPGTO)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDVENDA = " & SqlNum(mvarIDVENDA) & " AND " & VbNewLine 
      Sql = Sql & " IDPGTO = " & SqlNum(mvarIDPGTO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(Cstr(Ch_IDVENDA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPGTO) Then Sql = Sql & " IDPGTO = " & SqlNum(Cstr(Ch_IDPGTO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDPGTO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDVENDA, Ch_IDPGTO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarIDPGTO = XVal(!IDPGTO & "")
      mvarIDCONDPGTO = XVal(!IDCONDPGTO & "")
      mvarIDFORMAPGTO = XVal(!IDFORMAPGTO & "")
      mvarVLPGTO = XVal(!VLPGTO & "")
      mvarDTPGTO = xDate(!DTPGTO & "", True)
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDVENDA, mvarIDPGTO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDVENDA As String, Ch_IDPGTO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CPGTOSVENDA Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDVENDA = " & SqlNum(Ch_IDVENDA)
   Sql = Sql & " , IDPGTO = " & SqlNum(Ch_IDPGTO)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDVENDA = " & SqlNum(mvarIDVENDA)
   Sql = Sql & " and IDPGTO = " & SqlNum(mvarIDPGTO)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
