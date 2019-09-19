VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OMAQUINA"
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

Private mvarIDLOJA       As Double
Private mvarIDMAQUINA    As Double
Private mvarCODMAQUINA   As String
Private mvarIDTPMAQ      As Double
Private mvarNREGISTRO    As String
Private mvarNANVISA      As String
Private mvarDTOPERACAO   As String
Private mvarDTFECHAMENTO As String
Private mvarDISPAROS     As Double
Private mvarSITMAQUINA   As Double
Private mvarALTERSTAMP   As Double
Private mvarTIMESTAMP    As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDMAQUINA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMAQUINA <> vData)
   mvarIDMAQUINA = vData
End Property
Public Property Get IDMAQUINA() As Double
   IDMAQUINA = mvarIDMAQUINA
End Property
Public Property Let CODMAQUINA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODMAQUINA <> vData)
   mvarCODMAQUINA = vData
End Property
Public Property Get CODMAQUINA() As String
   CODMAQUINA = mvarCODMAQUINA
End Property
Public Property Let IDTPMAQ(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPMAQ <> vData)
   mvarIDTPMAQ = vData
End Property
Public Property Get IDTPMAQ() As Double
   IDTPMAQ = mvarIDTPMAQ
End Property
Public Property Let NREGISTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNREGISTRO <> vData)
   mvarNREGISTRO = vData
End Property
Public Property Get NREGISTRO() As String
   NREGISTRO = mvarNREGISTRO
End Property
Public Property Let NANVISA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNANVISA <> vData)
   mvarNANVISA = vData
End Property
Public Property Get NANVISA() As String
   NANVISA = mvarNANVISA
End Property
Public Property Let DTOPERACAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTOPERACAO <> vData)
   mvarDTOPERACAO = vData
End Property
Public Property Get DTOPERACAO() As String
   DTOPERACAO = mvarDTOPERACAO
End Property
Public Property Let DTFECHAMENTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTFECHAMENTO <> vData)
   mvarDTFECHAMENTO = vData
End Property
Public Property Get DTFECHAMENTO() As String
   DTFECHAMENTO = mvarDTFECHAMENTO
End Property
Public Property Let DISPAROS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarDISPAROS <> vData)
   mvarDISPAROS = vData
End Property
Public Property Get DISPAROS() As Double
   DISPAROS = mvarDISPAROS
End Property
Public Property Let SITMAQUINA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSITMAQUINA <> vData)
   mvarSITMAQUINA = vData
End Property
Public Property Get SITMAQUINA() As Double
   SITMAQUINA = mvarSITMAQUINA
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OMAQUINA Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDMAQUINA  = " & SqlNum(mvarIDMAQUINA) & VbNewLine 
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
 
   Sql = "Insert Into OMAQUINA (IDLOJA, IDMAQUINA, CODMAQUINA, IDTPMAQ" & VbNewLine 
   Sql = Sql & ", NREGISTRO, NANVISA, DTOPERACAO, DTFECHAMENTO, DISPAROS" & VbNewLine 
   Sql = Sql & ", SITMAQUINA, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDMAQUINA),0)+1 From OMAQUINA Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDMAQUINA) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarCODMAQUINA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTPMAQ) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNREGISTRO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNANVISA) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTOPERACAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTFECHAMENTO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarDISPAROS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSITMAQUINA) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDMAQUINA) As String
   Dim Sql As String
 
   Sql = "Delete From OMAQUINA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDMAQUINA) Then Sql = Sql & " IDMAQUINA = " & SqlNum(Cstr(Ch_IDMAQUINA)) & " AND " & VbNewLine 
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
 
   Sql = "Update OMAQUINA Set " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDMAQUINA    = " & SqlNum(mvarIDMAQUINA) & VbNewLine 
   Sql = Sql & " , CODMAQUINA   = " & SqlStr(mvarCODMAQUINA) & VbNewLine 
   Sql = Sql & " , IDTPMAQ      = " & SqlNum(mvarIDTPMAQ) & VbNewLine 
   Sql = Sql & " , NREGISTRO    = " & SqlStr(mvarNREGISTRO) & VbNewLine 
   Sql = Sql & " , NANVISA      = " & SqlStr(mvarNANVISA) & VbNewLine 
   Sql = Sql & " , DTOPERACAO   = " & SqlDate(mvarDTOPERACAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTFECHAMENTO = " & SqlDate(mvarDTFECHAMENTO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DISPAROS     = " & SqlNum(mvarDISPAROS) & VbNewLine 
   Sql = Sql & " , SITMAQUINA   = " & SqlNum(mvarSITMAQUINA) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP   = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP    = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OMAQUINA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDMAQUINA  = " & SqlNum(mvarIDMAQUINA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDMAQUINA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDMAQUINA, CODMAQUINA, IDTPMAQ" & VbNewLine 
   Sql = Sql & ", NREGISTRO, NANVISA, DTOPERACAO, DTFECHAMENTO, DISPAROS" & VbNewLine 
   Sql = Sql & ", SITMAQUINA, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OMAQUINA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDMAQUINA)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDMAQUINA = " & SqlNum(mvarIDMAQUINA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDMAQUINA) Then Sql = Sql & " IDMAQUINA = " & SqlNum(Cstr(Ch_IDMAQUINA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDMAQUINA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDMAQUINA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDMAQUINA = XVal(!IDMAQUINA & "")
      mvarCODMAQUINA = !CODMAQUINA & ""
      mvarIDTPMAQ = XVal(!IDTPMAQ & "")
      mvarNREGISTRO = !NREGISTRO & ""
      mvarNANVISA = !NANVISA & ""
      mvarDTOPERACAO = xDate(!DTOPERACAO & "", True)
      mvarDTFECHAMENTO = xDate(!DTFECHAMENTO & "", True)
      mvarDISPAROS = XVal(!DISPAROS & "")
      mvarSITMAQUINA = XVal(!SITMAQUINA & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDMAQUINA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDMAQUINA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OMAQUINA Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDMAQUINA = " & SqlNum(Ch_IDMAQUINA)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDMAQUINA = " & SqlNum(mvarIDMAQUINA)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
