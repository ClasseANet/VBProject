VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OMAQMANUT"
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

Private mvarIDLOJA     As Double
Private mvarIDMANUT    As Double
Private mvarIDTPMANUT  As Double
Private mvarIDMAQUINA  As Double
Private mvarIDMANIPULO As Double
Private mvarDTMANUT    As String
Private mvarDSCMANUT   As String
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDMANUT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMANUT <> vData)
   mvarIDMANUT = vData
End Property
Public Property Get IDMANUT() As Double
   IDMANUT = mvarIDMANUT
End Property
Public Property Let IDTPMANUT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPMANUT <> vData)
   mvarIDTPMANUT = vData
End Property
Public Property Get IDTPMANUT() As Double
   IDTPMANUT = mvarIDTPMANUT
End Property
Public Property Let IDMAQUINA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMAQUINA <> vData)
   mvarIDMAQUINA = vData
End Property
Public Property Get IDMAQUINA() As Double
   IDMAQUINA = mvarIDMAQUINA
End Property
Public Property Let IDMANIPULO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMANIPULO <> vData)
   mvarIDMANIPULO = vData
End Property
Public Property Get IDMANIPULO() As Double
   IDMANIPULO = mvarIDMANIPULO
End Property
Public Property Let DTMANUT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTMANUT <> vData)
   mvarDTMANUT = vData
End Property
Public Property Get DTMANUT() As String
   DTMANUT = mvarDTMANUT
End Property
Public Property Let DSCMANUT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCMANUT <> vData)
   mvarDSCMANUT = vData
End Property
Public Property Get DSCMANUT() As String
   DSCMANUT = mvarDSCMANUT
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OMAQMANUT Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDMANUT  = " & SqlNum(mvarIDMANUT) & VbNewLine 
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
 
   Sql = "Insert Into OMAQMANUT (IDLOJA, IDMANUT, IDTPMANUT, IDMAQUINA" & VbNewLine 
   Sql = Sql & ", IDMANIPULO, DTMANUT, DSCMANUT, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDMANUT),0)+1 From OMAQMANUT Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDMANUT) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDTPMANUT) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDMAQUINA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDMANIPULO) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTMANUT, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCMANUT) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDMANUT) As String
   Dim Sql As String
 
   Sql = "Delete From OMAQMANUT" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDMANUT) Then Sql = Sql & " IDMANUT = " & SqlNum(Cstr(Ch_IDMANUT)) & " AND " & VbNewLine 
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
 
   Sql = "Update OMAQMANUT Set " & VbNewLine 
   Sql = Sql & " IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDMANUT    = " & SqlNum(mvarIDMANUT) & VbNewLine 
   Sql = Sql & " , IDTPMANUT  = " & SqlNum(mvarIDTPMANUT) & VbNewLine 
   Sql = Sql & " , IDMAQUINA  = " & SqlNum(mvarIDMAQUINA) & VbNewLine 
   Sql = Sql & " , IDMANIPULO = " & SqlNum(mvarIDMANIPULO) & VbNewLine 
   Sql = Sql & " , DTMANUT    = " & SqlDate(mvarDTMANUT, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DSCMANUT   = " & SqlStr(mvarDSCMANUT) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP  = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OMAQMANUT'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDMANUT  = " & SqlNum(mvarIDMANUT) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDMANUT, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDMANUT, IDTPMANUT, IDMAQUINA" & VbNewLine 
   Sql = Sql & ", IDMANIPULO, DTMANUT, DSCMANUT, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OMAQMANUT" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDMANUT)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDMANUT = " & SqlNum(mvarIDMANUT) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDMANUT) Then Sql = Sql & " IDMANUT = " & SqlNum(Cstr(Ch_IDMANUT)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDMANUT, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDMANUT, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDMANUT = XVal(!IDMANUT & "")
      mvarIDTPMANUT = XVal(!IDTPMANUT & "")
      mvarIDMAQUINA = XVal(!IDMAQUINA & "")
      mvarIDMANIPULO = XVal(!IDMANIPULO & "")
      mvarDTMANUT = xDate(!DTMANUT & "", True)
      mvarDSCMANUT = !DSCMANUT & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDMANUT), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDMANUT As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OMAQMANUT Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDMANUT = " & SqlNum(Ch_IDMANUT)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDMANUT = " & SqlNum(mvarIDMANUT)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
