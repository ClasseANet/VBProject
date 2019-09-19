VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_FLAN_PGTO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb         As DS_BANCO
Private mvarRS          As Object 

Private mvarQryInsert   As String
Private mvarQryUpDate   As String
Private mvarQryDelete   As String
Private mvarQrySelect   As String
Private mvarQrySave     As String

Private mvarisDirt      As Boolean

Private mvarIDLOJA     As Double
Private mvarIDCONTA    As Double
Private mvarIDLAN      As Double
Private mvarIDPGTO     As Double
Private mvarIDVENDA    As Double
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCONTA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCONTA <> vData)
   mvarIDCONTA = vData
End Property
Public Property Get IDCONTA() As Double
   IDCONTA = mvarIDCONTA
End Property
Public Property Let IDLAN(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLAN <> vData)
   mvarIDLAN = vData
End Property
Public Property Get IDLAN() As Double
   IDLAN = mvarIDLAN
End Property
Public Property Let IDPGTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPGTO <> vData)
   mvarIDPGTO = vData
End Property
Public Property Get IDPGTO() As Double
   IDPGTO = mvarIDPGTO
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
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
Public Property Get QryInsert() As String
   Dim Sql As String
 
   Sql = "Insert Into FLAN_PGTO (IDLOJA, IDCONTA, IDLAN, IDPGTO" & VbNewLine 
   Sql = Sql & ", IDVENDA, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCONTA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDLAN) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPGTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(1) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarXDb.Sysdate(eSysDate.Data_Hora)) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDCONTA, Optional Ch_IDLAN, Optional Ch_IDPGTO, Optional Ch_IDVENDA) As String
   Dim Sql As String
 
   Sql = "Delete From FLAN_PGTO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDCONTA) Then Sql = Sql & " IDCONTA = " & SqlNum(Cstr(Ch_IDCONTA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDLAN) Then Sql = Sql & " IDLAN = " & SqlNum(Cstr(Ch_IDLAN)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDPGTO) Then Sql = Sql & " IDPGTO = " & SqlNum(Cstr(Ch_IDPGTO)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(Cstr(Ch_IDVENDA)) & " AND " & VbNewLine 
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
 
   Sql = "Update FLAN_PGTO Set " & VbNewLine 
   Sql = Sql & " IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDCONTA    = " & SqlNum(mvarIDCONTA) & VbNewLine 
   Sql = Sql & " , IDLAN      = " & SqlNum(mvarIDLAN) & VbNewLine 
   Sql = Sql & " , IDPGTO     = " & SqlNum(mvarIDPGTO) & VbNewLine 
   Sql = Sql & " , IDVENDA    = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP = " & SqlNum(1) & VbNewLine 
   Sql = Sql & " , TIMESTAMP  = " & SqlStr(mvarXDb.Sysdate(eSysDate.Data_Hora)) & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDCONTA  = " & SqlNum(mvarIDCONTA) & VbNewLine 
   Sql = Sql & " And IDLAN    = " & SqlNum(mvarIDLAN) & VbNewLine 
   Sql = Sql & " And IDPGTO   = " & SqlNum(mvarIDPGTO) & VbNewLine 
   Sql = Sql & " And IDVENDA  = " & SqlNum(mvarIDVENDA) & VbNewLine 
   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave() As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDCONTA, Optional Ch_IDLAN, Optional Ch_IDPGTO, Optional Ch_IDVENDA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDCONTA, IDLAN, IDPGTO" & VbNewLine 
   Sql = Sql & ", IDVENDA, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From FLAN_PGTO" & VbNewLine 
   If IsMissing(Ch_WHERE) Then
      Sql = Sql & " Where " & VbNewLine 
      If IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDCONTA)  And IsMissing(Ch_IDLAN)  And IsMissing(Ch_IDPGTO)  And IsMissing(Ch_IDVENDA)  Then 
         Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & VbNewLine 
         Sql = Sql & " And IDCONTA = " & SqlNum(mvarIDCONTA) & VbNewLine 
         Sql = Sql & " And IDLAN = " & SqlNum(mvarIDLAN) & VbNewLine 
         Sql = Sql & " And IDPGTO = " & SqlNum(mvarIDPGTO) & VbNewLine 
         Sql = Sql & " And IDVENDA = " & SqlNum(mvarIDVENDA) & VbNewLine 
      Else 
         If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
         If Not isMissing(Ch_IDCONTA) Then Sql = Sql & " IDCONTA = " & SqlNum(Cstr(Ch_IDCONTA)) & " AND " & VbNewLine 
         If Not isMissing(Ch_IDLAN) Then Sql = Sql & " IDLAN = " & SqlNum(Cstr(Ch_IDLAN)) & " AND " & VbNewLine 
         If Not isMissing(Ch_IDPGTO) Then Sql = Sql & " IDPGTO = " & SqlNum(Cstr(Ch_IDPGTO)) & " AND " & VbNewLine 
         If Not isMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(Cstr(Ch_IDVENDA)) & " AND " & VbNewLine 
         Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
      End If 
   Else
      If Trim(Ch_WHERE) <> "" Then
         Sql = Sql & " Where " & Ch_WHERE & vbNewLine
      End If
   End If
   If Not IsMissing(Ch_ORDERBY) Then
      If Trim(Ch_ORDERBY) <> "" Then
         Sql = Sql & " Order By " & Ch_ORDERBY & vbNewLine
      End If
   End If

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDCONTA, Optional Ch_IDLAN, Optional Ch_IDPGTO, Optional Ch_IDVENDA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDCONTA, Ch_IDLAN, Ch_IDPGTO, Ch_IDVENDA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCONTA = XVal(!IDCONTA & "")
      mvarIDLAN = XVal(!IDLAN & "")
      mvarIDPGTO = XVal(!IDPGTO & "")
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = FormatDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDCONTA, mvarIDLAN, mvarIDPGTO, mvarIDVENDA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDCONTA As String, Ch_IDLAN As String, Ch_IDPGTO As String, Ch_IDVENDA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FLAN_PGTO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDCONTA = " & SqlNum(Ch_IDCONTA)
   Sql = Sql & " , IDLAN = " & SqlNum(Ch_IDLAN)
   Sql = Sql & " , IDPGTO = " & SqlNum(Ch_IDPGTO)
   Sql = Sql & " , IDVENDA = " & SqlNum(Ch_IDVENDA)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDCONTA = " & SqlNum(mvarIDCONTA)
   Sql = Sql & " and IDLAN = " & SqlNum(mvarIDLAN)
   Sql = Sql & " and IDPGTO = " & SqlNum(mvarIDPGTO)
   Sql = Sql & " and IDVENDA = " & SqlNum(mvarIDVENDA)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
