VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OTRATAMENTO_PROD"
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

Private mvarIDLOJA         As Double
Private mvarID             As Double
Private mvarIDTPSERVICO    As Double
Private mvarIDTPTRATAMENTO As Double
Private mvarIDAREA         As Double
Private mvarIDPROD         As Double
Private mvarALTERSTAMP     As Double
Private mvarTIMESTAMP      As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let ID(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarID <> vData)
   mvarID = vData
End Property
Public Property Get ID() As Double
   ID = mvarID
End Property
Public Property Let IDTPSERVICO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPSERVICO <> vData)
   mvarIDTPSERVICO = vData
End Property
Public Property Get IDTPSERVICO() As Double
   IDTPSERVICO = mvarIDTPSERVICO
End Property
Public Property Let IDTPTRATAMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPTRATAMENTO <> vData)
   mvarIDTPTRATAMENTO = vData
End Property
Public Property Get IDTPTRATAMENTO() As Double
   IDTPTRATAMENTO = mvarIDTPTRATAMENTO
End Property
Public Property Let IDAREA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAREA <> vData)
   mvarIDAREA = vData
End Property
Public Property Get IDAREA() As Double
   IDAREA = mvarIDAREA
End Property
Public Property Let IDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROD <> vData)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Double
   IDPROD = mvarIDPROD
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OTRATAMENTO_PROD Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And ID           = " & SqlNum(mvarID) & VbNewLine 
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
 
   Sql = "Insert Into OTRATAMENTO_PROD (IDLOJA, ID, IDTPSERVICO, IDTPTRATAMENTO" & VbNewLine 
   Sql = Sql & ", IDAREA, IDPROD, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(ID),0)+1 From OTRATAMENTO_PROD Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarID) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDTPSERVICO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDAREA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_ID) As String
   Dim Sql As String
 
   Sql = "Delete From OTRATAMENTO_PROD" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_ID) Then Sql = Sql & " ID = " & SqlNum(Cstr(Ch_ID)) & " AND " & VbNewLine 
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
 
   Sql = "Update OTRATAMENTO_PROD Set " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , ID             = " & SqlNum(mvarID) & VbNewLine 
   Sql = Sql & " , IDTPSERVICO    = " & SqlNum(mvarIDTPSERVICO) & VbNewLine 
   Sql = Sql & " , IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   Sql = Sql & " , IDAREA         = " & SqlNum(mvarIDAREA) & VbNewLine 
   Sql = Sql & " , IDPROD         = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP     = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP      = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OTRATAMENTO_PROD'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And ID           = " & SqlNum(mvarID) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_ID, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, ID, IDTPSERVICO, IDTPTRATAMENTO" & VbNewLine 
   Sql = Sql & ", IDAREA, IDPROD, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OTRATAMENTO_PROD" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_ID)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " ID = " & SqlNum(mvarID) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_ID) Then Sql = Sql & " ID = " & SqlNum(Cstr(Ch_ID)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_ID, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_ID, Ch_WHERE, Ch_ORDERBY)
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
      mvarID = XVal(!ID & "")
      mvarIDTPSERVICO = XVal(!IDTPSERVICO & "")
      mvarIDTPTRATAMENTO = XVal(!IDTPTRATAMENTO & "")
      mvarIDAREA = XVal(!IDAREA & "")
      mvarIDPROD = XVal(!IDPROD & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarID), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_ID As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OTRATAMENTO_PROD Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , ID = " & SqlNum(Ch_ID)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and ID = " & SqlNum(mvarID)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub