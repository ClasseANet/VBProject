VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_FMOEDA"
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

Private mvarIDMOEDA    As Double
Private mvarNMMOEDA    As String
Private mvarIDPAIS     As Double
Private mvarSIGLAMOEDA As String
Private mvarIDAUXILIAR As String
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let IDMOEDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMOEDA <> vData)
   mvarIDMOEDA = vData
End Property
Public Property Get IDMOEDA() As Double
   IDMOEDA = mvarIDMOEDA
End Property
Public Property Let NMMOEDA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMMOEDA <> vData)
   mvarNMMOEDA = vData
End Property
Public Property Get NMMOEDA() As String
   NMMOEDA = mvarNMMOEDA
End Property
Public Property Let IDPAIS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAIS <> vData)
   mvarIDPAIS = vData
End Property
Public Property Get IDPAIS() As Double
   IDPAIS = mvarIDPAIS
End Property
Public Property Let SIGLAMOEDA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSIGLAMOEDA <> vData)
   mvarSIGLAMOEDA = vData
End Property
Public Property Get SIGLAMOEDA() As String
   SIGLAMOEDA = mvarSIGLAMOEDA
End Property
Public Property Let IDAUXILIAR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAUXILIAR <> vData)
   mvarIDAUXILIAR = vData
End Property
Public Property Get IDAUXILIAR() As String
   IDAUXILIAR = mvarIDAUXILIAR
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update FMOEDA Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDMOEDA      = " & SqlNum(mvarIDMOEDA) & VbNewLine 
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
 
   Sql = "Insert Into FMOEDA (IDMOEDA, NMMOEDA, IDPAIS, SIGLAMOEDA" & VbNewLine 
   Sql = Sql & ", IDAUXILIAR, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   If pAutoId Then
      Sql = Sql &  "(Select isNull(Max(IDMOEDA),0)+1 From FMOEDA)"
   Else
      Sql = Sql &  SqlNum(mvarIDMOEDA) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarNMMOEDA) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDPAIS = 0, "Null", SqlNum(mvarIDPAIS)) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSIGLAMOEDA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDAUXILIAR) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", " &  SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", " &  SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDMOEDA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From FMOEDA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDMOEDA)  Then 
      Sql = Sql & " IDMOEDA = " & SqlNum(mvarIDMOEDA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDMOEDA) Then Sql = Sql & " IDMOEDA = " & SqlNum(Cstr(Ch_IDMOEDA)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDMOEDA = " & mvarIDMOEDA
   sTag = sTag & "|NMMOEDA = " & mvarNMMOEDA
   sTag = sTag & "|IDPAIS = " & mvarIDPAIS
   sTag = sTag & "|SIGLAMOEDA = " & mvarSIGLAMOEDA
   sTag = sTag & "|IDAUXILIAR = " & mvarIDAUXILIAR
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
 
   Sql = "Update FMOEDA Set " & VbNewLine 
   Sql = Sql & " IDMOEDA    = " & SqlNum(mvarIDMOEDA) & VbNewLine 
   Sql = Sql & " , NMMOEDA    = " & SqlStr(mvarNMMOEDA) & VbNewLine 
   Sql = Sql & " , IDPAIS     = " & IIf(mvarIDPAIS = 0, "Null", SqlNum(mvarIDPAIS)) & VbNewLine 
   Sql = Sql & " , SIGLAMOEDA = " & SqlStr(mvarSIGLAMOEDA) & VbNewLine 
   Sql = Sql & " , IDAUXILIAR = " & SqlStr(mvarIDAUXILIAR) & VbNewLine 
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
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FMOEDA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDMOEDA      = " & SqlNum(mvarIDMOEDA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDMOEDA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDMOEDA, NMMOEDA, IDPAIS, SIGLAMOEDA" & VbNewLine 
   Sql = Sql & ", IDAUXILIAR, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From FMOEDA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDMOEDA)  Then 
      Sql = Sql & " IDMOEDA = " & SqlNum(mvarIDMOEDA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDMOEDA) Then Sql = Sql & " IDMOEDA = " & SqlNum(Cstr(Ch_IDMOEDA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDMOEDA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDMOEDA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDMOEDA = XVal(!IDMOEDA & "")
      mvarNMMOEDA = !NMMOEDA & ""
      mvarIDPAIS = XVal(!IDPAIS & "")
      mvarSIGLAMOEDA = !SIGLAMOEDA & ""
      mvarIDAUXILIAR = !IDAUXILIAR & ""
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDMOEDA = 0
   mvarNMMOEDA = ""
   mvarIDPAIS = 0
   mvarSIGLAMOEDA = ""
   mvarIDAUXILIAR = ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDMOEDA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDMOEDA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FMOEDA Set "
   Sql = Sql & " IDMOEDA = " & SqlNum(Ch_IDMOEDA)
   Sql = Sql & " Where 
   Sql = Sql & " IDMOEDA = " & SqlNum(mvarIDMOEDA)
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
