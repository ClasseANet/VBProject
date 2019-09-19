VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_SESTPROD"
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

Private mvarIDLOJA        As Double
Private mvarIDEST         As Double
Private mvarIDPROD        As Double
Private mvarUNIDCONTROLE  As String
Private mvarSLDATUAL      As Double
Private mvarSLDDISPONIVEL As Double
Private mvarSLDFECHAMENTO As Double
Private mvarALTERSTAMP    As Double
Private mvarTIMESTAMP     As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDEST(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEST <> vData)
   mvarIDEST = vData
End Property
Public Property Get IDEST() As Double
   IDEST = mvarIDEST
End Property
Public Property Let IDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROD <> vData)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Double
   IDPROD = mvarIDPROD
End Property
Public Property Let UNIDCONTROLE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUNIDCONTROLE <> vData)
   mvarUNIDCONTROLE = vData
End Property
Public Property Get UNIDCONTROLE() As String
   UNIDCONTROLE = mvarUNIDCONTROLE
End Property
Public Property Let SLDATUAL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSLDATUAL <> vData)
   mvarSLDATUAL = vData
End Property
Public Property Get SLDATUAL() As Double
   SLDATUAL = mvarSLDATUAL
End Property
Public Property Let SLDDISPONIVEL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSLDDISPONIVEL <> vData)
   mvarSLDDISPONIVEL = vData
End Property
Public Property Get SLDDISPONIVEL() As Double
   SLDDISPONIVEL = mvarSLDDISPONIVEL
End Property
Public Property Let SLDFECHAMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSLDFECHAMENTO <> vData)
   mvarSLDFECHAMENTO = vData
End Property
Public Property Get SLDFECHAMENTO() As Double
   SLDFECHAMENTO = mvarSLDFECHAMENTO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update SESTPROD Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDEST       = " & SqlNum(mvarIDEST) & VbNewLine 
   Sql = Sql & " And IDPROD      = " & SqlNum(mvarIDPROD) & VbNewLine 
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
 
   Sql = "Insert Into SESTPROD (IDLOJA, IDEST, IDPROD, UNIDCONTROLE" & VbNewLine 
   Sql = Sql & ", SLDATUAL, SLDDISPONIVEL, SLDFECHAMENTO, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDEST) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUNIDCONTROLE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSLDATUAL) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSLDDISPONIVEL) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSLDFECHAMENTO) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDEST, Optional Ch_IDPROD, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From SESTPROD" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDEST)  And IsMissing(Ch_IDPROD)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDEST = " & SqlNum(mvarIDEST) & " AND " & VbNewLine 
      Sql = Sql & " IDPROD = " & SqlNum(mvarIDPROD) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDEST) Then Sql = Sql & " IDEST = " & SqlNum(Cstr(Ch_IDEST)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPROD) Then Sql = Sql & " IDPROD = " & SqlNum(Cstr(Ch_IDPROD)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDEST = " & mvarIDEST
   sTag = sTag & "|IDPROD = " & mvarIDPROD
   sTag = sTag & "|UNIDCONTROLE = " & mvarUNIDCONTROLE
   sTag = sTag & "|SLDATUAL = " & mvarSLDATUAL
   sTag = sTag & "|SLDDISPONIVEL = " & mvarSLDDISPONIVEL
   sTag = sTag & "|SLDFECHAMENTO = " & mvarSLDFECHAMENTO
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
 
   Sql = "Update SESTPROD Set " & VbNewLine 
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDEST         = " & SqlNum(mvarIDEST) & VbNewLine 
   Sql = Sql & " , IDPROD        = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " , UNIDCONTROLE  = " & SqlStr(mvarUNIDCONTROLE) & VbNewLine 
   Sql = Sql & " , SLDATUAL      = " & SqlNum(mvarSLDATUAL) & VbNewLine 
   Sql = Sql & " , SLDDISPONIVEL = " & SqlNum(mvarSLDDISPONIVEL) & VbNewLine 
   Sql = Sql & " , SLDFECHAMENTO = " & SqlNum(mvarSLDFECHAMENTO) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP    = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP    = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP     = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP     = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'SESTPROD'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDEST       = " & SqlNum(mvarIDEST) & VbNewLine 
   Sql = Sql & " And IDPROD      = " & SqlNum(mvarIDPROD) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDEST, Optional Ch_IDPROD, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDEST, IDPROD, UNIDCONTROLE" & VbNewLine 
   Sql = Sql & ", SLDATUAL, SLDDISPONIVEL, SLDFECHAMENTO, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From SESTPROD" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDEST)  And IsMissing(Ch_IDPROD)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDEST = " & SqlNum(mvarIDEST) & " AND " & VbNewLine 
      Sql = Sql & " IDPROD = " & SqlNum(mvarIDPROD) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDEST) Then Sql = Sql & " IDEST = " & SqlNum(Cstr(Ch_IDEST)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPROD) Then Sql = Sql & " IDPROD = " & SqlNum(Cstr(Ch_IDPROD)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDEST, Optional Ch_IDPROD, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDEST, Ch_IDPROD, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDEST = XVal(!IDEST & "")
      mvarIDPROD = XVal(!IDPROD & "")
      mvarUNIDCONTROLE = !UNIDCONTROLE & ""
      mvarSLDATUAL = XVal(!SLDATUAL & "")
      mvarSLDDISPONIVEL = XVal(!SLDDISPONIVEL & "")
      mvarSLDFECHAMENTO = XVal(!SLDFECHAMENTO & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDEST = 0
   mvarIDPROD = 0
   mvarUNIDCONTROLE = ""
   mvarSLDATUAL = 0
   mvarSLDDISPONIVEL = 0
   mvarSLDFECHAMENTO = 0
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDEST, mvarIDPROD), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDEST As String, Ch_IDPROD As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update SESTPROD Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDEST = " & SqlNum(Ch_IDEST)
   Sql = Sql & " , IDPROD = " & SqlNum(Ch_IDPROD)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDEST = " & SqlNum(mvarIDEST)
   Sql = Sql & " and IDPROD = " & SqlNum(mvarIDPROD)
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
