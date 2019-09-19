VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CPROMO_PROD"
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
Private mvarIDPROMO      As Double
Private mvarIDPROD       As Double
Private mvarQTDPROD      As Double
Private mvarNMPROD       As String
Private mvarUNIDCONTROLE As String
Private mvarVLUNIT       As Double
Private mvarALTERSTAMP   As Double
Private mvarTIMESTAMP    As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDPROMO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROMO <> vData)
   mvarIDPROMO = vData
End Property
Public Property Get IDPROMO() As Double
   IDPROMO = mvarIDPROMO
End Property
Public Property Let IDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROD <> vData)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Double
   IDPROD = mvarIDPROD
End Property
Public Property Let QTDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDPROD <> vData)
   mvarQTDPROD = vData
End Property
Public Property Get QTDPROD() As Double
   QTDPROD = mvarQTDPROD
End Property
Public Property Let NMPROD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMPROD <> vData)
   mvarNMPROD = vData
End Property
Public Property Get NMPROD() As String
   NMPROD = mvarNMPROD
End Property
Public Property Let UNIDCONTROLE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUNIDCONTROLE <> vData)
   mvarUNIDCONTROLE = vData
End Property
Public Property Get UNIDCONTROLE() As String
   UNIDCONTROLE = mvarUNIDCONTROLE
End Property
Public Property Let VLUNIT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLUNIT <> vData)
   mvarVLUNIT = vData
End Property
Public Property Get VLUNIT() As Double
   VLUNIT = mvarVLUNIT
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update CPROMO_PROD Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDPROMO    = " & SqlNum(mvarIDPROMO) & vbNewLine
   Sql = Sql & " And IDPROD     = " & SqlNum(mvarIDPROD) & vbNewLine
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
Public Property Get QryInsert(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into CPROMO_PROD (IDLOJA, IDPROMO, IDPROD, QTDPROD" & vbNewLine
   Sql = Sql & ", NMPROD, UNIDCONTROLE, VLUNIT, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Values " & vbNewLine
   Sql = Sql & "(" & vbNewLine
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDPROMO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDPROD) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarQTDPROD) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarNMPROD) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarUNIDCONTROLE) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLUNIT) & vbNewLine
   Sql = Sql & ", " & "1" & vbNewLine
   Sql = Sql & ", " & "GetDate()" & vbNewLine
   Sql = Sql & ")" & vbNewLine

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDPROD, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From CPROMO_PROD" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDPROMO) And IsMissing(Ch_IDPROD) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDPROMO = " & SqlNum(mvarIDPROMO) & " AND " & vbNewLine
      Sql = Sql & " IDPROD = " & SqlNum(mvarIDPROD) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDPROMO) Then Sql = Sql & " IDPROMO = " & SqlNum(CStr(Ch_IDPROMO)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDPROD) Then Sql = Sql & " IDPROD = " & SqlNum(CStr(Ch_IDPROD)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = ""
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDPROMO = " & mvarIDPROMO
   sTag = sTag & "|IDPROD = " & mvarIDPROD
   sTag = sTag & "|QTDPROD = " & mvarQTDPROD
   sTag = sTag & "|NMPROD = " & mvarNMPROD
   sTag = sTag & "|UNIDCONTROLE = " & mvarUNIDCONTROLE
   sTag = sTag & "|VLUNIT = " & mvarVLUNIT
   sTag = sTag & "|Where = " & IIf(IsMissing(Ch_WHERE), "", Ch_WHERE)
   sTag = sTag & "|"

   MyDelRow.QUERY = Sql & ";"
   MyDelRow.TAG = sTag
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update CPROMO_PROD Set " & vbNewLine
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDPROMO      = " & SqlNum(mvarIDPROMO) & vbNewLine
   Sql = Sql & " , IDPROD       = " & SqlNum(mvarIDPROD) & vbNewLine
   Sql = Sql & " , QTDPROD      = " & SqlNum(mvarQTDPROD) & vbNewLine
   Sql = Sql & " , NMPROD       = " & SqlStr(mvarNMPROD) & vbNewLine
   Sql = Sql & " , UNIDCONTROLE = " & SqlStr(mvarUNIDCONTROLE) & vbNewLine
   Sql = Sql & " , VLUNIT       = " & SqlNum(mvarVLUNIT) & vbNewLine
   If pSinc Then
      Sql = Sql & ", ALTERSTAMP   = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP   = 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", TIMESTAMP    = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP    = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CPROMO_PROD'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDPROMO    = " & SqlNum(mvarIDPROMO) & vbNewLine
   Sql = Sql & " And IDPROD     = " & SqlNum(mvarIDPROD) & vbNewLine

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & vbNewLine
   Sql = Sql & Me.QryUpDate(pSinc:=pSinc)
   Sql = Sql & " Else " & vbNewLine
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId, pSinc:=pSinc)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDPROD, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDPROMO, IDPROD, QTDPROD" & vbNewLine
   Sql = Sql & ", NMPROD, UNIDCONTROLE, VLUNIT, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From CPROMO_PROD" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDPROMO) And IsMissing(Ch_IDPROD) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDPROMO = " & SqlNum(mvarIDPROMO) & " AND " & vbNewLine
      Sql = Sql & " IDPROD = " & SqlNum(mvarIDPROD) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDPROMO) Then Sql = Sql & " IDPROMO = " & SqlNum(CStr(Ch_IDPROMO)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDPROD) Then Sql = Sql & " IDPROD = " & SqlNum(CStr(Ch_IDPROD)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDPROD, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDPROMO, Ch_IDPROD, Ch_WHERE, Ch_ORDERBY)
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
    Call ShowError(Sql)
    Pesquisar = False
End Function
Public Sub Popula(Optional pRcSet)
   If IsMissing(pRcSet) Then Set pRcSet = mvarRS
   With pRcSet
      mvarIDLOJA = xVal(!IDLOJA & "")
      mvarIDPROMO = xVal(!IDPROMO & "")
      mvarIDPROD = xVal(!IDPROD & "")
      mvarQTDPROD = xVal(!QTDPROD & "")
      mvarNMPROD = !NMPROD & ""
      mvarUNIDCONTROLE = !UNIDCONTROLE & ""
      mvarVLUNIT = xVal(!VLUNIT & "")
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDPROMO = 0
   mvarIDPROD = 0
   mvarQTDPROD = 0
   mvarNMPROD = ""
   mvarUNIDCONTROLE = ""
   mvarVLUNIT = 0
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
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId As Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDPROMO, mvarIDPROD), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDPROMO As String, Ch_IDPROD As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CPROMO_PROD Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDPROMO = " & SqlNum(Ch_IDPROMO)
   Sql = Sql & " , IDPROD = " & SqlNum(Ch_IDPROD)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDPROMO = " & SqlNum(mvarIDPROMO)
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

