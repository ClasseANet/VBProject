VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_SETORESCLI"
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

Private mvarIDSETOR   As Double
Private mvarCODSETOR  As String
Private mvarDSCSETOR  As String
Private mvarIDPAI     As Double
Private mvarNIVELOS   As Double
Private mvarIDCLIENTE As Double
Public Property Get IDSETOR() As Double
   IDSETOR = mvarIDSETOR
End Property
Public Property Let CODSETOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSETOR <> vData)
   mvarCODSETOR = vData
End Property
Public Property Get CODSETOR() As String
   CODSETOR = mvarCODSETOR
End Property
Public Property Let DSCSETOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCSETOR <> vData)
   mvarDSCSETOR = vData
End Property
Public Property Get DSCSETOR() As String
   DSCSETOR = mvarDSCSETOR
End Property
Public Property Let IDPAI(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAI <> vData)
   mvarIDPAI = vData
End Property
Public Property Get IDPAI() As Double
   IDPAI = mvarIDPAI
End Property
Public Property Let NIVELOS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNIVELOS <> vData)
   mvarNIVELOS = vData
End Property
Public Property Get NIVELOS() As Double
   NIVELOS = mvarNIVELOS
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
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
 
   Sql = "Insert Into SETORESCLI (CODSETOR, DSCSETOR, IDPAI" & VbNewLine 
   Sql = Sql & ", NIVELOS, IDCLIENTE) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarCODSETOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCSETOR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarNIVELOS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDSETOR, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From SETORESCLI" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDSETOR)  Then 
      Sql = Sql & " IDSETOR = " & SqlNum(mvarIDSETOR) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDSETOR) Then Sql = Sql & " IDSETOR = " & SqlNum(Cstr(Ch_IDSETOR)) & " AND " & VbNewLine 
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
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update SETORESCLI Set " & VbNewLine 
   Sql = Sql & " CODSETOR  = " & SqlStr(mvarCODSETOR) & VbNewLine 
   Sql = Sql & " , DSCSETOR  = " & SqlStr(mvarDSCSETOR) & VbNewLine 
   Sql = Sql & " , IDPAI     = " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & " , NIVELOS   = " & SqlNum(mvarNIVELOS) & VbNewLine 
   Sql = Sql & " , IDCLIENTE = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'SETORESCLI'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDSETOR     = " & SqlNum(mvarIDSETOR) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDSETOR, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDSETOR, CODSETOR, DSCSETOR, IDPAI" & VbNewLine 
   Sql = Sql & ", NIVELOS, IDCLIENTE" & VbNewLine 
   Sql = Sql &" From SETORESCLI" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDSETOR)  Then 
      Sql = Sql & " IDSETOR = " & SqlNum(mvarIDSETOR) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDSETOR) Then Sql = Sql & " IDSETOR = " & SqlNum(Cstr(Ch_IDSETOR)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDSETOR, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDSETOR, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDSETOR = XVal(!IDSETOR & "")
      mvarCODSETOR = !CODSETOR & ""
      mvarDSCSETOR = !DSCSETOR & ""
      mvarIDPAI = XVal(!IDPAI & "")
      mvarNIVELOS = XVal(!NIVELOS & "")
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDSETOR = 0
   mvarCODSETOR = ""
   mvarDSCSETOR = ""
   mvarIDPAI = 0
   mvarNIVELOS = 0
   mvarIDCLIENTE = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDSETOR), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
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
