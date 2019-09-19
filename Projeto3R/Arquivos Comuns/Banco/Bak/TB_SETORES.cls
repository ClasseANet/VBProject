VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_SETORES"
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

Private mvarIDSETOR    As Double
Private mvarCODSETOR   As String
Private mvarDSCSETOR   As String
Private mvarIDPAI      As Double
Private mvarNIVELOS    As Double
Private mvarPRIORIDADE As Double
Private mvarIDEMPR     As Double
Private mvarIDCOLIGADA As Double
Private mvarIDAUXILIAR As String
Private mvarIDUSURESP  As String
Private mvarIDEMPRESA  As Double
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
Public Property Let PRIORIDADE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPRIORIDADE <> vData)
   mvarPRIORIDADE = vData
End Property
Public Property Get PRIORIDADE() As Double
   PRIORIDADE = mvarPRIORIDADE
End Property
Public Property Let IDEMPR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEMPR <> vData)
   mvarIDEMPR = vData
End Property
Public Property Get IDEMPR() As Double
   IDEMPR = mvarIDEMPR
End Property
Public Property Let IDCOLIGADA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCOLIGADA <> vData)
   mvarIDCOLIGADA = vData
End Property
Public Property Get IDCOLIGADA() As Double
   IDCOLIGADA = mvarIDCOLIGADA
End Property
Public Property Let IDAUXILIAR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAUXILIAR <> vData)
   mvarIDAUXILIAR = vData
End Property
Public Property Get IDAUXILIAR() As String
   IDAUXILIAR = mvarIDAUXILIAR
End Property
Public Property Let IDUSURESP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSURESP <> vData)
   mvarIDUSURESP = vData
End Property
Public Property Get IDUSURESP() As String
   IDUSURESP = mvarIDUSURESP
End Property
Public Property Let IDEMPRESA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEMPRESA <> vData)
   mvarIDEMPRESA = vData
End Property
Public Property Get IDEMPRESA() As Double
   IDEMPRESA = mvarIDEMPRESA
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
 
   Sql = "Insert Into SETORES (CODSETOR, DSCSETOR, IDPAI" & VbNewLine 
   Sql = Sql & ", NIVELOS, PRIORIDADE, IDEMPR, IDCOLIGADA, IDAUXILIAR" & VbNewLine 
   Sql = Sql & ", IDUSURESP, IDEMPRESA) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarCODSETOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCSETOR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarNIVELOS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPRIORIDADE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDEMPR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDAUXILIAR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSURESP) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDEMPRESA) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDSETOR) As String
   Dim Sql As String
 
   Sql = "Delete From SETORES" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDSETOR) Then Sql = Sql & " IDSETOR = " & SqlNum(Cstr(Ch_IDSETOR)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update SETORES Set " & VbNewLine 
   Sql = Sql & " CODSETOR   = " & SqlStr(mvarCODSETOR) & VbNewLine 
   Sql = Sql & " , DSCSETOR   = " & SqlStr(mvarDSCSETOR) & VbNewLine 
   Sql = Sql & " , IDPAI      = " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & " , NIVELOS    = " & SqlNum(mvarNIVELOS) & VbNewLine 
   Sql = Sql & " , PRIORIDADE = " & SqlNum(mvarPRIORIDADE) & VbNewLine 
   Sql = Sql & " , IDEMPR     = " & SqlNum(mvarIDEMPR) & VbNewLine 
   Sql = Sql & " , IDCOLIGADA = " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & " , IDAUXILIAR = " & SqlStr(mvarIDAUXILIAR) & VbNewLine 
   Sql = Sql & " , IDUSURESP  = " & SqlStr(mvarIDUSURESP) & VbNewLine 
   Sql = Sql & " , IDEMPRESA  = " & SqlNum(mvarIDEMPRESA) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'SETORES'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDSETOR      = " & SqlNum(mvarIDSETOR) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDSETOR, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDSETOR, CODSETOR, DSCSETOR, IDPAI" & VbNewLine 
   Sql = Sql & ", NIVELOS, PRIORIDADE, IDEMPR, IDCOLIGADA, IDAUXILIAR" & VbNewLine 
   Sql = Sql & ", IDUSURESP, IDEMPRESA" & VbNewLine 
   Sql = Sql &" From SETORES" & VbNewLine 
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
      mvarPRIORIDADE = XVal(!PRIORIDADE & "")
      mvarIDEMPR = XVal(!IDEMPR & "")
      mvarIDCOLIGADA = XVal(!IDCOLIGADA & "")
      mvarIDAUXILIAR = !IDAUXILIAR & ""
      mvarIDUSURESP = !IDUSURESP & ""
      mvarIDEMPRESA = XVal(!IDEMPRESA & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDSETOR), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
