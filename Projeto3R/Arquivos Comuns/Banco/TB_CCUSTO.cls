VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_CCUSTO"
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

Private mvarIDCCUSTO   As Double
Private mvarDSCCCUSTO  As String
Private mvarTPAPROV    As String
Private mvarCODCCUSTO  As String
Private mvarIDPAI      As Double
Private mvarIDSETOR    As Double
Private mvarGRUPO      As String
Private mvarSITCCUSTO  As String
Private mvarIDPROJ     As Double
Private mvarIDEMPR     As Double
Private mvarIDAUXILIAR As String
Private mvarIDCOLIGADA As Double
Private mvarIDUSURESP  As String
Public Property Let IDCCUSTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCCUSTO <> vData)
   mvarIDCCUSTO = vData
End Property
Public Property Get IDCCUSTO() As Double
   IDCCUSTO = mvarIDCCUSTO
End Property
Public Property Let DSCCCUSTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCCCUSTO <> vData)
   mvarDSCCCUSTO = vData
End Property
Public Property Get DSCCCUSTO() As String
   DSCCCUSTO = mvarDSCCCUSTO
End Property
Public Property Let TPAPROV(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPAPROV <> vData)
   mvarTPAPROV = vData
End Property
Public Property Get TPAPROV() As String
   TPAPROV = mvarTPAPROV
End Property
Public Property Let CODCCUSTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODCCUSTO <> vData)
   mvarCODCCUSTO = vData
End Property
Public Property Get CODCCUSTO() As String
   CODCCUSTO = mvarCODCCUSTO
End Property
Public Property Let IDPAI(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAI <> vData)
   mvarIDPAI = vData
End Property
Public Property Get IDPAI() As Double
   IDPAI = mvarIDPAI
End Property
Public Property Let IDSETOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSETOR <> vData)
   mvarIDSETOR = vData
End Property
Public Property Get IDSETOR() As Double
   IDSETOR = mvarIDSETOR
End Property
Public Property Let GRUPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarGRUPO <> vData)
   mvarGRUPO = vData
End Property
Public Property Get GRUPO() As String
   GRUPO = mvarGRUPO
End Property
Public Property Let SITCCUSTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITCCUSTO <> vData)
   mvarSITCCUSTO = vData
End Property
Public Property Get SITCCUSTO() As String
   SITCCUSTO = mvarSITCCUSTO
End Property
Public Property Let IDPROJ(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROJ <> vData)
   mvarIDPROJ = vData
End Property
Public Property Get IDPROJ() As Double
   IDPROJ = mvarIDPROJ
End Property
Public Property Let IDEMPR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEMPR <> vData)
   mvarIDEMPR = vData
End Property
Public Property Get IDEMPR() As Double
   IDEMPR = mvarIDEMPR
End Property
Public Property Let IDAUXILIAR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAUXILIAR <> vData)
   mvarIDAUXILIAR = vData
End Property
Public Property Get IDAUXILIAR() As String
   IDAUXILIAR = mvarIDAUXILIAR
End Property
Public Property Let IDCOLIGADA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCOLIGADA <> vData)
   mvarIDCOLIGADA = vData
End Property
Public Property Get IDCOLIGADA() As Double
   IDCOLIGADA = mvarIDCOLIGADA
End Property
Public Property Let IDUSURESP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSURESP <> vData)
   mvarIDUSURESP = vData
End Property
Public Property Get IDUSURESP() As String
   IDUSURESP = mvarIDUSURESP
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
 
   Sql = "Insert Into CCUSTO (IDCCUSTO, DSCCCUSTO, TPAPROV, CODCCUSTO" & VbNewLine 
   Sql = Sql & ", IDPAI, IDSETOR, GRUPO, SITCCUSTO, IDPROJ" & VbNewLine 
   Sql = Sql & ", IDEMPR, IDAUXILIAR, IDCOLIGADA, IDUSURESP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   If pAutoId Then
      Sql = Sql &  "(Select isNull(Max(IDCCUSTO),0)+1 From CCUSTO)"
   Else
      Sql = Sql &  SqlNum(mvarIDCCUSTO) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarDSCCCUSTO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTPAPROV) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODCCUSTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDSETOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarGRUPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSITCCUSTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPROJ) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDEMPR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDAUXILIAR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSURESP) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDCCUSTO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From CCUSTO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDCCUSTO)  Then 
      Sql = Sql & " IDCCUSTO = " & SqlNum(mvarIDCCUSTO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDCCUSTO) Then Sql = Sql & " IDCCUSTO = " & SqlNum(Cstr(Ch_IDCCUSTO)) & " AND " & VbNewLine 
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
 
   Sql = "Update CCUSTO Set " & VbNewLine 
   Sql = Sql & " IDCCUSTO   = " & SqlNum(mvarIDCCUSTO) & VbNewLine 
   Sql = Sql & " , DSCCCUSTO  = " & SqlStr(mvarDSCCCUSTO) & VbNewLine 
   Sql = Sql & " , TPAPROV    = " & SqlStr(mvarTPAPROV) & VbNewLine 
   Sql = Sql & " , CODCCUSTO  = " & SqlStr(mvarCODCCUSTO) & VbNewLine 
   Sql = Sql & " , IDPAI      = " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & " , IDSETOR    = " & SqlNum(mvarIDSETOR) & VbNewLine 
   Sql = Sql & " , GRUPO      = " & SqlStr(mvarGRUPO) & VbNewLine 
   Sql = Sql & " , SITCCUSTO  = " & SqlStr(mvarSITCCUSTO) & VbNewLine 
   Sql = Sql & " , IDPROJ     = " & SqlNum(mvarIDPROJ) & VbNewLine 
   Sql = Sql & " , IDEMPR     = " & SqlNum(mvarIDEMPR) & VbNewLine 
   Sql = Sql & " , IDAUXILIAR = " & SqlStr(mvarIDAUXILIAR) & VbNewLine 
   Sql = Sql & " , IDCOLIGADA = " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & " , IDUSURESP  = " & SqlStr(mvarIDUSURESP) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CCUSTO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDCCUSTO     = " & SqlNum(mvarIDCCUSTO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDCCUSTO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDCCUSTO, DSCCCUSTO, TPAPROV, CODCCUSTO" & VbNewLine 
   Sql = Sql & ", IDPAI, IDSETOR, GRUPO, SITCCUSTO, IDPROJ" & VbNewLine 
   Sql = Sql & ", IDEMPR, IDAUXILIAR, IDCOLIGADA, IDUSURESP" & VbNewLine 
   Sql = Sql &" From CCUSTO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDCCUSTO)  Then 
      Sql = Sql & " IDCCUSTO = " & SqlNum(mvarIDCCUSTO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDCCUSTO) Then Sql = Sql & " IDCCUSTO = " & SqlNum(Cstr(Ch_IDCCUSTO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDCCUSTO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDCCUSTO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCCUSTO = XVal(!IDCCUSTO & "")
      mvarDSCCCUSTO = !DSCCCUSTO & ""
      mvarTPAPROV = !TPAPROV & ""
      mvarCODCCUSTO = !CODCCUSTO & ""
      mvarIDPAI = XVal(!IDPAI & "")
      mvarIDSETOR = XVal(!IDSETOR & "")
      mvarGRUPO = !GRUPO & ""
      mvarSITCCUSTO = !SITCCUSTO & ""
      mvarIDPROJ = XVal(!IDPROJ & "")
      mvarIDEMPR = XVal(!IDEMPR & "")
      mvarIDAUXILIAR = !IDAUXILIAR & ""
      mvarIDCOLIGADA = XVal(!IDCOLIGADA & "")
      mvarIDUSURESP = !IDUSURESP & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDCCUSTO = 0
   mvarDSCCCUSTO = ""
   mvarTPAPROV = ""
   mvarCODCCUSTO = ""
   mvarIDPAI = 0
   mvarIDSETOR = 0
   mvarGRUPO = ""
   mvarSITCCUSTO = ""
   mvarIDPROJ = 0
   mvarIDEMPR = 0
   mvarIDAUXILIAR = ""
   mvarIDCOLIGADA = 0
   mvarIDUSURESP = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDCCUSTO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDCCUSTO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CCUSTO Set "
   Sql = Sql & " IDCCUSTO = " & SqlNum(Ch_IDCCUSTO)
   Sql = Sql & " Where 
   Sql = Sql & " IDCCUSTO = " & SqlNum(mvarIDCCUSTO)
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
