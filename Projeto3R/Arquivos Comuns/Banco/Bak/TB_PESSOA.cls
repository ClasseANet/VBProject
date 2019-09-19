VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_PESSOA"
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

Private mvarIDPESSOA    As Double
Private mvarCODPESSOA   As String
Private mvarNMPESSOA    As String
Private mvarRG          As String
Private mvarRGORGAO     As String
Private mvarRGEMISSAO   As String
Private mvarCPF         As String
Private mvarENDERECO    As String
Private mvarBAIRRO      As String
Private mvarCIDADE      As String
Private mvarUF          As String
Private mvarCOMPLEMENTO As String
Private mvarTELEFONE    As String
Private mvarCELULAR     As String
Private mvarIDEMPRESA   As Double
Private mvarFOTO        As Object
Private mvarAPELIDO     As String
Private mvarCRACHA      As String
Private mvarIDSETOR     As Double
Private mvarIDCCUSTO    As Double
Private mvarIDDEPTO     As Double
Private mvarEMAIL       As String
Private mvarNIVELOS     As Double
Private mvarLOCALTRAB   As String
Private mvarTPCLI       As String
Public Property Get IDPESSOA() As Double
   IDPESSOA = mvarIDPESSOA
End Property
Public Property Let CODPESSOA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODPESSOA <> vData)
   mvarCODPESSOA = vData
End Property
Public Property Get CODPESSOA() As String
   CODPESSOA = mvarCODPESSOA
End Property
Public Property Let NMPESSOA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMPESSOA <> vData)
   mvarNMPESSOA = vData
End Property
Public Property Get NMPESSOA() As String
   NMPESSOA = mvarNMPESSOA
End Property
Public Property Let RG(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRG <> vData)
   mvarRG = vData
End Property
Public Property Get RG() As String
   RG = mvarRG
End Property
Public Property Let RGORGAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRGORGAO <> vData)
   mvarRGORGAO = vData
End Property
Public Property Get RGORGAO() As String
   RGORGAO = mvarRGORGAO
End Property
Public Property Let RGEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRGEMISSAO <> vData)
   mvarRGEMISSAO = vData
End Property
Public Property Get RGEMISSAO() As String
   RGEMISSAO = mvarRGEMISSAO
End Property
Public Property Let CPF(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCPF <> vData)
   mvarCPF = vData
End Property
Public Property Get CPF() As String
   CPF = mvarCPF
End Property
Public Property Let ENDERECO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarENDERECO <> vData)
   mvarENDERECO = vData
End Property
Public Property Get ENDERECO() As String
   ENDERECO = mvarENDERECO
End Property
Public Property Let BAIRRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarBAIRRO <> vData)
   mvarBAIRRO = vData
End Property
Public Property Get BAIRRO() As String
   BAIRRO = mvarBAIRRO
End Property
Public Property Let CIDADE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCIDADE <> vData)
   mvarCIDADE = vData
End Property
Public Property Get CIDADE() As String
   CIDADE = mvarCIDADE
End Property
Public Property Let UF(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUF <> vData)
   mvarUF = vData
End Property
Public Property Get UF() As String
   UF = mvarUF
End Property
Public Property Let COMPLEMENTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCOMPLEMENTO <> vData)
   mvarCOMPLEMENTO = vData
End Property
Public Property Get COMPLEMENTO() As String
   COMPLEMENTO = mvarCOMPLEMENTO
End Property
Public Property Let TELEFONE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTELEFONE <> vData)
   mvarTELEFONE = vData
End Property
Public Property Get TELEFONE() As String
   TELEFONE = mvarTELEFONE
End Property
Public Property Let CELULAR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCELULAR <> vData)
   mvarCELULAR = vData
End Property
Public Property Get CELULAR() As String
   CELULAR = mvarCELULAR
End Property
Public Property Let IDEMPRESA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEMPRESA <> vData)
   mvarIDEMPRESA = vData
End Property
Public Property Get IDEMPRESA() As Double
   IDEMPRESA = mvarIDEMPRESA
End Property
Public Property Set FOTO(ByVal vData As Object)
   If Not mvarisDirt Then mvarisDirt = (mvarFOTO <> vData)
   Set mvarFOTO = vData
End Property
Public Property Get FOTO() As Object
   Set FOTO = mvarFOTO
End Property
Public Property Let APELIDO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarAPELIDO <> vData)
   mvarAPELIDO = vData
End Property
Public Property Get APELIDO() As String
   APELIDO = mvarAPELIDO
End Property
Public Property Let CRACHA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCRACHA <> vData)
   mvarCRACHA = vData
End Property
Public Property Get CRACHA() As String
   CRACHA = mvarCRACHA
End Property
Public Property Let IDSETOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSETOR <> vData)
   mvarIDSETOR = vData
End Property
Public Property Get IDSETOR() As Double
   IDSETOR = mvarIDSETOR
End Property
Public Property Let IDCCUSTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCCUSTO <> vData)
   mvarIDCCUSTO = vData
End Property
Public Property Get IDCCUSTO() As Double
   IDCCUSTO = mvarIDCCUSTO
End Property
Public Property Let IDDEPTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDEPTO <> vData)
   mvarIDDEPTO = vData
End Property
Public Property Get IDDEPTO() As Double
   IDDEPTO = mvarIDDEPTO
End Property
Public Property Let EMAIL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEMAIL <> vData)
   mvarEMAIL = vData
End Property
Public Property Get EMAIL() As String
   EMAIL = mvarEMAIL
End Property
Public Property Let NIVELOS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNIVELOS <> vData)
   mvarNIVELOS = vData
End Property
Public Property Get NIVELOS() As Double
   NIVELOS = mvarNIVELOS
End Property
Public Property Let LOCALTRAB(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLOCALTRAB <> vData)
   mvarLOCALTRAB = vData
End Property
Public Property Get LOCALTRAB() As String
   LOCALTRAB = mvarLOCALTRAB
End Property
Public Property Let TPCLI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPCLI <> vData)
   mvarTPCLI = vData
End Property
Public Property Get TPCLI() As String
   TPCLI = mvarTPCLI
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
 
   Sql = "Insert Into PESSOA (CODPESSOA, NMPESSOA, RG" & VbNewLine 
   Sql = Sql & ", RGORGAO, RGEMISSAO, CPF, ENDERECO, BAIRRO" & VbNewLine 
   Sql = Sql & ", CIDADE, UF, COMPLEMENTO, TELEFONE, CELULAR" & VbNewLine 
   Sql = Sql & ", IDEMPRESA, APELIDO, CRACHA, IDSETOR" & VbNewLine 
   Sql = Sql & ", IDCCUSTO, IDDEPTO, EMAIL, NIVELOS, LOCALTRAB" & VbNewLine 
   Sql = Sql & ", TPCLI) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarCODPESSOA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMPESSOA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarRG) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarRGORGAO) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarRGEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCPF) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarENDERECO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUF) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCOMPLEMENTO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTELEFONE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCELULAR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDEMPRESA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarAPELIDO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCRACHA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDSETOR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCCUSTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDDEPTO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarNIVELOS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLOCALTRAB) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTPCLI) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPESSOA) As String
   Dim Sql As String
 
   Sql = "Delete From PESSOA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDPESSOA) Then Sql = Sql & " IDPESSOA = " & SqlNum(Cstr(Ch_IDPESSOA)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update PESSOA Set " & VbNewLine 
   Sql = Sql & " CODPESSOA   = " & SqlStr(mvarCODPESSOA) & VbNewLine 
   Sql = Sql & " , NMPESSOA    = " & SqlStr(mvarNMPESSOA) & VbNewLine 
   Sql = Sql & " , RG          = " & SqlStr(mvarRG) & VbNewLine 
   Sql = Sql & " , RGORGAO     = " & SqlStr(mvarRGORGAO) & VbNewLine 
   Sql = Sql & " , RGEMISSAO   = " & SqlDate(mvarRGEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , CPF         = " & SqlStr(mvarCPF) & VbNewLine 
   Sql = Sql & " , ENDERECO    = " & SqlStr(mvarENDERECO) & VbNewLine 
   Sql = Sql & " , BAIRRO      = " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & " , CIDADE      = " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & " , UF          = " & SqlStr(mvarUF) & VbNewLine 
   Sql = Sql & " , COMPLEMENTO = " & SqlStr(mvarCOMPLEMENTO) & VbNewLine 
   Sql = Sql & " , TELEFONE    = " & SqlStr(mvarTELEFONE) & VbNewLine 
   Sql = Sql & " , CELULAR     = " & SqlStr(mvarCELULAR) & VbNewLine 
   Sql = Sql & " , IDEMPRESA   = " & SqlNum(mvarIDEMPRESA) & VbNewLine 
   Sql = Sql & " , APELIDO     = " & SqlStr(mvarAPELIDO) & VbNewLine 
   Sql = Sql & " , CRACHA      = " & SqlStr(mvarCRACHA) & VbNewLine 
   Sql = Sql & " , IDSETOR     = " & SqlNum(mvarIDSETOR) & VbNewLine 
   Sql = Sql & " , IDCCUSTO    = " & SqlNum(mvarIDCCUSTO) & VbNewLine 
   Sql = Sql & " , IDDEPTO     = " & SqlNum(mvarIDDEPTO) & VbNewLine 
   Sql = Sql & " , EMAIL       = " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & " , NIVELOS     = " & SqlNum(mvarNIVELOS) & VbNewLine 
   Sql = Sql & " , LOCALTRAB   = " & SqlStr(mvarLOCALTRAB) & VbNewLine 
   Sql = Sql & " , TPCLI       = " & SqlStr(mvarTPCLI) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'PESSOA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPESSOA      = " & SqlNum(mvarIDPESSOA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDPESSOA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPESSOA, CODPESSOA, NMPESSOA, RG" & VbNewLine 
   Sql = Sql & ", RGORGAO, RGEMISSAO, CPF, ENDERECO, BAIRRO" & VbNewLine 
   Sql = Sql & ", CIDADE, UF, COMPLEMENTO, TELEFONE, CELULAR" & VbNewLine 
   Sql = Sql & ", IDEMPRESA, APELIDO, CRACHA, IDSETOR" & VbNewLine 
   Sql = Sql & ", IDCCUSTO, IDDEPTO, EMAIL, NIVELOS, LOCALTRAB" & VbNewLine 
   Sql = Sql & ", TPCLI" & VbNewLine 
   Sql = Sql &" From PESSOA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPESSOA)  Then 
      Sql = Sql & " IDPESSOA = " & SqlNum(mvarIDPESSOA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDPESSOA) Then Sql = Sql & " IDPESSOA = " & SqlNum(Cstr(Ch_IDPESSOA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDPESSOA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPESSOA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPESSOA = XVal(!IDPESSOA & "")
      mvarCODPESSOA = !CODPESSOA & ""
      mvarNMPESSOA = !NMPESSOA & ""
      mvarRG = !RG & ""
      mvarRGORGAO = !RGORGAO & ""
      mvarRGEMISSAO = xDate(!RGEMISSAO & "", True)
      mvarCPF = !CPF & ""
      mvarENDERECO = !ENDERECO & ""
      mvarBAIRRO = !BAIRRO & ""
      mvarCIDADE = !CIDADE & ""
      mvarUF = !UF & ""
      mvarCOMPLEMENTO = !COMPLEMENTO & ""
      mvarTELEFONE = !TELEFONE & ""
      mvarCELULAR = !CELULAR & ""
      mvarIDEMPRESA = XVal(!IDEMPRESA & "")
      mvarIDEMPRESA = XVal(!IDEMPRESA & "")
      mvarAPELIDO = !APELIDO & ""
      mvarCRACHA = !CRACHA & ""
      mvarIDSETOR = XVal(!IDSETOR & "")
      mvarIDCCUSTO = XVal(!IDCCUSTO & "")
      mvarIDDEPTO = XVal(!IDDEPTO & "")
      mvarEMAIL = !EMAIL & ""
      mvarNIVELOS = XVal(!NIVELOS & "")
      mvarLOCALTRAB = !LOCALTRAB & ""
      mvarTPCLI = !TPCLI & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPESSOA), ComCOMMIT)
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
