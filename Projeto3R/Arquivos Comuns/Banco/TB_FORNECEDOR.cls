VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_FORNECEDOR"
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
Private mvarIDFOR        As Double
Private mvarNMFOR        As String
Private mvarRZFOR        As String
Private mvarCNPJFOR      As String
Private mvarINSCMUNIC    As String
Private mvarTELFOR1      As String
Private mvarTELFOR2      As String
Private mvarFAXFOR       As String
Private mvarEMAIL        As String
Private mvarENDFOR       As String
Private mvarBAIFOR       As String
Private mvarCIDFOR       As String
Private mvarESTADOFOR    As String
Private mvarCEPFOR       As String
Private mvarNMCONTATO    As String
Private mvarCARGOCONTATO As String
Private mvarCONDPGTO     As String
Private mvarGRUPO        As String
Private mvarFORMAPGTO    As String
Private mvarSIGLAPAIS    As String
Private mvarDTCADASTRO   As String
Private mvarQLFFOR       As String
Private mvarIDUSU        As String
Private mvarINSCEST      As String
Private mvarINSCOUTROS   As String
Private mvarPJ           As String
Private mvarOBS          As String
Private mvarTELCONTATO1  As String
Private mvarTELCONTATO2  As String
Private mvarFAXCONTATO   As String
Private mvarEMAILCONTATO As String
Private mvarNUMBANCO     As String
Private mvarNMBANCO      As String
Private mvarNUMAGENCIA   As String
Private mvarNMAGENCIA    As String
Private mvarNUMCONTA     As String
Private mvarOPCSIMPLES   As String
Private mvarSERVICO      As String
Private mvarMATERIAL     As String
Private mvarVERIFDOC     As String
Private mvarDSCPAIS      As String
Private mvarIDPAIS       As Double
Private mvarIDMATRIZ     As Double
Private mvarIDAUXILIAR   As String
Private mvarDVCONTA      As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDFOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFOR <> vData)
   mvarIDFOR = vData
End Property
Public Property Get IDFOR() As Double
   IDFOR = mvarIDFOR
End Property
Public Property Let NMFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMFOR <> vData)
   mvarNMFOR = vData
End Property
Public Property Get NMFOR() As String
   NMFOR = mvarNMFOR
End Property
Public Property Let RZFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRZFOR <> vData)
   mvarRZFOR = vData
End Property
Public Property Get RZFOR() As String
   RZFOR = mvarRZFOR
End Property
Public Property Let CNPJFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCNPJFOR <> vData)
   mvarCNPJFOR = vData
End Property
Public Property Get CNPJFOR() As String
   CNPJFOR = mvarCNPJFOR
End Property
Public Property Let INSCMUNIC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarINSCMUNIC <> vData)
   mvarINSCMUNIC = vData
End Property
Public Property Get INSCMUNIC() As String
   INSCMUNIC = mvarINSCMUNIC
End Property
Public Property Let TELFOR1(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTELFOR1 <> vData)
   mvarTELFOR1 = vData
End Property
Public Property Get TELFOR1() As String
   TELFOR1 = mvarTELFOR1
End Property
Public Property Let TELFOR2(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTELFOR2 <> vData)
   mvarTELFOR2 = vData
End Property
Public Property Get TELFOR2() As String
   TELFOR2 = mvarTELFOR2
End Property
Public Property Let FAXFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFAXFOR <> vData)
   mvarFAXFOR = vData
End Property
Public Property Get FAXFOR() As String
   FAXFOR = mvarFAXFOR
End Property
Public Property Let EMAIL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEMAIL <> vData)
   mvarEMAIL = vData
End Property
Public Property Get EMAIL() As String
   EMAIL = mvarEMAIL
End Property
Public Property Let ENDFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarENDFOR <> vData)
   mvarENDFOR = vData
End Property
Public Property Get ENDFOR() As String
   ENDFOR = mvarENDFOR
End Property
Public Property Let BAIFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarBAIFOR <> vData)
   mvarBAIFOR = vData
End Property
Public Property Get BAIFOR() As String
   BAIFOR = mvarBAIFOR
End Property
Public Property Let CIDFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCIDFOR <> vData)
   mvarCIDFOR = vData
End Property
Public Property Get CIDFOR() As String
   CIDFOR = mvarCIDFOR
End Property
Public Property Let ESTADOFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarESTADOFOR <> vData)
   mvarESTADOFOR = vData
End Property
Public Property Get ESTADOFOR() As String
   ESTADOFOR = mvarESTADOFOR
End Property
Public Property Let CEPFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCEPFOR <> vData)
   mvarCEPFOR = vData
End Property
Public Property Get CEPFOR() As String
   CEPFOR = mvarCEPFOR
End Property
Public Property Let NMCONTATO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMCONTATO <> vData)
   mvarNMCONTATO = vData
End Property
Public Property Get NMCONTATO() As String
   NMCONTATO = mvarNMCONTATO
End Property
Public Property Let CARGOCONTATO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCARGOCONTATO <> vData)
   mvarCARGOCONTATO = vData
End Property
Public Property Get CARGOCONTATO() As String
   CARGOCONTATO = mvarCARGOCONTATO
End Property
Public Property Let CONDPGTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCONDPGTO <> vData)
   mvarCONDPGTO = vData
End Property
Public Property Get CONDPGTO() As String
   CONDPGTO = mvarCONDPGTO
End Property
Public Property Let GRUPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarGRUPO <> vData)
   mvarGRUPO = vData
End Property
Public Property Get GRUPO() As String
   GRUPO = mvarGRUPO
End Property
Public Property Let FORMAPGTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFORMAPGTO <> vData)
   mvarFORMAPGTO = vData
End Property
Public Property Get FORMAPGTO() As String
   FORMAPGTO = mvarFORMAPGTO
End Property
Public Property Let SIGLAPAIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSIGLAPAIS <> vData)
   mvarSIGLAPAIS = vData
End Property
Public Property Get SIGLAPAIS() As String
   SIGLAPAIS = mvarSIGLAPAIS
End Property
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let QLFFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarQLFFOR <> vData)
   mvarQLFFOR = vData
End Property
Public Property Get QLFFOR() As String
   QLFFOR = mvarQLFFOR
End Property
Public Property Let IDUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSU <> vData)
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let INSCEST(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarINSCEST <> vData)
   mvarINSCEST = vData
End Property
Public Property Get INSCEST() As String
   INSCEST = mvarINSCEST
End Property
Public Property Let INSCOUTROS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarINSCOUTROS <> vData)
   mvarINSCOUTROS = vData
End Property
Public Property Get INSCOUTROS() As String
   INSCOUTROS = mvarINSCOUTROS
End Property
Public Property Let PJ(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPJ <> vData)
   mvarPJ = vData
End Property
Public Property Get PJ() As String
   PJ = mvarPJ
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
End Property
Public Property Let TELCONTATO1(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTELCONTATO1 <> vData)
   mvarTELCONTATO1 = vData
End Property
Public Property Get TELCONTATO1() As String
   TELCONTATO1 = mvarTELCONTATO1
End Property
Public Property Let TELCONTATO2(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTELCONTATO2 <> vData)
   mvarTELCONTATO2 = vData
End Property
Public Property Get TELCONTATO2() As String
   TELCONTATO2 = mvarTELCONTATO2
End Property
Public Property Let FAXCONTATO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFAXCONTATO <> vData)
   mvarFAXCONTATO = vData
End Property
Public Property Get FAXCONTATO() As String
   FAXCONTATO = mvarFAXCONTATO
End Property
Public Property Let EMAILCONTATO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEMAILCONTATO <> vData)
   mvarEMAILCONTATO = vData
End Property
Public Property Get EMAILCONTATO() As String
   EMAILCONTATO = mvarEMAILCONTATO
End Property
Public Property Let NUMBANCO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMBANCO <> vData)
   mvarNUMBANCO = vData
End Property
Public Property Get NUMBANCO() As String
   NUMBANCO = mvarNUMBANCO
End Property
Public Property Let NMBANCO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMBANCO <> vData)
   mvarNMBANCO = vData
End Property
Public Property Get NMBANCO() As String
   NMBANCO = mvarNMBANCO
End Property
Public Property Let NUMAGENCIA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMAGENCIA <> vData)
   mvarNUMAGENCIA = vData
End Property
Public Property Get NUMAGENCIA() As String
   NUMAGENCIA = mvarNUMAGENCIA
End Property
Public Property Let NMAGENCIA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMAGENCIA <> vData)
   mvarNMAGENCIA = vData
End Property
Public Property Get NMAGENCIA() As String
   NMAGENCIA = mvarNMAGENCIA
End Property
Public Property Let NUMCONTA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMCONTA <> vData)
   mvarNUMCONTA = vData
End Property
Public Property Get NUMCONTA() As String
   NUMCONTA = mvarNUMCONTA
End Property
Public Property Let OPCSIMPLES(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOPCSIMPLES <> vData)
   mvarOPCSIMPLES = vData
End Property
Public Property Get OPCSIMPLES() As String
   OPCSIMPLES = mvarOPCSIMPLES
End Property
Public Property Let SERVICO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSERVICO <> vData)
   mvarSERVICO = vData
End Property
Public Property Get SERVICO() As String
   SERVICO = mvarSERVICO
End Property
Public Property Let MATERIAL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMATERIAL <> vData)
   mvarMATERIAL = vData
End Property
Public Property Get MATERIAL() As String
   MATERIAL = mvarMATERIAL
End Property
Public Property Let VERIFDOC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVERIFDOC <> vData)
   mvarVERIFDOC = vData
End Property
Public Property Get VERIFDOC() As String
   VERIFDOC = mvarVERIFDOC
End Property
Public Property Let DSCPAIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCPAIS <> vData)
   mvarDSCPAIS = vData
End Property
Public Property Get DSCPAIS() As String
   DSCPAIS = mvarDSCPAIS
End Property
Public Property Let IDPAIS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAIS <> vData)
   mvarIDPAIS = vData
End Property
Public Property Get IDPAIS() As Double
   IDPAIS = mvarIDPAIS
End Property
Public Property Let IDMATRIZ(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMATRIZ <> vData)
   mvarIDMATRIZ = vData
End Property
Public Property Get IDMATRIZ() As Double
   IDMATRIZ = mvarIDMATRIZ
End Property
Public Property Let IDAUXILIAR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAUXILIAR <> vData)
   mvarIDAUXILIAR = vData
End Property
Public Property Get IDAUXILIAR() As String
   IDAUXILIAR = mvarIDAUXILIAR
End Property
Public Property Let DVCONTA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDVCONTA <> vData)
   mvarDVCONTA = vData
End Property
Public Property Get DVCONTA() As String
   DVCONTA = mvarDVCONTA
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
 
   Sql = "Insert Into FORNECEDOR (IDLOJA, IDFOR, NMFOR, RZFOR" & VbNewLine 
   Sql = Sql & ", CNPJFOR, INSCMUNIC, TELFOR1, TELFOR2, FAXFOR" & VbNewLine 
   Sql = Sql & ", EMAIL, ENDFOR, BAIFOR, CIDFOR, ESTADOFOR" & VbNewLine 
   Sql = Sql & ", CEPFOR, NMCONTATO, CARGOCONTATO, CONDPGTO, GRUPO" & VbNewLine 
   Sql = Sql & ", FORMAPGTO, SIGLAPAIS, DTCADASTRO, QLFFOR, IDUSU" & VbNewLine 
   Sql = Sql & ", INSCEST, INSCOUTROS, PJ, OBS, TELCONTATO1" & VbNewLine 
   Sql = Sql & ", TELCONTATO2, FAXCONTATO, EMAILCONTATO, NUMBANCO, NMBANCO" & VbNewLine 
   Sql = Sql & ", NUMAGENCIA, NMAGENCIA, NUMCONTA, OPCSIMPLES, SERVICO" & VbNewLine 
   Sql = Sql & ", MATERIAL, VERIFDOC, DSCPAIS, IDPAIS, IDMATRIZ" & VbNewLine 
   Sql = Sql & ", IDAUXILIAR, DVCONTA) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarRZFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCNPJFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarINSCMUNIC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTELFOR1) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTELFOR2) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFAXFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarENDFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarBAIFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCIDFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarESTADOFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCEPFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMCONTATO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCARGOCONTATO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCONDPGTO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarGRUPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFORMAPGTO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSIGLAPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarQLFFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarINSCEST) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarINSCOUTROS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPJ) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTELCONTATO1) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTELCONTATO2) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFAXCONTATO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMAILCONTATO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNUMBANCO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMBANCO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNUMAGENCIA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMAGENCIA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNUMCONTA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOPCSIMPLES) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSERVICO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMATERIAL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVERIFDOC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDMATRIZ) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDAUXILIAR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDVCONTA) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDFOR, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From FORNECEDOR" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDFOR)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDFOR = " & SqlNum(mvarIDFOR) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFOR) Then Sql = Sql & " IDFOR = " & SqlNum(Cstr(Ch_IDFOR)) & " AND " & VbNewLine 
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
 
   Sql = "Update FORNECEDOR Set " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDFOR        = " & SqlNum(mvarIDFOR) & VbNewLine 
   Sql = Sql & " , NMFOR        = " & SqlStr(mvarNMFOR) & VbNewLine 
   Sql = Sql & " , RZFOR        = " & SqlStr(mvarRZFOR) & VbNewLine 
   Sql = Sql & " , CNPJFOR      = " & SqlStr(mvarCNPJFOR) & VbNewLine 
   Sql = Sql & " , INSCMUNIC    = " & SqlStr(mvarINSCMUNIC) & VbNewLine 
   Sql = Sql & " , TELFOR1      = " & SqlStr(mvarTELFOR1) & VbNewLine 
   Sql = Sql & " , TELFOR2      = " & SqlStr(mvarTELFOR2) & VbNewLine 
   Sql = Sql & " , FAXFOR       = " & SqlStr(mvarFAXFOR) & VbNewLine 
   Sql = Sql & " , EMAIL        = " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & " , ENDFOR       = " & SqlStr(mvarENDFOR) & VbNewLine 
   Sql = Sql & " , BAIFOR       = " & SqlStr(mvarBAIFOR) & VbNewLine 
   Sql = Sql & " , CIDFOR       = " & SqlStr(mvarCIDFOR) & VbNewLine 
   Sql = Sql & " , ESTADOFOR    = " & SqlStr(mvarESTADOFOR) & VbNewLine 
   Sql = Sql & " , CEPFOR       = " & SqlStr(mvarCEPFOR) & VbNewLine 
   Sql = Sql & " , NMCONTATO    = " & SqlStr(mvarNMCONTATO) & VbNewLine 
   Sql = Sql & " , CARGOCONTATO = " & SqlStr(mvarCARGOCONTATO) & VbNewLine 
   Sql = Sql & " , CONDPGTO     = " & SqlStr(mvarCONDPGTO) & VbNewLine 
   Sql = Sql & " , GRUPO        = " & SqlStr(mvarGRUPO) & VbNewLine 
   Sql = Sql & " , FORMAPGTO    = " & SqlStr(mvarFORMAPGTO) & VbNewLine 
   Sql = Sql & " , SIGLAPAIS    = " & SqlStr(mvarSIGLAPAIS) & VbNewLine 
   Sql = Sql & " , DTCADASTRO   = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , QLFFOR       = " & SqlStr(mvarQLFFOR) & VbNewLine 
   Sql = Sql & " , IDUSU        = " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & " , INSCEST      = " & SqlStr(mvarINSCEST) & VbNewLine 
   Sql = Sql & " , INSCOUTROS   = " & SqlStr(mvarINSCOUTROS) & VbNewLine 
   Sql = Sql & " , PJ           = " & SqlStr(mvarPJ) & VbNewLine 
   Sql = Sql & " , OBS          = " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & " , TELCONTATO1  = " & SqlStr(mvarTELCONTATO1) & VbNewLine 
   Sql = Sql & " , TELCONTATO2  = " & SqlStr(mvarTELCONTATO2) & VbNewLine 
   Sql = Sql & " , FAXCONTATO   = " & SqlStr(mvarFAXCONTATO) & VbNewLine 
   Sql = Sql & " , EMAILCONTATO = " & SqlStr(mvarEMAILCONTATO) & VbNewLine 
   Sql = Sql & " , NUMBANCO     = " & SqlStr(mvarNUMBANCO) & VbNewLine 
   Sql = Sql & " , NMBANCO      = " & SqlStr(mvarNMBANCO) & VbNewLine 
   Sql = Sql & " , NUMAGENCIA   = " & SqlStr(mvarNUMAGENCIA) & VbNewLine 
   Sql = Sql & " , NMAGENCIA    = " & SqlStr(mvarNMAGENCIA) & VbNewLine 
   Sql = Sql & " , NUMCONTA     = " & SqlStr(mvarNUMCONTA) & VbNewLine 
   Sql = Sql & " , OPCSIMPLES   = " & SqlStr(mvarOPCSIMPLES) & VbNewLine 
   Sql = Sql & " , SERVICO      = " & SqlStr(mvarSERVICO) & VbNewLine 
   Sql = Sql & " , MATERIAL     = " & SqlStr(mvarMATERIAL) & VbNewLine 
   Sql = Sql & " , VERIFDOC     = " & SqlStr(mvarVERIFDOC) & VbNewLine 
   Sql = Sql & " , DSCPAIS      = " & SqlStr(mvarDSCPAIS) & VbNewLine 
   Sql = Sql & " , IDPAIS       = " & SqlNum(mvarIDPAIS) & VbNewLine 
   Sql = Sql & " , IDMATRIZ     = " & SqlNum(mvarIDMATRIZ) & VbNewLine 
   Sql = Sql & " , IDAUXILIAR   = " & SqlStr(mvarIDAUXILIAR) & VbNewLine 
   Sql = Sql & " , DVCONTA      = " & SqlStr(mvarDVCONTA) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FORNECEDOR'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDFOR      = " & SqlNum(mvarIDFOR) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDFOR, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDFOR, NMFOR, RZFOR" & VbNewLine 
   Sql = Sql & ", CNPJFOR, INSCMUNIC, TELFOR1, TELFOR2, FAXFOR" & VbNewLine 
   Sql = Sql & ", EMAIL, ENDFOR, BAIFOR, CIDFOR, ESTADOFOR" & VbNewLine 
   Sql = Sql & ", CEPFOR, NMCONTATO, CARGOCONTATO, CONDPGTO, GRUPO" & VbNewLine 
   Sql = Sql & ", FORMAPGTO, SIGLAPAIS, DTCADASTRO, QLFFOR, IDUSU" & VbNewLine 
   Sql = Sql & ", INSCEST, INSCOUTROS, PJ, OBS, TELCONTATO1" & VbNewLine 
   Sql = Sql & ", TELCONTATO2, FAXCONTATO, EMAILCONTATO, NUMBANCO, NMBANCO" & VbNewLine 
   Sql = Sql & ", NUMAGENCIA, NMAGENCIA, NUMCONTA, OPCSIMPLES, SERVICO" & VbNewLine 
   Sql = Sql & ", MATERIAL, VERIFDOC, DSCPAIS, IDPAIS, IDMATRIZ" & VbNewLine 
   Sql = Sql & ", IDAUXILIAR, DVCONTA" & VbNewLine 
   Sql = Sql &" From FORNECEDOR" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDFOR)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDFOR = " & SqlNum(mvarIDFOR) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFOR) Then Sql = Sql & " IDFOR = " & SqlNum(Cstr(Ch_IDFOR)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDFOR, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDFOR, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDFOR = XVal(!IDFOR & "")
      mvarNMFOR = !NMFOR & ""
      mvarRZFOR = !RZFOR & ""
      mvarCNPJFOR = !CNPJFOR & ""
      mvarINSCMUNIC = !INSCMUNIC & ""
      mvarTELFOR1 = !TELFOR1 & ""
      mvarTELFOR2 = !TELFOR2 & ""
      mvarFAXFOR = !FAXFOR & ""
      mvarEMAIL = !EMAIL & ""
      mvarENDFOR = !ENDFOR & ""
      mvarBAIFOR = !BAIFOR & ""
      mvarCIDFOR = !CIDFOR & ""
      mvarESTADOFOR = !ESTADOFOR & ""
      mvarCEPFOR = !CEPFOR & ""
      mvarNMCONTATO = !NMCONTATO & ""
      mvarCARGOCONTATO = !CARGOCONTATO & ""
      mvarCONDPGTO = !CONDPGTO & ""
      mvarGRUPO = !GRUPO & ""
      mvarFORMAPGTO = !FORMAPGTO & ""
      mvarSIGLAPAIS = !SIGLAPAIS & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarQLFFOR = !QLFFOR & ""
      mvarIDUSU = !IDUSU & ""
      mvarINSCEST = !INSCEST & ""
      mvarINSCOUTROS = !INSCOUTROS & ""
      mvarPJ = !PJ & ""
      mvarOBS = !OBS & ""
      mvarTELCONTATO1 = !TELCONTATO1 & ""
      mvarTELCONTATO2 = !TELCONTATO2 & ""
      mvarFAXCONTATO = !FAXCONTATO & ""
      mvarEMAILCONTATO = !EMAILCONTATO & ""
      mvarNUMBANCO = !NUMBANCO & ""
      mvarNMBANCO = !NMBANCO & ""
      mvarNUMAGENCIA = !NUMAGENCIA & ""
      mvarNMAGENCIA = !NMAGENCIA & ""
      mvarNUMCONTA = !NUMCONTA & ""
      mvarOPCSIMPLES = !OPCSIMPLES & ""
      mvarSERVICO = !SERVICO & ""
      mvarMATERIAL = !MATERIAL & ""
      mvarVERIFDOC = !VERIFDOC & ""
      mvarDSCPAIS = !DSCPAIS & ""
      mvarIDPAIS = XVal(!IDPAIS & "")
      mvarIDMATRIZ = XVal(!IDMATRIZ & "")
      mvarIDAUXILIAR = !IDAUXILIAR & ""
      mvarDVCONTA = !DVCONTA & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDFOR = 0
   mvarNMFOR = ""
   mvarRZFOR = ""
   mvarCNPJFOR = ""
   mvarINSCMUNIC = ""
   mvarTELFOR1 = ""
   mvarTELFOR2 = ""
   mvarFAXFOR = ""
   mvarEMAIL = ""
   mvarENDFOR = ""
   mvarBAIFOR = ""
   mvarCIDFOR = ""
   mvarESTADOFOR = ""
   mvarCEPFOR = ""
   mvarNMCONTATO = ""
   mvarCARGOCONTATO = ""
   mvarCONDPGTO = ""
   mvarGRUPO = ""
   mvarFORMAPGTO = ""
   mvarSIGLAPAIS = ""
   mvarDTCADASTRO = ""
   mvarQLFFOR = ""
   mvarIDUSU = ""
   mvarINSCEST = ""
   mvarINSCOUTROS = ""
   mvarPJ = ""
   mvarOBS = ""
   mvarTELCONTATO1 = ""
   mvarTELCONTATO2 = ""
   mvarFAXCONTATO = ""
   mvarEMAILCONTATO = ""
   mvarNUMBANCO = ""
   mvarNMBANCO = ""
   mvarNUMAGENCIA = ""
   mvarNMAGENCIA = ""
   mvarNUMCONTA = ""
   mvarOPCSIMPLES = ""
   mvarSERVICO = ""
   mvarMATERIAL = ""
   mvarVERIFDOC = ""
   mvarDSCPAIS = ""
   mvarIDPAIS = 0
   mvarIDMATRIZ = 0
   mvarIDAUXILIAR = ""
   mvarDVCONTA = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDFOR), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDFOR As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FORNECEDOR Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDFOR = " & SqlNum(Ch_IDFOR)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDFOR = " & SqlNum(mvarIDFOR)
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
