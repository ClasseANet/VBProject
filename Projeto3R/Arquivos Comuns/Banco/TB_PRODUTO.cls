VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_PRODUTO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb            As DS_BANCO
Private mvarRS             As Object 

Private mvarQryInsert      As String
Private mvarQryUpDate      As String
Private mvarQryDelete      As String
Private mvarQrySelect      As String
Private mvarQrySave        As String

Private mvarisDirt         As Boolean

Private mvarIDPROD       As Double
Private mvarCODPROD      As String
Private mvarNMPROD       As String
Private mvarDSCPROD      As String
Private mvarSIGLAUNID    As String
Private mvarEGRUPO       As String
Private mvarIDPAI        As Double
Private mvarIDUSU        As String
Private mvarDTCADASTRO   As String
Private mvarVLESTIMADO   As Double
Private mvarIDGRPFOR     As Double
Private mvarIDDESP       As String
Private mvarMSKCOD       As String
Private mvarQTDCOMPRA    As Double
Private mvarUNIDCOMPRA   As String
Private mvarCODBARRA     As String
Private mvarDTDESATIVO   As String
Private mvarVLULTCOMPRA  As Double
Private mvarVLMEDIO      As Double
Private mvarQTDESTOQUE   As Double
Private mvarQTDMIN       As Double
Private mvarQTDMAX       As Double
Private mvarPESOLIQ      As Double
Private mvarQTDRP        As Double
Private mvarPESOBRUTO    As Double
Private mvarCOMPRIMENTO  As Double
Private mvarLARGURA      As Double
Private mvarESPESSURA    As Double
Private mvarEMATPRIMA    As String
Private mvarEPRODACABADO As String
Private mvarEICMS        As String
Private mvarSITPROD      As String
Private mvarIDFAM        As Double
Private mvarINFORMATICA  As String
Private mvarATIVOFIXO    As String
Private mvarQUALIFICADO  As String
Private mvarIDDISCIPLINA As Double
Private mvarNMPRODIng    As String
Private mvarDSCPRODIng   As String
Private mvarTXCONVUNID   As String
Private mvarDIAMETRO     As Double
Private mvarOBSSMS       As String
Private mvarIDGRPSMS     As Double
Private mvarTAG          As String
Private mvarIDCLASSE     As Double
Private mvarIDFABRICANTE As Double
Private mvarALTURA       As Double
Private mvarTIPOACO      As String
Public Property Let IDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROD <> vData)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Double
   IDPROD = mvarIDPROD
End Property
Public Property Let CODPROD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODPROD <> vData)
   mvarCODPROD = vData
End Property
Public Property Get CODPROD() As String
   CODPROD = mvarCODPROD
End Property
Public Property Let NMPROD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMPROD <> vData)
   mvarNMPROD = vData
End Property
Public Property Get NMPROD() As String
   NMPROD = mvarNMPROD
End Property
Public Property Let DSCPROD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCPROD <> vData)
   mvarDSCPROD = vData
End Property
Public Property Get DSCPROD() As String
   DSCPROD = mvarDSCPROD
End Property
Public Property Let SIGLAUNID(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSIGLAUNID <> vData)
   mvarSIGLAUNID = vData
End Property
Public Property Get SIGLAUNID() As String
   SIGLAUNID = mvarSIGLAUNID
End Property
Public Property Let EGRUPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEGRUPO <> vData)
   mvarEGRUPO = vData
End Property
Public Property Get EGRUPO() As String
   EGRUPO = mvarEGRUPO
End Property
Public Property Let IDPAI(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAI <> vData)
   mvarIDPAI = vData
End Property
Public Property Get IDPAI() As Double
   IDPAI = mvarIDPAI
End Property
Public Property Let IDUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSU <> vData)
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let VLESTIMADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLESTIMADO <> vData)
   mvarVLESTIMADO = vData
End Property
Public Property Get VLESTIMADO() As Double
   VLESTIMADO = mvarVLESTIMADO
End Property
Public Property Let IDGRPFOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDGRPFOR <> vData)
   mvarIDGRPFOR = vData
End Property
Public Property Get IDGRPFOR() As Double
   IDGRPFOR = mvarIDGRPFOR
End Property
Public Property Let IDDESP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDESP <> vData)
   mvarIDDESP = vData
End Property
Public Property Get IDDESP() As String
   IDDESP = mvarIDDESP
End Property
Public Property Let MSKCOD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMSKCOD <> vData)
   mvarMSKCOD = vData
End Property
Public Property Get MSKCOD() As String
   MSKCOD = mvarMSKCOD
End Property
Public Property Let QTDCOMPRA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDCOMPRA <> vData)
   mvarQTDCOMPRA = vData
End Property
Public Property Get QTDCOMPRA() As Double
   QTDCOMPRA = mvarQTDCOMPRA
End Property
Public Property Let UNIDCOMPRA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUNIDCOMPRA <> vData)
   mvarUNIDCOMPRA = vData
End Property
Public Property Get UNIDCOMPRA() As String
   UNIDCOMPRA = mvarUNIDCOMPRA
End Property
Public Property Let CODBARRA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODBARRA <> vData)
   mvarCODBARRA = vData
End Property
Public Property Get CODBARRA() As String
   CODBARRA = mvarCODBARRA
End Property
Public Property Let DTDESATIVO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTDESATIVO <> vData)
   mvarDTDESATIVO = vData
End Property
Public Property Get DTDESATIVO() As String
   DTDESATIVO = mvarDTDESATIVO
End Property
Public Property Let VLULTCOMPRA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLULTCOMPRA <> vData)
   mvarVLULTCOMPRA = vData
End Property
Public Property Get VLULTCOMPRA() As Double
   VLULTCOMPRA = mvarVLULTCOMPRA
End Property
Public Property Let VLMEDIO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLMEDIO <> vData)
   mvarVLMEDIO = vData
End Property
Public Property Get VLMEDIO() As Double
   VLMEDIO = mvarVLMEDIO
End Property
Public Property Let QTDESTOQUE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDESTOQUE <> vData)
   mvarQTDESTOQUE = vData
End Property
Public Property Get QTDESTOQUE() As Double
   QTDESTOQUE = mvarQTDESTOQUE
End Property
Public Property Let QTDMIN(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDMIN <> vData)
   mvarQTDMIN = vData
End Property
Public Property Get QTDMIN() As Double
   QTDMIN = mvarQTDMIN
End Property
Public Property Let QTDMAX(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDMAX <> vData)
   mvarQTDMAX = vData
End Property
Public Property Get QTDMAX() As Double
   QTDMAX = mvarQTDMAX
End Property
Public Property Let PESOLIQ(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPESOLIQ <> vData)
   mvarPESOLIQ = vData
End Property
Public Property Get PESOLIQ() As Double
   PESOLIQ = mvarPESOLIQ
End Property
Public Property Let QTDRP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDRP <> vData)
   mvarQTDRP = vData
End Property
Public Property Get QTDRP() As Double
   QTDRP = mvarQTDRP
End Property
Public Property Let PESOBRUTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPESOBRUTO <> vData)
   mvarPESOBRUTO = vData
End Property
Public Property Get PESOBRUTO() As Double
   PESOBRUTO = mvarPESOBRUTO
End Property
Public Property Let COMPRIMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarCOMPRIMENTO <> vData)
   mvarCOMPRIMENTO = vData
End Property
Public Property Get COMPRIMENTO() As Double
   COMPRIMENTO = mvarCOMPRIMENTO
End Property
Public Property Let LARGURA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarLARGURA <> vData)
   mvarLARGURA = vData
End Property
Public Property Get LARGURA() As Double
   LARGURA = mvarLARGURA
End Property
Public Property Let ESPESSURA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarESPESSURA <> vData)
   mvarESPESSURA = vData
End Property
Public Property Get ESPESSURA() As Double
   ESPESSURA = mvarESPESSURA
End Property
Public Property Let EMATPRIMA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEMATPRIMA <> vData)
   mvarEMATPRIMA = vData
End Property
Public Property Get EMATPRIMA() As String
   EMATPRIMA = mvarEMATPRIMA
End Property
Public Property Let EPRODACABADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEPRODACABADO <> vData)
   mvarEPRODACABADO = vData
End Property
Public Property Get EPRODACABADO() As String
   EPRODACABADO = mvarEPRODACABADO
End Property
Public Property Let EICMS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEICMS <> vData)
   mvarEICMS = vData
End Property
Public Property Get EICMS() As String
   EICMS = mvarEICMS
End Property
Public Property Let SITPROD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITPROD <> vData)
   mvarSITPROD = vData
End Property
Public Property Get SITPROD() As String
   SITPROD = mvarSITPROD
End Property
Public Property Let IDFAM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFAM <> vData)
   mvarIDFAM = vData
End Property
Public Property Get IDFAM() As Double
   IDFAM = mvarIDFAM
End Property
Public Property Let INFORMATICA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarINFORMATICA <> vData)
   mvarINFORMATICA = vData
End Property
Public Property Get INFORMATICA() As String
   INFORMATICA = mvarINFORMATICA
End Property
Public Property Let ATIVOFIXO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVOFIXO <> vData)
   mvarATIVOFIXO = vData
End Property
Public Property Get ATIVOFIXO() As String
   ATIVOFIXO = mvarATIVOFIXO
End Property
Public Property Let QUALIFICADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarQUALIFICADO <> vData)
   mvarQUALIFICADO = vData
End Property
Public Property Get QUALIFICADO() As String
   QUALIFICADO = mvarQUALIFICADO
End Property
Public Property Let IDDISCIPLINA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDISCIPLINA <> vData)
   mvarIDDISCIPLINA = vData
End Property
Public Property Get IDDISCIPLINA() As Double
   IDDISCIPLINA = mvarIDDISCIPLINA
End Property
Public Property Let NMPRODIng(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMPRODIng <> vData)
   mvarNMPRODIng = vData
End Property
Public Property Get NMPRODIng() As String
   NMPRODIng = mvarNMPRODIng
End Property
Public Property Let DSCPRODIng(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCPRODIng <> vData)
   mvarDSCPRODIng = vData
End Property
Public Property Get DSCPRODIng() As String
   DSCPRODIng = mvarDSCPRODIng
End Property
Public Property Let TXCONVUNID(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTXCONVUNID <> vData)
   mvarTXCONVUNID = vData
End Property
Public Property Get TXCONVUNID() As String
   TXCONVUNID = mvarTXCONVUNID
End Property
Public Property Let DIAMETRO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarDIAMETRO <> vData)
   mvarDIAMETRO = vData
End Property
Public Property Get DIAMETRO() As Double
   DIAMETRO = mvarDIAMETRO
End Property
Public Property Let OBSSMS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBSSMS <> vData)
   mvarOBSSMS = vData
End Property
Public Property Get OBSSMS() As String
   OBSSMS = mvarOBSSMS
End Property
Public Property Let IDGRPSMS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDGRPSMS <> vData)
   mvarIDGRPSMS = vData
End Property
Public Property Get IDGRPSMS() As Double
   IDGRPSMS = mvarIDGRPSMS
End Property
Public Property Let TAG(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTAG <> vData)
   mvarTAG = vData
End Property
Public Property Get TAG() As String
   TAG = mvarTAG
End Property
Public Property Let IDCLASSE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLASSE <> vData)
   mvarIDCLASSE = vData
End Property
Public Property Get IDCLASSE() As Double
   IDCLASSE = mvarIDCLASSE
End Property
Public Property Let IDFABRICANTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFABRICANTE <> vData)
   mvarIDFABRICANTE = vData
End Property
Public Property Get IDFABRICANTE() As Double
   IDFABRICANTE = mvarIDFABRICANTE
End Property
Public Property Let ALTURA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarALTURA <> vData)
   mvarALTURA = vData
End Property
Public Property Get ALTURA() As Double
   ALTURA = mvarALTURA
End Property
Public Property Let TIPOACO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPOACO <> vData)
   mvarTIPOACO = vData
End Property
Public Property Get TIPOACO() As String
   TIPOACO = mvarTIPOACO
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
 
   Sql = "Insert Into PRODUTO (IDPROD, CODPROD, NMPROD, DSCPROD" & VbNewLine 
   Sql = Sql & ", SIGLAUNID, EGRUPO, IDPAI, IDUSU, DTCADASTRO" & VbNewLine 
   Sql = Sql & ", VLESTIMADO, IDGRPFOR, IDDESP, MSKCOD, QTDCOMPRA" & VbNewLine 
   Sql = Sql & ", UNIDCOMPRA, CODBARRA, DTDESATIVO, VLULTCOMPRA, VLMEDIO" & VbNewLine 
   Sql = Sql & ", QTDESTOQUE, QTDMIN, QTDMAX, PESOLIQ, QTDRP" & VbNewLine 
   Sql = Sql & ", PESOBRUTO, COMPRIMENTO, LARGURA, ESPESSURA, EMATPRIMA" & VbNewLine 
   Sql = Sql & ", EPRODACABADO, EICMS, SITPROD, IDFAM, INFORMATICA" & VbNewLine 
   Sql = Sql & ", ATIVOFIXO, QUALIFICADO, IDDISCIPLINA, NMPRODIng, DSCPRODIng" & VbNewLine 
   Sql = Sql & ", TXCONVUNID, DIAMETRO, OBSSMS, IDGRPSMS, TAG" & VbNewLine 
   Sql = Sql & ", IDCLASSE, IDFABRICANTE, ALTURA, TIPOACO) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  "isNull(Max(IDPROD),0)+1" & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSIGLAUNID) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEGRUPO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLESTIMADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDGRPFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDDESP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMSKCOD) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDCOMPRA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUNIDCOMPRA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODBARRA) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTDESATIVO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLULTCOMPRA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLMEDIO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDESTOQUE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDMIN) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDMAX) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPESOLIQ) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDRP) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPESOBRUTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarCOMPRIMENTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarLARGURA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarESPESSURA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMATPRIMA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEPRODACABADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEICMS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSITPROD) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFAM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarINFORMATICA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarATIVOFIXO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarQUALIFICADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDDISCIPLINA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMPRODIng) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCPRODIng) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTXCONVUNID) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarDIAMETRO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBSSMS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDGRPSMS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTAG) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCLASSE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFABRICANTE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarALTURA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPOACO) & VbNewLine 
   Sql = Sql & " From PRODUTO" & VbNewLine 
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPROD) As String
   Dim Sql As String
 
   Sql = "Delete From PRODUTO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDPROD) Then Sql = Sql & " IDPROD = " & SqlNum(Cstr(Ch_IDPROD)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update PRODUTO Set " & VbNewLine 
   Sql = Sql & " IDPROD       = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " , CODPROD      = " & SqlStr(mvarCODPROD) & VbNewLine 
   Sql = Sql & " , NMPROD       = " & SqlStr(mvarNMPROD) & VbNewLine 
   Sql = Sql & " , DSCPROD      = " & SqlStr(mvarDSCPROD) & VbNewLine 
   Sql = Sql & " , SIGLAUNID    = " & SqlStr(mvarSIGLAUNID) & VbNewLine 
   Sql = Sql & " , EGRUPO       = " & SqlStr(mvarEGRUPO) & VbNewLine 
   Sql = Sql & " , IDPAI        = " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & " , IDUSU        = " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & " , DTCADASTRO   = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VLESTIMADO   = " & SqlNum(mvarVLESTIMADO) & VbNewLine 
   Sql = Sql & " , IDGRPFOR     = " & SqlNum(mvarIDGRPFOR) & VbNewLine 
   Sql = Sql & " , IDDESP       = " & SqlStr(mvarIDDESP) & VbNewLine 
   Sql = Sql & " , MSKCOD       = " & SqlStr(mvarMSKCOD) & VbNewLine 
   Sql = Sql & " , QTDCOMPRA    = " & SqlNum(mvarQTDCOMPRA) & VbNewLine 
   Sql = Sql & " , UNIDCOMPRA   = " & SqlStr(mvarUNIDCOMPRA) & VbNewLine 
   Sql = Sql & " , CODBARRA     = " & SqlStr(mvarCODBARRA) & VbNewLine 
   Sql = Sql & " , DTDESATIVO   = " & SqlDate(mvarDTDESATIVO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VLULTCOMPRA  = " & SqlNum(mvarVLULTCOMPRA) & VbNewLine 
   Sql = Sql & " , VLMEDIO      = " & SqlNum(mvarVLMEDIO) & VbNewLine 
   Sql = Sql & " , QTDESTOQUE   = " & SqlNum(mvarQTDESTOQUE) & VbNewLine 
   Sql = Sql & " , QTDMIN       = " & SqlNum(mvarQTDMIN) & VbNewLine 
   Sql = Sql & " , QTDMAX       = " & SqlNum(mvarQTDMAX) & VbNewLine 
   Sql = Sql & " , PESOLIQ      = " & SqlNum(mvarPESOLIQ) & VbNewLine 
   Sql = Sql & " , QTDRP        = " & SqlNum(mvarQTDRP) & VbNewLine 
   Sql = Sql & " , PESOBRUTO    = " & SqlNum(mvarPESOBRUTO) & VbNewLine 
   Sql = Sql & " , COMPRIMENTO  = " & SqlNum(mvarCOMPRIMENTO) & VbNewLine 
   Sql = Sql & " , LARGURA      = " & SqlNum(mvarLARGURA) & VbNewLine 
   Sql = Sql & " , ESPESSURA    = " & SqlNum(mvarESPESSURA) & VbNewLine 
   Sql = Sql & " , EMATPRIMA    = " & SqlStr(mvarEMATPRIMA) & VbNewLine 
   Sql = Sql & " , EPRODACABADO = " & SqlStr(mvarEPRODACABADO) & VbNewLine 
   Sql = Sql & " , EICMS        = " & SqlStr(mvarEICMS) & VbNewLine 
   Sql = Sql & " , SITPROD      = " & SqlStr(mvarSITPROD) & VbNewLine 
   Sql = Sql & " , IDFAM        = " & SqlNum(mvarIDFAM) & VbNewLine 
   Sql = Sql & " , INFORMATICA  = " & SqlStr(mvarINFORMATICA) & VbNewLine 
   Sql = Sql & " , ATIVOFIXO    = " & SqlStr(mvarATIVOFIXO) & VbNewLine 
   Sql = Sql & " , QUALIFICADO  = " & SqlStr(mvarQUALIFICADO) & VbNewLine 
   Sql = Sql & " , IDDISCIPLINA = " & SqlNum(mvarIDDISCIPLINA) & VbNewLine 
   Sql = Sql & " , NMPRODIng    = " & SqlStr(mvarNMPRODIng) & VbNewLine 
   Sql = Sql & " , DSCPRODIng   = " & SqlStr(mvarDSCPRODIng) & VbNewLine 
   Sql = Sql & " , TXCONVUNID   = " & SqlStr(mvarTXCONVUNID) & VbNewLine 
   Sql = Sql & " , DIAMETRO     = " & SqlNum(mvarDIAMETRO) & VbNewLine 
   Sql = Sql & " , OBSSMS       = " & SqlStr(mvarOBSSMS) & VbNewLine 
   Sql = Sql & " , IDGRPSMS     = " & SqlNum(mvarIDGRPSMS) & VbNewLine 
   Sql = Sql & " , TAG          = " & SqlStr(mvarTAG) & VbNewLine 
   Sql = Sql & " , IDCLASSE     = " & SqlNum(mvarIDCLASSE) & VbNewLine 
   Sql = Sql & " , IDFABRICANTE = " & SqlNum(mvarIDFABRICANTE) & VbNewLine 
   Sql = Sql & " , ALTURA       = " & SqlNum(mvarALTURA) & VbNewLine 
   Sql = Sql & " , TIPOACO      = " & SqlStr(mvarTIPOACO) & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPROD         = " & SqlNum(mvarIDPROD) & VbNewLine 
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
Public Property Get QrySelect( Optional Ch_IDPROD, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Select IDPROD, CODPROD, NMPROD, DSCPROD" & VbNewLine 
   Sql = Sql & ", SIGLAUNID, EGRUPO, IDPAI, IDUSU, DTCADASTRO" & VbNewLine 
   Sql = Sql & ", VLESTIMADO, IDGRPFOR, IDDESP, MSKCOD, QTDCOMPRA" & VbNewLine 
   Sql = Sql & ", UNIDCOMPRA, CODBARRA, DTDESATIVO, VLULTCOMPRA, VLMEDIO" & VbNewLine 
   Sql = Sql & ", QTDESTOQUE, QTDMIN, QTDMAX, PESOLIQ, QTDRP" & VbNewLine 
   Sql = Sql & ", PESOBRUTO, COMPRIMENTO, LARGURA, ESPESSURA, EMATPRIMA" & VbNewLine 
   Sql = Sql & ", EPRODACABADO, EICMS, SITPROD, IDFAM, INFORMATICA" & VbNewLine 
   Sql = Sql & ", ATIVOFIXO, QUALIFICADO, IDDISCIPLINA, NMPRODIng, DSCPRODIng" & VbNewLine 
   Sql = Sql & ", TXCONVUNID, DIAMETRO, OBSSMS, IDGRPSMS, TAG" & VbNewLine 
   Sql = Sql & ", IDCLASSE, IDFABRICANTE, ALTURA, TIPOACO" & VbNewLine 
   Sql = Sql &" From PRODUTO" & VbNewLine 
   If IsMissing(Ch_WHERE) Then
      Sql = Sql & " Where " & VbNewLine 
      If IsMissing(Ch_IDPROD)  Then 
         Sql = Sql & " IDPROD = " & SqlNum(mvarIDPROD) & VbNewLine 
      Else 
         If Not isMissing(Ch_IDPROD) Then Sql = Sql & " IDPROD = " & SqlNum(Cstr(Ch_IDPROD)) & " AND " & VbNewLine 
         Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
      End If 
   Else
      If Trim(Ch_WHERE) <> "" Then
         Sql = Sql & " Where " & Ch_WHERE & vbNewLine
      End If
   End If

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar( Optional Ch_IDPROD, Optional Ch_WHERE) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPROD, Ch_WHERE)
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
      mvarIDPROD = XVal(!IDPROD & "")
      mvarCODPROD = !CODPROD & ""
      mvarNMPROD = !NMPROD & ""
      mvarDSCPROD = !DSCPROD & ""
      mvarSIGLAUNID = !SIGLAUNID & ""
      mvarEGRUPO = !EGRUPO & ""
      mvarIDPAI = XVal(!IDPAI & "")
      mvarIDUSU = !IDUSU & ""
      mvarDTCADASTRO = FormatDate(!DTCADASTRO & "", True)
      mvarVLESTIMADO = XVal(!VLESTIMADO & "")
      mvarIDGRPFOR = XVal(!IDGRPFOR & "")
      mvarIDDESP = !IDDESP & ""
      mvarMSKCOD = !MSKCOD & ""
      mvarQTDCOMPRA = XVal(!QTDCOMPRA & "")
      mvarUNIDCOMPRA = !UNIDCOMPRA & ""
      mvarCODBARRA = !CODBARRA & ""
      mvarDTDESATIVO = FormatDate(!DTDESATIVO & "", True)
      mvarVLULTCOMPRA = XVal(!VLULTCOMPRA & "")
      mvarVLMEDIO = XVal(!VLMEDIO & "")
      mvarQTDESTOQUE = XVal(!QTDESTOQUE & "")
      mvarQTDMIN = XVal(!QTDMIN & "")
      mvarQTDMAX = XVal(!QTDMAX & "")
      mvarPESOLIQ = XVal(!PESOLIQ & "")
      mvarQTDRP = XVal(!QTDRP & "")
      mvarPESOBRUTO = XVal(!PESOBRUTO & "")
      mvarCOMPRIMENTO = XVal(!COMPRIMENTO & "")
      mvarLARGURA = XVal(!LARGURA & "")
      mvarESPESSURA = XVal(!ESPESSURA & "")
      mvarEMATPRIMA = !EMATPRIMA & ""
      mvarEPRODACABADO = !EPRODACABADO & ""
      mvarEICMS = !EICMS & ""
      mvarSITPROD = !SITPROD & ""
      mvarIDFAM = XVal(!IDFAM & "")
      mvarINFORMATICA = !INFORMATICA & ""
      mvarATIVOFIXO = !ATIVOFIXO & ""
      mvarQUALIFICADO = !QUALIFICADO & ""
      mvarIDDISCIPLINA = XVal(!IDDISCIPLINA & "")
      mvarNMPRODIng = !NMPRODIng & ""
      mvarDSCPRODIng = !DSCPRODIng & ""
      mvarTXCONVUNID = !TXCONVUNID & ""
      mvarDIAMETRO = XVal(!DIAMETRO & "")
      mvarOBSSMS = !OBSSMS & ""
      mvarIDGRPSMS = XVal(!IDGRPSMS & "")
      mvarTAG = !TAG & ""
      mvarIDCLASSE = XVal(!IDCLASSE & "")
      mvarIDFABRICANTE = XVal(!IDFABRICANTE & "")
      mvarALTURA = XVal(!ALTURA & "")
      mvarTIPOACO = !TIPOACO & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPROD), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDPROD As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update PRODUTO Set "
   Sql = Sql & " IDPROD = " & SqlNum(Ch_IDPROD)
   Sql = Sql & " Where 
   Sql = Sql & " IDPROD = " & SqlNum(mvarIDPROD)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
