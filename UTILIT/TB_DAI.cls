VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TB_DAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private mvardBase As Object
Private mvarEXISTE As Integer

Private mvarQryInsert As String
Private mvarQryUpDate As String
Private mvarQryDelete As String

Private mvarDAI_DV As String
Private mvarDAI_NUM As String
Private mvarDAI_SEQ As String
Private mvarBAN_COD As String
Private mvarDAI_COMPL As String
Private mvarCLI_NOM As String
Private mvarCLI_NUM As String
Private mvarCLI_PF_PJ As String
Private mvarCLI_TIP_COB As String
Private mvarCRG_NUM As String
Private mvarDAI_DAT_EMI As String
Private mvarDAI_DAT_ENTRG As String
Private mvarDAI_DAT_ISEN As String
Private mvarDAI_DAT_LIB As String
Private mvarDAI_DAT_PGTO As String
Private mvarDAI_DAT_VENCTO As String
Private mvarDAI_HOR As String
Private mvarDAI_IDT_REIMP As String
Private mvarDAI_MOT_ALT As String
Private mvarDAI_MOT_VENC As String
Private mvarDAI_NUM_MAN As String
Private mvarDAI_NUM_PERI As String
Private mvarDAI_PES_BRU As String
Private mvarDAI_PES_LIQ As String
Private mvarDAI_QTD_VOL As String
Private mvarDAI_SIT As String
Private mvarDAI_TIP_TRAT As String
Private mvarDAI_TOT_CIF As String
Private mvarDAI_TOT_FOB As String
Private mvarDAI_TOT_FRT As String
Private mvarDAI_TOT_SGR As String
Private mvarDAI_VLR_ABAT As String
Private mvarDAI_VLR_ATAERO As String
Private mvarDAI_VLR_PGTO As String
Private mvarDAI_VLR_TARIF As String
Private mvarDAT_REC_CRG As String
Private mvarDET_COD As String
Private mvarDFAT_NUM As String
Private mvarDLIB_DAT As String
Private mvarDLIB_NUM As String
Private mvarDLIB_TIP As String
Private mvarDSPCH_NUM As String
Private mvarENQ_COD As String
Private mvarESP_MIN_PRE As String
Private mvarESP_NUM_DOC As String
Private mvarHCRG_NUM As String
Private mvarIMP_COD As String
Private mvarNUM_TERMO As String
Private mvarPESS_COD As String
Private mvarTXA_USD_REAL As String
Private mvarDAI_VLR_RET As String
Private mvarDAI_VLR_COR As String
Private mvarDAI_VLR_MUL As String
Private mvarDAI_VLR_JUR As String
Private mvarDAI_MOT_BAIXA As String
Private mvarDAI_DAT_APR As String
Private mvarDAI_VLR_APR As String
Private mvarDAI_IDT_TABN As String
Private mvarAER_COD As String
Public Property Let AER_COD(ByVal vData As String)
   mvarAER_COD = vData
End Property
Public Property Get AER_COD() As String
   AER_COD = mvarAER_COD
End Property
Public Property Let DAI_DV(ByVal vData As String)
   mvarDAI_DV = vData
End Property
Public Property Get DAI_DV() As String
   DAI_DV = mvarDAI_DV
End Property
Public Property Let DAI_NUM(ByVal vData As String)
   mvarDAI_NUM = vData
End Property
Public Property Get DAI_NUM() As String
   DAI_NUM = mvarDAI_NUM
End Property
Public Property Let DAI_SEQ(ByVal vData As String)
   mvarDAI_SEQ = vData
End Property
Public Property Get DAI_SEQ() As String
   DAI_SEQ = mvarDAI_SEQ
End Property
Public Property Let BAN_COD(ByVal vData As String)
   mvarBAN_COD = vData
End Property
Public Property Get BAN_COD() As String
   BAN_COD = mvarBAN_COD
End Property
Public Property Let DAI_COMPL(ByVal vData As String)
   mvarDAI_COMPL = vData
End Property
Public Property Get DAI_COMPL() As String
   DAI_COMPL = mvarDAI_COMPL
End Property
Public Property Let CLI_NOM(ByVal vData As String)
   mvarCLI_NOM = vData
End Property
Public Property Get CLI_NOM() As String
   CLI_NOM = mvarCLI_NOM
End Property
Public Property Let CLI_NUM(ByVal vData As String)
   mvarCLI_NUM = vData
End Property
Public Property Get CLI_NUM() As String
   CLI_NUM = mvarCLI_NUM
End Property
Public Property Let CLI_PF_PJ(ByVal vData As String)
   mvarCLI_PF_PJ = vData
End Property
Public Property Get CLI_PF_PJ() As String
   CLI_PF_PJ = mvarCLI_PF_PJ
End Property
Public Property Let CLI_TIP_COB(ByVal vData As String)
   mvarCLI_TIP_COB = vData
End Property
Public Property Get CLI_TIP_COB() As String
   CLI_TIP_COB = mvarCLI_TIP_COB
End Property
Public Property Let CRG_NUM(ByVal vData As String)
   mvarCRG_NUM = vData
End Property
Public Property Get CRG_NUM() As String
   CRG_NUM = mvarCRG_NUM
End Property
Public Property Let DAI_DAT_EMI(ByVal vData As String)
   mvarDAI_DAT_EMI = vData
End Property
Public Property Get DAI_DAT_EMI() As String
   DAI_DAT_EMI = mvarDAI_DAT_EMI
End Property
Public Property Let DAI_DAT_ENTRG(ByVal vData As String)
   mvarDAI_DAT_ENTRG = vData
End Property
Public Property Get DAI_DAT_ENTRG() As String
   DAI_DAT_ENTRG = mvarDAI_DAT_ENTRG
End Property
Public Property Let DAI_DAT_ISEN(ByVal vData As String)
   mvarDAI_DAT_ISEN = vData
End Property
Public Property Get DAI_DAT_ISEN() As String
   DAI_DAT_ISEN = mvarDAI_DAT_ISEN
End Property
Public Property Let DAI_DAT_LIB(ByVal vData As String)
   mvarDAI_DAT_LIB = vData
End Property
Public Property Get DAI_DAT_LIB() As String
   DAI_DAT_LIB = mvarDAI_DAT_LIB
End Property
Public Property Let DAI_DAT_PGTO(ByVal vData As String)
   mvarDAI_DAT_PGTO = vData
End Property
Public Property Get DAI_DAT_PGTO() As String
   DAI_DAT_PGTO = mvarDAI_DAT_PGTO
End Property
Public Property Let DAI_DAT_VENCTO(ByVal vData As String)
   mvarDAI_DAT_VENCTO = vData
End Property
Public Property Get DAI_DAT_VENCTO() As String
   DAI_DAT_VENCTO = mvarDAI_DAT_VENCTO
End Property
Public Property Let DAI_HOR(ByVal vData As String)
   mvarDAI_HOR = vData
End Property
Public Property Get DAI_HOR() As String
   DAI_HOR = mvarDAI_HOR
End Property
Public Property Let DAI_IDT_REIMP(ByVal vData As String)
   mvarDAI_IDT_REIMP = vData
End Property
Public Property Get DAI_IDT_REIMP() As String
   DAI_IDT_REIMP = mvarDAI_IDT_REIMP
End Property
Public Property Let DAI_MOT_ALT(ByVal vData As String)
   mvarDAI_MOT_ALT = vData
End Property
Public Property Get DAI_MOT_ALT() As String
   DAI_MOT_ALT = mvarDAI_MOT_ALT
End Property
Public Property Let DAI_MOT_VENC(ByVal vData As String)
   mvarDAI_MOT_VENC = vData
End Property
Public Property Get DAI_MOT_VENC() As String
   DAI_MOT_VENC = mvarDAI_MOT_VENC
End Property
Public Property Let DAI_NUM_MAN(ByVal vData As String)
   mvarDAI_NUM_MAN = vData
End Property
Public Property Get DAI_NUM_MAN() As String
   DAI_NUM_MAN = mvarDAI_NUM_MAN
End Property
Public Property Let DAI_NUM_PERI(ByVal vData As String)
   mvarDAI_NUM_PERI = vData
End Property
Public Property Get DAI_NUM_PERI() As String
   DAI_NUM_PERI = mvarDAI_NUM_PERI
End Property
Public Property Let DAI_PES_BRU(ByVal vData As String)
   mvarDAI_PES_BRU = vData
End Property
Public Property Get DAI_PES_BRU() As String
   DAI_PES_BRU = mvarDAI_PES_BRU
End Property
Public Property Let DAI_PES_LIQ(ByVal vData As String)
   mvarDAI_PES_LIQ = vData
End Property
Public Property Get DAI_PES_LIQ() As String
   DAI_PES_LIQ = mvarDAI_PES_LIQ
End Property
Public Property Let DAI_QTD_VOL(ByVal vData As String)
   mvarDAI_QTD_VOL = vData
End Property
Public Property Get DAI_QTD_VOL() As String
   DAI_QTD_VOL = mvarDAI_QTD_VOL
End Property
Public Property Let DAI_SIT(ByVal vData As String)
   mvarDAI_SIT = vData
End Property
Public Property Get DAI_SIT() As String
   DAI_SIT = mvarDAI_SIT
End Property
Public Property Let DAI_TIP_TRAT(ByVal vData As String)
   mvarDAI_TIP_TRAT = vData
End Property
Public Property Get DAI_TIP_TRAT() As String
   DAI_TIP_TRAT = mvarDAI_TIP_TRAT
End Property
Public Property Let DAI_TOT_CIF(ByVal vData As String)
   mvarDAI_TOT_CIF = vData
End Property
Public Property Get DAI_TOT_CIF() As String
   DAI_TOT_CIF = mvarDAI_TOT_CIF
End Property
Public Property Let DAI_TOT_FOB(ByVal vData As String)
   mvarDAI_TOT_FOB = vData
End Property
Public Property Get DAI_TOT_FOB() As String
   DAI_TOT_FOB = mvarDAI_TOT_FOB
End Property
Public Property Let DAI_TOT_FRT(ByVal vData As String)
   mvarDAI_TOT_FRT = vData
End Property
Public Property Get DAI_TOT_FRT() As String
   DAI_TOT_FRT = mvarDAI_TOT_FRT
End Property
Public Property Let DAI_TOT_SGR(ByVal vData As String)
   mvarDAI_TOT_SGR = vData
End Property
Public Property Get DAI_TOT_SGR() As String
   DAI_TOT_SGR = mvarDAI_TOT_SGR
End Property
Public Property Let DAI_VLR_ABAT(ByVal vData As String)
   mvarDAI_VLR_ABAT = vData
End Property
Public Property Get DAI_VLR_ABAT() As String
   DAI_VLR_ABAT = mvarDAI_VLR_ABAT
End Property
Public Property Let DAI_VLR_ATAERO(ByVal vData As String)
   mvarDAI_VLR_ATAERO = vData
End Property
Public Property Get DAI_VLR_ATAERO() As String
   DAI_VLR_ATAERO = mvarDAI_VLR_ATAERO
End Property
Public Property Let DAI_VLR_PGTO(ByVal vData As String)
   mvarDAI_VLR_PGTO = vData
End Property
Public Property Get DAI_VLR_PGTO() As String
   DAI_VLR_PGTO = mvarDAI_VLR_PGTO
End Property
Public Property Let DAI_VLR_TARIF(ByVal vData As String)
   mvarDAI_VLR_TARIF = vData
End Property
Public Property Get DAI_VLR_TARIF() As String
   DAI_VLR_TARIF = mvarDAI_VLR_TARIF
End Property
Public Property Let DAT_REC_CRG(ByVal vData As String)
   mvarDAT_REC_CRG = vData
End Property
Public Property Get DAT_REC_CRG() As String
   DAT_REC_CRG = mvarDAT_REC_CRG
End Property
Public Property Let DET_COD(ByVal vData As String)
   mvarDET_COD = vData
End Property
Public Property Get DET_COD() As String
   DET_COD = mvarDET_COD
End Property
Public Property Let DFAT_NUM(ByVal vData As String)
   mvarDFAT_NUM = vData
End Property
Public Property Get DFAT_NUM() As String
   DFAT_NUM = mvarDFAT_NUM
End Property
Public Property Let DLIB_DAT(ByVal vData As String)
   mvarDLIB_DAT = vData
End Property
Public Property Get DLIB_DAT() As String
   DLIB_DAT = mvarDLIB_DAT
End Property
Public Property Let DLIB_NUM(ByVal vData As String)
   mvarDLIB_NUM = vData
End Property
Public Property Get DLIB_NUM() As String
   DLIB_NUM = mvarDLIB_NUM
End Property
Public Property Let DLIB_TIP(ByVal vData As String)
   mvarDLIB_TIP = vData
End Property
Public Property Get DLIB_TIP() As String
   DLIB_TIP = mvarDLIB_TIP
End Property
Public Property Let DSPCH_NUM(ByVal vData As String)
   mvarDSPCH_NUM = vData
End Property
Public Property Get DSPCH_NUM() As String
   DSPCH_NUM = mvarDSPCH_NUM
End Property
Public Property Let ENQ_COD(ByVal vData As String)
   mvarENQ_COD = vData
End Property
Public Property Get ENQ_COD() As String
   ENQ_COD = mvarENQ_COD
End Property
Public Property Let ESP_MIN_PRE(ByVal vData As String)
   mvarESP_MIN_PRE = vData
End Property
Public Property Get ESP_MIN_PRE() As String
   ESP_MIN_PRE = mvarESP_MIN_PRE
End Property
Public Property Let ESP_NUM_DOC(ByVal vData As String)
   mvarESP_NUM_DOC = vData
End Property
Public Property Get ESP_NUM_DOC() As String
   ESP_NUM_DOC = mvarESP_NUM_DOC
End Property
Public Property Let HCRG_NUM(ByVal vData As String)
   mvarHCRG_NUM = vData
End Property
Public Property Get HCRG_NUM() As String
   HCRG_NUM = mvarHCRG_NUM
End Property
Public Property Let IMP_COD(ByVal vData As String)
   mvarIMP_COD = vData
End Property
Public Property Get IMP_COD() As String
   IMP_COD = mvarIMP_COD
End Property
Public Property Let NUM_TERMO(ByVal vData As String)
   mvarNUM_TERMO = vData
End Property
Public Property Get NUM_TERMO() As String
   NUM_TERMO = mvarNUM_TERMO
End Property
Public Property Let PESS_COD(ByVal vData As String)
   mvarPESS_COD = vData
End Property
Public Property Get PESS_COD() As String
   PESS_COD = mvarPESS_COD
End Property
Public Property Let TXA_USD_REAL(ByVal vData As String)
   mvarTXA_USD_REAL = vData
End Property
Public Property Get TXA_USD_REAL() As String
   TXA_USD_REAL = mvarTXA_USD_REAL
End Property
Public Property Let DAI_VLR_RET(ByVal vData As String)
   mvarDAI_VLR_RET = vData
End Property
Public Property Get DAI_VLR_RET() As String
   DAI_VLR_RET = mvarDAI_VLR_RET
End Property
Public Property Let DAI_VLR_COR(ByVal vData As String)
   mvarDAI_VLR_COR = vData
End Property
Public Property Get DAI_VLR_COR() As String
   DAI_VLR_COR = mvarDAI_VLR_COR
End Property
Public Property Let DAI_VLR_MUL(ByVal vData As String)
   mvarDAI_VLR_MUL = vData
End Property
Public Property Get DAI_VLR_MUL() As String
   DAI_VLR_MUL = mvarDAI_VLR_MUL
End Property
Public Property Let DAI_VLR_JUR(ByVal vData As String)
   mvarDAI_VLR_JUR = vData
End Property
Public Property Get DAI_VLR_JUR() As String
   DAI_VLR_JUR = mvarDAI_VLR_JUR
End Property
Public Property Let DAI_MOT_BAIXA(ByVal vData As String)
   mvarDAI_MOT_BAIXA = vData
End Property
Public Property Get DAI_MOT_BAIXA() As String
   DAI_MOT_BAIXA = mvarDAI_MOT_BAIXA
End Property
Public Property Let DAI_DAT_APR(ByVal vData As String)
   mvarDAI_DAT_APR = vData
End Property
Public Property Get DAI_DAT_APR() As String
   DAI_DAT_APR = mvarDAI_DAT_APR
End Property
Public Property Let DAI_VLR_APR(ByVal vData As String)
   mvarDAI_VLR_APR = vData
End Property
Public Property Get DAI_VLR_APR() As String
   DAI_VLR_APR = mvarDAI_VLR_APR
End Property
Public Property Let DAI_IDT_TABN(ByVal vData As String)
   mvarDAI_IDT_TABN = vData
End Property
Public Property Get DAI_IDT_TABN() As String
   DAI_IDT_TABN = mvarDAI_IDT_TABN
End Property
Public Property Get EXISTE() As Integer
   EXISTE = mvarEXISTE
End Property
Public Property Let dBase(ByVal vData As Object)
   Set mvardBase = vData
End Property
Public Property Get dBase() As Object
   Set dBase = mvardBase
End Property
Public Property Get QryInsert() As String
   Dim Sql$
   Sql = "insert into DAI (DAI_DV,DAI_NUM,DAI_SEQ,BAN_COD,DAI_COMPL,CLI_NOM,CLI_NUM,CLI_PF_PJ,CLI_TIP_COB,CRG_NUM,DAI_DAT_EMI,DAI_DAT_ENTRG,DAI_DAT_ISEN,DAI_DAT_LIB,DAI_DAT_PGTO,DAI_DAT_VENCTO,DAI_HOR,DAI_IDT_REIMP,DAI_MOT_ALT,DAI_MOT_VENC,DAI_NUM_MAN,DAI_NUM_PERI,DAI_PES_BRU,DAI_PES_LIQ,DAI_QTD_VOL,DAI_SIT,DAI_TIP_TRAT,DAI_TOT_CIF,DAI_TOT_FOB,DAI_TOT_FRT,DAI_TOT_SGR,DAI_VLR_ABAT,DAI_VLR_ATAERO,DAI_VLR_PGTO,DAI_VLR_TARIF,DAT_REC_CRG,DET_COD,DFAT_NUM,DLIB_DAT,DLIB_NUM,DLIB_TIP,DSPCH_NUM,ENQ_COD,ESP_MIN_PRE,ESP_NUM_DOC,HCRG_NUM,IMP_COD,NUM_TERMO,PESS_COD,TXA_USD_REAL,DAI_VLR_RET,DAI_VLR_COR,DAI_VLR_MUL,DAI_VLR_JUR,DAI_MOT_BAIXA,DAI_DAT_APR,DAI_VLR_APR,DAI_IDT_TABN) "
   Sql = Sql & " Values "
   Sql = Sql & "("
   Sql = Sql & Aspas(mvarDAI_DV)
   Sql = Sql & ", " & Aspas(mvarDAI_NUM)
   Sql = Sql & ", " & Aspas(mvarDAI_SEQ)
   Sql = Sql & ", " & Aspas(mvarBAN_COD)
   Sql = Sql & ", " & Aspas(mvarDAI_COMPL)
   Sql = Sql & ", " & Aspas(mvarCLI_NOM)
   Sql = Sql & ", " & Aspas(mvarCLI_NUM)
   Sql = Sql & ", " & Aspas(mvarCLI_PF_PJ)
   Sql = Sql & ", " & Aspas(mvarCLI_TIP_COB)
   Sql = Sql & ", " & Aspas(mvarCRG_NUM)
   Sql = Sql & ", " & SqlDate(mvarDAI_DAT_EMI)
   Sql = Sql & ", " & SqlDate(mvarDAI_DAT_ENTRG)
   Sql = Sql & ", " & SqlDate(mvarDAI_DAT_ISEN)
   Sql = Sql & ", " & SqlDate(mvarDAI_DAT_LIB)
   Sql = Sql & ", " & SqlDate(mvarDAI_DAT_PGTO)
   Sql = Sql & ", " & SqlDate(mvarDAI_DAT_VENCTO)
   Sql = Sql & ", " & Aspas(mvarDAI_HOR)
   Sql = Sql & ", " & Aspas(mvarDAI_IDT_REIMP)
   Sql = Sql & ", " & Aspas(mvarDAI_MOT_ALT)
   Sql = Sql & ", " & Aspas(mvarDAI_MOT_VENC)
   Sql = Sql & ", " & Aspas(mvarDAI_NUM_MAN)
   Sql = Sql & ", " & mvarDAI_NUM_PERI
   Sql = Sql & ", " & mvarDAI_PES_BRU
   Sql = Sql & ", " & mvarDAI_PES_LIQ
   Sql = Sql & ", " & mvarDAI_QTD_VOL
   Sql = Sql & ", " & Aspas(mvarDAI_SIT)
   Sql = Sql & ", " & Aspas(mvarDAI_TIP_TRAT)
   Sql = Sql & ", " & mvarDAI_TOT_CIF
   Sql = Sql & ", " & mvarDAI_TOT_FOB
   Sql = Sql & ", " & mvarDAI_TOT_FRT
   Sql = Sql & ", " & mvarDAI_TOT_SGR
   Sql = Sql & ", " & mvarDAI_VLR_ABAT
   Sql = Sql & ", " & mvarDAI_VLR_ATAERO
   Sql = Sql & ", " & mvarDAI_VLR_PGTO
   Sql = Sql & ", " & mvarDAI_VLR_TARIF
   Sql = Sql & ", " & SqlDate(mvarDAT_REC_CRG)
   Sql = Sql & ", " & Aspas(mvarDET_COD)
   Sql = Sql & ", " & Aspas(mvarDFAT_NUM)
   Sql = Sql & ", " & SqlDate(mvarDLIB_DAT)
   Sql = Sql & ", " & Aspas(mvarDLIB_NUM)
   Sql = Sql & ", " & Aspas(mvarDLIB_TIP)
   Sql = Sql & ", " & Aspas(mvarDSPCH_NUM)
   Sql = Sql & ", " & Aspas(mvarENQ_COD)
   Sql = Sql & ", " & Aspas(mvarESP_MIN_PRE)
   Sql = Sql & ", " & Aspas(mvarESP_NUM_DOC)
   Sql = Sql & ", " & Aspas(mvarHCRG_NUM)
   Sql = Sql & ", " & Aspas(mvarIMP_COD)
   Sql = Sql & ", " & Aspas(mvarNUM_TERMO)
   Sql = Sql & ", " & Aspas(mvarPESS_COD)
   Sql = Sql & ", " & mvarTXA_USD_REAL
   Sql = Sql & ", " & mvarDAI_VLR_RET
   Sql = Sql & ", " & mvarDAI_VLR_COR
   Sql = Sql & ", " & mvarDAI_VLR_MUL
   Sql = Sql & ", " & mvarDAI_VLR_JUR
   Sql = Sql & ", " & Aspas(mvarDAI_MOT_BAIXA)
   Sql = Sql & ", " & SqlDate(mvarDAI_DAT_APR)
   Sql = Sql & ", " & mvarDAI_VLR_APR
   Sql = Sql & ", " & Aspas(mvarDAI_IDT_TABN)
   Sql = Sql & ")"
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete() As String
   Dim Sql$
   If Trim(mvarDAI_NUM) = "" Then Exit Property
   If Trim(mvarDAI_SEQ) = "" Then Exit Property
   If Trim(mvarDAI_DV) = "" Then Exit Property
   Sql$ = "Delete * From DAI"
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " DAI_NUM = " & Aspas(mvarDAI_NUM)
   Sql$ = Sql$ & " and DAI_SEQ = " & Aspas(mvarDAI_SEQ)
   Sql$ = Sql$ & " and DAI_DV = " & Aspas(mvarDAI_DV)
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql$
   Sql$ = "update DAI set "
   Sql$ = Sql$ & " DAI_DV = " & Aspas(mvarDAI_DV)
   Sql$ = Sql$ & " , DAI_NUM = " & Aspas(mvarDAI_NUM)
   Sql$ = Sql$ & " , DAI_SEQ = " & Aspas(mvarDAI_SEQ)
   Sql$ = Sql$ & " , BAN_COD = " & Aspas(mvarBAN_COD)
   Sql$ = Sql$ & " , DAI_COMPL = " & Aspas(mvarDAI_COMPL)
   Sql$ = Sql$ & " , CLI_NOM = " & Aspas(mvarCLI_NOM)
   Sql$ = Sql$ & " , CLI_NUM = " & Aspas(mvarCLI_NUM)
   Sql$ = Sql$ & " , CLI_PF_PJ = " & Aspas(mvarCLI_PF_PJ)
   Sql$ = Sql$ & " , CLI_TIP_COB = " & Aspas(mvarCLI_TIP_COB)
   Sql$ = Sql$ & " , CRG_NUM = " & Aspas(mvarCRG_NUM)
   Sql$ = Sql$ & " , DAI_DAT_EMI = " & SqlDate(mvarDAI_DAT_EMI)
   Sql$ = Sql$ & " , DAI_DAT_ENTRG = " & SqlDate(mvarDAI_DAT_ENTRG)
   Sql$ = Sql$ & " , DAI_DAT_ISEN = " & SqlDate(mvarDAI_DAT_ISEN)
   Sql$ = Sql$ & " , DAI_DAT_LIB = " & SqlDate(mvarDAI_DAT_LIB)
   Sql$ = Sql$ & " , DAI_DAT_PGTO = " & SqlDate(mvarDAI_DAT_PGTO)
   Sql$ = Sql$ & " , DAI_DAT_VENCTO = " & SqlDate(mvarDAI_DAT_VENCTO)
   Sql$ = Sql$ & " , DAI_HOR = " & Aspas(mvarDAI_HOR)
   Sql$ = Sql$ & " , DAI_IDT_REIMP = " & Aspas(mvarDAI_IDT_REIMP)
   Sql$ = Sql$ & " , DAI_MOT_ALT = " & Aspas(mvarDAI_MOT_ALT)
   Sql$ = Sql$ & " , DAI_MOT_VENC = " & Aspas(mvarDAI_MOT_VENC)
   Sql$ = Sql$ & " , DAI_NUM_MAN = " & Aspas(mvarDAI_NUM_MAN)
   Sql$ = Sql$ & " , DAI_NUM_PERI = " & mvarDAI_NUM_PERI
   Sql$ = Sql$ & " , DAI_PES_BRU = " & mvarDAI_PES_BRU
   Sql$ = Sql$ & " , DAI_PES_LIQ = " & mvarDAI_PES_LIQ
   Sql$ = Sql$ & " , DAI_QTD_VOL = " & mvarDAI_QTD_VOL
   Sql$ = Sql$ & " , DAI_SIT = " & Aspas(mvarDAI_SIT)
   Sql$ = Sql$ & " , DAI_TIP_TRAT = " & Aspas(mvarDAI_TIP_TRAT)
   Sql$ = Sql$ & " , DAI_TOT_CIF = " & mvarDAI_TOT_CIF
   Sql$ = Sql$ & " , DAI_TOT_FOB = " & mvarDAI_TOT_FOB
   Sql$ = Sql$ & " , DAI_TOT_FRT = " & mvarDAI_TOT_FRT
   Sql$ = Sql$ & " , DAI_TOT_SGR = " & mvarDAI_TOT_SGR
   Sql$ = Sql$ & " , DAI_VLR_ABAT = " & mvarDAI_VLR_ABAT
   Sql$ = Sql$ & " , DAI_VLR_ATAERO = " & mvarDAI_VLR_ATAERO
   Sql$ = Sql$ & " , DAI_VLR_PGTO = " & mvarDAI_VLR_PGTO
   Sql$ = Sql$ & " , DAI_VLR_TARIF = " & mvarDAI_VLR_TARIF
   Sql$ = Sql$ & " , DAT_REC_CRG = " & SqlDate(mvarDAT_REC_CRG)
   Sql$ = Sql$ & " , DET_COD = " & Aspas(mvarDET_COD)
   Sql$ = Sql$ & " , DFAT_NUM = " & Aspas(mvarDFAT_NUM)
   Sql$ = Sql$ & " , DLIB_DAT = " & SqlDate(mvarDLIB_DAT)
   Sql$ = Sql$ & " , DLIB_NUM = " & Aspas(mvarDLIB_NUM)
   Sql$ = Sql$ & " , DLIB_TIP = " & Aspas(mvarDLIB_TIP)
   Sql$ = Sql$ & " , DSPCH_NUM = " & Aspas(mvarDSPCH_NUM)
   Sql$ = Sql$ & " , ENQ_COD = " & Aspas(mvarENQ_COD)
   Sql$ = Sql$ & " , ESP_MIN_PRE = " & Aspas(mvarESP_MIN_PRE)
   Sql$ = Sql$ & " , ESP_NUM_DOC = " & Aspas(mvarESP_NUM_DOC)
   Sql$ = Sql$ & " , HCRG_NUM = " & Aspas(mvarHCRG_NUM)
   Sql$ = Sql$ & " , IMP_COD = " & Aspas(mvarIMP_COD)
   Sql$ = Sql$ & " , NUM_TERMO = " & Aspas(mvarNUM_TERMO)
   Sql$ = Sql$ & " , PESS_COD = " & Aspas(mvarPESS_COD)
   Sql$ = Sql$ & " , TXA_USD_REAL = " & mvarTXA_USD_REAL
   Sql$ = Sql$ & " , DAI_VLR_RET = " & mvarDAI_VLR_RET
   Sql$ = Sql$ & " , DAI_VLR_COR = " & mvarDAI_VLR_COR
   Sql$ = Sql$ & " , DAI_VLR_MUL = " & mvarDAI_VLR_MUL
   Sql$ = Sql$ & " , DAI_VLR_JUR = " & mvarDAI_VLR_JUR
   Sql$ = Sql$ & " , DAI_MOT_BAIXA = " & Aspas(mvarDAI_MOT_BAIXA)
   Sql$ = Sql$ & " , DAI_DAT_APR = " & SqlDate(mvarDAI_DAT_APR)
   Sql$ = Sql$ & " , DAI_VLR_APR = " & mvarDAI_VLR_APR
   Sql$ = Sql$ & " , DAI_IDT_TABN = " & Aspas(mvarDAI_IDT_TABN)
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " DAI_NUM = " & Aspas(mvarDAI_NUM)
   Sql$ = Sql$ & " and DAI_SEQ = " & Aspas(mvarDAI_SEQ)
   Sql$ = Sql$ & " and DAI_DV = " & Aspas(mvarDAI_DV)
   mvarQryUpDate = Sql$
   QryUpDate = mvarQryUpDate
End Property
Public Function Incluir() As Integer
   Call mvardBase.Executa(Me.QryInsert)
   Incluir = mvardBase.CodeSql
End Function
Public Function Gravar(Optional ByVal ExibeResult = True) As Variant
   Dim Result
   Select Case mvarEXISTE
      Case ALTERACAO: Result = Alterar
      Case INCLUSAO: Result = Incluir()
   End Select
   If Not ExibeResult Then Exit Function
   If Result = FOUND Then
      Call ExibirAviso(LoadMsg(34), LoadMsg(57))
   Else
      Call ExibirAviso(LoadMsg(48), LoadMsg(57))
   End If
End Function
Public Function GetSelect(Ch_DAI_NUM$, Ch_DAI_SEQ$, Ch_DAI_DV$) As Integer
   Dim Sql$
   mvarDAI_NUM = Ch_DAI_NUM
   mvarDAI_SEQ = Ch_DAI_SEQ
   mvarDAI_DV = Ch_DAI_DV

   Sql$ = "select distinct DAI_DV,DAI_NUM,DAI_SEQ,BAN_COD,DAI_COMPL,CLI_NOM,CLI_NUM,CLI_PF_PJ,CLI_TIP_COB,CRG_NUM,DAI_DAT_EMI,DAI_DAT_ENTRG,DAI_DAT_ISEN,DAI_DAT_LIB,DAI_DAT_PGTO,DAI_DAT_VENCTO,DAI_HOR,DAI_IDT_REIMP,DAI_MOT_ALT,DAI_MOT_VENC,DAI_NUM_MAN,DAI_NUM_PERI,DAI_PES_BRU,DAI_PES_LIQ,DAI_QTD_VOL,DAI_SIT,DAI_TIP_TRAT,DAI_TOT_CIF,DAI_TOT_FOB,DAI_TOT_FRT,DAI_TOT_SGR,DAI_VLR_ABAT,DAI_VLR_ATAERO,DAI_VLR_PGTO,DAI_VLR_TARIF,DAT_REC_CRG,DET_COD,DFAT_NUM,DLIB_DAT,DLIB_NUM,DLIB_TIP,DSPCH_NUM,ENQ_COD,ESP_MIN_PRE,ESP_NUM_DOC,HCRG_NUM,IMP_COD,NUM_TERMO,PESS_COD,TXA_USD_REAL,DAI_VLR_RET,DAI_VLR_COR,DAI_VLR_MUL,DAI_VLR_JUR,DAI_MOT_BAIXA,DAI_DAT_APR,DAI_VLR_APR,DAI_IDT_TABN"
   Sql$ = Sql$ & " From DAI"
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " DAI_NUM = " & Aspas(mvarDAI_NUM)
   Sql$ = Sql$ & " and DAI_SEQ = " & Aspas(mvarDAI_SEQ)
   Sql$ = Sql$ & " and DAI_DV = " & Aspas(mvarDAI_DV)
   Call mvardBase.AbreTabela(Sql$)
   Select Case mvardBase.CodeSql
      Case FOUND
         GetSelect = ALTERACAO
         With mvardBase.Dys
            mvarDAI_DV = !DAI_DV
            mvarDAI_NUM = !DAI_NUM & ""
            mvarDAI_SEQ = !DAI_SEQ & ""
            mvarBAN_COD = !BAN_COD & ""
            mvarDAI_COMPL = !DAI_COMPL & ""
            mvarCLI_NOM = !CLI_NOM & ""
            mvarCLI_NUM = !CLI_NUM & ""
            mvarCLI_PF_PJ = !CLI_PF_PJ & ""
            mvarCLI_TIP_COB = !CLI_TIP_COB & ""
            mvarCRG_NUM = !CRG_NUM & ""
            mvarDAI_DAT_EMI = Format(!DAI_DAT_EMI & "", "DD/MM/YYYY")
            mvarDAI_DAT_ENTRG = Format(!DAI_DAT_ENTRG & "", "DD/MM/YYYY")
            mvarDAI_DAT_ISEN = Format(!DAI_DAT_ISEN & "", "DD/MM/YYYY")
            mvarDAI_DAT_LIB = Format(!DAI_DAT_LIB & "", "DD/MM/YYYY")
            mvarDAI_DAT_PGTO = Format(!DAI_DAT_PGTO & "", "DD/MM/YYYY")
            mvarDAI_DAT_VENCTO = Format(!DAI_DAT_VENCTO & "", "DD/MM/YYYY")
            mvarDAI_HOR = !DAI_HOR & ""
            mvarDAI_IDT_REIMP = !DAI_IDT_REIMP & ""
            mvarDAI_MOT_ALT = !DAI_MOT_ALT & ""
            mvarDAI_MOT_VENC = !DAI_MOT_VENC & ""
            mvarDAI_NUM_MAN = !DAI_NUM_MAN & ""
            mvarDAI_NUM_PERI = !DAI_NUM_PERI & ""
            mvarDAI_PES_BRU = !DAI_PES_BRU & ""
            mvarDAI_PES_LIQ = !DAI_PES_LIQ & ""
            mvarDAI_QTD_VOL = !DAI_QTD_VOL & ""
            mvarDAI_SIT = !DAI_SIT & ""
            mvarDAI_TIP_TRAT = !DAI_TIP_TRAT & ""
            mvarDAI_TOT_CIF = !DAI_TOT_CIF & ""
            mvarDAI_TOT_FOB = !DAI_TOT_FOB & ""
            mvarDAI_TOT_FRT = !DAI_TOT_FRT & ""
            mvarDAI_TOT_SGR = !DAI_TOT_SGR & ""
            mvarDAI_VLR_ABAT = !DAI_VLR_ABAT & ""
            mvarDAI_VLR_ATAERO = !DAI_VLR_ATAERO & ""
            mvarDAI_VLR_PGTO = !DAI_VLR_PGTO & ""
            mvarDAI_VLR_TARIF = !DAI_VLR_TARIF & ""
            mvarDAT_REC_CRG = Format(!DAT_REC_CRG & "", "DD/MM/YYYY")
            mvarDET_COD = !DET_COD & ""
            mvarDFAT_NUM = !DFAT_NUM & ""
            mvarDLIB_DAT = Format(!DLIB_DAT & "", "DD/MM/YYYY")
            mvarDLIB_NUM = !DLIB_NUM & ""
            mvarDLIB_TIP = !DLIB_TIP & ""
            mvarDSPCH_NUM = !DSPCH_NUM & ""
            mvarENQ_COD = !ENQ_COD & ""
            mvarESP_MIN_PRE = !ESP_MIN_PRE & ""
            mvarESP_NUM_DOC = !ESP_NUM_DOC & ""
            mvarHCRG_NUM = !HCRG_NUM & ""
            mvarIMP_COD = !IMP_COD & ""
            mvarNUM_TERMO = !NUM_TERMO & ""
            mvarPESS_COD = !PESS_COD & ""
            mvarTXA_USD_REAL = !TXA_USD_REAL & ""
            mvarDAI_VLR_RET = !DAI_VLR_RET & ""
            mvarDAI_VLR_COR = !DAI_VLR_COR & ""
            mvarDAI_VLR_MUL = !DAI_VLR_MUL & ""
            mvarDAI_VLR_JUR = !DAI_VLR_JUR & ""
            mvarDAI_MOT_BAIXA = !DAI_MOT_BAIXA & ""
            mvarDAI_DAT_APR = Format(!DAI_DAT_APR & "", "DD/MM/YYYY")
            mvarDAI_VLR_APR = !DAI_VLR_APR & ""
            mvarDAI_IDT_TABN = !DAI_IDT_TABN & ""
            .Close
         End With
      Case NOTFOUND: GetSelect = INCLUSAO
      Case ERRO: GetSelect = ERRO
   End Select
   mvarEXISTE = GetSelect
End Function
Public Function Excluir() As Variant
   If mvardBase.GravarTabela(Me.QryDelete, EXCLUSAO) = vbYes Then
      Excluir = True
   End If
End Function
Public Function Alterar() As Integer
   Call mvardBase.Executa(Me.QryUpDate)
   Alterar = mvardBase.CodeSql
End Function
Public Function AlterarChave(Ch_DAI_NUM$, Ch_DAI_SEQ$, Ch_DAI_DV$) As Integer
   Dim Sql$
   Sql$ = "update DAI set "
   Sql$ = Sql$ & " DAI_NUM = " & Aspas(Ch_DAI_NUM)
   Sql$ = Sql$ & " , DAI_SEQ = " & Aspas(Ch_DAI_SEQ)
   Sql$ = Sql$ & " , DAI_DV = " & Aspas(Ch_DAI_DV)
   Sql$ = Sql$ + " Where "
   Sql$ = Sql$ & " DAI_NUM = " & Aspas(mvarDAI_NUM)
   Sql$ = Sql$ & " and DAI_SEQ = " & Aspas(mvarDAI_SEQ)
   Sql$ = Sql$ & " and DAI_DV = " & Aspas(mvarDAI_DV)
   Call mvardBase.Executa(Sql)
   AlterarChave = mvardBase.CodeSql
End Function
Private Sub Class_Terminate()
   Set mvardBase = Nothing
End Sub
