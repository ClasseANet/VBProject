VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TB_CRG_TERMO"
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

Private mvarCRG_NUM As String
Private mvarHCRG_NUM As String
Private mvarNUM_TERMO As String
Private mvarAER_ORI As String
Private mvarCIA_SIG As String
Private mvarCLI_NOM As String
Private mvarCRTE_CLI_PRF As String
Private mvarCRTE_DAT_ARM As String
Private mvarCRTE_DAT_MANTRA As String
Private mvarCRTE_DAT_REC As String
Private mvarCRTE_DES As String
Private mvarCRTE_DSP_ANT As String
Private mvarCRTE_HOR_REC As String
Private mvarCRTE_IDT_BAG As String
Private mvarCRTE_IDT_MNT As String
Private mvarCRTE_MOE_FOB As String
Private mvarCRTE_MOE_FRT As String
Private mvarCRTE_MOE_SGR As String
Private mvarCRTE_PES_DCL As String
Private mvarCRTE_PES_REC As String
Private mvarCRTE_PES_SALDO As String
Private mvarCRTE_PES_TMP As String
Private mvarCRTE_SIT_REC As String
Private mvarCRTE_TIP As String
Private mvarCRTE_TIP_DISP As String
Private mvarCRTE_TIP_DOC As String
Private mvarCRTE_TIP_TRAT As String
Private mvarCRTE_TOT_PAR As String
Private mvarCRTE_VLR_FOB As String
Private mvarCRTE_VLR_FRT As String
Private mvarCRTE_VLR_SGR As String
Private mvarCRTE_VOL_DCL As String
Private mvarCRTE_VOL_REC As String
Private mvarCRTE_VOL_SALDO As String
Private mvarCRTE_VOL_TMP As String
Private mvarEMB_COD As String
Private mvarNAT_CRG As String
Private mvarPESO_UND As String
Private mvarRALF_COD As String
Private mvarAER_DST As String
Private mvarPESS_CPF As String
Private mvarCRTE_DISP_RESP As String
Private mvarCRTE_DISP_DAT As String
Private mvarCRTE_DISP_RECF As String
Private mvarRESP_NUM As String
Private mvarRES_NOM As String
Private mvarTIPO_DOC_RESP As String
Private mvarRESP_NOM As String
Private mvarDOC_SUFRAMA As String
Private mvarCRTE_MAND_JUD As String
Private mvarCRTE_DAT_JUD As String
Private mvarRESP_PLACA_VEI As String
Private mvarRESP_CRACHA As String
Private mvarAER_COD As String

Private mPCGs As Collection
Private mDAIs As Collection
Public Property Let AER_COD(ByVal vData As String)
   mvarAER_COD = vData
End Property
Public Property Get AER_COD() As String
   AER_COD = mvarAER_COD
End Property
Public Property Let CRG_NUM(ByVal vData As String)
   mvarCRG_NUM = vData
End Property
Public Property Get CRG_NUM() As String
   CRG_NUM = mvarCRG_NUM
End Property
Public Property Let HCRG_NUM(ByVal vData As String)
   vData = IIf(Trim(vData) = "", " ", vData)
   mvarHCRG_NUM = vData
End Property
Public Property Get HCRG_NUM() As String
   HCRG_NUM = mvarHCRG_NUM
End Property
Public Property Let NUM_TERMO(ByVal vData As String)
   mvarNUM_TERMO = vData
End Property
Public Property Get NUM_TERMO() As String
   NUM_TERMO = mvarNUM_TERMO
End Property
Public Property Let AER_ORI(ByVal vData As String)
   mvarAER_ORI = vData
End Property
Public Property Get AER_ORI() As String
   AER_ORI = mvarAER_ORI
End Property
Public Property Let CIA_SIG(ByVal vData As String)
   mvarCIA_SIG = vData
End Property
Public Property Get CIA_SIG() As String
   CIA_SIG = mvarCIA_SIG
End Property
Public Property Let CLI_NOM(ByVal vData As String)
   mvarCLI_NOM = vData
End Property
Public Property Get CLI_NOM() As String
   CLI_NOM = mvarCLI_NOM
End Property
Public Property Let CRTE_CLI_PRF(ByVal vData As String)
   mvarCRTE_CLI_PRF = vData
End Property
Public Property Get CRTE_CLI_PRF() As String
   CRTE_CLI_PRF = mvarCRTE_CLI_PRF
End Property
Public Property Let CRTE_DAT_ARM(ByVal vData As String)
   mvarCRTE_DAT_ARM = vData
End Property
Public Property Get CRTE_DAT_ARM() As String
   CRTE_DAT_ARM = mvarCRTE_DAT_ARM
End Property
Public Property Let CRTE_DAT_MANTRA(ByVal vData As String)
   mvarCRTE_DAT_MANTRA = vData
End Property
Public Property Get CRTE_DAT_MANTRA() As String
   CRTE_DAT_MANTRA = mvarCRTE_DAT_MANTRA
End Property
Public Property Let CRTE_DAT_REC(ByVal vData As String)
   mvarCRTE_DAT_REC = vData
End Property
Public Property Get CRTE_DAT_REC() As String
   CRTE_DAT_REC = mvarCRTE_DAT_REC
End Property
Public Property Let CRTE_DES(ByVal vData As String)
   mvarCRTE_DES = vData
End Property
Public Property Get CRTE_DES() As String
   CRTE_DES = mvarCRTE_DES
End Property
Public Property Let CRTE_DSP_ANT(ByVal vData As String)
   mvarCRTE_DSP_ANT = vData
End Property
Public Property Get CRTE_DSP_ANT() As String
   CRTE_DSP_ANT = mvarCRTE_DSP_ANT
End Property
Public Property Let CRTE_HOR_REC(ByVal vData As String)
   mvarCRTE_HOR_REC = vData
End Property
Public Property Get CRTE_HOR_REC() As String
   CRTE_HOR_REC = mvarCRTE_HOR_REC
End Property
Public Property Let CRTE_IDT_BAG(ByVal vData As String)
   mvarCRTE_IDT_BAG = vData
End Property
Public Property Get CRTE_IDT_BAG() As String
   CRTE_IDT_BAG = mvarCRTE_IDT_BAG
End Property
Public Property Let CRTE_IDT_MNT(ByVal vData As String)
   mvarCRTE_IDT_MNT = vData
End Property
Public Property Get CRTE_IDT_MNT() As String
   CRTE_IDT_MNT = mvarCRTE_IDT_MNT
End Property
Public Property Let CRTE_MOE_FOB(ByVal vData As String)
   mvarCRTE_MOE_FOB = vData
End Property
Public Property Get CRTE_MOE_FOB() As String
   CRTE_MOE_FOB = mvarCRTE_MOE_FOB
End Property
Public Property Let CRTE_MOE_FRT(ByVal vData As String)
   mvarCRTE_MOE_FRT = vData
End Property
Public Property Get CRTE_MOE_FRT() As String
   CRTE_MOE_FRT = mvarCRTE_MOE_FRT
End Property
Public Property Let CRTE_MOE_SGR(ByVal vData As String)
   mvarCRTE_MOE_SGR = vData
End Property
Public Property Get CRTE_MOE_SGR() As String
   CRTE_MOE_SGR = mvarCRTE_MOE_SGR
End Property
Public Property Let CRTE_PES_DCL(ByVal vData As String)
   mvarCRTE_PES_DCL = vData
End Property
Public Property Get CRTE_PES_DCL() As String
   CRTE_PES_DCL = mvarCRTE_PES_DCL
End Property
Public Property Let CRTE_PES_REC(ByVal vData As String)
   mvarCRTE_PES_REC = vData
End Property
Public Property Get CRTE_PES_REC() As String
   CRTE_PES_REC = mvarCRTE_PES_REC
End Property
Public Property Let CRTE_PES_SALDO(ByVal vData As String)
   mvarCRTE_PES_SALDO = vData
End Property
Public Property Get CRTE_PES_SALDO() As String
   CRTE_PES_SALDO = mvarCRTE_PES_SALDO
End Property
Public Property Let CRTE_PES_TMP(ByVal vData As String)
   mvarCRTE_PES_TMP = vData
End Property
Public Property Get CRTE_PES_TMP() As String
   CRTE_PES_TMP = mvarCRTE_PES_TMP
End Property
Public Property Let CRTE_SIT_REC(ByVal vData As String)
   mvarCRTE_SIT_REC = vData
End Property
Public Property Get CRTE_SIT_REC() As String
   CRTE_SIT_REC = mvarCRTE_SIT_REC
End Property
Public Property Let CRTE_TIP(ByVal vData As String)
   mvarCRTE_TIP = vData
End Property
Public Property Get CRTE_TIP() As String
   CRTE_TIP = mvarCRTE_TIP
End Property
Public Property Let CRTE_TIP_DISP(ByVal vData As String)
   mvarCRTE_TIP_DISP = vData
End Property
Public Property Get CRTE_TIP_DISP() As String
   CRTE_TIP_DISP = mvarCRTE_TIP_DISP
End Property
Public Property Let CRTE_TIP_DOC(ByVal vData As String)
   mvarCRTE_TIP_DOC = vData
End Property
Public Property Get CRTE_TIP_DOC() As String
   CRTE_TIP_DOC = mvarCRTE_TIP_DOC
End Property
Public Property Let CRTE_TIP_TRAT(ByVal vData As String)
   mvarCRTE_TIP_TRAT = vData
End Property
Public Property Get CRTE_TIP_TRAT() As String
   CRTE_TIP_TRAT = mvarCRTE_TIP_TRAT
End Property
Public Property Let CRTE_TOT_PAR(ByVal vData As String)
   mvarCRTE_TOT_PAR = vData
End Property
Public Property Get CRTE_TOT_PAR() As String
   CRTE_TOT_PAR = mvarCRTE_TOT_PAR
End Property
Public Property Let CRTE_VLR_FOB(ByVal vData As String)
   mvarCRTE_VLR_FOB = vData
End Property
Public Property Get CRTE_VLR_FOB() As String
   CRTE_VLR_FOB = mvarCRTE_VLR_FOB
End Property
Public Property Let CRTE_VLR_FRT(ByVal vData As String)
   mvarCRTE_VLR_FRT = vData
End Property
Public Property Get CRTE_VLR_FRT() As String
   CRTE_VLR_FRT = mvarCRTE_VLR_FRT
End Property
Public Property Let CRTE_VLR_SGR(ByVal vData As String)
   mvarCRTE_VLR_SGR = vData
End Property
Public Property Get CRTE_VLR_SGR() As String
   CRTE_VLR_SGR = mvarCRTE_VLR_SGR
End Property
Public Property Let CRTE_VOL_DCL(ByVal vData As String)
   mvarCRTE_VOL_DCL = vData
End Property
Public Property Get CRTE_VOL_DCL() As String
   CRTE_VOL_DCL = mvarCRTE_VOL_DCL
End Property
Public Property Let CRTE_VOL_REC(ByVal vData As String)
   mvarCRTE_VOL_REC = vData
End Property
Public Property Get CRTE_VOL_REC() As String
   CRTE_VOL_REC = mvarCRTE_VOL_REC
End Property
Public Property Let CRTE_VOL_SALDO(ByVal vData As String)
   mvarCRTE_VOL_SALDO = vData
End Property
Public Property Get CRTE_VOL_SALDO() As String
   CRTE_VOL_SALDO = mvarCRTE_VOL_SALDO
End Property
Public Property Let CRTE_VOL_TMP(ByVal vData As String)
   mvarCRTE_VOL_TMP = vData
End Property
Public Property Get CRTE_VOL_TMP() As String
   CRTE_VOL_TMP = mvarCRTE_VOL_TMP
End Property
Public Property Let EMB_COD(ByVal vData As String)
   mvarEMB_COD = vData
End Property
Public Property Get EMB_COD() As String
   EMB_COD = mvarEMB_COD
End Property
Public Property Let NAT_CRG(ByVal vData As String)
   mvarNAT_CRG = vData
End Property
Public Property Get NAT_CRG() As String
   NAT_CRG = mvarNAT_CRG
End Property
Public Property Let PESO_UND(ByVal vData As String)
   mvarPESO_UND = vData
End Property
Public Property Get PESO_UND() As String
   PESO_UND = mvarPESO_UND
End Property
Public Property Let RALF_COD(ByVal vData As String)
   mvarRALF_COD = vData
End Property
Public Property Get RALF_COD() As String
   RALF_COD = mvarRALF_COD
End Property
Public Property Let AER_DST(ByVal vData As String)
   mvarAER_DST = vData
End Property
Public Property Get AER_DST() As String
   AER_DST = mvarAER_DST
End Property
Public Property Let PESS_CPF(ByVal vData As String)
   mvarPESS_CPF = vData
End Property
Public Property Get PESS_CPF() As String
   PESS_CPF = mvarPESS_CPF
End Property
Public Property Let CRTE_DISP_RESP(ByVal vData As String)
   mvarCRTE_DISP_RESP = vData
End Property
Public Property Get CRTE_DISP_RESP() As String
   CRTE_DISP_RESP = mvarCRTE_DISP_RESP
End Property
Public Property Let CRTE_DISP_DAT(ByVal vData As String)
   mvarCRTE_DISP_DAT = vData
End Property
Public Property Get CRTE_DISP_DAT() As String
   CRTE_DISP_DAT = mvarCRTE_DISP_DAT
End Property
Public Property Let CRTE_DISP_RECF(ByVal vData As String)
   mvarCRTE_DISP_RECF = vData
End Property
Public Property Get CRTE_DISP_RECF() As String
   CRTE_DISP_RECF = mvarCRTE_DISP_RECF
End Property
Public Property Let RESP_NUM(ByVal vData As String)
   mvarRESP_NUM = vData
End Property
Public Property Get RESP_NUM() As String
   RESP_NUM = mvarRESP_NUM
End Property
Public Property Let RES_NOM(ByVal vData As String)
   mvarRES_NOM = vData
End Property
Public Property Get RES_NOM() As String
   RES_NOM = mvarRES_NOM
End Property
Public Property Let TIPO_DOC_RESP(ByVal vData As String)
   mvarTIPO_DOC_RESP = vData
End Property
Public Property Get TIPO_DOC_RESP() As String
   TIPO_DOC_RESP = mvarTIPO_DOC_RESP
End Property
Public Property Let RESP_NOM(ByVal vData As String)
   mvarRESP_NOM = vData
End Property
Public Property Get RESP_NOM() As String
   RESP_NOM = mvarRESP_NOM
End Property
Public Property Let DOC_SUFRAMA(ByVal vData As String)
   mvarDOC_SUFRAMA = vData
End Property
Public Property Get DOC_SUFRAMA() As String
   DOC_SUFRAMA = mvarDOC_SUFRAMA
End Property
Public Property Let CRTE_MAND_JUD(ByVal vData As String)
   mvarCRTE_MAND_JUD = vData
End Property
Public Property Get CRTE_MAND_JUD() As String
   CRTE_MAND_JUD = mvarCRTE_MAND_JUD
End Property
Public Property Let CRTE_DAT_JUD(ByVal vData As String)
   mvarCRTE_DAT_JUD = vData
End Property
Public Property Get CRTE_DAT_JUD() As String
   CRTE_DAT_JUD = mvarCRTE_DAT_JUD
End Property
Public Property Let RESP_PLACA_VEI(ByVal vData As String)
   mvarRESP_PLACA_VEI = vData
End Property
Public Property Get RESP_PLACA_VEI() As String
   RESP_PLACA_VEI = mvarRESP_PLACA_VEI
End Property
Public Property Let RESP_CRACHA(ByVal vData As String)
   mvarRESP_CRACHA = vData
End Property
Public Property Get RESP_CRACHA() As String
   RESP_CRACHA = mvarRESP_CRACHA
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
Public Property Get PCGs() As Collection
   Dim Sql$

   If mPCGs Is Nothing Then
      Set mPCGs = New Collection
      Sql = "SELECT PCG_NINF, CRG_NUM, HCRG_NUM, NUM_TERMO, SET_NUM"
      Sql = Sql & ", ARM_COD, BOX_NUM, PCG_VOL, PCG_PES, HIS_PCG_DAT"
      Sql = Sql & ", HIS_PCG_SIT, PCG_IDT_DSIC, HIS_PCG_HOR, PCG_DES "
      Sql = Sql & ", PCG_TEMP, AER_COD "
      Sql = Sql & " FROM PCG_IMP "
      Sql = Sql & " WHERE NUM_TERMO = " & Aspas(mvarNUM_TERMO)
      Sql = Sql & " AND CRG_NUM = " & Aspas(mvarCRG_NUM)
      Sql = Sql & " AND HCRG_NUM = " & Aspas(IIf(mvarHCRG_NUM = "", "' '", mvarHCRG_NUM))
      Sql = Sql & " AND AER_COD = " & Aspas(mvarAER_COD)
      Call mvardBase.AbreTabela(Sql)
      If mvardBase.CodeSql Then
         With mvardBase.Dys
            While Not .EOF = True
               Dim MyPCG As New TB_PCG_IMP
               
               MyPCG.dBase = mvardBase
               MyPCG.PCG_NINF = .Fields("PCG_NINF") & ""
               MyPCG.CRG_NUM = .Fields("CRG_NUM") & ""
               MyPCG.HCRG_NUM = .Fields("HCRG_NUM") & ""
               MyPCG.NUM_TERMO = .Fields("NUM_TERMO") & ""
               MyPCG.ARM_COD = .Fields("ARM_COD") & ""
               MyPCG.BOX_NUM = .Fields("BOX_NUM") & ""
               MyPCG.SET_NUM = .Fields("SET_NUM") & ""
               MyPCG.HIS_PCG_SIT = .Fields("HIS_PCG_SIT") & ""
               MyPCG.HIS_PCG_DAT = .Fields("HIS_PCG_DAT") & ""
               MyPCG.HIS_PCG_HOR = .Fields("HIS_PCG_HOR") & ""
               MyPCG.PCG_IDT_DSIC = .Fields("PCG_IDT_DSIC") & ""
               MyPCG.PCG_PES = .Fields("PCG_PES") & ""
               MyPCG.PCG_VOL = .Fields("PCG_VOL") & ""
               MyPCG.PCG_TEMP = .Fields("PCG_TEMP") & ""
               MyPCG.PCG_DES = .Fields("PCG_DES") & ""
               MyPCG.AER_COD = .Fields("AER_COD") & ""
               mPCGs.Add MyPCG, MyPCG.PCG_NINF
               
               .MoveNext
               Set MyPCG = Nothing
            Wend
            .Close
         End With
      End If
   End If

    Set PCGs = mPCGs
End Property
Public Property Set PCGs(vData As Collection)
    Set mPCGs = vData
End Property
Public Property Get DAIs() As Collection
   Dim Sql As String
   Dim sDAI As Recordset
   If mDAIs Is Nothing Then
      Set mDAIs = New Collection
      
      Sql$ = "select distinct AER_COD, DAI_DV, DAI_NUM, DAI_SEQ, BAN_COD"
      Sql$ = Sql$ & ", DAI_COMPL, CLI_NOM, CLI_NUM, CLI_PF_PJ"
      Sql$ = Sql$ & ", CLI_TIP_COB, CRG_NUM, DAI_DAT_EMI, DAI_DAT_ENTRG"
      Sql$ = Sql$ & ", DAI_DAT_ISEN, DAI_DAT_LIB, DAI_DAT_PGTO, DAI_DAT_VENCTO"
      Sql$ = Sql$ & ", DAI_HOR, DAI_IDT_REIMP, DAI_MOT_ALT, DAI_MOT_VENC"
      Sql$ = Sql$ & ", DAI_NUM_MAN, DAI_NUM_PERI, DAI_PES_BRU, DAI_PES_LIQ"
      Sql$ = Sql$ & ", DAI_QTD_VOL, DAI_SIT, DAI_TIP_TRAT, DAI_TOT_CIF"
      Sql$ = Sql$ & ", DAI_TOT_FOB, DAI_TOT_FRT, DAI_TOT_SGR, DAI_VLR_ABAT"
      Sql$ = Sql$ & ", DAI_VLR_ATAERO, DAI_VLR_PGTO, DAI_VLR_TARIF, DAT_REC_CRG"
      Sql$ = Sql$ & ", DET_COD, DFAT_NUM, DLIB_DAT, DLIB_NUM, DLIB_TIP, DSPCH_NUM"
      Sql$ = Sql$ & ", ENQ_COD, ESP_MIN_PRE, ESP_NUM_DOC, HCRG_NUM, IMP_COD"
      Sql$ = Sql$ & ", NUM_TERMO, PESS_COD, TXA_USD_REAL, DAI_VLR_RET, DAI_VLR_COR"
      Sql$ = Sql$ & ", DAI_VLR_MUL, DAI_VLR_JUR, DAI_MOT_BAIXA, DAI_IDT_TABN"
'      Sql$ = Sql$ & ", DAI_VLR_APR, DAI_DAT_APR"
      Sql$ = Sql$ & " From DAI"
      Sql$ = Sql$ & " Where "
      Sql$ = Sql$ & " NUM_TERMO = " & Aspas(mvarNUM_TERMO)
      Sql$ = Sql$ & " and CRG_NUM = " & Aspas(mvarCRG_NUM)
      Sql$ = Sql$ & " and HCRG_NUM = " & Aspas(IIf(mvarHCRG_NUM = "", "' '", mvarHCRG_NUM))
      Sql$ = Sql$ & " and AER_COD = " & Aspas(mvarAER_COD)
      
      Call mvardBase.AbreTabela(Sql)
      If mvardBase.CodeSql Then
         With mvardBase.Dys
            While Not .EOF
               Dim clDAI As New TB_DAI
               clDAI.dBase = mvardBase
               clDAI.AER_COD = .Fields("AER_COD") & ""
               clDAI.BAN_COD = .Fields("BAN_COD") & ""
               clDAI.CLI_NOM = .Fields("CLI_NOM") & ""
               clDAI.CLI_NUM = .Fields("CLI_NUM") & ""
               clDAI.CLI_PF_PJ = .Fields("CLI_PF_PJ") & ""
               clDAI.CLI_TIP_COB = .Fields("CLI_TIP_COB") & ""
               clDAI.DAI_COMPL = .Fields("DAI_COMPL") & ""
               clDAI.DAI_DAT_EMI = .Fields("DAI_DAT_EMI") & ""
               clDAI.DAI_DAT_ENTRG = .Fields("DAI_DAT_ENTRG") & ""
               clDAI.DAI_DAT_ISEN = .Fields("DAI_DAT_ISEN") & ""
               clDAI.DAI_DAT_LIB = .Fields("DAI_DAT_LIB") & ""
               clDAI.DAI_DAT_PGTO = .Fields("DAI_DAT_PGTO") & ""
               clDAI.DAI_DAT_VENCTO = .Fields("DAI_DAT_VENCTO") & ""
               clDAI.DAI_DV = .Fields("DAI_DV") & ""
               clDAI.DAI_HOR = .Fields("DAI_HOR") & ""
               clDAI.DAI_IDT_REIMP = .Fields("DAI_IDT_REIMP") & ""
               clDAI.DAI_IDT_TABN = .Fields("DAI_IDT_TABN") & ""
               clDAI.DAI_MOT_ALT = .Fields("DAI_MOT_ALT") & ""
               clDAI.DAI_MOT_BAIXA = .Fields("DAI_MOT_BAIXA") & ""
               clDAI.DAI_MOT_VENC = .Fields("DAI_MOT_VENC") & ""
               clDAI.DAI_NUM = .Fields("DAI_NUM") & ""
               clDAI.DAI_NUM_MAN = .Fields("DAI_NUM_MAN") & ""
               clDAI.DAI_NUM_PERI = .Fields("DAI_NUM_PERI") & ""
               clDAI.DAI_PES_BRU = .Fields("DAI_PES_BRU") & ""
               clDAI.DAI_PES_LIQ = .Fields("DAI_PES_LIQ") & ""
               clDAI.DAI_QTD_VOL = .Fields("DAI_QTD_VOL") & ""
               clDAI.DAI_SEQ = .Fields("DAI_SEQ") & ""
               clDAI.DAI_SIT = .Fields("DAI_SIT") & ""
               clDAI.DAI_TIP_TRAT = .Fields("DAI_TIP_TRAT") & ""
               clDAI.DAI_TOT_CIF = .Fields("DAI_TOT_CIF") & ""
               clDAI.DAI_TOT_FOB = .Fields("DAI_TOT_FOB") & ""
               clDAI.DAI_TOT_FRT = .Fields("DAI_TOT_FRT") & ""
               clDAI.DAI_TOT_SGR = .Fields("DAI_TOT_SGR") & ""
               clDAI.DAI_VLR_ABAT = .Fields("DAI_VLR_ABAT") & ""
               clDAI.DAI_VLR_ATAERO = .Fields("DAI_VLR_ATAERO") & ""
               clDAI.DAI_VLR_COR = .Fields("DAI_VLR_COR") & ""
               clDAI.DAI_VLR_JUR = .Fields("DAI_VLR_JUR") & ""
               clDAI.DAI_VLR_MUL = .Fields("DAI_VLR_MUL") & ""
               clDAI.DAI_VLR_PGTO = .Fields("DAI_VLR_PGTO") & ""
               clDAI.DAI_VLR_RET = .Fields("DAI_VLR_RET") & ""
               clDAI.DAI_VLR_TARIF = .Fields("DAI_VLR_TARIF") & ""
               clDAI.DAT_REC_CRG = .Fields("DAT_REC_CRG") & ""
               clDAI.DET_COD = .Fields("DET_COD") & ""
               clDAI.DFAT_NUM = .Fields("DFAT_NUM") & ""
               clDAI.DLIB_DAT = .Fields("DLIB_DAT") & ""
               clDAI.DLIB_NUM = .Fields("DLIB_NUM") & ""
               clDAI.DLIB_TIP = .Fields("DLIB_TIP") & ""
               clDAI.DSPCH_NUM = .Fields("DSPCH_NUM") & ""
               clDAI.ENQ_COD = .Fields("ENQ_COD") & ""
               clDAI.ESP_MIN_PRE = .Fields("ESP_MIN_PRE") & ""
               clDAI.ESP_NUM_DOC = .Fields("ESP_NUM_DOC") & ""
               clDAI.HCRG_NUM = .Fields("HCRG_NUM") & ""
               clDAI.IMP_COD = .Fields("IMP_COD") & ""
               clDAI.NUM_TERMO = .Fields("NUM_TERMO") & ""
               clDAI.PESS_COD = .Fields("PESS_COD") & ""
               clDAI.TXA_USD_REAL = .Fields("TXA_USD_REAL") & ""
               '            clDAI.DAI_DAT_APR = .Fields("DAI_DAT_APR") & ""
               '            clDAI.DAI_VLR_APR = .Fields("DAI_VLR_APR") & ""
               mDAIs.Add clDAI, clDAI.DAI_NUM & clDAI.DAI_SEQ & clDAI.DAI_DV
               .MoveNext
               Set clDAI = Nothing
            Wend
           .Close
         End With
      End If
   End If
   Set DAIs = mDAIs
End Property
Public Property Set DAIs(vData As Collection)
    Set mDAIs = vData
End Property
Public Property Get QryInsert() As String
   Dim Sql$
   Sql = "insert into CRG_TERMO (CRG_NUM, HCRG_NUM, NUM_TERMO, AER_ORI, CIA_SIG"
   Sql = Sql & ", CLI_NOM, CRTE_CLI_PRF,CRTE_DAT_ARM, CRTE_DAT_MANTRA, CRTE_DAT_REC"
   Sql = Sql & ", CRTE_DES, CRTE_DSP_ANT, CRTE_HOR_REC, CRTE_IDT_BAG, CRTE_IDT_MNT"
   Sql = Sql & ", CRTE_MOE_FOB, CRTE_MOE_FRT, CRTE_MOE_SGR, CRTE_PES_DCL, CRTE_PES_REC"
   Sql = Sql & ", CRTE_PES_SALDO, CRTE_PES_TMP, CRTE_SIT_REC, CRTE_TIP, CRTE_TIP_DISP"
   Sql = Sql & ", CRTE_TIP_DOC, CRTE_TIP_TRAT, CRTE_TOT_PAR, CRTE_VLR_FOB, CRTE_VLR_FRT"
   Sql = Sql & ", CRTE_VLR_SGR, CRTE_VOL_DCL, CRTE_VOL_REC, CRTE_VOL_SALDO, CRTE_VOL_TMP"
   Sql = Sql & ", EMB_COD, NAT_CRG, PESO_UND, RALF_COD, AER_DST, "
   Sql = Sql & ", PESS_CPF, CRTE_DISP_RESP, CRTE_DISP_DAT, CRTE_DISP_RECF, RESP_NUM"
   Sql = Sql & ", RES_NOM, TIPO_DOC_RESP, RESP_NOM, DOC_SUFRAMA, CRTE_MAND_JUD"
   Sql = Sql & ", CRTE_DAT_JUD, RESP_PLACA_VEI, RESP_CRACHA, AER_COD) "
   Sql = Sql & " Values "
   Sql = Sql & "("
   Sql = Sql & Aspas(mvarCRG_NUM)
   Sql = Sql & ", " & Aspas(mvarHCRG_NUM)
   Sql = Sql & ", " & Aspas(mvarNUM_TERMO)
   Sql = Sql & ", " & Aspas(mvarAER_ORI)
   Sql = Sql & ", " & Aspas(mvarCIA_SIG)
   Sql = Sql & ", " & Aspas(mvarCLI_NOM)
   Sql = Sql & ", " & Aspas(mvarCRTE_CLI_PRF)
   Sql = Sql & ", " & SqlDate(mvarCRTE_DAT_ARM)
   Sql = Sql & ", " & SqlDate(mvarCRTE_DAT_MANTRA)
   Sql = Sql & ", " & SqlDate(mvarCRTE_DAT_REC)
   Sql = Sql & ", " & Aspas(mvarCRTE_DES)
   Sql = Sql & ", " & Aspas(mvarCRTE_DSP_ANT)
   Sql = Sql & ", " & Aspas(mvarCRTE_HOR_REC)
   Sql = Sql & ", " & Aspas(mvarCRTE_IDT_BAG)
   Sql = Sql & ", " & Aspas(mvarCRTE_IDT_MNT)
   Sql = Sql & ", " & Aspas(mvarCRTE_MOE_FOB)
   Sql = Sql & ", " & Aspas(mvarCRTE_MOE_FRT)
   Sql = Sql & ", " & Aspas(mvarCRTE_MOE_SGR)
   Sql = Sql & ", " & mvarCRTE_PES_DCL
   Sql = Sql & ", " & mvarCRTE_PES_REC
   Sql = Sql & ", " & mvarCRTE_PES_SALDO
   Sql = Sql & ", " & mvarCRTE_PES_TMP
   Sql = Sql & ", " & Aspas(mvarCRTE_SIT_REC)
   Sql = Sql & ", " & Aspas(mvarCRTE_TIP)
   Sql = Sql & ", " & Aspas(mvarCRTE_TIP_DISP)
   Sql = Sql & ", " & Aspas(mvarCRTE_TIP_DOC)
   Sql = Sql & ", " & Aspas(mvarCRTE_TIP_TRAT)
   Sql = Sql & ", " & Aspas(mvarCRTE_TOT_PAR)
   Sql = Sql & ", " & mvarCRTE_VLR_FOB
   Sql = Sql & ", " & mvarCRTE_VLR_FRT
   Sql = Sql & ", " & mvarCRTE_VLR_SGR
   Sql = Sql & ", " & mvarCRTE_VOL_DCL
   Sql = Sql & ", " & mvarCRTE_VOL_REC
   Sql = Sql & ", " & mvarCRTE_VOL_SALDO
   Sql = Sql & ", " & mvarCRTE_VOL_TMP
   Sql = Sql & ", " & Aspas(mvarEMB_COD)
   Sql = Sql & ", " & Aspas(mvarNAT_CRG)
   Sql = Sql & ", " & Aspas(mvarPESO_UND)
   Sql = Sql & ", " & Aspas(mvarRALF_COD)
   Sql = Sql & ", " & Aspas(mvarAER_DST)
   Sql = Sql & ", " & Aspas(mvarPESS_CPF)
   Sql = Sql & ", " & Aspas(mvarCRTE_DISP_RESP)
   Sql = Sql & ", " & SqlDate(mvarCRTE_DISP_DAT)
   Sql = Sql & ", " & Aspas(mvarCRTE_DISP_RECF)
   Sql = Sql & ", " & Aspas(mvarRESP_NUM)
   Sql = Sql & ", " & Aspas(mvarRES_NOM)
   Sql = Sql & ", " & Aspas(mvarTIPO_DOC_RESP)
   Sql = Sql & ", " & Aspas(mvarRESP_NOM)
   Sql = Sql & ", " & Aspas(mvarDOC_SUFRAMA)
   Sql = Sql & ", " & Aspas(mvarCRTE_MAND_JUD)
   Sql = Sql & ", " & SqlDate(mvarCRTE_DAT_JUD)
   Sql = Sql & ", " & Aspas(mvarRESP_PLACA_VEI)
   Sql = Sql & ", " & Aspas(mvarRESP_CRACHA)
   Sql = Sql & ", " & Aspas(mvarAER_COD)
   Sql = Sql & ")"
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete() As String
   Dim Sql$
   If Trim(mvarCRG_NUM) = "" Then Exit Property
   If Trim(mvarHCRG_NUM) = "" Then Exit Property
   If Trim(mvarNUM_TERMO) = "" Then Exit Property
   Sql$ = "Delete * From CRG_TERMO"
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " CRG_NUM = " & Aspas(mvarCRG_NUM)
   Sql$ = Sql$ & " and HCRG_NUM = " & Aspas(IIf(mvarHCRG_NUM = "", "' '", mvarHCRG_NUM))
   Sql$ = Sql$ & " and NUM_TERMO = " & Aspas(mvarNUM_TERMO)
   Sql$ = Sql$ & " AER_COD = " & Aspas(mvarAER_COD)
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql$
   Sql$ = "update CRG_TERMO set "
   Sql$ = Sql$ & " CRG_NUM = " & Aspas(mvarCRG_NUM)
   Sql$ = Sql$ & " , HCRG_NUM = " & Aspas(mvarHCRG_NUM)
   Sql$ = Sql$ & " , NUM_TERMO = " & Aspas(mvarNUM_TERMO)
   Sql$ = Sql$ & " , AER_ORI = " & Aspas(mvarAER_ORI)
   Sql$ = Sql$ & " , CIA_SIG = " & Aspas(mvarCIA_SIG)
   Sql$ = Sql$ & " , CLI_NOM = " & Aspas(mvarCLI_NOM)
   Sql$ = Sql$ & " , CRTE_CLI_PRF = " & Aspas(mvarCRTE_CLI_PRF)
   Sql$ = Sql$ & " , CRTE_DAT_ARM = " & SqlDate(mvarCRTE_DAT_ARM)
   Sql$ = Sql$ & " , CRTE_DAT_MANTRA = " & SqlDate(mvarCRTE_DAT_MANTRA)
   Sql$ = Sql$ & " , CRTE_DAT_REC = " & SqlDate(mvarCRTE_DAT_REC)
   Sql$ = Sql$ & " , CRTE_DES = " & Aspas(mvarCRTE_DES)
   Sql$ = Sql$ & " , CRTE_DSP_ANT = " & Aspas(mvarCRTE_DSP_ANT)
   Sql$ = Sql$ & " , CRTE_HOR_REC = " & Aspas(mvarCRTE_HOR_REC)
   Sql$ = Sql$ & " , CRTE_IDT_BAG = " & Aspas(mvarCRTE_IDT_BAG)
   Sql$ = Sql$ & " , CRTE_IDT_MNT = " & Aspas(mvarCRTE_IDT_MNT)
   Sql$ = Sql$ & " , CRTE_MOE_FOB = " & Aspas(mvarCRTE_MOE_FOB)
   Sql$ = Sql$ & " , CRTE_MOE_FRT = " & Aspas(mvarCRTE_MOE_FRT)
   Sql$ = Sql$ & " , CRTE_MOE_SGR = " & Aspas(mvarCRTE_MOE_SGR)
   Sql$ = Sql$ & " , CRTE_PES_DCL = " & mvarCRTE_PES_DCL
   Sql$ = Sql$ & " , CRTE_PES_REC = " & mvarCRTE_PES_REC
   Sql$ = Sql$ & " , CRTE_PES_SALDO = " & mvarCRTE_PES_SALDO
   Sql$ = Sql$ & " , CRTE_PES_TMP = " & mvarCRTE_PES_TMP
   Sql$ = Sql$ & " , CRTE_SIT_REC = " & Aspas(mvarCRTE_SIT_REC)
   Sql$ = Sql$ & " , CRTE_TIP = " & Aspas(mvarCRTE_TIP)
   Sql$ = Sql$ & " , CRTE_TIP_DISP = " & Aspas(mvarCRTE_TIP_DISP)
   Sql$ = Sql$ & " , CRTE_TIP_DOC = " & Aspas(mvarCRTE_TIP_DOC)
   Sql$ = Sql$ & " , CRTE_TIP_TRAT = " & Aspas(mvarCRTE_TIP_TRAT)
   Sql$ = Sql$ & " , CRTE_TOT_PAR = " & Aspas(mvarCRTE_TOT_PAR)
   Sql$ = Sql$ & " , CRTE_VLR_FOB = " & mvarCRTE_VLR_FOB
   Sql$ = Sql$ & " , CRTE_VLR_FRT = " & mvarCRTE_VLR_FRT
   Sql$ = Sql$ & " , CRTE_VLR_SGR = " & mvarCRTE_VLR_SGR
   Sql$ = Sql$ & " , CRTE_VOL_DCL = " & mvarCRTE_VOL_DCL
   Sql$ = Sql$ & " , CRTE_VOL_REC = " & mvarCRTE_VOL_REC
   Sql$ = Sql$ & " , CRTE_VOL_SALDO = " & mvarCRTE_VOL_SALDO
   Sql$ = Sql$ & " , CRTE_VOL_TMP = " & mvarCRTE_VOL_TMP
   Sql$ = Sql$ & " , EMB_COD = " & Aspas(mvarEMB_COD)
   Sql$ = Sql$ & " , NAT_CRG = " & Aspas(mvarNAT_CRG)
   Sql$ = Sql$ & " , PESO_UND = " & Aspas(mvarPESO_UND)
   Sql$ = Sql$ & " , RALF_COD = " & Aspas(mvarRALF_COD)
   Sql$ = Sql$ & " , AER_DST = " & Aspas(mvarAER_DST)
   Sql$ = Sql$ & " , PESS_CPF = " & Aspas(mvarPESS_CPF)
   Sql$ = Sql$ & " , CRTE_DISP_RESP = " & Aspas(mvarCRTE_DISP_RESP)
   Sql$ = Sql$ & " , CRTE_DISP_DAT = " & SqlDate(mvarCRTE_DISP_DAT)
   Sql$ = Sql$ & " , CRTE_DISP_RECF = " & Aspas(mvarCRTE_DISP_RECF)
   Sql$ = Sql$ & " , RESP_NUM = " & Aspas(mvarRESP_NUM)
   Sql$ = Sql$ & " , RES_NOM = " & Aspas(mvarRES_NOM)
   Sql$ = Sql$ & " , TIPO_DOC_RESP = " & Aspas(mvarTIPO_DOC_RESP)
   Sql$ = Sql$ & " , RESP_NOM = " & Aspas(mvarRESP_NOM)
   Sql$ = Sql$ & " , DOC_SUFRAMA = " & Aspas(mvarDOC_SUFRAMA)
   Sql$ = Sql$ & " , CRTE_MAND_JUD = " & Aspas(mvarCRTE_MAND_JUD)
   Sql$ = Sql$ & " , CRTE_DAT_JUD = " & SqlDate(mvarCRTE_DAT_JUD)
   Sql$ = Sql$ & " , RESP_PLACA_VEI = " & Aspas(mvarRESP_PLACA_VEI)
   Sql$ = Sql$ & " , RESP_CRACHA = " & Aspas(mvarRESP_CRACHA)
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " CRG_NUM = " & Aspas(mvarCRG_NUM)
   Sql$ = Sql$ & " and HCRG_NUM = " & Aspas(IIf(mvarHCRG_NUM = "", "' '", mvarHCRG_NUM))
   Sql$ = Sql$ & " and NUM_TERMO = " & Aspas(mvarNUM_TERMO)
   Sql$ = Sql$ & " AER_COD = " & Aspas(mvarAER_COD)
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
Public Function GetSelect(Ch_CRG_NUM$, Ch_HCRG_NUM$, Ch_NUM_TERMO$, Ch_AER_COD$) As Integer
   Dim Sql$
   If Ch_CRG_NUM$ = "" Or Ch_NUM_TERMO$ = "" Then Exit Function
   mvarCRG_NUM = Ch_CRG_NUM
   mvarHCRG_NUM = Ch_HCRG_NUM
   mvarNUM_TERMO = Ch_NUM_TERMO
   mvarAER_COD = Ch_AER_COD
   
   Sql$ = "select distinct CRG_NUM, HCRG_NUM, NUM_TERMO, AER_ORI, CIA_SIG"
   Sql$ = Sql$ & ", CLI_NOM, CRTE_CLI_PRF, CRTE_DAT_ARM, CRTE_DAT_MANTRA"
   Sql$ = Sql$ & ", CRTE_DAT_REC, CRTE_DES, CRTE_DSP_ANT, CRTE_HOR_REC"
   Sql$ = Sql$ & ", CRTE_IDT_BAG, CRTE_IDT_MNT, CRTE_MOE_FOB, CRTE_MOE_FRT"
   Sql$ = Sql$ & ", CRTE_MOE_SGR, CRTE_PES_DCL, CRTE_PES_REC, CRTE_PES_SALDO"
   Sql$ = Sql$ & ", CRTE_PES_TMP, CRTE_SIT_REC, CRTE_TIP, CRTE_TIP_DISP"
   Sql$ = Sql$ & ", CRTE_TIP_DOC, CRTE_TIP_TRAT, CRTE_TOT_PAR, CRTE_VLR_FOB"
   Sql$ = Sql$ & ", CRTE_VLR_FRT, CRTE_VLR_SGR, CRTE_VOL_DCL, CRTE_VOL_REC"
   Sql$ = Sql$ & ", CRTE_VOL_SALDO, CRTE_VOL_TMP, EMB_COD, NAT_CRG, PESO_UND"
   Sql$ = Sql$ & ", RALF_COD, AER_DST, PESS_CPF, CRTE_DISP_RESP, CRTE_DISP_DAT"

   Sql$ = Sql$ & ", CRTE_DISP_RECF, RESP_NUM, RESP_NOM, TIPO_DOC_RESP, RESP_NOM"
   Sql$ = Sql$ & ", DOC_SUFRAMA, RESP_PLACA_VEI, RESP_CRACHA, AER_COD"
   'Sql$ = Sql$ & ", CRTE_MAND_JUD, CRTE_DAT_JUD"
   Sql$ = Sql$ & " From CRG_TERMO"
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " CRG_NUM = " & Aspas(mvarCRG_NUM)
   Sql$ = Sql$ & " and HCRG_NUM = " & Aspas(IIf(mvarHCRG_NUM = "", "' '", mvarHCRG_NUM))
   Sql$ = Sql$ & " and NUM_TERMO = " & Aspas(mvarNUM_TERMO)
   Sql$ = Sql$ & " and AER_COD = " & Aspas(mvarAER_COD)
   Call mvardBase.AbreTabela(Sql$)
   Select Case mvardBase.CodeSql
      Case FOUND
         GetSelect = ALTERACAO
         With mvardBase.Dys
            mvarCRG_NUM = !CRG_NUM
            mvarHCRG_NUM = !HCRG_NUM & ""
            mvarNUM_TERMO = !NUM_TERMO & ""
            mvarAER_COD = !AER_COD & ""
            mvarAER_ORI = !AER_ORI & ""
            mvarCIA_SIG = !CIA_SIG & ""
            mvarCLI_NOM = !CLI_NOM & ""
            mvarCRTE_CLI_PRF = !CRTE_CLI_PRF & ""
            mvarCRTE_DAT_ARM = Format(!CRTE_DAT_ARM & "", "DD/MM/YYYY")
            mvarCRTE_DAT_MANTRA = Format(!CRTE_DAT_MANTRA & "", "DD/MM/YYYY")
            mvarCRTE_DAT_REC = Format(!CRTE_DAT_REC & "", "DD/MM/YYYY")
            mvarCRTE_DES = !CRTE_DES & ""
            mvarCRTE_DSP_ANT = !CRTE_DSP_ANT & ""
            mvarCRTE_HOR_REC = !CRTE_HOR_REC & ""
            mvarCRTE_IDT_BAG = !CRTE_IDT_BAG & ""
            mvarCRTE_IDT_MNT = !CRTE_IDT_MNT & ""
            mvarCRTE_MOE_FOB = !CRTE_MOE_FOB & ""
            mvarCRTE_MOE_FRT = !CRTE_MOE_FRT & ""
            mvarCRTE_MOE_SGR = !CRTE_MOE_SGR & ""
            mvarCRTE_PES_DCL = !CRTE_PES_DCL & ""
            mvarCRTE_PES_REC = !CRTE_PES_REC & ""
            mvarCRTE_PES_SALDO = !CRTE_PES_SALDO & ""
            mvarCRTE_PES_TMP = !CRTE_PES_TMP & ""
            mvarCRTE_SIT_REC = !CRTE_SIT_REC & ""
            mvarCRTE_TIP = !CRTE_TIP & ""
            mvarCRTE_TIP_DISP = !CRTE_TIP_DISP & ""
            mvarCRTE_TIP_DOC = !CRTE_TIP_DOC & ""
            mvarCRTE_TIP_TRAT = !CRTE_TIP_TRAT & ""
            mvarCRTE_TOT_PAR = !CRTE_TOT_PAR & ""
            mvarCRTE_VLR_FOB = !CRTE_VLR_FOB & ""
            mvarCRTE_VLR_FRT = !CRTE_VLR_FRT & ""
            mvarCRTE_VLR_SGR = !CRTE_VLR_SGR & ""
            mvarCRTE_VOL_DCL = !CRTE_VOL_DCL & ""
            mvarCRTE_VOL_REC = !CRTE_VOL_REC & ""
            mvarCRTE_VOL_SALDO = !CRTE_VOL_SALDO & ""
            mvarCRTE_VOL_TMP = !CRTE_VOL_TMP & ""
            mvarEMB_COD = !EMB_COD & ""
            mvarNAT_CRG = !NAT_CRG & ""
            mvarPESO_UND = !PESO_UND & ""
            mvarRALF_COD = !RALF_COD & ""
            mvarAER_DST = !AER_DST & ""
            mvarPESS_CPF = !PESS_CPF & ""
            mvarCRTE_DISP_RESP = !CRTE_DISP_RESP & ""
            mvarCRTE_DISP_DAT = Format(!CRTE_DISP_DAT & "", "DD/MM/YYYY")
            mvarCRTE_DISP_RECF = !CRTE_DISP_RECF & ""
            mvarRESP_NUM = !RESP_NUM & ""
            mvarRESP_NOM = !RESP_NOM & ""
            mvarTIPO_DOC_RESP = !TIPO_DOC_RESP & ""
            mvarRESP_NOM = !RESP_NOM & ""
            mvarDOC_SUFRAMA = !DOC_SUFRAMA & ""
'            mvarCRTE_MAND_JUD = !CRTE_MAND_JUD & ""
'            mvarCRTE_DAT_JUD = Format(!CRTE_DAT_JUD & "", "DD/MM/YYYY")
            mvarRESP_PLACA_VEI = !RESP_PLACA_VEI & ""
            mvarRESP_CRACHA = !RESP_CRACHA & ""
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
Public Function AlterarChave(Ch_CRG_NUM$, Ch_HCRG_NUM$, Ch_NUM_TERMO$) As Integer
   Dim Sql$
   Sql$ = "update CRG_TERMO set "
   Sql$ = Sql$ & " CRG_NUM = " & Aspas(Ch_CRG_NUM)
   Sql$ = Sql$ & " , HCRG_NUM = " & Aspas(Ch_HCRG_NUM)
   Sql$ = Sql$ & " , NUM_TERMO = " & Aspas(Ch_NUM_TERMO)
   Sql$ = Sql$ + " Where "
   Sql$ = Sql$ & " CRG_NUM = " & Aspas(mvarCRG_NUM)
   Sql$ = Sql$ & " and HCRG_NUM = " & Aspas(IIf(mvarHCRG_NUM = "", "' '", mvarHCRG_NUM))
   Sql$ = Sql$ & " and NUM_TERMO = " & Aspas(mvarNUM_TERMO)
   Sql$ = Sql$ & " AER_COD = " & Aspas(mvarAER_COD)
   Call mvardBase.Executa(Sql)
   AlterarChave = mvardBase.CodeSql
End Function
Private Sub Class_Terminate()
  Set mvardBase = Nothing
  Set mPCGs = Nothing
  'Set mAvarias = Nothing
End Sub
Private Sub Class_Initialize()
    mvarPESO_UND = "KG"
    mvarCRTE_TIP_TRAT = "6" 'Local
    mvarNAT_CRG = "NOR"
    mvarCRTE_IDT_MNT = "N"
    mvarCRTE_DSP_ANT = "N"
    mvarCRTE_IDT_BAG = "N"
    mvarCRTE_TIP_DISP = "S"
    mvarCRTE_TOT_PAR = "T"
    mvarCRTE_CLI_PRF = "N"
End Sub
