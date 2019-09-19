VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_NOTAFISCAL"
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

Private mvarNUMNOTA       As String
Private mvarSERIE         As String
Private mvarIDFOR         As Double
Private mvarDTEMISSAO     As String
Private mvarDTSAIDA       As String
Private mvarIDMOEDA       As Double
Private mvarDTRECEB       As String
Private mvarDTVENC        As String
Private mvarCONDPAG       As String
Private mvarVLTOTAL       As Double
Private mvarFORMAPGTO     As String
Private mvarVLOUTROS      As Double
Private mvarDTCADASTRO    As String
Private mvarVLFRETE       As Double
Private mvarVLICMS        As Double
Private mvarVLIPI         As Double
Private mvarCNPJFOR       As String
Private mvarNMFOR         As String
Private mvarVLBASEICMS    As Double
Private mvarTIPONF        As String
Private mvarVLSEGURO      As Double
Private mvarIDSUBFOR      As Double
Private mvarCNPJSUBFOR    As String
Private mvarSUBFORNECEDOR As String
Private mvarSITNF         As String
Private mvarOBS           As String
Private mvarPRZPAG        As Double
Private mvarVLISS         As Double
Private mvarVLIRRF        As Double
Private mvarVLINSS        As Double
Private mvarINSPECIONADO  As String
Public Property Let NUMNOTA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMNOTA <> vData)
   mvarNUMNOTA = vData
End Property
Public Property Get NUMNOTA() As String
   NUMNOTA = mvarNUMNOTA
End Property
Public Property Let SERIE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSERIE <> vData)
   mvarSERIE = vData
End Property
Public Property Get SERIE() As String
   SERIE = mvarSERIE
End Property
Public Property Let IDFOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFOR <> vData)
   mvarIDFOR = vData
End Property
Public Property Get IDFOR() As Double
   IDFOR = mvarIDFOR
End Property
Public Property Let DTEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEMISSAO <> vData)
   mvarDTEMISSAO = vData
End Property
Public Property Get DTEMISSAO() As String
   DTEMISSAO = mvarDTEMISSAO
End Property
Public Property Let DTSAIDA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTSAIDA <> vData)
   mvarDTSAIDA = vData
End Property
Public Property Get DTSAIDA() As String
   DTSAIDA = mvarDTSAIDA
End Property
Public Property Let IDMOEDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMOEDA <> vData)
   mvarIDMOEDA = vData
End Property
Public Property Get IDMOEDA() As Double
   IDMOEDA = mvarIDMOEDA
End Property
Public Property Let DTRECEB(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTRECEB <> vData)
   mvarDTRECEB = vData
End Property
Public Property Get DTRECEB() As String
   DTRECEB = mvarDTRECEB
End Property
Public Property Let DTVENC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTVENC <> vData)
   mvarDTVENC = vData
End Property
Public Property Get DTVENC() As String
   DTVENC = mvarDTVENC
End Property
Public Property Let CONDPAG(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCONDPAG <> vData)
   mvarCONDPAG = vData
End Property
Public Property Get CONDPAG() As String
   CONDPAG = mvarCONDPAG
End Property
Public Property Let VLTOTAL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLTOTAL <> vData)
   mvarVLTOTAL = vData
End Property
Public Property Get VLTOTAL() As Double
   VLTOTAL = mvarVLTOTAL
End Property
Public Property Let FORMAPGTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFORMAPGTO <> vData)
   mvarFORMAPGTO = vData
End Property
Public Property Get FORMAPGTO() As String
   FORMAPGTO = mvarFORMAPGTO
End Property
Public Property Let VLOUTROS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLOUTROS <> vData)
   mvarVLOUTROS = vData
End Property
Public Property Get VLOUTROS() As Double
   VLOUTROS = mvarVLOUTROS
End Property
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let VLFRETE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLFRETE <> vData)
   mvarVLFRETE = vData
End Property
Public Property Get VLFRETE() As Double
   VLFRETE = mvarVLFRETE
End Property
Public Property Let VLICMS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLICMS <> vData)
   mvarVLICMS = vData
End Property
Public Property Get VLICMS() As Double
   VLICMS = mvarVLICMS
End Property
Public Property Let VLIPI(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLIPI <> vData)
   mvarVLIPI = vData
End Property
Public Property Get VLIPI() As Double
   VLIPI = mvarVLIPI
End Property
Public Property Let CNPJFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCNPJFOR <> vData)
   mvarCNPJFOR = vData
End Property
Public Property Get CNPJFOR() As String
   CNPJFOR = mvarCNPJFOR
End Property
Public Property Let NMFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMFOR <> vData)
   mvarNMFOR = vData
End Property
Public Property Get NMFOR() As String
   NMFOR = mvarNMFOR
End Property
Public Property Let VLBASEICMS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLBASEICMS <> vData)
   mvarVLBASEICMS = vData
End Property
Public Property Get VLBASEICMS() As Double
   VLBASEICMS = mvarVLBASEICMS
End Property
Public Property Let TIPONF(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPONF <> vData)
   mvarTIPONF = vData
End Property
Public Property Get TIPONF() As String
   TIPONF = mvarTIPONF
End Property
Public Property Let VLSEGURO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLSEGURO <> vData)
   mvarVLSEGURO = vData
End Property
Public Property Get VLSEGURO() As Double
   VLSEGURO = mvarVLSEGURO
End Property
Public Property Let IDSUBFOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSUBFOR <> vData)
   mvarIDSUBFOR = vData
End Property
Public Property Get IDSUBFOR() As Double
   IDSUBFOR = mvarIDSUBFOR
End Property
Public Property Let CNPJSUBFOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCNPJSUBFOR <> vData)
   mvarCNPJSUBFOR = vData
End Property
Public Property Get CNPJSUBFOR() As String
   CNPJSUBFOR = mvarCNPJSUBFOR
End Property
Public Property Let SUBFORNECEDOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSUBFORNECEDOR <> vData)
   mvarSUBFORNECEDOR = vData
End Property
Public Property Get SUBFORNECEDOR() As String
   SUBFORNECEDOR = mvarSUBFORNECEDOR
End Property
Public Property Let SITNF(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITNF <> vData)
   mvarSITNF = vData
End Property
Public Property Get SITNF() As String
   SITNF = mvarSITNF
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
End Property
Public Property Let PRZPAG(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPRZPAG <> vData)
   mvarPRZPAG = vData
End Property
Public Property Get PRZPAG() As Double
   PRZPAG = mvarPRZPAG
End Property
Public Property Let VLISS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLISS <> vData)
   mvarVLISS = vData
End Property
Public Property Get VLISS() As Double
   VLISS = mvarVLISS
End Property
Public Property Let VLIRRF(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLIRRF <> vData)
   mvarVLIRRF = vData
End Property
Public Property Get VLIRRF() As Double
   VLIRRF = mvarVLIRRF
End Property
Public Property Let VLINSS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLINSS <> vData)
   mvarVLINSS = vData
End Property
Public Property Get VLINSS() As Double
   VLINSS = mvarVLINSS
End Property
Public Property Let INSPECIONADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarINSPECIONADO <> vData)
   mvarINSPECIONADO = vData
End Property
Public Property Get INSPECIONADO() As String
   INSPECIONADO = mvarINSPECIONADO
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
 
   Sql = "Insert Into NOTAFISCAL (NUMNOTA, SERIE, IDFOR, DTEMISSAO" & VbNewLine 
   Sql = Sql & ", DTSAIDA, IDMOEDA, DTRECEB, DTVENC, CONDPAG" & VbNewLine 
   Sql = Sql & ", VLTOTAL, FORMAPGTO, VLOUTROS, DTCADASTRO, VLFRETE" & VbNewLine 
   Sql = Sql & ", VLICMS, VLIPI, CNPJFOR, NMFOR, VLBASEICMS" & VbNewLine 
   Sql = Sql & ", TIPONF, VLSEGURO, IDSUBFOR, CNPJSUBFOR, SUBFORNECEDOR" & VbNewLine 
   Sql = Sql & ", SITNF, OBS, PRZPAG, VLISS, VLIRRF" & VbNewLine 
   Sql = Sql & ", VLINSS, INSPECIONADO) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarNUMNOTA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSERIE) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDFOR),0)+1 From NOTAFISCAL Where NUMNOTA = " & SqlStr(mvarNUMNOTA) & " And SERIE = " & SqlStr(mvarSERIE) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDFOR) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTSAIDA, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDMOEDA) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTRECEB, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTVENC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCONDPAG) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLTOTAL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFORMAPGTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLOUTROS) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLFRETE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLICMS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLIPI) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCNPJFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMFOR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLBASEICMS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPONF) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLSEGURO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDSUBFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCNPJSUBFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSUBFORNECEDOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSITNF) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPRZPAG) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLISS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLIRRF) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLINSS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarINSPECIONADO) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_NUMNOTA, Optional Ch_SERIE, Optional Ch_IDFOR) As String
   Dim Sql As String
 
   Sql = "Delete From NOTAFISCAL" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_NUMNOTA) Then Sql = Sql & " NUMNOTA = " & SqlStr(Cstr(Ch_NUMNOTA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_SERIE) Then Sql = Sql & " SERIE = " & SqlStr(Cstr(Ch_SERIE)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDFOR) Then Sql = Sql & " IDFOR = " & SqlNum(Cstr(Ch_IDFOR)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update NOTAFISCAL Set " & VbNewLine 
   Sql = Sql & " NUMNOTA       = " & SqlStr(mvarNUMNOTA) & VbNewLine 
   Sql = Sql & " , SERIE         = " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & " , IDFOR         = " & SqlNum(mvarIDFOR) & VbNewLine 
   Sql = Sql & " , DTEMISSAO     = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTSAIDA       = " & SqlDate(mvarDTSAIDA, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , IDMOEDA       = " & SqlNum(mvarIDMOEDA) & VbNewLine 
   Sql = Sql & " , DTRECEB       = " & SqlDate(mvarDTRECEB, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTVENC        = " & SqlDate(mvarDTVENC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , CONDPAG       = " & SqlStr(mvarCONDPAG) & VbNewLine 
   Sql = Sql & " , VLTOTAL       = " & SqlNum(mvarVLTOTAL) & VbNewLine 
   Sql = Sql & " , FORMAPGTO     = " & SqlStr(mvarFORMAPGTO) & VbNewLine 
   Sql = Sql & " , VLOUTROS      = " & SqlNum(mvarVLOUTROS) & VbNewLine 
   Sql = Sql & " , DTCADASTRO    = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VLFRETE       = " & SqlNum(mvarVLFRETE) & VbNewLine 
   Sql = Sql & " , VLICMS        = " & SqlNum(mvarVLICMS) & VbNewLine 
   Sql = Sql & " , VLIPI         = " & SqlNum(mvarVLIPI) & VbNewLine 
   Sql = Sql & " , CNPJFOR       = " & SqlStr(mvarCNPJFOR) & VbNewLine 
   Sql = Sql & " , NMFOR         = " & SqlStr(mvarNMFOR) & VbNewLine 
   Sql = Sql & " , VLBASEICMS    = " & SqlNum(mvarVLBASEICMS) & VbNewLine 
   Sql = Sql & " , TIPONF        = " & SqlStr(mvarTIPONF) & VbNewLine 
   Sql = Sql & " , VLSEGURO      = " & SqlNum(mvarVLSEGURO) & VbNewLine 
   Sql = Sql & " , IDSUBFOR      = " & SqlNum(mvarIDSUBFOR) & VbNewLine 
   Sql = Sql & " , CNPJSUBFOR    = " & SqlStr(mvarCNPJSUBFOR) & VbNewLine 
   Sql = Sql & " , SUBFORNECEDOR = " & SqlStr(mvarSUBFORNECEDOR) & VbNewLine 
   Sql = Sql & " , SITNF         = " & SqlStr(mvarSITNF) & VbNewLine 
   Sql = Sql & " , OBS           = " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & " , PRZPAG        = " & SqlNum(mvarPRZPAG) & VbNewLine 
   Sql = Sql & " , VLISS         = " & SqlNum(mvarVLISS) & VbNewLine 
   Sql = Sql & " , VLIRRF        = " & SqlNum(mvarVLIRRF) & VbNewLine 
   Sql = Sql & " , VLINSS        = " & SqlNum(mvarVLINSS) & VbNewLine 
   Sql = Sql & " , INSPECIONADO  = " & SqlStr(mvarINSPECIONADO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'NOTAFISCAL'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " NUMNOTA         = " & SqlStr(mvarNUMNOTA) & VbNewLine 
   Sql = Sql & " And SERIE       = " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & " And IDFOR       = " & SqlNum(mvarIDFOR) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_NUMNOTA, Optional Ch_SERIE, Optional Ch_IDFOR, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select NUMNOTA, SERIE, IDFOR, DTEMISSAO" & VbNewLine 
   Sql = Sql & ", DTSAIDA, IDMOEDA, DTRECEB, DTVENC, CONDPAG" & VbNewLine 
   Sql = Sql & ", VLTOTAL, FORMAPGTO, VLOUTROS, DTCADASTRO, VLFRETE" & VbNewLine 
   Sql = Sql & ", VLICMS, VLIPI, CNPJFOR, NMFOR, VLBASEICMS" & VbNewLine 
   Sql = Sql & ", TIPONF, VLSEGURO, IDSUBFOR, CNPJSUBFOR, SUBFORNECEDOR" & VbNewLine 
   Sql = Sql & ", SITNF, OBS, PRZPAG, VLISS, VLIRRF" & VbNewLine 
   Sql = Sql & ", VLINSS, INSPECIONADO" & VbNewLine 
   Sql = Sql &" From NOTAFISCAL" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_NUMNOTA)  And IsMissing(Ch_SERIE)  And IsMissing(Ch_IDFOR)  Then 
      Sql = Sql & " NUMNOTA = " & SqlStr(mvarNUMNOTA) & " AND " & VbNewLine 
      Sql = Sql & " SERIE = " & SqlStr(mvarSERIE) & " AND " & VbNewLine 
      Sql = Sql & " IDFOR = " & SqlNum(mvarIDFOR) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_NUMNOTA) Then Sql = Sql & " NUMNOTA = " & SqlStr(Cstr(Ch_NUMNOTA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_SERIE) Then Sql = Sql & " SERIE = " & SqlStr(Cstr(Ch_SERIE)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_NUMNOTA, Optional Ch_SERIE, Optional Ch_IDFOR, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_NUMNOTA, Ch_SERIE, Ch_IDFOR, Ch_WHERE, Ch_ORDERBY)
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
      mvarNUMNOTA = !NUMNOTA & ""
      mvarSERIE = !SERIE & ""
      mvarIDFOR = XVal(!IDFOR & "")
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarDTSAIDA = xDate(!DTSAIDA & "", True)
      mvarIDMOEDA = XVal(!IDMOEDA & "")
      mvarDTRECEB = xDate(!DTRECEB & "", True)
      mvarDTVENC = xDate(!DTVENC & "", True)
      mvarCONDPAG = !CONDPAG & ""
      mvarVLTOTAL = XVal(!VLTOTAL & "")
      mvarFORMAPGTO = !FORMAPGTO & ""
      mvarVLOUTROS = XVal(!VLOUTROS & "")
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarVLFRETE = XVal(!VLFRETE & "")
      mvarVLICMS = XVal(!VLICMS & "")
      mvarVLIPI = XVal(!VLIPI & "")
      mvarCNPJFOR = !CNPJFOR & ""
      mvarNMFOR = !NMFOR & ""
      mvarVLBASEICMS = XVal(!VLBASEICMS & "")
      mvarTIPONF = !TIPONF & ""
      mvarVLSEGURO = XVal(!VLSEGURO & "")
      mvarIDSUBFOR = XVal(!IDSUBFOR & "")
      mvarCNPJSUBFOR = !CNPJSUBFOR & ""
      mvarSUBFORNECEDOR = !SUBFORNECEDOR & ""
      mvarSITNF = !SITNF & ""
      mvarOBS = !OBS & ""
      mvarPRZPAG = XVal(!PRZPAG & "")
      mvarVLISS = XVal(!VLISS & "")
      mvarVLIRRF = XVal(!VLIRRF & "")
      mvarVLINSS = XVal(!VLINSS & "")
      mvarINSPECIONADO = !INSPECIONADO & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarNUMNOTA, mvarSERIE, mvarIDFOR), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_NUMNOTA As String, Ch_SERIE As String, Ch_IDFOR As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update NOTAFISCAL Set "
   Sql = Sql & " NUMNOTA = " & SqlStr(Ch_NUMNOTA)
   Sql = Sql & " , SERIE = " & SqlStr(Ch_SERIE)
   Sql = Sql & " , IDFOR = " & SqlNum(Ch_IDFOR)
   Sql = Sql & " Where 
   Sql = Sql & " NUMNOTA = " & SqlStr(mvarNUMNOTA)
   Sql = Sql & " and SERIE = " & SqlStr(mvarSERIE)
   Sql = Sql & " and IDFOR = " & SqlNum(mvarIDFOR)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
