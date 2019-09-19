VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_SPRODUTO"
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

Private mvarIDPROD       As Double
Private mvarIDLOJA       As Double
Private mvarCODPROD      As String
Private mvarNMPROD       As String
Private mvarDSCPROD      As String
Private mvarUNIDCONTROLE As String
Private mvarUNIDCOMPRA   As String
Private mvarUNIDVENDA    As String
Private mvarTXCONVUNID   As Double
Private mvarCODBARRA     As String
Private mvarVLULTCOMPRA  As Double
Private mvarVLMEDIO      As Double
Private mvarVLVENDA      As Double
Private mvarQTDMIN       As Double
Private mvarQTDMAX       As Double
Private mvarQTDCOMPRA    As Double
Private mvarPESOLIQ      As Double
Private mvarPESOBRUTO    As Double
Private mvarCOMPRIMENTO  As Double
Private mvarLARGURA      As Double
Private mvarESPESSURA    As Double
Private mvarEMATPRIMA    As Double
Private mvarEPRODACABADO As Double
Private mvarEICMS        As Double
Private mvarEVENDA       As Double
Private mvarESERVICO     As Double
Private mvarATIVOFIXO    As Double
Private mvarSITPROD      As String
Private mvarDTDESATIVO   As String
Private mvarTAG          As String
Private mvarIDUSU        As String
Private mvarDTCADASTRO   As String
Private mvarALTERSTAMP   As Double
Private mvarTIMESTAMP    As String
Public Property Let IDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROD <> vData)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Double
   IDPROD = mvarIDPROD
End Property
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
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
Public Property Let UNIDCONTROLE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUNIDCONTROLE <> vData)
   mvarUNIDCONTROLE = vData
End Property
Public Property Get UNIDCONTROLE() As String
   UNIDCONTROLE = mvarUNIDCONTROLE
End Property
Public Property Let UNIDCOMPRA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUNIDCOMPRA <> vData)
   mvarUNIDCOMPRA = vData
End Property
Public Property Get UNIDCOMPRA() As String
   UNIDCOMPRA = mvarUNIDCOMPRA
End Property
Public Property Let UNIDVENDA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUNIDVENDA <> vData)
   mvarUNIDVENDA = vData
End Property
Public Property Get UNIDVENDA() As String
   UNIDVENDA = mvarUNIDVENDA
End Property
Public Property Let TXCONVUNID(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTXCONVUNID <> vData)
   mvarTXCONVUNID = vData
End Property
Public Property Get TXCONVUNID() As Double
   TXCONVUNID = mvarTXCONVUNID
End Property
Public Property Let CODBARRA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODBARRA <> vData)
   mvarCODBARRA = vData
End Property
Public Property Get CODBARRA() As String
   CODBARRA = mvarCODBARRA
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
Public Property Let VLVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLVENDA <> vData)
   mvarVLVENDA = vData
End Property
Public Property Get VLVENDA() As Double
   VLVENDA = mvarVLVENDA
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
Public Property Let QTDCOMPRA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDCOMPRA <> vData)
   mvarQTDCOMPRA = vData
End Property
Public Property Get QTDCOMPRA() As Double
   QTDCOMPRA = mvarQTDCOMPRA
End Property
Public Property Let PESOLIQ(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPESOLIQ <> vData)
   mvarPESOLIQ = vData
End Property
Public Property Get PESOLIQ() As Double
   PESOLIQ = mvarPESOLIQ
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
Public Property Let EMATPRIMA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarEMATPRIMA <> vData)
   mvarEMATPRIMA = vData
End Property
Public Property Get EMATPRIMA() As Double
   EMATPRIMA = mvarEMATPRIMA
End Property
Public Property Let EPRODACABADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarEPRODACABADO <> vData)
   mvarEPRODACABADO = vData
End Property
Public Property Get EPRODACABADO() As Double
   EPRODACABADO = mvarEPRODACABADO
End Property
Public Property Let EICMS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarEICMS <> vData)
   mvarEICMS = vData
End Property
Public Property Get EICMS() As Double
   EICMS = mvarEICMS
End Property
Public Property Let EVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarEVENDA <> vData)
   mvarEVENDA = vData
End Property
Public Property Get EVENDA() As Double
   EVENDA = mvarEVENDA
End Property
Public Property Let ESERVICO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarESERVICO <> vData)
   mvarESERVICO = vData
End Property
Public Property Get ESERVICO() As Double
   ESERVICO = mvarESERVICO
End Property
Public Property Let ATIVOFIXO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVOFIXO <> vData)
   mvarATIVOFIXO = vData
End Property
Public Property Get ATIVOFIXO() As Double
   ATIVOFIXO = mvarATIVOFIXO
End Property
Public Property Let SITPROD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITPROD <> vData)
   mvarSITPROD = vData
End Property
Public Property Get SITPROD() As String
   SITPROD = mvarSITPROD
End Property
Public Property Let DTDESATIVO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTDESATIVO <> vData)
   mvarDTDESATIVO = vData
End Property
Public Property Get DTDESATIVO() As String
   DTDESATIVO = mvarDTDESATIVO
End Property
Public Property Let TAG(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTAG <> vData)
   mvarTAG = vData
End Property
Public Property Get TAG() As String
   TAG = mvarTAG
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
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update SPRODUTO Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPROD         = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " And IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True) As String
   Dim Sql As String
 
   Sql = "Insert Into SPRODUTO (IDPROD, IDLOJA, CODPROD, NMPROD" & VbNewLine 
   Sql = Sql & ", DSCPROD, UNIDCONTROLE, UNIDCOMPRA, UNIDVENDA, TXCONVUNID" & VbNewLine 
   Sql = Sql & ", CODBARRA, VLULTCOMPRA, VLMEDIO, VLVENDA, QTDMIN" & VbNewLine 
   Sql = Sql & ", QTDMAX, QTDCOMPRA, PESOLIQ, PESOBRUTO, COMPRIMENTO" & VbNewLine 
   Sql = Sql & ", LARGURA, ESPESSURA, EMATPRIMA, EPRODACABADO, EICMS" & VbNewLine 
   Sql = Sql & ", EVENDA, ESERVICO, ATIVOFIXO, SITPROD, DTDESATIVO" & VbNewLine 
   Sql = Sql & ", TAG, IDUSU, DTCADASTRO, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   If pAutoId Then
      Sql = Sql &  "(Select isNull(Max(IDPROD),0)+1 From SPRODUTO Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql &  SqlNum(mvarIDPROD) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUNIDCONTROLE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUNIDCOMPRA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUNIDVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTXCONVUNID) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODBARRA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLULTCOMPRA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLMEDIO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDMIN) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDMAX) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDCOMPRA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPESOLIQ) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPESOBRUTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarCOMPRIMENTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarLARGURA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarESPESSURA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarEMATPRIMA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarEPRODACABADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarEICMS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarEVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarESERVICO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarATIVOFIXO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSITPROD) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTDESATIVO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTAG) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPROD, Optional Ch_IDLOJA) As String
   Dim Sql As String
 
   Sql = "Delete From SPRODUTO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDPROD) Then Sql = Sql & " IDPROD = " & SqlNum(Cstr(Ch_IDPROD)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   MyDelRow.Query = Sql & ";"
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update SPRODUTO Set " & VbNewLine 
   Sql = Sql & " IDPROD       = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " , IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , CODPROD      = " & SqlStr(mvarCODPROD) & VbNewLine 
   Sql = Sql & " , NMPROD       = " & SqlStr(mvarNMPROD) & VbNewLine 
   Sql = Sql & " , DSCPROD      = " & SqlStr(mvarDSCPROD) & VbNewLine 
   Sql = Sql & " , UNIDCONTROLE = " & SqlStr(mvarUNIDCONTROLE) & VbNewLine 
   Sql = Sql & " , UNIDCOMPRA   = " & SqlStr(mvarUNIDCOMPRA) & VbNewLine 
   Sql = Sql & " , UNIDVENDA    = " & SqlStr(mvarUNIDVENDA) & VbNewLine 
   Sql = Sql & " , TXCONVUNID   = " & SqlNum(mvarTXCONVUNID) & VbNewLine 
   Sql = Sql & " , CODBARRA     = " & SqlStr(mvarCODBARRA) & VbNewLine 
   Sql = Sql & " , VLULTCOMPRA  = " & SqlNum(mvarVLULTCOMPRA) & VbNewLine 
   Sql = Sql & " , VLMEDIO      = " & SqlNum(mvarVLMEDIO) & VbNewLine 
   Sql = Sql & " , VLVENDA      = " & SqlNum(mvarVLVENDA) & VbNewLine 
   Sql = Sql & " , QTDMIN       = " & SqlNum(mvarQTDMIN) & VbNewLine 
   Sql = Sql & " , QTDMAX       = " & SqlNum(mvarQTDMAX) & VbNewLine 
   Sql = Sql & " , QTDCOMPRA    = " & SqlNum(mvarQTDCOMPRA) & VbNewLine 
   Sql = Sql & " , PESOLIQ      = " & SqlNum(mvarPESOLIQ) & VbNewLine 
   Sql = Sql & " , PESOBRUTO    = " & SqlNum(mvarPESOBRUTO) & VbNewLine 
   Sql = Sql & " , COMPRIMENTO  = " & SqlNum(mvarCOMPRIMENTO) & VbNewLine 
   Sql = Sql & " , LARGURA      = " & SqlNum(mvarLARGURA) & VbNewLine 
   Sql = Sql & " , ESPESSURA    = " & SqlNum(mvarESPESSURA) & VbNewLine 
   Sql = Sql & " , EMATPRIMA    = " & SqlNum(mvarEMATPRIMA) & VbNewLine 
   Sql = Sql & " , EPRODACABADO = " & SqlNum(mvarEPRODACABADO) & VbNewLine 
   Sql = Sql & " , EICMS        = " & SqlNum(mvarEICMS) & VbNewLine 
   Sql = Sql & " , EVENDA       = " & SqlNum(mvarEVENDA) & VbNewLine 
   Sql = Sql & " , ESERVICO     = " & SqlNum(mvarESERVICO) & VbNewLine 
   Sql = Sql & " , ATIVOFIXO    = " & SqlNum(mvarATIVOFIXO) & VbNewLine 
   Sql = Sql & " , SITPROD      = " & SqlStr(mvarSITPROD) & VbNewLine 
   Sql = Sql & " , DTDESATIVO   = " & SqlDate(mvarDTDESATIVO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , TAG          = " & SqlStr(mvarTAG) & VbNewLine 
   Sql = Sql & " , IDUSU        = " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & " , DTCADASTRO   = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP   = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP    = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'SPRODUTO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPROD         = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " And IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDPROD, Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPROD, IDLOJA, CODPROD, NMPROD" & VbNewLine 
   Sql = Sql & ", DSCPROD, UNIDCONTROLE, UNIDCOMPRA, UNIDVENDA, TXCONVUNID" & VbNewLine 
   Sql = Sql & ", CODBARRA, VLULTCOMPRA, VLMEDIO, VLVENDA, QTDMIN" & VbNewLine 
   Sql = Sql & ", QTDMAX, QTDCOMPRA, PESOLIQ, PESOBRUTO, COMPRIMENTO" & VbNewLine 
   Sql = Sql & ", LARGURA, ESPESSURA, EMATPRIMA, EPRODACABADO, EICMS" & VbNewLine 
   Sql = Sql & ", EVENDA, ESERVICO, ATIVOFIXO, SITPROD, DTDESATIVO" & VbNewLine 
   Sql = Sql & ", TAG, IDUSU, DTCADASTRO, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From SPRODUTO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPROD)  And IsMissing(Ch_IDLOJA)  Then 
      Sql = Sql & " IDPROD = " & SqlNum(mvarIDPROD) & " AND " & VbNewLine 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDPROD) Then Sql = Sql & " IDPROD = " & SqlNum(Cstr(Ch_IDPROD)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDPROD, Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPROD, Ch_IDLOJA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDLOJA = XVal(!IDLOJA & "")
      mvarCODPROD = !CODPROD & ""
      mvarNMPROD = !NMPROD & ""
      mvarDSCPROD = !DSCPROD & ""
      mvarUNIDCONTROLE = !UNIDCONTROLE & ""
      mvarUNIDCOMPRA = !UNIDCOMPRA & ""
      mvarUNIDVENDA = !UNIDVENDA & ""
      mvarTXCONVUNID = XVal(!TXCONVUNID & "")
      mvarCODBARRA = !CODBARRA & ""
      mvarVLULTCOMPRA = XVal(!VLULTCOMPRA & "")
      mvarVLMEDIO = XVal(!VLMEDIO & "")
      mvarVLVENDA = XVal(!VLVENDA & "")
      mvarQTDMIN = XVal(!QTDMIN & "")
      mvarQTDMAX = XVal(!QTDMAX & "")
      mvarQTDCOMPRA = XVal(!QTDCOMPRA & "")
      mvarPESOLIQ = XVal(!PESOLIQ & "")
      mvarPESOBRUTO = XVal(!PESOBRUTO & "")
      mvarCOMPRIMENTO = XVal(!COMPRIMENTO & "")
      mvarLARGURA = XVal(!LARGURA & "")
      mvarESPESSURA = XVal(!ESPESSURA & "")
      mvarEMATPRIMA = XVal(!EMATPRIMA & "")
      mvarEPRODACABADO = XVal(!EPRODACABADO & "")
      mvarEICMS = XVal(!EICMS & "")
      mvarEVENDA = XVal(!EVENDA & "")
      mvarESERVICO = XVal(!ESERVICO & "")
      mvarATIVOFIXO = XVal(!ATIVOFIXO & "")
      mvarSITPROD = !SITPROD & ""
      mvarDTDESATIVO = xDate(!DTDESATIVO & "", True)
      mvarTAG = !TAG & ""
      mvarIDUSU = !IDUSU & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPROD, mvarIDLOJA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDPROD As String, Ch_IDLOJA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update SPRODUTO Set "
   Sql = Sql & " IDPROD = " & SqlNum(Ch_IDPROD)
   Sql = Sql & " , IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " Where 
   Sql = Sql & " IDPROD = " & SqlNum(mvarIDPROD)
   Sql = Sql & " and IDLOJA = " & SqlNum(mvarIDLOJA)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
