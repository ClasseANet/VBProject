VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_OCLIENTE"
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
Private mvarIDCLIENTE    As Double
Private mvarIDCOLIGADA   As Double
Private mvarIDLOJA0      As Double
Private mvarIDTPCONHEC   As Double
Private mvarIDFUNC       As Double
Private mvarDSCTPCONHEC  As String
Private mvarNOME         As String
Private mvarDTNASC       As String
Private mvarREGISTRO     As String
Private mvarPROFISSAO    As String
Private mvarFOTOTIPO     As Double
Private mvarSEXO         As String
Private mvarTEL1         As String
Private mvarTEL2         As String
Private mvarFAX          As String
Private mvarEMAIL        As String
Private mvarENDERECO     As String
Private mvarBAIRRO       As String
Private mvarCIDADE       As String
Private mvarESTADO       As String
Private mvarPAIS         As String
Private mvarCEP          As String
Private mvarEMPRESA      As String
Private mvarCARGO        As String
Private mvarOBS          As String
Private mvarATIVO        As Double
Private mvarISENTO       As Double
Private mvarMOTIVOINAT   As String
Private mvarDTCADASTRO   As String
Private mvarFLGMARKETING As Double
Private mvarFLGAGENDA    As Double
Private mvarTIMESTAMP    As String
Private mvarALTERSTAMP   As Double
Private mvarNFE          As Double
Private mvarOPER1        As String
Private mvarOPER2        As String
Private mvarOPERF        As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDCOLIGADA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCOLIGADA <> vData)
   mvarIDCOLIGADA = vData
End Property
Public Property Get IDCOLIGADA() As Double
   IDCOLIGADA = mvarIDCOLIGADA
End Property
Public Property Let IDLOJA0(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA0 <> vData)
   mvarIDLOJA0 = vData
End Property
Public Property Get IDLOJA0() As Double
   IDLOJA0 = mvarIDLOJA0
End Property
Public Property Let IDTPCONHEC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPCONHEC <> vData)
   mvarIDTPCONHEC = vData
End Property
Public Property Get IDTPCONHEC() As Double
   IDTPCONHEC = mvarIDTPCONHEC
End Property
Public Property Let IDFUNC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFUNC <> vData)
   mvarIDFUNC = vData
End Property
Public Property Get IDFUNC() As Double
   IDFUNC = mvarIDFUNC
End Property
Public Property Let DSCTPCONHEC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCTPCONHEC <> vData)
   mvarDSCTPCONHEC = vData
End Property
Public Property Get DSCTPCONHEC() As String
   DSCTPCONHEC = mvarDSCTPCONHEC
End Property
Public Property Let NOME(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNOME <> vData)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Public Property Let DTNASC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTNASC <> vData)
   mvarDTNASC = vData
End Property
Public Property Get DTNASC() As String
   DTNASC = mvarDTNASC
End Property
Public Property Let REGISTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarREGISTRO <> vData)
   mvarREGISTRO = vData
End Property
Public Property Get REGISTRO() As String
   REGISTRO = mvarREGISTRO
End Property
Public Property Let PROFISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPROFISSAO <> vData)
   mvarPROFISSAO = vData
End Property
Public Property Get PROFISSAO() As String
   PROFISSAO = mvarPROFISSAO
End Property
Public Property Let FOTOTIPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFOTOTIPO <> vData)
   mvarFOTOTIPO = vData
End Property
Public Property Get FOTOTIPO() As Double
   FOTOTIPO = mvarFOTOTIPO
End Property
Public Property Let SEXO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSEXO <> vData)
   mvarSEXO = vData
End Property
Public Property Get SEXO() As String
   SEXO = mvarSEXO
End Property
Public Property Let TEL1(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTEL1 <> vData)
   mvarTEL1 = vData
End Property
Public Property Get TEL1() As String
   TEL1 = mvarTEL1
End Property
Public Property Let TEL2(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTEL2 <> vData)
   mvarTEL2 = vData
End Property
Public Property Get TEL2() As String
   TEL2 = mvarTEL2
End Property
Public Property Let FAX(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFAX <> vData)
   mvarFAX = vData
End Property
Public Property Get FAX() As String
   FAX = mvarFAX
End Property
Public Property Let EMAIL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEMAIL <> vData)
   mvarEMAIL = vData
End Property
Public Property Get EMAIL() As String
   EMAIL = mvarEMAIL
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
Public Property Let ESTADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarESTADO <> vData)
   mvarESTADO = vData
End Property
Public Property Get ESTADO() As String
   ESTADO = mvarESTADO
End Property
Public Property Let PAIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPAIS <> vData)
   mvarPAIS = vData
End Property
Public Property Get PAIS() As String
   PAIS = mvarPAIS
End Property
Public Property Let CEP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCEP <> vData)
   mvarCEP = vData
End Property
Public Property Get CEP() As String
   CEP = mvarCEP
End Property
Public Property Let EMPRESA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEMPRESA <> vData)
   mvarEMPRESA = vData
End Property
Public Property Get EMPRESA() As String
   EMPRESA = mvarEMPRESA
End Property
Public Property Let CARGO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCARGO <> vData)
   mvarCARGO = vData
End Property
Public Property Get CARGO() As String
   CARGO = mvarCARGO
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
End Property
Public Property Let ATIVO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVO <> vData)
   mvarATIVO = vData
End Property
Public Property Get ATIVO() As Double
   ATIVO = mvarATIVO
End Property
Public Property Let ISENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarISENTO <> vData)
   mvarISENTO = vData
End Property
Public Property Get ISENTO() As Double
   ISENTO = mvarISENTO
End Property
Public Property Let MOTIVOINAT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMOTIVOINAT <> vData)
   mvarMOTIVOINAT = vData
End Property
Public Property Get MOTIVOINAT() As String
   MOTIVOINAT = mvarMOTIVOINAT
End Property
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let FLGMARKETING(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGMARKETING <> vData)
   mvarFLGMARKETING = vData
End Property
Public Property Get FLGMARKETING() As Double
   FLGMARKETING = mvarFLGMARKETING
End Property
Public Property Let FLGAGENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGAGENDA <> vData)
   mvarFLGAGENDA = vData
End Property
Public Property Get FLGAGENDA() As Double
   FLGAGENDA = mvarFLGAGENDA
End Property
Public Property Let TIMESTAMP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIMESTAMP <> vData)
   mvarTIMESTAMP = vData
End Property
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OCLIENTE Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDCLIENTE  = " & SqlNum(mvarIDCLIENTE) & vbNewLine
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
Public Property Let NFE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNFE <> vData)
   mvarNFE = vData
End Property
Public Property Get NFE() As Double
   NFE = mvarNFE
End Property
Public Property Let OPER1(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOPER1 <> vData)
   mvarOPER1 = vData
End Property
Public Property Get OPER1() As String
   OPER1 = mvarOPER1
End Property
Public Property Let OPER2(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOPER2 <> vData)
   mvarOPER2 = vData
End Property
Public Property Get OPER2() As String
   OPER2 = mvarOPER2
End Property
Public Property Let OPERF(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOPERF <> vData)
   mvarOPERF = vData
End Property
Public Property Get OPERF() As String
   OPERF = mvarOPERF
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
Public Property Get QryInsert(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into OCLIENTE (IDLOJA, IDCLIENTE, IDCOLIGADA, IDLOJA0" & vbNewLine
   Sql = Sql & ", IDTPCONHEC, IDFUNC, DSCTPCONHEC, NOME, DTNASC" & vbNewLine
   Sql = Sql & ", REGISTRO, PROFISSAO, FOTOTIPO, SEXO, TEL1, TEL2" & vbNewLine
   Sql = Sql & ", FAX, EMAIL, ENDERECO, BAIRRO, CIDADE" & vbNewLine
   Sql = Sql & ", ESTADO, PAIS, CEP, EMPRESA, CARGO" & vbNewLine
   Sql = Sql & ", OBS, ATIVO, ISENTO, MOTIVOINAT, DTCADASTRO" & vbNewLine
   Sql = Sql & ", FLGMARKETING, FLGAGENDA, TIMESTAMP, ALTERSTAMP, NFE" & vbNewLine
   Sql = Sql & ", OPER1, OPER2, OPERF) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDCLIENTE),0)+1 From OCLIENTE Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & vbNewLine
   End If
   Sql = Sql & ", " & SqlNum(mvarIDCOLIGADA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDLOJA0) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDTPCONHEC = 0, "Null", SqlNum(mvarIDTPCONHEC)) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDFUNC = 0, "Null", SqlNum(mvarIDFUNC)) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarDSCTPCONHEC) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarNOME) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTNASC, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarREGISTRO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarPROFISSAO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFOTOTIPO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarSEXO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTEL1) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTEL2) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarFAX) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarEMAIL) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarENDERECO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarBAIRRO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCIDADE) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarESTADO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarPAIS) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCEP) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarEMPRESA) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCARGO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarOBS) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarATIVO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarISENTO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarMOTIVOINAT) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGMARKETING) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGAGENDA) & vbNewLine
   If pSinc Then
      Sql = Sql & ", " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If
   Sql = Sql & ", " & SqlNum(mvarNFE) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarOPER1) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarOPER2) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarOPERF) & vbNewLine

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDCLIENTE, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OCLIENTE" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDCLIENTE) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDCLIENTE = " & SqlNum(mvarIDCLIENTE) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDCLIENTE) Then Sql = Sql & " IDCLIENTE = " & SqlNum(CStr(Ch_IDCLIENTE)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = ""
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|IDCOLIGADA = " & mvarIDCOLIGADA
   sTag = sTag & "|IDLOJA0 = " & mvarIDLOJA0
   sTag = sTag & "|IDTPCONHEC = " & mvarIDTPCONHEC
   sTag = sTag & "|IDFUNC = " & mvarIDFUNC
   sTag = sTag & "|DSCTPCONHEC = " & mvarDSCTPCONHEC
   sTag = sTag & "|NOME = " & mvarNOME
   sTag = sTag & "|DTNASC = " & mvarDTNASC
   sTag = sTag & "|REGISTRO = " & mvarREGISTRO
   sTag = sTag & "|PROFISSAO = " & mvarPROFISSAO
   sTag = sTag & "|FOTOTIPO = " & mvarFOTOTIPO
   sTag = sTag & "|SEXO = " & mvarSEXO
   sTag = sTag & "|TEL1 = " & mvarTEL1
   sTag = sTag & "|TEL2 = " & mvarTEL2
   sTag = sTag & "|FAX = " & mvarFAX
   sTag = sTag & "|EMAIL = " & mvarEMAIL
   sTag = sTag & "|ENDERECO = " & mvarENDERECO
   sTag = sTag & "|BAIRRO = " & mvarBAIRRO
   sTag = sTag & "|CIDADE = " & mvarCIDADE
   sTag = sTag & "|ESTADO = " & mvarESTADO
   sTag = sTag & "|PAIS = " & mvarPAIS
   sTag = sTag & "|CEP = " & mvarCEP
   sTag = sTag & "|EMPRESA = " & mvarEMPRESA
   sTag = sTag & "|CARGO = " & mvarCARGO
   sTag = sTag & "|OBS = " & mvarOBS
   sTag = sTag & "|ATIVO = " & mvarATIVO
   sTag = sTag & "|ISENTO = " & mvarISENTO
   sTag = sTag & "|MOTIVOINAT = " & mvarMOTIVOINAT
   sTag = sTag & "|DTCADASTRO = " & mvarDTCADASTRO
   sTag = sTag & "|FLGMARKETING = " & mvarFLGMARKETING
   sTag = sTag & "|FLGAGENDA = " & mvarFLGAGENDA
   sTag = sTag & "|NFE = " & mvarNFE
   sTag = sTag & "|OPER1 = " & mvarOPER1
   sTag = sTag & "|OPER2 = " & mvarOPER2
   sTag = sTag & "|OPERF = " & mvarOPERF
   sTag = sTag & "|Where = " & IIf(IsMissing(Ch_WHERE), "", Ch_WHERE)
   sTag = sTag & "|"

   MyDelRow.QUERY = Sql & ";"
   MyDelRow.TAG = sTag
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update OCLIENTE Set " & vbNewLine
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDCLIENTE    = " & SqlNum(mvarIDCLIENTE) & vbNewLine
   Sql = Sql & " , IDCOLIGADA   = " & SqlNum(mvarIDCOLIGADA) & vbNewLine
   Sql = Sql & " , IDLOJA0      = " & SqlNum(mvarIDLOJA0) & vbNewLine
   Sql = Sql & " , IDTPCONHEC   = " & IIf(mvarIDTPCONHEC = 0, "Null", SqlNum(mvarIDTPCONHEC)) & vbNewLine
   Sql = Sql & " , IDFUNC       = " & IIf(mvarIDFUNC = 0, "Null", SqlNum(mvarIDFUNC)) & vbNewLine
   Sql = Sql & " , DSCTPCONHEC  = " & SqlStr(mvarDSCTPCONHEC) & vbNewLine
   Sql = Sql & " , NOME         = " & SqlStr(mvarNOME) & vbNewLine
   Sql = Sql & " , DTNASC       = " & SqlDate(mvarDTNASC, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , REGISTRO     = " & SqlStr(mvarREGISTRO) & vbNewLine
   Sql = Sql & " , PROFISSAO     = " & SqlStr(mvarPROFISSAO) & vbNewLine
   Sql = Sql & " , FOTOTIPO     = " & SqlNum(mvarFOTOTIPO) & vbNewLine
   Sql = Sql & " , SEXO         = " & SqlStr(mvarSEXO) & vbNewLine
   Sql = Sql & " , TEL1         = " & SqlStr(mvarTEL1) & vbNewLine
   Sql = Sql & " , TEL2         = " & SqlStr(mvarTEL2) & vbNewLine
   Sql = Sql & " , FAX          = " & SqlStr(mvarFAX) & vbNewLine
   Sql = Sql & " , EMAIL        = " & SqlStr(mvarEMAIL) & vbNewLine
   Sql = Sql & " , ENDERECO     = " & SqlStr(mvarENDERECO) & vbNewLine
   Sql = Sql & " , BAIRRO       = " & SqlStr(mvarBAIRRO) & vbNewLine
   Sql = Sql & " , CIDADE       = " & SqlStr(mvarCIDADE) & vbNewLine
   Sql = Sql & " , ESTADO       = " & SqlStr(mvarESTADO) & vbNewLine
   Sql = Sql & " , PAIS         = " & SqlStr(mvarPAIS) & vbNewLine
   Sql = Sql & " , CEP          = " & SqlStr(mvarCEP) & vbNewLine
   Sql = Sql & " , EMPRESA      = " & SqlStr(mvarEMPRESA) & vbNewLine
   Sql = Sql & " , CARGO        = " & SqlStr(mvarCARGO) & vbNewLine
   Sql = Sql & " , OBS          = " & SqlStr(mvarOBS) & vbNewLine
   Sql = Sql & " , ATIVO        = " & SqlNum(mvarATIVO) & vbNewLine
   Sql = Sql & " , ISENTO       = " & SqlNum(mvarISENTO) & vbNewLine
   Sql = Sql & " , MOTIVOINAT   = " & SqlStr(mvarMOTIVOINAT) & vbNewLine
   Sql = Sql & " , DTCADASTRO   = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , FLGMARKETING = " & SqlNum(mvarFLGMARKETING) & vbNewLine
   Sql = Sql & " , FLGAGENDA    = " & SqlNum(mvarFLGAGENDA) & vbNewLine
   If pSinc Then
      Sql = Sql & ", TIMESTAMP    = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP    = GetDate()" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", ALTERSTAMP   = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP   = 1" & vbNewLine
   End If
   Sql = Sql & " , NFE          = " & SqlNum(mvarNFE) & vbNewLine
   Sql = Sql & " , OPER1        = " & SqlStr(mvarOPER1) & vbNewLine
   Sql = Sql & " , OPER2        = " & SqlStr(mvarOPER2) & vbNewLine
   Sql = Sql & " , OPERF        = " & SqlStr(mvarOPERF) & vbNewLine
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OCLIENTE'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDCLIENTE  = " & SqlNum(mvarIDCLIENTE) & vbNewLine

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & vbNewLine
   Sql = Sql & Me.QryUpDate(pSinc:=pSinc)
   Sql = Sql & " Else " & vbNewLine
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId, pSinc:=pSinc)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDCLIENTE, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDCLIENTE, IDCOLIGADA, IDLOJA0" & vbNewLine
   Sql = Sql & ", IDTPCONHEC, IDFUNC, DSCTPCONHEC, NOME, DTNASC" & vbNewLine
   Sql = Sql & ", REGISTRO, PROFISSAO, FOTOTIPO, SEXO, TEL1, TEL2" & vbNewLine
   Sql = Sql & ", FAX, EMAIL, ENDERECO, BAIRRO, CIDADE" & vbNewLine
   Sql = Sql & ", ESTADO, PAIS, CEP, EMPRESA, CARGO" & vbNewLine
   Sql = Sql & ", OBS, ATIVO, ISENTO, MOTIVOINAT, DTCADASTRO" & vbNewLine
   Sql = Sql & ", FLGMARKETING, FLGAGENDA, TIMESTAMP, ALTERSTAMP, NFE" & vbNewLine
   Sql = Sql & ", OPER1, OPER2, OPERF" & vbNewLine
   Sql = Sql & " From OCLIENTE" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDCLIENTE) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDCLIENTE = " & SqlNum(mvarIDCLIENTE) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDCLIENTE) Then Sql = Sql & " IDCLIENTE = " & SqlNum(CStr(Ch_IDCLIENTE)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDCLIENTE, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDCLIENTE, Ch_WHERE, Ch_ORDERBY)
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
    Call ShowError(Sql)
    Pesquisar = False
End Function
Public Sub Popula(Optional pRcSet)
   If IsMissing(pRcSet) Then Set pRcSet = mvarRS
   With pRcSet
      mvarIDLOJA = xVal(!IDLOJA & "")
      mvarIDCLIENTE = xVal(!IDCLIENTE & "")
      mvarIDCOLIGADA = xVal(!IDCOLIGADA & "")
      mvarIDLOJA0 = xVal(!IDLOJA0 & "")
      mvarIDTPCONHEC = xVal(!IDTPCONHEC & "")
      mvarIDFUNC = xVal(!IDFUNC & "")
      mvarDSCTPCONHEC = !DSCTPCONHEC & ""
      mvarNOME = !NOME & ""
      mvarDTNASC = xDate(!DTNASC & "", True)
      mvarREGISTRO = !REGISTRO & ""
      mvarPROFISSAO = !PROFISSAO & ""
      mvarFOTOTIPO = xVal(!FOTOTIPO & "")
      mvarSEXO = !SEXO & ""
      mvarTEL1 = !TEL1 & ""
      mvarTEL2 = !TEL2 & ""
      mvarFAX = !FAX & ""
      mvarEMAIL = !EMAIL & ""
      mvarENDERECO = !ENDERECO & ""
      mvarBAIRRO = !BAIRRO & ""
      mvarCIDADE = !CIDADE & ""
      mvarESTADO = !ESTADO & ""
      mvarPAIS = !PAIS & ""
      mvarCEP = !CEP & ""
      mvarEMPRESA = !EMPRESA & ""
      mvarCARGO = !CARGO & ""
      mvarOBS = !OBS & ""
      mvarATIVO = xVal(!ATIVO & "")
      mvarISENTO = xVal(!ISENTO & "")
      mvarMOTIVOINAT = !MOTIVOINAT & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarFLGMARKETING = xVal(!FLGMARKETING & "")
      mvarFLGAGENDA = xVal(!FLGAGENDA & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarNFE = xVal(!NFE & "")
      mvarOPER1 = !OPER1 & ""
      mvarOPER2 = !OPER2 & ""
      mvarOPERF = !OPERF & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDCLIENTE = 0
   mvarIDCOLIGADA = 0
   mvarIDLOJA0 = 0
   mvarIDTPCONHEC = 0
   mvarIDFUNC = 0
   mvarDSCTPCONHEC = ""
   mvarNOME = ""
   mvarDTNASC = ""
   mvarREGISTRO = ""
   mvarPROFISSAO = ""
   mvarFOTOTIPO = 0
   mvarSEXO = ""
   mvarTEL1 = ""
   mvarTEL2 = ""
   mvarFAX = ""
   mvarEMAIL = ""
   mvarENDERECO = ""
   mvarBAIRRO = ""
   mvarCIDADE = ""
   mvarESTADO = ""
   mvarPAIS = ""
   mvarCEP = ""
   mvarEMPRESA = ""
   mvarCARGO = ""
   mvarOBS = ""
   mvarATIVO = 0
   mvarISENTO = 0
   mvarMOTIVOINAT = ""
   mvarDTCADASTRO = ""
   mvarFLGMARKETING = 0
   mvarFLGAGENDA = 0
   mvarTIMESTAMP = ""
   mvarALTERSTAMP = 0
   mvarNFE = 0
   mvarOPER1 = ""
   mvarOPER2 = ""
   mvarOPERF = ""

   On Error Resume Next
   Call Class_Initialize
End Sub
Public Function Salvar(Optional ComCOMMIT As Boolean = True, Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId, pSinc:=pSinc), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId As Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDCLIENTE), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDCLIENTE As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OCLIENTE Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDCLIENTE = " & SqlNum(Ch_IDCLIENTE)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDCLIENTE = " & SqlNum(mvarIDCLIENTE)
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
