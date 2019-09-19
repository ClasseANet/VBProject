VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
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
Private mvarDSCTPCONHEC  As String
Private mvarNOME         As String
Private mvarDTNASC       As String
Private mvarREGISTRO     As String
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
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OCLIENTE Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDCLIENTE  = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
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
 
   Sql = "Insert Into OCLIENTE (IDLOJA, IDCLIENTE, IDCOLIGADA, IDLOJA0" & VbNewLine 
   Sql = Sql & ", IDTPCONHEC, DSCTPCONHEC, NOME, DTNASC, REGISTRO" & VbNewLine 
   Sql = Sql & ", FOTOTIPO, SEXO, TEL1, TEL2, FAX" & VbNewLine 
   Sql = Sql & ", EMAIL, ENDERECO, BAIRRO, CIDADE, ESTADO" & VbNewLine 
   Sql = Sql & ", PAIS, CEP, EMPRESA, CARGO, OBS" & VbNewLine 
   Sql = Sql & ", ATIVO, ISENTO, MOTIVOINAT, DTCADASTRO, FLGMARKETING" & VbNewLine 
   Sql = Sql & ", FLGAGENDA, TIMESTAMP, ALTERSTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDCLIENTE),0)+1 From OCLIENTE Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDLOJA0) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTPCONHEC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCTPCONHEC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTNASC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarREGISTRO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFOTOTIPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSEXO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTEL1) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTEL2) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFAX) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarENDERECO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarESTADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCEP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMPRESA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCARGO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarATIVO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarISENTO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMOTIVOINAT) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGMARKETING) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGAGENDA) & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDCLIENTE) As String
   Dim Sql As String
 
   Sql = "Delete From OCLIENTE" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDCLIENTE) Then Sql = Sql & " IDCLIENTE = " & SqlNum(Cstr(Ch_IDCLIENTE)) & " AND " & VbNewLine 
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
 
   Sql = "Update OCLIENTE Set " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDCLIENTE    = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & " , IDCOLIGADA   = " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & " , IDLOJA0      = " & SqlNum(mvarIDLOJA0) & VbNewLine 
   Sql = Sql & " , IDTPCONHEC   = " & SqlNum(mvarIDTPCONHEC) & VbNewLine 
   Sql = Sql & " , DSCTPCONHEC  = " & SqlStr(mvarDSCTPCONHEC) & VbNewLine 
   Sql = Sql & " , NOME         = " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & " , DTNASC       = " & SqlDate(mvarDTNASC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , REGISTRO     = " & SqlStr(mvarREGISTRO) & VbNewLine 
   Sql = Sql & " , FOTOTIPO     = " & SqlNum(mvarFOTOTIPO) & VbNewLine 
   Sql = Sql & " , SEXO         = " & SqlStr(mvarSEXO) & VbNewLine 
   Sql = Sql & " , TEL1         = " & SqlStr(mvarTEL1) & VbNewLine 
   Sql = Sql & " , TEL2         = " & SqlStr(mvarTEL2) & VbNewLine 
   Sql = Sql & " , FAX          = " & SqlStr(mvarFAX) & VbNewLine 
   Sql = Sql & " , EMAIL        = " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & " , ENDERECO     = " & SqlStr(mvarENDERECO) & VbNewLine 
   Sql = Sql & " , BAIRRO       = " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & " , CIDADE       = " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & " , ESTADO       = " & SqlStr(mvarESTADO) & VbNewLine 
   Sql = Sql & " , PAIS         = " & SqlStr(mvarPAIS) & VbNewLine 
   Sql = Sql & " , CEP          = " & SqlStr(mvarCEP) & VbNewLine 
   Sql = Sql & " , EMPRESA      = " & SqlStr(mvarEMPRESA) & VbNewLine 
   Sql = Sql & " , CARGO        = " & SqlStr(mvarCARGO) & VbNewLine 
   Sql = Sql & " , OBS          = " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & " , ATIVO        = " & SqlNum(mvarATIVO) & VbNewLine 
   Sql = Sql & " , ISENTO       = " & SqlNum(mvarISENTO) & VbNewLine 
   Sql = Sql & " , MOTIVOINAT   = " & SqlStr(mvarMOTIVOINAT) & VbNewLine 
   Sql = Sql & " , DTCADASTRO   = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , FLGMARKETING = " & SqlNum(mvarFLGMARKETING) & VbNewLine 
   Sql = Sql & " , FLGAGENDA    = " & SqlNum(mvarFLGAGENDA) & VbNewLine 
   Sql = Sql & " , TIMESTAMP    = " & "GetDate()" & VbNewLine 
   Sql = Sql & " , ALTERSTAMP   = " & "1" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OCLIENTE'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDCLIENTE  = " & SqlNum(mvarIDCLIENTE) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDCLIENTE, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDCLIENTE, IDCOLIGADA, IDLOJA0" & VbNewLine 
   Sql = Sql & ", IDTPCONHEC, DSCTPCONHEC, NOME, DTNASC, REGISTRO" & VbNewLine 
   Sql = Sql & ", FOTOTIPO, SEXO, TEL1, TEL2, FAX" & VbNewLine 
   Sql = Sql & ", EMAIL, ENDERECO, BAIRRO, CIDADE, ESTADO" & VbNewLine 
   Sql = Sql & ", PAIS, CEP, EMPRESA, CARGO, OBS" & VbNewLine 
   Sql = Sql & ", ATIVO, ISENTO, MOTIVOINAT, DTCADASTRO, FLGMARKETING" & VbNewLine 
   Sql = Sql & ", FLGAGENDA, TIMESTAMP, ALTERSTAMP" & VbNewLine 
   Sql = Sql &" From OCLIENTE" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDCLIENTE)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDCLIENTE = " & SqlNum(mvarIDCLIENTE) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCLIENTE) Then Sql = Sql & " IDCLIENTE = " & SqlNum(Cstr(Ch_IDCLIENTE)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDCLIENTE, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDCLIENTE, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
      mvarIDCOLIGADA = XVal(!IDCOLIGADA & "")
      mvarIDLOJA0 = XVal(!IDLOJA0 & "")
      mvarIDTPCONHEC = XVal(!IDTPCONHEC & "")
      mvarDSCTPCONHEC = !DSCTPCONHEC & ""
      mvarNOME = !NOME & ""
      mvarDTNASC = xDate(!DTNASC & "", True)
      mvarREGISTRO = !REGISTRO & ""
      mvarFOTOTIPO = XVal(!FOTOTIPO & "")
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
      mvarATIVO = XVal(!ATIVO & "")
      mvarISENTO = XVal(!ISENTO & "")
      mvarMOTIVOINAT = !MOTIVOINAT & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarFLGMARKETING = XVal(!FLGMARKETING & "")
      mvarFLGAGENDA = XVal(!FLGAGENDA & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDCLIENTE), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDCLIENTE As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OCLIENTE Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDCLIENTE = " & SqlNum(Ch_IDCLIENTE)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDCLIENTE = " & SqlNum(mvarIDCLIENTE)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
