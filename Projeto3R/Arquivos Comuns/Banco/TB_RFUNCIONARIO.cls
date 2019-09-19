VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_RFUNCIONARIO"
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

Private mvarIDLOJA         As Double
Private mvarIDFUNCIONARIO  As Double
Private mvarIDCOLIGADA     As Double
Private mvarIDLOJA0        As Double
Private mvarIDFINGER       As Double
Private mvarCHAPA          As String
Private mvarNOME           As String
Private mvarDTADMISSAO     As String
Private mvarDTDEMISSAO     As String
Private mvarFLGCERTIFICADO As Double
Private mvarSITFUNC        As String
Private mvarCOMPROD        As Double
Private mvarVLCOMPROD      As Double
Private mvarTPCOMPROD      As Double
Private mvarCOMSERV        As Double
Private mvarVLCOMSERV      As Double
Private mvarTPCOMSERV      As Double
Private mvarOBS            As String
Private mvarTELEFONE       As String
Private mvarCELULAR        As String
Private mvarEMAIL          As String
Private mvarENDERECO       As String
Private mvarBAIRRO         As String
Private mvarCIDADE         As String
Private mvarESTADO         As String
Private mvarCEP            As String
Private mvarPAIS           As String
Private mvarDTNASC         As String
Private mvarSALARIO        As Double
Private mvarSENHA          As String
Private mvarDTCADASTRO     As String
Private mvarALTERSTAMP     As Double
Private mvarTIMESTAMP      As String
Private mvarDIAFOLGA       As Double
Private mvarFLGVENDA       As Double
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDFUNCIONARIO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFUNCIONARIO <> vData)
   mvarIDFUNCIONARIO = vData
End Property
Public Property Get IDFUNCIONARIO() As Double
   IDFUNCIONARIO = mvarIDFUNCIONARIO
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
Public Property Let IDFINGER(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFINGER <> vData)
   mvarIDFINGER = vData
End Property
Public Property Get IDFINGER() As Double
   IDFINGER = mvarIDFINGER
End Property
Public Property Let CHAPA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCHAPA <> vData)
   mvarCHAPA = vData
End Property
Public Property Get CHAPA() As String
   CHAPA = mvarCHAPA
End Property
Public Property Let NOME(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNOME <> vData)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Public Property Let DTADMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTADMISSAO <> vData)
   mvarDTADMISSAO = vData
End Property
Public Property Get DTADMISSAO() As String
   DTADMISSAO = mvarDTADMISSAO
End Property
Public Property Let DTDEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTDEMISSAO <> vData)
   mvarDTDEMISSAO = vData
End Property
Public Property Get DTDEMISSAO() As String
   DTDEMISSAO = mvarDTDEMISSAO
End Property
Public Property Let FLGCERTIFICADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGCERTIFICADO <> vData)
   mvarFLGCERTIFICADO = vData
End Property
Public Property Get FLGCERTIFICADO() As Double
   FLGCERTIFICADO = mvarFLGCERTIFICADO
End Property
Public Property Let SITFUNC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITFUNC <> vData)
   mvarSITFUNC = vData
End Property
Public Property Get SITFUNC() As String
   SITFUNC = mvarSITFUNC
End Property
Public Property Let COMPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarCOMPROD <> vData)
   mvarCOMPROD = vData
End Property
Public Property Get COMPROD() As Double
   COMPROD = mvarCOMPROD
End Property
Public Property Let VLCOMPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLCOMPROD <> vData)
   mvarVLCOMPROD = vData
End Property
Public Property Get VLCOMPROD() As Double
   VLCOMPROD = mvarVLCOMPROD
End Property
Public Property Let TPCOMPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTPCOMPROD <> vData)
   mvarTPCOMPROD = vData
End Property
Public Property Get TPCOMPROD() As Double
   TPCOMPROD = mvarTPCOMPROD
End Property
Public Property Let COMSERV(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarCOMSERV <> vData)
   mvarCOMSERV = vData
End Property
Public Property Get COMSERV() As Double
   COMSERV = mvarCOMSERV
End Property
Public Property Let VLCOMSERV(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLCOMSERV <> vData)
   mvarVLCOMSERV = vData
End Property
Public Property Get VLCOMSERV() As Double
   VLCOMSERV = mvarVLCOMSERV
End Property
Public Property Let TPCOMSERV(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTPCOMSERV <> vData)
   mvarTPCOMSERV = vData
End Property
Public Property Get TPCOMSERV() As Double
   TPCOMSERV = mvarTPCOMSERV
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
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
Public Property Let CEP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCEP <> vData)
   mvarCEP = vData
End Property
Public Property Get CEP() As String
   CEP = mvarCEP
End Property
Public Property Let PAIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPAIS <> vData)
   mvarPAIS = vData
End Property
Public Property Get PAIS() As String
   PAIS = mvarPAIS
End Property
Public Property Let DTNASC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTNASC <> vData)
   mvarDTNASC = vData
End Property
Public Property Get DTNASC() As String
   DTNASC = mvarDTNASC
End Property
Public Property Let SALARIO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSALARIO <> vData)
   mvarSALARIO = vData
End Property
Public Property Get SALARIO() As Double
   SALARIO = mvarSALARIO
End Property
Public Property Let SENHA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSENHA <> vData)
   mvarSENHA = vData
End Property
Public Property Get SENHA() As String
   SENHA = mvarSENHA
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
 
   Sql = "Update RFUNCIONARIO Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & VbNewLine 
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
Public Property Let TIMESTAMP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIMESTAMP <> vData)
   mvarTIMESTAMP = vData
End Property
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
End Property
Public Property Let DIAFOLGA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarDIAFOLGA <> vData)
   mvarDIAFOLGA = vData
End Property
Public Property Get DIAFOLGA() As Double
   DIAFOLGA = mvarDIAFOLGA
End Property
Public Property Let FLGVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGVENDA <> vData)
   mvarFLGVENDA = vData
End Property
Public Property Get FLGVENDA() As Double
   FLGVENDA = mvarFLGVENDA
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
 
   Sql = "Insert Into RFUNCIONARIO (IDLOJA, IDFUNCIONARIO, IDCOLIGADA, IDLOJA0" & VbNewLine 
   Sql = Sql & ", IDFINGER, CHAPA, NOME, DTADMISSAO, DTDEMISSAO" & VbNewLine 
   Sql = Sql & ", FLGCERTIFICADO, SITFUNC, COMPROD, VLCOMPROD, TPCOMPROD" & VbNewLine 
   Sql = Sql & ", COMSERV, VLCOMSERV, TPCOMSERV, OBS, TELEFONE" & VbNewLine 
   Sql = Sql & ", CELULAR, EMAIL, ENDERECO, BAIRRO, CIDADE" & VbNewLine 
   Sql = Sql & ", ESTADO, CEP, PAIS, DTNASC, SALARIO" & VbNewLine 
   Sql = Sql & ", SENHA, DTCADASTRO, ALTERSTAMP, TIMESTAMP, DIAFOLGA" & VbNewLine 
   Sql = Sql & ", FLGVENDA) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDFUNCIONARIO),0)+1 From RFUNCIONARIO Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDFUNCIONARIO) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDLOJA0) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFINGER) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCHAPA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTADMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTDEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGCERTIFICADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSITFUNC) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarCOMPROD) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLCOMPROD) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTPCOMPROD) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarCOMSERV) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLCOMSERV) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTPCOMSERV) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTELEFONE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCELULAR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarENDERECO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarESTADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCEP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTNASC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSALARIO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSENHA) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", " &  SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", " &  SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If
   Sql = Sql & ", " & SqlNum(mvarDIAFOLGA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGVENDA) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDFUNCIONARIO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From RFUNCIONARIO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDFUNCIONARIO)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFUNCIONARIO) Then Sql = Sql & " IDFUNCIONARIO = " & SqlNum(Cstr(Ch_IDFUNCIONARIO)) & " AND " & VbNewLine 
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

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = "
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDFUNCIONARIO = " & mvarIDFUNCIONARIO
   sTag = sTag & "|IDCOLIGADA = " & mvarIDCOLIGADA
   sTag = sTag & "|IDLOJA0 = " & mvarIDLOJA0
   sTag = sTag & "|IDFINGER = " & mvarIDFINGER
   sTag = sTag & "|CHAPA = " & mvarCHAPA
   sTag = sTag & "|NOME = " & mvarNOME
   sTag = sTag & "|DTADMISSAO = " & mvarDTADMISSAO
   sTag = sTag & "|DTDEMISSAO = " & mvarDTDEMISSAO
   sTag = sTag & "|FLGCERTIFICADO = " & mvarFLGCERTIFICADO
   sTag = sTag & "|SITFUNC = " & mvarSITFUNC
   sTag = sTag & "|COMPROD = " & mvarCOMPROD
   sTag = sTag & "|VLCOMPROD = " & mvarVLCOMPROD
   sTag = sTag & "|TPCOMPROD = " & mvarTPCOMPROD
   sTag = sTag & "|COMSERV = " & mvarCOMSERV
   sTag = sTag & "|VLCOMSERV = " & mvarVLCOMSERV
   sTag = sTag & "|TPCOMSERV = " & mvarTPCOMSERV
   sTag = sTag & "|OBS = " & mvarOBS
   sTag = sTag & "|TELEFONE = " & mvarTELEFONE
   sTag = sTag & "|CELULAR = " & mvarCELULAR
   sTag = sTag & "|EMAIL = " & mvarEMAIL
   sTag = sTag & "|ENDERECO = " & mvarENDERECO
   sTag = sTag & "|BAIRRO = " & mvarBAIRRO
   sTag = sTag & "|CIDADE = " & mvarCIDADE
   sTag = sTag & "|ESTADO = " & mvarESTADO
   sTag = sTag & "|CEP = " & mvarCEP
   sTag = sTag & "|PAIS = " & mvarPAIS
   sTag = sTag & "|DTNASC = " & mvarDTNASC
   sTag = sTag & "|SALARIO = " & mvarSALARIO
   sTag = sTag & "|SENHA = " & mvarSENHA
   sTag = sTag & "|DTCADASTRO = " & mvarDTCADASTRO
   sTag = sTag & "|DIAFOLGA = " & mvarDIAFOLGA
   sTag = sTag & "|FLGVENDA = " & mvarFLGVENDA
   sTag = sTag & "|Where = " & IIf(IsMissing(Ch_WHERE), "", Ch_WHERE)
   sTag = sTag & "|"

   MyDelRow.Query = Sql & ";"
   MyDelRow.Tag = sTag
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update RFUNCIONARIO Set " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDFUNCIONARIO  = " & SqlNum(mvarIDFUNCIONARIO) & VbNewLine 
   Sql = Sql & " , IDCOLIGADA     = " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & " , IDLOJA0        = " & SqlNum(mvarIDLOJA0) & VbNewLine 
   Sql = Sql & " , IDFINGER       = " & SqlNum(mvarIDFINGER) & VbNewLine 
   Sql = Sql & " , CHAPA          = " & SqlStr(mvarCHAPA) & VbNewLine 
   Sql = Sql & " , NOME           = " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & " , DTADMISSAO     = " & SqlDate(mvarDTADMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTDEMISSAO     = " & SqlDate(mvarDTDEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , FLGCERTIFICADO = " & SqlNum(mvarFLGCERTIFICADO) & VbNewLine 
   Sql = Sql & " , SITFUNC        = " & SqlStr(mvarSITFUNC) & VbNewLine 
   Sql = Sql & " , COMPROD        = " & SqlNum(mvarCOMPROD) & VbNewLine 
   Sql = Sql & " , VLCOMPROD      = " & SqlNum(mvarVLCOMPROD) & VbNewLine 
   Sql = Sql & " , TPCOMPROD      = " & SqlNum(mvarTPCOMPROD) & VbNewLine 
   Sql = Sql & " , COMSERV        = " & SqlNum(mvarCOMSERV) & VbNewLine 
   Sql = Sql & " , VLCOMSERV      = " & SqlNum(mvarVLCOMSERV) & VbNewLine 
   Sql = Sql & " , TPCOMSERV      = " & SqlNum(mvarTPCOMSERV) & VbNewLine 
   Sql = Sql & " , OBS            = " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & " , TELEFONE       = " & SqlStr(mvarTELEFONE) & VbNewLine 
   Sql = Sql & " , CELULAR        = " & SqlStr(mvarCELULAR) & VbNewLine 
   Sql = Sql & " , EMAIL          = " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & " , ENDERECO       = " & SqlStr(mvarENDERECO) & VbNewLine 
   Sql = Sql & " , BAIRRO         = " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & " , CIDADE         = " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & " , ESTADO         = " & SqlStr(mvarESTADO) & VbNewLine 
   Sql = Sql & " , CEP            = " & SqlStr(mvarCEP) & VbNewLine 
   Sql = Sql & " , PAIS           = " & SqlStr(mvarPAIS) & VbNewLine 
   Sql = Sql & " , DTNASC         = " & SqlDate(mvarDTNASC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , SALARIO        = " & SqlNum(mvarSALARIO) & VbNewLine 
   Sql = Sql & " , SENHA          = " & SqlStr(mvarSENHA) & VbNewLine 
   Sql = Sql & " , DTCADASTRO     = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP     = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP     = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP      = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP      = GetDate()" & vbNewLine
   End If
   Sql = Sql & " , DIAFOLGA       = " & SqlNum(mvarDIAFOLGA) & VbNewLine 
   Sql = Sql & " , FLGVENDA       = " & SqlNum(mvarFLGVENDA) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'RFUNCIONARIO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDFUNCIONARIO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDFUNCIONARIO, IDCOLIGADA, IDLOJA0" & VbNewLine 
   Sql = Sql & ", IDFINGER, CHAPA, NOME, DTADMISSAO, DTDEMISSAO" & VbNewLine 
   Sql = Sql & ", FLGCERTIFICADO, SITFUNC, COMPROD, VLCOMPROD, TPCOMPROD" & VbNewLine 
   Sql = Sql & ", COMSERV, VLCOMSERV, TPCOMSERV, OBS, TELEFONE" & VbNewLine 
   Sql = Sql & ", CELULAR, EMAIL, ENDERECO, BAIRRO, CIDADE" & VbNewLine 
   Sql = Sql & ", ESTADO, CEP, PAIS, DTNASC, SALARIO" & VbNewLine 
   Sql = Sql & ", SENHA, DTCADASTRO, ALTERSTAMP, TIMESTAMP, DIAFOLGA" & VbNewLine 
   Sql = Sql & ", FLGVENDA" & VbNewLine 
   Sql = Sql &" From RFUNCIONARIO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDFUNCIONARIO)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFUNCIONARIO) Then Sql = Sql & " IDFUNCIONARIO = " & SqlNum(Cstr(Ch_IDFUNCIONARIO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDFUNCIONARIO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDFUNCIONARIO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDFUNCIONARIO = XVal(!IDFUNCIONARIO & "")
      mvarIDCOLIGADA = XVal(!IDCOLIGADA & "")
      mvarIDLOJA0 = XVal(!IDLOJA0 & "")
      mvarIDFINGER = XVal(!IDFINGER & "")
      mvarCHAPA = !CHAPA & ""
      mvarNOME = !NOME & ""
      mvarDTADMISSAO = xDate(!DTADMISSAO & "", True)
      mvarDTDEMISSAO = xDate(!DTDEMISSAO & "", True)
      mvarFLGCERTIFICADO = XVal(!FLGCERTIFICADO & "")
      mvarSITFUNC = !SITFUNC & ""
      mvarCOMPROD = XVal(!COMPROD & "")
      mvarVLCOMPROD = XVal(!VLCOMPROD & "")
      mvarTPCOMPROD = XVal(!TPCOMPROD & "")
      mvarCOMSERV = XVal(!COMSERV & "")
      mvarVLCOMSERV = XVal(!VLCOMSERV & "")
      mvarTPCOMSERV = XVal(!TPCOMSERV & "")
      mvarOBS = !OBS & ""
      mvarTELEFONE = !TELEFONE & ""
      mvarCELULAR = !CELULAR & ""
      mvarEMAIL = !EMAIL & ""
      mvarENDERECO = !ENDERECO & ""
      mvarBAIRRO = !BAIRRO & ""
      mvarCIDADE = !CIDADE & ""
      mvarESTADO = !ESTADO & ""
      mvarCEP = !CEP & ""
      mvarPAIS = !PAIS & ""
      mvarDTNASC = xDate(!DTNASC & "", True)
      mvarSALARIO = XVal(!SALARIO & "")
      mvarSENHA = !SENHA & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarDIAFOLGA = XVal(!DIAFOLGA & "")
      mvarFLGVENDA = XVal(!FLGVENDA & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDFUNCIONARIO = 0
   mvarIDCOLIGADA = 0
   mvarIDLOJA0 = 0
   mvarIDFINGER = 0
   mvarCHAPA = ""
   mvarNOME = ""
   mvarDTADMISSAO = ""
   mvarDTDEMISSAO = ""
   mvarFLGCERTIFICADO = 0
   mvarSITFUNC = ""
   mvarCOMPROD = 0
   mvarVLCOMPROD = 0
   mvarTPCOMPROD = 0
   mvarCOMSERV = 0
   mvarVLCOMSERV = 0
   mvarTPCOMSERV = 0
   mvarOBS = ""
   mvarTELEFONE = ""
   mvarCELULAR = ""
   mvarEMAIL = ""
   mvarENDERECO = ""
   mvarBAIRRO = ""
   mvarCIDADE = ""
   mvarESTADO = ""
   mvarCEP = ""
   mvarPAIS = ""
   mvarDTNASC = ""
   mvarSALARIO = 0
   mvarSENHA = ""
   mvarDTCADASTRO = ""
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""
   mvarDIAFOLGA = 0
   mvarFLGVENDA = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDFUNCIONARIO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDFUNCIONARIO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update RFUNCIONARIO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDFUNCIONARIO = " & SqlNum(Ch_IDFUNCIONARIO)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO)
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
