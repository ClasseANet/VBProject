VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_OLOJA"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarXDb               As Object
Private mvarRS                As Object

Private mvarQryInsert         As String
Private mvarQryUpDate         As String
Private mvarQryDelete         As String
Private mvarQrySelect         As String
Private mvarQrySave           As String

Private mvarisDirt            As Boolean

Private mvarIDLOJA          As Double
Private mvarIDCOLIGADA      As Double
Private mvarCNPJ            As String
Private mvarNOME            As String
Private mvarENDERECO        As String
Private mvarBAIRRO          As String
Private mvarCIDADE          As String
Private mvarESTADO          As String
Private mvarINSCEST         As String
Private mvarINSCMUNIC       As String
Private mvarEMAIL           As String
Private mvarTELEFONE1       As String
Private mvarTELEFONE2       As String
Private mvarCEP             As String
Private mvarFAX             As String
Private mvarIDPAIS          As Double
Private mvarNMCONTATO       As String
Private mvarCARGOCONTATO    As String
Private mvarDTOPERACAO      As String
Private mvarDIMENSAO        As String
Private mvarRAZAO           As String
Private mvarNMFANTASIA      As String
Private mvarRZABREV         As String
Private mvarFLGFILIAL       As Double
Private mvarCODSERVMUNIC    As String
Private mvarCODSERVFEDERAL  As String
Private mvarREGESPECIALTRIB As Double
Private mvarSIMPLES         As Double
Private mvarTRIBUTACAO      As Double
Private mvarINCENTIVOCULT   As Double
Private mvarATIVO           As Double
Private mvarALTERSTAMP      As Double
Private mvarTIMESTAMP       As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCOLIGADA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCOLIGADA <> vData)
   mvarIDCOLIGADA = vData
End Property
Public Property Get IDCOLIGADA() As Double
   IDCOLIGADA = mvarIDCOLIGADA
End Property
Public Property Let CNPJ(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCNPJ <> vData)
   mvarCNPJ = vData
End Property
Public Property Get CNPJ() As String
   CNPJ = mvarCNPJ
End Property
Public Property Let NOME(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNOME <> vData)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
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
Public Property Let INSCEST(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarINSCEST <> vData)
   mvarINSCEST = vData
End Property
Public Property Get INSCEST() As String
   INSCEST = mvarINSCEST
End Property
Public Property Let INSCMUNIC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarINSCMUNIC <> vData)
   mvarINSCMUNIC = vData
End Property
Public Property Get INSCMUNIC() As String
   INSCMUNIC = mvarINSCMUNIC
End Property
Public Property Let EMAIL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEMAIL <> vData)
   mvarEMAIL = vData
End Property
Public Property Get EMAIL() As String
   EMAIL = mvarEMAIL
End Property
Public Property Let TELEFONE1(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTELEFONE1 <> vData)
   mvarTELEFONE1 = vData
End Property
Public Property Get TELEFONE1() As String
   TELEFONE1 = mvarTELEFONE1
End Property
Public Property Let TELEFONE2(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTELEFONE2 <> vData)
   mvarTELEFONE2 = vData
End Property
Public Property Get TELEFONE2() As String
   TELEFONE2 = mvarTELEFONE2
End Property
Public Property Let CEP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCEP <> vData)
   mvarCEP = vData
End Property
Public Property Get CEP() As String
   CEP = mvarCEP
End Property
Public Property Let FAX(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFAX <> vData)
   mvarFAX = vData
End Property
Public Property Get FAX() As String
   FAX = mvarFAX
End Property
Public Property Let IDPAIS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAIS <> vData)
   mvarIDPAIS = vData
End Property
Public Property Get IDPAIS() As Double
   IDPAIS = mvarIDPAIS
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
Public Property Let DTOPERACAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTOPERACAO <> vData)
   mvarDTOPERACAO = vData
End Property
Public Property Get DTOPERACAO() As String
   DTOPERACAO = mvarDTOPERACAO
End Property
Public Property Let DIMENSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDIMENSAO <> vData)
   mvarDIMENSAO = vData
End Property
Public Property Get DIMENSAO() As String
   DIMENSAO = mvarDIMENSAO
End Property
Public Property Let RAZAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRAZAO <> vData)
   mvarRAZAO = vData
End Property
Public Property Get RAZAO() As String
   RAZAO = mvarRAZAO
End Property
Public Property Let NMFANTASIA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMFANTASIA <> vData)
   mvarNMFANTASIA = vData
End Property
Public Property Get NMFANTASIA() As String
   NMFANTASIA = mvarNMFANTASIA
End Property
Public Property Let RZABREV(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRZABREV <> vData)
   mvarRZABREV = vData
End Property
Public Property Get RZABREV() As String
   RZABREV = mvarRZABREV
End Property
Public Property Let FLGFILIAL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGFILIAL <> vData)
   mvarFLGFILIAL = vData
End Property
Public Property Get FLGFILIAL() As Double
   FLGFILIAL = mvarFLGFILIAL
End Property
Public Property Let CODSERVMUNIC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSERVMUNIC <> vData)
   mvarCODSERVMUNIC = vData
End Property
Public Property Get CODSERVMUNIC() As String
   CODSERVMUNIC = mvarCODSERVMUNIC
End Property
Public Property Let CODSERVFEDERAL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSERVFEDERAL <> vData)
   mvarCODSERVFEDERAL = vData
End Property
Public Property Get CODSERVFEDERAL() As String
   CODSERVFEDERAL = mvarCODSERVFEDERAL
End Property
Public Property Let REGESPECIALTRIB(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarREGESPECIALTRIB <> vData)
   mvarREGESPECIALTRIB = vData
End Property
Public Property Get REGESPECIALTRIB() As Double
   REGESPECIALTRIB = mvarREGESPECIALTRIB
End Property
Public Property Let SIMPLES(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSIMPLES <> vData)
   mvarSIMPLES = vData
End Property
Public Property Get SIMPLES() As Double
   SIMPLES = mvarSIMPLES
End Property
Public Property Let TRIBUTACAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTRIBUTACAO <> vData)
   mvarTRIBUTACAO = vData
End Property
Public Property Get TRIBUTACAO() As Double
   TRIBUTACAO = mvarTRIBUTACAO
End Property
Public Property Let INCENTIVOCULT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarINCENTIVOCULT <> vData)
   mvarINCENTIVOCULT = vData
End Property
Public Property Get INCENTIVOCULT() As Double
   INCENTIVOCULT = mvarINCENTIVOCULT
End Property
Public Property Let ATIVO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVO <> vData)
   mvarATIVO = vData
End Property
Public Property Get ATIVO() As Double
   ATIVO = mvarATIVO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OLOJA Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA            = " & SqlNum(mvarIDLOJA) & vbNewLine
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
 
   Sql = "Insert Into OLOJA (IDLOJA, IDCOLIGADA, CNPJ, NOME" & vbNewLine
   Sql = Sql & ", ENDERECO, BAIRRO, CIDADE, ESTADO, INSCEST" & vbNewLine
   Sql = Sql & ", INSCMUNIC, EMAIL, TELEFONE1, TELEFONE2, CEP" & vbNewLine
   Sql = Sql & ", FAX, IDPAIS, NMCONTATO, CARGOCONTATO, DTOPERACAO" & vbNewLine
   Sql = Sql & ", DIMENSAO, RAZAO, NMFANTASIA, RZABREV, FLGFILIAL" & vbNewLine
   Sql = Sql & ", CODSERVMUNIC, CODSERVFEDERAL, REGESPECIALTRIB, SIMPLES, TRIBUTACAO" & vbNewLine
   Sql = Sql & ", INCENTIVOCULT, ATIVO, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   If pAutoId Then
      Sql = Sql & "(Select isNull(Max(IDLOJA),0)+1 From OLOJA)"
   Else
      Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   End If
   Sql = Sql & ", " & SqlNum(mvarIDCOLIGADA) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCNPJ) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarNOME) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarENDERECO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarBAIRRO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCIDADE) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarESTADO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarINSCEST) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarINSCMUNIC) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarEMAIL) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTELEFONE1) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTELEFONE2) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCEP) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarFAX) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDPAIS) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarNMCONTATO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCARGOCONTATO) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTOPERACAO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarDIMENSAO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarRAZAO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarNMFANTASIA) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarRZABREV) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGFILIAL) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCODSERVMUNIC) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCODSERVFEDERAL) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarREGESPECIALTRIB) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarSIMPLES) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarTRIBUTACAO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarINCENTIVOCULT) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarATIVO) & vbNewLine
   If pSinc Then
      Sql = Sql & ", " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OLOJA" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
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
   sTag = sTag & "|IDCOLIGADA = " & mvarIDCOLIGADA
   sTag = sTag & "|CNPJ = " & mvarCNPJ
   sTag = sTag & "|NOME = " & mvarNOME
   sTag = sTag & "|ENDERECO = " & mvarENDERECO
   sTag = sTag & "|BAIRRO = " & mvarBAIRRO
   sTag = sTag & "|CIDADE = " & mvarCIDADE
   sTag = sTag & "|ESTADO = " & mvarESTADO
   sTag = sTag & "|INSCEST = " & mvarINSCEST
   sTag = sTag & "|INSCMUNIC = " & mvarINSCMUNIC
   sTag = sTag & "|EMAIL = " & mvarEMAIL
   sTag = sTag & "|TELEFONE1 = " & mvarTELEFONE1
   sTag = sTag & "|TELEFONE2 = " & mvarTELEFONE2
   sTag = sTag & "|CEP = " & mvarCEP
   sTag = sTag & "|FAX = " & mvarFAX
   sTag = sTag & "|IDPAIS = " & mvarIDPAIS
   sTag = sTag & "|NMCONTATO = " & mvarNMCONTATO
   sTag = sTag & "|CARGOCONTATO = " & mvarCARGOCONTATO
   sTag = sTag & "|DTOPERACAO = " & mvarDTOPERACAO
   sTag = sTag & "|DIMENSAO = " & mvarDIMENSAO
   sTag = sTag & "|RAZAO = " & mvarRAZAO
   sTag = sTag & "|NMFANTASIA = " & mvarNMFANTASIA
   sTag = sTag & "|RZABREV = " & mvarRZABREV
   sTag = sTag & "|FLGFILIAL = " & mvarFLGFILIAL
   sTag = sTag & "|CODSERVMUNIC = " & mvarCODSERVMUNIC
   sTag = sTag & "|CODSERVFEDERAL = " & mvarCODSERVFEDERAL
   sTag = sTag & "|REGESPECIALTRIB = " & mvarREGESPECIALTRIB
   sTag = sTag & "|SIMPLES = " & mvarSIMPLES
   sTag = sTag & "|TRIBUTACAO = " & mvarTRIBUTACAO
   sTag = sTag & "|INCENTIVOCULT = " & mvarINCENTIVOCULT
   sTag = sTag & "|ATIVO = " & mvarATIVO
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
 
   Sql = "Update OLOJA Set " & vbNewLine
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDCOLIGADA      = " & SqlNum(mvarIDCOLIGADA) & vbNewLine
   Sql = Sql & " , CNPJ            = " & SqlStr(mvarCNPJ) & vbNewLine
   Sql = Sql & " , NOME            = " & SqlStr(mvarNOME) & vbNewLine
   Sql = Sql & " , ENDERECO        = " & SqlStr(mvarENDERECO) & vbNewLine
   Sql = Sql & " , BAIRRO          = " & SqlStr(mvarBAIRRO) & vbNewLine
   Sql = Sql & " , CIDADE          = " & SqlStr(mvarCIDADE) & vbNewLine
   Sql = Sql & " , ESTADO          = " & SqlStr(mvarESTADO) & vbNewLine
   Sql = Sql & " , INSCEST         = " & SqlStr(mvarINSCEST) & vbNewLine
   Sql = Sql & " , INSCMUNIC       = " & SqlStr(mvarINSCMUNIC) & vbNewLine
   Sql = Sql & " , EMAIL           = " & SqlStr(mvarEMAIL) & vbNewLine
   Sql = Sql & " , TELEFONE1       = " & SqlStr(mvarTELEFONE1) & vbNewLine
   Sql = Sql & " , TELEFONE2       = " & SqlStr(mvarTELEFONE2) & vbNewLine
   Sql = Sql & " , CEP             = " & SqlStr(mvarCEP) & vbNewLine
   Sql = Sql & " , FAX             = " & SqlStr(mvarFAX) & vbNewLine
   Sql = Sql & " , IDPAIS          = " & SqlNum(mvarIDPAIS) & vbNewLine
   Sql = Sql & " , NMCONTATO       = " & SqlStr(mvarNMCONTATO) & vbNewLine
   Sql = Sql & " , CARGOCONTATO    = " & SqlStr(mvarCARGOCONTATO) & vbNewLine
   Sql = Sql & " , DTOPERACAO      = " & SqlDate(mvarDTOPERACAO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , DIMENSAO        = " & SqlStr(mvarDIMENSAO) & vbNewLine
   Sql = Sql & " , RAZAO           = " & SqlStr(mvarRAZAO) & vbNewLine
   Sql = Sql & " , NMFANTASIA      = " & SqlStr(mvarNMFANTASIA) & vbNewLine
   Sql = Sql & " , RZABREV         = " & SqlStr(mvarRZABREV) & vbNewLine
   Sql = Sql & " , FLGFILIAL       = " & SqlNum(mvarFLGFILIAL) & vbNewLine
   Sql = Sql & " , CODSERVMUNIC    = " & SqlStr(mvarCODSERVMUNIC) & vbNewLine
   Sql = Sql & " , CODSERVFEDERAL  = " & SqlStr(mvarCODSERVFEDERAL) & vbNewLine
   Sql = Sql & " , REGESPECIALTRIB = " & SqlNum(mvarREGESPECIALTRIB) & vbNewLine
   Sql = Sql & " , SIMPLES         = " & SqlNum(mvarSIMPLES) & vbNewLine
   Sql = Sql & " , TRIBUTACAO      = " & SqlNum(mvarTRIBUTACAO) & vbNewLine
   Sql = Sql & " , INCENTIVOCULT   = " & SqlNum(mvarINCENTIVOCULT) & vbNewLine
   Sql = Sql & " , ATIVO           = " & SqlNum(mvarATIVO) & vbNewLine
   If pSinc Then
      Sql = Sql & ", ALTERSTAMP      = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP      = 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", TIMESTAMP       = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP       = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OLOJA'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA            = " & SqlNum(mvarIDLOJA) & vbNewLine

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
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDCOLIGADA, CNPJ, NOME" & vbNewLine
   Sql = Sql & ", ENDERECO, BAIRRO, CIDADE, ESTADO, INSCEST" & vbNewLine
   Sql = Sql & ", INSCMUNIC, EMAIL, TELEFONE1, TELEFONE2, CEP" & vbNewLine
   Sql = Sql & ", FAX, IDPAIS, NMCONTATO, CARGOCONTATO, DTOPERACAO" & vbNewLine
   Sql = Sql & ", DIMENSAO, RAZAO, NMFANTASIA, RZABREV, FLGFILIAL" & vbNewLine
   Sql = Sql & ", CODSERVMUNIC, CODSERVFEDERAL, REGESPECIALTRIB, SIMPLES, TRIBUTACAO" & vbNewLine
   Sql = Sql & ", INCENTIVOCULT, ATIVO, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From OLOJA" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
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
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCOLIGADA = xVal(!IDCOLIGADA & "")
      mvarCNPJ = !CNPJ & ""
      mvarNOME = !NOME & ""
      mvarENDERECO = !ENDERECO & ""
      mvarBAIRRO = !BAIRRO & ""
      mvarCIDADE = !CIDADE & ""
      mvarESTADO = !ESTADO & ""
      mvarINSCEST = !INSCEST & ""
      mvarINSCMUNIC = !INSCMUNIC & ""
      mvarEMAIL = !EMAIL & ""
      mvarTELEFONE1 = !TELEFONE1 & ""
      mvarTELEFONE2 = !TELEFONE2 & ""
      mvarCEP = !CEP & ""
      mvarFAX = !FAX & ""
      mvarIDPAIS = xVal(!IDPAIS & "")
      mvarNMCONTATO = !NMCONTATO & ""
      mvarCARGOCONTATO = !CARGOCONTATO & ""
      mvarDTOPERACAO = xDate(!DTOPERACAO & "", True)
      mvarDIMENSAO = !DIMENSAO & ""
      mvarRAZAO = !RAZAO & ""
      mvarNMFANTASIA = !NMFANTASIA & ""
      mvarRZABREV = !RZABREV & ""
      mvarFLGFILIAL = xVal(!FLGFILIAL & "")
      mvarCODSERVMUNIC = !CODSERVMUNIC & ""
      mvarCODSERVFEDERAL = !CODSERVFEDERAL & ""
      mvarREGESPECIALTRIB = xVal(!REGESPECIALTRIB & "")
      mvarSIMPLES = xVal(!SIMPLES & "")
      mvarTRIBUTACAO = xVal(!TRIBUTACAO & "")
      mvarINCENTIVOCULT = xVal(!INCENTIVOCULT & "")
      mvarATIVO = xVal(!ATIVO & "")
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDCOLIGADA = 0
   mvarCNPJ = ""
   mvarNOME = ""
   mvarENDERECO = ""
   mvarBAIRRO = ""
   mvarCIDADE = ""
   mvarESTADO = ""
   mvarINSCEST = ""
   mvarINSCMUNIC = ""
   mvarEMAIL = ""
   mvarTELEFONE1 = ""
   mvarTELEFONE2 = ""
   mvarCEP = ""
   mvarFAX = ""
   mvarIDPAIS = 0
   mvarNMCONTATO = ""
   mvarCARGOCONTATO = ""
   mvarDTOPERACAO = ""
   mvarDIMENSAO = ""
   mvarRAZAO = ""
   mvarNMFANTASIA = ""
   mvarRZABREV = ""
   mvarFLGFILIAL = 0
   mvarCODSERVMUNIC = ""
   mvarCODSERVFEDERAL = ""
   mvarREGESPECIALTRIB = 0
   mvarSIMPLES = 0
   mvarTRIBUTACAO = 0
   mvarINCENTIVOCULT = 0
   mvarATIVO = 0
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OLOJA Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
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
