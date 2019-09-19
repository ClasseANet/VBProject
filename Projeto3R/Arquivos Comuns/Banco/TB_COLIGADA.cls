VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_COLIGADA"
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

Private mvarIDCOLIGADA   As Double
Private mvarNMCOLIGADA   As String
Private mvarRAZAO        As String
Private mvarCNPJ         As String
Private mvarINSCMUNIC    As String
Private mvarTELEFONE1    As String
Private mvarTELEFONE2    As String
Private mvarFAX          As String
Private mvarEMAIL        As String
Private mvarLOGRADOURO   As String
Private mvarBAIRRO       As String
Private mvarCIDADE       As String
Private mvarESTADO       As String
Private mvarCEP          As String
Private mvarNMCONTATO    As String
Private mvarCARGOCONTATO As String
Private mvarPAIS         As String
Private mvarSIGLAPAIS    As String
Private mvarINSCEST      As String
Private mvarINSCOUTROS   As String
Private mvarDTCADASTRO   As String
Private mvarPJ           As String
Private mvarENDCORREIO   As String
Private mvarIDPAIS       As Double
Private mvarIDEMPRESA    As Double
Private mvarIDAUX        As Double
Private mvarTAG          As String
Private mvarTIMESTAMP    As String
Private mvarALTERSTAMP   As Double
Public Property Get IDCOLIGADA() As Double
   IDCOLIGADA = mvarIDCOLIGADA
End Property
Public Property Let NMCOLIGADA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMCOLIGADA <> vData)
   mvarNMCOLIGADA = vData
End Property
Public Property Get NMCOLIGADA() As String
   NMCOLIGADA = mvarNMCOLIGADA
End Property
Public Property Let RAZAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarRAZAO <> vData)
   mvarRAZAO = vData
End Property
Public Property Get RAZAO() As String
   RAZAO = mvarRAZAO
End Property
Public Property Let CNPJ(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCNPJ <> vData)
   mvarCNPJ = vData
End Property
Public Property Get CNPJ() As String
   CNPJ = mvarCNPJ
End Property
Public Property Let INSCMUNIC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarINSCMUNIC <> vData)
   mvarINSCMUNIC = vData
End Property
Public Property Get INSCMUNIC() As String
   INSCMUNIC = mvarINSCMUNIC
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
Public Property Let LOGRADOURO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLOGRADOURO <> vData)
   mvarLOGRADOURO = vData
End Property
Public Property Get LOGRADOURO() As String
   LOGRADOURO = mvarLOGRADOURO
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
Public Property Let PAIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPAIS <> vData)
   mvarPAIS = vData
End Property
Public Property Get PAIS() As String
   PAIS = mvarPAIS
End Property
Public Property Let SIGLAPAIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSIGLAPAIS <> vData)
   mvarSIGLAPAIS = vData
End Property
Public Property Get SIGLAPAIS() As String
   SIGLAPAIS = mvarSIGLAPAIS
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
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let PJ(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPJ <> vData)
   mvarPJ = vData
End Property
Public Property Get PJ() As String
   PJ = mvarPJ
End Property
Public Property Let ENDCORREIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarENDCORREIO <> vData)
   mvarENDCORREIO = vData
End Property
Public Property Get ENDCORREIO() As String
   ENDCORREIO = mvarENDCORREIO
End Property
Public Property Let IDPAIS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAIS <> vData)
   mvarIDPAIS = vData
End Property
Public Property Get IDPAIS() As Double
   IDPAIS = mvarIDPAIS
End Property
Public Property Let IDEMPRESA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEMPRESA <> vData)
   mvarIDEMPRESA = vData
End Property
Public Property Get IDEMPRESA() As Double
   IDEMPRESA = mvarIDEMPRESA
End Property
Public Property Let IDAUX(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAUX <> vData)
   mvarIDAUX = vData
End Property
Public Property Get IDAUX() As Double
   IDAUX = mvarIDAUX
End Property
Public Property Let TAG(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTAG <> vData)
   mvarTAG = vData
End Property
Public Property Get TAG() As String
   TAG = mvarTAG
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
 
   Sql = "Update COLIGADA Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDCOLIGADA     = " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into COLIGADA (NMCOLIGADA, RAZAO, CNPJ" & VbNewLine 
   Sql = Sql & ", INSCMUNIC, TELEFONE1, TELEFONE2, FAX, EMAIL" & VbNewLine 
   Sql = Sql & ", LOGRADOURO, BAIRRO, CIDADE, ESTADO, CEP" & VbNewLine 
   Sql = Sql & ", NMCONTATO, CARGOCONTATO, PAIS, SIGLAPAIS, INSCEST" & VbNewLine 
   Sql = Sql & ", INSCOUTROS, DTCADASTRO, PJ, ENDCORREIO, IDPAIS" & VbNewLine 
   Sql = Sql & ", IDEMPRESA, IDAUX, TAG, TIMESTAMP, ALTERSTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarNMCOLIGADA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarRAZAO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCNPJ) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarINSCMUNIC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTELEFONE1) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTELEFONE2) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFAX) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLOGRADOURO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarESTADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCEP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMCONTATO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCARGOCONTATO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSIGLAPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarINSCEST) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarINSCOUTROS) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPJ) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarENDCORREIO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDEMPRESA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDAUX) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTAG) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", " &  SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", " &  SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDCOLIGADA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From COLIGADA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDCOLIGADA)  Then 
      Sql = Sql & " IDCOLIGADA = " & SqlNum(mvarIDCOLIGADA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDCOLIGADA) Then Sql = Sql & " IDCOLIGADA = " & SqlNum(Cstr(Ch_IDCOLIGADA)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDCOLIGADA = " & mvarIDCOLIGADA
   sTag = sTag & "|NMCOLIGADA = " & mvarNMCOLIGADA
   sTag = sTag & "|RAZAO = " & mvarRAZAO
   sTag = sTag & "|CNPJ = " & mvarCNPJ
   sTag = sTag & "|INSCMUNIC = " & mvarINSCMUNIC
   sTag = sTag & "|TELEFONE1 = " & mvarTELEFONE1
   sTag = sTag & "|TELEFONE2 = " & mvarTELEFONE2
   sTag = sTag & "|FAX = " & mvarFAX
   sTag = sTag & "|EMAIL = " & mvarEMAIL
   sTag = sTag & "|LOGRADOURO = " & mvarLOGRADOURO
   sTag = sTag & "|BAIRRO = " & mvarBAIRRO
   sTag = sTag & "|CIDADE = " & mvarCIDADE
   sTag = sTag & "|ESTADO = " & mvarESTADO
   sTag = sTag & "|CEP = " & mvarCEP
   sTag = sTag & "|NMCONTATO = " & mvarNMCONTATO
   sTag = sTag & "|CARGOCONTATO = " & mvarCARGOCONTATO
   sTag = sTag & "|PAIS = " & mvarPAIS
   sTag = sTag & "|SIGLAPAIS = " & mvarSIGLAPAIS
   sTag = sTag & "|INSCEST = " & mvarINSCEST
   sTag = sTag & "|INSCOUTROS = " & mvarINSCOUTROS
   sTag = sTag & "|DTCADASTRO = " & mvarDTCADASTRO
   sTag = sTag & "|PJ = " & mvarPJ
   sTag = sTag & "|ENDCORREIO = " & mvarENDCORREIO
   sTag = sTag & "|IDPAIS = " & mvarIDPAIS
   sTag = sTag & "|IDEMPRESA = " & mvarIDEMPRESA
   sTag = sTag & "|IDAUX = " & mvarIDAUX
   sTag = sTag & "|TAG = " & mvarTAG
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
 
   Sql = "Update COLIGADA Set " & VbNewLine 
   Sql = Sql & " NMCOLIGADA   = " & SqlStr(mvarNMCOLIGADA) & VbNewLine 
   Sql = Sql & " , RAZAO        = " & SqlStr(mvarRAZAO) & VbNewLine 
   Sql = Sql & " , CNPJ         = " & SqlStr(mvarCNPJ) & VbNewLine 
   Sql = Sql & " , INSCMUNIC    = " & SqlStr(mvarINSCMUNIC) & VbNewLine 
   Sql = Sql & " , TELEFONE1    = " & SqlStr(mvarTELEFONE1) & VbNewLine 
   Sql = Sql & " , TELEFONE2    = " & SqlStr(mvarTELEFONE2) & VbNewLine 
   Sql = Sql & " , FAX          = " & SqlStr(mvarFAX) & VbNewLine 
   Sql = Sql & " , EMAIL        = " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & " , LOGRADOURO   = " & SqlStr(mvarLOGRADOURO) & VbNewLine 
   Sql = Sql & " , BAIRRO       = " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & " , CIDADE       = " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & " , ESTADO       = " & SqlStr(mvarESTADO) & VbNewLine 
   Sql = Sql & " , CEP          = " & SqlStr(mvarCEP) & VbNewLine 
   Sql = Sql & " , NMCONTATO    = " & SqlStr(mvarNMCONTATO) & VbNewLine 
   Sql = Sql & " , CARGOCONTATO = " & SqlStr(mvarCARGOCONTATO) & VbNewLine 
   Sql = Sql & " , PAIS         = " & SqlStr(mvarPAIS) & VbNewLine 
   Sql = Sql & " , SIGLAPAIS    = " & SqlStr(mvarSIGLAPAIS) & VbNewLine 
   Sql = Sql & " , INSCEST      = " & SqlStr(mvarINSCEST) & VbNewLine 
   Sql = Sql & " , INSCOUTROS   = " & SqlStr(mvarINSCOUTROS) & VbNewLine 
   Sql = Sql & " , DTCADASTRO   = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , PJ           = " & SqlStr(mvarPJ) & VbNewLine 
   Sql = Sql & " , ENDCORREIO   = " & SqlStr(mvarENDCORREIO) & VbNewLine 
   Sql = Sql & " , IDPAIS       = " & SqlNum(mvarIDPAIS) & VbNewLine 
   Sql = Sql & " , IDEMPRESA    = " & SqlNum(mvarIDEMPRESA) & VbNewLine 
   Sql = Sql & " , IDAUX        = " & SqlNum(mvarIDAUX) & VbNewLine 
   Sql = Sql & " , TAG          = " & SqlStr(mvarTAG) & VbNewLine 
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
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'COLIGADA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDCOLIGADA     = " & SqlNum(mvarIDCOLIGADA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDCOLIGADA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDCOLIGADA, NMCOLIGADA, RAZAO, CNPJ" & VbNewLine 
   Sql = Sql & ", INSCMUNIC, TELEFONE1, TELEFONE2, FAX, EMAIL" & VbNewLine 
   Sql = Sql & ", LOGRADOURO, BAIRRO, CIDADE, ESTADO, CEP" & VbNewLine 
   Sql = Sql & ", NMCONTATO, CARGOCONTATO, PAIS, SIGLAPAIS, INSCEST" & VbNewLine 
   Sql = Sql & ", INSCOUTROS, DTCADASTRO, PJ, ENDCORREIO, IDPAIS" & VbNewLine 
   Sql = Sql & ", IDEMPRESA, IDAUX, TAG, TIMESTAMP, ALTERSTAMP" & VbNewLine 
   Sql = Sql &" From COLIGADA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDCOLIGADA)  Then 
      Sql = Sql & " IDCOLIGADA = " & SqlNum(mvarIDCOLIGADA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDCOLIGADA) Then Sql = Sql & " IDCOLIGADA = " & SqlNum(Cstr(Ch_IDCOLIGADA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDCOLIGADA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDCOLIGADA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCOLIGADA = XVal(!IDCOLIGADA & "")
      mvarNMCOLIGADA = !NMCOLIGADA & ""
      mvarRAZAO = !RAZAO & ""
      mvarCNPJ = !CNPJ & ""
      mvarINSCMUNIC = !INSCMUNIC & ""
      mvarTELEFONE1 = !TELEFONE1 & ""
      mvarTELEFONE2 = !TELEFONE2 & ""
      mvarFAX = !FAX & ""
      mvarEMAIL = !EMAIL & ""
      mvarLOGRADOURO = !LOGRADOURO & ""
      mvarBAIRRO = !BAIRRO & ""
      mvarCIDADE = !CIDADE & ""
      mvarESTADO = !ESTADO & ""
      mvarCEP = !CEP & ""
      mvarNMCONTATO = !NMCONTATO & ""
      mvarCARGOCONTATO = !CARGOCONTATO & ""
      mvarPAIS = !PAIS & ""
      mvarSIGLAPAIS = !SIGLAPAIS & ""
      mvarINSCEST = !INSCEST & ""
      mvarINSCOUTROS = !INSCOUTROS & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarPJ = !PJ & ""
      mvarENDCORREIO = !ENDCORREIO & ""
      mvarIDPAIS = XVal(!IDPAIS & "")
      mvarIDEMPRESA = XVal(!IDEMPRESA & "")
      mvarIDAUX = XVal(!IDAUX & "")
      mvarTAG = !TAG & ""
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDCOLIGADA = 0
   mvarNMCOLIGADA = ""
   mvarRAZAO = ""
   mvarCNPJ = ""
   mvarINSCMUNIC = ""
   mvarTELEFONE1 = ""
   mvarTELEFONE2 = ""
   mvarFAX = ""
   mvarEMAIL = ""
   mvarLOGRADOURO = ""
   mvarBAIRRO = ""
   mvarCIDADE = ""
   mvarESTADO = ""
   mvarCEP = ""
   mvarNMCONTATO = ""
   mvarCARGOCONTATO = ""
   mvarPAIS = ""
   mvarSIGLAPAIS = ""
   mvarINSCEST = ""
   mvarINSCOUTROS = ""
   mvarDTCADASTRO = ""
   mvarPJ = ""
   mvarENDCORREIO = ""
   mvarIDPAIS = 0
   mvarIDEMPRESA = 0
   mvarIDAUX = 0
   mvarTAG = ""
   mvarTIMESTAMP = ""
   mvarALTERSTAMP = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDCOLIGADA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
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
