VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OCONTATO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb         As Object 
Private mvarRS          As Object 

Private mvarQryInsert   As String
Private mvarQryUpDate   As String
Private mvarQryDelete   As String
Private mvarQrySelect   As String
Private mvarQrySave     As String

Private mvarisDirt      As Boolean

Private mvarIDLOJA     As Double
Private mvarIDCONTATO  As Double
Private mvarIDCOLIGADA As Double
Private mvarIDLOJA0    As Double
Private mvarEMPRESA    As String
Private mvarREGISTRO   As String
Private mvarTEL1       As String
Private mvarTEL2       As String
Private mvarFAX        As String
Private mvarEMAIL      As String
Private mvarENDERECO   As String
Private mvarBAIRRO     As String
Private mvarCIDADE     As String
Private mvarESTADO     As String
Private mvarCEP        As String
Private mvarCARGO      As String
Private mvarDTCADASTRO As String
Private mvarPJ         As Double
Private mvarPAIS       As String
Private mvarOBS        As String
Private mvarNOME       As String
Private mvarDTNASC     As String
Private mvarATIVO      As Double
Private mvarTIMESTAMP  As String
Private mvarALTERSTAMP As Double
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCONTATO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCONTATO <> vData)
   mvarIDCONTATO = vData
End Property
Public Property Get IDCONTATO() As Double
   IDCONTATO = mvarIDCONTATO
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
Public Property Let EMPRESA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEMPRESA <> vData)
   mvarEMPRESA = vData
End Property
Public Property Get EMPRESA() As String
   EMPRESA = mvarEMPRESA
End Property
Public Property Let REGISTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarREGISTRO <> vData)
   mvarREGISTRO = vData
End Property
Public Property Get REGISTRO() As String
   REGISTRO = mvarREGISTRO
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
Public Property Let CEP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCEP <> vData)
   mvarCEP = vData
End Property
Public Property Get CEP() As String
   CEP = mvarCEP
End Property
Public Property Let CARGO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCARGO <> vData)
   mvarCARGO = vData
End Property
Public Property Get CARGO() As String
   CARGO = mvarCARGO
End Property
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let PJ(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPJ <> vData)
   mvarPJ = vData
End Property
Public Property Get PJ() As Double
   PJ = mvarPJ
End Property
Public Property Let PAIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPAIS <> vData)
   mvarPAIS = vData
End Property
Public Property Get PAIS() As String
   PAIS = mvarPAIS
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
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
Public Property Let ATIVO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVO <> vData)
   mvarATIVO = vData
End Property
Public Property Get ATIVO() As Double
   ATIVO = mvarATIVO
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
 
   Sql = "Update OCONTATO Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDCONTATO = " & SqlNum(mvarIDCONTATO) & VbNewLine 
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
 
   Sql = "Insert Into OCONTATO (IDLOJA, IDCONTATO, IDCOLIGADA, IDLOJA0" & VbNewLine 
   Sql = Sql & ", EMPRESA, REGISTRO, TEL1, TEL2, FAX" & VbNewLine 
   Sql = Sql & ", EMAIL, ENDERECO, BAIRRO, CIDADE, ESTADO" & VbNewLine 
   Sql = Sql & ", CEP, CARGO, DTCADASTRO, PJ, PAIS" & VbNewLine 
   Sql = Sql & ", OBS, NOME, DTNASC, ATIVO, TIMESTAMP" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDCONTATO),0)+1 From OCONTATO Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDCONTATO) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDLOJA0) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMPRESA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarREGISTRO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTEL1) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTEL2) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFAX) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarENDERECO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarESTADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCEP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCARGO) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarPJ) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPAIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTNASC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarATIVO) & VbNewLine 
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
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDCONTATO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OCONTATO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDCONTATO)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDCONTATO = " & SqlNum(mvarIDCONTATO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCONTATO) Then Sql = Sql & " IDCONTATO = " & SqlNum(Cstr(Ch_IDCONTATO)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDCONTATO = " & mvarIDCONTATO
   sTag = sTag & "|IDCOLIGADA = " & mvarIDCOLIGADA
   sTag = sTag & "|IDLOJA0 = " & mvarIDLOJA0
   sTag = sTag & "|EMPRESA = " & mvarEMPRESA
   sTag = sTag & "|REGISTRO = " & mvarREGISTRO
   sTag = sTag & "|TEL1 = " & mvarTEL1
   sTag = sTag & "|TEL2 = " & mvarTEL2
   sTag = sTag & "|FAX = " & mvarFAX
   sTag = sTag & "|EMAIL = " & mvarEMAIL
   sTag = sTag & "|ENDERECO = " & mvarENDERECO
   sTag = sTag & "|BAIRRO = " & mvarBAIRRO
   sTag = sTag & "|CIDADE = " & mvarCIDADE
   sTag = sTag & "|ESTADO = " & mvarESTADO
   sTag = sTag & "|CEP = " & mvarCEP
   sTag = sTag & "|CARGO = " & mvarCARGO
   sTag = sTag & "|DTCADASTRO = " & mvarDTCADASTRO
   sTag = sTag & "|PJ = " & mvarPJ
   sTag = sTag & "|PAIS = " & mvarPAIS
   sTag = sTag & "|OBS = " & mvarOBS
   sTag = sTag & "|NOME = " & mvarNOME
   sTag = sTag & "|DTNASC = " & mvarDTNASC
   sTag = sTag & "|ATIVO = " & mvarATIVO
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
 
   Sql = "Update OCONTATO Set " & VbNewLine 
   Sql = Sql & " IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDCONTATO  = " & SqlNum(mvarIDCONTATO) & VbNewLine 
   Sql = Sql & " , IDCOLIGADA = " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & " , IDLOJA0    = " & SqlNum(mvarIDLOJA0) & VbNewLine 
   Sql = Sql & " , EMPRESA    = " & SqlStr(mvarEMPRESA) & VbNewLine 
   Sql = Sql & " , REGISTRO   = " & SqlStr(mvarREGISTRO) & VbNewLine 
   Sql = Sql & " , TEL1       = " & SqlStr(mvarTEL1) & VbNewLine 
   Sql = Sql & " , TEL2       = " & SqlStr(mvarTEL2) & VbNewLine 
   Sql = Sql & " , FAX        = " & SqlStr(mvarFAX) & VbNewLine 
   Sql = Sql & " , EMAIL      = " & SqlStr(mvarEMAIL) & VbNewLine 
   Sql = Sql & " , ENDERECO   = " & SqlStr(mvarENDERECO) & VbNewLine 
   Sql = Sql & " , BAIRRO     = " & SqlStr(mvarBAIRRO) & VbNewLine 
   Sql = Sql & " , CIDADE     = " & SqlStr(mvarCIDADE) & VbNewLine 
   Sql = Sql & " , ESTADO     = " & SqlStr(mvarESTADO) & VbNewLine 
   Sql = Sql & " , CEP        = " & SqlStr(mvarCEP) & VbNewLine 
   Sql = Sql & " , CARGO      = " & SqlStr(mvarCARGO) & VbNewLine 
   Sql = Sql & " , DTCADASTRO = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , PJ         = " & SqlNum(mvarPJ) & VbNewLine 
   Sql = Sql & " , PAIS       = " & SqlStr(mvarPAIS) & VbNewLine 
   Sql = Sql & " , OBS        = " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & " , NOME       = " & SqlStr(mvarNOME) & VbNewLine 
   Sql = Sql & " , DTNASC     = " & SqlDate(mvarDTNASC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , ATIVO      = " & SqlNum(mvarATIVO) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP  = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP  = GetDate()" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP = 1" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OCONTATO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDCONTATO = " & SqlNum(mvarIDCONTATO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDCONTATO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDCONTATO, IDCOLIGADA, IDLOJA0" & VbNewLine 
   Sql = Sql & ", EMPRESA, REGISTRO, TEL1, TEL2, FAX" & VbNewLine 
   Sql = Sql & ", EMAIL, ENDERECO, BAIRRO, CIDADE, ESTADO" & VbNewLine 
   Sql = Sql & ", CEP, CARGO, DTCADASTRO, PJ, PAIS" & VbNewLine 
   Sql = Sql & ", OBS, NOME, DTNASC, ATIVO, TIMESTAMP" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP" & VbNewLine 
   Sql = Sql &" From OCONTATO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDCONTATO)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDCONTATO = " & SqlNum(mvarIDCONTATO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCONTATO) Then Sql = Sql & " IDCONTATO = " & SqlNum(Cstr(Ch_IDCONTATO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDCONTATO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDCONTATO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCONTATO = XVal(!IDCONTATO & "")
      mvarIDCOLIGADA = XVal(!IDCOLIGADA & "")
      mvarIDLOJA0 = XVal(!IDLOJA0 & "")
      mvarEMPRESA = !EMPRESA & ""
      mvarREGISTRO = !REGISTRO & ""
      mvarTEL1 = !TEL1 & ""
      mvarTEL2 = !TEL2 & ""
      mvarFAX = !FAX & ""
      mvarEMAIL = !EMAIL & ""
      mvarENDERECO = !ENDERECO & ""
      mvarBAIRRO = !BAIRRO & ""
      mvarCIDADE = !CIDADE & ""
      mvarESTADO = !ESTADO & ""
      mvarCEP = !CEP & ""
      mvarCARGO = !CARGO & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarPJ = XVal(!PJ & "")
      mvarPAIS = !PAIS & ""
      mvarOBS = !OBS & ""
      mvarNOME = !NOME & ""
      mvarDTNASC = xDate(!DTNASC & "", True)
      mvarATIVO = XVal(!ATIVO & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDCONTATO = 0
   mvarIDCOLIGADA = 0
   mvarIDLOJA0 = 0
   mvarEMPRESA = ""
   mvarREGISTRO = ""
   mvarTEL1 = ""
   mvarTEL2 = ""
   mvarFAX = ""
   mvarEMAIL = ""
   mvarENDERECO = ""
   mvarBAIRRO = ""
   mvarCIDADE = ""
   mvarESTADO = ""
   mvarCEP = ""
   mvarCARGO = ""
   mvarDTCADASTRO = ""
   mvarPJ = 0
   mvarPAIS = ""
   mvarOBS = ""
   mvarNOME = ""
   mvarDTNASC = ""
   mvarATIVO = 0
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDCONTATO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDCONTATO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OCONTATO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDCONTATO = " & SqlNum(Ch_IDCONTATO)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDCONTATO = " & SqlNum(mvarIDCONTATO)
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
