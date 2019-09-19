VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_FLAN"
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
Private mvarIDCONTA      As Double
Private mvarIDLAN        As Double
Private mvarIDVENDA      As Double
Private mvarIDDESP       As Double
Private mvarIDSUBDESP    As Double
Private mvarIDPGTO       As Double
Private mvarNDOC         As String
Private mvarVALOR        As Double
Private mvarHISTORICO    As String
Private mvarNPARCELA     As Double
Private mvarTPARCELA     As Double
Private mvarIDFAVORECIDO As Double
Private mvarFAVORECIDO   As String
Private mvarTPFAVORECIDO As String
Private mvarDTCADASTRO   As String
Private mvarDTEMISSAO    As String
Private mvarDTVENCIMENTO As String
Private mvarDTBAIXA      As String
Private mvarDTCONTAB     As String
Private mvarTPLAN        As String
Private mvarTPTRANSA     As String
Private mvarIDTRANSF     As Double
Private mvarCODSIT       As String
Private mvarIDPAI        As Double
Private mvarFLGDELETE    As Double
Private mvarFLGEXPORT    As Double
Private mvarOBS          As String
Private mvarSALDO        As Double
Private mvarTPDOC        As String
Private mvarALTERSTAMP   As Double
Private mvarTIMESTAMP    As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCONTA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCONTA <> vData)
   mvarIDCONTA = vData
End Property
Public Property Get IDCONTA() As Double
   IDCONTA = mvarIDCONTA
End Property
Public Property Let IDLAN(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLAN <> vData)
   mvarIDLAN = vData
End Property
Public Property Get IDLAN() As Double
   IDLAN = mvarIDLAN
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDDESP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDESP <> vData)
   mvarIDDESP = vData
End Property
Public Property Get IDDESP() As Double
   IDDESP = mvarIDDESP
End Property
Public Property Let IDSUBDESP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSUBDESP <> vData)
   mvarIDSUBDESP = vData
End Property
Public Property Get IDSUBDESP() As Double
   IDSUBDESP = mvarIDSUBDESP
End Property
Public Property Let IDPGTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPGTO <> vData)
   mvarIDPGTO = vData
End Property
Public Property Get IDPGTO() As Double
   IDPGTO = mvarIDPGTO
End Property
Public Property Let NDOC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNDOC <> vData)
   mvarNDOC = vData
End Property
Public Property Get NDOC() As String
   NDOC = mvarNDOC
End Property
Public Property Let VALOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVALOR <> vData)
   mvarVALOR = vData
End Property
Public Property Get VALOR() As Double
   VALOR = mvarVALOR
End Property
Public Property Let HISTORICO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHISTORICO <> vData)
   mvarHISTORICO = vData
End Property
Public Property Get HISTORICO() As String
   HISTORICO = mvarHISTORICO
End Property
Public Property Let NPARCELA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNPARCELA <> vData)
   mvarNPARCELA = vData
End Property
Public Property Get NPARCELA() As Double
   NPARCELA = mvarNPARCELA
End Property
Public Property Let TPARCELA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTPARCELA <> vData)
   mvarTPARCELA = vData
End Property
Public Property Get TPARCELA() As Double
   TPARCELA = mvarTPARCELA
End Property
Public Property Let IDFAVORECIDO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFAVORECIDO <> vData)
   mvarIDFAVORECIDO = vData
End Property
Public Property Get IDFAVORECIDO() As Double
   IDFAVORECIDO = mvarIDFAVORECIDO
End Property
Public Property Let FAVORECIDO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFAVORECIDO <> vData)
   mvarFAVORECIDO = vData
End Property
Public Property Get FAVORECIDO() As String
   FAVORECIDO = mvarFAVORECIDO
End Property
Public Property Let TPFAVORECIDO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPFAVORECIDO <> vData)
   mvarTPFAVORECIDO = vData
End Property
Public Property Get TPFAVORECIDO() As String
   TPFAVORECIDO = mvarTPFAVORECIDO
End Property
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let DTEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEMISSAO <> vData)
   mvarDTEMISSAO = vData
End Property
Public Property Get DTEMISSAO() As String
   DTEMISSAO = mvarDTEMISSAO
End Property
Public Property Let DTVENCIMENTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTVENCIMENTO <> vData)
   mvarDTVENCIMENTO = vData
End Property
Public Property Get DTVENCIMENTO() As String
   DTVENCIMENTO = mvarDTVENCIMENTO
End Property
Public Property Let DTBAIXA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTBAIXA <> vData)
   mvarDTBAIXA = vData
End Property
Public Property Get DTBAIXA() As String
   DTBAIXA = mvarDTBAIXA
End Property
Public Property Let DTCONTAB(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCONTAB <> vData)
   mvarDTCONTAB = vData
End Property
Public Property Get DTCONTAB() As String
   DTCONTAB = mvarDTCONTAB
End Property
Public Property Let TPLAN(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPLAN <> vData)
   mvarTPLAN = vData
End Property
Public Property Get TPLAN() As String
   TPLAN = mvarTPLAN
End Property
Public Property Let TPTRANSA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPTRANSA <> vData)
   mvarTPTRANSA = vData
End Property
Public Property Get TPTRANSA() As String
   TPTRANSA = mvarTPTRANSA
End Property
Public Property Let IDTRANSF(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTRANSF <> vData)
   mvarIDTRANSF = vData
End Property
Public Property Get IDTRANSF() As Double
   IDTRANSF = mvarIDTRANSF
End Property
Public Property Let CODSIT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIT <> vData)
   mvarCODSIT = vData
End Property
Public Property Get CODSIT() As String
   CODSIT = mvarCODSIT
End Property
Public Property Let IDPAI(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAI <> vData)
   mvarIDPAI = vData
End Property
Public Property Get IDPAI() As Double
   IDPAI = mvarIDPAI
End Property
Public Property Let FLGDELETE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGDELETE <> vData)
   mvarFLGDELETE = vData
End Property
Public Property Get FLGDELETE() As Double
   FLGDELETE = mvarFLGDELETE
End Property
Public Property Let FLGEXPORT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGEXPORT <> vData)
   mvarFLGEXPORT = vData
End Property
Public Property Get FLGEXPORT() As Double
   FLGEXPORT = mvarFLGEXPORT
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
End Property
Public Property Let SALDO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSALDO <> vData)
   mvarSALDO = vData
End Property
Public Property Get SALDO() As Double
   SALDO = mvarSALDO
End Property
Public Property Let TPDOC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPDOC <> vData)
   mvarTPDOC = vData
End Property
Public Property Get TPDOC() As String
   TPDOC = mvarTPDOC
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update FLAN Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDCONTA    = " & SqlNum(mvarIDCONTA) & vbNewLine
   Sql = Sql & " And IDLAN      = " & SqlNum(mvarIDLAN) & vbNewLine
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
Public Property Get QryInsert(Optional pAutoId As Boolean = True) As String
   Dim Sql As String
 
   Sql = "Insert Into FLAN (IDLOJA, IDCONTA, IDLAN, IDVENDA" & vbNewLine
   Sql = Sql & ", IDDESP, IDSUBDESP, IDPGTO, NDOC, VALOR" & vbNewLine
   Sql = Sql & ", HISTORICO, NPARCELA, TPARCELA, IDFAVORECIDO, FAVORECIDO" & vbNewLine
   Sql = Sql & ", TPFAVORECIDO, DTCADASTRO, DTEMISSAO, DTVENCIMENTO, DTBAIXA" & vbNewLine
   Sql = Sql & ", DTCONTAB, TPLAN, TPTRANSA, IDTRANSF, CODSIT" & vbNewLine
   Sql = Sql & ", IDPAI, FLGDELETE, FLGEXPORT, OBS, SALDO" & vbNewLine
   Sql = Sql & ", TPDOC, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDCONTA) & vbNewLine
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDLAN),0)+1 From FLAN Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDCONTA = " & SqlNum(mvarIDCONTA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDLAN) & vbNewLine
   End If
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDDESP = 0, "Null", SqlNum(mvarIDDESP)) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDSUBDESP = 0, "Null", SqlNum(mvarIDSUBDESP)) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDPGTO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarNDOC) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVALOR) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarHISTORICO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarNPARCELA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarTPARCELA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDFAVORECIDO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarFAVORECIDO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTPFAVORECIDO) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTVENCIMENTO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTBAIXA, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTCONTAB, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTPLAN) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTPTRANSA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDTRANSF) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCODSIT) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDPAI) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGDELETE) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGEXPORT) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarOBS) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarSALDO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTPDOC) & vbNewLine
   Sql = Sql & ", " & "1" & vbNewLine
   Sql = Sql & ", " & "GetDate()" & vbNewLine

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDCONTA, Optional Ch_IDLAN) As String
   Dim Sql As String
 
   Sql = "Delete From FLAN" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
   If Not IsMissing(Ch_IDCONTA) Then Sql = Sql & " IDCONTA = " & SqlNum(CStr(Ch_IDCONTA)) & " AND " & vbNewLine
   If Not IsMissing(Ch_IDLAN) Then Sql = Sql & " IDLAN = " & SqlNum(CStr(Ch_IDLAN)) & " AND " & vbNewLine
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   MyDelRow.QUERY = Sql & ";"
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update FLAN Set " & vbNewLine
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDCONTA      = " & SqlNum(mvarIDCONTA) & vbNewLine
   Sql = Sql & " , IDLAN        = " & SqlNum(mvarIDLAN) & vbNewLine
   Sql = Sql & " , IDVENDA      = " & SqlNum(mvarIDVENDA) & vbNewLine
   Sql = Sql & " , IDDESP       = " & IIf(mvarIDDESP = 0, "Null", SqlNum(mvarIDDESP)) & vbNewLine
   Sql = Sql & " , IDSUBDESP    = " & IIf(mvarIDSUBDESP = 0, "Null", SqlNum(mvarIDSUBDESP)) & vbNewLine
   Sql = Sql & " , IDPGTO       = " & SqlNum(mvarIDPGTO) & vbNewLine
   Sql = Sql & " , NDOC         = " & SqlStr(mvarNDOC) & vbNewLine
   Sql = Sql & " , VALOR        = " & SqlNum(mvarVALOR) & vbNewLine
   Sql = Sql & " , HISTORICO    = " & SqlStr(mvarHISTORICO) & vbNewLine
   Sql = Sql & " , NPARCELA     = " & SqlNum(mvarNPARCELA) & vbNewLine
   Sql = Sql & " , TPARCELA     = " & SqlNum(mvarTPARCELA) & vbNewLine
   Sql = Sql & " , IDFAVORECIDO = " & SqlNum(mvarIDFAVORECIDO) & vbNewLine
   Sql = Sql & " , FAVORECIDO   = " & SqlStr(mvarFAVORECIDO) & vbNewLine
   Sql = Sql & " , TPFAVORECIDO = " & SqlStr(mvarTPFAVORECIDO) & vbNewLine
   Sql = Sql & " , DTCADASTRO   = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , DTEMISSAO    = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , DTVENCIMENTO = " & SqlDate(mvarDTVENCIMENTO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , DTBAIXA      = " & SqlDate(mvarDTBAIXA, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , DTCONTAB     = " & SqlDate(mvarDTCONTAB, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , TPLAN        = " & SqlStr(mvarTPLAN) & vbNewLine
   Sql = Sql & " , TPTRANSA     = " & SqlStr(mvarTPTRANSA) & vbNewLine
   Sql = Sql & " , IDTRANSF     = " & SqlNum(mvarIDTRANSF) & vbNewLine
   Sql = Sql & " , CODSIT       = " & SqlStr(mvarCODSIT) & vbNewLine
   Sql = Sql & " , IDPAI        = " & SqlNum(mvarIDPAI) & vbNewLine
   Sql = Sql & " , FLGDELETE    = " & SqlNum(mvarFLGDELETE) & vbNewLine
   Sql = Sql & " , FLGEXPORT    = " & SqlNum(mvarFLGEXPORT) & vbNewLine
   Sql = Sql & " , OBS          = " & SqlStr(mvarOBS) & vbNewLine
   Sql = Sql & " , SALDO        = " & SqlNum(mvarSALDO) & vbNewLine
   Sql = Sql & " , TPDOC        = " & SqlStr(mvarTPDOC) & vbNewLine
   Sql = Sql & " , ALTERSTAMP   = " & "1" & vbNewLine
   Sql = Sql & " , TIMESTAMP    = " & "GetDate()" & vbNewLine
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FLAN'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDCONTA    = " & SqlNum(mvarIDCONTA) & vbNewLine
   Sql = Sql & " And IDLAN      = " & SqlNum(mvarIDLAN) & vbNewLine

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & vbNewLine
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & vbNewLine
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDCONTA, Optional Ch_IDLAN, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDCONTA, IDLAN, IDVENDA" & vbNewLine
   Sql = Sql & ", IDDESP, IDSUBDESP, IDPGTO, NDOC, VALOR" & vbNewLine
   Sql = Sql & ", HISTORICO, NPARCELA, TPARCELA, IDFAVORECIDO, FAVORECIDO" & vbNewLine
   Sql = Sql & ", TPFAVORECIDO, DTCADASTRO, DTEMISSAO, DTVENCIMENTO, DTBAIXA" & vbNewLine
   Sql = Sql & ", DTCONTAB, TPLAN, TPTRANSA, IDTRANSF, CODSIT" & vbNewLine
   Sql = Sql & ", IDPAI, FLGDELETE, FLGEXPORT, OBS, SALDO" & vbNewLine
   Sql = Sql & ", TPDOC, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From FLAN" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDCONTA) And IsMissing(Ch_IDLAN) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDCONTA = " & SqlNum(mvarIDCONTA) & " AND " & vbNewLine
      Sql = Sql & " IDLAN = " & SqlNum(mvarIDLAN) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDCONTA) Then Sql = Sql & " IDCONTA = " & SqlNum(CStr(Ch_IDCONTA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDLAN) Then Sql = Sql & " IDLAN = " & SqlNum(CStr(Ch_IDLAN)) & " AND " & vbNewLine
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
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDCONTA, Optional Ch_IDLAN, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDCONTA, Ch_IDLAN, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCONTA = xVal(!IDCONTA & "")
      mvarIDLAN = xVal(!IDLAN & "")
      mvarIDVENDA = xVal(!IDVENDA & "")
      mvarIDDESP = xVal(!IDDESP & "")
      mvarIDSUBDESP = xVal(!IDSUBDESP & "")
      mvarIDPGTO = xVal(!IDPGTO & "")
      mvarNDOC = !NDOC & ""
      mvarVALOR = xVal(!VALOR & "")
      mvarHISTORICO = !HISTORICO & ""
      mvarNPARCELA = xVal(!NPARCELA & "")
      mvarTPARCELA = xVal(!TPARCELA & "")
      mvarIDFAVORECIDO = xVal(!IDFAVORECIDO & "")
      mvarFAVORECIDO = !FAVORECIDO & ""
      mvarTPFAVORECIDO = !TPFAVORECIDO & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarDTVENCIMENTO = xDate(!DTVENCIMENTO & "", True)
      mvarDTBAIXA = xDate(!DTBAIXA & "", True)
      mvarDTCONTAB = xDate(!DTCONTAB & "", True)
      mvarTPLAN = !TPLAN & ""
      mvarTPTRANSA = !TPTRANSA & ""
      mvarIDTRANSF = xVal(!IDTRANSF & "")
      mvarCODSIT = !CODSIT & ""
      mvarIDPAI = xVal(!IDPAI & "")
      mvarFLGDELETE = xVal(!FLGDELETE & "")
      mvarFLGEXPORT = xVal(!FLGEXPORT & "")
      mvarOBS = !OBS & ""
      mvarSALDO = xVal(!SALDO & "")
      mvarTPDOC = !TPDOC & ""
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True, Optional pAutoId As Boolean = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId As Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDCONTA, mvarIDLAN), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDCONTA As String, Ch_IDLAN As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FLAN Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDCONTA = " & SqlNum(Ch_IDCONTA)
   Sql = Sql & " , IDLAN = " & SqlNum(Ch_IDLAN)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDCONTA = " & SqlNum(mvarIDCONTA)
   Sql = Sql & " and IDLAN = " & SqlNum(mvarIDLAN)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
