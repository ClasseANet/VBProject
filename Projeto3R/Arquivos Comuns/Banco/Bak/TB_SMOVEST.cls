VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_SMOVEST"
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
Private mvarIDMOV        As Double
Private mvarIDEST        As Double
Private mvarIDPROD       As Double
Private mvarDTMOV        As String
Private mvarQTDITEM      As Double
Private mvarUNIDCONTROLE As String
Private mvarTPDOC        As String
Private mvarNUMDOC       As String
Private mvarIDFOR        As Double
Private mvarITEMDOC      As Double
Private mvarFAVORECIDO   As String
Private mvarTPLAN        As String
Private mvarTPTRANSA     As String
Private mvarIDTRANSF     As Double
Private mvarSALDO        As Double
Private mvarIDVENDA      As Double
Private mvarOBS          As String
Private mvarFLGDELETE    As Double
Private mvarDTCADASTRO   As String
Private mvarCODSIT       As String
Private mvarALTERSTAMP   As Double
Private mvarTIMESTAMP    As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDMOV(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMOV <> vData)
   mvarIDMOV = vData
End Property
Public Property Get IDMOV() As Double
   IDMOV = mvarIDMOV
End Property
Public Property Let IDEST(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEST <> vData)
   mvarIDEST = vData
End Property
Public Property Get IDEST() As Double
   IDEST = mvarIDEST
End Property
Public Property Let IDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROD <> vData)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Double
   IDPROD = mvarIDPROD
End Property
Public Property Let DTMOV(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTMOV <> vData)
   mvarDTMOV = vData
End Property
Public Property Get DTMOV() As String
   DTMOV = mvarDTMOV
End Property
Public Property Let QTDITEM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDITEM <> vData)
   mvarQTDITEM = vData
End Property
Public Property Get QTDITEM() As Double
   QTDITEM = mvarQTDITEM
End Property
Public Property Let UNIDCONTROLE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUNIDCONTROLE <> vData)
   mvarUNIDCONTROLE = vData
End Property
Public Property Get UNIDCONTROLE() As String
   UNIDCONTROLE = mvarUNIDCONTROLE
End Property
Public Property Let TPDOC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPDOC <> vData)
   mvarTPDOC = vData
End Property
Public Property Get TPDOC() As String
   TPDOC = mvarTPDOC
End Property
Public Property Let NUMDOC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMDOC <> vData)
   mvarNUMDOC = vData
End Property
Public Property Get NUMDOC() As String
   NUMDOC = mvarNUMDOC
End Property
Public Property Let IDFOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFOR <> vData)
   mvarIDFOR = vData
End Property
Public Property Get IDFOR() As Double
   IDFOR = mvarIDFOR
End Property
Public Property Let ITEMDOC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarITEMDOC <> vData)
   mvarITEMDOC = vData
End Property
Public Property Get ITEMDOC() As Double
   ITEMDOC = mvarITEMDOC
End Property
Public Property Let FAVORECIDO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFAVORECIDO <> vData)
   mvarFAVORECIDO = vData
End Property
Public Property Get FAVORECIDO() As String
   FAVORECIDO = mvarFAVORECIDO
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
Public Property Let SALDO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSALDO <> vData)
   mvarSALDO = vData
End Property
Public Property Get SALDO() As Double
   SALDO = mvarSALDO
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
End Property
Public Property Let FLGDELETE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGDELETE <> vData)
   mvarFLGDELETE = vData
End Property
Public Property Get FLGDELETE() As Double
   FLGDELETE = mvarFLGDELETE
End Property
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let CODSIT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIT <> vData)
   mvarCODSIT = vData
End Property
Public Property Get CODSIT() As String
   CODSIT = mvarCODSIT
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update SMOVEST Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDMOV      = " & SqlNum(mvarIDMOV) & VbNewLine 
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
 
   Sql = "Insert Into SMOVEST (IDLOJA, IDMOV, IDEST, IDPROD" & VbNewLine 
   Sql = Sql & ", DTMOV, QTDITEM, UNIDCONTROLE, TPDOC, NUMDOC" & VbNewLine 
   Sql = Sql & ", IDFOR, ITEMDOC, FAVORECIDO, TPLAN, TPTRANSA" & VbNewLine 
   Sql = Sql & ", IDTRANSF, SALDO, IDVENDA, OBS, FLGDELETE" & VbNewLine 
   Sql = Sql & ", DTCADASTRO, CODSIT, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDMOV),0)+1 From SMOVEST Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDMOV) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDEST) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTMOV, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDITEM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUNIDCONTROLE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTPDOC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNUMDOC) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFOR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarITEMDOC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFAVORECIDO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTPLAN) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTPTRANSA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTRANSF) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSALDO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGDELETE) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODSIT) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDMOV) As String
   Dim Sql As String
 
   Sql = "Delete From SMOVEST" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDMOV) Then Sql = Sql & " IDMOV = " & SqlNum(Cstr(Ch_IDMOV)) & " AND " & VbNewLine 
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
 
   Sql = "Update SMOVEST Set " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDMOV        = " & SqlNum(mvarIDMOV) & VbNewLine 
   Sql = Sql & " , IDEST        = " & SqlNum(mvarIDEST) & VbNewLine 
   Sql = Sql & " , IDPROD       = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " , DTMOV        = " & SqlDate(mvarDTMOV, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , QTDITEM      = " & SqlNum(mvarQTDITEM) & VbNewLine 
   Sql = Sql & " , UNIDCONTROLE = " & SqlStr(mvarUNIDCONTROLE) & VbNewLine 
   Sql = Sql & " , TPDOC        = " & SqlStr(mvarTPDOC) & VbNewLine 
   Sql = Sql & " , NUMDOC       = " & SqlStr(mvarNUMDOC) & VbNewLine 
   Sql = Sql & " , IDFOR        = " & SqlNum(mvarIDFOR) & VbNewLine 
   Sql = Sql & " , ITEMDOC      = " & SqlNum(mvarITEMDOC) & VbNewLine 
   Sql = Sql & " , FAVORECIDO   = " & SqlStr(mvarFAVORECIDO) & VbNewLine 
   Sql = Sql & " , TPLAN        = " & SqlStr(mvarTPLAN) & VbNewLine 
   Sql = Sql & " , TPTRANSA     = " & SqlStr(mvarTPTRANSA) & VbNewLine 
   Sql = Sql & " , IDTRANSF     = " & SqlNum(mvarIDTRANSF) & VbNewLine 
   Sql = Sql & " , SALDO        = " & SqlNum(mvarSALDO) & VbNewLine 
   Sql = Sql & " , IDVENDA      = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " , OBS          = " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & " , FLGDELETE    = " & SqlNum(mvarFLGDELETE) & VbNewLine 
   Sql = Sql & " , DTCADASTRO   = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , CODSIT       = " & SqlStr(mvarCODSIT) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP   = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP    = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'SMOVEST'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDMOV      = " & SqlNum(mvarIDMOV) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDMOV, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDMOV, IDEST, IDPROD" & VbNewLine 
   Sql = Sql & ", DTMOV, QTDITEM, UNIDCONTROLE, TPDOC, NUMDOC" & VbNewLine 
   Sql = Sql & ", IDFOR, ITEMDOC, FAVORECIDO, TPLAN, TPTRANSA" & VbNewLine 
   Sql = Sql & ", IDTRANSF, SALDO, IDVENDA, OBS, FLGDELETE" & VbNewLine 
   Sql = Sql & ", DTCADASTRO, CODSIT, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From SMOVEST" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDMOV)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDMOV = " & SqlNum(mvarIDMOV) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDMOV) Then Sql = Sql & " IDMOV = " & SqlNum(Cstr(Ch_IDMOV)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDMOV, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDMOV, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDMOV = XVal(!IDMOV & "")
      mvarIDEST = XVal(!IDEST & "")
      mvarIDPROD = XVal(!IDPROD & "")
      mvarDTMOV = xDate(!DTMOV & "", True)
      mvarQTDITEM = XVal(!QTDITEM & "")
      mvarUNIDCONTROLE = !UNIDCONTROLE & ""
      mvarTPDOC = !TPDOC & ""
      mvarNUMDOC = !NUMDOC & ""
      mvarIDFOR = XVal(!IDFOR & "")
      mvarITEMDOC = XVal(!ITEMDOC & "")
      mvarFAVORECIDO = !FAVORECIDO & ""
      mvarTPLAN = !TPLAN & ""
      mvarTPTRANSA = !TPTRANSA & ""
      mvarIDTRANSF = XVal(!IDTRANSF & "")
      mvarSALDO = XVal(!SALDO & "")
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarOBS = !OBS & ""
      mvarFLGDELETE = XVal(!FLGDELETE & "")
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarCODSIT = !CODSIT & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDMOV), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDMOV As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update SMOVEST Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDMOV = " & SqlNum(Ch_IDMOV)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDMOV = " & SqlNum(mvarIDMOV)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
