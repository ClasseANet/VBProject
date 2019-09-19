VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_RBANCOHH"
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

Private mvarIDLOJA        As Double
Private mvarIDFUNCIONARIO As Double
Private mvarDTPONTO       As String
Private mvarHHINI         As String
Private mvarHHFIM         As String
Private mvarHHESPERADO    As Double
Private mvarHHTRAB        As Double
Private mvarSALDODIA      As Double
Private mvarIDMOVHH       As Double
Private mvarHHABONADO     As Double
Private mvarIDABONO       As Double
Private mvarFLGDIA        As Double
Private mvarHHREFEICAO    As Double
Private mvarALTERSTAMP    As Double
Private mvarTIMESTAMP     As String
Private mvarACUMULADO     As Double
Private mvarFLGFALTA      As Double
Private mvarOBS           As String
Private mvarFLGZERASALDO  As Double
Private mvarAPROVADO      As Double
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
Public Property Let DTPONTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTPONTO <> vData)
   mvarDTPONTO = vData
End Property
Public Property Get DTPONTO() As String
   DTPONTO = mvarDTPONTO
End Property
Public Property Let HHINI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHHINI <> vData)
   mvarHHINI = vData
End Property
Public Property Get HHINI() As String
   HHINI = mvarHHINI
End Property
Public Property Let HHFIM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHHFIM <> vData)
   mvarHHFIM = vData
End Property
Public Property Get HHFIM() As String
   HHFIM = mvarHHFIM
End Property
Public Property Let HHESPERADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarHHESPERADO <> vData)
   mvarHHESPERADO = vData
End Property
Public Property Get HHESPERADO() As Double
   HHESPERADO = mvarHHESPERADO
End Property
Public Property Let HHTRAB(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarHHTRAB <> vData)
   mvarHHTRAB = vData
End Property
Public Property Get HHTRAB() As Double
   HHTRAB = mvarHHTRAB
End Property
Public Property Let SALDODIA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSALDODIA <> vData)
   mvarSALDODIA = vData
End Property
Public Property Get SALDODIA() As Double
   SALDODIA = mvarSALDODIA
End Property
Public Property Let IDMOVHH(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMOVHH <> vData)
   mvarIDMOVHH = vData
End Property
Public Property Get IDMOVHH() As Double
   IDMOVHH = mvarIDMOVHH
End Property
Public Property Let HHABONADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarHHABONADO <> vData)
   mvarHHABONADO = vData
End Property
Public Property Get HHABONADO() As Double
   HHABONADO = mvarHHABONADO
End Property
Public Property Let IDABONO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDABONO <> vData)
   mvarIDABONO = vData
End Property
Public Property Get IDABONO() As Double
   IDABONO = mvarIDABONO
End Property
Public Property Let FLGDIA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGDIA <> vData)
   mvarFLGDIA = vData
End Property
Public Property Get FLGDIA() As Double
   FLGDIA = mvarFLGDIA
End Property
Public Property Let HHREFEICAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarHHREFEICAO <> vData)
   mvarHHREFEICAO = vData
End Property
Public Property Get HHREFEICAO() As Double
   HHREFEICAO = mvarHHREFEICAO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update RBANCOHH Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & VbNewLine 
   Sql = Sql & " And IDMOVHH     = " & SqlNum(mvarIDMOVHH) & VbNewLine 
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
Public Property Let ACUMULADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarACUMULADO <> vData)
   mvarACUMULADO = vData
End Property
Public Property Get ACUMULADO() As Double
   ACUMULADO = mvarACUMULADO
End Property
Public Property Let FLGFALTA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGFALTA <> vData)
   mvarFLGFALTA = vData
End Property
Public Property Get FLGFALTA() As Double
   FLGFALTA = mvarFLGFALTA
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
End Property
Public Property Let FLGZERASALDO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGZERASALDO <> vData)
   mvarFLGZERASALDO = vData
End Property
Public Property Get FLGZERASALDO() As Double
   FLGZERASALDO = mvarFLGZERASALDO
End Property
Public Property Let APROVADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarAPROVADO <> vData)
   mvarAPROVADO = vData
End Property
Public Property Get APROVADO() As Double
   APROVADO = mvarAPROVADO
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
 
   Sql = "Insert Into RBANCOHH (IDLOJA, IDFUNCIONARIO, DTPONTO, HHINI" & VbNewLine 
   Sql = Sql & ", HHFIM, HHESPERADO, HHTRAB, SALDODIA, IDMOVHH" & VbNewLine 
   Sql = Sql & ", HHABONADO, IDABONO, FLGDIA, HHREFEICAO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP, ACUMULADO, FLGFALTA, OBS, FLGZERASALDO" & VbNewLine 
   Sql = Sql & ", APROVADO) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFUNCIONARIO) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTPONTO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarHHINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarHHFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarHHESPERADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarHHTRAB) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSALDODIA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDMOVHH),0)+1 From RBANCOHH Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDMOVHH) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarHHABONADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDABONO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGDIA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarHHREFEICAO) & VbNewLine 
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
   Sql = Sql & ", " & SqlNum(mvarACUMULADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGFALTA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGZERASALDO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarAPROVADO) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDFUNCIONARIO, Optional Ch_IDMOVHH, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From RBANCOHH" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDFUNCIONARIO)  And IsMissing(Ch_IDMOVHH)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & " AND " & VbNewLine 
      Sql = Sql & " IDMOVHH = " & SqlNum(mvarIDMOVHH) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFUNCIONARIO) Then Sql = Sql & " IDFUNCIONARIO = " & SqlNum(Cstr(Ch_IDFUNCIONARIO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDMOVHH) Then Sql = Sql & " IDMOVHH = " & SqlNum(Cstr(Ch_IDMOVHH)) & " AND " & VbNewLine 
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
   sTag = sTag & "|DTPONTO = " & mvarDTPONTO
   sTag = sTag & "|HHINI = " & mvarHHINI
   sTag = sTag & "|HHFIM = " & mvarHHFIM
   sTag = sTag & "|HHESPERADO = " & mvarHHESPERADO
   sTag = sTag & "|HHTRAB = " & mvarHHTRAB
   sTag = sTag & "|SALDODIA = " & mvarSALDODIA
   sTag = sTag & "|IDMOVHH = " & mvarIDMOVHH
   sTag = sTag & "|HHABONADO = " & mvarHHABONADO
   sTag = sTag & "|IDABONO = " & mvarIDABONO
   sTag = sTag & "|FLGDIA = " & mvarFLGDIA
   sTag = sTag & "|HHREFEICAO = " & mvarHHREFEICAO
   sTag = sTag & "|ACUMULADO = " & mvarACUMULADO
   sTag = sTag & "|FLGFALTA = " & mvarFLGFALTA
   sTag = sTag & "|OBS = " & mvarOBS
   sTag = sTag & "|FLGZERASALDO = " & mvarFLGZERASALDO
   sTag = sTag & "|APROVADO = " & mvarAPROVADO
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
 
   Sql = "Update RBANCOHH Set " & VbNewLine 
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & VbNewLine 
   Sql = Sql & " , DTPONTO       = " & SqlDate(mvarDTPONTO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , HHINI         = " & SqlDate(mvarHHINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , HHFIM         = " & SqlDate(mvarHHFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , HHESPERADO    = " & SqlNum(mvarHHESPERADO) & VbNewLine 
   Sql = Sql & " , HHTRAB        = " & SqlNum(mvarHHTRAB) & VbNewLine 
   Sql = Sql & " , SALDODIA      = " & SqlNum(mvarSALDODIA) & VbNewLine 
   Sql = Sql & " , IDMOVHH       = " & SqlNum(mvarIDMOVHH) & VbNewLine 
   Sql = Sql & " , HHABONADO     = " & SqlNum(mvarHHABONADO) & VbNewLine 
   Sql = Sql & " , IDABONO       = " & SqlNum(mvarIDABONO) & VbNewLine 
   Sql = Sql & " , FLGDIA        = " & SqlNum(mvarFLGDIA) & VbNewLine 
   Sql = Sql & " , HHREFEICAO    = " & SqlNum(mvarHHREFEICAO) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP    = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP    = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP     = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP     = GetDate()" & vbNewLine
   End If
   Sql = Sql & " , ACUMULADO     = " & SqlNum(mvarACUMULADO) & VbNewLine 
   Sql = Sql & " , FLGFALTA      = " & SqlNum(mvarFLGFALTA) & VbNewLine 
   Sql = Sql & " , OBS           = " & SqlStr(mvarOBS) & VbNewLine 
   Sql = Sql & " , FLGZERASALDO  = " & SqlNum(mvarFLGZERASALDO) & VbNewLine 
   Sql = Sql & " , APROVADO      = " & SqlNum(mvarAPROVADO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'RBANCOHH'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & VbNewLine 
   Sql = Sql & " And IDMOVHH     = " & SqlNum(mvarIDMOVHH) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDFUNCIONARIO, Optional Ch_IDMOVHH, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDFUNCIONARIO, DTPONTO, HHINI" & VbNewLine 
   Sql = Sql & ", HHFIM, HHESPERADO, HHTRAB, SALDODIA, IDMOVHH" & VbNewLine 
   Sql = Sql & ", HHABONADO, IDABONO, FLGDIA, HHREFEICAO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP, ACUMULADO, FLGFALTA, OBS, FLGZERASALDO" & VbNewLine 
   Sql = Sql & ", APROVADO" & VbNewLine 
   Sql = Sql &" From RBANCOHH" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDFUNCIONARIO)  And IsMissing(Ch_IDMOVHH)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO) & " AND " & VbNewLine 
      Sql = Sql & " IDMOVHH = " & SqlNum(mvarIDMOVHH) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFUNCIONARIO) Then Sql = Sql & " IDFUNCIONARIO = " & SqlNum(Cstr(Ch_IDFUNCIONARIO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDMOVHH) Then Sql = Sql & " IDMOVHH = " & SqlNum(Cstr(Ch_IDMOVHH)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDFUNCIONARIO, Optional Ch_IDMOVHH, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDFUNCIONARIO, Ch_IDMOVHH, Ch_WHERE, Ch_ORDERBY)
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
      mvarDTPONTO = xDate(!DTPONTO & "", True)
      mvarHHINI = xDate(!HHINI & "", True)
      mvarHHFIM = xDate(!HHFIM & "", True)
      mvarHHESPERADO = XVal(!HHESPERADO & "")
      mvarHHTRAB = XVal(!HHTRAB & "")
      mvarSALDODIA = XVal(!SALDODIA & "")
      mvarIDMOVHH = XVal(!IDMOVHH & "")
      mvarHHABONADO = XVal(!HHABONADO & "")
      mvarIDABONO = XVal(!IDABONO & "")
      mvarFLGDIA = XVal(!FLGDIA & "")
      mvarHHREFEICAO = XVal(!HHREFEICAO & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarACUMULADO = XVal(!ACUMULADO & "")
      mvarFLGFALTA = XVal(!FLGFALTA & "")
      mvarOBS = !OBS & ""
      mvarFLGZERASALDO = XVal(!FLGZERASALDO & "")
      mvarAPROVADO = XVal(!APROVADO & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDFUNCIONARIO = 0
   mvarDTPONTO = ""
   mvarHHINI = ""
   mvarHHFIM = ""
   mvarHHESPERADO = 0
   mvarHHTRAB = 0
   mvarSALDODIA = 0
   mvarIDMOVHH = 0
   mvarHHABONADO = 0
   mvarIDABONO = 0
   mvarFLGDIA = 0
   mvarHHREFEICAO = 0
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""
   mvarACUMULADO = 0
   mvarFLGFALTA = 0
   mvarOBS = ""
   mvarFLGZERASALDO = 0
   mvarAPROVADO = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDFUNCIONARIO, mvarIDMOVHH), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDFUNCIONARIO As String, Ch_IDMOVHH As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update RBANCOHH Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDFUNCIONARIO = " & SqlNum(Ch_IDFUNCIONARIO)
   Sql = Sql & " , IDMOVHH = " & SqlNum(Ch_IDMOVHH)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDFUNCIONARIO = " & SqlNum(mvarIDFUNCIONARIO)
   Sql = Sql & " and IDMOVHH = " & SqlNum(mvarIDMOVHH)
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
