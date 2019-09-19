VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_FRECIBO"
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
Private mvarSERIE           As String
Private mvarIDRECIBO        As Double
Private mvarIDVENDA         As Double
Private mvarDTEMISSAO       As String
Private mvarVLTOTAL         As Double
Private mvarTIPO            As Double
Private mvarIDLOTE          As Double
Private mvarSTATUS          As Double
Private mvarNATUREZAOP      As Double
Private mvarASSINATURA      As String
Private mvarREGESPECIALTRIB As Double
Private mvarSIMPLES         As Double
Private mvarIDRECIBO0       As Double
Private mvarSERIE0          As String
Private mvarTRIBUTACAO      As Double
Private mvarINCENTIVOCULT   As Double
Private mvarCODSERVFEDERAL  As String
Private mvarCODSERVMUNIC    As String
Private mvarDSCSERV         As String
Private mvarALTERSTAMP      As Double
Private mvarTIMESTAMP       As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let SERIE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSERIE <> vData)
   mvarSERIE = vData
End Property
Public Property Get SERIE() As String
   SERIE = mvarSERIE
End Property
Public Property Let IDRECIBO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDRECIBO <> vData)
   mvarIDRECIBO = vData
End Property
Public Property Get IDRECIBO() As Double
   IDRECIBO = mvarIDRECIBO
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let DTEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEMISSAO <> vData)
   mvarDTEMISSAO = vData
End Property
Public Property Get DTEMISSAO() As String
   DTEMISSAO = mvarDTEMISSAO
End Property
Public Property Let VLTOTAL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLTOTAL <> vData)
   mvarVLTOTAL = vData
End Property
Public Property Get VLTOTAL() As Double
   VLTOTAL = mvarVLTOTAL
End Property
Public Property Let TIPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPO <> vData)
   mvarTIPO = vData
End Property
Public Property Get TIPO() As Double
   TIPO = mvarTIPO
End Property
Public Property Let IDLOTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOTE <> vData)
   mvarIDLOTE = vData
End Property
Public Property Get IDLOTE() As Double
   IDLOTE = mvarIDLOTE
End Property
Public Property Let STATUS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSTATUS <> vData)
   mvarSTATUS = vData
End Property
Public Property Get STATUS() As Double
   STATUS = mvarSTATUS
End Property
Public Property Let NATUREZAOP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNATUREZAOP <> vData)
   mvarNATUREZAOP = vData
End Property
Public Property Get NATUREZAOP() As Double
   NATUREZAOP = mvarNATUREZAOP
End Property
Public Property Let ASSINATURA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarASSINATURA <> vData)
   mvarASSINATURA = vData
End Property
Public Property Get ASSINATURA() As String
   ASSINATURA = mvarASSINATURA
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
Public Property Let IDRECIBO0(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDRECIBO0 <> vData)
   mvarIDRECIBO0 = vData
End Property
Public Property Get IDRECIBO0() As Double
   IDRECIBO0 = mvarIDRECIBO0
End Property
Public Property Let SERIE0(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSERIE0 <> vData)
   mvarSERIE0 = vData
End Property
Public Property Get SERIE0() As String
   SERIE0 = mvarSERIE0
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
Public Property Let CODSERVFEDERAL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSERVFEDERAL <> vData)
   mvarCODSERVFEDERAL = vData
End Property
Public Property Get CODSERVFEDERAL() As String
   CODSERVFEDERAL = mvarCODSERVFEDERAL
End Property
Public Property Let CODSERVMUNIC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSERVMUNIC <> vData)
   mvarCODSERVMUNIC = vData
End Property
Public Property Get CODSERVMUNIC() As String
   CODSERVMUNIC = mvarCODSERVMUNIC
End Property
Public Property Let DSCSERV(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCSERV <> vData)
   mvarDSCSERV = vData
End Property
Public Property Get DSCSERV() As String
   DSCSERV = mvarDSCSERV
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update FRECIBO Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA            = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And SERIE         = " & SqlStr(mvarSERIE) & vbNewLine
   Sql = Sql & " And IDRECIBO      = " & SqlNum(mvarIDRECIBO) & vbNewLine
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
 
   Sql = "Insert Into FRECIBO (IDLOJA, SERIE, IDRECIBO, IDVENDA" & vbNewLine
   Sql = Sql & ", DTEMISSAO, VLTOTAL, TIPO, IDLOTE, STATUS" & vbNewLine
   Sql = Sql & ", NATUREZAOP, ASSINATURA, REGESPECIALTRIB, SIMPLES, IDRECIBO0" & vbNewLine
   Sql = Sql & ", SERIE0, TRIBUTACAO, INCENTIVOCULT, CODSERVFEDERAL, CODSERVMUNIC" & vbNewLine
   Sql = Sql & ", DSCSERV, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarSERIE) & vbNewLine
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDRECIBO),0)+1 From FRECIBO Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And SERIE = " & SqlStr(mvarSERIE) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDRECIBO) & vbNewLine
   End If
   Sql = Sql & ", " & IIf(mvarIDVENDA = 0, "Null", SqlNum(mvarIDVENDA)) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLTOTAL) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarTIPO) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDLOTE = 0, "Null", SqlNum(mvarIDLOTE)) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarSTATUS) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarNATUREZAOP) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarASSINATURA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarREGESPECIALTRIB) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarSIMPLES) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDRECIBO0) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarSERIE0) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarTRIBUTACAO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarINCENTIVOCULT) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCODSERVFEDERAL) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCODSERVMUNIC) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarDSCSERV) & vbNewLine
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
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_IDRECIBO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From FRECIBO" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_SERIE) And IsMissing(Ch_IDRECIBO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " SERIE = " & SqlStr(mvarSERIE) & " AND " & vbNewLine
      Sql = Sql & " IDRECIBO = " & SqlNum(mvarIDRECIBO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_SERIE) Then Sql = Sql & " SERIE = " & SqlStr(CStr(Ch_SERIE)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDRECIBO) Then Sql = Sql & " IDRECIBO = " & SqlNum(CStr(Ch_IDRECIBO)) & " AND " & vbNewLine
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
   sTag = sTag & "|SERIE = " & mvarSERIE
   sTag = sTag & "|IDRECIBO = " & mvarIDRECIBO
   sTag = sTag & "|IDVENDA = " & mvarIDVENDA
   sTag = sTag & "|DTEMISSAO = " & mvarDTEMISSAO
   sTag = sTag & "|VLTOTAL = " & mvarVLTOTAL
   sTag = sTag & "|TIPO = " & mvarTIPO
   sTag = sTag & "|IDLOTE = " & mvarIDLOTE
   sTag = sTag & "|STATUS = " & mvarSTATUS
   sTag = sTag & "|NATUREZAOP = " & mvarNATUREZAOP
   sTag = sTag & "|ASSINATURA = " & mvarASSINATURA
   sTag = sTag & "|REGESPECIALTRIB = " & mvarREGESPECIALTRIB
   sTag = sTag & "|SIMPLES = " & mvarSIMPLES
   sTag = sTag & "|IDRECIBO0 = " & mvarIDRECIBO0
   sTag = sTag & "|SERIE0 = " & mvarSERIE0
   sTag = sTag & "|TRIBUTACAO = " & mvarTRIBUTACAO
   sTag = sTag & "|INCENTIVOCULT = " & mvarINCENTIVOCULT
   sTag = sTag & "|CODSERVFEDERAL = " & mvarCODSERVFEDERAL
   sTag = sTag & "|CODSERVMUNIC = " & mvarCODSERVMUNIC
   sTag = sTag & "|DSCSERV = " & mvarDSCSERV
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
 
   Sql = "Update FRECIBO Set " & vbNewLine
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , SERIE           = " & SqlStr(mvarSERIE) & vbNewLine
   Sql = Sql & " , IDRECIBO        = " & SqlNum(mvarIDRECIBO) & vbNewLine
   Sql = Sql & " , IDVENDA         = " & IIf(mvarIDVENDA = 0, "Null", SqlNum(mvarIDVENDA)) & vbNewLine
   Sql = Sql & " , DTEMISSAO       = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , VLTOTAL         = " & SqlNum(mvarVLTOTAL) & vbNewLine
   Sql = Sql & " , TIPO            = " & SqlNum(mvarTIPO) & vbNewLine
   Sql = Sql & " , IDLOTE          = " & IIf(mvarIDLOTE = 0, "Null", SqlNum(mvarIDLOTE)) & vbNewLine
   Sql = Sql & " , STATUS          = " & SqlNum(mvarSTATUS) & vbNewLine
   Sql = Sql & " , NATUREZAOP      = " & SqlNum(mvarNATUREZAOP) & vbNewLine
   Sql = Sql & " , ASSINATURA      = " & SqlStr(mvarASSINATURA) & vbNewLine
   Sql = Sql & " , REGESPECIALTRIB = " & SqlNum(mvarREGESPECIALTRIB) & vbNewLine
   Sql = Sql & " , SIMPLES         = " & SqlNum(mvarSIMPLES) & vbNewLine
   Sql = Sql & " , IDRECIBO0       = " & SqlNum(mvarIDRECIBO0) & vbNewLine
   Sql = Sql & " , SERIE0          = " & SqlStr(mvarSERIE0) & vbNewLine
   Sql = Sql & " , TRIBUTACAO      = " & SqlNum(mvarTRIBUTACAO) & vbNewLine
   Sql = Sql & " , INCENTIVOCULT   = " & SqlNum(mvarINCENTIVOCULT) & vbNewLine
   Sql = Sql & " , CODSERVFEDERAL  = " & SqlStr(mvarCODSERVFEDERAL) & vbNewLine
   Sql = Sql & " , CODSERVMUNIC    = " & SqlStr(mvarCODSERVMUNIC) & vbNewLine
   Sql = Sql & " , DSCSERV         = " & SqlStr(mvarDSCSERV) & vbNewLine
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
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FRECIBO'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA            = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And SERIE         = " & SqlStr(mvarSERIE) & vbNewLine
   Sql = Sql & " And IDRECIBO      = " & SqlNum(mvarIDRECIBO) & vbNewLine

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
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_IDRECIBO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, SERIE, IDRECIBO, IDVENDA" & vbNewLine
   Sql = Sql & ", DTEMISSAO, VLTOTAL, TIPO, IDLOTE, STATUS" & vbNewLine
   Sql = Sql & ", NATUREZAOP, ASSINATURA, REGESPECIALTRIB, SIMPLES, IDRECIBO0" & vbNewLine
   Sql = Sql & ", SERIE0, TRIBUTACAO, INCENTIVOCULT, CODSERVFEDERAL, CODSERVMUNIC" & vbNewLine
   Sql = Sql & ", DSCSERV, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From FRECIBO" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_SERIE) And IsMissing(Ch_IDRECIBO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " SERIE = " & SqlStr(mvarSERIE) & " AND " & vbNewLine
      Sql = Sql & " IDRECIBO = " & SqlNum(mvarIDRECIBO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_SERIE) Then Sql = Sql & " SERIE = " & SqlStr(CStr(Ch_SERIE)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDRECIBO) Then Sql = Sql & " IDRECIBO = " & SqlNum(CStr(Ch_IDRECIBO)) & " AND " & vbNewLine
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
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_IDRECIBO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_SERIE, Ch_IDRECIBO, Ch_WHERE, Ch_ORDERBY)
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
      mvarSERIE = !SERIE & ""
      mvarIDRECIBO = xVal(!IDRECIBO & "")
      mvarIDVENDA = xVal(!IDVENDA & "")
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarVLTOTAL = xVal(!VLTOTAL & "")
      mvarTIPO = xVal(!TIPO & "")
      mvarIDLOTE = xVal(!IDLOTE & "")
      mvarSTATUS = xVal(!STATUS & "")
      mvarNATUREZAOP = xVal(!NATUREZAOP & "")
      mvarASSINATURA = !ASSINATURA & ""
      mvarREGESPECIALTRIB = xVal(!REGESPECIALTRIB & "")
      mvarSIMPLES = xVal(!SIMPLES & "")
      mvarIDRECIBO0 = xVal(!IDRECIBO0 & "")
      mvarSERIE0 = !SERIE0 & ""
      mvarTRIBUTACAO = xVal(!TRIBUTACAO & "")
      mvarINCENTIVOCULT = xVal(!INCENTIVOCULT & "")
      mvarCODSERVFEDERAL = !CODSERVFEDERAL & ""
      mvarCODSERVMUNIC = !CODSERVMUNIC & ""
      mvarDSCSERV = !DSCSERV & ""
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarSERIE = ""
   mvarIDRECIBO = 0
   mvarIDVENDA = 0
   mvarDTEMISSAO = ""
   mvarVLTOTAL = 0
   mvarTIPO = 0
   mvarIDLOTE = 0
   mvarSTATUS = 0
   mvarNATUREZAOP = 0
   mvarASSINATURA = ""
   mvarREGESPECIALTRIB = 0
   mvarSIMPLES = 0
   mvarIDRECIBO0 = 0
   mvarSERIE0 = ""
   mvarTRIBUTACAO = 0
   mvarINCENTIVOCULT = 0
   mvarCODSERVFEDERAL = ""
   mvarCODSERVMUNIC = ""
   mvarDSCSERV = ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarSERIE, mvarIDRECIBO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_SERIE As String, Ch_IDRECIBO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FRECIBO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , SERIE = " & SqlStr(Ch_SERIE)
   Sql = Sql & " , IDRECIBO = " & SqlNum(Ch_IDRECIBO)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and SERIE = " & SqlStr(mvarSERIE)
   Sql = Sql & " and IDRECIBO = " & SqlNum(mvarIDRECIBO)
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
