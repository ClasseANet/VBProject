VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CPROMOCAO"
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
Private mvarIDPROMO    As Double
Private mvarDSCPROMO   As String
Private mvarATIVO      As Double
Private mvarVLTOTAL    As Double
Private mvarVLDESC     As Double
Private mvarVALOR      As Double
Private mvarDTINI      As String
Private mvarDTFIM      As String
Private mvarDTINIV     As String
Private mvarDTFIMV     As String
Private mvarNPARCELA   As Double
Private mvarFLGSERV    As Double
Private mvarFLGTRAT    As Double
Private mvarFLGAREA    As Double
Private mvarSERVIN     As String
Private mvarTRATIN     As String
Private mvarAREAIN     As String
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDPROMO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROMO <> vData)
   mvarIDPROMO = vData
End Property
Public Property Get IDPROMO() As Double
   IDPROMO = mvarIDPROMO
End Property
Public Property Let DSCPROMO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCPROMO <> vData)
   mvarDSCPROMO = vData
End Property
Public Property Get DSCPROMO() As String
   DSCPROMO = mvarDSCPROMO
End Property
Public Property Let ATIVO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVO <> vData)
   mvarATIVO = vData
End Property
Public Property Get ATIVO() As Double
   ATIVO = mvarATIVO
End Property
Public Property Let VLTOTAL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLTOTAL <> vData)
   mvarVLTOTAL = vData
End Property
Public Property Get VLTOTAL() As Double
   VLTOTAL = mvarVLTOTAL
End Property
Public Property Let VLDESC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLDESC <> vData)
   mvarVLDESC = vData
End Property
Public Property Get VLDESC() As Double
   VLDESC = mvarVLDESC
End Property
Public Property Let VALOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVALOR <> vData)
   mvarVALOR = vData
End Property
Public Property Get VALOR() As Double
   VALOR = mvarVALOR
End Property
Public Property Let DTINI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTINI <> vData)
   mvarDTINI = vData
End Property
Public Property Get DTINI() As String
   DTINI = mvarDTINI
End Property
Public Property Let DTFIM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTFIM <> vData)
   mvarDTFIM = vData
End Property
Public Property Get DTFIM() As String
   DTFIM = mvarDTFIM
End Property
Public Property Let DTINIV(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTINIV <> vData)
   mvarDTINIV = vData
End Property
Public Property Get DTINIV() As String
   DTINIV = mvarDTINIV
End Property
Public Property Let DTFIMV(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTFIMV <> vData)
   mvarDTFIMV = vData
End Property
Public Property Get DTFIMV() As String
   DTFIMV = mvarDTFIMV
End Property
Public Property Let NPARCELA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNPARCELA <> vData)
   mvarNPARCELA = vData
End Property
Public Property Get NPARCELA() As Double
   NPARCELA = mvarNPARCELA
End Property
Public Property Let FLGSERV(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGSERV <> vData)
   mvarFLGSERV = vData
End Property
Public Property Get FLGSERV() As Double
   FLGSERV = mvarFLGSERV
End Property
Public Property Let FLGTRAT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGTRAT <> vData)
   mvarFLGTRAT = vData
End Property
Public Property Get FLGTRAT() As Double
   FLGTRAT = mvarFLGTRAT
End Property
Public Property Let FLGAREA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGAREA <> vData)
   mvarFLGAREA = vData
End Property
Public Property Get FLGAREA() As Double
   FLGAREA = mvarFLGAREA
End Property
Public Property Let SERVIN(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSERVIN <> vData)
   mvarSERVIN = vData
End Property
Public Property Get SERVIN() As String
   SERVIN = mvarSERVIN
End Property
Public Property Let TRATIN(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTRATIN <> vData)
   mvarTRATIN = vData
End Property
Public Property Get TRATIN() As String
   TRATIN = mvarTRATIN
End Property
Public Property Let AREAIN(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarAREAIN <> vData)
   mvarAREAIN = vData
End Property
Public Property Get AREAIN() As String
   AREAIN = mvarAREAIN
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update CPROMOCAO Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDPROMO  = " & SqlNum(mvarIDPROMO) & vbNewLine
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
 
   Sql = "Insert Into CPROMOCAO (IDLOJA, IDPROMO, DSCPROMO, ATIVO, VLTOTAL" & vbNewLine
   Sql = Sql & ", VLDESC, VALOR, DTINI, DTFIM, DTINIV" & vbNewLine
   Sql = Sql & ", DTFIMV, NPARCELA, FLGSERV, FLGTRAT, FLGAREA" & vbNewLine
   Sql = Sql & ", SERVIN, TRATIN, AREAIN, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDPROMO),0)+1 From CPROMOCAO Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDPROMO) & vbNewLine
   End If
   Sql = Sql & ", " & SqlStr(mvarDSCPROMO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarATIVO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLTOTAL) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLDESC) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVALOR) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTINI, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTINIV, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTFIMV, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarNPARCELA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGSERV) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGTRAT) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGAREA) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarSERVIN) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarTRATIN) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarAREAIN) & vbNewLine
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
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From CPROMOCAO" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDPROMO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDPROMO = " & SqlNum(mvarIDPROMO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDPROMO) Then Sql = Sql & " IDPROMO = " & SqlNum(CStr(Ch_IDPROMO)) & " AND " & vbNewLine
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
   sTag = sTag & "|IDPROMO = " & mvarIDPROMO
   sTag = sTag & "|DSCPROMO = " & mvarDSCPROMO
   sTag = sTag & "|ATIVO = " & mvarATIVO
   sTag = sTag & "|VLTOTAL = " & mvarVLTOTAL
   sTag = sTag & "|VLDESC = " & mvarVLDESC
   sTag = sTag & "|VALOR = " & mvarVALOR
   sTag = sTag & "|DTINI = " & mvarDTINI
   sTag = sTag & "|DTFIM = " & mvarDTFIM
   sTag = sTag & "|DTINIV = " & mvarDTINIV
   sTag = sTag & "|DTFIMV = " & mvarDTFIMV
   sTag = sTag & "|NPARCELA = " & mvarNPARCELA
   sTag = sTag & "|FLGSERV = " & mvarFLGSERV
   sTag = sTag & "|FLGTRAT = " & mvarFLGTRAT
   sTag = sTag & "|FLGAREA = " & mvarFLGAREA
   sTag = sTag & "|SERVIN = " & mvarSERVIN
   sTag = sTag & "|TRATIN = " & mvarTRATIN
   sTag = sTag & "|AREAIN = " & mvarAREAIN
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
 
   Sql = "Update CPROMOCAO Set " & vbNewLine
   Sql = Sql & " IDLOJA     = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDPROMO    = " & SqlNum(mvarIDPROMO) & vbNewLine
   Sql = Sql & " , DSCPROMO   = " & SqlStr(mvarDSCPROMO) & vbNewLine
   Sql = Sql & " , ATIVO      = " & SqlNum(mvarATIVO) & vbNewLine
   Sql = Sql & " , VLTOTAL    = " & SqlNum(mvarVLTOTAL) & vbNewLine
   Sql = Sql & " , VLDESC     = " & SqlNum(mvarVLDESC) & vbNewLine
   Sql = Sql & " , VALOR      = " & SqlNum(mvarVALOR) & vbNewLine
   Sql = Sql & " , DTINI      = " & SqlDate(mvarDTINI, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , DTFIM      = " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , DTINIV     = " & SqlDate(mvarDTINIV, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , DTFIMV     = " & SqlDate(mvarDTFIMV, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , NPARCELA   = " & SqlNum(mvarNPARCELA) & vbNewLine
   Sql = Sql & " , FLGSERV    = " & SqlNum(mvarFLGSERV) & vbNewLine
   Sql = Sql & " , FLGTRAT    = " & SqlNum(mvarFLGTRAT) & vbNewLine
   Sql = Sql & " , FLGAREA    = " & SqlNum(mvarFLGAREA) & vbNewLine
   Sql = Sql & " , SERVIN     = " & SqlStr(mvarSERVIN) & vbNewLine
   Sql = Sql & " , TRATIN     = " & SqlStr(mvarTRATIN) & vbNewLine
   Sql = Sql & " , AREAIN     = " & SqlStr(mvarAREAIN) & vbNewLine
   If pSinc Then
      Sql = Sql & ", ALTERSTAMP = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP = 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", TIMESTAMP  = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP  = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CPROMOCAO'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDPROMO  = " & SqlNum(mvarIDPROMO) & vbNewLine

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
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDPROMO, DSCPROMO, ATIVO, VLTOTAL" & vbNewLine
   Sql = Sql & ", VLDESC, VALOR, DTINI, DTFIM, DTINIV" & vbNewLine
   Sql = Sql & ", DTFIMV, NPARCELA, FLGSERV, FLGTRAT, FLGAREA" & vbNewLine
   Sql = Sql & ", SERVIN, TRATIN, AREAIN, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From CPROMOCAO" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDPROMO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDPROMO = " & SqlNum(mvarIDPROMO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDPROMO) Then Sql = Sql & " IDPROMO = " & SqlNum(CStr(Ch_IDPROMO)) & " AND " & vbNewLine
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
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDPROMO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPROMO = xVal(!IDPROMO & "")
      mvarDSCPROMO = !DSCPROMO & ""
      mvarATIVO = xVal(!ATIVO & "")
      mvarVLTOTAL = xVal(!VLTOTAL & "")
      mvarVLDESC = xVal(!VLDESC & "")
      mvarVALOR = xVal(!VALOR & "")
      mvarDTINI = xDate(!DTINI & "", True)
      mvarDTFIM = xDate(!DTFIM & "", True)
      mvarDTINIV = xDate(!DTINIV & "", True)
      mvarDTFIMV = xDate(!DTFIMV & "", True)
      mvarNPARCELA = xVal(!NPARCELA & "")
      mvarFLGSERV = xVal(!FLGSERV & "")
      mvarFLGTRAT = xVal(!FLGTRAT & "")
      mvarFLGAREA = xVal(!FLGAREA & "")
      mvarSERVIN = !SERVIN & ""
      mvarTRATIN = !TRATIN & ""
      mvarAREAIN = !AREAIN & ""
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDPROMO = 0
   mvarDSCPROMO = ""
   mvarATIVO = 0
   mvarVLTOTAL = 0
   mvarVLDESC = 0
   mvarVALOR = 0
   mvarDTINI = ""
   mvarDTFIM = ""
   mvarDTINIV = ""
   mvarDTFIMV = ""
   mvarNPARCELA = 0
   mvarFLGSERV = 0
   mvarFLGTRAT = 0
   mvarFLGAREA = 0
   mvarSERVIN = ""
   mvarTRATIN = ""
   mvarAREAIN = ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDPROMO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDPROMO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CPROMOCAO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDPROMO = " & SqlNum(Ch_IDPROMO)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDPROMO = " & SqlNum(mvarIDPROMO)
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
