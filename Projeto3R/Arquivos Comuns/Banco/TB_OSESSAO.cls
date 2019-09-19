VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_OSESSAO"
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
Private mvarIDATENDIMENTO  As Double
Private mvarIDSESSAO       As Double
Private mvarIDTPSERVICO    As Double
Private mvarIDTPTRATAMENTO As Double
Private mvarIDMANIPULO     As Double
Private mvarIDDIRECAO      As Double
Private mvarIDAREA         As Double
Private mvarIDVENDA        As Double
Private mvarIDITEM         As Double
Private mvarPOTENCIA       As Double
Private mvarDISPAROS       As Double
Private mvarFOTOTIPO       As Double
Private mvarNUMSESSAO      As Double
Private mvarTEMPOSESSAO    As Double
Private mvarALTERSTAMP     As Double
Private mvarTIMESTAMP      As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDATENDIMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDATENDIMENTO <> vData)
   mvarIDATENDIMENTO = vData
End Property
Public Property Get IDATENDIMENTO() As Double
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDSESSAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSESSAO <> vData)
   mvarIDSESSAO = vData
End Property
Public Property Get IDSESSAO() As Double
   IDSESSAO = mvarIDSESSAO
End Property
Public Property Let IDTPSERVICO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPSERVICO <> vData)
   mvarIDTPSERVICO = vData
End Property
Public Property Get IDTPSERVICO() As Double
   IDTPSERVICO = mvarIDTPSERVICO
End Property
Public Property Let IDTPTRATAMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPTRATAMENTO <> vData)
   mvarIDTPTRATAMENTO = vData
End Property
Public Property Get IDTPTRATAMENTO() As Double
   IDTPTRATAMENTO = mvarIDTPTRATAMENTO
End Property
Public Property Let IDMANIPULO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMANIPULO <> vData)
   mvarIDMANIPULO = vData
End Property
Public Property Get IDMANIPULO() As Double
   IDMANIPULO = mvarIDMANIPULO
End Property
Public Property Let IDDIRECAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDIRECAO <> vData)
   mvarIDDIRECAO = vData
End Property
Public Property Get IDDIRECAO() As Double
   IDDIRECAO = mvarIDDIRECAO
End Property
Public Property Let IDAREA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAREA <> vData)
   mvarIDAREA = vData
End Property
Public Property Get IDAREA() As Double
   IDAREA = mvarIDAREA
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDITEM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDITEM <> vData)
   mvarIDITEM = vData
End Property
Public Property Get IDITEM() As Double
   IDITEM = mvarIDITEM
End Property
Public Property Let POTENCIA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarPOTENCIA <> vData)
   mvarPOTENCIA = vData
End Property
Public Property Get POTENCIA() As Double
   POTENCIA = mvarPOTENCIA
End Property
Public Property Let DISPAROS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarDISPAROS <> vData)
   mvarDISPAROS = vData
End Property
Public Property Get DISPAROS() As Double
   DISPAROS = mvarDISPAROS
End Property
Public Property Let FOTOTIPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFOTOTIPO <> vData)
   mvarFOTOTIPO = vData
End Property
Public Property Get FOTOTIPO() As Double
   FOTOTIPO = mvarFOTOTIPO
End Property
Public Property Let NUMSESSAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMSESSAO <> vData)
   mvarNUMSESSAO = vData
End Property
Public Property Get NUMSESSAO() As Double
   NUMSESSAO = mvarNUMSESSAO
End Property
Public Property Let TEMPOSESSAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTEMPOSESSAO <> vData)
   mvarTEMPOSESSAO = vData
End Property
Public Property Get TEMPOSESSAO() As Double
   TEMPOSESSAO = mvarTEMPOSESSAO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OSESSAO Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & vbNewLine
   Sql = Sql & " And IDSESSAO     = " & SqlNum(mvarIDSESSAO) & vbNewLine
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
 
   Sql = "Insert Into OSESSAO (IDLOJA, IDATENDIMENTO, IDSESSAO, IDTPSERVICO" & vbNewLine
   Sql = Sql & ", IDTPTRATAMENTO, IDMANIPULO, IDDIRECAO, IDAREA, IDVENDA" & vbNewLine
   Sql = Sql & ", IDITEM, POTENCIA, DISPAROS, FOTOTIPO, NUMSESSAO" & vbNewLine
   Sql = Sql & ", TEMPOSESSAO, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDATENDIMENTO) & vbNewLine
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDSESSAO),0)+1 From OSESSAO Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDSESSAO) & vbNewLine
   End If
   Sql = Sql & ", " & IIf(mvarIDTPSERVICO = 0, "Null", SqlNum(mvarIDTPSERVICO)) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDTPTRATAMENTO = 0, "Null", SqlNum(mvarIDTPTRATAMENTO)) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDMANIPULO) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDDIRECAO = 0, "Null", SqlNum(mvarIDDIRECAO)) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDAREA = 0, "Null", SqlNum(mvarIDAREA)) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDITEM) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarPOTENCIA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarDISPAROS) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFOTOTIPO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarNUMSESSAO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarTEMPOSESSAO) & vbNewLine
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
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDATENDIMENTO, Optional Ch_IDSESSAO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OSESSAO" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDATENDIMENTO) And IsMissing(Ch_IDSESSAO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & " AND " & vbNewLine
      Sql = Sql & " IDSESSAO = " & SqlNum(mvarIDSESSAO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDATENDIMENTO) Then Sql = Sql & " IDATENDIMENTO = " & SqlNum(CStr(Ch_IDATENDIMENTO)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDSESSAO) Then Sql = Sql & " IDSESSAO = " & SqlNum(CStr(Ch_IDSESSAO)) & " AND " & vbNewLine
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
   sTag = sTag & "|IDATENDIMENTO = " & mvarIDATENDIMENTO
   sTag = sTag & "|IDSESSAO = " & mvarIDSESSAO
   sTag = sTag & "|IDTPSERVICO = " & mvarIDTPSERVICO
   sTag = sTag & "|IDTPTRATAMENTO = " & mvarIDTPTRATAMENTO
   sTag = sTag & "|IDMANIPULO = " & mvarIDMANIPULO
   sTag = sTag & "|IDDIRECAO = " & mvarIDDIRECAO
   sTag = sTag & "|IDAREA = " & mvarIDAREA
   sTag = sTag & "|IDVENDA = " & mvarIDVENDA
   sTag = sTag & "|IDITEM = " & mvarIDITEM
   sTag = sTag & "|POTENCIA = " & mvarPOTENCIA
   sTag = sTag & "|DISPAROS = " & mvarDISPAROS
   sTag = sTag & "|FOTOTIPO = " & mvarFOTOTIPO
   sTag = sTag & "|NUMSESSAO = " & mvarNUMSESSAO
   sTag = sTag & "|TEMPOSESSAO = " & mvarTEMPOSESSAO
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
 
   Sql = "Update OSESSAO Set " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDATENDIMENTO  = " & SqlNum(mvarIDATENDIMENTO) & vbNewLine
   Sql = Sql & " , IDSESSAO       = " & SqlNum(mvarIDSESSAO) & vbNewLine
   Sql = Sql & " , IDTPSERVICO    = " & IIf(mvarIDTPSERVICO = 0, "Null", SqlNum(mvarIDTPSERVICO)) & vbNewLine
   Sql = Sql & " , IDTPTRATAMENTO = " & IIf(mvarIDTPTRATAMENTO = 0, "Null", SqlNum(mvarIDTPTRATAMENTO)) & vbNewLine
   Sql = Sql & " , IDMANIPULO     = " & SqlNum(mvarIDMANIPULO) & vbNewLine
   Sql = Sql & " , IDDIRECAO      = " & IIf(mvarIDDIRECAO = 0, "Null", SqlNum(mvarIDDIRECAO)) & vbNewLine
   Sql = Sql & " , IDAREA         = " & IIf(mvarIDAREA = 0, "Null", SqlNum(mvarIDAREA)) & vbNewLine
   Sql = Sql & " , IDVENDA        = " & SqlNum(mvarIDVENDA) & vbNewLine
   Sql = Sql & " , IDITEM         = " & SqlNum(mvarIDITEM) & vbNewLine
   Sql = Sql & " , POTENCIA       = " & SqlNum(mvarPOTENCIA) & vbNewLine
   Sql = Sql & " , DISPAROS       = " & SqlNum(mvarDISPAROS) & vbNewLine
   Sql = Sql & " , FOTOTIPO       = " & SqlNum(mvarFOTOTIPO) & vbNewLine
   Sql = Sql & " , NUMSESSAO      = " & SqlNum(mvarNUMSESSAO) & vbNewLine
   Sql = Sql & " , TEMPOSESSAO    = " & SqlNum(mvarTEMPOSESSAO) & vbNewLine
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
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OSESSAO'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & vbNewLine
   Sql = Sql & " And IDSESSAO     = " & SqlNum(mvarIDSESSAO) & vbNewLine

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
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDATENDIMENTO, Optional Ch_IDSESSAO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDATENDIMENTO, IDSESSAO, IDTPSERVICO" & vbNewLine
   Sql = Sql & ", IDTPTRATAMENTO, IDMANIPULO, IDDIRECAO, IDAREA, IDVENDA" & vbNewLine
   Sql = Sql & ", IDITEM, POTENCIA, DISPAROS, FOTOTIPO, NUMSESSAO" & vbNewLine
   Sql = Sql & ", TEMPOSESSAO, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From OSESSAO" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDATENDIMENTO) And IsMissing(Ch_IDSESSAO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & " AND " & vbNewLine
      Sql = Sql & " IDSESSAO = " & SqlNum(mvarIDSESSAO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDATENDIMENTO) Then Sql = Sql & " IDATENDIMENTO = " & SqlNum(CStr(Ch_IDATENDIMENTO)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDSESSAO) Then Sql = Sql & " IDSESSAO = " & SqlNum(CStr(Ch_IDSESSAO)) & " AND " & vbNewLine
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
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDATENDIMENTO, Optional Ch_IDSESSAO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDATENDIMENTO, Ch_IDSESSAO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDATENDIMENTO = xVal(!IDATENDIMENTO & "")
      mvarIDSESSAO = xVal(!IDSESSAO & "")
      mvarIDTPSERVICO = xVal(!IDTPSERVICO & "")
      mvarIDTPTRATAMENTO = xVal(!IDTPTRATAMENTO & "")
      mvarIDMANIPULO = xVal(!IDMANIPULO & "")
      mvarIDDIRECAO = xVal(!IDDIRECAO & "")
      mvarIDAREA = xVal(!IDAREA & "")
      mvarIDVENDA = xVal(!IDVENDA & "")
      mvarIDITEM = xVal(!IDITEM & "")
      mvarPOTENCIA = xVal(!POTENCIA & "")
      mvarDISPAROS = xVal(!DISPAROS & "")
      mvarFOTOTIPO = xVal(!FOTOTIPO & "")
      mvarNUMSESSAO = xVal(!NUMSESSAO & "")
      mvarTEMPOSESSAO = xVal(!TEMPOSESSAO & "")
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDATENDIMENTO = 0
   mvarIDSESSAO = 0
   mvarIDTPSERVICO = 0
   mvarIDTPTRATAMENTO = 0
   mvarIDMANIPULO = 0
   mvarIDDIRECAO = 0
   mvarIDAREA = 0
   mvarIDVENDA = 0
   mvarIDITEM = 0
   mvarPOTENCIA = 0
   mvarDISPAROS = 0
   mvarFOTOTIPO = 0
   mvarNUMSESSAO = 0
   mvarTEMPOSESSAO = 0
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
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDATENDIMENTO, mvarIDSESSAO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDATENDIMENTO As String, Ch_IDSESSAO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OSESSAO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDATENDIMENTO = " & SqlNum(Ch_IDATENDIMENTO)
   Sql = Sql & " , IDSESSAO = " & SqlNum(Ch_IDSESSAO)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO)
   Sql = Sql & " and IDSESSAO = " & SqlNum(mvarIDSESSAO)
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
