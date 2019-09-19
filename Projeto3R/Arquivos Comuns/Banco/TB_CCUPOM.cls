VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_CCUPOM"
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
Private mvarIDPROMO        As Double
Private mvarIDCUPOM        As Double
Private mvarCODCUPOM       As String
Private mvarDTEMISSAO      As String
Private mvarDTVENC         As String
Private mvarVLDESC         As Double
Private mvarATIVO          As Double
Private mvarALTERSTAMP     As Double
Private mvarTIMESTAMP      As String
Private mvarIDPACOTE       As Double
Private mvarIDATENDIMENTO  As Double
Private mvarIDSESSAO       As Double
Private mvarIDVENDA        As Double
Private mvarIDCLIENTE      As Double
Private mvarIDTPSERVICO    As Double
Private mvarIDTPTRATAMENTO As Double
Private mvarIDAREA         As Double
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
Public Property Let IDCUPOM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCUPOM <> vData)
   mvarIDCUPOM = vData
End Property
Public Property Get IDCUPOM() As Double
   IDCUPOM = mvarIDCUPOM
End Property
Public Property Let CODCUPOM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODCUPOM <> vData)
   mvarCODCUPOM = vData
End Property
Public Property Get CODCUPOM() As String
   CODCUPOM = mvarCODCUPOM
End Property
Public Property Let DTEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEMISSAO <> vData)
   mvarDTEMISSAO = vData
End Property
Public Property Get DTEMISSAO() As String
   DTEMISSAO = mvarDTEMISSAO
End Property
Public Property Let DTVENC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTVENC <> vData)
   mvarDTVENC = vData
End Property
Public Property Get DTVENC() As String
   DTVENC = mvarDTVENC
End Property
Public Property Let VLDESC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLDESC <> vData)
   mvarVLDESC = vData
End Property
Public Property Get VLDESC() As Double
   VLDESC = mvarVLDESC
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
 
   Sql = "Update CCUPOM Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDPROMO      = " & SqlNum(mvarIDPROMO) & VbNewLine 
   Sql = Sql & " And IDCUPOM      = " & SqlNum(mvarIDCUPOM) & VbNewLine 
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
Public Property Let IDPACOTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPACOTE <> vData)
   mvarIDPACOTE = vData
End Property
Public Property Get IDPACOTE() As Double
   IDPACOTE = mvarIDPACOTE
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
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
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
Public Property Let IDAREA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAREA <> vData)
   mvarIDAREA = vData
End Property
Public Property Get IDAREA() As Double
   IDAREA = mvarIDAREA
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
 
   Sql = "Insert Into CCUPOM (IDLOJA, IDPROMO, IDCUPOM, CODCUPOM" & VbNewLine 
   Sql = Sql & ", DTEMISSAO, DTVENC, VLDESC, ATIVO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP, IDPACOTE, IDATENDIMENTO, IDSESSAO, IDVENDA" & VbNewLine 
   Sql = Sql & ", IDCLIENTE, IDTPSERVICO, IDTPTRATAMENTO, IDAREA) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPROMO) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDCUPOM),0)+1 From CCUPOM Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDPROMO = " & SqlNum(mvarIDPROMO) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDCUPOM) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarCODCUPOM) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTVENC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarATIVO) & VbNewLine 
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
   Sql = Sql & ", " & SqlNum(mvarIDPACOTE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDATENDIMENTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDSESSAO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTPSERVICO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDAREA) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDCUPOM, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From CCUPOM" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDPROMO)  And IsMissing(Ch_IDCUPOM)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDPROMO = " & SqlNum(mvarIDPROMO) & " AND " & VbNewLine 
      Sql = Sql & " IDCUPOM = " & SqlNum(mvarIDCUPOM) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPROMO) Then Sql = Sql & " IDPROMO = " & SqlNum(Cstr(Ch_IDPROMO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCUPOM) Then Sql = Sql & " IDCUPOM = " & SqlNum(Cstr(Ch_IDCUPOM)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDPROMO = " & mvarIDPROMO
   sTag = sTag & "|IDCUPOM = " & mvarIDCUPOM
   sTag = sTag & "|CODCUPOM = " & mvarCODCUPOM
   sTag = sTag & "|DTEMISSAO = " & mvarDTEMISSAO
   sTag = sTag & "|DTVENC = " & mvarDTVENC
   sTag = sTag & "|VLDESC = " & mvarVLDESC
   sTag = sTag & "|ATIVO = " & mvarATIVO
   sTag = sTag & "|IDPACOTE = " & mvarIDPACOTE
   sTag = sTag & "|IDATENDIMENTO = " & mvarIDATENDIMENTO
   sTag = sTag & "|IDSESSAO = " & mvarIDSESSAO
   sTag = sTag & "|IDVENDA = " & mvarIDVENDA
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|IDTPSERVICO = " & mvarIDTPSERVICO
   sTag = sTag & "|IDTPTRATAMENTO = " & mvarIDTPTRATAMENTO
   sTag = sTag & "|IDAREA = " & mvarIDAREA
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
 
   Sql = "Update CCUPOM Set " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDPROMO        = " & SqlNum(mvarIDPROMO) & VbNewLine 
   Sql = Sql & " , IDCUPOM        = " & SqlNum(mvarIDCUPOM) & VbNewLine 
   Sql = Sql & " , CODCUPOM       = " & SqlStr(mvarCODCUPOM) & VbNewLine 
   Sql = Sql & " , DTEMISSAO      = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTVENC         = " & SqlDate(mvarDTVENC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VLDESC         = " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & " , ATIVO          = " & SqlNum(mvarATIVO) & VbNewLine 
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
   Sql = Sql & " , IDPACOTE       = " & SqlNum(mvarIDPACOTE) & VbNewLine 
   Sql = Sql & " , IDATENDIMENTO  = " & SqlNum(mvarIDATENDIMENTO) & VbNewLine 
   Sql = Sql & " , IDSESSAO       = " & SqlNum(mvarIDSESSAO) & VbNewLine 
   Sql = Sql & " , IDVENDA        = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " , IDCLIENTE      = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & " , IDTPSERVICO    = " & SqlNum(mvarIDTPSERVICO) & VbNewLine 
   Sql = Sql & " , IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   Sql = Sql & " , IDAREA         = " & SqlNum(mvarIDAREA) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CCUPOM'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDPROMO      = " & SqlNum(mvarIDPROMO) & VbNewLine 
   Sql = Sql & " And IDCUPOM      = " & SqlNum(mvarIDCUPOM) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDCUPOM, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDPROMO, IDCUPOM, CODCUPOM" & VbNewLine 
   Sql = Sql & ", DTEMISSAO, DTVENC, VLDESC, ATIVO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP, IDPACOTE, IDATENDIMENTO, IDSESSAO, IDVENDA" & VbNewLine 
   Sql = Sql & ", IDCLIENTE, IDTPSERVICO, IDTPTRATAMENTO, IDAREA" & VbNewLine 
   Sql = Sql &" From CCUPOM" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDPROMO)  And IsMissing(Ch_IDCUPOM)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDPROMO = " & SqlNum(mvarIDPROMO) & " AND " & VbNewLine 
      Sql = Sql & " IDCUPOM = " & SqlNum(mvarIDCUPOM) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPROMO) Then Sql = Sql & " IDPROMO = " & SqlNum(Cstr(Ch_IDPROMO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCUPOM) Then Sql = Sql & " IDCUPOM = " & SqlNum(Cstr(Ch_IDCUPOM)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDCUPOM, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDPROMO, Ch_IDCUPOM, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPROMO = XVal(!IDPROMO & "")
      mvarIDCUPOM = XVal(!IDCUPOM & "")
      mvarCODCUPOM = !CODCUPOM & ""
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarDTVENC = xDate(!DTVENC & "", True)
      mvarVLDESC = XVal(!VLDESC & "")
      mvarATIVO = XVal(!ATIVO & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarIDPACOTE = XVal(!IDPACOTE & "")
      mvarIDATENDIMENTO = XVal(!IDATENDIMENTO & "")
      mvarIDSESSAO = XVal(!IDSESSAO & "")
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
      mvarIDTPSERVICO = XVal(!IDTPSERVICO & "")
      mvarIDTPTRATAMENTO = XVal(!IDTPTRATAMENTO & "")
      mvarIDAREA = XVal(!IDAREA & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDPROMO = 0
   mvarIDCUPOM = 0
   mvarCODCUPOM = ""
   mvarDTEMISSAO = ""
   mvarDTVENC = ""
   mvarVLDESC = 0
   mvarATIVO = 0
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""
   mvarIDPACOTE = 0
   mvarIDATENDIMENTO = 0
   mvarIDSESSAO = 0
   mvarIDVENDA = 0
   mvarIDCLIENTE = 0
   mvarIDTPSERVICO = 0
   mvarIDTPTRATAMENTO = 0
   mvarIDAREA = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDPROMO, mvarIDCUPOM), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDPROMO As String, Ch_IDCUPOM As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CCUPOM Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDPROMO = " & SqlNum(Ch_IDPROMO)
   Sql = Sql & " , IDCUPOM = " & SqlNum(Ch_IDCUPOM)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDPROMO = " & SqlNum(mvarIDPROMO)
   Sql = Sql & " and IDCUPOM = " & SqlNum(mvarIDCUPOM)
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
