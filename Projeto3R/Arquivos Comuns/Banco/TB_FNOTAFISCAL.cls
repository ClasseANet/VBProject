VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_FNOTAFISCAL"
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
Private mvarSERIE        As String
Private mvarNUMNOTA      As Double
Private mvarIDVENDA      As Double
Private mvarIDRECIBO     As Double
Private mvarDTEMISSAO    As String
Private mvarVLTOTAL      As Double
Private mvarFLGENVIADA   As Double
Private mvarFLGCANCELADA As Double
Private mvarALTERSTAMP   As Double
Private mvarTIMESTAMP    As String
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
Public Property Let NUMNOTA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMNOTA <> vData)
   mvarNUMNOTA = vData
End Property
Public Property Get NUMNOTA() As Double
   NUMNOTA = mvarNUMNOTA
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDRECIBO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDRECIBO <> vData)
   mvarIDRECIBO = vData
End Property
Public Property Get IDRECIBO() As Double
   IDRECIBO = mvarIDRECIBO
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
Public Property Let FLGENVIADA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGENVIADA <> vData)
   mvarFLGENVIADA = vData
End Property
Public Property Get FLGENVIADA() As Double
   FLGENVIADA = mvarFLGENVIADA
End Property
Public Property Let FLGCANCELADA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGCANCELADA <> vData)
   mvarFLGCANCELADA = vData
End Property
Public Property Get FLGCANCELADA() As Double
   FLGCANCELADA = mvarFLGCANCELADA
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update FNOTAFISCAL Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And SERIE      = " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & " And NUMNOTA    = " & SqlNum(mvarNUMNOTA) & VbNewLine 
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into FNOTAFISCAL (IDLOJA, SERIE, NUMNOTA, IDVENDA" & VbNewLine 
   Sql = Sql & ", IDRECIBO, DTEMISSAO, VLTOTAL, FLGENVIADA, FLGCANCELADA" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarNUMNOTA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDRECIBO) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLTOTAL) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGENVIADA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFLGCANCELADA) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_NUMNOTA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From FNOTAFISCAL" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_SERIE)  And IsMissing(Ch_NUMNOTA)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " SERIE = " & SqlStr(mvarSERIE) & " AND " & VbNewLine 
      Sql = Sql & " NUMNOTA = " & SqlNum(mvarNUMNOTA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_SERIE) Then Sql = Sql & " SERIE = " & SqlStr(Cstr(Ch_SERIE)) & " AND " & VbNewLine 
      If Not isMissing(Ch_NUMNOTA) Then Sql = Sql & " NUMNOTA = " & SqlNum(Cstr(Ch_NUMNOTA)) & " AND " & VbNewLine 
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
   sTag = sTag & "|SERIE = " & mvarSERIE
   sTag = sTag & "|NUMNOTA = " & mvarNUMNOTA
   sTag = sTag & "|IDVENDA = " & mvarIDVENDA
   sTag = sTag & "|IDRECIBO = " & mvarIDRECIBO
   sTag = sTag & "|DTEMISSAO = " & mvarDTEMISSAO
   sTag = sTag & "|VLTOTAL = " & mvarVLTOTAL
   sTag = sTag & "|FLGENVIADA = " & mvarFLGENVIADA
   sTag = sTag & "|FLGCANCELADA = " & mvarFLGCANCELADA
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
 
   Sql = "Update FNOTAFISCAL Set " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , SERIE        = " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & " , NUMNOTA      = " & SqlNum(mvarNUMNOTA) & VbNewLine 
   Sql = Sql & " , IDVENDA      = " & IIf(mvarIDVENDA = 0, "Null", SqlNum(mvarIDVENDA)) & VbNewLine 
   Sql = Sql & " , IDRECIBO     = " & IIf(mvarIDRECIBO = 0, "Null", SqlNum(mvarIDRECIBO)) & VbNewLine 
   Sql = Sql & " , DTEMISSAO    = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VLTOTAL      = " & SqlNum(mvarVLTOTAL) & VbNewLine 
   Sql = Sql & " , FLGENVIADA   = " & SqlNum(mvarFLGENVIADA) & VbNewLine 
   Sql = Sql & " , FLGCANCELADA = " & SqlNum(mvarFLGCANCELADA) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP   = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP   = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP    = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP    = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FNOTAFISCAL'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And SERIE      = " & SqlStr(mvarSERIE) & VbNewLine 
   Sql = Sql & " And NUMNOTA    = " & SqlNum(mvarNUMNOTA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_NUMNOTA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, SERIE, NUMNOTA, IDVENDA" & VbNewLine 
   Sql = Sql & ", IDRECIBO, DTEMISSAO, VLTOTAL, FLGENVIADA, FLGCANCELADA" & VbNewLine 
   Sql = Sql & ", ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From FNOTAFISCAL" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_SERIE)  And IsMissing(Ch_NUMNOTA)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " SERIE = " & SqlStr(mvarSERIE) & " AND " & VbNewLine 
      Sql = Sql & " NUMNOTA = " & SqlNum(mvarNUMNOTA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_SERIE) Then Sql = Sql & " SERIE = " & SqlStr(Cstr(Ch_SERIE)) & " AND " & VbNewLine 
      If Not isMissing(Ch_NUMNOTA) Then Sql = Sql & " NUMNOTA = " & SqlNum(Cstr(Ch_NUMNOTA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_SERIE, Optional Ch_NUMNOTA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_SERIE, Ch_NUMNOTA, Ch_WHERE, Ch_ORDERBY)
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
      mvarSERIE = !SERIE & ""
      mvarNUMNOTA = XVal(!NUMNOTA & "")
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarIDRECIBO = XVal(!IDRECIBO & "")
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarVLTOTAL = XVal(!VLTOTAL & "")
      mvarFLGENVIADA = XVal(!FLGENVIADA & "")
      mvarFLGCANCELADA = XVal(!FLGCANCELADA & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarSERIE = ""
   mvarNUMNOTA = 0
   mvarIDVENDA = 0
   mvarIDRECIBO = 0
   mvarDTEMISSAO = ""
   mvarVLTOTAL = 0
   mvarFLGENVIADA = 0
   mvarFLGCANCELADA = 0
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
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId as Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarSERIE, mvarNUMNOTA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_SERIE As String, Ch_NUMNOTA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FNOTAFISCAL Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , SERIE = " & SqlStr(Ch_SERIE)
   Sql = Sql & " , NUMNOTA = " & SqlNum(Ch_NUMNOTA)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and SERIE = " & SqlStr(mvarSERIE)
   Sql = Sql & " and NUMNOTA = " & SqlNum(mvarNUMNOTA)
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
