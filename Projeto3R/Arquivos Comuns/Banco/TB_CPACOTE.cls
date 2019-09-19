VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_CPACOTE"
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
Private mvarIDVENDA    As Double
Private mvarIDPACOTE   As Double
Private mvarDSCPACOTE  As String
Private mvarIDCLIENTE  As Double
Private mvarIDPROMO    As Double
Private mvarDTEMISSAO  As String
Private mvarVALOR      As Double
Private mvarVLDESC     As Double
Private mvarVLTOTAL    As Double
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDPACOTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPACOTE <> vData)
   mvarIDPACOTE = vData
End Property
Public Property Get IDPACOTE() As Double
   IDPACOTE = mvarIDPACOTE
End Property
Public Property Let DSCPACOTE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCPACOTE <> vData)
   mvarDSCPACOTE = vData
End Property
Public Property Get DSCPACOTE() As String
   DSCPACOTE = mvarDSCPACOTE
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDPROMO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROMO <> vData)
   mvarIDPROMO = vData
End Property
Public Property Get IDPROMO() As Double
   IDPROMO = mvarIDPROMO
End Property
Public Property Let DTEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEMISSAO <> vData)
   mvarDTEMISSAO = vData
End Property
Public Property Get DTEMISSAO() As String
   DTEMISSAO = mvarDTEMISSAO
End Property
Public Property Let VALOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVALOR <> vData)
   mvarVALOR = vData
End Property
Public Property Get VALOR() As Double
   VALOR = mvarVALOR
End Property
Public Property Let VLDESC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLDESC <> vData)
   mvarVLDESC = vData
End Property
Public Property Get VLDESC() As Double
   VLDESC = mvarVLDESC
End Property
Public Property Let VLTOTAL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLTOTAL <> vData)
   mvarVLTOTAL = vData
End Property
Public Property Get VLTOTAL() As Double
   VLTOTAL = mvarVLTOTAL
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update CPACOTE Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDVENDA  = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " And IDPACOTE = " & SqlNum(mvarIDPACOTE) & VbNewLine 
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
 
   Sql = "Insert Into CPACOTE (IDLOJA, IDVENDA, IDPACOTE, DSCPACOTE" & VbNewLine 
   Sql = Sql & ", IDCLIENTE, IDPROMO, DTEMISSAO, VALOR, VLDESC" & VbNewLine 
   Sql = Sql & ", VLTOTAL, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDPACOTE),0)+1 From CPACOTE Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDVENDA = " & SqlNum(mvarIDVENDA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDPACOTE) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarDSCPACOTE) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDCLIENTE = 0, "Null", SqlNum(mvarIDCLIENTE)) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDPROMO = 0, "Null", SqlNum(mvarIDPROMO)) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVALOR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLTOTAL) & VbNewLine 
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

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDPACOTE, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From CPACOTE" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDVENDA)  And IsMissing(Ch_IDPACOTE)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDVENDA = " & SqlNum(mvarIDVENDA) & " AND " & VbNewLine 
      Sql = Sql & " IDPACOTE = " & SqlNum(mvarIDPACOTE) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(Cstr(Ch_IDVENDA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPACOTE) Then Sql = Sql & " IDPACOTE = " & SqlNum(Cstr(Ch_IDPACOTE)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDVENDA = " & mvarIDVENDA
   sTag = sTag & "|IDPACOTE = " & mvarIDPACOTE
   sTag = sTag & "|DSCPACOTE = " & mvarDSCPACOTE
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|IDPROMO = " & mvarIDPROMO
   sTag = sTag & "|DTEMISSAO = " & mvarDTEMISSAO
   sTag = sTag & "|VALOR = " & mvarVALOR
   sTag = sTag & "|VLDESC = " & mvarVLDESC
   sTag = sTag & "|VLTOTAL = " & mvarVLTOTAL
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
 
   Sql = "Update CPACOTE Set " & VbNewLine 
   Sql = Sql & " IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDVENDA    = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " , IDPACOTE   = " & SqlNum(mvarIDPACOTE) & VbNewLine 
   Sql = Sql & " , DSCPACOTE  = " & SqlStr(mvarDSCPACOTE) & VbNewLine 
   Sql = Sql & " , IDCLIENTE  = " & IIf(mvarIDCLIENTE = 0, "Null", SqlNum(mvarIDCLIENTE)) & VbNewLine 
   Sql = Sql & " , IDPROMO    = " & IIf(mvarIDPROMO = 0, "Null", SqlNum(mvarIDPROMO)) & VbNewLine 
   Sql = Sql & " , DTEMISSAO  = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VALOR      = " & SqlNum(mvarVALOR) & VbNewLine 
   Sql = Sql & " , VLDESC     = " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & " , VLTOTAL    = " & SqlNum(mvarVLTOTAL) & VbNewLine 
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
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CPACOTE'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDVENDA  = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " And IDPACOTE = " & SqlNum(mvarIDPACOTE) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDPACOTE, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDVENDA, IDPACOTE, DSCPACOTE" & VbNewLine 
   Sql = Sql & ", IDCLIENTE, IDPROMO, DTEMISSAO, VALOR, VLDESC" & VbNewLine 
   Sql = Sql & ", VLTOTAL, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From CPACOTE" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDVENDA)  And IsMissing(Ch_IDPACOTE)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDVENDA = " & SqlNum(mvarIDVENDA) & " AND " & VbNewLine 
      Sql = Sql & " IDPACOTE = " & SqlNum(mvarIDPACOTE) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(Cstr(Ch_IDVENDA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPACOTE) Then Sql = Sql & " IDPACOTE = " & SqlNum(Cstr(Ch_IDPACOTE)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDPACOTE, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDVENDA, Ch_IDPACOTE, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarIDPACOTE = XVal(!IDPACOTE & "")
      mvarDSCPACOTE = !DSCPACOTE & ""
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
      mvarIDPROMO = XVal(!IDPROMO & "")
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarVALOR = XVal(!VALOR & "")
      mvarVLDESC = XVal(!VLDESC & "")
      mvarVLTOTAL = XVal(!VLTOTAL & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDVENDA = 0
   mvarIDPACOTE = 0
   mvarDSCPACOTE = ""
   mvarIDCLIENTE = 0
   mvarIDPROMO = 0
   mvarDTEMISSAO = ""
   mvarVALOR = 0
   mvarVLDESC = 0
   mvarVLTOTAL = 0
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDVENDA, mvarIDPACOTE), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDVENDA As String, Ch_IDPACOTE As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CPACOTE Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDVENDA = " & SqlNum(Ch_IDVENDA)
   Sql = Sql & " , IDPACOTE = " & SqlNum(Ch_IDPACOTE)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDVENDA = " & SqlNum(mvarIDVENDA)
   Sql = Sql & " and IDPACOTE = " & SqlNum(mvarIDPACOTE)
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
