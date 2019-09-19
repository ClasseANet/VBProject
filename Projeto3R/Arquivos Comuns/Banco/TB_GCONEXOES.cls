VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GCONEXOES"
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

Private mvarIDCONEXAO   As Double
Private mvarNMSISTEMA   As String
Private mvarTITULO      As String
Private mvarTIPOSGBD    As Double
Private mvarTIPOCONEXAO As String
Private mvarNMSERVIDOR  As String
Private mvarNMBANCO     As String
Private mvarUSUARIO     As String
Private mvarSENHA       As String
Private mvarFLGPADRAO   As String
Public Property Get IDCONEXAO() As Double
   IDCONEXAO = mvarIDCONEXAO
End Property
Public Property Let NMSISTEMA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMSISTEMA <> vData)
   mvarNMSISTEMA = vData
End Property
Public Property Get NMSISTEMA() As String
   NMSISTEMA = mvarNMSISTEMA
End Property
Public Property Let TITULO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTITULO <> vData)
   mvarTITULO = vData
End Property
Public Property Get TITULO() As String
   TITULO = mvarTITULO
End Property
Public Property Let TIPOSGBD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPOSGBD <> vData)
   mvarTIPOSGBD = vData
End Property
Public Property Get TIPOSGBD() As Double
   TIPOSGBD = mvarTIPOSGBD
End Property
Public Property Let TIPOCONEXAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPOCONEXAO <> vData)
   mvarTIPOCONEXAO = vData
End Property
Public Property Get TIPOCONEXAO() As String
   TIPOCONEXAO = mvarTIPOCONEXAO
End Property
Public Property Let NMSERVIDOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMSERVIDOR <> vData)
   mvarNMSERVIDOR = vData
End Property
Public Property Get NMSERVIDOR() As String
   NMSERVIDOR = mvarNMSERVIDOR
End Property
Public Property Let NMBANCO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMBANCO <> vData)
   mvarNMBANCO = vData
End Property
Public Property Get NMBANCO() As String
   NMBANCO = mvarNMBANCO
End Property
Public Property Let USUARIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUSUARIO <> vData)
   mvarUSUARIO = vData
End Property
Public Property Get USUARIO() As String
   USUARIO = mvarUSUARIO
End Property
Public Property Let SENHA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSENHA <> vData)
   mvarSENHA = vData
End Property
Public Property Get SENHA() As String
   SENHA = mvarSENHA
End Property
Public Property Let FLGPADRAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGPADRAO <> vData)
   mvarFLGPADRAO = vData
End Property
Public Property Get FLGPADRAO() As String
   FLGPADRAO = mvarFLGPADRAO
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
 
   Sql = "Insert Into GCONEXOES (NMSISTEMA, TITULO, TIPOSGBD" & VbNewLine 
   Sql = Sql & ", TIPOCONEXAO, NMSERVIDOR, NMBANCO, USUARIO, SENHA" & VbNewLine 
   Sql = Sql & ", FLGPADRAO) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarNMSISTEMA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTITULO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTIPOSGBD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPOCONEXAO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMSERVIDOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMBANCO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUSUARIO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSENHA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarFLGPADRAO) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDCONEXAO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From GCONEXOES" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDCONEXAO)  Then 
      Sql = Sql & " IDCONEXAO = " & SqlNum(mvarIDCONEXAO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDCONEXAO) Then Sql = Sql & " IDCONEXAO = " & SqlNum(Cstr(Ch_IDCONEXAO)) & " AND " & VbNewLine 
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
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update GCONEXOES Set " & VbNewLine 
   Sql = Sql & " NMSISTEMA   = " & SqlStr(mvarNMSISTEMA) & VbNewLine 
   Sql = Sql & " , TITULO      = " & SqlStr(mvarTITULO) & VbNewLine 
   Sql = Sql & " , TIPOSGBD    = " & SqlNum(mvarTIPOSGBD) & VbNewLine 
   Sql = Sql & " , TIPOCONEXAO = " & SqlStr(mvarTIPOCONEXAO) & VbNewLine 
   Sql = Sql & " , NMSERVIDOR  = " & SqlStr(mvarNMSERVIDOR) & VbNewLine 
   Sql = Sql & " , NMBANCO     = " & SqlStr(mvarNMBANCO) & VbNewLine 
   Sql = Sql & " , USUARIO     = " & SqlStr(mvarUSUARIO) & VbNewLine 
   Sql = Sql & " , SENHA       = " & SqlStr(mvarSENHA) & VbNewLine 
   Sql = Sql & " , FLGPADRAO   = " & SqlStr(mvarFLGPADRAO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GCONEXOES'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDCONEXAO     = " & SqlNum(mvarIDCONEXAO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDCONEXAO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDCONEXAO, NMSISTEMA, TITULO, TIPOSGBD" & VbNewLine 
   Sql = Sql & ", TIPOCONEXAO, NMSERVIDOR, NMBANCO, USUARIO, SENHA" & VbNewLine 
   Sql = Sql & ", FLGPADRAO" & VbNewLine 
   Sql = Sql &" From GCONEXOES" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDCONEXAO)  Then 
      Sql = Sql & " IDCONEXAO = " & SqlNum(mvarIDCONEXAO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDCONEXAO) Then Sql = Sql & " IDCONEXAO = " & SqlNum(Cstr(Ch_IDCONEXAO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDCONEXAO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDCONEXAO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCONEXAO = XVal(!IDCONEXAO & "")
      mvarNMSISTEMA = !NMSISTEMA & ""
      mvarTITULO = !TITULO & ""
      mvarTIPOSGBD = XVal(!TIPOSGBD & "")
      mvarTIPOCONEXAO = !TIPOCONEXAO & ""
      mvarNMSERVIDOR = !NMSERVIDOR & ""
      mvarNMBANCO = !NMBANCO & ""
      mvarUSUARIO = !USUARIO & ""
      mvarSENHA = !SENHA & ""
      mvarFLGPADRAO = !FLGPADRAO & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDCONEXAO = 0
   mvarNMSISTEMA = ""
   mvarTITULO = ""
   mvarTIPOSGBD = 0
   mvarTIPOCONEXAO = ""
   mvarNMSERVIDOR = ""
   mvarNMBANCO = ""
   mvarUSUARIO = ""
   mvarSENHA = ""
   mvarFLGPADRAO = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDCONEXAO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
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
