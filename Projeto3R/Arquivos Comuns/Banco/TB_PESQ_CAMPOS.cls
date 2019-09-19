VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_PESQ_CAMPOS"
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

Private mvarIDPESQUISA   As Double
Private mvarIDCAMPO      As Double
Private mvarTABELA       As String
Private mvarCAMPO        As String
Private mvarOPERADOR     As String
Private mvarVALOR        As String
Private mvarCONCATENADOR As String
Private mvarTIPO         As String
Private mvarLISTA        As String
Private mvarDSCCAMPO     As String
Public Property Let IDPESQUISA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPESQUISA <> vData)
   mvarIDPESQUISA = vData
End Property
Public Property Get IDPESQUISA() As Double
   IDPESQUISA = mvarIDPESQUISA
End Property
Public Property Let IDCAMPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCAMPO <> vData)
   mvarIDCAMPO = vData
End Property
Public Property Get IDCAMPO() As Double
   IDCAMPO = mvarIDCAMPO
End Property
Public Property Let TABELA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTABELA <> vData)
   mvarTABELA = vData
End Property
Public Property Get TABELA() As String
   TABELA = mvarTABELA
End Property
Public Property Let CAMPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCAMPO <> vData)
   mvarCAMPO = vData
End Property
Public Property Get CAMPO() As String
   CAMPO = mvarCAMPO
End Property
Public Property Let OPERADOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOPERADOR <> vData)
   mvarOPERADOR = vData
End Property
Public Property Get OPERADOR() As String
   OPERADOR = mvarOPERADOR
End Property
Public Property Let VALOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVALOR <> vData)
   mvarVALOR = vData
End Property
Public Property Get VALOR() As String
   VALOR = mvarVALOR
End Property
Public Property Let CONCATENADOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCONCATENADOR <> vData)
   mvarCONCATENADOR = vData
End Property
Public Property Get CONCATENADOR() As String
   CONCATENADOR = mvarCONCATENADOR
End Property
Public Property Let TIPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPO <> vData)
   mvarTIPO = vData
End Property
Public Property Get TIPO() As String
   TIPO = mvarTIPO
End Property
Public Property Let LISTA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLISTA <> vData)
   mvarLISTA = vData
End Property
Public Property Get LISTA() As String
   LISTA = mvarLISTA
End Property
Public Property Let DSCCAMPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCCAMPO <> vData)
   mvarDSCCAMPO = vData
End Property
Public Property Get DSCCAMPO() As String
   DSCCAMPO = mvarDSCCAMPO
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
 
   Sql = "Insert Into PESQ_CAMPOS (IDPESQUISA, IDCAMPO, TABELA, CAMPO" & VbNewLine 
   Sql = Sql & ", OPERADOR, VALOR, CONCATENADOR, TIPO, LISTA" & VbNewLine 
   Sql = Sql & ", DSCCAMPO) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDPESQUISA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDCAMPO),0)+1 From PESQ_CAMPOS Where IDPESQUISA = " & SqlNum(mvarIDPESQUISA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDCAMPO) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarTABELA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCAMPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOPERADOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVALOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCONCATENADOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLISTA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCCAMPO) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPESQUISA, Optional Ch_IDCAMPO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From PESQ_CAMPOS" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPESQUISA)  And IsMissing(Ch_IDCAMPO)  Then 
      Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA) & " AND " & VbNewLine 
      Sql = Sql & " IDCAMPO = " & SqlNum(mvarIDCAMPO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDPESQUISA) Then Sql = Sql & " IDPESQUISA = " & SqlNum(Cstr(Ch_IDPESQUISA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCAMPO) Then Sql = Sql & " IDCAMPO = " & SqlNum(Cstr(Ch_IDCAMPO)) & " AND " & VbNewLine 
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
 
   Sql = "Update PESQ_CAMPOS Set " & VbNewLine 
   Sql = Sql & " IDPESQUISA   = " & SqlNum(mvarIDPESQUISA) & VbNewLine 
   Sql = Sql & " , IDCAMPO      = " & SqlNum(mvarIDCAMPO) & VbNewLine 
   Sql = Sql & " , TABELA       = " & SqlStr(mvarTABELA) & VbNewLine 
   Sql = Sql & " , CAMPO        = " & SqlStr(mvarCAMPO) & VbNewLine 
   Sql = Sql & " , OPERADOR     = " & SqlStr(mvarOPERADOR) & VbNewLine 
   Sql = Sql & " , VALOR        = " & SqlStr(mvarVALOR) & VbNewLine 
   Sql = Sql & " , CONCATENADOR = " & SqlStr(mvarCONCATENADOR) & VbNewLine 
   Sql = Sql & " , TIPO         = " & SqlStr(mvarTIPO) & VbNewLine 
   Sql = Sql & " , LISTA        = " & SqlStr(mvarLISTA) & VbNewLine 
   Sql = Sql & " , DSCCAMPO     = " & SqlStr(mvarDSCCAMPO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'PESQ_CAMPOS'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPESQUISA     = " & SqlNum(mvarIDPESQUISA) & VbNewLine 
   Sql = Sql & " And IDCAMPO    = " & SqlNum(mvarIDCAMPO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDPESQUISA, Optional Ch_IDCAMPO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPESQUISA, IDCAMPO, TABELA, CAMPO" & VbNewLine 
   Sql = Sql & ", OPERADOR, VALOR, CONCATENADOR, TIPO, LISTA" & VbNewLine 
   Sql = Sql & ", DSCCAMPO" & VbNewLine 
   Sql = Sql &" From PESQ_CAMPOS" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPESQUISA)  And IsMissing(Ch_IDCAMPO)  Then 
      Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA) & " AND " & VbNewLine 
      Sql = Sql & " IDCAMPO = " & SqlNum(mvarIDCAMPO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDPESQUISA) Then Sql = Sql & " IDPESQUISA = " & SqlNum(Cstr(Ch_IDPESQUISA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCAMPO) Then Sql = Sql & " IDCAMPO = " & SqlNum(Cstr(Ch_IDCAMPO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDPESQUISA, Optional Ch_IDCAMPO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPESQUISA, Ch_IDCAMPO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPESQUISA = XVal(!IDPESQUISA & "")
      mvarIDCAMPO = XVal(!IDCAMPO & "")
      mvarTABELA = !TABELA & ""
      mvarCAMPO = !CAMPO & ""
      mvarOPERADOR = !OPERADOR & ""
      mvarVALOR = !VALOR & ""
      mvarCONCATENADOR = !CONCATENADOR & ""
      mvarTIPO = !TIPO & ""
      mvarLISTA = !LISTA & ""
      mvarDSCCAMPO = !DSCCAMPO & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDPESQUISA = 0
   mvarIDCAMPO = 0
   mvarTABELA = ""
   mvarCAMPO = ""
   mvarOPERADOR = ""
   mvarVALOR = ""
   mvarCONCATENADOR = ""
   mvarTIPO = ""
   mvarLISTA = ""
   mvarDSCCAMPO = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPESQUISA, mvarIDCAMPO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDPESQUISA As String, Ch_IDCAMPO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update PESQ_CAMPOS Set "
   Sql = Sql & " IDPESQUISA = " & SqlNum(Ch_IDPESQUISA)
   Sql = Sql & " , IDCAMPO = " & SqlNum(Ch_IDCAMPO)
   Sql = Sql & " Where 
   Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA)
   Sql = Sql & " and IDCAMPO = " & SqlNum(mvarIDCAMPO)
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
