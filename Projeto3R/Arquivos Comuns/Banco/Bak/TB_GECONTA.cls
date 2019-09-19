VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GECONTA"
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

Private mvarIDECONTA     As Double
Private mvarTIPOSERVIDOR As Double
Private mvarSMTP         As String
Private mvarPOP          As String
Private mvarNMUSUARIO    As String
Private mvarENDMAIL      As String
Private mvarUSUARIO      As String
Private mvarSENHA        As String
Public Property Get IDECONTA() As Double
   IDECONTA = mvarIDECONTA
End Property
Public Property Let TIPOSERVIDOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPOSERVIDOR <> vData)
   mvarTIPOSERVIDOR = vData
End Property
Public Property Get TIPOSERVIDOR() As Double
   TIPOSERVIDOR = mvarTIPOSERVIDOR
End Property
Public Property Let SMTP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSMTP <> vData)
   mvarSMTP = vData
End Property
Public Property Get SMTP() As String
   SMTP = mvarSMTP
End Property
Public Property Let POP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarPOP <> vData)
   mvarPOP = vData
End Property
Public Property Get POP() As String
   POP = mvarPOP
End Property
Public Property Let NMUSUARIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMUSUARIO <> vData)
   mvarNMUSUARIO = vData
End Property
Public Property Get NMUSUARIO() As String
   NMUSUARIO = mvarNMUSUARIO
End Property
Public Property Let ENDMAIL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarENDMAIL <> vData)
   mvarENDMAIL = vData
End Property
Public Property Get ENDMAIL() As String
   ENDMAIL = mvarENDMAIL
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
 
   Sql = "Insert Into GECONTA (TIPOSERVIDOR, SMTP, POP" & VbNewLine 
   Sql = Sql & ", NMUSUARIO, ENDMAIL, USUARIO, SENHA) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarTIPOSERVIDOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSMTP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarPOP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMUSUARIO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarENDMAIL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUSUARIO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSENHA) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete() As String
   Dim Sql As String
 
   Sql = "Delete From GECONTA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GECONTA Set " & VbNewLine 
   Sql = Sql & " TIPOSERVIDOR = " & SqlNum(mvarTIPOSERVIDOR) & VbNewLine 
   Sql = Sql & " , SMTP         = " & SqlStr(mvarSMTP) & VbNewLine 
   Sql = Sql & " , POP          = " & SqlStr(mvarPOP) & VbNewLine 
   Sql = Sql & " , NMUSUARIO    = " & SqlStr(mvarNMUSUARIO) & VbNewLine 
   Sql = Sql & " , ENDMAIL      = " & SqlStr(mvarENDMAIL) & VbNewLine 
   Sql = Sql & " , USUARIO      = " & SqlStr(mvarUSUARIO) & VbNewLine 
   Sql = Sql & " , SENHA        = " & SqlStr(mvarSENHA) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GECONTA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 

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
Public Property Get QrySelect(Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDECONTA, TIPOSERVIDOR, SMTP, POP" & VbNewLine 
   Sql = Sql & ", NMUSUARIO, ENDMAIL, USUARIO, SENHA" & VbNewLine 
   Sql = Sql &" From GECONTA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY 

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar(Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_WHERE, Ch_ORDERBY)
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
      mvarIDECONTA = XVal(!IDECONTA & "")
      mvarTIPOSERVIDOR = XVal(!TIPOSERVIDOR & "")
      mvarSMTP = !SMTP & ""
      mvarPOP = !POP & ""
      mvarNMUSUARIO = !NMUSUARIO & ""
      mvarENDMAIL = !ENDMAIL & ""
      mvarUSUARIO = !USUARIO & ""
      mvarSENHA = !SENHA & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete(), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GECONTA Set "
   Sql = Sql & " Where 
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
