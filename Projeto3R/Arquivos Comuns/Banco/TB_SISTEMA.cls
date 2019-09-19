VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_SISTEMA"
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

Private mvarCODSIS        As String
Private mvarDSCSIS        As String
Private mvarSISPAI        As String
Private mvarCODVERSAOSIS  As String
Private mvarATIVO         As Double
Private mvarIDMODUDEFAULT As Double
Public Property Let CODSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIS <> vData)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let DSCSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCSIS <> vData)
   mvarDSCSIS = vData
End Property
Public Property Get DSCSIS() As String
   DSCSIS = mvarDSCSIS
End Property
Public Property Let SISPAI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSISPAI <> vData)
   mvarSISPAI = vData
End Property
Public Property Get SISPAI() As String
   SISPAI = mvarSISPAI
End Property
Public Property Let CODVERSAOSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODVERSAOSIS <> vData)
   mvarCODVERSAOSIS = vData
End Property
Public Property Get CODVERSAOSIS() As String
   CODVERSAOSIS = mvarCODVERSAOSIS
End Property
Public Property Let ATIVO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVO <> vData)
   mvarATIVO = vData
End Property
Public Property Get ATIVO() As Double
   ATIVO = mvarATIVO
End Property
Public Property Let IDMODUDEFAULT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMODUDEFAULT <> vData)
   mvarIDMODUDEFAULT = vData
End Property
Public Property Get IDMODUDEFAULT() As Double
   IDMODUDEFAULT = mvarIDMODUDEFAULT
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
 
   Sql = "Insert Into SISTEMA (CODSIS, DSCSIS, SISPAI, CODVERSAOSIS" & VbNewLine 
   Sql = Sql & ", ATIVO, IDMODUDEFAULT) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSISPAI) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODVERSAOSIS) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarATIVO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDMODUDEFAULT) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_CODSIS, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From SISTEMA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_CODSIS)  Then 
      Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
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
 
   Sql = "Update SISTEMA Set " & VbNewLine 
   Sql = Sql & " CODSIS        = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , DSCSIS        = " & SqlStr(mvarDSCSIS) & VbNewLine 
   Sql = Sql & " , SISPAI        = " & SqlStr(mvarSISPAI) & VbNewLine 
   Sql = Sql & " , CODVERSAOSIS  = " & SqlStr(mvarCODVERSAOSIS) & VbNewLine 
   Sql = Sql & " , ATIVO         = " & SqlNum(mvarATIVO) & VbNewLine 
   Sql = Sql & " , IDMODUDEFAULT = " & SqlNum(mvarIDMODUDEFAULT) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'SISTEMA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " CODSIS          = " & SqlStr(mvarCODSIS) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_CODSIS, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select CODSIS, DSCSIS, SISPAI, CODVERSAOSIS" & VbNewLine 
   Sql = Sql & ", ATIVO, IDMODUDEFAULT" & VbNewLine 
   Sql = Sql &" From SISTEMA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_CODSIS)  Then 
      Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_CODSIS, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_CODSIS, Ch_WHERE, Ch_ORDERBY)
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
      mvarCODSIS = !CODSIS & ""
      mvarDSCSIS = !DSCSIS & ""
      mvarSISPAI = !SISPAI & ""
      mvarCODVERSAOSIS = !CODVERSAOSIS & ""
      mvarATIVO = XVal(!ATIVO & "")
      mvarIDMODUDEFAULT = XVal(!IDMODUDEFAULT & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarCODSIS = ""
   mvarDSCSIS = ""
   mvarSISPAI = ""
   mvarCODVERSAOSIS = ""
   mvarATIVO = 0
   mvarIDMODUDEFAULT = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarCODSIS), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_CODSIS As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update SISTEMA Set "
   Sql = Sql & " CODSIS = " & SqlStr(Ch_CODSIS)
   Sql = Sql & " Where 
   Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS)
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
