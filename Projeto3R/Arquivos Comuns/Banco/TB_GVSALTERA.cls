VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GVSALTERA"
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

Private mvarIDVSALTERA   As Double
Private mvarCODSIS       As String
Private mvarCODVERSAOSIS As String
Private mvarTITVSALTERA  As String
Private mvarDSCVSALTERA  As String
Private mvarDTVSALTERA   As String
Public Property Get IDVSALTERA() As Double
   IDVSALTERA = mvarIDVSALTERA
End Property
Public Property Let CODSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIS <> vData)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let CODVERSAOSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODVERSAOSIS <> vData)
   mvarCODVERSAOSIS = vData
End Property
Public Property Get CODVERSAOSIS() As String
   CODVERSAOSIS = mvarCODVERSAOSIS
End Property
Public Property Let TITVSALTERA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTITVSALTERA <> vData)
   mvarTITVSALTERA = vData
End Property
Public Property Get TITVSALTERA() As String
   TITVSALTERA = mvarTITVSALTERA
End Property
Public Property Let DSCVSALTERA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCVSALTERA <> vData)
   mvarDSCVSALTERA = vData
End Property
Public Property Get DSCVSALTERA() As String
   DSCVSALTERA = mvarDSCVSALTERA
End Property
Public Property Let DTVSALTERA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTVSALTERA <> vData)
   mvarDTVSALTERA = vData
End Property
Public Property Get DTVSALTERA() As String
   DTVSALTERA = mvarDTVSALTERA
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
 
   Sql = "Insert Into GVSALTERA (CODSIS, CODVERSAOSIS, TITVSALTERA" & VbNewLine 
   Sql = Sql & ", DSCVSALTERA, DTVSALTERA) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODVERSAOSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTITVSALTERA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCVSALTERA) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTVSALTERA, eSysDate.Data_Hora) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDVSALTERA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From GVSALTERA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDVSALTERA)  Then 
      Sql = Sql & " IDVSALTERA = " & SqlNum(mvarIDVSALTERA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDVSALTERA) Then Sql = Sql & " IDVSALTERA = " & SqlNum(Cstr(Ch_IDVSALTERA)) & " AND " & VbNewLine 
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
 
   Sql = "Update GVSALTERA Set " & VbNewLine 
   Sql = Sql & " CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , CODVERSAOSIS = " & SqlStr(mvarCODVERSAOSIS) & VbNewLine 
   Sql = Sql & " , TITVSALTERA  = " & SqlStr(mvarTITVSALTERA) & VbNewLine 
   Sql = Sql & " , DSCVSALTERA  = " & SqlStr(mvarDSCVSALTERA) & VbNewLine 
   Sql = Sql & " , DTVSALTERA   = " & SqlDate(mvarDTVSALTERA, eSysDate.Data_Hora) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GVSALTERA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDVSALTERA     = " & SqlNum(mvarIDVSALTERA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDVSALTERA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDVSALTERA, CODSIS, CODVERSAOSIS, TITVSALTERA" & VbNewLine 
   Sql = Sql & ", DSCVSALTERA, DTVSALTERA" & VbNewLine 
   Sql = Sql &" From GVSALTERA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDVSALTERA)  Then 
      Sql = Sql & " IDVSALTERA = " & SqlNum(mvarIDVSALTERA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDVSALTERA) Then Sql = Sql & " IDVSALTERA = " & SqlNum(Cstr(Ch_IDVSALTERA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDVSALTERA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDVSALTERA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDVSALTERA = XVal(!IDVSALTERA & "")
      mvarCODSIS = !CODSIS & ""
      mvarCODVERSAOSIS = !CODVERSAOSIS & ""
      mvarTITVSALTERA = !TITVSALTERA & ""
      mvarDSCVSALTERA = !DSCVSALTERA & ""
      mvarDTVSALTERA = xDate(!DTVSALTERA & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDVSALTERA = 0
   mvarCODSIS = ""
   mvarCODVERSAOSIS = ""
   mvarTITVSALTERA = ""
   mvarDSCVSALTERA = ""
   mvarDTVSALTERA = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDVSALTERA), ComCOMMIT)
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
