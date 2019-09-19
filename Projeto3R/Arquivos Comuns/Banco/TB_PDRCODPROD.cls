VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_PDRCODPROD"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb         As DS_BANCO
Private mvarRS          As Object 

Private mvarQryInsert   As String
Private mvarQryUpDate   As String
Private mvarQryDelete   As String
Private mvarQrySelect   As String
Private mvarQrySave     As String

Private mvarisDirt      As Boolean

Private mvarIDPDRCOD  As Double
Private mvarEPADRAO   As String
Private mvarIDFOR     As Double
Private mvarMSKPDRCOD As String
Private mvarDSCPDRCOD As String
Private mvarNMPDR     As String
Public Property Let IDPDRCOD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPDRCOD <> vData)
   mvarIDPDRCOD = vData
End Property
Public Property Get IDPDRCOD() As Double
   IDPDRCOD = mvarIDPDRCOD
End Property
Public Property Let EPADRAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarEPADRAO <> vData)
   mvarEPADRAO = vData
End Property
Public Property Get EPADRAO() As String
   EPADRAO = mvarEPADRAO
End Property
Public Property Let IDFOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFOR <> vData)
   mvarIDFOR = vData
End Property
Public Property Get IDFOR() As Double
   IDFOR = mvarIDFOR
End Property
Public Property Let MSKPDRCOD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMSKPDRCOD <> vData)
   mvarMSKPDRCOD = vData
End Property
Public Property Get MSKPDRCOD() As String
   MSKPDRCOD = mvarMSKPDRCOD
End Property
Public Property Let DSCPDRCOD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCPDRCOD <> vData)
   mvarDSCPDRCOD = vData
End Property
Public Property Get DSCPDRCOD() As String
   DSCPDRCOD = mvarDSCPDRCOD
End Property
Public Property Let NMPDR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMPDR <> vData)
   mvarNMPDR = vData
End Property
Public Property Get NMPDR() As String
   NMPDR = mvarNMPDR
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
Public Property Get QryInsert() As String
   Dim Sql As String
 
   Sql = "Insert Into PDRCODPROD (IDPDRCOD, EPADRAO, IDFOR, MSKPDRCOD" & VbNewLine 
   Sql = Sql & ", DSCPDRCOD, NMPDR) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  "isNull(Max(IDPDRCOD),0)+1" & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarEPADRAO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMSKPDRCOD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCPDRCOD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMPDR) & VbNewLine 
   Sql = Sql & " From PDRCODPROD" & VbNewLine 
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPDRCOD) As String
   Dim Sql As String
 
   Sql = "Delete From PDRCODPROD" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDPDRCOD) Then Sql = Sql & " IDPDRCOD = " & SqlNum(Cstr(Ch_IDPDRCOD)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update PDRCODPROD Set " & VbNewLine 
   Sql = Sql & " IDPDRCOD  = " & SqlNum(mvarIDPDRCOD) & VbNewLine 
   Sql = Sql & " , EPADRAO   = " & SqlStr(mvarEPADRAO) & VbNewLine 
   Sql = Sql & " , IDFOR     = " & SqlNum(mvarIDFOR) & VbNewLine 
   Sql = Sql & " , MSKPDRCOD = " & SqlStr(mvarMSKPDRCOD) & VbNewLine 
   Sql = Sql & " , DSCPDRCOD = " & SqlStr(mvarDSCPDRCOD) & VbNewLine 
   Sql = Sql & " , NMPDR     = " & SqlStr(mvarNMPDR) & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPDRCOD    = " & SqlNum(mvarIDPDRCOD) & VbNewLine 
   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave() As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_IDPDRCOD, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPDRCOD, EPADRAO, IDFOR, MSKPDRCOD" & VbNewLine 
   Sql = Sql & ", DSCPDRCOD, NMPDR" & VbNewLine 
   Sql = Sql &" From PDRCODPROD" & VbNewLine 
   If IsMissing(Ch_WHERE) Then
      Sql = Sql & " Where " & VbNewLine 
      If IsMissing(Ch_IDPDRCOD)  Then 
         Sql = Sql & " IDPDRCOD = " & SqlNum(mvarIDPDRCOD) & VbNewLine 
      Else 
         If Not isMissing(Ch_IDPDRCOD) Then Sql = Sql & " IDPDRCOD = " & SqlNum(Cstr(Ch_IDPDRCOD)) & " AND " & VbNewLine 
         Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
      End If 
   Else
      If Trim(Ch_WHERE) <> "" Then
         Sql = Sql & " Where " & Ch_WHERE & vbNewLine
      End If
   End If
   If Not IsMissing(Ch_ORDERBY) Then
      If Trim(Ch_ORDERBY) <> "" Then
         Sql = Sql & " Order By " & Ch_ORDERBY & vbNewLine
      End If
   End If

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar( Optional Ch_IDPDRCOD, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPDRCOD, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPDRCOD = XVal(!IDPDRCOD & "")
      mvarEPADRAO = !EPADRAO & ""
      mvarIDFOR = XVal(!IDFOR & "")
      mvarMSKPDRCOD = !MSKPDRCOD & ""
      mvarDSCPDRCOD = !DSCPDRCOD & ""
      mvarNMPDR = !NMPDR & ""
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPDRCOD), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDPDRCOD As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update PDRCODPROD Set "
   Sql = Sql & " IDPDRCOD = " & SqlNum(Ch_IDPDRCOD)
   Sql = Sql & " Where 
   Sql = Sql & " IDPDRCOD = " & SqlNum(mvarIDPDRCOD)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
