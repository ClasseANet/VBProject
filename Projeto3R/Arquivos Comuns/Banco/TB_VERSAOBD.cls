VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_VERSAOBD"
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

Private mvarIDBD   As Double
Private mvarDSCBD  As String
Private mvarVSBD   As String
Private mvarATUBD  As String
Private mvarDTATU  As String
Private mvarARQATU As String
Public Property Let IDBD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDBD <> vData)
   mvarIDBD = vData
End Property
Public Property Get IDBD() As Double
   IDBD = mvarIDBD
End Property
Public Property Let DSCBD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCBD <> vData)
   mvarDSCBD = vData
End Property
Public Property Get DSCBD() As String
   DSCBD = mvarDSCBD
End Property
Public Property Let VSBD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVSBD <> vData)
   mvarVSBD = vData
End Property
Public Property Get VSBD() As String
   VSBD = mvarVSBD
End Property
Public Property Let ATUBD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarATUBD <> vData)
   mvarATUBD = vData
End Property
Public Property Get ATUBD() As String
   ATUBD = mvarATUBD
End Property
Public Property Let DTATU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTATU <> vData)
   mvarDTATU = vData
End Property
Public Property Get DTATU() As String
   DTATU = mvarDTATU
End Property
Public Property Let ARQATU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarARQATU <> vData)
   mvarARQATU = vData
End Property
Public Property Get ARQATU() As String
   ARQATU = mvarARQATU
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
 
   Sql = "Insert Into VERSAOBD (IDBD, DSCBD, VSBD, ATUBD" & VbNewLine 
   Sql = Sql & ", DTATU, ARQATU) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   If pAutoId Then
      Sql = Sql &  "(Select isNull(Max(IDBD),0)+1 From VERSAOBD)"
   Else
      Sql = Sql &  SqlNum(mvarIDBD) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarDSCBD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVSBD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarATUBD) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTATU, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarARQATU) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDBD, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From VERSAOBD" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDBD)  Then 
      Sql = Sql & " IDBD = " & SqlNum(mvarIDBD) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDBD) Then Sql = Sql & " IDBD = " & SqlNum(Cstr(Ch_IDBD)) & " AND " & VbNewLine 
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
 
   Sql = "Update VERSAOBD Set " & VbNewLine 
   Sql = Sql & " IDBD   = " & SqlNum(mvarIDBD) & VbNewLine 
   Sql = Sql & " , DSCBD  = " & SqlStr(mvarDSCBD) & VbNewLine 
   Sql = Sql & " , VSBD   = " & SqlStr(mvarVSBD) & VbNewLine 
   Sql = Sql & " , ATUBD  = " & SqlStr(mvarATUBD) & VbNewLine 
   Sql = Sql & " , DTATU  = " & SqlDate(mvarDTATU, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , ARQATU = " & SqlStr(mvarARQATU) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'VERSAOBD'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDBD     = " & SqlNum(mvarIDBD) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDBD, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDBD, DSCBD, VSBD, ATUBD" & VbNewLine 
   Sql = Sql & ", DTATU, ARQATU" & VbNewLine 
   Sql = Sql &" From VERSAOBD" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDBD)  Then 
      Sql = Sql & " IDBD = " & SqlNum(mvarIDBD) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDBD) Then Sql = Sql & " IDBD = " & SqlNum(Cstr(Ch_IDBD)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDBD, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDBD, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDBD = XVal(!IDBD & "")
      mvarDSCBD = !DSCBD & ""
      mvarVSBD = !VSBD & ""
      mvarATUBD = !ATUBD & ""
      mvarDTATU = xDate(!DTATU & "", True)
      mvarARQATU = !ARQATU & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDBD = 0
   mvarDSCBD = ""
   mvarVSBD = ""
   mvarATUBD = ""
   mvarDTATU = ""
   mvarARQATU = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDBD), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDBD As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update VERSAOBD Set "
   Sql = Sql & " IDBD = " & SqlNum(Ch_IDBD)
   Sql = Sql & " Where 
   Sql = Sql & " IDBD = " & SqlNum(mvarIDBD)
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
