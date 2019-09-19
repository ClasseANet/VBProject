VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_DDPASSOS"
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

Private mvarIDTABELAO   As Double
Private mvarIDTABELAD   As Double
Private mvarDSCCAMPOREL As String
Private mvarCAMINHO     As String
Public Property Let IDTABELAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTABELAO <> vData)
   mvarIDTABELAO = vData
End Property
Public Property Get IDTABELAO() As Double
   IDTABELAO = mvarIDTABELAO
End Property
Public Property Let IDTABELAD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTABELAD <> vData)
   mvarIDTABELAD = vData
End Property
Public Property Get IDTABELAD() As Double
   IDTABELAD = mvarIDTABELAD
End Property
Public Property Let DSCCAMPOREL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCCAMPOREL <> vData)
   mvarDSCCAMPOREL = vData
End Property
Public Property Get DSCCAMPOREL() As String
   DSCCAMPOREL = mvarDSCCAMPOREL
End Property
Public Property Let CAMINHO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCAMINHO <> vData)
   mvarCAMINHO = vData
End Property
Public Property Get CAMINHO() As String
   CAMINHO = mvarCAMINHO
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
 
   Sql = "Insert Into DDPASSOS (IDTABELAO, IDTABELAD, DSCCAMPOREL, CAMINHO) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDTABELAO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTABELAD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCCAMPOREL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCAMINHO) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDTABELAO, Optional Ch_IDTABELAD, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From DDPASSOS" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDTABELAO)  And IsMissing(Ch_IDTABELAD)  Then 
      Sql = Sql & " IDTABELAO = " & SqlNum(mvarIDTABELAO) & " AND " & VbNewLine 
      Sql = Sql & " IDTABELAD = " & SqlNum(mvarIDTABELAD) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDTABELAO) Then Sql = Sql & " IDTABELAO = " & SqlNum(Cstr(Ch_IDTABELAO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDTABELAD) Then Sql = Sql & " IDTABELAD = " & SqlNum(Cstr(Ch_IDTABELAD)) & " AND " & VbNewLine 
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
 
   Sql = "Update DDPASSOS Set " & VbNewLine 
   Sql = Sql & " IDTABELAO   = " & SqlNum(mvarIDTABELAO) & VbNewLine 
   Sql = Sql & " , IDTABELAD   = " & SqlNum(mvarIDTABELAD) & VbNewLine 
   Sql = Sql & " , DSCCAMPOREL = " & SqlStr(mvarDSCCAMPOREL) & VbNewLine 
   Sql = Sql & " , CAMINHO     = " & SqlStr(mvarCAMINHO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'DDPASSOS'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDTABELAO     = " & SqlNum(mvarIDTABELAO) & VbNewLine 
   Sql = Sql & " And IDTABELAD = " & SqlNum(mvarIDTABELAD) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDTABELAO, Optional Ch_IDTABELAD, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDTABELAO, IDTABELAD, DSCCAMPOREL, CAMINHO" & VbNewLine 
   Sql = Sql &" From DDPASSOS" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDTABELAO)  And IsMissing(Ch_IDTABELAD)  Then 
      Sql = Sql & " IDTABELAO = " & SqlNum(mvarIDTABELAO) & " AND " & VbNewLine 
      Sql = Sql & " IDTABELAD = " & SqlNum(mvarIDTABELAD) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDTABELAO) Then Sql = Sql & " IDTABELAO = " & SqlNum(Cstr(Ch_IDTABELAO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDTABELAD) Then Sql = Sql & " IDTABELAD = " & SqlNum(Cstr(Ch_IDTABELAD)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDTABELAO, Optional Ch_IDTABELAD, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDTABELAO, Ch_IDTABELAD, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDTABELAO = XVal(!IDTABELAO & "")
      mvarIDTABELAD = XVal(!IDTABELAD & "")
      mvarDSCCAMPOREL = !DSCCAMPOREL & ""
      mvarCAMINHO = !CAMINHO & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDTABELAO = 0
   mvarIDTABELAD = 0
   mvarDSCCAMPOREL = ""
   mvarCAMINHO = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDTABELAO, mvarIDTABELAD), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDTABELAO As String, Ch_IDTABELAD As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update DDPASSOS Set "
   Sql = Sql & " IDTABELAO = " & SqlNum(Ch_IDTABELAO)
   Sql = Sql & " , IDTABELAD = " & SqlNum(Ch_IDTABELAD)
   Sql = Sql & " Where 
   Sql = Sql & " IDTABELAO = " & SqlNum(mvarIDTABELAO)
   Sql = Sql & " and IDTABELAD = " & SqlNum(mvarIDTABELAD)
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
