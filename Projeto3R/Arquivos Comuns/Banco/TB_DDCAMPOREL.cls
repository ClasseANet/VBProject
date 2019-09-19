VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_DDCAMPOREL"
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

Private mvarIDTABREL As Double
Private mvarIDCAMPO  As Double
Private mvarIDCAMPOC As Double
Public Property Let IDTABREL(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTABREL <> vData)
   mvarIDTABREL = vData
End Property
Public Property Get IDTABREL() As Double
   IDTABREL = mvarIDTABREL
End Property
Public Property Let IDCAMPO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCAMPO <> vData)
   mvarIDCAMPO = vData
End Property
Public Property Get IDCAMPO() As Double
   IDCAMPO = mvarIDCAMPO
End Property
Public Property Let IDCAMPOC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCAMPOC <> vData)
   mvarIDCAMPOC = vData
End Property
Public Property Get IDCAMPOC() As Double
   IDCAMPOC = mvarIDCAMPOC
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
 
   Sql = "Insert Into DDCAMPOREL (IDTABREL, IDCAMPO, IDCAMPOC) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDTABREL) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCAMPO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCAMPOC) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDTABREL, Optional Ch_IDCAMPO, Optional Ch_IDCAMPOC, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From DDCAMPOREL" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDTABREL)  And IsMissing(Ch_IDCAMPO)  And IsMissing(Ch_IDCAMPOC)  Then 
      Sql = Sql & " IDTABREL = " & SqlNum(mvarIDTABREL) & " AND " & VbNewLine 
      Sql = Sql & " IDCAMPO = " & SqlNum(mvarIDCAMPO) & " AND " & VbNewLine 
      Sql = Sql & " IDCAMPOC = " & SqlNum(mvarIDCAMPOC) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDTABREL) Then Sql = Sql & " IDTABREL = " & SqlNum(Cstr(Ch_IDTABREL)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCAMPO) Then Sql = Sql & " IDCAMPO = " & SqlNum(Cstr(Ch_IDCAMPO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCAMPOC) Then Sql = Sql & " IDCAMPOC = " & SqlNum(Cstr(Ch_IDCAMPOC)) & " AND " & VbNewLine 
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
 
   Sql = "Update DDCAMPOREL Set " & VbNewLine 
   Sql = Sql & " IDTABREL = " & SqlNum(mvarIDTABREL) & VbNewLine 
   Sql = Sql & " , IDCAMPO  = " & SqlNum(mvarIDCAMPO) & VbNewLine 
   Sql = Sql & " , IDCAMPOC = " & SqlNum(mvarIDCAMPOC) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'DDCAMPOREL'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDTABREL   = " & SqlNum(mvarIDTABREL) & VbNewLine 
   Sql = Sql & " And IDCAMPO = " & SqlNum(mvarIDCAMPO) & VbNewLine 
   Sql = Sql & " And IDCAMPOC = " & SqlNum(mvarIDCAMPOC) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDTABREL, Optional Ch_IDCAMPO, Optional Ch_IDCAMPOC, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDTABREL, IDCAMPO, IDCAMPOC" & VbNewLine 
   Sql = Sql &" From DDCAMPOREL" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDTABREL)  And IsMissing(Ch_IDCAMPO)  And IsMissing(Ch_IDCAMPOC)  Then 
      Sql = Sql & " IDTABREL = " & SqlNum(mvarIDTABREL) & " AND " & VbNewLine 
      Sql = Sql & " IDCAMPO = " & SqlNum(mvarIDCAMPO) & " AND " & VbNewLine 
      Sql = Sql & " IDCAMPOC = " & SqlNum(mvarIDCAMPOC) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDTABREL) Then Sql = Sql & " IDTABREL = " & SqlNum(Cstr(Ch_IDTABREL)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCAMPO) Then Sql = Sql & " IDCAMPO = " & SqlNum(Cstr(Ch_IDCAMPO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCAMPOC) Then Sql = Sql & " IDCAMPOC = " & SqlNum(Cstr(Ch_IDCAMPOC)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDTABREL, Optional Ch_IDCAMPO, Optional Ch_IDCAMPOC, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDTABREL, Ch_IDCAMPO, Ch_IDCAMPOC, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDTABREL = XVal(!IDTABREL & "")
      mvarIDCAMPO = XVal(!IDCAMPO & "")
      mvarIDCAMPOC = XVal(!IDCAMPOC & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDTABREL = 0
   mvarIDCAMPO = 0
   mvarIDCAMPOC = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDTABREL, mvarIDCAMPO, mvarIDCAMPOC), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDTABREL As String, Ch_IDCAMPO As String, Ch_IDCAMPOC As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update DDCAMPOREL Set "
   Sql = Sql & " IDTABREL = " & SqlNum(Ch_IDTABREL)
   Sql = Sql & " , IDCAMPO = " & SqlNum(Ch_IDCAMPO)
   Sql = Sql & " , IDCAMPOC = " & SqlNum(Ch_IDCAMPOC)
   Sql = Sql & " Where 
   Sql = Sql & " IDTABREL = " & SqlNum(mvarIDTABREL)
   Sql = Sql & " and IDCAMPO = " & SqlNum(mvarIDCAMPO)
   Sql = Sql & " and IDCAMPOC = " & SqlNum(mvarIDCAMPOC)
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
