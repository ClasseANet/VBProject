VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GRPACESSO"
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

Private mvarID        As Double
Private mvarIDMODU    As String
Private mvarIDGRUPO   As String
Private mvarCODACESSO As String
Private mvarUSUDELEGA As String
Public Property Let ID(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarID <> vData)
   mvarID = vData
End Property
Public Property Get ID() As Double
   ID = mvarID
End Property
Public Property Let IDMODU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMODU <> vData)
   mvarIDMODU = vData
End Property
Public Property Get IDMODU() As String
   IDMODU = mvarIDMODU
End Property
Public Property Let IDGRUPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDGRUPO <> vData)
   mvarIDGRUPO = vData
End Property
Public Property Get IDGRUPO() As String
   IDGRUPO = mvarIDGRUPO
End Property
Public Property Let CODACESSO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODACESSO <> vData)
   mvarCODACESSO = vData
End Property
Public Property Get CODACESSO() As String
   CODACESSO = mvarCODACESSO
End Property
Public Property Let USUDELEGA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUSUDELEGA <> vData)
   mvarUSUDELEGA = vData
End Property
Public Property Get USUDELEGA() As String
   USUDELEGA = mvarUSUDELEGA
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
 
   Sql = "Insert Into GRPACESSO (ID, IDMODU, IDGRUPO, CODACESSO" & VbNewLine 
   Sql = Sql & ", USUDELEGA) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarID) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDGRUPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODACESSO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUSUDELEGA) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDMODU, Optional Ch_IDGRUPO) As String
   Dim Sql As String
 
   Sql = "Delete From GRPACESSO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDMODU) Then Sql = Sql & " IDMODU = " & SqlStr(Cstr(Ch_IDMODU)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDGRUPO) Then Sql = Sql & " IDGRUPO = " & SqlStr(Cstr(Ch_IDGRUPO)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GRPACESSO Set " & VbNewLine 
   Sql = Sql & " ID        = " & SqlNum(mvarID) & VbNewLine 
   Sql = Sql & " , IDMODU    = " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & " , IDGRUPO   = " & SqlStr(mvarIDGRUPO) & VbNewLine 
   Sql = Sql & " , CODACESSO = " & SqlStr(mvarCODACESSO) & VbNewLine 
   Sql = Sql & " , USUDELEGA = " & SqlStr(mvarUSUDELEGA) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GRPACESSO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDMODU      = " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & " And IDGRUPO = " & SqlStr(mvarIDGRUPO) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDMODU, Optional Ch_IDGRUPO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select ID, IDMODU, IDGRUPO, CODACESSO" & VbNewLine 
   Sql = Sql & ", USUDELEGA" & VbNewLine 
   Sql = Sql &" From GRPACESSO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDMODU)  And IsMissing(Ch_IDGRUPO)  Then 
      Sql = Sql & " IDMODU = " & SqlStr(mvarIDMODU) & " AND " & VbNewLine 
      Sql = Sql & " IDGRUPO = " & SqlStr(mvarIDGRUPO) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDMODU) Then Sql = Sql & " IDMODU = " & SqlStr(Cstr(Ch_IDMODU)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDGRUPO) Then Sql = Sql & " IDGRUPO = " & SqlStr(Cstr(Ch_IDGRUPO)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDMODU, Optional Ch_IDGRUPO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDMODU, Ch_IDGRUPO, Ch_WHERE, Ch_ORDERBY)
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
      mvarID = XVal(!ID & "")
      mvarIDMODU = !IDMODU & ""
      mvarIDGRUPO = !IDGRUPO & ""
      mvarCODACESSO = !CODACESSO & ""
      mvarUSUDELEGA = !USUDELEGA & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDMODU, mvarIDGRUPO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDMODU As String, Ch_IDGRUPO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GRPACESSO Set "
   Sql = Sql & " IDMODU = " & SqlStr(Ch_IDMODU)
   Sql = Sql & " , IDGRUPO = " & SqlStr(Ch_IDGRUPO)
   Sql = Sql & " Where 
   Sql = Sql & " IDMODU = " & SqlStr(mvarIDMODU)
   Sql = Sql & " and IDGRUPO = " & SqlStr(mvarIDGRUPO)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
