VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GVERSAOSIS_USUARIO"
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

Private mvarCODSIS       As String
Private mvarCODVERSAOSIS As String
Private mvarIDUSU        As String
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
Public Property Let IDUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSU <> vData)
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
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
 
   Sql = "Insert Into GVERSAOSIS_USUARIO (CODSIS, CODVERSAOSIS, IDUSU) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODVERSAOSIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_CODSIS, Optional Ch_CODVERSAOSIS, Optional Ch_IDUSU) As String
   Dim Sql As String
 
   Sql = "Delete From GVERSAOSIS_USUARIO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
   if Not isMissing(Ch_CODVERSAOSIS) Then Sql = Sql & " CODVERSAOSIS = " & SqlStr(Cstr(Ch_CODVERSAOSIS)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDUSU) Then Sql = Sql & " IDUSU = " & SqlStr(Cstr(Ch_IDUSU)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GVERSAOSIS_USUARIO Set " & VbNewLine 
   Sql = Sql & " CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , CODVERSAOSIS = " & SqlStr(mvarCODVERSAOSIS) & VbNewLine 
   Sql = Sql & " , IDUSU        = " & SqlStr(mvarIDUSU) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GVERSAOSIS_USUARIO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " CODSIS         = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " And CODVERSAOSIS = " & SqlStr(mvarCODVERSAOSIS) & VbNewLine 
   Sql = Sql & " And IDUSU      = " & SqlStr(mvarIDUSU) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_CODSIS, Optional Ch_CODVERSAOSIS, Optional Ch_IDUSU, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select CODSIS, CODVERSAOSIS, IDUSU" & VbNewLine 
   Sql = Sql &" From GVERSAOSIS_USUARIO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_CODSIS)  And IsMissing(Ch_CODVERSAOSIS)  And IsMissing(Ch_IDUSU)  Then 
      Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS) & " AND " & VbNewLine 
      Sql = Sql & " CODVERSAOSIS = " & SqlStr(mvarCODVERSAOSIS) & " AND " & VbNewLine 
      Sql = Sql & " IDUSU = " & SqlStr(mvarIDUSU) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_CODSIS) Then Sql = Sql & " CODSIS = " & SqlStr(Cstr(Ch_CODSIS)) & " AND " & VbNewLine 
      If Not isMissing(Ch_CODVERSAOSIS) Then Sql = Sql & " CODVERSAOSIS = " & SqlStr(Cstr(Ch_CODVERSAOSIS)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDUSU) Then Sql = Sql & " IDUSU = " & SqlStr(Cstr(Ch_IDUSU)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_CODSIS, Optional Ch_CODVERSAOSIS, Optional Ch_IDUSU, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_CODSIS, Ch_CODVERSAOSIS, Ch_IDUSU, Ch_WHERE, Ch_ORDERBY)
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
      mvarCODVERSAOSIS = !CODVERSAOSIS & ""
      mvarIDUSU = !IDUSU & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarCODSIS, mvarCODVERSAOSIS, mvarIDUSU), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_CODSIS As String, Ch_CODVERSAOSIS As String, Ch_IDUSU As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GVERSAOSIS_USUARIO Set "
   Sql = Sql & " CODSIS = " & SqlStr(Ch_CODSIS)
   Sql = Sql & " , CODVERSAOSIS = " & SqlStr(Ch_CODVERSAOSIS)
   Sql = Sql & " , IDUSU = " & SqlStr(Ch_IDUSU)
   Sql = Sql & " Where 
   Sql = Sql & " CODSIS = " & SqlStr(mvarCODSIS)
   Sql = Sql & " and CODVERSAOSIS = " & SqlStr(mvarCODVERSAOSIS)
   Sql = Sql & " and IDUSU = " & SqlStr(mvarIDUSU)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
