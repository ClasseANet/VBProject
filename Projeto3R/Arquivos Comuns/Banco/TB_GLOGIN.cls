VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GLOGIN"
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

Private mvarLOGINID      As Object
Private mvarIDUSU        As String
Private mvarCOMPUTERNAME As String
Private mvarCODSIS       As String
Private mvarLOGINTIME    As String
Private mvarCLIENTTYPE   As String
Private mvarDOMINIO      As String
Public Property Set LOGINID(ByVal vData As Object)
   If Not mvarisDirt Then mvarisDirt = (mvarLOGINID <> vData)
   Set mvarLOGINID = vData
End Property
Public Property Get LOGINID() As Object
   Set LOGINID = mvarLOGINID
End Property
Public Property Let IDUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSU <> vData)
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let COMPUTERNAME(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCOMPUTERNAME <> vData)
   mvarCOMPUTERNAME = vData
End Property
Public Property Get COMPUTERNAME() As String
   COMPUTERNAME = mvarCOMPUTERNAME
End Property
Public Property Let CODSIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODSIS <> vData)
   mvarCODSIS = vData
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let LOGINTIME(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLOGINTIME <> vData)
   mvarLOGINTIME = vData
End Property
Public Property Get LOGINTIME() As String
   LOGINTIME = mvarLOGINTIME
End Property
Public Property Let CLIENTTYPE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCLIENTTYPE <> vData)
   mvarCLIENTTYPE = vData
End Property
Public Property Get CLIENTTYPE() As String
   CLIENTTYPE = mvarCLIENTTYPE
End Property
Public Property Let DOMINIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDOMINIO <> vData)
   mvarDOMINIO = vData
End Property
Public Property Get DOMINIO() As String
   DOMINIO = mvarDOMINIO
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
 
   Sql = "Insert Into GLOGIN (IDUSU, COMPUTERNAME, CODSIS" & VbNewLine 
   Sql = Sql & ", LOGINTIME, CLIENTTYPE, DOMINIO) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &   & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCOMPUTERNAME) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarLOGINTIME, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCLIENTTYPE) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDOMINIO) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_LOGINID, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From GLOGIN" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_LOGINID)  Then 
      Sql = Sql & " LOGINID = " & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_LOGINID) Then Sql = Sql & " LOGINID = " & " AND " & VbNewLine 
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
 
   Sql = "Update GLOGIN Set " & VbNewLine 
   Sql = Sql & " , IDUSU        = " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & " , COMPUTERNAME = " & SqlStr(mvarCOMPUTERNAME) & VbNewLine 
   Sql = Sql & " , CODSIS       = " & SqlStr(mvarCODSIS) & VbNewLine 
   Sql = Sql & " , LOGINTIME    = " & SqlDate(mvarLOGINTIME, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , CLIENTTYPE   = " & SqlStr(mvarCLIENTTYPE) & VbNewLine 
   Sql = Sql & " , DOMINIO      = " & SqlStr(mvarDOMINIO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GLOGIN'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " LOGINID        = " & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_LOGINID, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select , IDUSU, COMPUTERNAME, CODSIS" & VbNewLine 
   Sql = Sql & ", LOGINTIME, CLIENTTYPE, DOMINIO" & VbNewLine 
   Sql = Sql &" From GLOGIN" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_LOGINID)  Then 
      Sql = Sql & " LOGINID = " & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_LOGINID) Then Sql = Sql & " LOGINID = " & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_LOGINID, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_LOGINID, Ch_WHERE, Ch_ORDERBY)
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

      mvarIDUSU = !IDUSU & ""
      mvarCOMPUTERNAME = !COMPUTERNAME & ""
      mvarCODSIS = !CODSIS & ""
      mvarLOGINTIME = xDate(!LOGINTIME & "", True)
      mvarCLIENTTYPE = !CLIENTTYPE & ""
      mvarDOMINIO = !DOMINIO & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()

   mvarIDUSU = ""
   mvarCOMPUTERNAME = ""
   mvarCODSIS = ""
   mvarLOGINTIME = ""
   mvarCLIENTTYPE = ""
   mvarDOMINIO = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarLOGINID), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_LOGINID As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GLOGIN Set "
   Sql = Sql & " LOGINID = "
   Sql = Sql & " Where 
   Sql = Sql & " LOGINID = "
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
