VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_FLOTERPS"
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

Private mvarIDLOJA     As Double
Private mvarIDLOTE     As Double
Private mvarNUMLOTE    As String
Private mvarDTINI      As String
Private mvarDTFIM      As String
Private mvarDTENVIO    As String
Private mvarSITLOTE    As Double
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDLOTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOTE <> vData)
   mvarIDLOTE = vData
End Property
Public Property Get IDLOTE() As Double
   IDLOTE = mvarIDLOTE
End Property
Public Property Let NUMLOTE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMLOTE <> vData)
   mvarNUMLOTE = vData
End Property
Public Property Get NUMLOTE() As String
   NUMLOTE = mvarNUMLOTE
End Property
Public Property Let DTINI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTINI <> vData)
   mvarDTINI = vData
End Property
Public Property Get DTINI() As String
   DTINI = mvarDTINI
End Property
Public Property Let DTFIM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTFIM <> vData)
   mvarDTFIM = vData
End Property
Public Property Get DTFIM() As String
   DTFIM = mvarDTFIM
End Property
Public Property Let DTENVIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTENVIO <> vData)
   mvarDTENVIO = vData
End Property
Public Property Get DTENVIO() As String
   DTENVIO = mvarDTENVIO
End Property
Public Property Let SITLOTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSITLOTE <> vData)
   mvarSITLOTE = vData
End Property
Public Property Get SITLOTE() As Double
   SITLOTE = mvarSITLOTE
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update FLOTERPS Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDLOTE   = " & SqlNum(mvarIDLOTE) & VbNewLine 
   If Not mvarXDb Is Nothing Then
      If mvarXDb.Conectado Then
         If mvarXDb.Executa(Sql, True) Then
            mvarALTERSTAMP = vData
         End If
      End If
   End If
End Property
Public Property Get ALTERSTAMP() As Double
   ALTERSTAMP = mvarALTERSTAMP
End Property
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
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
 
   Sql = "Insert Into FLOTERPS (IDLOJA, IDLOTE, NUMLOTE, DTINI" & VbNewLine 
   Sql = Sql & ", DTFIM, DTENVIO, SITLOTE, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDLOTE),0)+1 From FLOTERPS Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDLOTE) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarNUMLOTE) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDTENVIO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSITLOTE) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDLOTE) As String
   Dim Sql As String
 
   Sql = "Delete From FLOTERPS" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDLOTE) Then Sql = Sql & " IDLOTE = " & SqlNum(Cstr(Ch_IDLOTE)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   MyDelRow.Query = Sql & ";"
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update FLOTERPS Set " & VbNewLine 
   Sql = Sql & " IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDLOTE     = " & SqlNum(mvarIDLOTE) & VbNewLine 
   Sql = Sql & " , NUMLOTE    = " & SqlStr(mvarNUMLOTE) & VbNewLine 
   Sql = Sql & " , DTINI      = " & SqlDate(mvarDTINI, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTFIM      = " & SqlDate(mvarDTFIM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTENVIO    = " & SqlStr(mvarDTENVIO) & VbNewLine 
   Sql = Sql & " , SITLOTE    = " & SqlNum(mvarSITLOTE) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP  = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FLOTERPS'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDLOTE   = " & SqlNum(mvarIDLOTE) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDLOTE, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDLOTE, NUMLOTE, DTINI" & VbNewLine 
   Sql = Sql & ", DTFIM, DTENVIO, SITLOTE, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From FLOTERPS" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDLOTE)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDLOTE = " & SqlNum(mvarIDLOTE) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDLOTE) Then Sql = Sql & " IDLOTE = " & SqlNum(Cstr(Ch_IDLOTE)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDLOTE, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDLOTE, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDLOJA = XVal(!IDLOJA & "")
      mvarIDLOTE = XVal(!IDLOTE & "")
      mvarNUMLOTE = !NUMLOTE & ""
      mvarDTINI = xDate(!DTINI & "", True)
      mvarDTFIM = xDate(!DTFIM & "", True)
      mvarDTENVIO = !DTENVIO & ""
      mvarSITLOTE = XVal(!SITLOTE & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDLOTE), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDLOTE As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FLOTERPS Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDLOTE = " & SqlNum(Ch_IDLOTE)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDLOTE = " & SqlNum(mvarIDLOTE)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
