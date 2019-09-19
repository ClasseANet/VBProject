VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_PPNITEM"
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
Private mvarIDPN       As Double
Private mvarITEM       As Double
Private mvarDTITEM     As String
Private mvarVLAREA     As Double
Private mvarVLPROD     As Double
Private mvarVLFAT      As Double
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDPN(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPN <> vData)
   mvarIDPN = vData
End Property
Public Property Get IDPN() As Double
   IDPN = mvarIDPN
End Property
Public Property Let ITEM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarITEM <> vData)
   mvarITEM = vData
End Property
Public Property Get ITEM() As Double
   ITEM = mvarITEM
End Property
Public Property Let DTITEM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTITEM <> vData)
   mvarDTITEM = vData
End Property
Public Property Get DTITEM() As String
   DTITEM = mvarDTITEM
End Property
Public Property Let VLAREA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLAREA <> vData)
   mvarVLAREA = vData
End Property
Public Property Get VLAREA() As Double
   VLAREA = mvarVLAREA
End Property
Public Property Let VLPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLPROD <> vData)
   mvarVLPROD = vData
End Property
Public Property Get VLPROD() As Double
   VLPROD = mvarVLPROD
End Property
Public Property Let VLFAT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLFAT <> vData)
   mvarVLFAT = vData
End Property
Public Property Get VLFAT() As Double
   VLFAT = mvarVLFAT
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update PPNITEM Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDPN     = " & SqlNum(mvarIDPN) & VbNewLine 
   Sql = Sql & " And ITEM     = " & SqlNum(mvarITEM) & VbNewLine 
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
 
   Sql = "Insert Into PPNITEM (IDLOJA, IDPN, ITEM, DTITEM" & VbNewLine 
   Sql = Sql & ", VLAREA, VLPROD, VLFAT, ALTERSTAMP, TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPN) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarITEM) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTITEM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLAREA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLPROD) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLFAT) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDPN, Optional Ch_ITEM) As String
   Dim Sql As String
 
   Sql = "Delete From PPNITEM" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDPN) Then Sql = Sql & " IDPN = " & SqlNum(Cstr(Ch_IDPN)) & " AND " & VbNewLine 
   if Not isMissing(Ch_ITEM) Then Sql = Sql & " ITEM = " & SqlNum(Cstr(Ch_ITEM)) & " AND " & VbNewLine 
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
 
   Sql = "Update PPNITEM Set " & VbNewLine 
   Sql = Sql & " IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDPN       = " & SqlNum(mvarIDPN) & VbNewLine 
   Sql = Sql & " , ITEM       = " & SqlNum(mvarITEM) & VbNewLine 
   Sql = Sql & " , DTITEM     = " & SqlDate(mvarDTITEM, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VLAREA     = " & SqlNum(mvarVLAREA) & VbNewLine 
   Sql = Sql & " , VLPROD     = " & SqlNum(mvarVLPROD) & VbNewLine 
   Sql = Sql & " , VLFAT      = " & SqlNum(mvarVLFAT) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP  = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'PPNITEM'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDPN     = " & SqlNum(mvarIDPN) & VbNewLine 
   Sql = Sql & " And ITEM     = " & SqlNum(mvarITEM) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDPN, Optional Ch_ITEM, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDPN, ITEM, DTITEM" & VbNewLine 
   Sql = Sql & ", VLAREA, VLPROD, VLFAT, ALTERSTAMP, TIMESTAMP" & VbNewLine 
   Sql = Sql &" From PPNITEM" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDPN)  And IsMissing(Ch_ITEM)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDPN = " & SqlNum(mvarIDPN) & " AND " & VbNewLine 
      Sql = Sql & " ITEM = " & SqlNum(mvarITEM) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPN) Then Sql = Sql & " IDPN = " & SqlNum(Cstr(Ch_IDPN)) & " AND " & VbNewLine 
      If Not isMissing(Ch_ITEM) Then Sql = Sql & " ITEM = " & SqlNum(Cstr(Ch_ITEM)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDPN, Optional Ch_ITEM, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDPN, Ch_ITEM, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPN = XVal(!IDPN & "")
      mvarITEM = XVal(!ITEM & "")
      mvarDTITEM = xDate(!DTITEM & "", True)
      mvarVLAREA = XVal(!VLAREA & "")
      mvarVLPROD = XVal(!VLPROD & "")
      mvarVLFAT = XVal(!VLFAT & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDPN, mvarITEM), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDPN As String, Ch_ITEM As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update PPNITEM Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDPN = " & SqlNum(Ch_IDPN)
   Sql = Sql & " , ITEM = " & SqlNum(Ch_ITEM)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDPN = " & SqlNum(mvarIDPN)
   Sql = Sql & " and ITEM = " & SqlNum(mvarITEM)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
