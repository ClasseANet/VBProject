VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_VLATRIBUTO"
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

Private mvarIDATR   As Double
Private mvarIDPROD  As Double
Private mvarVLATR   As String
Private mvarIDFILHO As Double
Public Property Let IDATR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDATR <> vData)
   mvarIDATR = vData
End Property
Public Property Get IDATR() As Double
   IDATR = mvarIDATR
End Property
Public Property Let IDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROD <> vData)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Double
   IDPROD = mvarIDPROD
End Property
Public Property Let VLATR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVLATR <> vData)
   mvarVLATR = vData
End Property
Public Property Get VLATR() As String
   VLATR = mvarVLATR
End Property
Public Property Let IDFILHO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFILHO <> vData)
   mvarIDFILHO = vData
End Property
Public Property Get IDFILHO() As Double
   IDFILHO = mvarIDFILHO
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
 
   Sql = "Insert Into VLATRIBUTO (IDATR, IDPROD, VLATR, IDFILHO) " & VbNewLine 
   Sql = Sql & " Values " & VbNewLine 
   Sql = Sql & "(" & VbNewLine 
   Sql = Sql &  SqlNum(mvarIDATR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVLATR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDFILHO) & VbNewLine 
   Sql = Sql & ")" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From VLATRIBUTO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update VLATRIBUTO Set " & VbNewLine 
   Sql = Sql & " IDATR   = " & SqlNum(mvarIDATR) & VbNewLine 
   Sql = Sql & " , IDPROD  = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " , VLATR   = " & SqlStr(mvarVLATR) & VbNewLine 
   Sql = Sql & " , IDFILHO = " & SqlNum(mvarIDFILHO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'VLATRIBUTO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 

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
Public Property Get QrySelect(Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDATR, IDPROD, VLATR, IDFILHO" & VbNewLine 
   Sql = Sql &" From VLATRIBUTO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY 

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar(Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_WHERE, Ch_ORDERBY)
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
      mvarIDATR = XVal(!IDATR & "")
      mvarIDPROD = XVal(!IDPROD & "")
      mvarVLATR = !VLATR & ""
      mvarIDFILHO = XVal(!IDFILHO & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDATR = 0
   mvarIDPROD = 0
   mvarVLATR = ""
   mvarIDFILHO = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete(), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update VLATRIBUTO Set "
   Sql = Sql & " Where 
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
