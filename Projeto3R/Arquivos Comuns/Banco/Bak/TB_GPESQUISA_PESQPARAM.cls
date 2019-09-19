VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GPESQUISA_PESQPARAM"
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

Private mvarIDPESQUISA  As Double
Private mvarIDPARAM     As Double
Private mvarLISTA       As String
Private mvarSELMULT     As String
Private mvarLOV         As String
Private mvarOBRIGATORIO As String
Private mvarSCALEWIDTH  As Double
Private mvarSCALEHEIGHT As Double
Private mvarORDEM       As Double
Public Property Let IDPESQUISA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPESQUISA <> vData)
   mvarIDPESQUISA = vData
End Property
Public Property Get IDPESQUISA() As Double
   IDPESQUISA = mvarIDPESQUISA
End Property
Public Property Let IDPARAM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPARAM <> vData)
   mvarIDPARAM = vData
End Property
Public Property Get IDPARAM() As Double
   IDPARAM = mvarIDPARAM
End Property
Public Property Let LISTA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLISTA <> vData)
   mvarLISTA = vData
End Property
Public Property Get LISTA() As String
   LISTA = mvarLISTA
End Property
Public Property Let SELMULT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSELMULT <> vData)
   mvarSELMULT = vData
End Property
Public Property Get SELMULT() As String
   SELMULT = mvarSELMULT
End Property
Public Property Let LOV(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLOV <> vData)
   mvarLOV = vData
End Property
Public Property Get LOV() As String
   LOV = mvarLOV
End Property
Public Property Let OBRIGATORIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBRIGATORIO <> vData)
   mvarOBRIGATORIO = vData
End Property
Public Property Get OBRIGATORIO() As String
   OBRIGATORIO = mvarOBRIGATORIO
End Property
Public Property Let SCALEWIDTH(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSCALEWIDTH <> vData)
   mvarSCALEWIDTH = vData
End Property
Public Property Get SCALEWIDTH() As Double
   SCALEWIDTH = mvarSCALEWIDTH
End Property
Public Property Let SCALEHEIGHT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSCALEHEIGHT <> vData)
   mvarSCALEHEIGHT = vData
End Property
Public Property Get SCALEHEIGHT() As Double
   SCALEHEIGHT = mvarSCALEHEIGHT
End Property
Public Property Let ORDEM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarORDEM <> vData)
   mvarORDEM = vData
End Property
Public Property Get ORDEM() As Double
   ORDEM = mvarORDEM
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
 
   Sql = "Insert Into GPESQUISA_PESQPARAM (IDPESQUISA, IDPARAM, LISTA, SELMULT" & VbNewLine 
   Sql = Sql & ", LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, ORDEM) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   If pAutoId Then
      Sql = Sql &  "(Select isNull(Max(IDPESQUISA),0)+1 From GPESQUISA_PESQPARAM Where IDPARAM = " & SqlNum(mvarIDPARAM) & ")" & vbNewLine
   Else
      Sql = Sql &  SqlNum(mvarIDPESQUISA) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDPARAM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLISTA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSELMULT) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLOV) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBRIGATORIO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSCALEWIDTH) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSCALEHEIGHT) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarORDEM) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPESQUISA, Optional Ch_IDPARAM) As String
   Dim Sql As String
 
   Sql = "Delete From GPESQUISA_PESQPARAM" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDPESQUISA) Then Sql = Sql & " IDPESQUISA = " & SqlNum(Cstr(Ch_IDPESQUISA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDPARAM) Then Sql = Sql & " IDPARAM = " & SqlNum(Cstr(Ch_IDPARAM)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GPESQUISA_PESQPARAM Set " & VbNewLine 
   Sql = Sql & " IDPESQUISA  = " & SqlNum(mvarIDPESQUISA) & VbNewLine 
   Sql = Sql & " , IDPARAM     = " & SqlNum(mvarIDPARAM) & VbNewLine 
   Sql = Sql & " , LISTA       = " & SqlStr(mvarLISTA) & VbNewLine 
   Sql = Sql & " , SELMULT     = " & SqlStr(mvarSELMULT) & VbNewLine 
   Sql = Sql & " , LOV         = " & SqlStr(mvarLOV) & VbNewLine 
   Sql = Sql & " , OBRIGATORIO = " & SqlStr(mvarOBRIGATORIO) & VbNewLine 
   Sql = Sql & " , SCALEWIDTH  = " & SqlNum(mvarSCALEWIDTH) & VbNewLine 
   Sql = Sql & " , SCALEHEIGHT = " & SqlNum(mvarSCALEHEIGHT) & VbNewLine 
   Sql = Sql & " , ORDEM       = " & SqlNum(mvarORDEM) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GPESQUISA_PESQPARAM'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPESQUISA    = " & SqlNum(mvarIDPESQUISA) & VbNewLine 
   Sql = Sql & " And IDPARAM   = " & SqlNum(mvarIDPARAM) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDPESQUISA, Optional Ch_IDPARAM, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPESQUISA, IDPARAM, LISTA, SELMULT" & VbNewLine 
   Sql = Sql & ", LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, ORDEM" & VbNewLine 
   Sql = Sql &" From GPESQUISA_PESQPARAM" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPESQUISA)  And IsMissing(Ch_IDPARAM)  Then 
      Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA) & " AND " & VbNewLine 
      Sql = Sql & " IDPARAM = " & SqlNum(mvarIDPARAM) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDPESQUISA) Then Sql = Sql & " IDPESQUISA = " & SqlNum(Cstr(Ch_IDPESQUISA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPARAM) Then Sql = Sql & " IDPARAM = " & SqlNum(Cstr(Ch_IDPARAM)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDPESQUISA, Optional Ch_IDPARAM, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPESQUISA, Ch_IDPARAM, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPESQUISA = XVal(!IDPESQUISA & "")
      mvarIDPARAM = XVal(!IDPARAM & "")
      mvarLISTA = !LISTA & ""
      mvarSELMULT = !SELMULT & ""
      mvarLOV = !LOV & ""
      mvarOBRIGATORIO = !OBRIGATORIO & ""
      mvarSCALEWIDTH = XVal(!SCALEWIDTH & "")
      mvarSCALEHEIGHT = XVal(!SCALEHEIGHT & "")
      mvarORDEM = XVal(!ORDEM & "")
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPESQUISA, mvarIDPARAM), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDPESQUISA As String, Ch_IDPARAM As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update GPESQUISA_PESQPARAM Set "
   Sql = Sql & " IDPESQUISA = " & SqlNum(Ch_IDPESQUISA)
   Sql = Sql & " , IDPARAM = " & SqlNum(Ch_IDPARAM)
   Sql = Sql & " Where 
   Sql = Sql & " IDPESQUISA = " & SqlNum(mvarIDPESQUISA)
   Sql = Sql & " and IDPARAM = " & SqlNum(mvarIDPARAM)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
