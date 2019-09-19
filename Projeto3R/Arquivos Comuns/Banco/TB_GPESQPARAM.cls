VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GPESQPARAM"
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

Private mvarIDPARAM     As Double
Private mvarCODPARAM    As String
Private mvarTITULO      As String
Private mvarTIPO        As String
Private mvarLISTA       As String
Private mvarSELMULT     As String
Private mvarLOV         As String
Private mvarOBRIGATORIO As String
Private mvarSCALEWIDTH  As Double
Private mvarSCALEHEIGHT As Double
Private mvarIDCONEXAO   As Double
Public Property Get IDPARAM() As Double
   IDPARAM = mvarIDPARAM
End Property
Public Property Let CODPARAM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODPARAM <> vData)
   mvarCODPARAM = vData
End Property
Public Property Get CODPARAM() As String
   CODPARAM = mvarCODPARAM
End Property
Public Property Let TITULO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTITULO <> vData)
   mvarTITULO = vData
End Property
Public Property Get TITULO() As String
   TITULO = mvarTITULO
End Property
Public Property Let TIPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPO <> vData)
   mvarTIPO = vData
End Property
Public Property Get TIPO() As String
   TIPO = mvarTIPO
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
Public Property Let IDCONEXAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCONEXAO <> vData)
   mvarIDCONEXAO = vData
End Property
Public Property Get IDCONEXAO() As Double
   IDCONEXAO = mvarIDCONEXAO
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
 
   Sql = "Insert Into GPESQPARAM (CODPARAM, TITULO, TIPO" & VbNewLine 
   Sql = Sql & ", LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH" & VbNewLine 
   Sql = Sql & ", SCALEHEIGHT, IDCONEXAO) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarCODPARAM) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTITULO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLISTA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSELMULT) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLOV) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBRIGATORIO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSCALEWIDTH) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSCALEHEIGHT) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCONEXAO) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDPARAM, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From GPESQPARAM" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPARAM)  Then 
      Sql = Sql & " IDPARAM = " & SqlNum(mvarIDPARAM) & " AND " & VbNewLine 
   Else 
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
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update GPESQPARAM Set " & VbNewLine 
   Sql = Sql & " CODPARAM    = " & SqlStr(mvarCODPARAM) & VbNewLine 
   Sql = Sql & " , TITULO      = " & SqlStr(mvarTITULO) & VbNewLine 
   Sql = Sql & " , TIPO        = " & SqlStr(mvarTIPO) & VbNewLine 
   Sql = Sql & " , LISTA       = " & SqlStr(mvarLISTA) & VbNewLine 
   Sql = Sql & " , SELMULT     = " & SqlStr(mvarSELMULT) & VbNewLine 
   Sql = Sql & " , LOV         = " & SqlStr(mvarLOV) & VbNewLine 
   Sql = Sql & " , OBRIGATORIO = " & SqlStr(mvarOBRIGATORIO) & VbNewLine 
   Sql = Sql & " , SCALEWIDTH  = " & SqlNum(mvarSCALEWIDTH) & VbNewLine 
   Sql = Sql & " , SCALEHEIGHT = " & SqlNum(mvarSCALEHEIGHT) & VbNewLine 
   Sql = Sql & " , IDCONEXAO   = " & SqlNum(mvarIDCONEXAO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GPESQPARAM'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDPARAM       = " & SqlNum(mvarIDPARAM) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDPARAM, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDPARAM, CODPARAM, TITULO, TIPO" & VbNewLine 
   Sql = Sql & ", LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH" & VbNewLine 
   Sql = Sql & ", SCALEHEIGHT, IDCONEXAO" & VbNewLine 
   Sql = Sql &" From GPESQPARAM" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDPARAM)  Then 
      Sql = Sql & " IDPARAM = " & SqlNum(mvarIDPARAM) & " AND " & VbNewLine 
   Else 
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
Public Function Pesquisar( Optional Ch_IDPARAM, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDPARAM, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPARAM = XVal(!IDPARAM & "")
      mvarCODPARAM = !CODPARAM & ""
      mvarTITULO = !TITULO & ""
      mvarTIPO = !TIPO & ""
      mvarLISTA = !LISTA & ""
      mvarSELMULT = !SELMULT & ""
      mvarLOV = !LOV & ""
      mvarOBRIGATORIO = !OBRIGATORIO & ""
      mvarSCALEWIDTH = XVal(!SCALEWIDTH & "")
      mvarSCALEHEIGHT = XVal(!SCALEHEIGHT & "")
      mvarIDCONEXAO = XVal(!IDCONEXAO & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDPARAM = 0
   mvarCODPARAM = ""
   mvarTITULO = ""
   mvarTIPO = ""
   mvarLISTA = ""
   mvarSELMULT = ""
   mvarLOV = ""
   mvarOBRIGATORIO = ""
   mvarSCALEWIDTH = 0
   mvarSCALEHEIGHT = 0
   mvarIDCONEXAO = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDPARAM), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
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
