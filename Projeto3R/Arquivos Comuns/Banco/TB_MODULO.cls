VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_MODULO"
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

Private mvarID          As Double
Private mvarIDMODU      As String
Private mvarDSCMODU     As String
Private mvarSITMODU     As String
Private mvarMODUPAI     As String
Private mvarMENUDEFAULT As String
Private mvarIDPAI       As Double
Private mvarVBSCRIPT    As String
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
Public Property Let DSCMODU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCMODU <> vData)
   mvarDSCMODU = vData
End Property
Public Property Get DSCMODU() As String
   DSCMODU = mvarDSCMODU
End Property
Public Property Let SITMODU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITMODU <> vData)
   mvarSITMODU = vData
End Property
Public Property Get SITMODU() As String
   SITMODU = mvarSITMODU
End Property
Public Property Let MODUPAI(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMODUPAI <> vData)
   mvarMODUPAI = vData
End Property
Public Property Get MODUPAI() As String
   MODUPAI = mvarMODUPAI
End Property
Public Property Let MENUDEFAULT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMENUDEFAULT <> vData)
   mvarMENUDEFAULT = vData
End Property
Public Property Get MENUDEFAULT() As String
   MENUDEFAULT = mvarMENUDEFAULT
End Property
Public Property Let IDPAI(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPAI <> vData)
   mvarIDPAI = vData
End Property
Public Property Get IDPAI() As Double
   IDPAI = mvarIDPAI
End Property
Public Property Let VBSCRIPT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVBSCRIPT <> vData)
   mvarVBSCRIPT = vData
End Property
Public Property Get VBSCRIPT() As String
   VBSCRIPT = mvarVBSCRIPT
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
 
   Sql = "Insert Into MODULO (IDMODU, DSCMODU, SITMODU" & VbNewLine 
   Sql = Sql & ", MODUPAI, MENUDEFAULT, IDPAI, VBSCRIPT) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCMODU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSITMODU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMODUPAI) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMENUDEFAULT) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVBSCRIPT) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDMODU, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From MODULO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDMODU)  Then 
      Sql = Sql & " IDMODU = " & SqlStr(mvarIDMODU) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDMODU) Then Sql = Sql & " IDMODU = " & SqlStr(Cstr(Ch_IDMODU)) & " AND " & VbNewLine 
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
 
   Sql = "Update MODULO Set " & VbNewLine 
   Sql = Sql & " IDMODU      = " & SqlStr(mvarIDMODU) & VbNewLine 
   Sql = Sql & " , DSCMODU     = " & SqlStr(mvarDSCMODU) & VbNewLine 
   Sql = Sql & " , SITMODU     = " & SqlStr(mvarSITMODU) & VbNewLine 
   Sql = Sql & " , MODUPAI     = " & SqlStr(mvarMODUPAI) & VbNewLine 
   Sql = Sql & " , MENUDEFAULT = " & SqlStr(mvarMENUDEFAULT) & VbNewLine 
   Sql = Sql & " , IDPAI       = " & SqlNum(mvarIDPAI) & VbNewLine 
   Sql = Sql & " , VBSCRIPT    = " & SqlStr(mvarVBSCRIPT) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'MODULO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDMODU        = " & SqlStr(mvarIDMODU) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDMODU, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select ID, IDMODU, DSCMODU, SITMODU" & VbNewLine 
   Sql = Sql & ", MODUPAI, MENUDEFAULT, IDPAI, VBSCRIPT" & VbNewLine 
   Sql = Sql &" From MODULO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDMODU)  Then 
      Sql = Sql & " IDMODU = " & SqlStr(mvarIDMODU) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDMODU) Then Sql = Sql & " IDMODU = " & SqlStr(Cstr(Ch_IDMODU)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDMODU, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDMODU, Ch_WHERE, Ch_ORDERBY)
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
      mvarDSCMODU = !DSCMODU & ""
      mvarSITMODU = !SITMODU & ""
      mvarMODUPAI = !MODUPAI & ""
      mvarMENUDEFAULT = !MENUDEFAULT & ""
      mvarIDPAI = XVal(!IDPAI & "")
      mvarVBSCRIPT = !VBSCRIPT & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarID = 0
   mvarIDMODU = ""
   mvarDSCMODU = ""
   mvarSITMODU = ""
   mvarMODUPAI = ""
   mvarMENUDEFAULT = ""
   mvarIDPAI = 0
   mvarVBSCRIPT = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDMODU), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDMODU As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update MODULO Set "
   Sql = Sql & " IDMODU = " & SqlStr(Ch_IDMODU)
   Sql = Sql & " Where 
   Sql = Sql & " IDMODU = " & SqlStr(mvarIDMODU)
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
