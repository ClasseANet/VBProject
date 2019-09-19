VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_GCAMPOCOMPL"
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

Private mvarIDCAMPOCOMPL As Double
Private mvarNMCAMPOCOMPL As String
Private mvarTABELA       As String
Private mvarTIPODADO     As String
Private mvarTAMANHODADO  As Double
Private mvarLISTADADO    As String
Private mvarOBRIGATORIO  As String
Private mvarSELMULT      As String
Public Property Get IDCAMPOCOMPL() As Double
   IDCAMPOCOMPL = mvarIDCAMPOCOMPL
End Property
Public Property Let NMCAMPOCOMPL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMCAMPOCOMPL <> vData)
   mvarNMCAMPOCOMPL = vData
End Property
Public Property Get NMCAMPOCOMPL() As String
   NMCAMPOCOMPL = mvarNMCAMPOCOMPL
End Property
Public Property Let TABELA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTABELA <> vData)
   mvarTABELA = vData
End Property
Public Property Get TABELA() As String
   TABELA = mvarTABELA
End Property
Public Property Let TIPODADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIPODADO <> vData)
   mvarTIPODADO = vData
End Property
Public Property Get TIPODADO() As String
   TIPODADO = mvarTIPODADO
End Property
Public Property Let TAMANHODADO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTAMANHODADO <> vData)
   mvarTAMANHODADO = vData
End Property
Public Property Get TAMANHODADO() As Double
   TAMANHODADO = mvarTAMANHODADO
End Property
Public Property Let LISTADADO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLISTADADO <> vData)
   mvarLISTADADO = vData
End Property
Public Property Get LISTADADO() As String
   LISTADADO = mvarLISTADADO
End Property
Public Property Let OBRIGATORIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBRIGATORIO <> vData)
   mvarOBRIGATORIO = vData
End Property
Public Property Get OBRIGATORIO() As String
   OBRIGATORIO = mvarOBRIGATORIO
End Property
Public Property Let SELMULT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSELMULT <> vData)
   mvarSELMULT = vData
End Property
Public Property Get SELMULT() As String
   SELMULT = mvarSELMULT
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
 
   Sql = "Insert Into GCAMPOCOMPL (NMCAMPOCOMPL, TABELA, TIPODADO" & VbNewLine 
   Sql = Sql & ", TAMANHODADO, LISTADADO, OBRIGATORIO, SELMULT) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarNMCAMPOCOMPL) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTABELA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTIPODADO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTAMANHODADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLISTADADO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarOBRIGATORIO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSELMULT) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDCAMPOCOMPL) As String
   Dim Sql As String
 
   Sql = "Delete From GCAMPOCOMPL" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDCAMPOCOMPL) Then Sql = Sql & " IDCAMPOCOMPL = " & SqlNum(Cstr(Ch_IDCAMPOCOMPL)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update GCAMPOCOMPL Set " & VbNewLine 
   Sql = Sql & " NMCAMPOCOMPL = " & SqlStr(mvarNMCAMPOCOMPL) & VbNewLine 
   Sql = Sql & " , TABELA       = " & SqlStr(mvarTABELA) & VbNewLine 
   Sql = Sql & " , TIPODADO     = " & SqlStr(mvarTIPODADO) & VbNewLine 
   Sql = Sql & " , TAMANHODADO  = " & SqlNum(mvarTAMANHODADO) & VbNewLine 
   Sql = Sql & " , LISTADADO    = " & SqlStr(mvarLISTADADO) & VbNewLine 
   Sql = Sql & " , OBRIGATORIO  = " & SqlStr(mvarOBRIGATORIO) & VbNewLine 
   Sql = Sql & " , SELMULT      = " & SqlStr(mvarSELMULT) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'GCAMPOCOMPL'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDCAMPOCOMPL   = " & SqlNum(mvarIDCAMPOCOMPL) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDCAMPOCOMPL, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDCAMPOCOMPL, NMCAMPOCOMPL, TABELA, TIPODADO" & VbNewLine 
   Sql = Sql & ", TAMANHODADO, LISTADADO, OBRIGATORIO, SELMULT" & VbNewLine 
   Sql = Sql &" From GCAMPOCOMPL" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDCAMPOCOMPL)  Then 
      Sql = Sql & " IDCAMPOCOMPL = " & SqlNum(mvarIDCAMPOCOMPL) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDCAMPOCOMPL) Then Sql = Sql & " IDCAMPOCOMPL = " & SqlNum(Cstr(Ch_IDCAMPOCOMPL)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDCAMPOCOMPL, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDCAMPOCOMPL, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDCAMPOCOMPL = XVal(!IDCAMPOCOMPL & "")
      mvarNMCAMPOCOMPL = !NMCAMPOCOMPL & ""
      mvarTABELA = !TABELA & ""
      mvarTIPODADO = !TIPODADO & ""
      mvarTAMANHODADO = XVal(!TAMANHODADO & "")
      mvarLISTADADO = !LISTADADO & ""
      mvarOBRIGATORIO = !OBRIGATORIO & ""
      mvarSELMULT = !SELMULT & ""
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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDCAMPOCOMPL), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
