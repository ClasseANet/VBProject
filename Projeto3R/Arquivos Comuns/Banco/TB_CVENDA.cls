VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CVENDA"
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

Private mvarIDLOJA        As Double
Private mvarIDVENDA       As Double
Private mvarIDCLIENTE     As Double
Private mvarIDFUNCIONARIO As Double
Private mvarIDMOEDA       As Double
Private mvarDTVENDA       As String
Private mvarVLITENS       As Double
Private mvarVLVENDA       As Double
Private mvarVLPGTO        As Double
Private mvarVLDESC        As Double
Private mvarVLACRESC      As Double
Private mvarVLTROCO       As Double
Private mvarSITVENDA      As String
Private mvarFLGCANCELADA  As Double
Private mvarALTERSTAMP    As Double
Private mvarTIMESTAMP     As String
Private mvarIDFATURA      As Double
Private mvarCODVENDA      As Double
Private mvarOBS           As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDFUNCIONARIO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFUNCIONARIO <> vData)
   mvarIDFUNCIONARIO = vData
End Property
Public Property Get IDFUNCIONARIO() As Double
   IDFUNCIONARIO = mvarIDFUNCIONARIO
End Property
Public Property Let IDMOEDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMOEDA <> vData)
   mvarIDMOEDA = vData
End Property
Public Property Get IDMOEDA() As Double
   IDMOEDA = mvarIDMOEDA
End Property
Public Property Let DTVENDA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTVENDA <> vData)
   mvarDTVENDA = vData
End Property
Public Property Get DTVENDA() As String
   DTVENDA = mvarDTVENDA
End Property
Public Property Let VLITENS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLITENS <> vData)
   mvarVLITENS = vData
End Property
Public Property Get VLITENS() As Double
   VLITENS = mvarVLITENS
End Property
Public Property Let VLVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLVENDA <> vData)
   mvarVLVENDA = vData
End Property
Public Property Get VLVENDA() As Double
   VLVENDA = mvarVLVENDA
End Property
Public Property Let VLPGTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLPGTO <> vData)
   mvarVLPGTO = vData
End Property
Public Property Get VLPGTO() As Double
   VLPGTO = mvarVLPGTO
End Property
Public Property Let VLDESC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLDESC <> vData)
   mvarVLDESC = vData
End Property
Public Property Get VLDESC() As Double
   VLDESC = mvarVLDESC
End Property
Public Property Let VLACRESC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLACRESC <> vData)
   mvarVLACRESC = vData
End Property
Public Property Get VLACRESC() As Double
   VLACRESC = mvarVLACRESC
End Property
Public Property Let VLTROCO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLTROCO <> vData)
   mvarVLTROCO = vData
End Property
Public Property Get VLTROCO() As Double
   VLTROCO = mvarVLTROCO
End Property
Public Property Let SITVENDA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITVENDA <> vData)
   mvarSITVENDA = vData
End Property
Public Property Get SITVENDA() As String
   SITVENDA = mvarSITVENDA
End Property
Public Property Let FLGCANCELADA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGCANCELADA <> vData)
   mvarFLGCANCELADA = vData
End Property
Public Property Get FLGCANCELADA() As Double
   FLGCANCELADA = mvarFLGCANCELADA
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update CVENDA Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDVENDA     = " & SqlNum(mvarIDVENDA) & vbNewLine
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
Public Property Let TIMESTAMP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIMESTAMP <> vData)
   mvarTIMESTAMP = vData
End Property
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
End Property
Public Property Let IDFATURA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFATURA <> vData)
   mvarIDFATURA = vData
End Property
Public Property Get IDFATURA() As Double
   IDFATURA = mvarIDFATURA
End Property
Public Property Let CODVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarCODVENDA <> vData)
   mvarCODVENDA = vData
End Property
Public Property Get CODVENDA() As Double
   CODVENDA = mvarCODVENDA
End Property
Public Property Let OBS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarOBS <> vData)
   mvarOBS = vData
End Property
Public Property Get OBS() As String
   OBS = mvarOBS
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
Public Property Get QryInsert(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into CVENDA (IDLOJA, IDVENDA, IDCLIENTE, IDFUNCIONARIO" & vbNewLine
   Sql = Sql & ", IDMOEDA, DTVENDA, VLITENS, VLVENDA, VLPGTO" & vbNewLine
   Sql = Sql & ", VLDESC, VLACRESC, VLTROCO, SITVENDA, FLGCANCELADA" & vbNewLine
   Sql = Sql & ", IDFATURA, CODVENDA, OBS, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDVENDA),0)+1 From CVENDA Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDVENDA) & vbNewLine
   End If
   Sql = Sql & ", " & IIf(mvarIDCLIENTE = 0, "Null", SqlNum(mvarIDCLIENTE)) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDFUNCIONARIO = 0, "Null", SqlNum(mvarIDFUNCIONARIO)) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDMOEDA = 0, "Null", SqlNum(mvarIDMOEDA)) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTVENDA, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLITENS) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLVENDA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLPGTO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLDESC) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLACRESC) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVLTROCO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarSITVENDA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGCANCELADA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDFATURA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarCODVENDA) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarOBS) & vbNewLine
   If pSinc Then
      Sql = Sql & ", " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From CVENDA" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDVENDA) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDVENDA = " & SqlNum(mvarIDVENDA) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(CStr(Ch_IDVENDA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = ""
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDVENDA = " & mvarIDVENDA
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|IDFUNCIONARIO = " & mvarIDFUNCIONARIO
   sTag = sTag & "|IDMOEDA = " & mvarIDMOEDA
   sTag = sTag & "|DTVENDA = " & mvarDTVENDA
   sTag = sTag & "|VLITENS = " & mvarVLITENS
   sTag = sTag & "|VLVENDA = " & mvarVLVENDA
   sTag = sTag & "|VLPGTO = " & mvarVLPGTO
   sTag = sTag & "|VLDESC = " & mvarVLDESC
   sTag = sTag & "|VLACRESC = " & mvarVLACRESC
   sTag = sTag & "|VLTROCO = " & mvarVLTROCO
   sTag = sTag & "|SITVENDA = " & mvarSITVENDA
   sTag = sTag & "|FLGCANCELADA = " & mvarFLGCANCELADA
   sTag = sTag & "|IDFATURA = " & mvarIDFATURA
   sTag = sTag & "|CODVENDA = " & mvarCODVENDA
   sTag = sTag & "|OBS = " & mvarOBS
   sTag = sTag & "|Where = " & IIf(IsMissing(Ch_WHERE), "", Ch_WHERE)
   sTag = sTag & "|"

   MyDelRow.QUERY = Sql & ";"
   MyDelRow.TAG = sTag
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update CVENDA Set " & vbNewLine
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDVENDA       = " & SqlNum(mvarIDVENDA) & vbNewLine
   Sql = Sql & " , IDCLIENTE     = " & IIf(mvarIDCLIENTE = 0, "Null", SqlNum(mvarIDCLIENTE)) & vbNewLine
   Sql = Sql & " , IDFUNCIONARIO = " & IIf(mvarIDFUNCIONARIO = 0, "Null", SqlNum(mvarIDFUNCIONARIO)) & vbNewLine
   Sql = Sql & " , IDMOEDA       = " & IIf(mvarIDMOEDA = 0, "Null", SqlNum(mvarIDMOEDA)) & vbNewLine
   Sql = Sql & " , DTVENDA       = " & SqlDate(mvarDTVENDA, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , VLITENS       = " & SqlNum(mvarVLITENS) & vbNewLine
   Sql = Sql & " , VLVENDA       = " & SqlNum(mvarVLVENDA) & vbNewLine
   Sql = Sql & " , VLPGTO        = " & SqlNum(mvarVLPGTO) & vbNewLine
   Sql = Sql & " , VLDESC        = " & SqlNum(mvarVLDESC) & vbNewLine
   Sql = Sql & " , VLACRESC      = " & SqlNum(mvarVLACRESC) & vbNewLine
   Sql = Sql & " , VLTROCO       = " & SqlNum(mvarVLTROCO) & vbNewLine
   Sql = Sql & " , SITVENDA      = " & SqlStr(mvarSITVENDA) & vbNewLine
   Sql = Sql & " , FLGCANCELADA  = " & SqlNum(mvarFLGCANCELADA) & vbNewLine
   Sql = Sql & " , IDFATURA      = " & SqlNum(mvarIDFATURA) & vbNewLine
   Sql = Sql & " , CODVENDA      = " & SqlNum(mvarCODVENDA) & vbNewLine
   Sql = Sql & " , OBS           = " & SqlStr(mvarOBS) & vbNewLine
   If pSinc Then
      Sql = Sql & ", ALTERSTAMP    = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP    = 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", TIMESTAMP     = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP     = GetDate()" & vbNewLine
   End If
  
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CVENDA'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDVENDA     = " & SqlNum(mvarIDVENDA) & vbNewLine

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & vbNewLine
   Sql = Sql & Me.QryUpDate(pSinc:=pSinc)
   Sql = Sql & " Else " & vbNewLine
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId, pSinc:=pSinc)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDVENDA, IDCLIENTE, IDFUNCIONARIO" & vbNewLine
   Sql = Sql & ", IDMOEDA, DTVENDA, VLITENS, VLVENDA, VLPGTO" & vbNewLine
   Sql = Sql & ", VLDESC, VLACRESC, VLTROCO, SITVENDA, FLGCANCELADA" & vbNewLine
   Sql = Sql & ", IDFATURA, CODVENDA, OBS, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From CVENDA" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDVENDA) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDVENDA = " & SqlNum(mvarIDVENDA) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(CStr(Ch_IDVENDA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDVENDA, Ch_WHERE, Ch_ORDERBY)
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
    Call ShowError(Sql)
    Pesquisar = False
End Function
Public Sub Popula(Optional pRcSet)
   If IsMissing(pRcSet) Then Set pRcSet = mvarRS
   With pRcSet
      mvarIDLOJA = xVal(!IDLOJA & "")
      mvarIDVENDA = xVal(!IDVENDA & "")
      mvarIDCLIENTE = xVal(!IDCLIENTE & "")
      mvarIDFUNCIONARIO = xVal(!IDFUNCIONARIO & "")
      mvarIDMOEDA = xVal(!IDMOEDA & "")
      mvarDTVENDA = xDate(!DTVENDA & "", True)
      mvarVLITENS = xVal(!VLITENS & "")
      mvarVLVENDA = xVal(!VLVENDA & "")
      mvarVLPGTO = xVal(!VLPGTO & "")
      mvarVLDESC = xVal(!VLDESC & "")
      mvarVLACRESC = xVal(!VLACRESC & "")
      mvarVLTROCO = xVal(!VLTROCO & "")
      mvarSITVENDA = !SITVENDA & ""
      mvarFLGCANCELADA = xVal(!FLGCANCELADA & "")
      mvarIDFATURA = xVal(!IDFATURA & "")
      mvarCODVENDA = xVal(!CODVENDA & "")
      mvarOBS = !OBS & ""
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDVENDA = 0
   mvarIDCLIENTE = 0
   mvarIDFUNCIONARIO = 0
   mvarIDMOEDA = 0
   mvarDTVENDA = ""
   mvarVLITENS = 0
   mvarVLVENDA = 0
   mvarVLPGTO = 0
   mvarVLDESC = 0
   mvarVLACRESC = 0
   mvarVLTROCO = 0
   mvarSITVENDA = ""
   mvarTIMESTAMP = ""
   mvarIDFATURA = 0
   mvarCODVENDA = 0
   mvarOBS = ""
   mvarFLGCANCELADA = 0
   mvarALTERSTAMP = 0

   On Error Resume Next
   Call Class_Initialize
End Sub
Public Function Salvar(Optional ComCOMMIT As Boolean = True, Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId, pSinc:=pSinc), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId As Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDVENDA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDVENDA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CVENDA Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDVENDA = " & SqlNum(Ch_IDVENDA)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDVENDA = " & SqlNum(mvarIDVENDA)
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
