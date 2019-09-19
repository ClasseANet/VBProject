VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_OTPTRATAMENTO"
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

Private mvarIDLOJA         As Double
Private mvarIDTPTRATAMENTO As Double
Private mvarDSCTRATAMENTO  As String
Private mvarFREQUENCIA     As Double
Private mvarFLGDISPARO     As Double
Private mvarALTERSTAMP     As Double
Private mvarTIMESTAMP      As String
Private mvarATIVO          As Double
Private mvarCOR            As String
Private mvarFLGAREA        As Double
Private mvarFLGAVALIACAO   As Double
Private mvarFLGDEL         As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDTPTRATAMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPTRATAMENTO <> vData)
   mvarIDTPTRATAMENTO = vData
End Property
Public Property Get IDTPTRATAMENTO() As Double
   IDTPTRATAMENTO = mvarIDTPTRATAMENTO
End Property
Public Property Let DSCTRATAMENTO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCTRATAMENTO <> vData)
   mvarDSCTRATAMENTO = vData
End Property
Public Property Get DSCTRATAMENTO() As String
   DSCTRATAMENTO = mvarDSCTRATAMENTO
End Property
Public Property Let FREQUENCIA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFREQUENCIA <> vData)
   mvarFREQUENCIA = vData
End Property
Public Property Get FREQUENCIA() As Double
   FREQUENCIA = mvarFREQUENCIA
End Property
Public Property Let FLGDISPARO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGDISPARO <> vData)
   mvarFLGDISPARO = vData
End Property
Public Property Get FLGDISPARO() As Double
   FLGDISPARO = mvarFLGDISPARO
End Property
Public Property Let FLGDEL(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGDEL <> vData)
   mvarFLGDEL = vData
End Property
Public Property Get FLGDEL() As String
   FLGDEL = mvarFLGDEL
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OTPTRATAMENTO Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & vbNewLine
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
Public Property Let ATIVO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVO <> vData)
   mvarATIVO = vData
End Property
Public Property Get ATIVO() As Double
   ATIVO = mvarATIVO
End Property
Public Property Let COR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCOR <> vData)
   mvarCOR = vData
End Property
Public Property Get COR() As String
   COR = mvarCOR
End Property
Public Property Let FLGAREA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGAREA <> vData)
   mvarFLGAREA = vData
End Property
Public Property Get FLGAREA() As Double
   FLGAREA = mvarFLGAREA
End Property
Public Property Let FLGAVALIACAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGAVALIACAO <> vData)
   mvarFLGAVALIACAO = vData
End Property
Public Property Get FLGAVALIACAO() As Double
   FLGAVALIACAO = mvarFLGAVALIACAO
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
 
   Sql = "Insert Into OTPTRATAMENTO (IDLOJA, IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA" & vbNewLine
   Sql = Sql & ", FLGDISPARO, ALTERSTAMP, TIMESTAMP, ATIVO, COR" & vbNewLine
   Sql = Sql & ", FLGAREA, FLGAVALIACAO, FLGDEL) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDTPTRATAMENTO),0)+1 From OTPTRATAMENTO Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDTPTRATAMENTO) & vbNewLine
   End If
   Sql = Sql & ", " & SqlStr(mvarDSCTRATAMENTO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFREQUENCIA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGDISPARO) & vbNewLine
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
   Sql = Sql & ", " & SqlNum(mvarATIVO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarCOR) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGAREA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGAVALIACAO) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarFLGDEL) & vbNewLine
   
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDTPTRATAMENTO, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OTPTRATAMENTO" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDTPTRATAMENTO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDTPTRATAMENTO) Then Sql = Sql & " IDTPTRATAMENTO = " & SqlNum(CStr(Ch_IDTPTRATAMENTO)) & " AND " & vbNewLine
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
   sTag = sTag & "|IDTPTRATAMENTO = " & mvarIDTPTRATAMENTO
   sTag = sTag & "|DSCTRATAMENTO = " & mvarDSCTRATAMENTO
   sTag = sTag & "|FREQUENCIA = " & mvarFREQUENCIA
   sTag = sTag & "|FLGDISPARO = " & mvarFLGDISPARO
   sTag = sTag & "|ATIVO = " & mvarATIVO
   sTag = sTag & "|COR = " & mvarCOR
   sTag = sTag & "|FLGAREA = " & mvarFLGAREA
   sTag = sTag & "|FLGAVALIACAO = " & mvarFLGAVALIACAO
   sTag = sTag & "|FLGDEL = " & mvarFLGDEL
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
 
   Sql = "Update OTPTRATAMENTO Set " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & vbNewLine
   Sql = Sql & " , DSCTRATAMENTO  = " & SqlStr(mvarDSCTRATAMENTO) & vbNewLine
   Sql = Sql & " , FREQUENCIA     = " & SqlNum(mvarFREQUENCIA) & vbNewLine
   Sql = Sql & " , FLGDISPARO     = " & SqlNum(mvarFLGDISPARO) & vbNewLine
   If pSinc Then
      Sql = Sql & ", ALTERSTAMP     = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP     = 1" & vbNewLine
   End If
   If pSinc Then
      Sql = Sql & ", TIMESTAMP      = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP      = GetDate()" & vbNewLine
   End If
   Sql = Sql & " , ATIVO          = " & SqlNum(mvarATIVO) & vbNewLine
   Sql = Sql & " , COR            = " & SqlStr(mvarCOR) & vbNewLine
   Sql = Sql & " , FLGAREA        = " & SqlNum(mvarFLGAREA) & vbNewLine
   Sql = Sql & " , FLGAVALIACAO   = " & SqlNum(mvarFLGAVALIACAO) & vbNewLine
   Sql = Sql & " , FLGDEL         = " & SqlStr(mvarFLGDEL) & vbNewLine
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OTPTRATAMENTO'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & vbNewLine

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
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDTPTRATAMENTO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA" & vbNewLine
   Sql = Sql & ", FLGDISPARO, ALTERSTAMP, TIMESTAMP, ATIVO, COR" & vbNewLine
   Sql = Sql & ", FLGAREA, FLGAVALIACAO, FLGDEL" & vbNewLine
   Sql = Sql & " From OTPTRATAMENTO" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDTPTRATAMENTO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDTPTRATAMENTO) Then Sql = Sql & " IDTPTRATAMENTO = " & SqlNum(CStr(Ch_IDTPTRATAMENTO)) & " AND " & vbNewLine
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
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDTPTRATAMENTO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDTPTRATAMENTO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDTPTRATAMENTO = xVal(!IDTPTRATAMENTO & "")
      mvarDSCTRATAMENTO = !DSCTRATAMENTO & ""
      mvarFREQUENCIA = xVal(!FREQUENCIA & "")
      mvarFLGDISPARO = xVal(!FLGDISPARO & "")
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarATIVO = xVal(!ATIVO & "")
      mvarCOR = !COR & ""
      mvarFLGAREA = xVal(!FLGAREA & "")
      mvarFLGAVALIACAO = xVal(!FLGAVALIACAO & "")
      mvarFLGDEL = !FLGDEL & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDTPTRATAMENTO = 0
   mvarDSCTRATAMENTO = ""
   mvarFREQUENCIA = 0
   mvarFLGDISPARO = 0
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""
   mvarATIVO = 0
   mvarCOR = ""
   mvarFLGAREA = 0
   mvarFLGAVALIACAO = 0
   mvarFLGDEL = "0"

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
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDTPTRATAMENTO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDTPTRATAMENTO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OTPTRATAMENTO Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDTPTRATAMENTO = " & SqlNum(Ch_IDTPTRATAMENTO)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO)
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
